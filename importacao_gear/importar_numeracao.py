#!/usr/bin/env python3
"""
Importação das planilhas de numeração (DET-DPL e Ordens de Serviço) para a
Central de Numeração do sistema.

Lê as abas DOCUMENTAÇÃO, VISTORIAS, PROJETOS e ORDEM DE SERVIÇO de qualquer
quantidade de arquivos e gera, em `saida/numeracao/`:

  numeros_utilizados.csv  um registro por número já emitido (não pode ser reutilizado)
  sequencias.csv          por sequência (tipo/regional/contrato + ano): qtd, último, PRÓXIMO número
  pendencias.csv          números repetidos, lacunas, ano divergente, formatos inválidos...
  vinculos_demandas.csv   (opcional, com --demandas) NTs/projetos citados nas demandas x registro
  revisao_numeracao.xlsx  tudo acima para conferência

Uso:
  python3 importar_numeracao.py NUMERACAO_*.xlsx
  python3 importar_numeracao.py --demandas saida/demandas.csv NUMERACAO_*.xlsx

Como as planilhas numeram
-------------------------
O número é uma FÓRMULA (COUNTIFS) que conta quantas linhas acima têm o mesmo:
  DOCUMENTO: TIPO DE DOCUMENTO          -> "NT GEAR.1052.2026"  (contador por TIPO, compartilhado entre setores)
  PROJETO:   REGIONAL                   -> "PROJETO GEAR RE 002.2026" (contador por REGIONAL, entre setores)
  OS:        CONTRATO                   -> "OS 001.IMP.TONY PLACAS.26" (contador por CONTRATO)
O ano vem da DATA da linha. A Central replica essas mesmas sequências.
"""
from __future__ import annotations

import argparse
import re
import sys
import unicodedata
from collections import defaultdict
from pathlib import Path

import pandas as pd

BASE = Path(__file__).resolve().parent

# Tipos de documento: sigla/alias usado na planilha -> (tipo canônico, sigla oficial)
# (aba INFORMAÇÕES, colunas TIPOS DE DOCUMENTOS / SIGLA, mais grafias antigas encontradas)
TIPOS = {
    "OF. INT.": ("OFÍCIO INTERNO", "OF. INT."),
    "OFICIO INTERNO": ("OFÍCIO INTERNO", "OF. INT."),
    "INF. TEC.": ("INFORMAÇÃO TÉCNICA", "INF. TÉC."),
    "INFORMACAO TECNICA": ("INFORMAÇÃO TÉCNICA", "INF. TÉC."),
    "NT": ("NOTA TÉCNICA", "NT"),
    "NOTA TECNICA": ("NOTA TÉCNICA", "NT"),
    "NOT.": ("NOTIFICAÇÃO", "NOT."),
    "NOTIFICACAO": ("NOTIFICAÇÃO", "NOT."),
    "OF": ("OFÍCIO", "OF"),
    "OFICIO": ("OFÍCIO", "OF"),
    "PAR. TEC.": ("PARECER TÉCNICO", "PAR. TÉC."),
    "PARECER TECNICO": ("PARECER TÉCNICO", "PAR. TÉC."),
    "REL. TEC.": ("RELATÓRIO TÉCNICO", "REL. TÉC."),
    "RELATORIO TECNICO": ("RELATÓRIO TÉCNICO", "REL. TÉC."),
    "DILIG.": ("DILIGÊNCIA", "DILIG."),
    "DILIGENCIA": ("DILIGÊNCIA", "DILIG."),
    "PAR.": ("PARECER", "PAR."),
    "PARECER": ("PARECER", "PAR."),
    "RESP. E-MAIL": ("RESPOSTA E-MAIL", "RESP. E-MAIL"),
    "ATA": ("ATA", "ATA"),
    "VIST.": ("VISTORIA", "VIST."),
    "VISTORIA": ("VISTORIA", "VIST."),
    "TEM": ("TERMO DE ENTREGA", "TEM"),
    "ET": ("ESTUDO TÉCNICO", "ET"),
    "AT. CAP. TEC.": ("ATESTADO DE CAPACIDADE TÉCNICA", "AT. CAP. TÉC."),
    "LAU. TEC.": ("LAUDO TÉCNICO", "LAU. TÉC."),
    "LAUDO TEC.": ("LAUDO TÉCNICO", "LAU. TÉC."),
}
SETORES = {"DET", "DPL", "GEAR", "GPRO", "GIMP", "GLIP", "GMOB", "GMON", "IMP"}
SETORES_ALIAS = {"GEMON": "GMON", "GESIN": "GSIN"}  # grafias antigas (2024)
REGIONAIS = {"EL": "ELDORADO", "IN": "INDUSTRIAL", "NA": "NACIONAL", "PE": "PETROLÂNDIA",
             "RE": "RESSACA", "RI": "RIACHO", "SE": "SEDE", "VF": "VARGEM DAS FLORES"}

RE_DOC = re.compile(r"^(?P<tipo>.+?)\s+(?P<setor>[A-Z]{2,6})\s*\.\s*(?P<seq>\d{1,5})\s*\.\s*(?P<ano>\d{4})$")
RE_PROJ = re.compile(r"^PROJETO\s+(?P<setor>[A-Z]{2,6})\s+(?P<reg>[A-Z]{2})\s+(?P<seq>\d{1,5})\s*\.\s*(?P<ano>\d{4})$")
RE_OS = re.compile(r"^OS\s+(?P<seq>\d{1,5})\s*\.\s*(?P<setor>[A-Z]+)\s*\.\s*(?P<contrato>.+?)\s*\.\s*(?P<ano>\d{2})$")

ABAS = {  # aba -> (família, coluna do número)
    "DOCUMENTAÇÃO": ("DOCUMENTO", "Nº DOCUMENTAÇÃO"),
    "VISTORIAS": ("DOCUMENTO", "Nº DOCUMENTAÇÃO"),
    "PROJETOS": ("PROJETO", "NÚMERO DE PROJETO"),
    "ORDEM DE SERVIÇO": ("OS", "NÚMERO DE OS"),
}


def sem_acento(s):
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def chave(s):
    return re.sub(r"\s+", " ", sem_acento(str(s)).upper()).strip()


def texto(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = re.sub(r"\s+", " ", str(v).replace("\xa0", " ")).strip()
    return s or None


def col(df, *nomes):
    k = {chave(c): c for c in df.columns}
    for n in nomes:
        if chave(n) in k:
            return k[chave(n)]
    # cabeçalho corrompido na planilha: "LOCAL DE REFERÊNOF. INT.A"
    for n in nomes:
        for kk, c in k.items():
            if kk.startswith(chave(n)[:12]):
                return c
    return None


PEND = []


def pend(reg, problema, acao, gravidade="AVISO"):
    PEND.append({"arquivo": reg.get("arquivo"), "aba": reg.get("aba"), "linha_excel": reg.get("linha_excel"),
                 "numero": reg.get("numero"), "problema": problema, "acao": acao, "gravidade": gravidade})


def ler_arquivo(caminho: Path):
    regs = []
    xls = pd.ExcelFile(caminho)
    for aba in xls.sheet_names:
        if chave(aba) not in {chave(a) for a in ABAS}:
            continue
        familia, col_num = next(v for a, v in ABAS.items() if chave(a) == chave(aba))
        df = pd.read_excel(xls, sheet_name=aba, dtype=object, keep_default_na=False, na_values=[""])  # "NA" = Nacional
        c_num = col(df, col_num)
        c = {
            "setor": col(df, "SETOR"), "data": col(df, "DATA"), "tipo": col(df, "TIPO DE DOCUMENTO"),
            "resp": col(df, "RESPONSAVEL", "RESPONSÁVEL"), "assunto": col(df, "ASSUNTO"),
            "local": col(df, "LOCAL DE REFERÊNCIA", "REFERÊNCIA"), "regional": col(df, "REGIONAIS"),
            "contrato": col(df, "CONTRATO"), "status": col(df, "STATUS"), "obs": col(df, "OBS"),
            "dimensao": col(df, "DIMENSÃO PROJ. ELABORADO (METRO LINEAR)", "1"),
        }
        g = lambda row, k: texto(row[c[k]]) if c[k] is not None else None  # noqa: E731
        for i, row in df.iterrows():
            num = texto(row[c_num])
            if num is None or num.upper().startswith("PREENCHA"):
                continue  # linha-modelo da fórmula, sem documento
            dt = pd.to_datetime(row[c["data"]], errors="coerce") if c["data"] else pd.NaT
            regs.append({
                "arquivo": caminho.name, "aba": aba, "linha_excel": i + 2, "familia": familia,
                "numero": num, "data": dt.date() if pd.notna(dt) else None,
                "setor": (g(row, "setor") or "").upper() or None, "tipo_planilha": g(row, "tipo"),
                "responsavel": g(row, "resp"), "assunto": g(row, "assunto"), "local": g(row, "local"),
                "regional": g(row, "regional"), "contrato": g(row, "contrato"), "status": g(row, "status"),
                "observacoes": g(row, "obs"), "dimensao_m": g(row, "dimensao") if familia == "PROJETO" else None,
            })
    return regs


def interpretar(r):
    """Decompõe o número e define o escopo da sequência."""
    r.update(tipo=None, sigla=None, setor_num=None, regional_sigla=None, sequencial=None, ano=None, escopo=None)
    if r["familia"] == "DOCUMENTO":
        m = RE_DOC.match(chave(r["numero"]))
        if not m:
            pend(r, "número fora do padrão 'TIPO SETOR.0000.AAAA'", "registrado sem sequência; conferir", "ERRO")
            return
        t = TIPOS.get(chave(m["tipo"]))
        if t is None:
            pend(r, f"tipo de documento desconhecido: '{m['tipo']}'", "incluir em TIPOS do script", "ERRO")
            t = (m["tipo"], m["tipo"])
        setor = SETORES_ALIAS.get(m["setor"], m["setor"])
        r.update(tipo=t[0], sigla=t[1], setor_num=setor, sequencial=int(m["seq"]), ano=int(m["ano"]))
        r["escopo"] = f"DOCUMENTO|{t[0]}|{r['ano']}"
        r["numero_padronizado"] = f"{t[1]} {setor}.{r['sequencial']:04d}.{r['ano']}"
    elif r["familia"] == "PROJETO":
        m = RE_PROJ.match(chave(r["numero"]))
        if not m:
            pend(r, "número fora do padrão 'PROJETO SETOR RG 000.AAAA'", "registrado sem sequência; conferir", "ERRO")
            return
        if m["reg"] not in REGIONAIS:
            pend(r, f"sigla de regional desconhecida: '{m['reg']}'", "conferir", "ERRO")
        r.update(tipo="PROJETO", sigla="PROJETO", setor_num=m["setor"], regional_sigla=m["reg"],
                 sequencial=int(m["seq"]), ano=int(m["ano"]))
        r["escopo"] = f"PROJETO|{m['reg']}|{r['ano']}"
        r["numero_padronizado"] = f"PROJETO {m['setor']} {m['reg']} {r['sequencial']:03d}.{r['ano']}"
    else:
        m = RE_OS.match(chave(r["numero"]))
        if not m:
            pend(r, "número fora do padrão 'OS 000.SETOR.CONTRATO.AA'", "registrado sem sequência; conferir", "ERRO")
            return
        ano = 2000 + int(m["ano"])
        r.update(tipo="ORDEM DE SERVIÇO", sigla="OS", setor_num=m["setor"], sequencial=int(m["seq"]), ano=ano,
                 contrato=m["contrato"])
        r["escopo"] = f"OS|{m['contrato']}|{ano}"
        r["numero_padronizado"] = f"OS {r['sequencial']:03d}.{m['setor']}.{m['contrato']}.{ano % 100:02d}"

    if r["setor_num"] and r["setor_num"] not in SETORES and r["setor_num"] not in SETORES_ALIAS.values():
        pend(r, f"setor desconhecido no número: '{r['setor_num']}'", "conferir se o setor existe", "AVISO")
    if r["data"] and r["ano"] and r["data"].year != r["ano"]:
        pend(r, f"ano do número ({r['ano']}) diferente do ano da data ({r['data']:%d/%m/%Y})",
             "mantido o número como emitido", "AVISO")
    ano_arq = re.search(r"(20\d\d)", r["arquivo"])
    if ano_arq and r["ano"] and int(ano_arq.group(1)) != r["ano"]:
        pend(r, f"número de {r['ano']} registrado na planilha de {ano_arq.group(1)}",
             "a fórmula contou esta linha junto com as do ano da planilha: risco de número repetido", "ERRO")


def analisar_sequencias(regs):
    por_escopo = defaultdict(list)
    for r in regs:
        if r.get("escopo"):
            por_escopo[r["escopo"]].append(r)
    seqs = []
    for esc, itens in sorted(por_escopo.items()):
        familia, chave_seq, ano = esc.split("|")
        nums = defaultdict(list)
        for r in itens:
            nums[r["sequencial"]].append(r)
        usados = sorted(nums)
        repetidos = {n: rs for n, rs in nums.items() if len(rs) > 1}
        for n, rs in repetidos.items():
            numeros_distintos = {x["numero_padronizado"] for x in rs}
            for x in rs:
                x["repetido"] = True
                pend(x, f"sequencial {n} usado {len(rs)} vezes em {chave_seq}/{ano}"
                        + (f" (números: {', '.join(sorted(numeros_distintos))})" if len(numeros_distintos) > 1 else ""),
                     "importado; confirmar nos documentos emitidos qual número cada um recebeu", "ERRO")
        lacunas = sorted(set(range(1, max(usados) + 1)) - set(usados)) if usados else []
        seqs.append({
            "familia": familia, "sequencia": chave_seq, "ano": int(ano), "qtd_registros": len(itens),
            "qtd_numeros_distintos": len(usados), "primeiro": usados[0] if usados else None,
            "ultimo": usados[-1] if usados else None, "proximo": (usados[-1] + 1) if usados else 1,
            "qtd_repetidos": len(repetidos), "qtd_lacunas": len(lacunas),
            "lacunas": ", ".join(map(str, lacunas[:50])) + (" ..." if len(lacunas) > 50 else ""),
        })
        if lacunas:
            pend({"arquivo": "", "aba": "", "linha_excel": "", "numero": f"{chave_seq}/{ano}"},
                 f"{len(lacunas)} número(s) sem registro na sequência",
                 "registrar na Central como NÃO UTILIZADO/CANCELADO, nunca reaproveitar", "AVISO")
    return pd.DataFrame(seqs)


def vincular_demandas(regs, demandas_csv: Path):
    """Procura, nas demandas importadas, as NTs/projetos citados e confere se existem no registro."""
    # a sequência é por TIPO (documentos) ou por REGIONAL (projetos): o setor não faz parte da chave
    idx = defaultdict(list)
    for r in regs:
        if r.get("sequencial") is None:
            continue
        if r["familia"] == "DOCUMENTO":
            idx[(r["tipo"], r["sequencial"], r["ano"])].append(r)
        elif r["familia"] == "PROJETO":
            idx[("PROJETO", r["regional_sigla"], r["sequencial"], r["ano"])].append(r)
    d = pd.read_csv(demandas_csv, dtype=str)
    re_ref = re.compile(r"(?P<tipo>NT|INF\.?\s*TEC(?:NICA)?\.?|OF\.?\s*INT\.?|REL\.?\s*TEC\.?|PARECER)\s*(?:N[º°O]?\s*)?"
                        r"(?P<setor>[A-Z]{3,5})\s*[.\s]\s*(?P<seq>\d{1,4})\s*[./]\s*(?P<ano>\d{4})")
    re_proj = re.compile(r"(?P<setor>G[A-Z]{2,4})\s+(?P<reg>[A-Z]{2})\s+(?P<seq>\d{1,3})\s*\.\s*(?P<ano>\d{4})")
    alias = {"NT": "NOTA TÉCNICA", "INF": "INFORMAÇÃO TÉCNICA", "OF": "OFÍCIO INTERNO", "REL": "RELATÓRIO TÉCNICO",
             "PARECER": "PARECER"}
    out = []

    def registrar(row, campo, citado, setor, achados):
        if not achados:
            sit, num = "NÃO ENCONTRADO NO REGISTRO", None
        else:
            num = " | ".join(x["numero_padronizado"] for x in achados)
            setores = {SETORES_ALIAS.get(x["setor_num"], x["setor_num"]) for x in achados}
            sit = "OK" if SETORES_ALIAS.get(setor, setor) in setores else "MESMO Nº COM OUTRO SETOR"
        out.append({"id_origem": row["id_origem"], "campo": campo, "citado": citado,
                    "numero_no_registro": num, "situacao": sit})

    for _, row in d.iterrows():
        for campo in ("nota_tecnica", "projeto"):
            txt = row.get(campo)
            if not isinstance(txt, str):
                continue
            t = chave(txt)
            for m in re_ref.finditer(t):
                tp = alias[re.match(r"[A-Z]+", m["tipo"]).group(0)]
                registrar(row, campo, m.group(0), m["setor"], idx.get((tp, int(m["seq"]), int(m["ano"]))))
            for m in re_proj.finditer(t):
                registrar(row, campo, m.group(0), m["setor"],
                          idx.get(("PROJETO", m["reg"], int(m["seq"]), int(m["ano"]))))
    return pd.DataFrame(out, columns=["id_origem", "campo", "citado", "numero_no_registro", "situacao"])


COLS = ["familia", "numero", "numero_padronizado", "tipo", "sigla", "setor_num", "regional_sigla", "contrato",
        "sequencial", "ano", "escopo", "data", "setor", "tipo_planilha", "responsavel", "assunto", "local",
        "regional", "status", "observacoes", "dimensao_m", "repetido", "arquivo", "aba", "linha_excel"]


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("planilhas", nargs="+", type=Path)
    ap.add_argument("--demandas", type=Path, help="demandas.csv gerado por importar.py (opcional)")
    ap.add_argument("--saida", type=Path, default=BASE / "saida" / "numeracao")
    a = ap.parse_args(argv)

    regs = []
    for p in sorted(a.planilhas, key=lambda p: p.name):
        lidos = ler_arquivo(p)
        print(f"  [{p.name}] {len(lidos)} números")
        regs += lidos
    for r in regs:
        r["repetido"] = False
        interpretar(r)
    seqs = analisar_sequencias(regs)

    a.saida.mkdir(parents=True, exist_ok=True)
    num = pd.DataFrame(regs).reindex(columns=COLS).sort_values(["familia", "escopo", "sequencial"], na_position="last")
    pen = pd.DataFrame(PEND, columns=["arquivo", "aba", "linha_excel", "numero", "problema", "acao", "gravidade"])
    num.to_csv(a.saida / "numeros_utilizados.csv", index=False, encoding="utf-8")
    seqs.to_csv(a.saida / "sequencias.csv", index=False, encoding="utf-8")
    pen.to_csv(a.saida / "pendencias.csv", index=False, encoding="utf-8-sig")

    vinc = None
    if a.demandas:
        vinc = vincular_demandas(regs, a.demandas)
        vinc.to_csv(a.saida / "vinculos_demandas.csv", index=False, encoding="utf-8")

    pen["tipo_problema"] = pen.problema.str.replace(r"'[^']*'|\d+|\(.*\)", "…", regex=True)
    resumo = pen.groupby(["gravidade", "tipo_problema"]).size().reset_index(name="qtd")
    with pd.ExcelWriter(a.saida / "revisao_numeracao.xlsx") as xw:
        seqs[seqs.ano == seqs.ano.max()].to_excel(xw, sheet_name="proximos_numeros", index=False)
        seqs.to_excel(xw, sheet_name="sequencias", index=False)
        resumo.to_excel(xw, sheet_name="pendencias_por_tipo", index=False)
        pen.drop(columns="tipo_problema").to_excel(xw, sheet_name="pendencias", index=False)
        num.to_excel(xw, sheet_name="numeros_utilizados", index=False)
        if vinc is not None:
            vinc.to_excel(xw, sheet_name="vinculos_demandas", index=False)

    print(f"\n  números registrados: {len(num)}  |  sequências: {len(seqs)}")
    print(resumo.to_string(index=False))
    if vinc is not None and len(vinc):
        print(f"\n  referências a documentos nas demandas: {len(vinc)}")
        print(vinc.situacao.value_counts().to_string())
    print(f"\nArquivos gravados em {a.saida}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
