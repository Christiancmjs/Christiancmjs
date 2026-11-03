#!/usr/bin/env python3
"""
Importação das planilhas "Demandas GEAR" (aba GERAL) para o sistema.

Lê uma ou mais planilhas .xlsx, padroniza colunas e valores segundo as
tabelas de domínio/de-para da pasta `mapeamentos/` e gera, em `saida/`:

  demandas.csv        uma linha por demanda (pronta para carga no banco)
  execucoes.csv       datas de execução por tipo de serviço (1:N com demandas)
  pendencias.csv      tudo o que foi corrigido, movido ou não pôde ser aproveitado
  revisao.xlsx        os três conjuntos acima + resumo, para conferência humana

Uso:
  python3 importar.py Demandas_GEAR_2024-2025.xlsx Demandas_GEAR_2026_12.xlsx
  python3 importar.py --saida outra_pasta planilha1.xlsx planilha2.xlsx

Nenhum dado é descartado em silêncio: todo valor alterado ou rejeitado é
registrado em pendencias.csv com o arquivo, a linha do Excel e o valor original.
"""
from __future__ import annotations

import argparse
import csv
import difflib
import hashlib
import re
import sys
import unicodedata
from collections import Counter, defaultdict
from datetime import date, datetime, time
from pathlib import Path

import pandas as pd

BASE = Path(__file__).resolve().parent
MAPS = BASE / "mapeamentos"

# ---------------------------------------------------------------------------
# 1. Mapeamento de colunas: nome canônico <- possíveis cabeçalhos na planilha
#    (comparação sem acento, sem espaços extras e sem diferenciar maiúsculas)
# ---------------------------------------------------------------------------
COLUNAS = {
    "data_recebimento": ["DATA RECEBIMENTO", "DATA DE RECEBIMENTO - DET"],
    "data_solicitacao": ["DATA SOLICITAÇÃO/OFÍCIO"],
    "canal_entrada": ["CANAL DE ENTRADA"],
    "remetente": ["REMETENTE"],
    "origem": ["ORIGEM", "ORIGEM DA DEMANDA"],
    "registro": ["REGISTRO"],
    "solicitante": ["SOLICITANTE"],
    "email_solicitante": ["E-MAIL SOLICITANTE"],
    "viabilidade": ["VIABILIDADE"],
    "status": ["STATUS"],
    "prioridade": ["PRIORIDADE"],
    "alteracao_prioridade_interno": ["SOLICITANTE DE ALTERAÇÃO DE PRIORIDADE - INTERNO"],
    "alteracao_prioridade_externo": ["SOLICITANTE DE ALTERAÇÃO DE PRIORIDADE - EXTERNO"],
    "data_alteracao_prioridade": ["DATA ALTERAÇÃO PRIORIDADE (REUNIÃO)"],
    "tipo_intervencao": ["TIPO DE INTERVENÇÃO"],
    "assunto": ["ASSUNTO"],
    "sub_assunto": ["SUB-ASSUNTO"],
    "descricao": ["DESCRIÇÃO DA SOLICITAÇÃO", "DESCRICAO DA SOLICITACAO"],
    "logradouro": ["RUA/AV"],
    "numero": ["NÚMERO", "Nº"],
    "referencia": ["REFERÊNCIA"],
    "bairro": ["BAIRRO"],
    "logradouro_completo": ["LOGRADOURO COMPLETO"],
    "programa_asfalto_novo": ["PROGRAMA ASFALTO NOVO"],
    "regional": ["REGIONAL"],
    "nota_tecnica": ["NOTA TÉCNICA", "NT/CI"],
    "projeto": ["PROJETO / PREVISAO", "PROJETO"],
    "medida_alternativa": ["MEDIDA ALTERNATIVA"],
    "respondido": ["RESPONDIDO"],
    "data_resposta_email": ["DATA DA RESPOSTA POR E-MAIL"],
    "data_resposta_gabinete": ["DATA RESPOSTA ENVIO GABINETE PRESIDÊNCIA"],
    "encaminhada_para": ["DEMANDA ENCAMINHADA PARA"],
    "data_encaminhamento": ["DATA ENCAMINHAMENTO"],
    "geocode": ["GEOCODE / PROTOCOLO", "GEOCODE"],
    "data_baixa_geocode": ["BAIXA GEOCODE"],
    "observacoes": ["OBSERVAÇÕES"],
    "data_execucao": ["DATA DE EXECUÇÃO DA SOLICITAÇÃO"],
    "latitude": ["LATITUDE"],
    "longitude": ["LONGITUDE"],
    "marcada_duplicada": ["DUPLICADAS"],
}
# Colunas de execução por tipo de serviço -> tabela `execucoes`
EXECUCOES = {
    "VERTICAL": "VERTICAL",
    "HORIZONTAL": "HORIZONTAL",
    "REDUTOR": "REDUTOR",
    "TRAVESSIA ELEVADA": "TRAVESSIA ELEVADA",
    "ABRIGO": "ABRIGO",
    "FAIXAS DE PANO": "FAIXAS DE PANO",
    "OBRAS FÍSICAS": "OBRAS FÍSICAS",
    "DISPOSITIVO ELETRÔNICO": "DISPOSITIVO ELETRÔNICO",
}
# ANO, MÊS, ColunaN etc. são derivadas/lixo e não são importadas.

CATEGORICOS = [
    "canal_entrada", "remetente", "origem", "viabilidade", "status", "prioridade",
    "alteracao_prioridade_interno", "alteracao_prioridade_externo", "tipo_intervencao",
    "regional", "programa_asfalto_novo", "sub_assunto",
]
DATAS = {  # campo -> guarda hora?
    "data_recebimento": False, "data_solicitacao": False, "data_alteracao_prioridade": False,
    "data_resposta_email": True, "data_resposta_gabinete": True, "data_encaminhamento": True,
    "data_baixa_geocode": False, "data_execucao": False,
}
# Textos que significam "vazio" em qualquer campo
VAZIOS = {"", "-", "--", "*", "**", "NAN", "NAT", "NONE", "N/A", "NA", "0",
          "NAO SE APLICA", "NAO CONSTA", "PENDENTE", "PEDENTE", "NENHUM", "NRNHUM",
          "CADASTRO PENDENTE", "OFICIO NAO RECEBIDO", ", ,  -", ",  -"}
CAMPOS_ONDE_NENHUM_E_VALOR = {"alteracao_prioridade_interno", "alteracao_prioridade_externo", "sub_assunto"}

# Caixa geográfica de Contagem/MG (com folga) para validar coordenadas
LAT_MIN, LAT_MAX = -20.30, -19.60
LON_MIN, LON_MAX = -44.50, -43.70
ANO_MIN, ANO_MAX = 2015, 2027

EMAIL_RE = re.compile(r"^[\w.+-]+@[\w-]+(\.[\w-]+)+$")
DATA_RE = re.compile(r"(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{2,4})(?:\D{0,6}?(\d{1,2})\s*[:;h]\s*(\d{2}))?")


# ---------------------------------------------------------------------------
# utilidades de texto
# ---------------------------------------------------------------------------
def sem_acento(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def limpar(v) -> str | None:
    """Converte qualquer célula em texto limpo (ou None)."""
    if v is None or (isinstance(v, float) and pd.isna(v)) or v is pd.NaT:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).replace("\xa0", " ").replace("\r", "")
    s = "\n".join(re.sub(r"[ \t]+", " ", ln).strip() for ln in s.split("\n")).strip()
    return s or None


def chave(s: str) -> str:
    """Forma de comparação: maiúsculas, sem acento, espaços simples."""
    return re.sub(r"\s+", " ", sem_acento(s).upper()).strip()


def eh_vazio(s: str | None, campo: str = "") -> bool:
    if s is None:
        return True
    k = chave(s)
    if k in ("NENHUM",) and campo in CAMPOS_ONDE_NENHUM_E_VALOR:
        return False
    return k in VAZIOS


# ---------------------------------------------------------------------------
# carga das tabelas de apoio
# ---------------------------------------------------------------------------
def ler_csv(nome):
    with open(MAPS / nome, encoding="utf-8", newline="") as f:
        return list(csv.DictReader(f))


DOMINIOS: dict[str, dict[str, str]] = defaultdict(dict)   # campo -> chave -> valor canônico
for r in ler_csv("dominios.csv"):
    DOMINIOS[r["campo"]][chave(r["valor"])] = r["valor"]

DEPARA: dict[tuple[str, str], dict] = {}
for r in ler_csv("valores.csv"):
    DEPARA[(r["campo"], chave(r["valor_origem"]))] = r

ASSUNTO_LEGADO = {chave(r["assunto_origem"]): r for r in ler_csv("assunto_legado.csv")}
SUB_PARA_ASSUNTO = {chave(r["sub_assunto"]): r["assunto_padrao"] for r in ler_csv("sub_assunto.csv")}


# ---------------------------------------------------------------------------
# registro de pendências
# ---------------------------------------------------------------------------
class Pendencias:
    def __init__(self):
        self.itens = []

    def add(self, linha_ctx, campo, valor, problema, acao, gravidade="AVISO"):
        self.itens.append({
            "id_origem": linha_ctx["id_origem"], "arquivo": linha_ctx["arquivo"],
            "linha_excel": linha_ctx["linha_excel"], "campo": campo,
            "valor_original": valor, "problema": problema, "acao": acao,
            "gravidade": gravidade,
        })


PEND = Pendencias()


# ---------------------------------------------------------------------------
# conversores
# ---------------------------------------------------------------------------
def categorizar(campo, bruto, ctx, rec):
    """Padroniza um valor categórico. Retorna o valor canônico ou None."""
    s = limpar(bruto)
    if eh_vazio(s, campo):
        return None
    # valores com quebra de linha: usa a primeira linha (ex.: "SIM\nVICE-PRESIDENTE")
    k = chave(s)
    dom = DOMINIOS.get(campo, {})
    if k in dom:
        return dom[k]
    regra = DEPARA.get((campo, k))
    if regra is None and "\n" in s:
        primeiro = chave(s.split("\n")[0])
        if primeiro in dom:
            PEND.add(ctx, campo, s, "mais de um valor na célula", f"usado '{dom[primeiro]}'")
            return dom[primeiro]
    if regra is not None:
        destino = regra["valor_destino"] or None
        if regra["mover_para"]:
            alvo = regra["mover_para"]
            if not rec.get(alvo):
                rec[alvo] = s if alvo == "motivo_prioridade" else categorizar(alvo, s, ctx, rec)
        if regra["revisar"] == "S":
            PEND.add(ctx, campo, s, "valor legado convertido por regra de-para",
                     f"-> '{destino or ''}'" + (f"; movido para {regra['mover_para']}" if regra["mover_para"] else ""))
        return destino
    if dom:
        parecido = difflib.get_close_matches(k, dom.keys(), n=1, cutoff=0.86)
        if parecido:
            v = dom[parecido[0]]
            PEND.add(ctx, campo, s, "grafia diferente da lista", f"corrigido automaticamente para '{v}'")
            return v
        PEND.add(ctx, campo, s, "valor fora da lista de domínio",
                 "campo deixado vazio; incluir em mapeamentos/valores.csv", "ERRO")
        return None
    return s.upper()


def tratar_assunto(bruto_assunto, bruto_sub, ctx, rec):
    """Converte ASSUNTO (taxonomia nova ou legada com sufixo _IM/_MA)."""
    sub = categorizar("sub_assunto", bruto_sub, ctx, rec)
    s = limpar(bruto_assunto)
    dom = DOMINIOS["assunto"]
    assunto = None
    if not eh_vazio(s):
        rec["assunto_original"] = s
        k = chave(s)
        sufixo = re.search(r"_(IM|MA)$", k)
        if sufixo:
            k = k[: sufixo.start()].strip()
            tipo = "IMPLANTAÇÃO" if sufixo.group(1) == "IM" else "MANUTENÇÃO"
            if not rec.get("tipo_intervencao"):
                rec["tipo_intervencao"] = tipo
        if k in dom:
            assunto = dom[k]
        elif k in ASSUNTO_LEGADO:
            r = ASSUNTO_LEGADO[k]
            assunto = r["assunto"]
            if not sub and r["sub_assunto"]:
                sub = r["sub_assunto"]
        else:
            # combinações ("VERTICAL E HORIZONTAL", "OBRAS E HORIZONTAL"): usa a 1ª parte
            partes = [p.strip() for p in re.split(r"\s+E\s+|,|/", k) if p.strip()]
            for p in partes:
                if p in dom or p in ASSUNTO_LEGADO:
                    assunto = dom.get(p) or ASSUNTO_LEGADO[p]["assunto"]
                    PEND.add(ctx, "assunto", s, "mais de um assunto na célula",
                             f"usado '{assunto}'; demais citados no texto original")
                    break
            if assunto is None:
                cands = list(dom.keys()) + list(ASSUNTO_LEGADO.keys())
                parecido = difflib.get_close_matches(k, cands, n=1, cutoff=0.86)
                if parecido:
                    p = parecido[0]
                    assunto = dom.get(p) or ASSUNTO_LEGADO[p]["assunto"]
                    PEND.add(ctx, "assunto", s, "grafia diferente da lista",
                             f"corrigido automaticamente para '{assunto}'")
                else:
                    PEND.add(ctx, "assunto", s, "assunto não reconhecido",
                             "deduzido do sub-assunto ou deixado vazio; incluir em mapeamentos/assunto_legado.csv",
                             "ERRO")
    if assunto is None and sub and chave(sub) in SUB_PARA_ASSUNTO:
        assunto = SUB_PARA_ASSUNTO[chave(sub)]
    rec["assunto"] = assunto
    rec["sub_assunto"] = sub


def para_data(campo, bruto, ctx, rec, com_hora=False):
    """Aceita datetime do Excel, 'dd/mm/aaaa', 'dd/mm/aaaa - hh:mm', 'ddmmaaaa'."""
    if bruto is None or (isinstance(bruto, float) and pd.isna(bruto)) or bruto is pd.NaT:
        return None
    if isinstance(bruto, (pd.Timestamp, datetime, date)):
        dt = pd.Timestamp(bruto).to_pydatetime()
        if not (ANO_MIN <= dt.year <= ANO_MAX):
            PEND.add(ctx, campo, str(bruto), f"ano fora do intervalo {ANO_MIN}-{ANO_MAX}", "campo deixado vazio", "ERRO")
            return None
        return dt if com_hora and dt.time() != time(0) else dt.date()
    s = limpar(bruto)
    if eh_vazio(s, campo):
        return None
    achados = DATA_RE.findall(s)
    if not achados and re.fullmatch(r"\d{8}", s):
        if s.startswith("20"):  # 20260811 (aaaammdd)
            achados = [(s[6:], s[4:6], s[:4], "", "")]
        else:  # 13012026 (ddmmaaaa)
            achados = [(s[:2], s[2:4], s[4:], "", "")]
    if not achados:
        PEND.add(ctx, campo, s, "texto em campo de data", "campo vazio; texto copiado para observações")
        anexar_obs(rec, f"{campo}: {s}")
        return None
    d, m, a, hh, mm = achados[0]
    a = int(a) + (2000 if len(a) == 2 else 0)
    try:
        dt = datetime(a, int(m), int(d), int(hh or 0), int(mm or 0))
    except ValueError:
        PEND.add(ctx, campo, s, "data inválida", "campo deixado vazio", "ERRO")
        return None
    if not (ANO_MIN <= dt.year <= ANO_MAX):
        PEND.add(ctx, campo, s, f"ano fora do intervalo {ANO_MIN}-{ANO_MAX}", "campo deixado vazio", "ERRO")
        return None
    if len(achados) > 1 or re.sub(DATA_RE, "", s).strip(" -–:\n"):
        PEND.add(ctx, campo, s, "célula com mais de uma data ou texto adicional",
                 "usada a primeira data; texto completo copiado para observações")
        anexar_obs(rec, f"{campo}: {s}")
    return dt if com_hora and (hh or mm) else dt.date()


def anexar_obs(rec, texto):
    rec["observacoes"] = (rec.get("observacoes") + "\n" if rec.get("observacoes") else "") + f"[IMPORTAÇÃO] {texto}"


def numero_para_digitos(v) -> str | None:
    if isinstance(v, int):
        return str(abs(v))
    if isinstance(v, float) and not pd.isna(v):
        return f"{abs(v):.0f}"
    s = limpar(v)
    if s and re.fullmatch(r"-?[\d.,]+", s):
        return re.sub(r"\D", "", s)
    return None


def para_coordenada(v):
    """Corrige coordenadas que perderam a vírgula decimal (ex.: -4408818745798690
    -> -44.08818745798690). Retorna float negativo ou None."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)) and abs(v) <= 180:
        return float(v)
    dig = numero_para_digitos(v)
    if not dig or len(dig) < 4:
        return None
    sinal = -1 if (isinstance(v, (int, float)) and v < 0) or str(v).strip().startswith("-") else 1
    return sinal * float(dig[:2] + "." + dig[2:])


def tratar_coordenadas(lat_b, lon_b, ctx, rec):
    a, b = para_coordenada(lat_b), para_coordenada(lon_b)
    if a is None and b is None:
        return
    lat = lon = None
    for v in (a, b):
        if v is None:
            continue
        if LAT_MIN <= v <= LAT_MAX:
            lat = v
        elif LON_MIN <= v <= LON_MAX:
            lon = v
    if lat is None or lon is None:
        PEND.add(ctx, "latitude/longitude", f"{lat_b} | {lon_b}",
                 "coordenada fora de Contagem ou ilegível", "coordenadas deixadas vazias", "ERRO")
        return
    if a != lat:  # colunas trocadas na planilha
        PEND.add(ctx, "latitude/longitude", f"{lat_b} | {lon_b}",
                 "latitude e longitude invertidas / sem separador decimal", "corrigido automaticamente", "INFO")
    rec["latitude"], rec["longitude"] = round(lat, 7), round(lon, 7)


def para_bool(campo, bruto, ctx, rec):
    s = limpar(bruto)
    if s is None:
        return None
    regra = DEPARA.get((campo, chave(s)))
    if regra:
        s = regra["valor_destino"]
        if not s:
            return None
    k = chave(s)
    if k in ("SIM", "S", "X"):
        return True
    if k in ("NAO", "N"):
        return False
    PEND.add(ctx, campo, s, "valor não é SIM/NÃO", "campo deixado vazio", "ERRO")
    return None


# ---------------------------------------------------------------------------
# leitura de uma planilha
# ---------------------------------------------------------------------------
def localizar_colunas(df: pd.DataFrame):
    por_chave = {chave(str(c)): c for c in df.columns}
    mapa, faltando = {}, []
    for canon, nomes in COLUNAS.items():
        col = next((por_chave[chave(n)] for n in nomes if chave(n) in por_chave), None)
        if col is None:
            faltando.append(canon)
        mapa[canon] = col
    exec_cols = {tipo: por_chave.get(chave(c)) for c, tipo in EXECUCOES.items()}
    return mapa, exec_cols, faltando


def processar_planilha(caminho: Path, aba="GERAL"):
    tag = re.sub(r"[^A-Za-z0-9]+", "_", caminho.stem).strip("_")
    df = pd.read_excel(caminho, sheet_name=aba, dtype=object)
    mapa, exec_cols, faltando = localizar_colunas(df)
    if faltando:
        print(f"  [{caminho.name}] colunas ausentes (ficarão vazias): {', '.join(faltando)}")
    demandas, execs, ignoradas = [], [], 0

    for idx, row in df.iterrows():
        linha_excel = idx + 2  # +1 cabeçalho, +1 base 1
        get = lambda c: row[mapa[c]] if mapa.get(c) is not None else None  # noqa: E731
        # linha vazia (ou só com valores-padrão de fórmulas/validação)?
        essenciais = ["data_recebimento", "descricao", "solicitante", "registro", "logradouro", "assunto"]
        if all(eh_vazio(limpar(get(c))) for c in essenciais):
            ignoradas += 1
            continue

        ctx = {"id_origem": f"{tag}:L{linha_excel}", "arquivo": caminho.name, "linha_excel": linha_excel}
        rec = {"id_origem": ctx["id_origem"], "arquivo_origem": caminho.name, "linha_excel": linha_excel,
               "observacoes": limpar(get("observacoes"))}

        # textos livres
        for c in ["registro", "solicitante", "descricao", "logradouro", "referencia", "bairro",
                  "nota_tecnica", "projeto", "medida_alternativa", "geocode", "encaminhada_para"]:
            s = limpar(get(c))
            rec[c] = None if eh_vazio(s, c) else s
        for c in ["solicitante", "logradouro", "bairro", "geocode", "encaminhada_para"]:
            if rec[c]:
                rec[c] = rec[c].upper()

        # número (2024-2025 trazia o nº na coluna REFERÊNCIA)
        num = limpar(get("numero"))
        rec["numero"] = None if eh_vazio(num) else num
        if not rec["numero"] and rec["referencia"] and re.fullmatch(r"\d{1,6}", rec["referencia"]):
            rec["numero"], rec["referencia"] = rec["referencia"], None
        if rec["numero"] and chave(rec["numero"]) in ("SN", "S/N", "S/Nº", "S N"):
            rec["numero"] = "S/N"

        # e-mail
        em = limpar(get("email_solicitante"))
        if not eh_vazio(em):
            em = em.lower().replace(" ", "")
            if EMAIL_RE.match(em):
                rec["email_solicitante"] = em
            else:
                rec["email_solicitante"] = None
                PEND.add(ctx, "email_solicitante", em, "e-mail inválido", "campo deixado vazio")
        # SOLICITANTE preenchido com e-mail (comum em 2024-2025)
        if rec["solicitante"] and EMAIL_RE.match(rec["solicitante"].lower()):
            if not rec.get("email_solicitante"):
                rec["email_solicitante"] = rec["solicitante"].lower()
            PEND.add(ctx, "solicitante", rec["solicitante"], "e-mail no campo SOLICITANTE",
                     "copiado para email_solicitante; nome do solicitante a completar", "INFO")

        # "DEMANDA ENCAMINHADA PARA" às vezes contém uma data
        enc_bruto = get("encaminhada_para")
        if isinstance(enc_bruto, (datetime, pd.Timestamp)):
            rec["encaminhada_para"] = None
            if eh_vazio(limpar(get("data_encaminhamento"))):
                rec["data_encaminhamento"] = para_data("data_encaminhamento", enc_bruto, ctx, rec, True)
            PEND.add(ctx, "encaminhada_para", str(enc_bruto), "data no campo de destino",
                     "movida para data_encaminhamento")

        # categóricos (ordem importa: origem/canal podem mover valores entre si)
        rec["motivo_prioridade"] = None
        for c in ["canal_entrada", "remetente", "origem", "viabilidade", "status", "prioridade",
                  "alteracao_prioridade_interno", "alteracao_prioridade_externo", "tipo_intervencao",
                  "regional", "programa_asfalto_novo"]:
            v = categorizar(c, get(c), ctx, rec)
            if v is not None or not rec.get(c):
                rec[c] = v
        tratar_assunto(get("assunto"), get("sub_assunto"), ctx, rec)

        # datas
        for c, com_hora in DATAS.items():
            if rec.get(c) is None:
                rec[c] = para_data(c, get(c), ctx, rec, com_hora)
        if rec["data_recebimento"] is None:
            PEND.add(ctx, "data_recebimento", limpar(get("data_recebimento")),
                     "demanda sem data de recebimento", "importada sem data; completar no sistema", "ERRO")

        rec["respondido"] = para_bool("respondido", get("respondido"), ctx, rec)
        rec["marcada_duplicada"] = bool(limpar(get("marcada_duplicada")))
        tratar_coordenadas(get("latitude"), get("longitude"), ctx, rec)

        # logradouro completo: recompõe quando vier vazio/lixo (", ,  - ")
        lc = limpar(get("logradouro_completo"))
        if lc and re.sub(r"[\s,\-]", "", lc):
            rec["logradouro_completo"] = lc.upper()
        else:
            partes = ", ".join(p for p in [rec["logradouro"], rec["numero"], rec["referencia"]] if p)
            rec["logradouro_completo"] = (partes + (f" - {rec['bairro']}" if rec["bairro"] else "")) or None

        demandas.append(rec)

        # execuções (1:N)
        for tipo, col in exec_cols.items():
            if col is None:
                continue
            d = para_data(f"execucao_{tipo}", row[col], ctx, rec)
            if d:
                execs.append({"id_origem": rec["id_origem"], "tipo_servico": tipo, "data_execucao": d})

    print(f"  [{caminho.name}] {len(demandas)} demandas, {len(execs)} execuções, {ignoradas} linhas vazias ignoradas")
    return demandas, execs, ignoradas


# ---------------------------------------------------------------------------
# duplicidades (não remove; apenas sinaliza para decisão humana)
# ---------------------------------------------------------------------------
RANK_STATUS = {"EXECUTADO": 4, "EXECUTADO COM CONDICIONANTE": 4, "FINALIZADO": 4, "INVIÁVEL": 4,
               "EXECUTADO PARCIAL": 3, "AGUARDA CRONOGRAMA": 2, "DIRECIONADO": 2, "EM ANÁLISE": 1}
CHAVE_DUP = ["data_recebimento", "solicitante", "descricao", "logradouro_completo", "assunto"]
NAO_MESCLAR = {"id_origem", "arquivo_origem", "linha_excel", "hash_conteudo", "ids_consolidados"}
ANEXAR_DIVERGENTES = ["geocode", "nota_tecnica", "projeto", "registro"]


def consolidar_duplicatas(demandas, execs):
    """Demandas com conteúdo idêntico (data, solicitante, descrição, endereço, assunto) viram UMA só.

    Fica a cópia mais completa (status mais avançado > mais campos preenchidos > a mais recente).
    Campos vazios dela são completados com os das cópias; protocolos/NTs diferentes vão para
    observações; execuções são transferidas. As cópias removidas vão para duplicatas_removidas.csv."""
    grupos = defaultdict(list)
    for ordem, r in enumerate(demandas):
        base = "|".join(chave(str(r.get(c) or "")) for c in CHAVE_DUP)
        r["hash_conteudo"] = hashlib.sha1(base.encode()).hexdigest()
        r["_ordem"] = ordem
        grupos[r["hash_conteudo"]].append(r)

    mantidas, removidas, destino = [], [], {}
    for rs in grupos.values():
        if len(rs) == 1:
            rs[0]["ids_consolidados"] = None
            mantidas.append(rs[0])
            continue
        fica = max(rs, key=lambda r: (RANK_STATUS.get(r.get("status"), 0),
                                      sum(v not in (None, "", False) for v in r.values()), r["_ordem"]))
        outras = [r for r in rs if r is not fica]
        for o in outras:
            for k, v in o.items():
                if k in NAO_MESCLAR or k.startswith("_") or v in (None, ""):
                    continue
                if fica.get(k) in (None, ""):
                    fica[k] = v
                elif k in ANEXAR_DIVERGENTES and chave(str(v)) != chave(str(fica[k])):
                    anexar_obs(fica, f"{k} da linha duplicada {o['id_origem']}: {v}")
                elif k in ("status", "viabilidade", "respondido") and v != fica[k]:
                    PEND.add({"id_origem": fica["id_origem"], "arquivo": fica["arquivo_origem"],
                              "linha_excel": fica["linha_excel"]}, k, str(v),
                             f"cópia duplicada {o['id_origem']} tinha {k} diferente",
                             f"mantido '{fica[k]}' (cópia mais avançada/completa)")
            o["consolidada_em"] = fica["id_origem"]
            destino[o["id_origem"]] = fica["id_origem"]
            removidas.append(o)
            PEND.add({"id_origem": o["id_origem"], "arquivo": o["arquivo_origem"], "linha_excel": o["linha_excel"]},
                     "*", "", "demanda duplicada (data, solicitante, descrição, endereço e assunto idênticos)",
                     f"não carregada; consolidada em {fica['id_origem']}", "INFO")
        fica["ids_consolidados"] = "; ".join(o["id_origem"] for o in outras)
        mantidas.append(fica)

    mantidas.sort(key=lambda r: r["_ordem"])
    vistos, novas_execs = set(), []
    for e in execs:
        e = dict(e, id_origem=destino.get(e["id_origem"], e["id_origem"]))
        k = (e["id_origem"], e["tipo_servico"], e["data_execucao"])
        if k not in vistos:
            vistos.add(k)
            novas_execs.append(e)
    return mantidas, removidas, novas_execs


def chave_endereco(r):
    num = (r.get("numero") or "").upper()
    if not r.get("logradouro") or not num or num in ("S/N", "0"):
        return None
    digitos = re.sub(r"\D", "", num)
    return f"{chave(r['logradouro'])}|{digitos}|{chave(r.get('bairro') or '')}"


def reaproveitar_coordenadas(demandas):
    """Demanda sem coordenada recebe a de outra demanda no MESMO endereço (rua + nº + bairro)."""
    pontos = defaultdict(list)
    for r in demandas:
        r["precisao_coordenada"] = "PLANILHA" if r.get("latitude") is not None else None
        k = chave_endereco(r)
        if k and r.get("latitude") is not None:
            pontos[k].append((r["latitude"], r["longitude"]))
    n = 0
    for r in demandas:
        if r.get("latitude") is not None:
            continue
        k = chave_endereco(r)
        if k in pontos:
            lats, lons = zip(*pontos[k])
            r["latitude"] = round(sorted(lats)[len(lats) // 2], 7)
            r["longitude"] = round(sorted(lons)[len(lons) // 2], 7)
            r["precisao_coordenada"] = "MESMO ENDEREÇO DE OUTRA DEMANDA"
            n += 1
    return n


# ---------------------------------------------------------------------------
ORDEM_SAIDA = [
    "id_origem", "arquivo_origem", "linha_excel", "data_recebimento", "data_solicitacao", "canal_entrada",
    "remetente", "origem", "registro", "solicitante", "email_solicitante", "viabilidade", "status",
    "prioridade", "motivo_prioridade", "alteracao_prioridade_interno", "alteracao_prioridade_externo",
    "data_alteracao_prioridade", "tipo_intervencao", "assunto", "sub_assunto", "assunto_original",
    "descricao", "logradouro", "numero", "referencia", "bairro", "logradouro_completo", "regional",
    "latitude", "longitude", "precisao_coordenada", "programa_asfalto_novo", "nota_tecnica", "projeto", "medida_alternativa",
    "respondido", "data_resposta_email", "data_resposta_gabinete", "encaminhada_para",
    "data_encaminhamento", "geocode", "data_baixa_geocode", "data_execucao", "observacoes",
    "marcada_duplicada", "ids_consolidados", "hash_conteudo",
]


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("planilhas", nargs="+", type=Path)
    ap.add_argument("--aba", default="GERAL")
    ap.add_argument("--saida", type=Path, default=BASE / "saida")
    a = ap.parse_args(argv)

    todas_d, todas_e, total_ign = [], [], 0
    # ordem cronológica: o nome com o ano mais antigo primeiro (a 1ª ocorrência vira a "original")
    def ano(p):
        m = re.search(r"(20\d\d)", p.name)
        return (m.group(1) if m else "9999", p.name)
    for p in sorted(a.planilhas, key=ano):
        d, e, ign = processar_planilha(p, a.aba)
        todas_d += d
        todas_e += e
        total_ign += ign
    todas_d, removidas, todas_e = consolidar_duplicatas(todas_d, todas_e)
    n_coord = reaproveitar_coordenadas(todas_d)

    a.saida.mkdir(parents=True, exist_ok=True)
    dem = pd.DataFrame(todas_d).reindex(columns=ORDEM_SAIDA)
    exe = pd.DataFrame(todas_e, columns=["id_origem", "tipo_servico", "data_execucao"])
    pen = pd.DataFrame(PEND.itens, columns=["id_origem", "arquivo", "linha_excel", "campo", "valor_original",
                                            "problema", "acao", "gravidade"])

    dem.to_csv(a.saida / "demandas.csv", index=False, encoding="utf-8")
    rem = pd.DataFrame(removidas).reindex(columns=ORDEM_SAIDA[:-2] + ["consolidada_em"])
    rem.to_csv(a.saida / "duplicatas_removidas.csv", index=False, encoding="utf-8-sig")
    exe.to_csv(a.saida / "execucoes.csv", index=False, encoding="utf-8")
    pen.to_csv(a.saida / "pendencias.csv", index=False, encoding="utf-8-sig")

    resumo = pd.concat([
        pd.DataFrame({"indicador": ["demandas importadas", "execuções", "linhas vazias ignoradas",
                                    "duplicatas consolidadas (não carregadas)",
                                    "demandas com coordenada", "  da planilha", "  reaproveitada do mesmo endereço",
                                    "demandas SEM coordenada",
                                    "pendências ERRO", "pendências AVISO", "pendências INFO"],
                      "valor": [len(dem), len(exe), total_ign, len(rem),
                                dem.latitude.notna().sum(), (dem.precisao_coordenada == "PLANILHA").sum(), n_coord,
                                dem.latitude.isna().sum(),
                                (pen.gravidade == "ERRO").sum(), (pen.gravidade == "AVISO").sum(),
                                (pen.gravidade == "INFO").sum()]}),
    ])
    por_problema = pen.groupby(["gravidade", "campo", "problema"]).size().reset_index(name="qtd") \
        .sort_values(["gravidade", "qtd"], ascending=[True, False])
    with pd.ExcelWriter(a.saida / "revisao.xlsx") as xw:
        resumo.to_excel(xw, sheet_name="resumo", index=False)
        por_problema.to_excel(xw, sheet_name="pendencias_por_tipo", index=False)
        pen.to_excel(xw, sheet_name="pendencias", index=False)
        dem.to_excel(xw, sheet_name="demandas", index=False)
        exe.to_excel(xw, sheet_name="execucoes", index=False)
        rem.to_excel(xw, sheet_name="duplicatas_removidas", index=False)

    print("\nResumo:")
    for _, r in resumo.iterrows():
        print(f"  {r.indicador:40} {r.valor}")
    print(f"\nArquivos gravados em {a.saida}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
