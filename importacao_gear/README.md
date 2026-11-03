# Importação das planilhas "Demandas GEAR" (aba GERAL)

Este guia explica como tratar as planilhas **Demandas_GEAR_2024-2025.xlsx** e
**Demandas_GEAR_2026_12.xlsx** para que as demandas entrem no sistema já
padronizadas. Tudo é feito por um script que pode ser executado de novo sem
duplicar registros.

## 1. O que foi encontrado nas planilhas

| | 2024-2025 | 2026 |
|---|---|---|
| Linhas com dados na aba GERAL | 10.967 | 3.340 (+116 linhas vazias que só têm valores-padrão) |
| Período (data de recebimento) | 30/05/2023 a 31/12/2025 | 09/04/2025 a 20/10/2026 |
| Sobreposição entre os arquivos | nenhuma (nenhum REGISTRO ou GEOCODE em comum) | |

Os dois arquivos **não têm o mesmo layout** e o de 2024-2025 usa uma
classificação antiga. As principais diferenças:

| Problema | Exemplo | Tratamento |
|---|---|---|
| Nomes de coluna diferentes | `DATA DE RECEBIMENTO - DET` × `DATA RECEBIMENTO`; `Nº` × `NÚMERO`; `NT/CI` × `NOTA TÉCNICA`; `GEOCODE` × `GEOCODE / PROTOCOLO` | os dois nomes vão para o mesmo campo (`COLUNAS` em `importar.py`) |
| ASSUNTO antigo (161 valores) com sufixo `_IM`/`_MA` | `REDUTOR DE VELOCIDADE_IM` | → assunto `DISPOSITIVOS DE VELOCIDADE`, sub-assunto `REDUTOR DE VELOCIDADE (QUEBRA-MOLAS)` e tipo `IMPLANTAÇÃO` (`mapeamentos/assunto_legado.csv`) |
| PRIORIDADE com o **motivo** no lugar do nível | `ESCOLA`, `DOT`, `GABINETE`, `VAGA PCD`… | → `ALTA`; o texto original vai para `motivo_prioridade` |
| Valores que não existem mais nas listas | STATUS `NÃO ANALISADO`, `EM TRATATIVA`; VIABILIDADE `DIRECIONADO` | convertidos pela tabela de-para e listados para revisão |
| Canal de entrada digitado em ORIGEM | ORIGEM = `E-MAIL` (1.390 linhas) | movido para `canal_entrada`; a origem fica vazia |
| Erros de digitação, acentos e espaços | `GABIBETE PRESIDÊNCIA`, `CAMARA MUNICIPAL`, `'SIM '`, `WHATSAAP` | corrigidos automaticamente pela lista de domínio da aba REFERÊNCIAS |
| Coordenadas sem vírgula decimal e **com latitude e longitude trocadas** | LATITUDE `-4408818745798690` | → longitude `-44.0881875`, latitude `-19.9223031` (5.502 corrigidas, 24 fora de Contagem ficam vazias) |
| Datas como texto, ou várias na mesma célula | `23/03/2026 - 08:31`, `09/01/2026\n13/03/2026`, `20260811`, `PENDENTE` | convertidas; com várias datas usa-se a primeira e o texto completo vai para observações |
| Nº do imóvel na coluna REFERÊNCIA | REFERÊNCIA = `118` | movido para `numero` |
| E-mail na coluna SOLICITANTE | `IARA.COURA@CONTAGEM.MG.GOV.BR` | copiado para `email_solicitante` |
| Colunas de execução (VERTICAL, HORIZONTAL, REDUTOR…) | uma coluna de data por tipo de serviço | viram a tabela `execucoes` (uma linha por serviço executado) |
| Colunas derivadas ou sem uso | `ANO`, `MÊS`, `Coluna1…32` | não importadas |
| GEOCODE repetido | mesmo protocolo em várias linhas | **não** é chave única; cada linha continua sendo uma demanda |

## 2. Passo a passo

```bash
pip install -r requirements.txt
python3 importar.py Demandas_GEAR_2024-2025.xlsx Demandas_GEAR_2026_12.xlsx
```

São gerados em `saida/`:

| Arquivo | Conteúdo |
|---|---|
| `demandas.csv` | uma linha por demanda, já padronizada e sem duplicatas (14.111 linhas) |
| `execucoes.csv` | datas de execução por tipo de serviço (1.333 linhas) |
| `duplicatas_removidas.csv` | as 196 cópias que não entram no sistema e em qual demanda cada uma foi consolidada |
| `pendencias.csv` | cada valor corrigido, movido ou descartado, com arquivo, linha do Excel, valor original e ação tomada |
| `revisao.xlsx` | resumo, pendências agrupadas por tipo, pendências detalhadas, demandas e execuções |

**Antes de carregar no sistema, revise `revisao.xlsx`:**

1. Aba `pendencias_por_tipo`: os itens **ERRO** (54 na execução atual) não
   puderam ser convertidos. Exemplos: bairro digitado em ASSUNTO, datas como
   `29/27/2026`, coordenadas fora do município. Corrija na planilha de origem
   ou acrescente uma regra em `mapeamentos/`, e rode o script de novo.
2. Itens **AVISO** com "valor legado convertido por regra de-para": confira se
   as regras marcadas com `revisar = S` em `mapeamentos/valores.csv` refletem a
   decisão da gerência (por exemplo: STATUS `NÃO ANALISADO` → `EM ANÁLISE`).
3. **Duplicatas não são carregadas.** Quando duas ou mais linhas têm data,
   solicitante, descrição, endereço e assunto idênticos, **só uma entra no
   sistema** (196 cópias consolidadas). Fica a cópia com o status mais
   avançado; no empate, a mais completa e, depois, a mais recente. Os campos
   vazios dela são completados com os das outras cópias. Protocolos, NTs ou
   registros diferentes são anotados em `observacoes`, e as execuções são
   transferidas para ela. A coluna `ids_consolidados` lista as linhas fundidas,
   e as cópias removidas ficam em `duplicatas_removidas.csv` (aba
   `duplicatas_removidas`) para consulta. A coluna `marcada_duplicada` mantém o
   "X" que já existia na planilha.

Depois, carregue no banco. O arquivo `schema.sql` traz o modelo para
PostgreSQL e os comandos `\copy` comentados no fim. Esse fluxo foi testado com
os dois arquivos: 14.111 demandas e 1.333 execuções carregadas, sem nenhum
valor fora das listas de domínio.

```bash
psql -d banco_do_sistema -f schema.sql
```

Se o sistema já tiver as próprias tabelas, use `demandas.csv` e
`execucoes.csv` como origem e associe as colunas, que estão nomeadas de forma
explícita.

## 3. Como ajustar as regras (sem mexer no código)

Todas as decisões de negócio estão em CSVs editáveis no Excel:

- `mapeamentos/dominios.csv`: valores válidos de cada campo. É a lista da aba
  REFERÊNCIAS mais valores que já estão em uso em 2026, como `WHATSAPP`,
  `REGIONAL`, `ORGÃO PÚBLICO` e `INFORMAÇÃO`.
- `mapeamentos/valores.csv`: de-para de valores antigos ou errados. Colunas:
  `campo, valor_origem, valor_destino, mover_para, revisar, observacao`.
  Com `valor_destino` vazio, o campo fica vazio. `mover_para` copia o valor
  original para outro campo, se esse campo estiver vazio.
- `mapeamentos/assunto_legado.csv`: ASSUNTO antigo (sem o sufixo `_IM`/`_MA`)
  → assunto e sub-assunto novos.
- `mapeamentos/sub_assunto.csv`: assunto-padrão de cada sub-assunto, usado
  quando a demanda só tem o sub-assunto.

Valores com diferença só de acento, maiúscula ou espaço são reconhecidos
sozinhos. Grafias muito parecidas (similaridade ≥ 86%) são corrigidas e
registradas como AVISO.

## 4. Planilhas futuras

- `id_origem` (`arquivo:linha`) identifica cada linha. Reimportar o mesmo
  arquivo não duplica registros (`ON CONFLICT (id_origem) DO NOTHING`).
  **Não reordene nem apague linhas** de uma planilha que já foi importada, pois
  os ids mudariam. Acrescente linhas sempre no final.
- Para evitar retrabalho, mantenha a validação de dados (listas suspensas da
  aba REFERÊNCIAS) em todas as colunas categóricas. Use um único formato de
  data (célula do tipo data, não texto) e deixe LATITUDE/LONGITUDE como número
  com vírgula decimal.
- Quando o sistema estiver no ar, o ideal é que as demandas novas sejam
  cadastradas direto nele e que a planilha deixe de ser a fonte.

## 4a. Coordenadas (mapa)

**Nem toda demanda tem coordenada.** Situação após a importação:

| | Demandas |
|---|---|
| Coordenada vinda da planilha (2024-2025, corrigida) | 5.409 |
| Reaproveitada de outra demanda no mesmo endereço (rua + nº + bairro) | 363 |
| **Sem coordenada** | **8.339** (5.168 de 2024-2025 e **todas** as 3.171 de 2026) |

O campo `precisao_coordenada` informa a origem de cada ponto. Das 8.339 sem
coordenada, 2.474 têm rua, número e bairro (geocodificação precisa), 2.725 têm
rua e bairro sem número (ponto aproximado ou cruzamento, a partir da
referência) e 2.996 não têm rua separada, mas quase todas trazem o endereço em
`logradouro_completo`.

Para completar o mapa é preciso **geocodificar** esses endereços com um
serviço externo. Opções, da mais precisa para a menos:

1. a base de logradouros e lotes da Prefeitura (geoprocessamento), se houver;
2. uma API de geocodificação (Google, HERE ou Mapbox), que exige chave e tem
   custo baixo para esse volume;
3. Nominatim/OpenStreetMap, gratuito mas mais lento e menos preciso em
   endereços de periferia.

Toda coordenada obtida assim deve ser gravada com `precisao_coordenada`
(`ENDEREÇO EXATO`, `RUA`, `BAIRRO`) e revisada no mapa pela equipe.

## 5. Central de Numeração (planilhas de numeração DET-DPL e Ordens de Serviço)

O objetivo é registrar **todo número já usado**, para que a Central continue
cada sequência a partir do último número e nunca repita um número já emitido.

```bash
python3 importar_numeracao.py --demandas saida/demandas.csv \
    "NUMERAÇÃO DE DOCUMENTAÇÃO 2024 - DET-DPL.xlsx" "NUMERAÇÃO DE DOCUMENTAÇÃO 2025 - DET-DPL.xlsx" \
    "NUMERAÇÃO DE DOCUMENTAÇÃO 2026 - DET-DPL.xlsx" \
    "NUMERAÇÃO DE ORDEM DE SERVIÇO 2025.xlsx" "NUMERAÇÃO DE ORDEM DE SERVIÇO 2026.xlsx"
```

O script lê as abas DOCUMENTAÇÃO, VISTORIAS, PROJETOS e ORDEM DE SERVIÇO e
grava os resultados em `saida/numeracao/`. Na execução atual foram 14.054
números em 69 sequências.

### Como as planilhas numeram (e por que isso importa)

Nenhum número foi digitado: todos são **fórmulas** (`COUNTIFS`) que contam as
linhas acima com o mesmo tipo, regional ou contrato.

| Família | Exemplo | Contador |
|---|---|---|
| Documento | `NT GEAR.1052.2026` | por **tipo de documento**, compartilhado entre setores: a NT 1051 é do GEAR, a 1045 é do GMON |
| Projeto | `PROJETO GEAR RE 002.2026` | por **regional** (GEAR e GPRO dividem o mesmo contador) |
| Ordem de serviço | `OS 001.IMP.TONY PLACAS.26` | por **contrato** |

Consequências:

- **O número exibido hoje pode não ser o que foi impresso no documento.** Se
  alguém inseriu, apagou, ordenou ou mudou o tipo de uma linha, todos os
  números abaixo dela foram recalculados. O script guarda o valor de agora e
  compara com as NTs e projetos citados nas demandas (`vinculos_demandas.csv`).
  Das 2.288 citações, 2.038 conferem, 243 têm o mesmo número com outro setor
  e 7 não existem no registro.
- **O ano não zera o contador; o arquivo é que zera.** Linhas de 2025 lançadas
  na planilha de 2026 consumiram os números 1 a 3 das NTs de 2026, que por
  isso começam em 4. Linhas de 2025 lançadas na planilha de 2024 também
  geraram números repetidos.
- **Siglas que mudaram no meio do caminho** (`DILIGÊNCIA` → `DILIG.`,
  `RELATÓRIO TÉCNICO` → `REL. TÉC.`) foram unificadas em um tipo canônico
  (`TIPOS` no script).

### O que revisar em `revisao_numeracao.xlsx`

1. Aba `proximos_numeros`: o último número de cada sequência de 2026 e o
   próximo que a Central vai emitir (por exemplo: NT 1057, OF. INT. 232,
   PROJETO NA 068, OS TONY PLACAS 061).
2. Pendências **ERRO**: 66 registros com número repetido. Exemplos:
   `NT GEAR.1361.2025` usado 5 vezes e `PARECER GLIP.3477.2024` usado 4 vezes.
   Também há números de outro ano lançados na planilha errada. Esses registros
   entram na Central como `CONFLITO`. Confirme no documento emitido (SEI,
   arquivo ou PDF) qual número cada um recebeu.
3. Lacunas: 43 números que ninguém usa, entre eles as NTs 1 a 3 de 2026.
   Entram como `NAO_UTILIZADO` e **nunca são reaproveitados**.

### Virada para a Central (sugestão)

1. Escolha um dia de corte. Na véspera, **congele as planilhas**: substitua as
   fórmulas por valores (Copiar → Colar valores) para que nenhum número mude.
2. Rode `importar_numeracao.py` e faça a carga com os comandos do fim de
   `schema.sql`. Eles criam as sequências com o último número de cada uma e
   registram os conflitos e as lacunas.
3. A partir do corte, números novos saem **só** da Central, pela função
   `num_proximo()`, que é atômica (dois usuários nunca recebem o mesmo
   número). As planilhas ficam somente para leitura.
4. Em 1º de janeiro, cada sequência recomeça em 1 automaticamente, porque o
   ano faz parte da chave.

> As planilhas e a pasta `saida/` têm dados pessoais (nomes, e-mails e
> telefones) e estão no `.gitignore`. Não as envie ao repositório.
