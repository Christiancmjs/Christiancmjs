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
| `demandas.csv` | uma linha por demanda, já padronizada (14.307 linhas) |
| `execucoes.csv` | datas de execução por tipo de serviço (1.346 linhas) |
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
3. A coluna `possivel_duplicata_de` aponta 196 linhas com conteúdo idêntico a
   outra (mesma data, solicitante, descrição, endereço e assunto). Elas **são
   importadas**, mas devem ser conferidas. A coluna `marcada_duplicada` traz a
   marcação "X" que já existia na planilha.

Depois, carregue no banco. O arquivo `schema.sql` traz o modelo para
PostgreSQL e os comandos `\copy` comentados no fim. Esse fluxo foi testado com
os dois arquivos: 14.307 demandas e 1.346 execuções carregadas, sem nenhum
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

> As planilhas e a pasta `saida/` têm dados pessoais (nomes, e-mails e
> telefones) e estão no `.gitignore`. Não as envie ao repositório.
