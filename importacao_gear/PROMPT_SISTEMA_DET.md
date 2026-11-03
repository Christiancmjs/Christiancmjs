Você vai trabalhar no repositório `Christiancmjs/sistema-det`, o Sistema DET:
Next.js 15, React 19 e TypeScript, hoje com todos os dados no navegador
(localStorage e IndexedDB). Trabalhe numa branch nova, nunca na principal.
Antes de mudar código, leia os arquivos citados, apresente um plano por fase
e espere a minha aprovação. Faça um commit por fase e rode `npm run build` e
`npm run lint` ao fim de cada fase.

Referência obrigatória: o repositório `Christiancmjs/Christiancmjs`, pasta
`importacao_gear/`. Ela contém o tratamento das planilhas já validado com os
dados reais:
- `importar.py`: tratamento das planilhas de demandas (aba GERAL);
- `importar_numeracao.py`: tratamento das planilhas de numeração;
- `mapeamentos/*.csv`: listas de domínio e tabelas de-para;
- `README.md`: decisões e números esperados.

Essa pasta é a fonte da verdade das regras. Porte essas regras para
TypeScript dentro do DET. Não invente regras novas.

Nunca coloque planilhas, CSVs de saída ou backups no repositório: eles têm
nomes, e-mails e telefones.

==========================================================================
FASE 1 — Correção rápida: Regional não preenche ao clicar no mapa
==========================================================================
Sintoma: em "Nova demanda", página 2, ao clicar num ponto do mapa o campo
Regional não mostra a regional do ponto.

Causa encontrada:
- `components/gear/demanda-form-wizard.tsx`, função `geocodeSearchSelected`:
  procura o bairro devolvido pelo Nominatim em `bairrosPorCep(nextCep)`
  (`lib/gear-integration-store.ts`). A comparação é feita por igualdade
  exata do nome em maiúsculas.
- Quando há bairros cadastrados em Admin > Listas, `bairrosPorCep` ignora o
  CEP e devolve a lista inteira. Se o nome não bate exatamente (acento,
  prefixos como "CONJUNTO" ou "JARDIM", nome do OSM diferente do oficial), o
  código usa `bairrosDisponiveis[0]`, que é o primeiro bairro da lista, com a
  regional errada. Se a lista estiver vazia, mantém a regional anterior
  (padrão "REGIONAL SEDE").
- O campo Regional é `readOnly`, então o usuário não consegue corrigir.
- O `useEffect` de `form.cep` repete o mesmo `[0]`.

Corrija assim:
1. Resolva a regional pelo ponto, nesta ordem:
   a) Se existir camada de limites das regionais carregada em
      `lib/map-layer-store.ts` (GeoJSON ou KML), use ponto-em-polígono. Crie
      uma configuração em Admin > Listas para indicar qual camada contém as
      regionais e qual propriedade traz o nome.
   b) Senão, compare o bairro devolvido pelo Nominatim (`suburb`,
      `neighbourhood`, `quarter`, `city_district`) com a lista de bairros. A
      comparação deve ignorar acento, maiúsculas e espaços, e aceitar
      similaridade de pelo menos 0,85.
   c) Senão, use a regional da demanda cadastrada mais próxima que tenha
      coordenada, se estiver a até 300 m.
   d) Nunca use `[0]` como palpite. Se nada for encontrado, deixe a Regional
      vazia, mostre "Regional não identificada — selecione" e transforme o
      campo num Select com as 8 regionais de `lib/det-regionais.ts`.
2. Mostre de onde veio a regional ("pelo limite da regional", "pelo bairro",
   "pela demanda vizinha" ou "manual") e permita a correção manual.
3. Valide com pontos em pelo menos 3 regionais diferentes.

==========================================================================
FASE 2 — Base de dados central (substituir localStorage)
==========================================================================
Problema: cada computador tem a própria base. Os usuários não veem as
mesmas demandas. A Central de Numeração conta separadamente em cada
navegador, então dois usuários podem emitir o MESMO número. Com 14 mil
demandas o DET já grava uma versão reduzida para não estourar a cota
(`writeGearDemandas` em `lib/client-demandas-store.ts`, chave
`det:gear-storage-warning`).

Objetivo: um banco central no servidor, compartilhado por todos.
1. Pergunte onde o sistema vai rodar (servidor da Prefeitura ou nuvem)
   antes de escolher a hospedagem. Proposta padrão: PostgreSQL, Prisma e
   Route Handlers do Next.js (`app/api/...`), com um `docker-compose.yml`
   para desenvolvimento local.
2. Faça o inventário de todas as chaves `det:*`, `det_*` e `gear*` do
   localStorage e do IndexedDB `det-sistema-db`. Use
   `buildBackupPayload` em `app/admin/configuracoes/page.tsx` e procure
   `localStorage` em `lib/` e `app/`. Modele as entidades principais:
   demandas, execuções, histórico/auditoria, documentos, numerações,
   projetos, ordens de serviço, usuários, listas (bairros, assuntos, tipos
   de documento) e camadas de mapa.
3. Crie uma camada de acesso a dados em `lib/data-access/` e troque as
   leituras e gravações diretas de localStorage nos stores
   (`client-demandas-store`, `det-numbering`, `det-admin-store`,
   `det-process-store`, `gear-integration-store`, `gimp-*`, `gmon-*`,
   `map-layer-store`, `det-audit-log`) por chamadas à API. Migre um store
   por vez. O localStorage pode continuar apenas como cache de preferências
   da tela.
4. Login e permissões passam a ser validados no servidor
   (`middleware.ts`, `lib/det-access.ts`). Senhas com hash.
5. Migração dos dados atuais: crie uma tela ou rota que receba o JSON de
   backup já exportado em Configurações > Banco de Dados e grave no banco
   sem duplicar registros.
6. Listas de demandas, mapas e dashboards precisam de paginação e filtros
   no servidor. Não carregue 14 mil registros inteiros no navegador.
7. Critério de aceite: dois navegadores diferentes veem os mesmos dados, e
   a gravação de 14 mil demandas não gera aviso de cota.

==========================================================================
FASE 3 — Importador de planilhas de demandas condizente com o tratamento
==========================================================================
Hoje o importador fica em `app/admin/configuracoes/page.tsx`
(`mapSpreadsheetRows`, `handleImportSpreadsheetRows`) e grava por meio de
`importDemandasCsvText` (`lib/client-demandas-store.ts`). Problemas atuais:
- Não padroniza valores: ASSUNTO legado, PRIORIDADE com motivo, ORIGEM com
  canal de entrada e assim por diante.
- Não corrige coordenadas sem separador decimal nem latitude e longitude
  trocadas.
- Atualiza demandas pelo GEOCODE. O GEOCODE NÃO é único: um protocolo cobre
  várias linhas. Por isso o importador cria geocodes artificiais
  `DET/TRANSCON...` para os repetidos e vazios (2.338 e 3.727 casos).
- Não consolida duplicatas de conteúdo.

Mova o processamento para o servidor (API) e implemente as regras de
`importacao_gear/importar.py`:
1. Mapeamento de colunas com os nomes alternativos dos dois layouts
   (`COLUNAS` em `importar.py`). Ignore `ANO`, `MÊS` e `ColunaN`.
2. Normalização de categorias: comparação sem acento, maiúsculas ou espaços
   extras; depois, de-para de `mapeamentos/valores.csv` (inclusive
   `mover_para`); depois, correção por similaridade de pelo menos 0,86,
   registrada como AVISO. Valor não reconhecido fica vazio e vira ERRO, sem
   ser descartado em silêncio. Os domínios ficam em
   `mapeamentos/dominios.csv`.
3. ASSUNTO legado com sufixo `_IM`/`_MA`: o sufixo define o tipo
   (implantação ou manutenção), e o nome passa por
   `mapeamentos/assunto_legado.csv` para virar assunto e sub-assunto. Se
   faltar o assunto, use `mapeamentos/sub_assunto.csv`. Guarde o texto
   original em `assuntoOriginal`.
4. PRIORIDADE com motivo (ESCOLA, DOT, GABINETE...): vira ALTA e o texto vai
   para `motivoPrioridade`.
5. Datas: aceite data do Excel, `dd/mm/aaaa`, `dd/mm/aaaa - hh:mm`,
   `aaaammdd` e `ddmmaaaa`. Com várias datas na célula, use a primeira e
   copie o texto completo para observações. Ano válido: 2015 a 2027.
6. Coordenadas: recupere o separador decimal pelos 2 primeiros dígitos e
   identifique latitude e longitude pela faixa de Contagem (latitude -20,30
   a -19,60; longitude -44,50 a -43,70), não pelo nome da coluna. Fora da
   faixa, deixe vazio e registre ERRO.
7. Nº do imóvel na coluna REFERÊNCIA vai para `numero`. E-mail na coluna
   SOLICITANTE vai para `email`. Linhas só com valores-padrão são
   ignoradas.
8. Colunas VERTICAL, HORIZONTAL, REDUTOR, TRAVESSIA ELEVADA, ABRIGO, FAIXAS
   DE PANO, OBRAS FÍSICAS e DISPOSITIVO ELETRÔNICO viram registros de
   execução vinculados à demanda.
9. Identidade da demanda importada: `arquivo:linha` (`idOrigem`) mais o
   hash do conteúdo. Reimportar o mesmo arquivo atualiza sem duplicar. O
   GEOCODE original é preservado como está, inclusive repetido. Vazio ou
   "PENDENTE" fica como pendente, sem geocode artificial. Se o DET precisar
   de um identificador próprio, crie um campo separado e não reaproveite o
   GEOCODE.
10. Duplicatas de conteúdo (mesmos data, solicitante, descrição, endereço
    completo e assunto): carregue só uma. Fica a de status mais avançado; no
    empate, a mais completa e, depois, a mais recente. Complete os campos
    vazios com os das cópias. GEOCODE, NT, projeto e registro diferentes vão
    para observações. Transfira as execuções e guarde `idsConsolidados`.
11. Coordenada de outra demanda no mesmo endereço (rua + número + bairro)
    pode ser reaproveitada. Registre `precisaoCoordenada` com um destes
    valores: `PLANILHA`, `MESMO ENDEREÇO`, `GEOCODIFICADO-EXATO`,
    `GEOCODIFICADO-RUA`, `GEOCODIFICADO-BAIRRO` ou `MANUAL`.
12. Marque cada demanda importada com `lote` (`LEGADO 2024-2025`,
    `LEGADO 2026` ou `SISTEMA`). Os painéis devem separar o passivo herdado
    do fluxo novo. Os indicadores de atraso só contam demandas recebidas
    depois da data de corte, que deve ser configurável.
13. As tabelas de-para e os domínios ficam editáveis em Admin > Listas, com
    os CSVs de `mapeamentos/` como carga inicial.
14. O assistente mostra, antes de confirmar: total lido, linhas ignoradas,
    duplicatas consolidadas, pendências por tipo (ERRO, AVISO, INFO) e
    download de um relatório de pendências com arquivo, linha, campo, valor
    original, problema e ação.

Teste com as planilhas reais `Demandas_GEAR_2024-2025.xlsx` e
`Demandas_GEAR_2026_12.xlsx` (aba GERAL). Eu forneço os arquivos; não os
coloque no repositório. O resultado deve bater com o script Python:
- 10.967 + 3.340 linhas válidas e 116 linhas vazias ignoradas;
- 196 duplicatas consolidadas, resultando em 14.111 demandas;
- 1.333 execuções;
- 5.409 coordenadas da planilha e 363 reaproveitadas;
- 54 pendências ERRO.

==========================================================================
FASE 4 — Importação da numeração para a Central de Numeração
==========================================================================
Hoje o destino NUMERAÇÃO do importador grava em
`det:importacao:numeracao:v1`, que a Central (`lib/det-numbering.ts`) não
lê. A Central recomeçaria do 1 e repetiria números.

Implemente as regras de `importacao_gear/importar_numeracao.py`:
1. Leia as abas DOCUMENTAÇÃO, VISTORIAS, PROJETOS e ORDEM DE SERVIÇO das
   planilhas "NUMERAÇÃO DE DOCUMENTAÇÃO 2024/2025/2026 - DET-DPL" e
   "NUMERAÇÃO DE ORDEM DE SERVIÇO 2025/2026". Os números são FÓRMULAS
   (`COUNTIFS`): use o valor calculado e ignore as linhas "PREENCHA ...".
2. Decomponha cada número:
   - documento `TIPO SETOR.0000.AAAA`;
   - projeto `PROJETO SETOR RG 000.AAAA`;
   - OS `OS 000.SETOR.CONTRATO.AA`.
   Unifique as siglas de tipo pela tabela `TIPOS` do script (`DILIGÊNCIA` =
   `DILIG.`, `RELATÓRIO TÉCNICO` = `REL. TÉC.` etc.).
3. Grave na própria Central, já no banco da Fase 2, todos os números
   emitidos com origem `PLANILHA`, arquivo e linha. Situações:
   - `UTILIZADO`: número normal;
   - `CONFLITO`: número repetido nas planilhas (66 casos);
   - `NÃO UTILIZADO`: lacuna na sequência (43 casos), nunca reaproveitada.
   Acrescente `CONFLITO` e `NÃO UTILIZADO` aos status da Central.
4. Regra do contador (DECIDIDA):
   - até 31/12/2026, manter a regra das planilhas:
     - documentos: um contador por TIPO + ANO, compartilhado entre setores
       (NT GEAR 1051 e NT GMON 1052 vêm do mesmo contador);
     - projetos: um contador por REGIONAL + ANO, no formato
       `PROJETO {SETOR} {RG} {000}.{AAAA}`;
     - OS: um contador por CONTRATO + ANO, no formato
       `OS {000}.{SETOR}.{CONTRATO}.{AA}`.
   - a partir de 01/01/2027, a regra deve ser configurável em Admin
     (escopo do contador e formato por família), com a data de vigência
     registrada.
   Ajuste `proximoNumeroDet` e `criarNumeracaoDet`. Hoje o contador é por
   setor + tipo, `extractSequencial` só entende `.0000.AAAA` e a checagem de
   colisão compara o texto inteiro. A colisão deve ser checada por
   (sequência, sequencial). A emissão precisa ser atômica no banco
   (transação ou `UPDATE ... RETURNING`), para que dois usuários
   simultâneos nunca recebam o mesmo número. Verifique também a numeração
   de OS em `lib/gimp-service-orders-store.ts` e a de projetos em
   `lib/det-projects-store.ts` e `lib/gear-numbering.ts`, para usarem a
   mesma Central.
5. Mostre na Central, por sequência: último número, próximo número,
   conflitos e lacunas.
6. Vincule as NTs e projetos citados nas demandas (campos NT/CI e PROJETO)
   aos números da Central. Na comparação, o setor não faz parte da chave.
   Liste as citações "mesmo número com outro setor" e "não encontrado".
7. Resultado esperado com as planilhas reais:
   - 14.054 números em 69 sequências, 66 conflitos e 43 lacunas;
   - próximos de 2026: NT 1057, OF. INT. 232, PROJETO NA 068, OS TONY
     PLACAS 061, OS TRANSINAL 018/2023 099;
   - vínculos com as demandas: 2.023 OK, 242 com outro setor e 7 não
     encontrados.
8. Na tela de importação, lembre o usuário de congelar as planilhas
   (Copiar > Colar valores) antes da importação final, e de que a partir
   do corte só a Central emite números.

==========================================================================
FASE 5 — Geocodificação das demandas sem coordenada (OpenStreetMap)
==========================================================================
Hoje a geocodificação roda no navegador (`lib/det-geocoding.ts`, Central
de Saúde DET, "Geocodificar tudo").
1. Leve a fila para o servidor (job ou worker), para não depender de uma
   aba aberta. Ela deve continuar de onde parou.
2. Respeite a política de uso do Nominatim: no máximo 1 requisição por
   segundo, User-Agent identificando o sistema e um e-mail de contato
   configurável, cache permanente de endereço para resultado e nenhuma
   consulta repetida.
3. Restrinja a busca a Contagem/MG (`viewbox` com `bounded=1`,
   `countrycodes=br`). Tente nesta ordem:
   - rua + número + bairro;
   - rua + bairro;
   - `logradouroCompleto`;
   - referência ou cruzamento;
   - bairro.
   Grave a `precisaoCoordenada` correspondente.
4. Resultado fora da faixa de Contagem é rejeitado. Toda coordenada obtida
   também define a regional pela mesma lógica da Fase 1.
5. Crie uma tela de revisão no mapa: filtro por precisão, ajuste do ponto
   arrastando o marcador (vira `MANUAL`) e um contador de pendentes.
6. Esperado: cerca de 8.339 demandas sem coordenada após a importação. Das
   que não tiverem endereço suficiente, liste quais e por quê.

==========================================================================
ENTREGA
==========================================================================
- Um arquivo `ALTERACOES-vX.Y.Z-...md` por fase, no padrão do repositório.
- Um resumo final com o que foi feito, o que ficou pendente e como fazer a
  virada: congelar planilhas, exportar o backup do navegador, migrar,
  importar as planilhas e liberar a Central.
- Testes automatizados para o normalizador de valores, as datas, as
  coordenadas, a consolidação de duplicatas e o cálculo do próximo número,
  usando os casos de exemplo citados acima.
