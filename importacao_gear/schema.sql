-- Modelo de dados sugerido (PostgreSQL) para receber a saída de importar.py
-- Adapte nomes/tipos ao banco do sistema se ele já tiver tabelas próprias.

CREATE TABLE IF NOT EXISTS demandas (
    id                            BIGSERIAL PRIMARY KEY,
    id_origem                     TEXT UNIQUE NOT NULL,   -- arquivo:linha; permite reimportar sem duplicar (upsert)
    arquivo_origem                TEXT NOT NULL,
    linha_excel                   INTEGER NOT NULL,
    data_recebimento              DATE,
    data_solicitacao              DATE,
    canal_entrada                 TEXT,
    remetente                     TEXT,
    origem                        TEXT,
    registro                      TEXT,
    solicitante                   TEXT,
    email_solicitante             TEXT,
    viabilidade                   TEXT,
    status                        TEXT,
    prioridade                    TEXT,
    motivo_prioridade             TEXT,                   -- texto legado de 2024-2025 (ESCOLA, DOT, GABINETE...)
    alteracao_prioridade_interno  TEXT,
    alteracao_prioridade_externo  TEXT,
    data_alteracao_prioridade     DATE,
    tipo_intervencao              TEXT,
    assunto                       TEXT,
    sub_assunto                   TEXT,
    assunto_original              TEXT,                   -- valor exatamente como estava na planilha
    descricao                     TEXT,
    logradouro                    TEXT,
    numero                        TEXT,
    referencia                    TEXT,
    bairro                        TEXT,
    logradouro_completo           TEXT,
    regional                      TEXT,
    latitude                      NUMERIC(10,7),
    longitude                     NUMERIC(10,7),
    precisao_coordenada           TEXT,                   -- PLANILHA | MESMO ENDEREÇO DE OUTRA DEMANDA | GEOCODIFICADO...
    programa_asfalto_novo         TEXT,
    nota_tecnica                  TEXT,
    projeto                       TEXT,
    medida_alternativa            TEXT,
    respondido                    BOOLEAN,
    data_resposta_email           TIMESTAMP,
    data_resposta_gabinete        TIMESTAMP,
    encaminhada_para              TEXT,
    data_encaminhamento           TIMESTAMP,
    geocode                       TEXT,                   -- NÃO é único: um protocolo pode cobrir várias linhas
    data_baixa_geocode            DATE,
    data_execucao                 DATE,
    observacoes                   TEXT,
    marcada_duplicada             BOOLEAN DEFAULT FALSE,  -- coluna DUPLICADAS = X na planilha
    ids_consolidados              TEXT,                   -- id_origem das cópias duplicadas fundidas nesta demanda
    hash_conteudo                 TEXT
);

CREATE TABLE IF NOT EXISTS execucoes (
    id             BIGSERIAL PRIMARY KEY,
    demanda_id     BIGINT NOT NULL REFERENCES demandas(id) ON DELETE CASCADE,
    tipo_servico   TEXT NOT NULL,   -- VERTICAL, HORIZONTAL, REDUTOR, TRAVESSIA ELEVADA, ABRIGO, ...
    data_execucao  DATE NOT NULL,
    UNIQUE (demanda_id, tipo_servico, data_execucao)
);

-- Tabela de domínios (listas da aba REFERÊNCIAS). Popular com mapeamentos/dominios.csv.
CREATE TABLE IF NOT EXISTS dominios (
    campo TEXT NOT NULL,
    valor TEXT NOT NULL,
    PRIMARY KEY (campo, valor)
);

CREATE INDEX IF NOT EXISTS ix_demandas_data      ON demandas (data_recebimento);
CREATE INDEX IF NOT EXISTS ix_demandas_status    ON demandas (status);
CREATE INDEX IF NOT EXISTS ix_demandas_regional  ON demandas (regional);
CREATE INDEX IF NOT EXISTS ix_demandas_geocode   ON demandas (geocode);

-- ---------------------------------------------------------------------------
-- Carga (psql). Executar a partir da pasta importacao_gear/ depois de rodar importar.py:
--
--   \copy dominios FROM 'mapeamentos/dominios.csv' CSV HEADER
--
--   CREATE TEMP TABLE stg_demandas (LIKE demandas INCLUDING DEFAULTS);
--   ALTER TABLE stg_demandas DROP COLUMN id;
--   \copy stg_demandas FROM 'saida/demandas.csv' CSV HEADER
--   INSERT INTO demandas SELECT nextval('demandas_id_seq'), s.* FROM stg_demandas s
--     ON CONFLICT (id_origem) DO NOTHING;           -- reimportação segura
--
--   CREATE TEMP TABLE stg_execucoes (id_origem TEXT, tipo_servico TEXT, data_execucao DATE);
--   \copy stg_execucoes FROM 'saida/execucoes.csv' CSV HEADER
--   INSERT INTO execucoes (demanda_id, tipo_servico, data_execucao)
--     SELECT d.id, e.tipo_servico, e.data_execucao
--       FROM stg_execucoes e JOIN demandas d USING (id_origem)
--     ON CONFLICT DO NOTHING;
-- ---------------------------------------------------------------------------

-- ===========================================================================
-- CENTRAL DE NUMERAÇÃO (saída de importar_numeracao.py)
-- ===========================================================================
-- Uma sequência = um contador. Igual às planilhas:
--   DOCUMENTO -> por tipo de documento + ano (compartilhado entre setores)
--   PROJETO   -> por regional + ano
--   OS        -> por contrato + ano
CREATE TABLE IF NOT EXISTS num_sequencias (
    id              BIGSERIAL PRIMARY KEY,
    familia         TEXT NOT NULL CHECK (familia IN ('DOCUMENTO','PROJETO','OS')),
    sequencia       TEXT NOT NULL,          -- 'NOTA TÉCNICA', 'RE', 'TONY PLACAS'...
    ano             INTEGER NOT NULL,
    ultimo_numero   INTEGER NOT NULL DEFAULT 0,
    UNIQUE (familia, sequencia, ano)
);

CREATE TABLE IF NOT EXISTS num_numeros (
    id                  BIGSERIAL PRIMARY KEY,
    sequencia_id        BIGINT REFERENCES num_sequencias(id),
    sequencial          INTEGER,
    numero              TEXT NOT NULL,      -- exatamente como foi emitido
    numero_padronizado  TEXT,
    tipo                TEXT,
    setor               TEXT,
    regional_sigla      TEXT,
    contrato            TEXT,
    data_emissao        DATE,
    responsavel         TEXT,
    assunto             TEXT,
    local               TEXT,
    regional            TEXT,
    situacao            TEXT NOT NULL DEFAULT 'EMITIDO'
                        CHECK (situacao IN ('EMITIDO','CANCELADO','NAO_UTILIZADO','CONFLITO')),
    origem              TEXT NOT NULL DEFAULT 'SISTEMA',  -- 'PLANILHA' para o legado importado
    arquivo_origem      TEXT,
    linha_excel         INTEGER,
    criado_em           TIMESTAMP NOT NULL DEFAULT now()
);
-- Sem UNIQUE em (sequencia_id, sequencial): o legado TEM números repetidos (situacao = 'CONFLITO').
-- Para números novos, a unicidade é garantida pela função abaixo e por este índice parcial:
CREATE UNIQUE INDEX IF NOT EXISTS ux_num_numeros_novos
    ON num_numeros (sequencia_id, sequencial) WHERE origem = 'SISTEMA';

-- Emissão atômica do próximo número (sem corrida entre dois usuários ao mesmo tempo)
CREATE OR REPLACE FUNCTION num_proximo(p_familia TEXT, p_sequencia TEXT, p_ano INTEGER)
RETURNS INTEGER LANGUAGE sql AS $$
    INSERT INTO num_sequencias (familia, sequencia, ano, ultimo_numero)
    VALUES (p_familia, p_sequencia, p_ano, 1)
    ON CONFLICT (familia, sequencia, ano)
    DO UPDATE SET ultimo_numero = num_sequencias.ultimo_numero + 1
    RETURNING ultimo_numero;
$$;

-- ---------------------------------------------------------------------------
-- Carga do legado (psql), a partir de importacao_gear/:
--
--   CREATE TEMP TABLE stg_num (familia TEXT, numero TEXT, numero_padronizado TEXT, tipo TEXT, sigla TEXT,
--     setor_num TEXT, regional_sigla TEXT, contrato TEXT, sequencial NUMERIC, ano NUMERIC, escopo TEXT,
--     data DATE, setor TEXT, tipo_planilha TEXT, responsavel TEXT, assunto TEXT, local TEXT, regional TEXT,
--     status TEXT, observacoes TEXT, dimensao_m TEXT, repetido BOOLEAN, arquivo TEXT, aba TEXT, linha_excel INTEGER);
--   \copy stg_num FROM 'saida/numeracao/numeros_utilizados.csv' CSV HEADER
--
--   INSERT INTO num_sequencias (familia, sequencia, ano, ultimo_numero)
--     SELECT familia, split_part(escopo,'|',2), ano::int, max(sequencial)::int
--       FROM stg_num WHERE escopo IS NOT NULL GROUP BY 1,2,3
--     ON CONFLICT (familia, sequencia, ano) DO UPDATE
--       SET ultimo_numero = GREATEST(num_sequencias.ultimo_numero, EXCLUDED.ultimo_numero);
--
--   INSERT INTO num_numeros (sequencia_id, sequencial, numero, numero_padronizado, tipo, setor, regional_sigla,
--       contrato, data_emissao, responsavel, assunto, local, regional, situacao, origem, arquivo_origem, linha_excel)
--     SELECT s.id, n.sequencial::int, n.numero, n.numero_padronizado, n.tipo, n.setor_num, n.regional_sigla,
--            n.contrato, n.data, n.responsavel, n.assunto, n.local, n.regional,
--            CASE WHEN n.repetido THEN 'CONFLITO' ELSE 'EMITIDO' END, 'PLANILHA', n.arquivo, n.linha_excel
--       FROM stg_num n
--       LEFT JOIN num_sequencias s ON s.familia = n.familia AND s.sequencia = split_part(n.escopo,'|',2)
--                                 AND s.ano = n.ano::int;
--
--   -- lacunas: registrar como NAO_UTILIZADO para nunca serem reaproveitadas
--   INSERT INTO num_numeros (sequencia_id, sequencial, numero, situacao, origem)
--     SELECT s.id, g, s.sequencia || ' ' || g || '/' || s.ano, 'NAO_UTILIZADO', 'PLANILHA'
--       FROM num_sequencias s, generate_series(1, s.ultimo_numero) g
--      WHERE NOT EXISTS (SELECT 1 FROM num_numeros x WHERE x.sequencia_id = s.id AND x.sequencial = g);
--
--   -- emitir um número novo:  SELECT num_proximo('DOCUMENTO', 'NOTA TÉCNICA', 2026);  -- 1057
-- ---------------------------------------------------------------------------
