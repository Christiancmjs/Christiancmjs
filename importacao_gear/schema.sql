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
    possivel_duplicata_de         TEXT,                   -- id_origem da 1ª ocorrência com conteúdo idêntico
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
