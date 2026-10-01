-- ==============================================================================
-- SCHEMA CENTRAL: triagem_vre.db
-- Sistema de Triagem VRE - Inspetoria Fiscal / SMAURB
-- Modelo Relacional SQLite Otimizado para Alta Performance (WAL Mode)
-- ==============================================================================

PRAGMA journal_mode = WAL;
PRAGMA synchronous = NORMAL;
PRAGMA cache_size = -64000;
PRAGMA temp_store = MEMORY;

-- 1. Backlog Ativo (Protocolos em aberto de todos os meses de competência ativos)
CREATE TABLE IF NOT EXISTS vre_backlog (
    protocolo_licenca TEXT PRIMARY KEY,
    protocolo_sil TEXT,
    cnpj_limpo TEXT NOT NULL,
    cnpj_formatado TEXT,
    nome_empresarial TEXT,
    mes_competencia TEXT NOT NULL,       -- Formato 'AAAA-MM'
    data_importacao TEXT NOT NULL,
    bairro_original TEXT,
    area TEXT,
    cnaes_licenciadas TEXT,
    fiscal_atribuido TEXT,               -- Preenchido por histórico ou pelo administrativo
    flag_restricao INTEGER DEFAULT 0,    -- 0 ou 1
    flag_noturno INTEGER DEFAULT 0,      -- 0 ou 1
    tipo_servico TEXT,                   -- 'Sem Restrição Livre', 'Com Restrição Livre', 'Noturno Livre', etc.
    status TEXT DEFAULT 'PENDENTE',       -- 'PENDENTE', 'DISTRIBUIDO', 'FINALIZADO'
    data_finalizacao TEXT,
    dados_extras_json TEXT               -- Preserva colunas originais do VRE em JSON
);

CREATE INDEX IF NOT EXISTS idx_backlog_cnpj ON vre_backlog(cnpj_limpo);
CREATE INDEX IF NOT EXISTS idx_backlog_status ON vre_backlog(status);
CREATE INDEX IF NOT EXISTS idx_backlog_competencia ON vre_backlog(mes_competencia);

-- 2. Histórico Permanente de Atendimentos (Consultas de fiscal anterior por CNPJ)
CREATE TABLE IF NOT EXISTS historico_fiscais (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    cnpj_limpo TEXT NOT NULL,
    fiscal TEXT NOT NULL,
    data_atendimento TEXT,
    origem_registro TEXT
);

CREATE INDEX IF NOT EXISTS idx_hist_cnpj ON historico_fiscais(cnpj_limpo);
CREATE UNIQUE INDEX IF NOT EXISTS uq_hist_cnpj_fiscal ON historico_fiscais(cnpj_limpo, fiscal);

-- 3. Tabela de Regras Normativas de CNAE
CREATE TABLE IF NOT EXISTS regras_cnae (
    cnae_7digitos TEXT PRIMARY KEY,      -- 7 dígitos numéricos limpos (ex: '4687701')
    cnae_formatado TEXT,                 -- Máscara oficial (ex: '4687-7/01')
    denominacao TEXT,
    restricao INTEGER DEFAULT 0,
    noturno INTEGER DEFAULT 0
);

CREATE INDEX IF NOT EXISTS idx_cnae_flags ON regras_cnae(restricao, noturno);

-- 4. Tabela de Bairros x Áreas Fiscais
CREATE TABLE IF NOT EXISTS bairros_areas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bairro_normalizado TEXT NOT NULL UNIQUE,
    area TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_bairro_norm ON bairros_areas(bairro_normalizado);

-- 5. Tabela de Controle de Competências (Visibilidade/Arquivamento por Mês na Planilha de Trabalho)
CREATE TABLE IF NOT EXISTS controle_competencias (
    mes_competencia TEXT PRIMARY KEY,
    arquivado INTEGER DEFAULT 0,          -- 1 = Arquivado (oculto da planilha ativa), 0 = Ativo (visível)
    atualizado_em DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_comp_arquivado ON controle_competencias(arquivado);
