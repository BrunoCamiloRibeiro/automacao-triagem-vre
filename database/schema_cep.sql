-- ==============================================================================
-- SCHEMA CEP X ÁREA: cep_areas.db
-- Base Dedicada de Mapeamento Dinâmico CEP x Área Municipal
-- Inspetoria Fiscal / SMAURB
-- ==============================================================================

PRAGMA journal_mode = WAL;
PRAGMA synchronous = NORMAL;
PRAGMA cache_size = -64000;
PRAGMA temp_store = MEMORY;

-- Tabela Principal de Mapeamento CEP x Área
CREATE TABLE IF NOT EXISTS cep_mapeamento (
    cep_limpo TEXT PRIMARY KEY,          -- 8 dígitos numéricos limpos
    cep_formatado TEXT,                  -- Ex: '15000-000'
    area TEXT NOT NULL,                  -- Ex: 'CENTRAL', 'HB', 'BOSQUE', etc.
    bairro_referencia TEXT,
    origem TEXT,                         -- 'MANUAL', 'INFERIDO_PLANILHA', 'LOGRADOURO', 'DEMONSTRATIVO'
    atualizado_em DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_cep_area ON cep_mapeamento(area);

-- Tabela Auxiliar de CEPs Pendentes de Mapeamento
CREATE TABLE IF NOT EXISTS cep_pendentes (
    cep_limpo TEXT PRIMARY KEY,
    cep_formatado TEXT,
    bairro_referencia TEXT,
    logradouro TEXT,
    ocorrencias INTEGER DEFAULT 1,
    criado_em DATETIME DEFAULT CURRENT_TIMESTAMP,
    atualizado_em DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_cep_pend_ocorrencias ON cep_pendentes(ocorrencias DESC);
