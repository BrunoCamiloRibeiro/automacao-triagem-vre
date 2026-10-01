# Modelo de Dados Relacional (MER / DER)

**Sistema**: Automação de Triagem VRE — SMAURB / Inspetoria Fiscal  
**Padrão**: Modelo Entidade-Relacionamento Relacional (Peter Chen / Codd)  
**Banco de Dados**: SQLite 3 (Modo WAL)  
**Versão**: 2.2.0  

---

## 1. Diagrama Entidade-Relacionamento (DER)

```mermaid
erDiagram
    %% =========================================================================
    %% BANCO CENTRAL: triagem_vre.db
    %% =========================================================================
    vre_backlog {
        TEXT protocolo_licenca PK "Protocolo único da licença VRE"
        TEXT protocolo_sil "Número de protocolo no sistema SIL"
        TEXT cnpj_limpo "14 dígitos numéricos normalizados"
        TEXT cnpj_formatado "Máscara XX.XXX.XXX/XXXX-XX"
        TEXT nome_empresarial "Razão social da empresa"
        TEXT mes_competencia "Competência no formato AAAA-MM"
        TEXT data_importacao "Data da inserção no sistema"
        TEXT bairro_original "Grafia do bairro conforme remessa"
        TEXT area "Área fiscal atribuída (HB, Central, etc)"
        TEXT cnaes_licenciadas "Lista de códigos CNAE licenciados"
        TEXT fiscal_atribuido "Fiscal designado por histórico/gestão"
        INTEGER flag_restricao "0 = Não, 1 = Possui restrição"
        INTEGER flag_noturno "0 = Não, 1 = Atividade noturna"
        TEXT tipo_servico "Classificação textual oficial"
        TEXT status "PENDENTE, DISTRIBUIDO, FINALIZADO"
        TEXT data_finalizacao "Data de baixa do protocolo"
        TEXT dados_extras_json "Payload JSON das colunas VRE originais"
    }

    historico_fiscais {
        INTEGER id PK "Identificador sequencial"
        TEXT cnpj_limpo "14 dígitos do contribuinte"
        TEXT fiscal "Nome do auditor/fiscal responsável"
        TEXT data_atendimento "Data ISO AAAA-MM-DD"
        TEXT origem_registro "Planilha ou lote de origem"
    }

    regras_cnae {
        TEXT cnae_7digitos PK "7 dígitos numéricos IBGE"
        TEXT cnae_formatado "Máscara XXXX-X/XX"
        TEXT denominacao "Descrição oficial da subclasse"
        INTEGER restricao "1 = Restrição de zoneamento municipal"
        INTEGER noturno "1 = Atividade com operação noturna"
    }

    bairros_areas {
        INTEGER id PK "Identificador sequencial"
        TEXT bairro_normalizado UK "Nome em maiúsculas sem acento (NFKD)"
        TEXT area "Área de fiscalização de posturas"
    }

    controle_competencias {
        TEXT mes_competencia PK "Mês no formato AAAA-MM"
        INTEGER arquivado "0 = Ativo/Visível, 1 = Arquivado/Oculto"
        DATETIME atualizado_em "Timestamp da última alteração"
    }

    %% =========================================================================
    %% BANCO SATÉLITE: cep_areas.db
    %% =========================================================================
    cep_mapeamento {
        TEXT cep_limpo PK "8 dígitos numéricos"
        TEXT cep_formatado "Máscara XXXXX-XXX"
        TEXT area "Área municipal vinculada"
        TEXT bairro_referencia "Bairro principal associado"
        TEXT origem "MANUAL, INFERIDO_PLANILHA, etc"
        DATETIME atualizado_em "Timestamp de atualização"
    }

    cep_pendentes {
        TEXT cep_limpo PK "8 dígitos do CEP sem área"
        TEXT cep_formatado "Máscara XXXXX-XXX"
        TEXT bairro_referencia "Bairro original informado no VRE"
        TEXT logradouro "Logradouro informado no VRE"
        INTEGER ocorrencias "Contador de protocolos afetados"
        DATETIME criado_em "Data do primeiro registro"
        DATETIME atualizado_em "Data da última ocorrência"
    }

    %% Relacionamentos lógicos
    vre_backlog }o--|| controle_competencias : "pertence a"
    vre_backlog }o--o| historico_fiscais : "herda fiscal por cnpj_limpo"
    vre_backlog }o--o| bairros_areas : "resolve área por bairro_normalizado"
    vre_backlog }o--o| cep_mapeamento : "fallback de área por cep_limpo"
    cep_pendentes ||--o| cep_mapeamento : "promovido a"
```

---

## 2. Dicionário de Dados Detalhado

### 2.1. Tabela `vre_backlog` (Banco Central)
Armazena a totalidade dos protocolos de licenciamento importados.

| Coluna | Tipo | Nulo? | Padrão | Descrição e Regras |
| :--- | :--- | :--- | :--- | :--- |
| `protocolo_licenca` | `TEXT` | Não | PK | Código unívoco da licença expedida pelo VRE (ex.: `B000100...`). |
| `protocolo_sil` | `TEXT` | Sim | `NULL` | Protocolo gerado no integrador estadual SIL. |
| `cnpj_limpo` | `TEXT` | Não | - | Chave de busca fiscal contendo exatamente 14 dígitos numéricos. |
| `cnpj_formatado` | `TEXT` | Sim | `NULL` | CNPJ formatado para exibição visual na planilha de trabalho. |
| `nome_empresarial` | `TEXT` | Sim | `NULL` | Razão social ou denominação da pessoa jurídica. |
| `mes_competencia` | `TEXT` | Não | - | Mês extraído da data oficial do VRE no formato `AAAA-MM`. |
| `data_importacao` | `TEXT` | Não | - | Data em que a linha foi ingerida pelo pipeline. |
| `bairro_original` | `TEXT` | Sim | `NULL` | Bairro textual exato como expedido no CSV do VRE. |
| `area` | `TEXT` | Sim | `NULL` | Região fiscal municipal (ex.: `HB`, `CENTRAL`, `BOSQUE`). |
| `cnaes_licenciadas`| `TEXT` | Sim | `NULL` | String contendo todos os CNAEs autorizados na licença. |
| `fiscal_atribuido` | `TEXT` | Sim | `NULL` | Fiscal que atendeu historicamente ou designado pela supervisão. |
| `flag_restricao` | `INTEGER` | Não | `0` | `1` se qualquer atividade possuir restrição de uso e ocupação do solo. |
| `flag_noturno` | `INTEGER` | Não | `0` | `1` se qualquer atividade operar no período noturno. |
| `tipo_servico` | `TEXT` | Sim | - | Classificação de atendimento (`Sem Restrição Livre`, etc.). |
| `status` | `TEXT` | Não | `'PENDENTE'`| Estado de atendimento (`PENDENTE`, `DISTRIBUIDO`, `FINALIZADO`). |
| `data_finalizacao` | `TEXT` | Sim | `NULL` | Data de conclusão da vistoria e arquivamento administrativo. |
| `dados_extras_json`| `TEXT` | Sim | `NULL` | Objeto JSON preservando integralmente todas as colunas do CSV VRE. |

---

### 2.2. Tabela `historico_fiscais` (Banco Central)
Indexa todos os atendimentos fiscais passados para garantia de continuidade na fiscalização.

| Coluna | Tipo | Nulo? | Padrão | Descrição e Regras |
| :--- | :--- | :--- | :--- | :--- |
| `id` | `INTEGER` | Não | PK Auto | Identificador primário sequencial. |
| `cnpj_limpo` | `TEXT` | Não | - | 14 dígitos numéricos indexados via B-Tree. |
| `fiscal` | `TEXT` | Não | - | Nome completo ou denominação oficial do fiscal. |
| `data_atendimento`| `TEXT` | Sim | `NULL` | Data ISO (`AAAA-MM-DD`) da realização do atendimento. |
| `origem_registro` | `TEXT` | Sim | `NULL` | Origem da carga (ex.: `Planilha_Legada_2025`, `Manual`). |

---

## 3. Otimizações de Desempenho e Pragmas do SQLite

Para viabilizar consultas em subsegundo mesmo com centenas de milhares de linhas, os seguintes pragmas de engenharia são configurados:

```sql
PRAGMA journal_mode = WAL;
PRAGMA synchronous = NORMAL;
PRAGMA cache_size = -64000;      -- Alocação de 64 MB de cache RAM para páginas B-Tree
PRAGMA temp_store = MEMORY;      -- Tabelas e ordenações temporárias mantidas em memória
```

### 3.1. Índices Estratégicos Criados
- `idx_backlog_cnpj` em `vre_backlog(cnpj_limpo)`: Busca imediata de duplicatas.
- `idx_backlog_status` em `vre_backlog(status)`: Acelera filtragem de protocolos em aberto.
- `idx_backlog_competencia` em `vre_backlog(mes_competencia)`: Filtragem instantânea na geração de planilhas.
- `idx_hist_cnpj` em `historico_fiscais(cnpj_limpo)`: Consulta de histórico em $O(\log N)$.
- `uq_hist_cnpj_fiscal` (UNIQUE): Impede duplicação do mesmo par (CNPJ x Fiscal).
- `idx_bairro_norm` em `bairros_areas(bairro_normalizado)`: Resolução direta por hash B-Tree.
