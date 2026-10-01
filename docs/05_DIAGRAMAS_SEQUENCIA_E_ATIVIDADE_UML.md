# Diagramas de Sequência e Atividades (UML 2.5)

**Sistema**: Automação de Triagem VRE — SMAURB / Inspetoria Fiscal  
**Padrão**: UML 2.5 Behavioral Models (OMG)  
**Versão**: 2.2.0  

---

## 1. Diagramas de Sequência

### 1.1. Sequência do Pipeline Headless (`triagem_vre`)
Ilustra a cadeia de execução instantânea disparada por duplo clique ou arraste de CSV:

```mermaid
sequenceDiagram
    autonumber
    actor User as Operador (Duplo Clique / Drag & Drop)
    participant Runner as HeadlessRunner (main_seco.py)
    participant Norm as NormalizerService
    participant DB as SQLite Central (triagem_vre.db)
    participant CEP as SQLite Satélite (cep_areas.db)
    participant Writer as ExcelWriterService (OpenPyXL)
    participant FS as Sistema de Arquivos

    User->>Runner: Dispara Executar_Triagem.bat [arquivo.csv]
    Runner->>FS: Varre e identifica encoding do CSV (UTF-8 / Latin-1)
    FS-->>Runner: Linhas decodificadas

    loop Para cada protocolo no CSV
        Runner->>Norm: normalizar_cnpj(cnpj_bruto)
        Norm-->>Runner: cnpj_limpo (14 dígitos)
        
        Runner->>DB: SELECT fiscal FROM historico_fiscais WHERE cnpj_limpo = ?
        DB-->>Runner: fiscal_anterior (ou vazio)
        
        Runner->>Norm: extrair_cnaes(cnaes_licenciadas)
        Norm-->>Runner: lista de CNAEs (7 dígitos)
        
        loop Matching de CNAEs com Curto-Circuito
            Runner->>DB: SELECT restricao, noturno FROM regras_cnae WHERE cnae = ?
            DB-->>Runner: flags (restricao, noturno)
            Note over Runner: Se (flag_r == 1 E flag_n == 1) ➔ Interrompe loop
        end
        
        Runner->>Norm: normalizar_bairro(bairro_bruto)
        Norm-->>Runner: bairro_norm (NFKD)
        
        Runner->>DB: SELECT area FROM bairros_areas WHERE bairro_normalizado = ?
        alt Bairro Encontrado
            DB-->>Runner: area_municipal
        else Bairro Desconhecido (Fallback)
            Runner->>CEP: SELECT area FROM cep_mapeamento WHERE cep_limpo = ?
            alt CEP Encontrado
                CEP-->>Runner: area_municipal
            else CEP Desconhecido
                Runner->>CEP: INSERT OR UPDATE em cep_pendentes
                CEP-->>Runner: Enfileirado
            end
        end

        Runner->>DB: INSERT OR IGNORE INTO vre_backlog (...)
        Note over Runner,DB: Preserva fiscal e status pré-existentes
    end

    Runner->>DB: SELECT * FROM vre_backlog WHERE status IN ('PENDENTE', 'DISTRIBUIDO')
    DB-->>Runner: protocolos_ativos

    Runner->>Writer: gerar_planilha_trabalho(protocolos_ativos)
    Writer->>FS: Grava em Planilha_Controle_VRE_Ativa.tmp.xlsx
    Writer->>FS: Renomeia atomicamente para Planilha_Controle_VRE_Ativa.xlsx
    Writer-->>Runner: Caminho do arquivo oficial
    Runner-->>User: Exibe mensagem verde de sucesso e encerra
```

---

### 1.2. Sequência de Mapeamento de CEP e Backfill Automático

```mermaid
sequenceDiagram
    autonumber
    actor Admin as Administrador
    participant GUI as Painel Administrativo (Tkinter)
    participant CepSvc as CepService
    participant CepDB as cep_areas.db
    participant CentralDB as triagem_vre.db

    Admin->>GUI: Informa CEP: "15042-015", Área: "CIDADE DA CRIANÇA"
    Admin->>GUI: Clica em "Salvar / Atualizar Mapeamento"
    GUI->>CepSvc: registrar_area_cep(cep, area, bairro)
    
    CepSvc->>CepDB: INSERT INTO cep_mapeamento (...) ON CONFLICT DO UPDATE
    CepSvc->>CepDB: DELETE FROM cep_pendentes WHERE cep_limpo = ?
    
    Note over CepSvc,CentralDB: Rotina de Backfill em Backlog Pendente
    CepSvc->>CentralDB: SELECT protocolo_licenca, dados_extras_json FROM vre_backlog WHERE area IS NULL
    CentralDB-->>CepSvc: protocolos_sem_area
    CepSvc->>CentralDB: UPDATE vre_backlog SET area = 'CIDADE DA CRIANÇA' WHERE protocolo_licenca IN (...)
    CentralDB-->>CepSvc: total_atualizados
    
    CepSvc-->>GUI: Retorna Sucesso (N protocolos atualizados)
    GUI-->>Admin: Exibe mensagem informativa com total de protocolos corrigidos
```

---

## 2. Diagramas de Atividades

### 2.1. Algoritmo de Herança Hierárquica de CNAEs (Seeding)
Demonstra como restrições de Grupos e Classes propagam-se para as subclasses da tabela `regras_cnae`:

```mermaid
flowchart TD
    Start([Início da Carga de CNAEs]) --> ReadRow[/Ler linha do arquivo Excel de CNAEs/]
    ReadRow --> IsValid{Linha válida de dados?}
    IsValid -- Não --> HasMoreRows{Mais linhas?}
    IsValid -- Sim --> CheckSec{É Seção nova?<br/>Letra A-U}
    
    CheckSec -- Sim --> ResetBranch["Resetar ramos inferiores<br/>(div, grp, cls = 0)"] --> HasMoreRows
    CheckSec -- Não --> CheckDiv{É Divisão nova?<br/>2 dígitos}
    
    CheckDiv -- Sim --> SetDiv["cur_div = flag_restr / flag_noturno<br/>(Reset grp, cls = 0)"] --> HasMoreRows
    CheckDiv -- Não --> CheckGrp{É Grupo novo?<br/>ex: 56.1}
    
    CheckGrp -- Sim --> SetGrp["cur_grp = flag | cur_div<br/>(Reset cls = 0)"] --> HasMoreRows
    CheckGrp -- Não --> CheckCls{É Classe nova?<br/>ex: 56.11-2}
    
    CheckCls -- Sim --> SetCls["cur_cls = flag | cur_grp"] --> HasMoreRows
    CheckCls -- Não --> CheckSub{É Subclasse?<br/>7 dígitos numéricos}
    
    CheckSub -- Sim --> InheritFlags["final_restr = flag | cur_cls_restr<br/>final_noturno = flag | cur_cls_noturno"]
    InheritFlags --> AddBatch[Adicionar à lista de inserção em lote]
    AddBatch --> HasMoreRows
    
    HasMoreRows -- Sim --> ReadRow
    HasMoreRows -- Não --> BulkInsert["Executar INSERT OR REPLACE em regras_cnae<br/>(Transação Única em SQLite)"]
    BulkInsert --> EndNode([Fim da Carga])
```

---

### 2.2. Algoritmo de Resolução de Área e Fallback com Raias (Swimlanes)

```mermaid
flowchart TD
    subgraph "Módulo de Normalização"
        A1[Recebe Bairro e CEP Brutos] --> A2[Aplica Decomposição NFKD no Bairro]
        A2 --> A3[Expande Abreviações JD, PQ, VL, RES]
    end

    subgraph "Banco Central (triagem_vre.db)"
        A3 --> B1[(Consulta tabela bairros_areas)]
        B1 --> B2{Bairro localizado?}
        B2 -- Sim --> B3[Retorna Área do Bairro]
    end

    subgraph "Banco Satélite (cep_areas.db)"
        B2 -- Não --> C1[Higieniza CEP para 8 dígitos]
        C1 --> C2[(Consulta tabela cep_mapeamento)]
        C2 --> C3{CEP localizado?}
        C3 -- Sim --> C4[Retorna Área pelo CEP]
        C3 -- Não --> C5[(Registra em cep_pendentes)]
        C5 --> C6[Retorna Área Vazia / Status PENDENTE]
    end

    B3 --> Fim([Área Resolvida com Sucesso])
    C4 --> Fim
    C6 --> FimPendente([Aguardando Resolução Administrativa])
```
