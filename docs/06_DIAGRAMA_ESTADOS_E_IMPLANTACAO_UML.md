# Diagramas de Estados e Implantação (UML 2.5)

**Sistema**: Automação de Triagem VRE — SMAURB / Inspetoria Fiscal  
**Padrão**: UML 2.5 State Machine & Deployment Models (OMG)  
**Versão**: 2.2.0  

---

## 1. Diagramas de Máquina de Estados (Statechart)

### 1.1. Ciclo de Vida do Protocolo de Licenciamento (`vre_backlog`)
Demonstra as transições de estado de um protocolo empresarial desde sua emissão no VRE estadual até o arquivamento definitivo na Inspetoria Fiscal:

```mermaid
stateDiagram-v2
    [*] --> Inexistente
    
    Inexistente --> Ingerido_Pendente: Importação via CSV (main_seco.py)
    note right of Ingerido_Pendente
        Fiscal atribuído por histórico (se houver)
        Área atribuída por Bairro/CEP
        Status gravado no SQLite: 'PENDENTE'
        (Na planilha Excel ativa, fica vazio)
    end note

    Ingerido_Pendente --> Distribuido: Fiscal assume ou supervisor atribui
    note right of Distribuido
        Status gravado no SQLite: 'DISTRIBUIDO'
        Registro visível na planilha de trabalho
    end note

    Distribuido --> Em_Andamento: Início da vistoria em campo
    
    Em_Andamento --> Deferido: Vistoria aprovada sem óbices
    Em_Andamento --> Indeferido: Irregularidade insanável de zoneamento/postura
    Em_Andamento --> Notificado: Prazo concedido para adequação
    
    Notificado --> Em_Andamento: Reinspeção fiscal
    
    Deferido --> Finalizado: Baixa do protocolo
    Indeferido --> Finalizado: Notificação formalizada
    
    Finalizado --> Arquivado: Mês de competência é arquivado no painel
    note right of Arquivado
        Oculto da planilha ativa
        Preservado 100% no banco central
    end note

    Arquivado --> Reativado: Reabertura de mês pelo gestor
    Reativado --> Finalizado
```

---

### 1.2. Ciclo de Vida da Competência Mensal (`controle_competencias`)

```mermaid
stateDiagram-v2
    [*] --> Competencia_Detectada: Primeiro protocolo com competência AAAA-MM
    
    Competencia_Detectada --> Ativa_Visivel: Inserido em controle_competencias (arquivado = 0)
    note right of Ativa_Visivel
        Protocolos deste mês são exportados
        para a Planilha_Controle_VRE_Ativa.xlsx
    end note

    Ativa_Visivel --> Arquivada_Oculta: Administrador marca arquivado = 1
    note right of Arquivada_Oculta
        Protocolos são omitidos da planilha ativa,
        mas preservados intactos no SQLite
        Backup gerado em Planilhas anteriores/
    end note

    Arquivada_Oculta --> Ativa_Visivel: Administrador marca arquivado = 0 (Reativação)
```

---

## 2. Diagrama de Implantação (Deployment Diagram)

Representa a topologia física e lógica de execução do sistema em ambiente corporativo municipal:

```mermaid
flowchart TD
    subgraph Workstation ["🖥️ Estação de Trabalho do Usuário (Windows 10/11)"]
        subgraph SO ["Ambiente do Sistema Operacional"]
            subgraph Lançadores ["Camada de Lançamento"]
                BAT["Executar_Triagem.bat<br/>(Lançador Interativo em Lote)"]
                VBS["Executar_Triagem_Silencioso.vbs<br/>(Lançador Invisível wscript.exe)"]
                PANEL["Painel_Administrativo.bat<br/>(Lançador da Interface Gráfica)"]
            end

            subgraph Runtime ["Runtime de Execução Python"]
                PYW["python.exe / pythonw.exe<br/>(Python 3.12 Portátil PSF Assinado)"]
                MODS["Bibliotecas: openpyxl, xlrd, et_xmlfile"]
            end

            subgraph Storage ["Sistema de Arquivos Local / Unidade de Rede"]
                DB1[("triagem_vre.db<br/>(SQLite Central / WAL Cache 64MB)")]
                DB2[("cep_areas.db<br/>(SQLite Satélite CEP x Área)")]
                XLSX["Planilha_Controle_VRE_Ativa.xlsx<br/>(Arquivo Oficial de Trabalho)"]
                BKP_DIR["Pasta Planilhas anteriores/<br/>(Backups Datados)"]
                CSV_DIR["Pasta entrada_csv/<br/>(Arquivos VRE Ingeridos)"]
            end

            subgraph MS_Excel ["Microsoft Office / Excel Desktop"]
                EXCEL_APP["excel.exe<br/>(Acesso diário pelos auditores e fiscais)"]
            end
        end
    end

    BAT --> PYW
    VBS --> PYW
    PANEL --> PYW
    PYW --> MODS
    MODS --> DB1
    MODS --> DB2
    MODS --> XLSX
    MODS --> BKP_DIR
    CSV_DIR --> MODS
    EXCEL_APP -. Abre e edita .-> XLSX
```

### 2.1. Características de Implantação
1. **Zero Privilégio Administrativo**: A instalação em modo portátil não requer permissões de Administrador do Windows nem modificação do Registro do sistema (`regedit`).
2. **Imunidade a Antivírus Corporativo (Symantec / Windows Defender)**:
   - Em vez de empacotar via executável monolítico (*dropper* temporário em `%TEMP%`), a solução portátil utiliza o binário oficial `python.exe` / `pythonw.exe`, que possui assinatura digital autêntica da *Python Software Foundation*.
3. **Resiliência a Quedas de Rede**: O uso do banco SQLite com `PRAGMA synchronous = NORMAL` e `journal_mode = WAL` assegura que oscilações transitórias na rede não corrompam os índices relacionais.
