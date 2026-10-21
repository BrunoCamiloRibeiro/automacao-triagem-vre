# 04 — Fluxos Operacionais e Modelagem de Dados

**Sistema**: Automação de Triagem VRE & Gestão Fiscal  
**Órgão**: Prefeitura Municipal de São José do Rio Preto — SP  
**Secretaria**: Secretaria Municipal de Habitação e Urbanismo (SMAURB) / Inspetoria Fiscal  
**Versão**: 2.2.0 (Patch v1.6)  
**Padrão de Engenharia**: Modelagem Comportamental e Estrutural / Diagramação Mermaid.js  

---

## 1. Visão Geral dos Fluxos do Sistema

O processamento do Sistema VRE é orquestrado através de pipelines desacoplados que integram ingestão de dados em lote, sanitização determinística, enriquecimento geográfico em camadas, persistência transacional ACID em SQLite (modo WAL) e renderização de planilhas formatadas via `openpyxl`.

Esta seção formaliza:
1. Os **Diagramas de Sequência** detalhando as interações temporais entre componentes.
2. O **Diagrama Estrutural de Classes e Módulos** do ecossistema Python.
3. O **Modelo Entidade-Relacionamento (DER)** das bases `triagem_vre.db` e `cep_areas.db`.
4. Os **Diagramas de Transição de Estados** que governam o ciclo de vida dos protocolos e das competências mensais.

---

## 2. Diagramas de Sequência

### 2.1. Fluxo 1: Ingestão Diária e Triagem Automatizada

O fluxo principal do `main_seco.py` opera a recepção e processamento de arquivos CSV depositados em `entrada_csv/`, aplicando normalização, classificação de risco, enriquecimento de área e persistência segura.

```mermaid
sequenceDiagram
    autonumber
    actor Operador as Fiscal / Operador
    participant Runner as runner.main_seco
    participant DB as core.database
    participant Norm as core.normalizer
    participant CEP as core.cep_service
    participant Excel as core.excel_writer
    participant FS as Sistema de Arquivos

    Operador->>FS: Deposita CSV diário em entrada_csv/
    Operador->>Runner: Executa iniciar_triagem.bat (ou main_seco.py)
    
    activate Runner
    Runner->>DB: inicializar_banco()
    activate DB
    DB-->>Runner: Conexão ativa (WAL mode, foreign_keys=ON)
    deactivate DB

    Runner->>FS: escanear_arquivos_csv("entrada_csv/")
    FS-->>Runner: Lista de caminhos de arquivos CSV

    loop Para cada arquivo CSV encontrado
        Runner->>Runner: Detectar encoding (utf-8, cp1252, latin-1)
        Runner->>FS: Ler linhas e cabeçalhos
        
        loop Para cada registro do CSV
            Runner->>Norm: normalizar_cnpj(cnpj_bruto)
            Norm-->>Runner: CNPJ formatado (00.000.000/0000-00)
            
            Runner->>Norm: normalizar_cep(cep_bruto)
            Norm-->>Runner: CEP sanitizado (00000-000)
            
            Runner->>Norm: normalizar_bairro(bairro_bruto)
            Norm-->>Runner: Bairro canônico em caixa alta
            
            Runner->>Norm: classificar_risco(cnae_principal, cnaes_secundarios)
            Norm-->>Runner: Risco Diurno, Risco Noturno, Grau de Risco
            
            Runner->>CEP: resolver_area_fiscal(bairro, logradouro, cep)
            activate CEP
            Note over CEP: Fallback em 3 Níveis:<br/>1. Bairro Canônico<br/>2. CEP Mapeado<br/>3. Fila de Pendências
            CEP-->>Runner: Área Fiscal (ex: "CENTRAL", "HB", "PENDENTE")
            deactivate CEP
            
            Runner->>DB: buscar_titular_anterior_cnpj(cnpj)
            DB-->>Runner: Fiscal titular histórico (se existir)
        end
        
        Runner->>DB: salvar_protocolos_triagem(lista_protocolos)
        activate DB
        Note over DB: Transação BEGIN IMMEDIATE<br/>INSERT OR IGNORE / ON CONFLICT
        DB-->>Runner: Total inseridos / atualizados
        deactivate DB
        
        Runner->>FS: Mover CSV para entrada_csv/processados/ (com timestamp)
    end

    Runner->>DB: sincronizar_com_planilha_ativa()
    Runner->>DB: obter_registros_ativos_ordenados()
    DB-->>Runner: Lista completa de registros ativos (visivel=1)

    Runner->>Excel: gerar_planilha_controle(registros, "Planilha_Controle_VRE_Ativa.xlsx")
    activate Excel
    Note over Excel: Injeção de Plain Values<br/>Aplicação das 15 Regras de Formatação<br/>Cálculo de Hash SHA-256
    Excel->>FS: Salva arquivo Excel formatado (.xlsx)
    Excel-->>Runner: Concluído com sucesso
    deactivate Excel

    Runner-->>Operador: Exibe sumário de execução e encerra
    deactivate Runner
```

---

### 2.2. Fluxo 2: Sincronização Bidirecional e Preservação de Edições Humanas

Garante que quaisquer intervenções manuais realizadas pelos fiscais na planilha Excel (mudança de status, anotações de fiscalização, atribuição nominal de fiscal) sejam lidas de volta para o SQLite antes de qualquer nova geração, eliminando sobrescritas acidentais.

```mermaid
sequenceDiagram
    autonumber
    actor Fiscal as Fiscal / Inspetor
    participant Excel as Planilha_Controle_VRE_Ativa.xlsx
    participant Runner as runner.main_seco
    participant DB as core.database
    participant Writer as core.excel_writer

    Fiscal->>Excel: Edita colunas manuais (Status, Fiscal Designado, Anotações)
    Fiscal->>Excel: Salva e fecha arquivo Excel

    Fiscal->>Runner: Inicia nova execução da rotina
    activate Runner
    
    Runner->>Runner: verificar_se_arquivo_em_uso(caminho_excel)
    alt Arquivo aberto por outro processo (Lock Win32)
        Runner-->>Fiscal: Exibe alerta: Feche o arquivo Excel antes de prosseguir
    else Arquivo liberado para leitura
        Runner->>DB: sincronizar_planilha_para_banco(caminho_excel)
        activate DB
        DB->>Excel: Abrir workbook em modo somente-leitura (data_only=True)
        Excel-->>DB: Linhas e células preenchidas
        
        loop Para cada linha de dados (da linha 2 em diante)
            DB->>DB: Extrair Protocolo (chave primária)
            DB->>DB: Extrair campos editáveis (Status, Fiscal, Parecer, Anotações)
            DB->>DB: Comparar com registro atual do banco
            opt Houve alteração humana
                DB->>DB: UPDATE triagem SET status=?, fiscal=?, parecer=?, ... WHERE protocolo=?
                Note over DB: Atualização protegida por transação<br/>Data de modificação atualizada
            end
        end
        DB-->>Runner: Total de registros sincronizados de volta
        deactivate DB
        
        Runner->>DB: consultar_dados_consolidados()
        DB-->>Runner: Dataset completo e atualizado
        
        Runner->>Writer: gerar_planilha_controle(dados_consolidados, caminho_excel)
        activate Writer
        Writer->>Excel: Sobrescreve com visual renovado e edições preservadas
        Writer-->>Runner: Sucesso
        deactivate Writer
    end
    
    Runner-->>Fiscal: Execução finalizada com dados íntegros
    deactivate Runner
```

---

### 2.3. Fluxo 3: Arquivamento e Reconfiguração de Visibilidade

O módulo de arquivamento permite à chefia ou inspetor retirar meses/competências encerradas da visualização operacional ativa, mantendo o banco leve e gerando um snapshot histórico autônomo.

```mermaid
sequenceDiagram
    autonumber
    actor Admin as Administrador / Chefe de Triagem
    participant GUI as admin.main_admin
    participant Service as admin.archive_service
    participant DB as triagem_vre.db
    participant FS as Sistema de Arquivos (Exportações)
    participant Writer as core.excel_writer

    Admin->>GUI: Abre Painel Administrativo VRE
    activate GUI
    GUI->>Service: listar_competencias_status()
    activate Service
    Service->>DB: SELECT competencia, COUNT(*), visivel FROM triagem GROUP BY competencia
    DB-->>Service: Agrupamento de competências e quantidades
    Service-->>GUI: Lista de competências (ex: 2026-08: 412, 2026-09: 530)
    deactivate Service
    GUI-->>Admin: Exibe tabela no Treeview com botões de ação

    Admin->>GUI: Seleciona "2026-08" e clica em "Arquivar Competência"
    GUI->>GUI: Exibe diálogo de confirmação modal
    Admin->>GUI: Confirma operação de arquivamento
    
    GUI->>Service: arquivar_competencia("2026-08", exportar_excel=True)
    activate Service
    
    Service->>DB: Inicia Transação (BEGIN IMMEDIATE)
    Service->>DB: UPDATE triagem SET visivel = 0 WHERE competencia = '2026-08'
    DB-->>Service: Linhas afetadas (412)
    Service->>DB: COMMIT
    
    opt Exportação do arquivo de arquivo histórico
        Service->>DB: SELECT * FROM triagem WHERE competencia = '2026-08'
        DB-->>Service: Registros completos da competência
        Service->>Writer: gerar_planilha_controle(registros, "Arquivados/VRE_Arquivado_2026-08.xlsx")
        Writer->>FS: Grava arquivo Excel autônomo da competência
        Writer-->>Service: Arquivo gerado
    end

    Service->>DB: SELECT * FROM triagem WHERE visivel = 1 ORDER BY data_hora DESC
    DB-->>Service: Registros remanescentes ativos
    Service->>Writer: gerar_planilha_controle(remanescentes, "Planilha_Controle_VRE_Ativa.xlsx")
    Writer->>FS: Sobrescreve planilha ativa (sem as linhas arquivadas)
    Writer-->>Service: Sucesso
    
    Service-->>GUI: Relatório: 412 registros arquivados com sucesso
    deactivate Service

    GUI->>GUI: Atualiza Treeview e estatísticas de tela
    GUI-->>Admin: Notificação visual de sucesso (Messagebox)
    deactivate GUI
```

---

### 2.4. Fluxo 4: Importação em Lote de Planilhas Históricas e Reconciliação

Executado para carga retroativa de históricos legados, garantindo preservação de decisões passadas e populando a memória de atribuição fiscal por CNPJ.

```mermaid
sequenceDiagram
    autonumber
    actor Operador as Operador do Sistema
    participant Importer as runner.importar_antigas
    participant Service as core.historico_importer
    participant Norm as core.normalizer
    participant DB as core.database
    participant FS as Pastas Históricas (.xlsx)

    Operador->>Importer: Executa importar_antigas.py
    activate Importer
    Importer->>FS: Localiza arquivos Excel legados
    FS-->>Importer: Lista de caminhos de arquivos

    loop Para cada planilha histórica encontrada
        Importer->>Service: processar_planilha_historica(caminho_arquivo)
        activate Service
        Service->>Service: Inspecionar cabeçalhos e mapear layout de colunas
        
        loop Para cada linha da planilha
            Service->>Norm: normalizar_cnpj(cnpj_legado)
            Norm-->>Service: CNPJ limpo
            
            Service->>Service: Converter datas (formato serial Excel ou dd/mm/aaaa)
            Service->>Service: Mapear status legado para canônico (ex: 'DEFER.' -> 'Deferido')
            
            Service->>DB: inserir_registro_historico(dados_registro)
            activate DB
            Note over DB: INSERT INTO triagem (...)<br/>ON CONFLICT(protocolo) DO UPDATE<br/>se data for mais recente
            DB-->>Service: Registro persistido
            deactivate DB
        end
        Service-->>Importer: Sumário da planilha (processadas, inseridas, ignoradas)
        deactivate Service
    end

    Importer->>DB: reconstruir_indices_e_otimizar()
    Importer-->>Operador: Processamento concluído com resumo final no terminal
    deactivate Importer
```

---

## 3. Diagrama Estrutural de Classes e Módulos

O sistema é construído sobre módulos coesos e de baixo acoplamento organizados em três camadas lógicas:
- **`core`**: Domínio e infraestrutura fundamental (banco, normalização, enriquecimento e geração Excel).
- **`runner`**: Casos de uso de automação em lote e execução CLI (diária e histórica).
- **`admin`**: Casos de uso de gestão, governança, arquivamento e interface gráfica (GUI).

```mermaid
classDiagram
    %% Core Package
    class Config {
        <<module: core.config>>
        +DATABASE_PATH: Path
        +DATABASE_CEP_PATH: Path
        +ENTRADA_CSV_DIR: Path
        +PROCESSADOS_DIR: Path
        +PLANILHA_ATIVA_PATH: Path
        +COLUNAS_PLANILHA: List[str]
        +ESTILOS_CABECALHO: dict
        +REGRAS_CONDICIONAIS: list
        +obter_caminho_absoluto(relativo: str) Path
    }

    class Database {
        <<module: core.database>>
        +inicializar_banco() None
        +obter_conexao(caminho: Path) Connection
        +salvar_protocolos_triagem(protocolos: List[dict]) int
        +sincronizar_planilha_para_banco(caminho_excel: Path) int
        +obter_registros_ativos_ordenados() List[dict]
        +buscar_titular_anterior_cnpj(cnpj: str) Optional[str]
        +executar_transacao(query: str, params: tuple) Any
    }

    class Normalizer {
        <<module: core.normalizer>>
        -_RE_DIGITOS: Pattern
        -_RE_ESPACOS: Pattern
        -_MAPA_BAIRROS: dict
        +normalizar_cnpj(cnpj: str) str
        +normalizar_cep(cep: str) str
        +normalizar_bairro(bairro: str) str
        +normalizar_texto(texto: str) str
        +classificar_risco(cnae_primario: str, cnaes_sec: str) Tuple[str, str, str]
    }

    class CepService {
        <<module: core.cep_service>>
        +inicializar_banco_cep() None
        +resolver_area_fiscal(bairro: str, logradouro: str, cep: str) str
        +buscar_area_por_bairro(bairro: str) Optional[str]
        +buscar_area_por_cep(cep: str) Optional[str]
        +registrar_cep_pendente(cep: str, bairro: str, log: str) None
        +consultar_cache_web(cep: str) Optional[dict]
    }

    class ExcelWriter {
        <<module: core.excel_writer>>
        +gerar_planilha_controle(dados: List[dict], destino: Path) None
        +aplicar_formatacao_condicional(worksheet: Worksheet) None
        +ajustar_largura_colunas(worksheet: Worksheet) None
        +calcular_sha256_arquivo(caminho: Path) str
    }

    class HistoricoImporter {
        <<module: core.historico_importer>>
        +processar_planilha_historica(caminho: Path) dict
        +converter_data_excel(valor_celula: Any) str
        +canonizar_status(status_bruto: str) str
    }

    %% Runner Package
    class MainSeco {
        <<module: runner.main_seco>>
        +executar_pipeline_completo() bool
        +escanear_arquivos_csv(pasta: Path) List[Path]
        +processar_arquivo_csv(caminho_csv: Path) int
        +mover_para_processados(caminho_csv: Path) Path
    }

    class ImportarAntigas {
        <<module: runner.importar_antigas>>
        +executar_importacao_em_lote(pasta_origem: Path) None
    }

    %% Admin Package
    class ArchiveService {
        <<module: admin.archive_service>>
        +listar_competencias() List[dict]
        +arquivar_competencia(competencia: str, exportar: bool) int
        +desarquivar_competencia(competencia: str) int
        +excluir_competencia(competencia: str) int
        +exportar_competencia_excel(competencia: str, destino: Path) Path
    }

    class PainelAdminVRE {
        <<class: admin.main_admin.PainelAdminVRE>>
        -root: Tk
        -tree_competencias: Treeview
        -btn_arquivar: Button
        -btn_desarquivar: Button
        -btn_exportar: Button
        +__init__(master: Tk)
        +carregar_dados() None
        +on_arquivar_click() None
        +on_desarquivar_click() None
        +on_exportar_click() None
        +exibir_sobre() None
    }

    %% Relationships
    MainSeco --> Config : consome configurações
    MainSeco --> Database : persiste dados
    MainSeco --> Normalizer : limpa entradas
    MainSeco --> CepService : enriquece áreas
    MainSeco --> ExcelWriter : gera planilha ativa

    ImportarAntigas --> HistoricoImporter : orquestra
    HistoricoImporter --> Normalizer : sanitiza
    HistoricoImporter --> Database : grava histórico

    ArchiveService --> Database : manipula flags visibilidade
    ArchiveService --> ExcelWriter : exporta snapshots
    PainelAdminVRE --> ArchiveService : invoca regras de negócio
```

---

## 4. Modelo Entidade-Relacionamento e Esquema de Banco de Dados

O ecossistema utiliza dois bancos SQLite desacoplados:
1. **`triagem_vre.db`** (Banco Central de Negócio): Armazena os protocolos de viabilidade/licenciamento, metadados cadastrais, status processual e atribuição fiscal.
2. **`cep_areas.db`** (Banco Satélite de Inteligência Geográfica): Armazena as bases cartográficas, mapeamento de CEPs, logradouros, bairros e fila de pendências para aprendizado contínuo.

```mermaid
erDiagram
    TRIAGEM {
        TEXT protocolo PK "Protocolo estadual do VRE (ex: SPP2600000001)"
        TEXT data_hora "Data e hora do envio do pedido no VRE"
        TEXT competencia "Ano-Mês derivado da data (ex: 2026-08)"
        TEXT protocolo_licenciamento "Número do protocolo no sistema de licenciamento"
        TEXT tipo_processamento "Tipo: Viabilidade, Licenciamento ou Regularização"
        TEXT cnpj "CNPJ normalizado (00.000.000/0000-00)"
        TEXT razao_social "Razão social da pessoa jurídica"
        TEXT nome_fantasia "Nome comercial / de fachada"
        TEXT cnae_principal "Código e descrição do CNAE principal"
        TEXT cnaes_secundarios "Lista consolidada de CNAEs secundários"
        TEXT risco_diurno "Grau de risco diurno (Baixo, Médio, Alto)"
        TEXT risco_noturno "Grau de risco noturno (Baixo, Médio, Alto)"
        TEXT grau_risco "Grau consolidado de risco sanitário/ambiental/posturas"
        TEXT logradouro "Nome da via pública normalizada"
        TEXT numero "Número predial"
        TEXT complemento "Apartamento, sala, bloco, galpão"
        TEXT bairro "Bairro normalizado em caixa alta"
        TEXT cep "CEP no formato 00000-000"
        TEXT area_fiscal "Área atribuída (CENTRAL, HB, BOSQUE, etc.)"
        TEXT fiscal_anterior "Fiscal que atendeu o CNPJ em processos prévios"
        TEXT fiscal_designado "Fiscal selecionado para a demanda atual"
        TEXT status "Status operacional (Não Iniciado, Em Análise, etc.)"
        TEXT parecer_fiscal "Decisão formal e despacho do auditor fiscal"
        TEXT anotacoes "Observações livres da equipe técnica"
        TEXT hash_linha "Hash SHA-256 para detecção de adulteração"
        INTEGER visivel "Flag booleana (1=ativo na planilha, 0=arquivado)"
        TEXT criado_em "Timestamp ISO-8601 de criação"
        TEXT atualizado_em "Timestamp ISO-8601 da última alteração"
    }

    BASE_ENDERECOS {
        INTEGER id PK "Identificador único autoincremento"
        TEXT cep "Código postal 00000-000"
        TEXT logradouro "Nome da rua/avenida"
        TEXT bairro "Nome do bairro cadastrado"
        TEXT area_fiscal "Área fiscal correspondente"
        TEXT fonte "Origem do mapeamento (Oficial, Histórico, Manual)"
    }

    CACHE_CEP_WEB {
        TEXT cep PK "CEP normalizado (00000-000)"
        TEXT logradouro "Logradouro retornado por API web"
        TEXT bairro "Bairro retornado por API web"
        TEXT localidade "Município (São José do Rio Preto)"
        TEXT uf "Unidade federativa (SP)"
        TEXT atualizado_em "Timestamp de coleta web"
    }

    CEPS_PENDENTES {
        INTEGER id PK "Identificador único autoincremento"
        TEXT cep "CEP que resultou em PENDENTE"
        TEXT bairro "Bairro informado no protocolo"
        TEXT logradouro "Logradouro informado no protocolo"
        INTEGER ocorrencias "Contador de vezes que o CEP apareceu"
        TEXT primeira_ocorrencia "Data do primeiro registro"
        TEXT ultima_ocorrencia "Data da ocorrência mais recente"
    }

    TRIAGEM }o..o| BASE_ENDERECOS : "resolve área fiscal via"
    TRIAGEM }o..o| CACHE_CEP_WEB : "enriquece endereço via"
    TRIAGEM }o..o| CEPS_PENDENTES : "alimenta ocorrências de"
```

---

## 5. Diagramas de Transição de Estados

### 5.1. Máquina de Estados: Ciclo de Vida do Protocolo

Cada solicitação inserida na tabela `triagem` transita por uma máquina de estados finita que dita seu tratamento operacional e as regras de formatação condicional aplicadas no Excel.

```mermaid
stateDiagram-v2
    [*] --> NaoIniciado: Ingestão via CSV (Status Inicial)
    
    NaoIniciado --> EmAnalise: Fiscal assume o protocolo ou inicia conferência
    NaoIniciado --> Indeferido: Inviabilidade imediata manifesta (Zoneamento Proibido)
    
    EmAnalise --> VistoriaAgendada: Fiscal agenda vistoria presencial 'in loco'
    EmAnalise --> Deferido: Documentação e zoneamento regulares
    EmAnalise --> Indeferido: Não atendimento a posturas municipais
    EmAnalise --> Notificado: Notificação preliminar emitida ao requerente
    
    VistoriaAgendada --> Vistoriado: Vistoria presencial executada pelo fiscal
    
    Vistoriado --> Deferido: Constatada conformidade do estabelecimento
    Vistoriado --> Indeferido: Irregularidades insanáveis no imóvel
    Vistoriado --> Notificado: Constatadas inconformidades sanáveis
    
    Notificado --> EmAnalise: Requerente apresenta cumprimento de exigências
    Notificado --> Indeferido: Decurso de prazo legal sem manifestação
    
    Deferido --> Arquivado: Competência mensal arquivada pela gestão
    Indeferido --> Arquivado: Competência mensal arquivada pela gestão
    
    Arquivado --> [*]
```

### 5.2. Tabela de Transições e Regras de Negócio do Protocolo

| Estado de Origem | Evento / Gatilho | Estado de Destino | Ação do Sistema / Estilo Visual |
| :--- | :--- | :--- | :--- |
| **(Início)** | Ingestão de novo CSV | `Não Iniciado` | Linha com fundo padrão (sem destaque específico). |
| `Não Iniciado` | Atribuição manual de fiscal | `Em Análise` | Fundo Amarelo Suave (`#FFF2CC`) indicando processo em andamento. |
| `Em Análise` | Identificação de necessidade de inspeção | `Vistoria Agendada` | Fundo Azul Suave (`#D9E1F2`) alertando agendamento externo. |
| `Vistoria Agendada` | Realização da diligência em campo | `Vistoriado` | Fundo Roxo Suave (`#E2EFDA`) sinalizando vistoria concluída. |
| `Em Análise` / `Vistoriado` | Emissão de intimação ou pendência | `Notificado` | Fundo Laranja Suave (`#FCE4D6`) alertando prazo de contribuinte. |
| `Em Análise` / `Vistoriado` | Despacho favorável do auditor | `Deferido` | Fundo Verde Claro (`#E2EFDA`) e fonte verde escuro (`#375623`). |
| `Qualquer Estado Ativo` | Despacho denegatório de alvará | `Indeferido` | Fundo Vermelho Suave (`#F8CBAD`) e fonte vermelha escura (`#C65911`). |
| `Deferido` / `Indeferido` | Arquivamento via Painel Admin | `Arquivado` | `visivel = 0`, oculto da planilha ativa, migrado para snapshot. |

---

### 5.3. Máquina de Estados: Ciclo de Vida da Competência Mensal

As competências mensais (calculadas pelo campo `competencia`, ex: `2026-08`) possuem um ciclo de vida de visualização e persistência:

```mermaid
stateDiagram-v2
    [*] --> Ativa: Ingestão de protocolos do mês corrente
    
    Ativa --> PassadaVisivel: Virada do mês calendário (competência fechada, porém mantida visível)
    
    PassadaVisivel --> Arquivada: Gestor executa arquivamento via Painel Admin
    Ativa --> Arquivada: Arquivamento direto de competência (fora do ciclo ativo)
    
    Arquivada --> PassadaVisivel: Operação de Desarquivamento (Restaurar)
    
    Arquivada --> Excluida: Operação irreversível de exclusão de competência (Hard Delete)
    Excluida --> [*]
```

### 5.4. Regras de Transição da Competência

1. **Ativa (Corrente)**:
   - Todo novo protocolo do mês corrente entra automaticamente com `visivel = 1`.
   - Aparece na `Planilha_Controle_VRE_Ativa.xlsx`.
   - É editável por qualquer fiscal.

2. **Passada Visível**:
   - Competência de meses anteriores que ainda possui pendências ou vistorias ativas.
   - Mantida com `visivel = 1` para permitir consultas e conclusões.

3. **Arquivada (Oculta)**:
   - Acionada pelo `ArchiveService.arquivar_competencia()`.
   - Altera `visivel = 0` para todos os registros correspondentes.
   - Gera snapshot imutável em `Arquivados/VRE_Arquivado_AAAA-MM.xlsx`.
   - Reduz o tamanho e o tempo de abertura da planilha de trabalho ativa.

4. **Restaurada**:
   - Caso um protocolo antigo precise ser reavaliado na rotina diária, a competência pode ser reativada (`visivel = 1`), reincorporando as linhas na planilha de controle.
