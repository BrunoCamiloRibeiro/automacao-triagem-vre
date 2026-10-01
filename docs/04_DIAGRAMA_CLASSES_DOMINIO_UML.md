# Modelo Estrutural e Diagrama de Classes (UML 2.5)

**Sistema**: Automação de Triagem VRE — SMAURB / Inspetoria Fiscal  
**Padrão**: UML 2.5 Class Model (OMG)  
**Versão**: 2.2.0  

---

## 1. Diagrama de Classes da Arquitetura Completa

```mermaid
classDiagram
    direction TB

    %% =========================================================================
    %% CAMADA DE DOMÍNIO E ENTIDADES
    %% =========================================================================
    class VreProtocol {
        +String protocolo_licenca
        +String protocolo_sil
        +String cnpj_limpo
        +String cnpj_formatado
        +String nome_empresarial
        +String mes_competencia
        +String data_importacao
        +String bairro_original
        +String area
        +String cnaes_licenciadas
        +String fiscal_atribuido
        +int flag_restricao
        +int flag_noturno
        +String tipo_servico
        +String status
        +String dados_extras_json
        +to_excel_row() List
    }

    class CnaeRule {
        +String cnae_7digitos
        +String cnae_formatado
        +String denominacao
        +int restricao
        +int noturno
    }

    class NeighborhoodArea {
        +int id
        +String bairro_normalizado
        +String area
    }

    class CepMapping {
        +String cep_limpo
        +String cep_formatado
        +String area
        +String bairro_referencia
        +String origem
        +DateTime atualizado_em
    }

    class PendingCep {
        +String cep_limpo
        +String cep_formatado
        +String bairro_referencia
        +String logradouro
        +int ocorrencias
        +DateTime criado_em
    }

    class FiscalHistory {
        +int id
        +String cnpj_limpo
        +String fiscal
        +String data_atendimento
        +String origem_registro
    }

    class CompetenceControl {
        +String mes_competencia
        +int arquivado
        +DateTime atualizado_em
    }

    %% =========================================================================
    %% CAMADA DE SERVIÇOS E LÓGICA DE NEGÓCIO
    %% =========================================================================
    class NormalizerService {
        <<utility>>
        +normalizar_cnpj(cnpj_bruto: Any) String
        +formatar_cnpj(cnpj_limpo: String) String
        +normalizar_cep(cep_bruto: Any) String
        +formatar_cep(cep_limpo: String) String
        +normalizar_bairro(nome_bairro: String) String
        +extrair_cnaes(cnaes_str: String) List~String~
        +formatar_cnae(cnae_7: String) String
        +montar_endereco_completo(...) String
    }

    class CepService {
        <<service>>
        +consultar_area_por_cep(cep: String) String
        +registrar_area_cep(...) bool
        +registrar_cep_pendente(...) void
        +resolver_area(...) Tuple~String, String~
        +reconciliar_areas_backlog(...) int
        +listar_ceps_mapeados(...) List~Dict~
        +listar_ceps_pendentes(...) List~Dict~
    }

    class ExcelWriterService {
        <<service>>
        +criar_template_padrao(caminho: Path) Path
        +gerar_planilha_trabalho(registros: List, saida: Path) Path
        -aplicar_regras_condicionais(ws: Worksheet, ref: String) void
    }

    class ArchiveService {
        <<service>>
        +definir_status_competencia(mes: String, arquivado: bool) void
        +listar_competencias_status() List~Dict~
        +buscar_historico_fiscais(termo: String) List~Dict~
        +obter_estatisticas_gerais() Dict~String, int~
        +aplicar_reorganizacao_planilha(saida: Path) Dict
    }

    class HistoricoImporterService {
        <<service>>
        +importar_planilha_historico(caminho: Path, forcar: bool) Dict
        +escanear_e_importar_planilhas_antigas() Dict
        +normalizar_data_atendimento(val: Any) String
    }

    %% =========================================================================
    %% CAMADA DE INFRAESTRUTURA E ACESSO A DADOS
    %% =========================================================================
    class DatabaseManager {
        <<infrastructure>>
        +connect_db(path: Path, pragmas: List) Connection
        +get_triagem_db(path: Path) Connection
        +get_cep_db(path: Path) Connection
        +db_transaction(conn: Connection) ContextManager
        +initialize_schema_if_needed(conn: Connection, is_cep: bool) void
    }

    class ConfigManager {
        <<configuration>>
        +Path BASE_DIR
        +Path CONFIGURACOES_DIR
        +Path TRIAGEM_DB_PATH
        +Path CEP_DB_PATH
        +Path OUTPUT_XLSX_PATH
        +Path TEMPLATE_XLSX_PATH
        +List CSV_ENCODINGS
        +ensure_directories() void
    }

    %% =========================================================================
    %% CAMADA DE APRESENTAÇÃO
    %% =========================================================================
    class HeadlessRunner {
        <<runner>>
        +main() void
        +processar_fluxo_seco(caminho_csv: Path, saida_xlsx: Path) int
        -ler_linhas_csv(caminho: Path) List~Dict~
        -localizar_arquivo_csv() Path
    }

    class AdminPanelGUI {
        <<presentation>>
        -Tk root
        -Notebook notebook
        +atualizar_tudo() void
        -criar_aba_historico() void
        -criar_aba_ceps() void
        -criar_aba_competencias() void
        -executar_busca_historico() void
        -salvar_mapeamento_cep() void
        -aplicar_reorganizacao() void
    }

    %% =========================================================================
    %% RELACIONAMENTOS E DEPENDÊNCIAS
    %% =========================================================================
    HeadlessRunner ..> DatabaseManager : utiliza
    HeadlessRunner ..> NormalizerService : sanitiza com
    HeadlessRunner ..> CepService : resolve áreas via
    HeadlessRunner ..> ExcelWriterService : gera relatório via
    HeadlessRunner o-- VreProtocol : processa

    AdminPanelGUI ..> ArchiveService : gerencia competências via
    AdminPanelGUI ..> CepService : gerencia CEPs via
    AdminPanelGUI ..> HistoricoImporterService : importa via

    ArchiveService ..> DatabaseManager : persiste em
    ArchiveService ..> ExcelWriterService : reconstrói via
    CepService ..> DatabaseManager : persiste em
    HistoricoImporterService ..> DatabaseManager : persiste em
    HistoricoImporterService ..> NormalizerService : higieniza com

    DatabaseManager ..> ConfigManager : obtém caminhos e pragmas
