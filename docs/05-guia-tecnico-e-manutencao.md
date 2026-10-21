# 05 — Guia Técnico, Manutenção e Registros de Decisão Arquitetural (ADRs)

**Sistema**: Automação de Triagem VRE & Gestão Fiscal  
**Órgão**: Prefeitura Municipal de São José do Rio Preto — SP  
**Secretaria**: Secretaria Municipal de Habitação e Urbanismo (SMAURB) / Inspetoria Fiscal  
**Versão**: 2.2.0 (Patch v1.6)  
**Padrão de Engenharia**: Guia de Operações, Troubleshooting e ADRs (Architecture Decision Records)  

---

## 1. Configuração do Ambiente de Execução

O sistema foi arquitetado para operar em dois ambientes distintos:
1. **Ambiente de Produção (Windows x64 Portátil)**: Execução autônoma em estações de trabalho de fiscais e operadores, sem necessidade de privilégios de administrador ou instalação prévia de Python no sistema operacional hospedeiro.
2. **Ambiente de Desenvolvimento e Testes (Linux / POSIX ou Windows com Git)**: Ambiente completo para engenharia de software, refatoração, execução de baterias de testes automatizados e controle de versão.

---

### 1.1. Estrutura do Pacote Portátil Windows

Na máquina de produção do fiscal, a aplicação reside em uma pasta de rede ou diretório local estruturado:

```
PROJETO VRE/
├── Configurações/
│   ├── database/
│   │   ├── triagem_vre.db         # Banco principal (WAL mode)
│   │   ├── cep_areas.db           # Banco satélite de inteligência geográfica
│   │   ├── schema.sql             # Script DDL de reconstrução do banco central
│   │   └── schema_cep.sql         # Script DDL de reconstrução do banco CEP
│   ├── python/                    # CPython 3.12 x64 Embedded Distribution
│   │   ├── python.exe
│   │   ├── Lib/site-packages/     # openpyxl, et_xmlfile
│   │   └── ...
│   └── src/                       # Código-fonte Python empacotado
│       ├── core/
│       ├── runner/
│       └── admin/
├── entrada_csv/                   # Recepção diária de arquivos do VRE
│   └── processados/               # Repositório de arquivos já ingeridos
├── Planilhas anteriores/         # Armazenamento de planilhas de competências passadas e backups
├── Planilha_Controle_VRE_Ativa.xlsx # Planilha viva de trabalho diário
├── Executar_Triagem.bat           # Lançador interativo no terminal
├── Executar_Triagem_Silencioso.vbs# Lançador em segundo plano (sem janela preta)
└── Painel_Administrativo.bat      # Lançador do GUI de governança
```

### 1.2. Configuração do Ambiente de Desenvolvimento (Linux / WSL)

Para desenvolvedores atuando no ambiente Linux:

1. **Requisitos de Sistema**:
   - Python 3.12 ou superior (`python3 --version`).
   - Pip e virtualenv (opcional, ou interpretador do sistema com pacotes instalados).
   - SQLite 3.35+ com suporte a cláusulas `ON CONFLICT` e `PRAGMA wal_checkpoint`.
   - Tkinter instalado no sistema (`sudo apt-get install python3-tk`).

2. **Instalação das Dependências**:
   ```bash
   pip install openpyxl
   ```

3. **Validação Rápida**:
   ```bash
   python3 -c "import openpyxl, sqlite3, tkinter; print('Ambiente íntegro!')"
   ```

---

## 2. Guia de Execução

### 2.1. Execução da Rotina Diária de Triagem

#### Via Terminal / Linha de Comando:
```bash
# Execução padrão utilizando caminhos definidos em core.config
python3 src/runner/main_seco.py
```

O comando realiza sequencialmente:
1. Conexão e verificação de integridade dos bancos SQLite.
2. Sincronização retroativa de eventuais alterações humanas feitas na `Planilha_Controle_VRE_Ativa.xlsx`.
3. Varredura da pasta `entrada_csv/` por arquivos `.csv`.
4. Sanitização e ingestão em lote com detecção de encoding (`utf-8`, `cp1252`, `latin-1`).
5. Classificação de CNAEs e resolução de áreas fiscais via CEP/Bairro.
6. Movimentação dos CSVs processados para `entrada_csv/processados/` com carimbo de data/hora.
7. Geração e gravação da nova `Planilha_Controle_VRE_Ativa.xlsx` com formatação condicional completa.

#### Via Atalho Windows:
- **`Executar_Triagem.bat`**: Invoca o script exibindo logs coloridos no console de comando do Windows.
- **`Executar_Triagem_Silencioso.vbs`**: Invoca o `pythonw.exe` em modo invisível (ideal para agendamento de tarefas no Windows Task Scheduler).

---

### 2.2. Execução da Importação de Planilhas Históricas

Para carregar planilhas legadas de meses anteriores e popular a memória institucional:

```bash
python3 src/runner/importar_antigas.py
```

O script:
- Varre a pasta configurada em busca de arquivos Excel históricos (`.xlsx`).
- Identifica layouts de colunas, higieniza CNPJs e datas seriais do Excel.
- Mapeia status legados para os termos oficiais (`Deferido`, `Indeferido`, etc.).
- Atualiza o banco `triagem_vre.db` com o histórico de fiscais anteriores.

---

### 2.3. Execução do Painel Administrativo (GUI)

Para abrir a interface gráfica de governança de competências e arquivamento:

```bash
python3 src/admin/main_admin.py
```

**Funcionalidades disponíveis no Painel**:
- **Tabela de Competências**: Exibe todas as competências encontradas no banco, total de linhas e se estão ativas na planilha ou arquivadas.
- **Arquivar Competência**: Torna os registros invisíveis (`visivel = 0`), gera um snapshot `.xlsx` autônomo na pasta `Arquivados/` e reescreve a planilha ativa sem os registros arquivados.
- **Desarquivar Competência**: Reverte a flag (`visivel = 1`), trazendo a competência de volta à planilha ativa.
- **Exportar Competência**: Gera uma cópia da competência selecionada em qualquer pasta escolhida pelo usuário.
- **Excluir Competência**: Remove definitivamente os registros do banco após tripla confirmação.

---

## 3. Protocolo de Testes Automatizados

A garantia de qualidade e **estrita ausência de regressões** é o pilar central da engenharia do projeto.

### 3.1. Execução da Suíte Completa de Testes

Devem ser executados os dois comandos abaixo no terminal da raiz do projeto:

```bash
# 1. Bateria Unitária Automatizada (37 testes unitários)
python3 -m unittest discover tests

# 2. Bateria Integrada e Teste de Carga (8 suítes completas de integração)
python3 tests/executar_bateria_testes.py
```

### 3.2. Critério de Aprovação
- **100% dos testes devem passar com status `OK`**.
- Zero falhas (`failures=0`), zero erros (`errors=0`).
- O tempo total de execução da suíte completa não deve exceder 5 segundos.

### 3.3. Diretrizes para Novos Testes
- **Anonimização Obrigatória**: Nunca utilize nomes reais de servidores públicos em fixtures de teste. Utilize identificadores canônicos conforme padrão: `Fiscal_Designado_Manual`, `Fiscal_Gama`, `Fiscal_Delta`, etc.
- **Isolamento de Estado**: Testes que manipulam banco de dados devem utilizar bancos em memória (`:memory:`) ou criar cópias temporárias com deleção garantida no `tearDown()`.
- **Validação de Hash**: Testes que verificam geração de planilhas devem calcular o hash SHA-256 e validar a presença das 15 regras condicionais e Plain Values.

---

## 4. Diagnóstico e Resolução de Problemas (Troubleshooting)

### 4.1. Erro de Permissão ao Gerar Planilha (`PermissionError: [Errno 13]`)
- **Sintoma**: A aplicação falha ao tentar sobrescrever `Planilha_Controle_VRE_Ativa.xlsx`.
- **Causa Raiz**: O Microsoft Excel ou LibreOffice mantém um bloqueio exclusivo de escrita (*mandatory file lock*) no arquivo enquanto ele estiver aberto na tela de algum fiscal.
- **Solução**:
  1. Fechar todas as instâncias do Excel que estejam visualizando o arquivo.
  2. Executar novamente a rotina.
  3. O código foi instrumentado para capturar essa exceção preventivamente e emitir mensagem explicativa amigável no console antes de abortar a operação.

### 4.2. Bloqueio Concorrente do Banco SQLite (`sqlite3.OperationalError: database is locked`)
- **Sintoma**: Falha ao tentar gravar ou atualizar dados no banco de dados.
- **Causa Raiz**: Outro processo está executando uma transação longa ou o banco não conseguiu fechar o arquivo de journal.
- **Solução**:
  1. O sistema utiliza `PRAGMA journal_mode = WAL` e `busy_timeout = 5000` (5 segundos) para mitigar colisões.
  2. Caso o bloqueio persista, encerre processos órfãos (`python` ou `pythonw`) no Gerenciador de Tarefas do Windows ou via `kill` no Linux.
  3. Verifique se os arquivos `triagem_vre.db-wal` e `triagem_vre.db-shm` possuem permissão de escrita no diretório.

### 4.3. Falha de Encoding na Leitura do CSV
- **Sintoma**: Caracteres corrompidos (ex: `SÃ£o JosÃ©` ou erro `UnicodeDecodeError`).
- **Causa Raiz**: O portal estadual VRE exporta relatórios ora em UTF-8 (com ou sem BOM), ora em Windows-1252 / CP1252, ora em ISO-8859-1 (Latin-1).
- **Solução**: O módulo `main_seco.py` já implementa um motor de detecção multinível com tentativa sequencial `['utf-8-sig', 'utf-8', 'cp1252', 'latin-1']`. Caso um novo formato surja, certifique-se de que os bytes sejam lidos em modo binário e decodificados com tratamento `errors='replace'`.

### 4.4. CEP ou Bairro Resultando em Área "PENDENTE"
- **Sintoma**: Na planilha de controle, a coluna **Área Fiscal** fica preenchida com o valor `PENDENTE`.
- **Causa Raiz**: O loteamento ou condomínio é novo e seu nome ou faixa de CEP ainda não consta na tabela `base_enderecos` do banco satélite `cep_areas.db`.
- **Solução**:
  1. O sistema insere automaticamente o CEP na tabela `ceps_pendentes` com contador de ocorrências.
  2. O administrador pode inspecionar a tabela e adicionar a correspondência na tabela `base_enderecos`:
     ```sql
     INSERT INTO base_enderecos (cep, logradouro, bairro, area_fiscal, fonte)
     VALUES ('15000-000', 'RUA EXEMPLO', 'NOVO BAIRRO', 'CENTRAL', 'Manual');
     ```

---

## 5. Procedimento de Aplicação de Patches (GEMINI.md)

Toda evolução do sistema é empacotada em arquivo `.zip` sequencial na pasta `Patchs de atualização/`.

### 5.1. Regras do Patch
1. **Nome Sequencial Incremental**: Ex: `Patch v1.6.zip`, `Patch v1.7.zip`.
2. **Localização**: Raiz do projeto em `Patchs de atualização/`.
3. **Estrutura Interna**: Deve conter apenas os arquivos modificados preservando o caminho relativo original (`src/core/database.py`, `docs/...`).
4. **Sem Arquivos Supérfluos**: Proibida a inclusão de scripts `.bat`, arquivos `.txt` ou binários temporários.

### 5.2. Como Aplicar em Produção (Windows Portátil):
1. Abrir a pasta **`Configurações/`** da instalação portátil.
2. Copiar o conteúdo do `.zip` para dentro de `Configurações/`.
3. Confirmar a substituição de todos os arquivos coincidentes.
4. A versão estará imediatamente atualizada e pronta para uso.

---

## 6. Registros de Decisão Arquitetural (ADRs)

---

### ADR-001: Utilização de Plain Values (Valores Pré-computados) no Excel
- **Status**: Aprovado e Implantado
- **Contexto**: A versão legada utilizava fórmulas dinâmicas (`PROCV`, `SE`, `ÍNDICE/CORRESP`) em mais de 3.000 linhas por 32 colunas. Ao abrir a planilha, o Excel recalculava dezenas de milhares de células, causando travamentos de até 4 minutos e corrupção de arquivos.
- **Decisão**: Toda inteligência de negócio, sanitização de texto, cálculo de datas e resolução de áreas fiscais é executada em Python antes da geração da planilha. As células são gravadas estritamente com **valores literais (Plain Values)**.
- **Consequências**: 
  - *Positivas*: Abertura instantânea da planilha em menos de 1 segundo; tamanho do arquivo reduzido em mais de 70%; compatibilidade universal com qualquer leitor de planilhas.
  - *Negativas*: Caso o usuário altere manualmente o bairro na planilha, a área fiscal não é recalculada na célula dinamicamente até a próxima execução da sincronização bidirecional.

---

### ADR-002: Adoção do SQLite com WAL Mode e Pragmas de Performance
- **Status**: Aprovado e Implantado
- **Contexto**: Havia necessidade de um banco de dados relacional confiável, com suporte ACID e sem complexidade de infraestrutura de rede corporativa (servidores dedicados, portas liberadas, autenticação LDAP).
- **Decisão**: Adoção do SQLite 3 embarcado configurado com:
  - `PRAGMA journal_mode = WAL` (Write-Ahead Logging).
  - `PRAGMA synchronous = NORMAL`.
  - `PRAGMA foreign_keys = ON`.
  - `PRAGMA busy_timeout = 5000`.
  - `BEGIN IMMEDIATE` para blocos de escrita transacional.
- **Consequências**:
  - *Positivas*: Zero instalação ou manutenção de servidores; suporte a leituras concorrentes simultâneas sem bloqueio de escritores; proteção contra corrupção em caso de encerramento abrupto.
  - *Negativas*: Exige sistema de arquivos local ou compartilhamento de rede com suporte confiável a file locks (SMBv3).

---

### ADR-003: Pipeline de Resolução Geográfica em Fallback de 3 Níveis
- **Status**: Aprovado e Implantado
- **Contexto**: O VRE recebe cadastros de empresas onde o CEP muitas vezes é genérico (`15000-000`) ou o bairro está com grafia heterogênea (ex: `Jd. Sta. Luzia`, `Jardim Santa Luzia`, `JD STA LUZIA`).
- **Decisão**: Implementar motor de resolução em cascata estrita:
  1. **Nível 1 (Bairro Canônico)**: Busca pelo nome do bairro normalizado na base cadastral.
  2. **Nível 2 (CEP Mapeado)**: Busca pelo CEP de 8 dígitos sanitizado na tabela de logradouros.
  3. **Nível 3 (Fila de Pendências)**: Atribui status de área `PENDENTE`, não interrompe o fluxo e registra o caso na fila para aprendizado da equipe.
- **Consequências**:
  - *Positivas*: Mais de 95% de assertividade automática imediata; nenhuma interrupção de fluxo em lote por falta de endereço.
  - *Negativas*: Requer revisão periódica da tabela de pendências pelos supervisores da Inspetoria.

---

### ADR-004: Manipulação de Planilhas via `openpyxl` em Python Puro
- **Status**: Aprovado e Implantado
- **Contexto**: Avaliou-se o uso de bibliotecas que acionam a API COM do Windows (`pywin32` / Excel via OLE Automation) versus bibliotecas em Python puro (`openpyxl`, `xlsxwriter`).
- **Decisão**: Padronização exclusiva em `openpyxl`.
- **Consequências**:
  - *Positivas*: Total independência da presença do Microsoft Excel instalado na máquina; funciona de forma idêntica em servidores Linux, contêineres e clientes Windows portáteis; menor consumo de memória RAM.
  - *Negativas*: `openpyxl` é ligeiramente mais lento para carregar arquivos massivos do que extensões em C, o que foi neutralizado pelo uso de Plain Values e leitura seletiva.

---

### ADR-005: Desacoplamento do Banco Satélite `cep_areas.db`
- **Status**: Aprovado e Implantado
- **Contexto**: O banco central `triagem_vre.db` recebe alterações diárias e backups frequentes. A base cartográfica de CEPs e logradouros da cidade possui centenas de milhares de linhas e raramente é modificada.
- **Decisão**: Separar a inteligência geográfica em um arquivo SQLite satélite independente: `Configurações/database/cep_areas.db`.
- **Consequências**:
  - *Positivas*: Backups do banco de triagem são rápidos e pequenos (poucos megabytes); a base de CEPs pode ser atualizada em lote anualmente sem interferir nas transações ativas dos fiscais.
  - *Negativas*: O sistema gerencia dois arquivos de conexão distintos em vez de um único banco.

---

### ADR-006: Sincronização Bidirecional com Verificação de Integridade
- **Status**: Aprovado e Implantado
- **Contexto**: Os fiscais preferem trabalhar visualmente na planilha Excel, alterando colunas de status e despachos diretamente nas células. Havia risco de sobrescrita caso uma nova ingestão diária gerasse uma nova planilha sem incorporar as edições do dia anterior.
- **Decisão**: Implementar a função `sincronizar_planilha_para_banco()` que é acionada **antes** de qualquer geração. Ela lê a planilha em disco, compara os campos editáveis com o banco e atualiza o SQLite com as intervenções humanas.
- **Consequências**:
  - *Positivas*: Preservação absoluta das anotações humanas; o SQLite mantém-se como a única fonte da verdade (*single source of truth*); liberdade ergonômica para a equipe fiscal.
  - *Negativas*: Exige que o fiscal salve e feche a planilha antes de rodar a rotina matinal de triagem.

---

### ADR-007: Arquitetura de Distribuição Portátil com Python Embutido
- **Status**: Aprovado e Implantado
- **Contexto**: As máquinas dos auditores e fiscais da Prefeitura possuem políticas restritivas de TI, bloqueando a instalação de softwares com privilégios de administrador (`UAC`).
- **Decisão**: Fornecer a solução como pasta portátil autocontida, incluindo o interpretador CPython Embedded x64 dentro de `Configurações/python/`, acionado por arquivos de lote `.bat` com caminhos relativos.
- **Consequências**:
  - *Positivas*: Instalação zero; basta descompactar e rodar; não interfere em configurações corporativas de TI; atualizações aplicadas via substituição de arquivos por `.zip`.
  - *Negativas*: O pacote inicial descompactado possui cerca de 60 MB devido ao interpretador embutido.

---

### ADR-008: Interface Administrativa Nativa em Tkinter com Proteção de Concorrência
- **Status**: Aprovado e Implantado
- **Contexto**: Para operações de supervisão e arquivamento, era necessária uma interface gráfica amigável para usuários não técnicos.
- **Decisão**: Desenvolvimento do `PainelAdminVRE` utilizando exclusivamente a biblioteca nativa `tkinter` / `ttk`, com delegação de operações pesadas para o `archive_service.py` e diálogos modais de confirmação.
- **Consequências**:
  - *Positivas*: Nenhuma dependência externa adicional (como PyQt ou Electron); inicialização instantânea; aparência nativa Windows.
  - *Negativas*: Limitações estéticas inerentes aos widgets do Tkinter, superadas com temas nativos `ttk` (`clam`/`vista`).
