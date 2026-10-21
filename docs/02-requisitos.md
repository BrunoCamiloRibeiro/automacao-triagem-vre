# 02 — Levantamento e Especificação de Requisitos

**Padrão**: IEEE 830 / ISO/IEC/IEEE 29148  
**Sistema**: Automação de Triagem VRE & Gestão Fiscal  
**Órgão**: Prefeitura Municipal de São José do Rio Preto — SP  
**Secretaria**: SMAURB / Inspetoria Fiscal  
**Versão**: 2.2.0 (Patch v1.6)  

---

## 1. Introdução e Rastreabilidade

Este documento detalha todos os **Requisitos Funcionais (RF)** e **Requisitos Não-Funcionais (RNF)** extraídos da implementação da automação VRE, cobrindo o comportamento estrito de ponta a ponta validado pelos testes automatizados de regressão.

---

## 2. Requisitos Funcionais (RF)

### 2.1. Ingestão e Sanitização de Dados (RF01 a RF06)

| ID | Requisito Funcional | Descrição Detalhada | Regras de Validação / Exceções |
| :--- | :--- | :--- | :--- |
| **RF01** | Ingestão Flexível de CSV | O sistema deve aceitar arquivos `.csv` por três vias: (1) argumento de linha de comando `sys.argv[1]`; (2) ação de arrastar e soltar (Drag & Drop) sobre o executável ou atalho `.bat`; (3) busca automática do arquivo mais recente na raiz (`BASE_DIR`) ou na pasta `Configurações/entrada_csv/`. | Se nenhum arquivo for encontrado, o sistema processa e regenera a planilha ativa mantendo todos os protocolos ativos já cadastrados no banco. |
| **RF02** | Autodetecção em Cascata de Encodings | O motor de leitura de CSV deve tentar abrir o arquivo sequencialmente usando as seguintes codificações: `utf-8-sig`, `utf-8`, `cp1252`, `iso-8859-1` e `latin-1`, com delimitador `;`. | Se nenhuma codificação decodificar o arquivo sem erros, uma exceção `ValueError` deve ser levantada e registrada em `triagem_erro.log`. |
| **RF03** | Sanitização Rigorosa de CNPJ | Toda entrada de CNPJ deve ser convertida para uma sequência estrita de 14 dígitos numéricos puros. O algoritmo deve tratar: valores `float` do Excel (ex: `44444444000144.0`), strings com terminação `.0`, pontuações e strings entre 11 e 13 dígitos completando zeros à esquerda via `zfill(14)`. | Valores fora do intervalo de 11 a 14 dígitos ou vazios devem retornar string vazia `""`, prevenindo cadastros inválidos. |
| **RF04** | Formatação Oficial de CNPJ | O sistema deve aplicar a máscara `XX.XXX.XXX/XXXX-XX` sobre os 14 dígitos normalizados. | Se a string não contiver exatamente 14 dígitos, o valor original deve ser preservado sem corrupção. |
| **RF05** | Higienização de CEP | O sistema deve higienizar o CEP para 8 dígitos numéricos puros, recuperando eventuais zeros à esquerda suprimidos pelo Excel (suportando comprimentos de 7 e 8 dígitos via `zfill(8)`). | Entradas com menos de 7 ou mais de 8 dígitos devem retornar string vazia `""`. |
| **RF06** | Fatiamento e Extração de CNAEs | O sistema deve extrair códigos CNAE informados em strings com 9 dígitos ou separadas por vírgula/ponto-e-vírgula, fatiando cada código para seus 7 primeiros dígitos numéricos base (classe/subclasse) e eliminando duplicatas mantendo a ordem original. | Códigos inválidos ou com menos de 7 dígitos numéricos devem ser ignorados. |

---

### 2.2. Regras de Negócio e Zoneamento Fiscal (RF07 a RF12)

| ID | Requisito Funcional | Descrição Detalhada | Regras de Validação / Exceções |
| :--- | :--- | :--- | :--- |
| **RF07** | Atribuição por Histórico Fiscal | O sistema deve consultar o banco `triagem_vre.db` e atribuir automaticamente o fiscal titular registrado para aquele CNPJ limpo em atendimentos anteriores (ordenados por `data_atendimento DESC, id DESC`), buscando em `historico_fiscais` e residualmente no backlog ativo. | Nomes genéricos ou de pendência (ex: 'N/A', 'Pendente', 'Sem Fiscal') são descartados pela lista `TERMOS_INVALIDOS_FISCAL`. |
| **RF08** | Avaliação com Curto-Circuito de CNAEs | Ao analisar a lista de CNAEs de uma empresa contra a tabela `regras_cnae`, o algoritmo deve avaliar as flags `restricao` e `noturno`. Caso ambas atinjam o valor `1` (`flag_restr = 1` e `flag_noturno = 1`), o loop de busca é interrompido imediatamente (*short-circuit*). | Otimiza o processamento em lote evitando consultas supérfluas no dicionário de cache. |
| **RF09** | Classificação em Coluna Única de Restrição | A Coluna 5 (E) da planilha deve receber a classificação sintetizada: (1) `"Noturno"` se `flag_noturno = 1`; (2) `"Com Restrição"` se `flag_restricao = 1` e `flag_noturno = 0`; (3) `"Sem Restrição"` nos demais casos. | Preserva qualquer valor customizado informado manualmente pelo operador (`RestricaoCustom`). |
| **RF10** | Sugestão e Tipificação de Serviço | Para protocolos deferidos ou em andamento, o campo `Tipo de Serviço` (Coluna D) deve seguir as regras: (1) Se for noturno com fiscal: `"Noturno"`; (2) Se for noturno sem fiscal: `"Noturno Livre"`; (3) Se possuir fiscal atribuído e não for noturno: `"Diurno Livre"`; (4) Se não tiver fiscal e tiver restrição: `"Com Restrição Livre"`; (5) Se não tiver fiscal e não tiver restrição: `"Sem Restrição Livre"`. | Caso o usuário tenha editado para termos específicos (ex: 'Noturno entregue', 'Cota', 'VP'), o valor customizado é rigorosamente preservado. |
| **RF11** | Omissão Estrita em Protocolos Indeferidos | Quando a postura for indeferida (`SituacaoPostura == 'INDEFERIDO'` ou `Status == 'Indeferido'`), o sistema deve deixar a coluna `Tipo de Serviço` (D) e a coluna `Restrição` (E) estritamente vazias. | Aplica-se tanto a novos pedidos do CSV quanto a registros passados do backlog através da função `limpar_indeferidos_e_normalizar_backlog`. |
| **RF12** | Isolamento da Coluna F (Data) | A Coluna 6 (F - Data) deve ser mantida vazia por padrão para novas atribuições da fiscalização, registrando apenas datas preenchidas manualmente pelo operador no Excel (`DataCol6`). A data oficial emitida pelo sistema estadual VRE deve constar exclusivamente na Coluna 14 (N - Data VRE). | O sistema executa uma migração única registrada em `controle_migracoes` para expurgar datas automáticas antigas da Coluna F. |

---

### 2.3. Resolução Geográfica e Banco Satélite de CEP (RF13 a RF16)

| ID | Requisito Funcional | Descrição Detalhada | Regras de Validação / Exceções |
| :--- | :--- | :--- | :--- |
| **RF13** | Nível 1: Resolução por Bairro Normalizado | O sistema deve normalizar o nome do bairro aplicando decomposição Unicode NFKD, remoção de acentos e expansão de abreviações consagradas (`JD` $\rightarrow$ `JARDIM`, `PQ` $\rightarrow$ `PARQUE`, `VL` $\rightarrow$ `VILA`, etc.), consultando a tabela `bairros_areas` do banco central. | Letras isoladas como 'V' ou 'J' não são abreviadas para evitar corromper numerais romanos (ex: 'DAMHA V', 'LIBERDADE IV'). |
| **RF14** | Nível 2: Fallback para Banco de CEPs | Se o bairro não for encontrado ou estiver em branco, o sistema consulta indexadamente o banco satélite `database/cep_areas.db` buscando a Área cadastrada para aquele CEP de 8 dígitos. | Se localizado, retorna o nome da Área harmonizado (ex: 'SCHIMITT' $\rightarrow$ 'SCHMITT', 'CIDADE CRIANÇA' $\rightarrow$ 'CIDADE DA CRIANÇA'). |
| **RF15** | Nível 3: Fila de CEPs Pendentes | Se nem Bairro nem CEP resolverem a Área, e o CEP for válido com 8 dígitos, o sistema cadastra ou incrementa o contador `ocorrencias` na tabela `cep_pendentes`, registrando bairro e logradouro de referência. | A Área do protocolo no backlog permanece vazia `""` para que o fiscal saiba que requer atenção. |
| **RF16** | Retroalimentação Automática (*Backfill*) | Assim que um operador cadastra a área de um CEP pendente no painel administrativo, o sistema varre o backlog e preenche retroativamente a Área de todos os protocolos pendentes que contenham aquele mesmo CEP. | Respeita a flag `AreaLimpaManualmente`: se o fiscal intencionalmente apagou a área de um protocolo no Excel, o backfill não reescreve a decisão. |

---

### 2.4. Persistência, Sincronização e Geração Excel (RF17 a RF21)

| ID | Requisito Funcional | Descrição Detalhada | Regras de Validação / Exceções |
| :--- | :--- | :--- | :--- |
| **RF17** | Proteção do Backlog (`INSERT OR IGNORE`) | Novos pedidos recebidos no CSV são inseridos na tabela `vre_backlog` com cláusula `INSERT OR IGNORE INTO vre_backlog` indexada pela chave primária `protocolo_licenca`. | Garante que reprocessar o mesmo arquivo CSV jamais sobrescreva status, fiscal atribuído ou notas inseridas pela equipe. |
| **RF18** | Sincronização Bidirecional via Hash SHA-256 | Antes de processar qualquer CSV ou alterar competências, o sistema compara o hash SHA-256 da planilha física em disco contra o hash registrado na tabela `controle_planilha_ativa`. Se houver divergência, lê todas as edições manuais e atualiza o banco de dados. | Se os hashes forem idênticos e `forcar=False`, encerra a sincronização em 0 milissegundos sem I/O desnecessário. |
| **RF19** | Geração com Plain Values e Zero Travamento | A planilha oficial `Planilha_Controle_VRE_Ativa.xlsx` deve ser gerada injetando exclusivamente valores estáticos de texto/número puro nas 32 colunas oficiais (A a AF), sem fórmulas lentas de busca. | Elimina travamentos de recálculo no Excel e reduz o tempo de abertura do arquivo de 40 segundos para menos de 1 segundo. |
| **RF20** | Formatação Condicional na Hierarquia Exata | As 15 regras de formatação condicional devem ser aplicadas em um único range contínuo (`A3:AF5000`), garantindo que status prioritários (Em Andamento verde, Indeferido rosa/vermelho) sobreponham cores de fiscal ou serviço. | As regras seguem a ordem estrita: `$B3="Em Andamento"` $\rightarrow$ `$B3="Indeferido"` ... $\rightarrow$ `$C3<>""` $\rightarrow$ `$D3="Noturno Livre"` $\rightarrow$ `$D3="Sem Restrição Livre"`. |
| **RF21** | Backups Preventivos Datados | Antes de sobrescrever a planilha de trabalho ativa, o sistema deve copiar o arquivo existente para a pasta `Planilhas anteriores/` com carimbo de data e hora no formato `Planilha_Controle_VRE_Ativa_backup_AAAAMMDD_HHMMSS.xlsx`. | Caso o arquivo esteja temporariamente bloqueado por outro programa, tenta cópia resiliente por stream de bytes binários. |

---

### 2.5. Arquivamento e Interface Gráfica (RF22 a RF26)

| ID | Requisito Funcional | Descrição Detalhada | Regras de Validação / Exceções |
| :--- | :--- | :--- | :--- |
| **RF22** | Arquivamento de Competências Mensais | O sistema deve agrupar protocolos por mês de competência (`mes_competencia` no formato `AAAA-MM`) e permitir ao gestor alternar o status entre `"Ativo"` (visível na planilha) e `"Arquivado"` (oculto da planilha). | Os dados arquivados nunca são apagados do SQLite, podendo ser reativados a qualquer momento pelo usuário. |
| **RF23** | Painel Administrativo com 3 Abas | O executável `config_vre` deve disponibilizar interface gráfica nativa em Tkinter com abas dedicadas para: (1) Histórico de Fiscais; (2) Gestão de CEP x Área; (3) Arquivamento de Competências. | Suporte a atalhos de teclado (Enter para pesquisar) e ordenação interativa de colunas ao clicar no cabeçalho. |
| **RF24** | Multithreading com Modal de Progresso | Operações de importação ou reorganização pesada devem ser executadas em `threading.Thread(daemon=True)` exibindo janela modal com barra de progresso indeterminada (`Progressbar`). | Previne que o sistema operacional marque a janela como "Não Respondendo" e sincroniza callbacks via `root.after(0, ...)`. |
| **RF25** | Carga e Deduplicação de Planilhas Antigas | O módulo deve permitir importar planilhas históricas em lote varrendo `Planilhas anteriores/` ou arquivo avulso, deduplicando pares (CNPJ x Fiscal) e (CEP x Área) e retendo as datas de atendimento mais recentes. | Registra o hash do arquivo em `historico_arquivos_importados` para evitar duplicidade de processamento. |
| **RF26** | Modo CLI Headless Alternativo | Caso a interface gráfica não possa ser carregada (ex: terminal SSH ou servidor Linux), o executável `config_vre` deve aceitar argumentos de linha de comando (`--cli`, `--buscar`, `--arquivar`, `--reativar`, `--aplicar`, `--mapear-cep`, `--listar-pendentes`). | Permite automação completa via scripts `.bat` ou rotinas agendadas no Cron/Task Scheduler. |

---

## 3. Requisitos Não-Funcionais (RNF)

| ID | Categoria | Requisito Não-Funcional | Métrica / Critério de Aceitação |
| :--- | :--- | :--- | :--- |
| **RNF01** | Desempenho | Tempo de Processamento de Lote | A triagem de um arquivo CSV diário com até 1.000 linhas deve ser concluída em **menos de 2,5 segundos** em máquina padrão de escritório (Intel Core i3 / 8GB RAM). |
| **RNF02** | Confiabilidade | Integridade Transacional ACID | Todas as operações de escrita devem ser executadas dentro do context manager `db_transaction` com `BEGIN IMMEDIATE` e `COMMIT/ROLLBACK` automáticos. |
| **RNF03** | Portabilidade | Operação Sem Instalação de Administrador | O software deve rodar a partir de pendrive ou pasta do usuário (`AppData` ou `Documentos`) sem exigir permissões de administrador do Windows (`UAC`) nem instalação de dependências globais. |
| **RNF04** | Usabilidade | Operação por Duplo Clique / Drag & Drop | O usuário operacional não deve precisar digitar comandos no terminal. O processamento deve ocorrer por duplo clique nos arquivos `.bat` ou arrastando o arquivo `.csv` sobre o ícone. |
| **RNF05** | Eficiência | Otimização de Memória e WAL Mode | O banco SQLite deve operar com `PRAGMA journal_mode = WAL`, `PRAGMA synchronous = NORMAL` e `PRAGMA cache_size = -64000` (64MB RAM), permitindo leituras simultâneas. |
| **RNF06** | Manutenibilidade | Clean Code e Padrão de Docblocks | 100% dos módulos, classes e funções devem possuir docstrings completas especificando Propósito, Parâmetros, Retorno, Exceções e Complexidade Big-O. |
| **RNF07** | Segurança | Proteção de Dados (LGPD) e Anonimização | Os testes automatizados e repositórios de código aberto (`projeto_vre_github`) devem utilizar nomes fictícios e anonimizados de fiscais (ex: `Fiscal_Gama`, `Fiscal_Delta`), preservando o sigilo de servidores públicos. |
| **RNF08** | Resiliência | Prevenção de Injeção SQL em Buscas Textuais | Todas as buscas textuais por `LIKE` no SQLite devem tratar e escapar caracteres especiais (`%`, `_`) com cláusula `ESCAPE '\\'`. |
| **RNF09** | Compatibilidade | Versões de Microsoft Excel Suportadas | As planilhas geradas pelo OpenPyXL devem abrir perfeitamente em Microsoft Office Excel 2010, 2013, 2016, 2019, 2021, Office 365 e LibreOffice Calc. |
| **RNF10** | Rastreabilidade | Política de Patches Sequenciais | Qualquer atualização de código deve gerar um pacote compactado em `Patchs de atualização/Patch v1.X.zip` contendo apenas os arquivos integrais modificados, mantendo as pastas relativas. |
