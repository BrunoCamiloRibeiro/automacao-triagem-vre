# Modelo de Casos de Uso (UML 2.5)

**Sistema**: Automação de Triagem VRE — SMAURB / Inspetoria Fiscal  
**Padrão**: UML 2.5 (OMG) / RUP Formal Use Case Specification  
**Versão**: 2.2.0  

---

## 1. Diagrama Geral de Casos de Uso

```mermaid
flowchart LR
    Operador(("👤 Operador de Triagem"))
    Fiscal(("👤 Auditor / Fiscal"))
    Admin(("👤 Administrador / Gestor"))

    subgraph SistemaVRE ["Sistema de Triagem VRE"]
        UC01["UC01: Processar Triagem Operacional<br/>(Modo Headless / Seco)"]
        UC02["UC02: Consultar Histórico por CNPJ / Razão"]
        UC03["UC03: Gerenciar Mapeamento CEP x Área"]
        UC04["UC04: Resolver Fila de CEPs Pendentes"]
        UC05["UC05: Controlar Arquivamento de Competências"]
        UC06["UC06: Importar Planilhas Legadas"]
        
        UC01_SUB1["<<include>><br/>Sanitizar Dados Cadastrais"]
        UC01_SUB2["<<include>><br/>Cruzar Histórico e CNAEs"]
        UC01_SUB3["<<include>><br/>Gerar Planilha Oficial"]
    end

    Operador --> UC01
    UC01 ..> UC01_SUB1
    UC01 ..> UC01_SUB2
    UC01 ..> UC01_SUB3

    Operador --> UC02
    Fiscal --> UC02
    Admin --> UC02

    Admin --> UC03
    Admin --> UC04
    Admin --> UC05
    Admin --> UC06
```

---

## 2. Especificação Detalhada dos Casos de Uso

### UC01: Processar Triagem Operacional (Modo Headless / Seco)
- **Ator Primário**: Operador de Triagem.
- **Objetivo**: Processar uma nova remessa de protocolos VRE de forma imediata e gerar a planilha de trabalho ativa atualizada.
- **Pré-condições**: Existência de um arquivo `.csv` válido do VRE na pasta `entrada_csv/` ou fornecido via argumento.
- **Pós-condições**: Tabela `vre_backlog` populada com novos protocolos sem sobrescrever registros já editados; arquivo `Planilha_Controle_VRE_Ativa.xlsx` gerado e pronto para uso.

#### Fluxo Principal:
1. O Operador aciona o script `Executar_Triagem.bat` (ou arrasta o arquivo CSV sobre ele).
2. O sistema detecta o caminho do CSV e identifica automaticamente o encoding do arquivo (UTF-8 / Latin-1).
3. Para cada linha do arquivo:
   - Sanitiza o CNPJ e busca eventual atendimento anterior em `historico_fiscais`.
   - Extrai os códigos CNAE em 7 dígitos e consulta flags de Restrição e Noturno em `regras_cnae` com curto-circuito.
   - Determina a classificação unificada da coluna `Restrição` ('Sem Restrição', 'Com Restrição' ou 'Noturno').
   - Normaliza o nome do Bairro via NFKD e busca a Área municipal correspondente.
   - Caso o bairro seja desconhecido, aciona fallback para `cep_areas.db`. Se o CEP também for desconhecido, registra em `cep_pendentes`.
   - Insere o registro em `vre_backlog` através de `INSERT OR IGNORE` sobre a chave primária `protocolo_licenca`.
4. O sistema consulta todos os protocolos com status ativo (`PENDENTE` ou `DISTRIBUIDO`) pertencentes a competências não arquivadas.
5. Injeta os registros como Plain Values no template oficial, aplica formatação condicional e salva atomicamente a planilha `Planilha_Controle_VRE_Ativa.xlsx`.
6. Encerra a execução retornando código de saída `0`.

#### Fluxos Alternativos e Exceções:
- **FA 1.1 — Nenhum CSV encontrado**: O sistema exibe mensagem orientando a colocação do arquivo na pasta `entrada_csv/` e encerra com aviso.
- **FE 1.2 — CSV com cabeçalhos corrompidos**: O sistema registra o erro no log e alerta o usuário sem comprometer a integridade do banco de dados.

---

### UC02: Consultar Histórico por CNPJ / Razão Social
- **Atores**: Operador de Triagem, Auditor Fiscal, Administrador.
- **Objetivo**: Localizar com agilidade atendimentos fiscais passados de um determinado contribuinte.
- **Pré-condições**: Banco `triagem_vre.db` inicializado.
- **Pós-condições**: Lista de atendimentos com CNPJ, Fiscal, Data e Origem exibida em tela.

#### Fluxo Principal:
1. O Usuário abre o Painel Administrativo (`Painel_Administrativo.bat`) e acessa a Aba 1 ("Busca de Histórico").
2. Digita um CNPJ (com ou sem pontuação) ou fragmento da Razão Social e pressiona Enter.
3. O sistema consulta simultaneamente a tabela `historico_fiscais` e o `vre_backlog`.
4. Apresenta os resultados ordenados por data decrescente na tabela interativa (Treeview).

---

### UC03: Gerenciar Mapeamento Dinâmico de CEP x Área
- **Ator Primário**: Administrador.
- **Objetivo**: Cadastrar ou atualizar a associação entre um CEP municipal e sua respectiva Área de fiscalização.
- **Pré-condições**: Acesso à Aba 2 do Painel Administrativo.
- **Pós-condições**: CEP registrado em `cep_mapeamento`; pendências do mesmo CEP removidas; protocolos pendentes com esse CEP atualizados automaticamente no backlog.

#### Fluxo Principal:
1. O Administrador acessa a Aba 2 ("Gestão de CEP x Área").
2. Preenche o CEP (8 dígitos), seleciona a Área municipal no seletor e informa o Bairro de referência.
3. Clica em "Salvar / Atualizar Mapeamento".
4. O sistema valida os campos, insere/atualiza no banco `cep_areas.db`, remove o CEP de `cep_pendentes` e executa rotina de *backfill* no `vre_backlog` para preencher áreas vazias de protocolos em aberto.
5. Exibe confirmação de sucesso com o total de protocolos pendentes atualizados.

---

### UC04: Resolver Fila de CEPs Pendentes
- **Ator Primário**: Administrador.
- **Objetivo**: Visualizar CEPs não resolvidos durante as triagens e vinculá-los às suas respectivas Áreas.
- **Pós-condições**: CEP transferido da fila de pendentes para a base oficial de mapeamento.

#### Fluxo Principal:
1. Na Aba 2, o Administrador clica em "Ver CEPs Pendentes".
2. O sistema exibe lista ordenada pelos CEPs com maior número de ocorrências nos protocolos VRE.
3. O Administrador seleciona um CEP pendente, clica em "Resolver", seleciona a Área correta e confirma.
4. O sistema processa o cadastro conforme o UC03 e atualiza a listagem em tempo real.

---

### UC05: Controlar Arquivamento de Competências
- **Ator Primário**: Administrador.
- **Objetivo**: Ocultar meses encerrados da planilha de trabalho ativa sem perder os dados cadastrais no banco SQLite.
- **Pós-condições**: Tabela `controle_competencias` atualizada; backup datado gerado em `Planilhas anteriores/`; nova planilha ativa gerada exclusivamente com os lotes ativos.

#### Fluxo Principal:
1. O Administrador acessa a Aba 3 ("Arquivamento de Competências").
2. O sistema lista todos os meses de competência registrados (`AAAA-MM`), seu status atual (Ativo/Arquivado) e volume de protocolos.
3. O Administrador marca como "Arquivado" os meses desejados e clica em "Aplicar Alterações na Planilha".
4. O sistema:
   - Salva cópia preventiva de segurança em `Planilhas anteriores/`.
   - Atualiza as flags na tabela `controle_competencias`.
   - Consulta apenas os protocolos ativos e reconstrói atomicamente a `Planilha_Controle_VRE_Ativa.xlsx`.
5. Exibe relatório com as competências ativas e total de linhas geradas.

---

### UC06: Importar e Deduplicar Planilhas Legadas
- **Ator Primário**: Administrador.
- **Objetivo**: Migrar históricos dispersos em planilhas Excel antigas para alimentar a inteligência do SQLite.
- **Pós-condições**: Pares únicos (CNPJ x Fiscal) consolidados em `historico_fiscais`; CEPs mapeados em `cep_mapeamento`.

#### Fluxo Principal:
1. O Administrador coloca arquivos Excel antigos na pasta `Planilhas anteriores/` e clica em "Importar Planilhas Antigas".
2. O sistema executa o processo em thread assíncrona com barra de progresso.
3. Para cada arquivo, normaliza datas, ignora cabeçalhos repetidos e deduplica registros.
4. Ao final, exibe relatório consolidado com total de linhas lidas, duplicatas descartadas e registros gravados.
