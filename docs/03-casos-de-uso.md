# 03 — Atores e Casos de Uso do Sistema

**Sistema**: Automação de Triagem VRE & Gestão Fiscal  
**Órgão**: Prefeitura Municipal de São José do Rio Preto — SP  
**Secretaria**: SMAURB / Inspetoria Fiscal  
**Versão**: 2.2.0 (Patch v1.6)  

---

## 1. Identificação de Atores do Sistema

O sistema atende a quatro atores principais, com responsabilidades operacionais distintas:

| Ator | Tipo | Papel e Responsabilidades no Ecossistema |
| :--- | :--- | :--- |
| **Operador de Triagem** | Humano / Operacional | Servidor público da Inspetoria Fiscal encarregado de baixar semanalmente ou diariamente o arquivo `.csv` de pedidos do portal VRE da JUCESP, acionar a triagem automatizada e disponibilizar a planilha oficial na rede interna. |
| **Inspetor / Auditor Fiscal** | Humano / Finalístico | Fiscal de posturas municipais que atua em campo. Utiliza a `Planilha_Controle_VRE_Ativa.xlsx` para acompanhar suas vistorias, alterar status de pedidos (Em Andamento, Deferido, Indeferido), preencher números de processo Sigm/OS e lançar a data de comparecimento. |
| **Gestor / Administrador** | Humano / Administrativo | Coordenador ou supervisor da fiscalização. Utiliza o Painel Administrativo (`config_vre.exe`) para resolver CEPs pendentes, consultar o histórico de CNPJs, importar planilhas de exercícios anteriores e arquivar meses concluídos para enxugar a planilha de trabalho. |
| **Agendador de Tarefas / CLI** | Sistema / Automatizado | Serviço do Windows (Task Scheduler) ou script em lote executando chamadas silenciosas sem interface gráfica para sincronização contínua. |

---

## 2. Diagrama Geral de Casos de Uso (Mermaid)

```mermaid
graph LR
    actor_op["Operador de Triagem"]
    actor_fisc["Inspetor Fiscal"]
    actor_adm["Gestor / Administrador"]
    actor_sys["Agendador / Sistema"]

    subgraph "Sistema de Automação VRE"
        UC01["UC01: Executar Triagem Diária de CSV"]
        UC02["UC02: Sincronizar Edições Manuais do Excel"]
        UC03["UC03: Pesquisar Histórico de Atendimentos"]
        UC04["UC04: Mapear e Resolver CEP x Área"]
        UC05["UC05: Controlar Arquivamento de Competências"]
        UC06["UC06: Importar Planilhas Históricas"]
        UC07["UC07: Executar Comandos Administrativos via CLI"]
    end

    actor_op --> UC01
    actor_fisc --> UC02
    actor_adm --> UC03
    actor_adm --> UC04
    actor_adm --> UC05
    actor_adm --> UC06
    actor_sys --> UC01
    actor_sys --> UC07

    UC01 ..> UC02 : <<include>>
    UC05 ..> UC02 : <<include>>
```

---

## 3. Especificação Detalhada dos Casos de Uso

---

### UC01: Executar Triagem Diária de Novos Pedidos (Headless)

- **Objetivo**: Processar os novos pedidos de licenciamento do VRE, cruzar dados com as bases de histórico, CNAE e bairros/CEPs e gerar a planilha de trabalho ativa atualizada.
- **Atores**: Operador de Triagem, Agendador / Sistema.
- **Pré-condições**:
  1. O arquivo `.csv` exportado do portal VRE está disponível na pasta `Configurações/entrada_csv/`, na raiz do projeto, ou é arrastado diretamente sobre o executável/atalho `.bat`.
  2. O banco central `triagem_vre.db` e o banco satélite `cep_areas.db` estão acessíveis.

#### Fluxo Principal:
1. O Operador inicia o processamento executando `Executar_Triagem.bat` ou arrastando o arquivo `.csv` sobre ele.
2. O sistema localiza o arquivo CSV de maior precedência.
3. **[UC02 Sincronizar Edições Manuais]**: O sistema verifica se a planilha `Planilha_Controle_VRE_Ativa.xlsx` foi modificada e salva as edições no banco.
4. O sistema lê o CSV testando os encodings em cascata (`utf-8-sig`, `utf-8`, `cp1252`, `iso-8859-1`, `latin-1`).
5. Para cada linha do CSV:
   - Higieniza o CNPJ para 14 dígitos e mascara.
   - Normaliza a data oficial do VRE (Coluna N) e calcula a competência `AAAA-MM`.
   - Consulta o histórico de fiscal titular por CNPJ.
   - Avalia a lista de CNAEs contra o cache de regras com algoritmo de curto-circuito.
   - Determina a sugestão de Tipo de Serviço e Restrição (omite ambos se a postura for Indeferida).
   - Executa a resolução geográfica em 3 níveis (Bairro normalizado $\rightarrow$ CEP mapeado $\rightarrow$ Fila de pendências).
   - Empacota os dados adicionais em formato JSON.
6. O sistema executa `INSERT OR IGNORE` no `vre_backlog`, protegendo registros já existentes.
7. O sistema executa a reconciliação automática de áreas e saneamento de indeferidos em todo o backlog.
8. O sistema realiza cópia de backup preventivo da planilha anterior em `Planilhas anteriores/`.
9. O sistema gera a nova `Planilha_Controle_VRE_Ativa.xlsx` com Plain Values e 15 regras de formatação condicional contendo todos os protocolos dos meses ativos.
10. O sistema move o CSV processado para a subpasta `Configurações/entrada_csv/processados/` com carimbo de data e hora.
11. O sistema finaliza com código de saída 0.

#### Fluxos Alternativos e Exceções:
- **FA01 — Nenhum CSV Novo Encontrado**:
  - O sistema exibe aviso no console informando que nenhum arquivo CSV novo foi detectado.
  - Sincroniza eventuais edições manuais e regenera a planilha ativa contendo os protocolos ativos existentes no banco de dados.
- **EX01 — Arquivo CSV Corrompido ou Encoding Inválido**:
  - O sistema registra o erro no arquivo `triagem_erro.log`.
  - Exibe mensagem de erro fatal e finaliza com código de saída 1 sem corromper o banco.
- **EX02 — Planilha de Saída Aberta pelo Usuário**:
  - Se a `Planilha_Controle_VRE_Ativa.xlsx` estiver aberta no Excel no momento da escrita, o sistema tenta a substituição e, caso haja bloqueio estrito de arquivo pelo Windows, emite alerta orientando o fechamento do arquivo.

#### Pós-condições:
- Planilha `Planilha_Controle_VRE_Ativa.xlsx` atualizada na raiz do projeto.
- Novos protocolos gravados no banco SQLite sem duplicações.
- Arquivo CSV arquivado em `processados/`.

---

### UC02: Sincronizar Edições Manuais do Excel com o Banco SQLite

- **Objetivo**: Ler e persistir as edições manuais feitas pelos fiscais na planilha ativa diretamente de volta para o banco de dados.
- **Atores**: Inspetor Fiscal, Operador de Triagem.
- **Pré-condições**: A planilha `Planilha_Controle_VRE_Ativa.xlsx` existe na raiz do projeto.

#### Fluxo Principal:
1. O sistema calcula o hash SHA-256 da planilha física em disco.
2. Compara com o hash registrado na tabela `controle_planilha_ativa`.
3. Se os hashes forem idênticos, o caso de uso encerra imediatamente (zero processamento).
4. Se o hash divergir:
   - Abre a planilha via OpenPyXL em modo `read_only=True` e `data_only=True`.
   - Mapeia as 32 colunas oficiais da linha 2 em diante.
   - Para cada linha:
     - Localiza o protocolo no `vre_backlog` por `protocolo_licenca` ou `protocolo_sil`.
     - Identifica se houve alteração em: `Status`, `Fiscal`, `Tipo de Serviço`, `Restrição`, `Data manual (Coluna F)`, `IPTU`, `Área`, `Protocolo Sigm / OS`, `Endereço`.
     - Preserva os valores editados no JSON de dados extras.
     - Se o fiscal titular foi atribuído/alterado, adiciona ao `historico_fiscais`.
     - Se o CEP foi associado a uma Área, registra em `cep_mapeamento` e remove de `cep_pendentes`.
5. Executa `UPDATE` em lote no `vre_backlog`.
6. Atualiza o hash SHA-256 registrado em `controle_planilha_ativa`.

#### Fluxos Alternativos e Exceções:
- **FA01 — Nova Linha Digitada Inteiramente à Mão no Excel**:
  - O sistema detecta que o protocolo de licença não existia no banco e executa `INSERT` criando o protocolo com prefixo `MANUAL_`.

#### Pós-condições:
- SQLite atualizado com todas as intervenções humanas.
- Histórico de fiscais e mapeamento de CEPs alimentados por aprendizado contínuo.

---

### UC03: Pesquisar Histórico de Atendimentos de Fiscais

- **Objetivo**: Permitir ao gestor consultar rapidamente atendimentos anteriores de empresas e fiscais titulares para dirimir dúvidas de competência e distribuição de processos.
- **Atores**: Gestor / Administrador.
- **Pré-condições**: Painel Administrativo (`config_vre.exe`) aberto na Aba 1 ("Histórico de Fiscais").

#### Fluxo Principal:
1. O Administrador digita um número de CNPJ (com ou sem máscara) ou o nome da Razão Social no campo de busca.
2. Opcionalmente, seleciona um fiscal específico no dropdown de filtro.
3. Clica em "Pesquisar" ou pressiona `Enter`.
4. O sistema executa busca no banco SQLite:
   - Se for CNPJ de 14 dígitos, busca correspondência exata em `historico_fiscais` e no `vre_backlog`.
   - Se for texto, busca por Razão Social com escape seguro de wildcards SQL.
5. Os resultados são exibidos na Treeview exibindo CNPJ formatado, Fiscal, Data do Atendimento, Origem e Localização do registro.
6. O Administrador pode clicar nos cabeçalhos das colunas para ordenar os resultados de forma ascendente ou descendente.

#### Pós-condições:
- Resultados exibidos na tela sem qualquer alteração nos dados do banco.

---

### UC04: Mapear e Resolver Correlações de CEP x Área

- **Objetivo**: Vincular um CEP a uma Área Fiscal municipal (ex: 'CENTRAL', 'HB', 'REPRESA'), resolvendo automaticamente pendências cadastrais.
- **Atores**: Gestor / Administrador.
- **Pré-condições**: Painel Administrativo aberto na Aba 2 ("Gestão de CEP x Área").

#### Fluxo Principal:
1. O Administrador visualiza a lista de "CEPs Pendentes de Mapeamento" no painel direito, ordenada pelo número de ocorrências.
2. Dá um clique no CEP desejado; o formulário de cadastro é preenchido automaticamente com o CEP e o bairro de referência.
3. O Administrador seleciona ou digita a Área Fiscal correspondente no dropdown (ex: 'BOSQUE').
4. Clica em "Salvar Mapeamento".
5. O sistema:
   - Valida os 8 dígitos do CEP e a obrigatoriedade da Área.
   - Salva a correlação no banco `cep_areas.db`.
   - Remove o CEP da tabela `cep_pendentes`.
   - Executa `backfill_area_backlog_por_cep`, atualizando a Área de todos os protocolos pendentes no backlog que possuam esse CEP.
6. O sistema emite mensagem de sucesso informando a quantidade de protocolos no backlog que foram atualizados.
7. As tabelas da tela são recarregadas automaticamente.

#### Fluxos Alternativos e Exceções:
- **FA01 — Excluir Mapeamento Existente**:
  - O Administrador seleciona um CEP na tabela de CEPs Mapeados e clica em "Excluir". O sistema solicita confirmação e remove o registro de `cep_mapeamento`.

#### Pós-condições:
- CEP mapeado com sucesso em `cep_areas.db`.
- Backlog atualizado retroativamente com a nova Área Fiscal.

---

### UC05: Controlar Visibilidade e Arquivar Meses de Competência

- **Objetivo**: Controlar quais meses de competência compõem a planilha de trabalho ativa, permitindo ocultar meses antigos para manter o arquivo ágil e enxuto.
- **Atores**: Gestor / Administrador.
- **Pré-condições**: Painel Administrativo aberto na Aba 3 ("Arquivamento de Competências").

#### Fluxo Principal:
1. O Administrador visualiza os KPIs de desempenho (Backlog total, Ativos na planilha, Pendentes, Finalizados) e a lista de meses cadastrados.
2. Cada mês exibe o status `"✅ Ativo (Visível na Planilha)"` ou `"📦 Arquivado (Oculto)"`.
3. O Administrador dá um duplo clique em um mês ou seleciona e clica em "Alternar Status / Arquivar".
4. O sistema atualiza o flag `arquivado` daquele mês na tabela `controle_competencias`.
5. O Administrador clica no botão `"💾 Aplicar Alterações na Planilha de Trabalho"`.
6. O sistema abre a janela modal de progresso em thread de background:
   - Salva backup preventivo da planilha em uso em `Planilhas anteriores/`.
   - Sincroniza edições manuais pendentes.
   - Consulta apenas os protocolos dos meses com `arquivado = 0`.
   - Reconstrói a `Planilha_Controle_VRE_Ativa.xlsx` e registra o novo hash.
7. A janela modal é fechada e uma mensagem de confirmação detalhada é apresentada.

#### Pós-condições:
- Planilha ativa regenerada contendo estritamente os meses ativos.
- Registros arquivados preservados com 100% de integridade no banco de dados.

---

### UC06: Importar em Lote Planilhas Históricas com Deduplicação

- **Objetivo**: Importar massivamente planilhas antigas em formato `.xlsx` ou `.xls` de anos e meses anteriores para popular a inteligência do sistema.
- **Atores**: Gestor / Administrador.
- **Pré-condições**: Planilhas legadas depositadas em `Planilhas anteriores/` ou selecionadas via diálogo de arquivos.

#### Fluxo Principal:
1. O Administrador clica em "📁 Importar da Pasta 'Planilhas anteriores'" ou "📥 Selecionar Planilha Excel...".
2. O sistema dispara a rotina em thread separada com tela modal de progresso.
3. Para cada arquivo:
   - Calcula o hash SHA-256 e verifica se já foi importado.
   - Lê as abas da planilha identificando dinamicamente os cabeçalhos.
   - Deduplica em memória os pares (CNPJ x Fiscal) e (CEP x Área), retendo as datas mais recentes.
   - Consolida os protocolos e insere via UPSERT no banco SQLite.
4. Concluída a leitura de todos os arquivos, reconcilia automaticamente áreas e fiscais no backlog.
5. Exibe resumo com total de arquivos, linhas brutas, duplicatas descartadas e novos dados memorizados.

#### Pós-condições:
- Banco de dados enriquecido com o histórico institucional de atendimentos fiscais.
