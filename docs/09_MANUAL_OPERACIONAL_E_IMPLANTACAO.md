# Manual de Operação e Implantação (Runbook)

**Sistema**: Automação de Triagem VRE — SMAURB / Inspetoria Fiscal  
**Versão**: 2.2.0  
**Público-Alvo**: Operadores de Triagem, Fiscais de Posturas e Administradores de TI  

---

## 1. Requisitos do Sistema

### 1.1. Estação de Trabalho do Usuário (Cliente)
- **Sistema Operacional**: Windows 10 ou 11 (64-bit) / Linux (Ubuntu 20.04+)
- **Software Adicional**: Microsoft Excel 2016 ou superior (para visualização da planilha)
- **Privilégios**: Usuário padrão (não requer privilégios de Administrador)
- **Memória RAM Mínima**: 4 GB (recomendado 8 GB)
- **Espaço em Disco**: 250 MB livres

---

## 2. Instalação e Preparação Inicial

### 2.1. Implantação Rápida (Ambiente Python Padrão)
1. Extraia o pacote em uma pasta de trabalho local ou compartilhamento de rede.
2. Abra o terminal na pasta raiz e instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```
3. Inicialize os bancos de dados SQLite e popule as tabelas normativas:
   ```bash
   python database/init_db.py --seed
   ```
4. A estrutura estará pronta para processar remessas.

---

## 3. Guia de Operação Diária

### 3.1. Processamento da Triagem Operacional (Modo Seco)
O processamento pode ser realizado de três maneiras:

#### Método A — Arraste de Arquivo (Mais Rápido e Intuitivo)
1. Localize o arquivo `.csv` baixado do VRE na sua pasta de Downloads.
2. Arraste o arquivo `.csv` e solte-o diretamente sobre o atalho **`Executar_Triagem.bat`**.
3. O terminal processará o lote em menos de 2 segundos, exibirá a confirmação em verde e fechará automaticamente.

#### Método B — Depósito na Pasta `entrada_csv/`
1. Salve o CSV baixado na pasta `entrada_csv/` (ou diretamente na raiz).
2. Dê um duplo clique no arquivo **`Executar_Triagem.bat`**.
3. O sistema detectará e processará automaticamente o CSV mais recente.

#### Método C — Execução 100% Silenciosa (Sem Terminal)
1. Dê um duplo clique em **`Executar_Triagem_Silencioso.vbs`**.
2. A triagem será executada em segundo plano de forma invisível, atualizando a planilha `Planilha_Controle_VRE_Ativa.xlsx`.

---

## 4. Guia do Painel Administrativo (`config_vre`)

Abra o painel dando um duplo clique em **`Painel_Administrativo.bat`**.

### 4.1. Aba 1 — Busca de Histórico
- **Finalidade**: Descobrir qual fiscal atendeu determinada empresa anteriormente.
- **Como Usar**:
  1. Digite o CNPJ ou parte da Razão Social no campo de pesquisa.
  2. Pressione Enter ou clique em **Pesquisar**.
  3. A listagem exibirá as vistorias anteriores ordenadas da mais recente para a mais antiga.

### 4.2. Aba 2 — Gestão de CEP x Área
- **Finalidade**: Vincular novos CEPs às respectivas Áreas fiscais.
- **Como Usar**:
  1. Digite o CEP (8 dígitos) e selecione a Área (HB, Central, Bosque, etc.).
  2. Clique em **Salvar / Atualizar Mapeamento**.
  3. O sistema automaticamente atualiza todos os protocolos do backlog que estavam pendentes desse CEP.
  4. Para inspecionar CEPs desconhecidos capturados nas remessas, clique em **Ver CEPs Pendentes**.

### 4.3. Aba 3 — Arquivamento de Competências
- **Finalidade**: Limpar a planilha de trabalho no final do mês, mantendo apenas os lotes em andamento.
- **Como Usar**:
  1. Localize os meses que já foram finalizados (ex.: `2025-05`).
  2. Altere o status para **Arquivado**.
  3. Clique em **Aplicar Alterações na Planilha**.
  4. O sistema salva uma cópia de segurança em `Planilhas anteriores/` e gera uma planilha limpa apenas com os meses ativos.

---

## 5. Resolução de Problemas (Troubleshooting)

### 5.1. Alerta de Bloqueio por Antivírus (Symantec Endpoint)
Caso o executável compilado `.exe` seja bloqueado por falso positivo de heurística:
- **Solução**: Utilize os lançadores nativos baseados em script (`Executar_Triagem.bat` e `Painel_Administrativo.bat`), que acionam o executável oficial `python.exe` assinado pela PSF, 100% aceito pelo antivírus.

### 5.2. Planilha Ativa Não Abre / Arquivo Bloqueado
- **Causa**: A planilha `Planilha_Controle_VRE_Ativa.xlsx` já está aberta no Excel por outro usuário.
- **Solução**: Feche o arquivo no Excel antes de rodar a triagem ou o arquivamento.

---

## 6. Procedimento de Backup e Recuperação

- **Bancos SQLite**: Os arquivos `triagem_vre.db` e `cep_areas.db` devem ser incluídos na rotina periódica de backup institucional do servidor.
- **Restauração**: Para restaurar o sistema, basta copiar os arquivos `.db` de volta para a pasta `database/`. O sistema restabelecerá o funcionamento imediatamente.
