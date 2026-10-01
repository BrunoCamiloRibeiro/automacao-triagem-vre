# Changelog

Todas as alterações notáveis neste projeto serão documentadas neste arquivo.

O formato baseia-se em [Keep a Changelog](https://keepachangelog.com/pt-BR/1.0.0/) e este projeto adere ao [Versionamento Semântico](https://semver.org/lang/pt-BR/).

---

## [2.2.0] - 2026-10-01

### Adicionado
- Estruturação modular limpa e desacoplada em subpacotes `core`, `runner` e `admin`.
- Suporte abrangente a configuração por variáveis de ambiente (`.env` e `os.getenv`).
- Documentação formal clássica de Engenharia de Software com padrão UML 2.5 completo (Casos de Uso, Classes, Sequência, Atividades, Estados, Componentes, Implantação e MER).
- Seeding com suporte automático a fallback demonstrativo sintético (`historico_demonstrativo.json`), permitindo testes e compilação sem dependência de dados reais sensíveis.
- Pipeline de integração contínua (GitHub Actions) com matriz de testes para Python 3.10, 3.11 e 3.12.
- Nova coluna consolidada de **Restrição** (coluna 5 da planilha) com as três classificações oficiais: 'Sem Restrição', 'Com Restrição' e 'Noturno'.

### Segurança & Higienização
- Sanitização total de dados sensíveis e pessoais (nomes de servidores e dados fiscais substituídos por padrões genéricos em testes e documentação).
- Remoção de comunicações corporativas privadas e planilhas legadas de contribuintes.
- Criação de templates de issues, pull requests, política de segurança (LGPD) e diretrizes de contribuição.

### Otimizado
- Pipeline de gravação Excel com injeção pura de Plain Values, eliminando 100% dos travamentos por fórmulas dinâmicas e referências circulares.
- Algoritmo de curto-circuito de CNAEs atingindo tempo de execução subsegundo.
- Resolução de Bairro com decomposição Unicode NFKD e fallback dinâmico para CEP.

---

## [2.1.0] - 2026-09-25

### Adicionado
- Modo headless operacional silencioso (`main_seco.py` e lançador VBS invisível).
- Banco satélite dedicado `cep_areas.db` com fila de resolução de pendências.
- Suporte a distribuição portátil com interpretador oficial embutido para eliminação de falsos positivos de antivírus corporativos.

---

## [2.0.0] - 2026-09-22

### Adicionado
- Painel Administrativo em Tkinter nativo (`config_vre`).
- Gestão e arquivamento de competências mensais com visibilidade controlada na planilha ativa e preservação perene no SQLite.
- Algoritmo de herança hierárquica normativa de CNAEs (Seção -> Divisão -> Grupo -> Classe -> Subclasse).
