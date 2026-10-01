# Guia de Contribuição — Automação de Triagem VRE

Agradecemos o interesse em contribuir com o projeto de Automação de Triagem do Via Rápida Empresa (VRE)!

Este documento estabelece as diretrizes para desenvolvimento, boas práticas de código, abertura de issues e submissão de Pull Requests.

---

## 🧭 Código de Conduta

Ao participar deste projeto, você concorda em seguir o nosso [Código de Conduta](CODE_OF_CONDUCT.md).

---

## 🛠️ Configuração do Ambiente de Desenvolvimento

### 1. Clonar o Repositório
```bash
git clone https://github.com/usuario-ou-orgao/automacao-triagem-vre.git
cd automacao-triagem-vre
```

### 2. Criar e Ativar Ambiente Virtual
```bash
python -m venv .venv

# Linux / macOS:
source .venv/bin/activate

# Windows (PowerShell):
.venv\Scripts\Activate.ps1
```

### 3. Instalar Dependências de Desenvolvimento
```bash
pip install -r requirements-dev.txt
```

### 4. Inicializar os Bancos SQLite Locais
```bash
python database/init_db.py --seed
```

---

## 🧪 Execução de Testes Automatizados

Antes de submeter qualquer código, garanta que todos os testes unitários e de integração estejam passando:

```bash
# Execução da suíte padrão unittest:
python -m unittest discover -s tests -p "test_*.py" -v

# Execução da bateria completa de cenários de ponta a ponta:
python tests/executar_bateria_testes.py

# Ou utilizando pytest com relatório:
pytest
```

---

## 📐 Padrões de Código e Estilo

- **PEP 8**: Siga estritamente as diretrizes da PEP 8.
- **Tipagem Estática**: Utilize anotações de tipo (`type hints`) em todas as funções e métodos.
- **Docstrings**: Mantenha docstrings detalhadas no padrão Google/NumPy explicando entradas, saídas e algoritmos.
- **Linhas e Formatação**: Limite de 100 caracteres por linha.
- **Linter e Formatador**: Execute o `ruff` ou `black`:
  ```bash
  ruff check .
  black --check .
  ```

---

## 🔒 Diretrizes de Privacidade e LGPD

1. **Nenhum Dado Real de Contribuinte**: Nunca versione arquivos `.csv`, `.xlsx` ou bancos SQLite contendo dados reais de pessoas físicas, jurídicas ou servidores municipais.
2. Utilize exclusivamente dados sintéticos e fixtures de teste localizados em `tests/test_data/`.
3. Verifique sempre o `git status` e `.gitignore` antes de efetuar commits.

---

## 🚀 Fluxo de Submissão (Pull Requests)

1. Crie uma branch para sua funcionalidade ou correção:
   ```bash
   git checkout -b feature/nome-da-feature
   # ou
   git checkout -b fix/nome-do-bug
   ```
2. Adicione testes para cobrir suas alterações.
3. Garanta que 100% dos testes passem localmente.
4. Escreva commits claros no padrão *Conventional Commits*:
   - `feat: adiciona suporte a nova regra de CNAE`
   - `fix: corrige parsing de complemento de endereço`
   - `docs: atualiza especificações de arquitetura`
5. Abra o Pull Request descrevendo claramente a motivação e a solução implementada.
