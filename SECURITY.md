# Política de Segurança e Proteção de Dados (LGPD)

## 🛡️ Versões Suportadas

Atualmente, apenas as seguintes versões recebem atualizações de segurança e correções normativas:

| Versão | Suporte Ativo |
| :--- | :--- |
| `2.2.x` | :white_check_mark: Sim |
| `< 2.2.0` | :x: Não |

---

## 🔒 Relato de Vulnerabilidades

A segurança e a integridade de dados públicos e cadastrais são tratadas com a máxima prioridade.

Caso você identifique uma vulnerabilidade de segurança, falha de isolamento ou potencial vazamento de dados:

1. **NÃO crie uma Issue pública** para relatar vulnerabilidades de segurança.
2. Envie um e-mail confidencial para a equipe de tecnologia responsável:
   - `suporte.fiscalizacao@riopreto.sp.gov.br` (ou canal institucional configurado).
3. Inclua no relato:
   - Descrição detalhada da vulnerabilidade;
   - Passos reprodutíveis ou script de prova de conceito (PoC);
   - Impacto estimado e possíveis mitigações.

Comprometemo-nos a responder e validar o relato em até **5 dias úteis**.

---

## 📜 Conformidade com a LGPD (Lei Geral de Proteção de Dados - Lei nº 13.709/2018)

- **Sanitização de Repositório**: Este repositório de código-fonte aberto não contém nem deve conter dados pessoais de contribuintes, munícipes ou servidores (CPFs, telefones pessoais, históricos reais de vistorias).
- **Armazenamento Local e Seguro**: Em ambientes de produção municipal, o banco SQLite armazena dados operacionais exclusivamente na máquina autorizada ou diretório de rede protegido com controle rigoroso de ACLs (Access Control Lists).
- **Tratamento de Dados Cadastrais**: As informações de CNPJs e licenças empresariais processadas decorrem do cumprimento de obrigação legal e exercício regular de direitos no âmbito da fiscalização tributária e de posturas (Art. 7º, II e III da LGPD).
