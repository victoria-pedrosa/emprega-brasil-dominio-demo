# Emprega Brasil Dominio

> Projeto de portfólio de **Victória Pedrosa** (Automação, Processos e Dados). Automação desenvolvida para um escritório de contabilidade; **esta é uma versão com dados fictícios** — nomes, CNPJs, e-mails e IDs internos foram substituídos.

## Problema de negócio
Cruzar os dados de empréstimo consignado (Emprega Brasil) com a folha da Domínio era manual.

## Antes x depois
| | Antes | Depois |
|---|---|---|
| Como é feito | Conferência manual de relatórios. | Script extrai e cruza as informações e gera o resultado. |

## Ganho
- Conferência da folha mais rápida.

## Tecnologias
PDF, Python, SQLite, openpyxl, pandas

## Arquivos
- `automacao_emprestimo.py`
- `requirements.txt`

## Como rodar
1. `pip install -r requirements.txt`
2. Copie `.env.exemplo` para `.env` e preencha os caminhos.
3. Execute o script principal.

## Autora
Victória Pedrosa — Product Owner do Time de IA, automação de processos contábeis e fiscais.
