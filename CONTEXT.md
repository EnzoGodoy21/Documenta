# Contexto do projeto — Documenta

Gerador de documentos `.docx` em massa a partir de um template Word.
O template define os placeholders; pasta `prints/` e arquivo de dados definem o conteúdo.
Nenhum arquivo de configuração extra é necessário para execução básica.

---

## Mapa de placeholders (confirmado)

| Tipo | Placeholder | Comportamento |
|---|---|---|
| `[TXT:chave]` | Texto automático ou via CSV | Ver tabela abaixo |
| `[IMG:chave]` | Imagem | `prints/chave_TABELA.png` (ou .jpg) |
| `[LEG:chave]` | Legenda de print | **Standby** — tipo reservado, estilo próprio a definir |

### [TXT:*] automáticos (sem configuração)

| Placeholder | Valor |
|---|---|
| `[TXT:NOME_TABELA]` | Nome da tabela |
| `[TXT:DATA]` | Data de geração: `04/04/2026` |
| `[TXT:DATA_HORA]` | Data e hora: `04/04/2026 14:30` |
| `[TXT:ANO]` | `2026` |
| `[TXT:MES]` | `04` |
| `[TXT:DIA]` | `04` |

### [IMG:*] — convenções de nome de arquivo

| Padrão | Comportamento |
|---|---|
| `chave_TABELA.png` | único print |
| `chave_1_TABELA.png`, `chave_2_TABELA.png`, ... | múltiplos numerados — um único `[IMG:chave]` no template insere todos em sequência |

Cada tabela pode ter N diferente; o template não precisa mudar.

---

### [TXT:*] via CSV (colunas extras do arquivo de tabelas)

Qualquer coluna extra no CSV de tabelas vira automaticamente um `[TXT:*]` por tabela.
O CSV serve ao mesmo tempo como filtro de execução e fonte de dados de texto.

```
nome_tabela ; RESPONSAVEL  ; DOMINIO   ; DESCRICAO
VENDAS      ; João Silva   ; Comercial ; Tabela de vendas do ERP
CLIENTES    ; Maria Santos ; CRM       ; Cadastro de clientes
```

Uso no template: `[TXT:RESPONSAVEL]`, `[TXT:DOMINIO]`, `[TXT:DESCRICAO]`

---

## Decisões de design

- **Sem CSV de controle obrigatório** — tabelas descobertas via filesystem (prints/ × template)
- **CSV de tabelas é opcional** — serve como filtro E como fonte de dados [TXT:*]
- **Template é a fonte da verdade** — define quais chaves IMG e TXT são esperadas
- **Idempotente** — pula tabelas com .docx em output/ (use --force para reprocessar)
- **Timestamp fixo por lote** — DATA/DATA_HORA iguais em todos os docs da mesma execução
- **[LEG:chave] reservado** — tipo separado de [TXT:*], estilo diferente (fonte menor), a implementar

---

## Estrutura de pastas

```
projeto/
├── CONTEXT.md               ← este arquivo
├── COMO_CRIAR_TEMPLATE.md   ← guia para criação do template
├── gera_word.py             ← script principal + CLI
├── interface.py             ← interface gráfica (tkinter)
├── requirements.txt         ← dependências Python
├── run.bat                  ← atalho Windows para abrir a interface
├── .gitignore
├── template/                ← único .docx de template
├── prints/                  ← imagens chave_TABELA.png (não versionado)
├── output/                  ← .docx gerados (não versionado)
└── logs/                    ← relatórios de execução (não versionado)
```

---

## Status de execução por tabela

| Status | Significado |
|---|---|
| `ok` | Documento gerado com todos os placeholders substituídos |
| `parcial` | Gerado, mas algum print ou texto estava ausente (placeholder visível no .docx) |
| `erro` | Falha crítica — documento não gerado |

---

## Como usar este arquivo

Ao iniciar uma nova sessão no Cowork relacionada a este projeto, diga:
> "Leia o CONTEXT.md e use como base para a tarefa"
