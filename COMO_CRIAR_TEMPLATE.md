# Como criar o template

O template é um arquivo `.docx` criado normalmente no Word.
Onde quiser inserir valores dinâmicos, escreva os placeholders abaixo.
O script substitui tudo automaticamente — nenhum arquivo de configuração é necessário.

---

## Placeholders disponíveis

### Texto automático (sem configuração)

| Placeholder | O que insere |
|---|---|
| `[TXT:NOME_TABELA]` | Nome da tabela (ex.: `VENDAS`) |
| `[TXT:DATA]` | Data de geração: `04/04/2026` |
| `[TXT:DATA_HORA]` | Data e hora: `04/04/2026 14:30` |
| `[TXT:ANO]` | Apenas o ano: `2026` |
| `[TXT:MES]` | Apenas o mês: `04` |
| `[TXT:DIA]` | Apenas o dia: `04` |

### Texto de dados externos (via `dados.csv`)

Qualquer coluna do arquivo de dados pode virar um placeholder.

| Placeholder | O que insere |
|---|---|
| `[TXT:RESPONSAVEL]` | Coluna `RESPONSAVEL` do `dados.csv` |
| `[TXT:DOMINIO]` | Coluna `DOMINIO` do `dados.csv` |
| `[TXT:DESCRICAO]` | Coluna `DESCRICAO` do `dados.csv` |
| `[TXT:...]` | Qualquer outra coluna |

### Imagens

| Placeholder | O que insere |
|---|---|
| `[IMG:visao_geral]` | `prints/visao_geral_TABELA.png` |
| `[IMG:distribuicao]` | `prints/distribuicao_TABELA.png` |
| `[IMG:chave]` | `prints/chave_TABELA.png` (ou `.jpg`) |

---

## Convenção de nome dos prints

```
chave_NOMETABELA.png
```

Exemplos:
```
visao_geral_VENDAS.png
distribuicao_CLIENTES.png
linhagem_PEDIDOS.png
```

As tabelas disponíveis são descobertas automaticamente cruzando os
arquivos em `prints/` com as chaves `[IMG:*]` do template.

---

## Exemplo de template

```
Tabela: [TXT:NOME_TABELA]              Documento gerado em: [TXT:DATA]
Responsável: [TXT:RESPONSAVEL]         Domínio: [TXT:DOMINIO]
Descrição: [TXT:DESCRICAO]


Visão Geral
[IMG:visao_geral]


Distribuição de Dados
[IMG:distribuicao]


Linhagem
[IMG:linhagem]
```

---

## Estrutura de pastas esperada

```
projeto/
├── template/
│   └── meu_template.docx     ← template com os placeholders
├── prints/
│   ├── visao_geral_VENDAS.png
│   ├── visao_geral_CLIENTES.png
│   ├── distribuicao_VENDAS.png
│   └── ...
├── output/                   ← .docx gerados (criado automaticamente)
├── logs/                     ← relatórios de execução (criado automaticamente)
├── gera_word.py
├── interface.py
└── run.bat
```

---

## Dicas

- Placeholders funcionam dentro de **títulos, parágrafos e células de tabela** do Word
- Se um placeholder não for substituído (print ausente, coluna faltando no CSV), ele **fica visível** no `.docx` gerado e o status fica como `PARCIAL` — fácil de identificar no painel de resultados
- O script é **idempotente**: numa segunda execução sem `--force`, tabelas que já têm `.docx` em `output/` são ignoradas
- Textos externos via `dados.csv` ainda não estão implementados — em desenvolvimento
