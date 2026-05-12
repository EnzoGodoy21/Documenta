# Como criar o template

O template é um arquivo `.docx` criado normalmente no Word.
Onde quiser inserir valores dinâmicos, escreva os placeholders abaixo.
Nenhum arquivo de configuração extra é necessário.

---

## Tipos de placeholder

| Tipo | Formato | Comportamento |
|---|---|---|
| Texto | `[TXT:chave]` | Substituído por texto (automático ou via CSV) |
| Imagem | `[IMG:chave]` | Substituído por imagem da pasta `prints/` |
| Legenda | `[LEG:chave]` | Reservado — estilo próprio, a implementar |

---

## [TXT:*] automáticos — sem configuração

```
[TXT:NOME_TABELA]   nome da tabela                      ex.: VENDAS
[TXT:DATA]          data de geração                     ex.: 04/04/2026
[TXT:DATA_HORA]     data e hora de geração              ex.: 04/04/2026 14:30
[TXT:ANO]           ano                                 ex.: 2026
[TXT:MES]           mês                                 ex.: 04
[TXT:DIA]           dia                                 ex.: 04
```

---

## [TXT:*] via CSV — dados por tabela

Adicione colunas extras no arquivo de tabelas (.csv) e use qualquer coluna
como placeholder no template. O CSV serve ao mesmo tempo como filtro de
execução e como fonte de dados de texto.

**Exemplo de CSV:**
```
nome_tabela ; RESPONSAVEL  ; DOMINIO   ; DESCRICAO
VENDAS      ; João Silva   ; Comercial ; Tabela de vendas do ERP
CLIENTES    ; Maria Santos ; CRM       ; Cadastro de clientes
```

**Uso no template:**
```
[TXT:RESPONSAVEL]   [TXT:DOMINIO]   [TXT:DESCRICAO]
```

Cada tabela recebe os valores da sua própria linha no CSV.

---

## [IMG:*] — imagens

```
[IMG:chave]   insere prints/chave_TABELA.png (ou .jpg / .jpeg)
```

**Convenção de nome dos prints:**
```
chave_NOMETABELA.png            ← único print por chave
chave_N_NOMETABELA.png          ← múltiplos prints numerados (N começa em 1)

Exemplos — único:
  visao_geral_VENDAS.png
  distribuicao_CLIENTES.png

Exemplos — múltiplos (N prints para a mesma chave):
  detalhe_1_VENDAS.png
  detalhe_2_VENDAS.png
  detalhe_3_VENDAS.png
```

Quando múltiplos prints numerados existem para uma chave, o template precisa
de apenas **um** `[IMG:detalhe]` — os demais são inseridos automaticamente
como parágrafos seguintes. Cada tabela pode ter um N diferente.

As tabelas disponíveis são descobertas automaticamente cruzando os arquivos
em `prints/` com as chaves `[IMG:*]` do template. Nenhuma lista manual é necessária.

---

## [LEG:*] — legenda de print (standby)

Tipo reservado para uso futuro. Aparecerá abaixo de prints com estilo
diferenciado (fonte menor, formatação específica). Não implementado ainda.

```
[LEG:visao_geral]   legenda vinculada ao [IMG:visao_geral]
```

---

## Exemplo de template

```
Tabela: [TXT:NOME_TABELA]              Gerado em: [TXT:DATA]
Responsável: [TXT:RESPONSAVEL]         Domínio: [TXT:DOMINIO]
Descrição: [TXT:DESCRICAO]


1. Visão Geral
[IMG:visao_geral]


2. Distribuição de Dados
[IMG:distribuicao]


3. Linhagem
[IMG:linhagem]
```

---

## Estrutura de pastas esperada

```
projeto/
├── template/
│   └── meu_template.docx
├── prints/
│   ├── visao_geral_VENDAS.png
│   ├── visao_geral_CLIENTES.png
│   ├── distribuicao_VENDAS.png
│   └── ...
├── tabelas.csv          ← opcional: filtro + dados [TXT:*]
├── output/              ← criado automaticamente
└── logs/                ← criado automaticamente
```

---

## Dicas

- Placeholders funcionam em parágrafos, títulos e células de tabela do Word
- Se um placeholder não for substituído, ele fica visível no `.docx` e o status fica `PARCIAL`
- O script é idempotente: numa segunda execução sem `--force`, tabelas já geradas são ignoradas
- O timestamp é o mesmo para todos os documentos do mesmo lote
