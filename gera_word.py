"""
gera_word.py — Gerador de documentos Word por tabela

Placeholders suportados no template:
  [TXT:NOME_TABELA]  → substituído pelo nome da tabela (texto)
  [TXT:outra_chave]  → substituído pelo valor correspondente (texto genérico)
  [IMG:chave]        → substituído por prints/chave_TABELA.png (imagem)

Uso:
  python gera_word.py                            # execução normal
  python gera_word.py --force                    # reprocessa tudo
  python gera_word.py --tabelas VENDAS CLIENTES  # só essas tabelas
  python gera_word.py --workers 8                # paralelismo
  python gera_word.py --template outro.docx      # template específico
"""

import os
import re
import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────

def iter_paragrafos(doc):
    """Itera sobre todos os parágrafos do documento, incluindo dentro de tabelas."""
    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p


def _substituir_texto_no_paragrafo(p, busca, substituto):
    """
    Substitui texto em um parágrafo, tratando casos onde o placeholder
    está dividido entre múltiplos runs.
    """
    if busca not in p.text:
        return

    # Caso simples: placeholder dentro de um único run
    for run in p.runs:
        if busca in run.text:
            run.text = run.text.replace(busca, substituto)
            return

    # Caso complexo: placeholder dividido entre runs — mescla tudo no primeiro
    texto_completo = "".join(r.text for r in p.runs)
    if busca in texto_completo:
        novo_texto = texto_completo.replace(busca, substituto)
        if p.runs:
            p.runs[0].text = novo_texto
            for run in p.runs[1:]:
                run.text = ""


def _substituir_imagem_no_paragrafo(p, placeholder, caminho_imagem):
    """
    Substitui um placeholder de imagem pelo arquivo de imagem correspondente.
    Centraliza e remove indentação.
    """
    # Limpa todos os runs
    for run in p.runs:
        run.text = ""

    # Zera indentação
    pPr = p._p.get_or_add_pPr()
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), "0")
    ind.set(qn("w:right"), "0")
    pPr.append(ind)

    # Centraliza
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Insere imagem
    run = p.add_run()
    run.add_picture(caminho_imagem, width=Inches(5.5))


# ─────────────────────────────────────────────────────────────
# AUTO-DESCOBERTA
# ─────────────────────────────────────────────────────────────

def descobrir_chaves(template_path: str) -> dict:
    """
    Escaneia o template e retorna as chaves encontradas separadas por tipo.

    Retorna:
      {
        "img": ["chave1", "chave2", ...],   # [IMG:chave]
        "txt": ["chave1", "chave2", ...],   # [TXT:chave]
      }

    Exemplos:
      [IMG:visao_geral]   → img: 'visao_geral'
      [TXT:NOME_TABELA]   → txt: 'NOME_TABELA'
    """
    doc = Document(template_path)
    chaves_img = set()
    chaves_txt = set()

    for p in iter_paragrafos(doc):
        chaves_img.update(re.findall(r"\[IMG:([^\]]+)\]", p.text))
        chaves_txt.update(re.findall(r"\[TXT:([^\]]+)\]", p.text))

    return {
        "img": sorted(chaves_img),
        "txt": sorted(chaves_txt),
    }


def descobrir_tabelas(prints_path: str, chaves_img: list) -> list:
    """
    Descobre os nomes de tabelas cruzando os arquivos em prints/
    com as chaves de imagem encontradas no template.

    Convenção: chave_NOMETABELA.png
    Tabelas são inferidas apenas via [IMG:*] — não via [TXT:*].
    """
    pasta = Path(prints_path)
    if not pasta.exists():
        return []

    tabelas = set()
    extensoes = {".png", ".jpg", ".jpeg"}

    for arquivo in pasta.iterdir():
        if arquivo.suffix.lower() not in extensoes:
            continue
        nome = arquivo.stem
        for chave in chaves_img:
            prefixo = chave + "_"
            if nome.startswith(prefixo):
                tabela = nome[len(prefixo):]
                if tabela:
                    tabelas.add(tabela)

    return sorted(tabelas)


def ler_tabelas_de_arquivo(path: str) -> list:
    """
    Lê nomes de tabelas de um arquivo .txt ou .csv.

    .txt — uma tabela por linha; linhas em branco e comentários (#) ignorados.
    .csv — lê a coluna 'nome_tabela' se existir; caso contrário usa a primeira coluna.
           Suporta delimitadores vírgula (,) e ponto-e-vírgula (;).
    """
    import csv as _csv

    ext = Path(path).suffix.lower()
    tabelas = []

    with open(path, encoding="utf-8-sig") as f:
        if ext == ".csv":
            amostra = f.read(2048)
            f.seek(0)
            try:
                dialeto = _csv.Sniffer().sniff(amostra, delimiters=",;")
            except _csv.Error:
                dialeto = _csv.excel
            reader = _csv.DictReader(f, dialect=dialeto)
            col = next((c for c in (reader.fieldnames or [])
                        if c.strip().lower() == "nome_tabela"), None)
            if col is None and reader.fieldnames:
                col = reader.fieldnames[0]
            for row in reader:
                if col:
                    val = row.get(col, "").strip()
                    if val:
                        tabelas.append(val)
        else:
            for linha in f:
                val = linha.strip()
                if val and not val.startswith("#"):
                    tabelas.append(val)

    return tabelas


def resolver_template(template_arg: str) -> str:
    """
    Resolve o caminho do template — aceita arquivo direto ou pasta
    (retorna o primeiro .docx encontrado).
    """
    p = Path(template_arg)
    if p.is_file():
        return str(p)
    if p.is_dir():
        candidatos = sorted(p.glob("*.docx"))
        if candidatos:
            return str(candidatos[0])
    raise FileNotFoundError(f"Template não encontrado: {template_arg}")


# ─────────────────────────────────────────────────────────────
# PROCESSAMENTO
# ─────────────────────────────────────────────────────────────

def _builtin_textos(tabela: str, timestamp: datetime = None) -> dict:
    """
    Valores de texto automáticos disponíveis em todo documento, sem necessidade de CSV.

    Placeholder         Exemplo de valor
    ──────────────────────────────────────────────────
    [TXT:NOME_TABELA]   VENDAS
    [TXT:DATA]          04/04/2026
    [TXT:DATA_HORA]     04/04/2026 14:30
    [TXT:ANO]           2026
    [TXT:MES]           04
    [TXT:DIA]           04
    """
    ts = timestamp or datetime.now()
    return {
        "NOME_TABELA": tabela,
        "DATA":        ts.strftime("%d/%m/%Y"),
        "DATA_HORA":   ts.strftime("%d/%m/%Y %H:%M"),
        "ANO":         ts.strftime("%Y"),
        "MES":         ts.strftime("%m"),
        "DIA":         ts.strftime("%d"),
    }


def processar_tabela(tabela, template_path, prints_path, output_path, chaves,
                     textos=None, callback=None, timestamp=None):
    """
    Processa uma tabela: abre o template, faz substituições e salva o .docx.

    Parâmetros:
      chaves     — dict {"img": [...], "txt": [...]} retornado por descobrir_chaves()
      textos     — dict opcional com valores vindos de CSV/dados externos
                   ex.: {"RESPONSAVEL": "Enzo", "DOMINIO": "Comercial"}
      timestamp  — datetime fixo para DATA/DATA_HORA (passado pelo caller para
                   garantir que todos os docs do mesmo lote tenham o mesmo valor).
                   Se None, usa datetime.now() no momento da chamada.

    Valores automáticos (sempre disponíveis, sem CSV):
      [TXT:NOME_TABELA]  [TXT:DATA]  [TXT:DATA_HORA]  [TXT:ANO]  [TXT:MES]  [TXT:DIA]

    Retorna um dict com:
      - tabela: nome da tabela
      - status: 'ok' | 'parcial' | 'erro'
      - prints_ok: chaves [IMG:*] inseridas com sucesso
      - prints_ausentes: chaves [IMG:*] sem arquivo correspondente
      - textos_ausentes: chaves [TXT:*] sem valor definido
      - erro: mensagem de erro (None se sem erro)
    """
    textos = textos or {}

    # Builtins primeiro, textos externos sobrescrevem se houver conflito
    valores_txt = {**_builtin_textos(tabela, timestamp), **textos}

    resultado = {
        "tabela": tabela,
        "status": "ok",
        "prints_ok": [],
        "prints_ausentes": [],
        "textos_ausentes": [],
        "erro": None,
    }

    try:
        doc = Document(template_path)

        for p in iter_paragrafos(doc):

            # ── substituição de texto [TXT:chave] ──────────────
            for chave in chaves.get("txt", []):
                placeholder = f"[TXT:{chave}]"
                if placeholder not in p.text:
                    continue
                if chave in valores_txt:
                    _substituir_texto_no_paragrafo(p, placeholder, valores_txt[chave])
                else:
                    # Mantém o placeholder e registra ausência
                    resultado["textos_ausentes"].append(chave)
                    resultado["status"] = "parcial"

            # ── substituição de imagens [IMG:chave] ────────────
            for chave in chaves.get("img", []):
                placeholder = f"[IMG:{chave}]"
                if placeholder not in p.text:
                    continue

                caminho = _achar_print(prints_path, chave, tabela)
                if caminho:
                    _substituir_imagem_no_paragrafo(p, placeholder, caminho)
                    resultado["prints_ok"].append(chave)
                else:
                    # Mantém o placeholder visível para facilitar identificação
                    resultado["prints_ausentes"].append(chave)
                    resultado["status"] = "parcial"

        # Salva o documento
        os.makedirs(output_path, exist_ok=True)
        caminho_saida = Path(output_path) / f"{tabela}.docx"
        doc.save(str(caminho_saida))

    except Exception as e:
        resultado["status"] = "erro"
        resultado["erro"] = str(e)

    if callback:
        callback(resultado)

    return resultado


def _achar_print(prints_path, chave, tabela):
    """Procura o arquivo de print para uma chave+tabela. Retorna o caminho ou None."""
    for ext in (".png", ".jpg", ".jpeg"):
        candidato = Path(prints_path) / f"{chave}_{tabela}{ext}"
        if candidato.exists():
            return str(candidato)
    return None


# ─────────────────────────────────────────────────────────────
# RELATÓRIO
# ─────────────────────────────────────────────────────────────

def gerar_relatorio(resultados, logs_path) -> str:
    """Gera um relatório .txt com o resumo da execução. Retorna o caminho do arquivo."""
    os.makedirs(logs_path, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    caminho = Path(logs_path) / f"relatorio_{timestamp}.txt"

    ok      = [r for r in resultados if r["status"] == "ok"]
    parcial = [r for r in resultados if r["status"] == "parcial"]
    erro    = [r for r in resultados if r["status"] == "erro"]

    with open(caminho, "w", encoding="utf-8") as f:
        f.write(f"Relatório de execução — {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
        f.write("=" * 60 + "\n\n")
        f.write(f"Total processado : {len(resultados)}\n")
        f.write(f"  OK             : {len(ok)}\n")
        f.write(f"  Parcial        : {len(parcial)}\n")
        f.write(f"  Erro           : {len(erro)}\n\n")

        if parcial:
            f.write("── PARCIAIS ──\n")
            for r in parcial:
                ausencias = []
                if r["prints_ausentes"]:
                    ausencias.append(f"IMG ausentes: {', '.join(r['prints_ausentes'])}")
                if r["textos_ausentes"]:
                    ausencias.append(f"TXT ausentes: {', '.join(r['textos_ausentes'])}")
                f.write(f"  {r['tabela']}: {' | '.join(ausencias)}\n")
            f.write("\n")

        if erro:
            f.write("── ERROS ──\n")
            for r in erro:
                f.write(f"  {r['tabela']}: {r['erro']}\n")
            f.write("\n")

        f.write("── DETALHES ──\n")
        for r in sorted(resultados, key=lambda x: x["tabela"]):
            f.write(f"  [{r['status'].upper():7}] {r['tabela']}\n")

    return str(caminho)


# ─────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────

EPILOG = """
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 COMO CRIAR O TEMPLATE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

O template é um arquivo .docx normal criado no Word.
Onde quiser inserir valores dinâmicos, use os placeholders:

  [TXT:NOME_TABELA]   → nome da tabela (automático)
  [TXT:DATA]          → data de geração: 04/04/2026
  [TXT:DATA_HORA]     → data e hora:     04/04/2026 14:30
  [TXT:ANO]           → apenas o ano:    2026
  [TXT:MES]           → apenas o mês:    04
  [TXT:DIA]           → apenas o dia:    04
  [TXT:coluna]        → qualquer coluna do dados.csv (ex.: [TXT:RESPONSAVEL])
  [IMG:chave]         → imagem da pasta prints/ com nome chave_TABELA.png

Exemplo de template:
  Tabela: [TXT:NOME_TABELA]          Documento gerado em: [TXT:DATA]
  Responsável: [TXT:RESPONSAVEL]     Domínio: [TXT:DOMINIO]

  Visão Geral
  [IMG:visao_geral]

  Distribuição de Dados
  [IMG:distribuicao]

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 CONVENÇÃO DE NOMES DOS PRINTS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  chave_NOMETABELA.png   (ou .jpg / .jpeg)

  Exemplos:
    visao_geral_VENDAS.png
    distribuicao_CLIENTES.png

As tabelas disponíveis são descobertas automaticamente
cruzando os arquivos em prints/ com as chaves do template.
Não é necessário nenhum arquivo de configuração.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 EXEMPLOS DE USO
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  # Execução padrão (processa só tabelas sem .docx em output/)
  python gera_word.py

  # Reprocessar tudo do zero
  python gera_word.py --force

  # Tabelas específicas por nome
  python gera_word.py --tabelas VENDAS CLIENTES PEDIDOS

  # Tabelas por arquivo (um nome por linha)
  python gera_word.py --tabelas-arquivo minhas_tabelas.txt

  # Template e pastas customizados
  python gera_word.py --template docs/meu_template.docx --prints capturas/ --output gerados/

  # Mais paralelismo para grandes volumes
  python gera_word.py --workers 8

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 STATUS DOS DOCUMENTOS GERADOS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  ok       Documento gerado com todos os placeholders substituídos
  parcial  Documento gerado, mas algum print ou texto estava ausente
           (o placeholder original fica visível no .docx para identificação)
  erro     Falha crítica — documento não foi gerado
"""


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Gera documentos .docx em massa a partir de um template Word,\n"
            "substituindo placeholders [TXT:*] e [IMG:*] por tabela.\n"
            "Use --help para ver exemplos e instruções de template."
        ),
        epilog=EPILOG,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--template", default="template/",
                        metavar="CAMINHO",
                        help="Arquivo .docx de template ou pasta que o contém  (padrão: template/)")
    parser.add_argument("--prints",   default="prints/",
                        metavar="PASTA",
                        help="Pasta com os prints chave_TABELA.png             (padrão: prints/)")
    parser.add_argument("--output",   default="output/",
                        metavar="PASTA",
                        help="Pasta de saída dos .docx gerados                 (padrão: output/)")
    parser.add_argument("--logs",     default="logs/",
                        metavar="PASTA",
                        help="Pasta para relatórios de execução                (padrão: logs/)")
    parser.add_argument("--force",    action="store_true",
                        help="Reprocessa tabelas que já têm .docx em output/")
    parser.add_argument("--tabelas",         nargs="+", metavar="TABELA",
                        help="Processa apenas as tabelas listadas (nomes separados por espaço)")
    parser.add_argument("--tabelas-arquivo", metavar="ARQUIVO",
                        help="Arquivo .txt ou .csv com nomes de tabelas, um por linha")
    parser.add_argument("--workers",  type=int, default=4,
                        metavar="N",
                        help="Número de workers paralelos                      (padrão: 4)")
    args = parser.parse_args()

    # Resolve template
    try:
        template_path = resolver_template(args.template)
    except FileNotFoundError as e:
        print(f"❌ {e}")
        return

    print(f"📄 Template : {template_path}")

    # Auto-descoberta
    chaves  = descobrir_chaves(template_path)
    tabelas = descobrir_tabelas(args.prints, chaves["img"])

    print(f"🖼️  IMG chaves: {chaves['img'] if chaves['img'] else '(nenhuma)'}")
    print(f"📝 TXT chaves: {chaves['txt'] if chaves['txt'] else '(nenhuma)'}")
    print(f"📊 Tabelas  : {len(tabelas)} encontrada(s)")

    # Filtros
    filtro = set()
    if args.tabelas:
        filtro.update(args.tabelas)
    if args.tabelas_arquivo:
        try:
            do_arquivo = ler_tabelas_de_arquivo(args.tabelas_arquivo)
            filtro.update(do_arquivo)
            print(f"📂 Arquivo  : {len(do_arquivo)} tabela(s) carregada(s) de {args.tabelas_arquivo}")
        except Exception as e:
            print(f"❌ Erro ao ler arquivo de tabelas: {e}")
            return
    if filtro:
        tabelas = [t for t in tabelas if t in filtro]
        print(f"🔍 Filtro   : {len(tabelas)} tabela(s) selecionada(s)")

    if not args.force:
        pendentes  = [t for t in tabelas if not (Path(args.output) / f"{t}.docx").exists()]
        ignoradas  = len(tabelas) - len(pendentes)
        if ignoradas:
            print(f"⏭️  Ignoradas: {ignoradas} já processada(s) (use --force para reprocessar)")
        tabelas = pendentes

    if not tabelas:
        print("✅ Nada a processar.")
        return

    # Timestamp fixo para todos os docs do lote — DATA/DATA_HORA consistentes
    ts = datetime.now()
    print(f"\n🚀 Processando {len(tabelas)} tabela(s) com {args.workers} worker(s)...")
    print(f"📅 Timestamp : {ts.strftime('%d/%m/%Y %H:%M:%S')}\n")

    resultados = []

    def on_done(resultado):
        resultados.append(resultado)
        icon = {"ok": "✅", "parcial": "⚠️ ", "erro": "❌"}.get(resultado["status"], "?")
        linha = f"{icon} [{resultado['status'].upper():7}] {resultado['tabela']}"
        if resultado["prints_ausentes"]:
            linha += f"  — IMG ausentes: {', '.join(resultado['prints_ausentes'])}"
        if resultado["textos_ausentes"]:
            linha += f"  — TXT ausentes: {', '.join(resultado['textos_ausentes'])}"
        if resultado["erro"]:
            linha += f"  — erro: {resultado['erro']}"
        print(linha)

    with ThreadPoolExecutor(max_workers=args.workers) as executor:
        futuros = {
            executor.submit(processar_tabela, t, template_path, args.prints,
                            args.output, chaves, None, on_done, ts): t
            for t in tabelas
        }
        for f in as_completed(futuros):
            pass  # resultados tratados no callback

    caminho_rel = gerar_relatorio(resultados, args.logs)

    ok      = sum(1 for r in resultados if r["status"] == "ok")
    parcial = sum(1 for r in resultados if r["status"] == "parcial")
    erro    = sum(1 for r in resultados if r["status"] == "erro")

    print(f"\n{'=' * 40}")
    print(f"✅ OK: {ok}  ⚠️  Parcial: {parcial}  ❌ Erro: {erro}")
    print(f"📋 Relatório: {caminho_rel}")


if __name__ == "__main__":
    main()
