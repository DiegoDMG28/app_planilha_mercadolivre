"""Importador de vendas do Mercado Livre -> banco_vendas.xlsx

Uso:
    python app_vendas.py                      (abre uma janela para escolher os arquivos)
    python app_vendas.py vendas1.xlsx ...     (sem janela)

Requer: pip install openpyxl
O banco (banco_vendas.xlsx) fica na mesma pasta deste arquivo.
"""
import os
import re
import sys
import unicodedata
import warnings
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Font, PatternFill

warnings.filterwarnings("ignore")

BANCO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "banco_vendas.xlsx")

# ---------------------------------------------------------------------------
# REGRAS para identificar modelo / cor / estampa a partir do título do anúncio.
# Cada regra: (texto a procurar, nome que vai para a tabela). Edite à vontade.
# ---------------------------------------------------------------------------
MODELOS = [("bolsinho", "Bag Bolsinho"), ("laminad", "Bag Laminada")]
MODELO_PADRAO = "Bag"
CORES = [("vermelh", "Vermelha"), ("pret", "Preta"), ("azul", "Azul"), ("cinza", "Cinza"),
         ("rosa", "Rosa"), ("branc", "Branca"), ("verd", "Verde"), ("amarel", "Amarela"),
         ("laranj", "Laranja"), ("roxo", "Roxa"), ("marrom", "Marrom")]
ESTAMPAS = [(r"camuflad", "Camuflada"), (r"ifood|\b99\b", "iFood")]
ESTAMPA_PADRAO = "Lisa"

# Tabela de códigos: nome -> (coluna do código na aba Códigos, nº de dígitos)
TABELAS = {"modelo": (1, 2), "cor": (4, 2), "estampa": (7, 1)}

CAB_VENDAS = ["N.º de venda", "Data", "Estado", "Situação", "ID produto", "Produto", "Unidades",
              "Valor da venda", "Envio", "Descontos", "Taxa ML", "Reembolsos", "Repasse ML",
              "Custo unit.", "CMV", "Lucro"]
CAB_PROD = ["ID", "Modelo", "Cor", "Estampa", "Produto", "Custo de produção", "Anúncio ML",
            "Título do anúncio"]
CAB_COD = ["Cód. modelo", "Modelo", "", "Cód. cor", "Cor", "", "Cód. estampa", "Estampa"]

BRL = '"R$" #,##0.00'
HDR_FILL = PatternFill("solid", fgColor="1F3864")
IN_FILL = PatternFill("solid", fgColor="FFF2CC")
SEC_FILL = PatternFill("solid", fgColor="D9E1F2")
MESES = {m: i for i, m in enumerate(
    ["janeiro", "fevereiro", "marco", "abril", "maio", "junho", "julho", "agosto",
     "setembro", "outubro", "novembro", "dezembro"], 1)}


def fonte(**kw):
    return Font(name="Arial", size=kw.pop("size", 10), **kw)


def norm(s):
    s = unicodedata.normalize("NFKD", str(s if s is not None else ""))
    return "".join(c for c in s if not unicodedata.combining(c)).lower().strip()


def chave(s):
    return re.sub(r"[^a-z0-9#]", "", norm(s))


def num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("R$", "")
    if not s:
        return None
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def parse_data(v):
    if isinstance(v, datetime):
        return v
    m = re.search(r"(\d{1,2}) de (\w+) de (\d{4})(?:\D+(\d{1,2}):(\d{2}))?", norm(v))
    if not m or m.group(2) not in MESES:
        return None
    return datetime(int(m.group(3)), MESES[m.group(2)], int(m.group(1)),
                    int(m.group(4) or 0), int(m.group(5) or 0))


# ---------------------------------------------------------------------------
# Leitura do relatório do Mercado Livre
# ---------------------------------------------------------------------------
PREFIXOS = {  # campo -> início do cabeçalho normalizado (1ª ocorrência na planilha)
    "id": "nodevenda", "data": "datadevenda", "estado": "estado", "unid": "unidades",
    "valor": "receitaporprodutos", "taxa": "tarifadevendaeimpostos",
    "rec_envio": "receitaporenvio", "tar_envio": "tarifasdeenvio",
    "env_med": "custodeenviocombase", "env_dif": "custopordiferencas",
    "desc": "descontosebonus", "reemb": "cancelamentoserembolsos", "total": "total",
    "anuncio": "#deanuncio", "titulo": "titulodoanuncio", "variacao": "variacao",
}


def ler_relatorio(caminho):
    ws = load_workbook(caminho, data_only=True).worksheets[0]
    linhas = list(ws.iter_rows(values_only=True))
    ini = next((i for i, r in enumerate(linhas[:40]) if any(chave(c) == "nodevenda" for c in r)), None)
    if ini is None:
        raise ValueError("não parece um relatório de vendas do Mercado Livre (coluna 'N.º de venda' não encontrada)")
    cab = [chave(c) for c in linhas[ini]]
    col = {}
    for campo, pref in PREFIXOS.items():
        exato = campo in ("estado", "unid", "id", "data")
        col[campo] = next((i for i, c in enumerate(cab) if (c == pref if exato else c.startswith(pref))), None)
    faltando = [c for c in ("id", "data", "estado", "unid", "valor", "taxa", "total", "titulo") if col[c] is None]
    if faltando:
        raise ValueError("colunas não encontradas: " + ", ".join(faltando))

    def g(r, campo):
        i = col.get(campo)
        return r[i] if i is not None and i < len(r) else None

    vendas = []
    for r in linhas[ini + 1:]:
        vid = g(r, "id")
        if vid is None or not str(vid).strip():
            continue
        vid = str(int(vid)) if isinstance(vid, float) else str(vid).strip()
        if not vid.isdigit():
            continue
        total = num(g(r, "total"))
        tem = total is not None
        n = lambda c: num(g(r, c)) or 0.0
        vendas.append({
            "id": vid, "data": parse_data(g(r, "data")), "estado": str(g(r, "estado") or "").strip(),
            "unid": num(g(r, "unid")),
            "valor": num(g(r, "valor")) if tem else None,
            "envio": -(n("rec_envio") + n("tar_envio") + n("env_med") + n("env_dif")) if tem else None,
            "desc": -n("desc") if tem else None,
            "taxa": -n("taxa") if tem else None,
            "reemb": -n("reemb") if tem else None,
            "repasse": total,
            "anuncio": str(g(r, "anuncio") or "").strip(),
            "titulo": str(g(r, "titulo") or "").strip(),
            "variacao": str(g(r, "variacao") or "").strip(),
        })
    return vendas


def identificar(titulo, variacao):
    t, v = norm(titulo), norm(variacao)

    def cor(txt):
        achados = [(m.start(), nome) for stem, nome in CORES for m in re.finditer(r"\b" + stem + r"\w*", txt)]
        return max(achados)[1] if achados else None

    modelo = next((n for p, n in MODELOS if p in t), MODELO_PADRAO)
    estampa = next((n for p, n in ESTAMPAS if re.search(p, t)), ESTAMPA_PADRAO)
    if "sem logo" in t:
        estampa += " Sem Logo"
    return modelo, cor(v) or cor(t) or "Sem cor", estampa


# ---------------------------------------------------------------------------
# Banco (Excel)
# ---------------------------------------------------------------------------
def cabecalho(ws, nomes, larguras):
    for i, (n, w) in enumerate(zip(nomes, larguras), 1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = w
        if n:
            c = ws.cell(1, i, n)
            c.font = fonte(bold=True, color="FFFFFF")
            c.fill = HDR_FILL
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"


def criar_banco():
    wb = Workbook()
    v = wb.active
    v.title = "Vendas"
    cabecalho(v, CAB_VENDAS, [20, 17, 22, 12, 11, 34, 9, 14, 12, 12, 12, 12, 14, 12, 12, 14])
    v.auto_filter.ref = "A1:P1"
    p = wb.create_sheet("Produtos")
    cabecalho(p, CAB_PROD, [10, 18, 14, 22, 44, 18, 16, 60])
    c = wb.create_sheet("Códigos")
    cabecalho(c, CAB_COD, [13, 24, 3, 13, 18, 3, 13, 24])
    for r in range(2, 301):
        for k in (1, 4, 7):
            c.cell(r, k).number_format = "@"
    for tabela, cod, nome in (("modelo", "01", "Bag Bolsinho"), ("cor", "02", "Vermelha"), ("estampa", "1", "iFood")):
        k = TABELAS[tabela][0]
        c.cell(2, k, cod)
        c.cell(2, k + 1, nome)
    return wb


def codigo(cods, tabela, nome):
    """Devolve o código do nome; se não existir, cria o próximo código livre."""
    k, dig = TABELAS[tabela]
    existentes, r = {}, 2
    while cods.cell(r, k).value not in (None, ""):
        existentes[norm(cods.cell(r, k + 1).value)] = str(cods.cell(r, k).value).zfill(dig)
        r += 1
    if norm(nome) in existentes:
        return existentes[norm(nome)]
    usados = {int(x) for x in existentes.values() if x.isdigit()}
    n = next(i for i in range(1, 10 ** 6) if i not in usados)
    cod = str(n).zfill(dig)
    c = cods.cell(r, k, cod)
    c.number_format = "@"
    cods.cell(r, k + 1, nome)
    return cod


def formulas_venda(r):
    prod = f"MATCH($E{r},Produtos!$A:$A,0)"
    custo = f"INDEX(Produtos!$F:$F,{prod})"
    return {
        4: f'=IF(M{r}="","Sem dados",IF(OR(ISNUMBER(SEARCH("devol",C{r})),ISNUMBER(SEARCH("cancel",C{r}))),"Devolvida","Válida"))',
        6: f'=IFERROR(INDEX(Produtos!$E:$E,{prod}),"")',
        14: f'=IFERROR(IF(ISNUMBER({custo}),{custo},""),"")',
        15: f'=IF(D{r}="Devolvida",0,IF(OR(D{r}="Sem dados",N{r}=""),"",G{r}*N{r}))',
        16: f'=IF(D{r}="Sem dados","",IF(D{r}="Devolvida",0,IF(O{r}="","",M{r}-O{r})))',
    }


CAMPOS_VENDA = {2: "data", 3: "estado", 7: "unid", 8: "valor", 9: "envio", 10: "desc",
                11: "taxa", 12: "reemb", 13: "repasse"}


def gravar_venda(ws, r, s, pid=None):
    if pid is not None:
        ws.cell(r, 1, s["id"]).number_format = "@"
        ws.cell(r, 5, pid).number_format = "@"
        for c, f in formulas_venda(r).items():
            ws.cell(r, c, f)
    for c, campo in CAMPOS_VENDA.items():
        val = s[campo]
        ws.cell(r, c, round(val, 2) if isinstance(val, float) and c > 7 else val)
    ws.cell(r, 2).number_format = "dd/mm/yyyy hh:mm"
    ws.cell(r, 7).number_format = "0"
    for c in range(8, 17):
        ws.cell(r, c).number_format = BRL
    for c in range(1, 17):
        ws.cell(r, c).font = fonte()


def mudou(antigo, novo):
    if isinstance(novo, float) or isinstance(antigo, float):
        return (antigo is None) != (novo is None) or (novo is not None and abs(float(antigo) - novo) > 0.005)
    return antigo != novo


def gravar_produto(ws, r, pid, s):
    ws.cell(r, 1, pid).number_format = "@"
    ws.cell(r, 2, f'=IFERROR(INDEX(\'Códigos\'!$B:$B,MATCH(LEFT($A{r},2),\'Códigos\'!$A:$A,0)),"?")')
    ws.cell(r, 3, f'=IFERROR(INDEX(\'Códigos\'!$E:$E,MATCH(MID($A{r},3,2),\'Códigos\'!$D:$D,0)),"?")')
    ws.cell(r, 4, f'=IFERROR(INDEX(\'Códigos\'!$H:$H,MATCH(MID($A{r},5,9),\'Códigos\'!$G:$G,0)),"?")')
    ws.cell(r, 5, f'=B{r}&" - "&C{r}&" - "&D{r}')
    ws.cell(r, 7, s["anuncio"])
    ws.cell(r, 8, s["titulo"])
    for c in range(1, 9):
        ws.cell(r, c).font = fonte()
    f = ws.cell(r, 6)
    f.number_format = BRL
    f.fill = IN_FILL
    f.font = fonte(color="0000FF")


# ---------------------------------------------------------------------------
# Aba Resumo (refeita a cada importação; tudo por fórmula, acompanha o banco)
# ---------------------------------------------------------------------------
def montar_resumo(wb):
    if "Resumo" in wb.sheetnames:
        del wb["Resumo"]
    ws = wb.create_sheet("Resumo", 0)
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEF", [44, 16, 18, 16, 16, 3]):
        ws.column_dimensions[col].width = w
    ws["A1"] = "Resumo de Vendas"
    ws["A1"].font = fonte(bold=True, size=14, color="1F3864")
    ws["A2"] = "Considera só vendas 'Válidas' (fora devolvidas/canceladas e sem dados). Lucro = Repasse ML − CMV; precisa do custo de produção na aba Produtos."
    ws["A2"].font = fonte(italic=True, size=9, color="595959")

    V = lambda c: f"Vendas!${c}:${c}"
    ok = f'{V("D")},"Válida"'
    soma = lambda c: f"=SUMIFS({V(c)},{ok})"

    def titulo(r, textos):
        for i, t in enumerate(textos):
            c = ws.cell(r, 1 + i, t)
            c.font = fonte(bold=True, color="FFFFFF")
            c.fill = HDR_FILL
            c.alignment = Alignment(horizontal="center" if i else "left")

    titulo(4, ["Indicadores", "Total"])
    kpis = [
        ("Pedidos válidos", f'=COUNTIFS({ok})', "0"),
        ("Unidades vendidas", soma("G"), "0"),
        ("Valor das vendas", soma("H"), BRL),
        ("(−) Taxa do ML", soma("K"), BRL),
        ("(−) Envio", soma("I"), BRL),
        ("(−) Descontos", soma("J"), BRL),
        ("(−) Reembolsos", soma("L"), BRL),
        ("Repasse do ML", soma("M"), BRL),
        ("(−) CMV (custo de produção)", soma("O"), BRL),
        ("Lucro total", soma("P"), BRL),
        ("Margem de lucro", f'=IFERROR(B14/SUMIFS({V("H")},{ok},{V("O")},">=0"),0)', "0.0%"),
        ("Ticket médio", "=IFERROR(B7/B5,0)", BRL),
        ("Devolvidas / canceladas", f'=COUNTIFS({V("D")},"Devolvida")', "0"),
        ("Sem dados no relatório do ML", f'=COUNTIFS({V("D")},"Sem dados")', "0"),
        ("Válidas sem custo de produção (fora do lucro)", f'=B5-COUNTIFS({ok},{V("O")},">=0")', "0"),
    ]
    for i, (rot, f, fmt) in enumerate(kpis):
        r = 5 + i
        ws.cell(r, 1, rot).font = fonte(bold=rot.startswith("="))
        c = ws.cell(r, 2, f)
        c.number_format = fmt
        c.font = fonte(bold=rot.startswith("="))
    ws["B14"].fill = PatternFill("solid", fgColor="E2EFDA")
    ws["A14"].fill = PatternFill("solid", fgColor="E2EFDA")

    titulo(21, ["Para onde vai o dinheiro", "R$"])
    for i, (rot, ref) in enumerate([("Taxa do ML", "B8"), ("Envio", "B9"),
                                    ("Custo de produção", "B13"), ("Lucro", "B14")]):
        ws.cell(22 + i, 1, rot).font = fonte()
        c = ws.cell(22 + i, 2, f"={ref}")
        c.number_format = BRL
        c.font = fonte()

    # por mês
    datas = [ws_d.value for ws_d in wb["Vendas"]["B"][1:] if isinstance(ws_d.value, datetime)]
    meses = []
    if datas:
        y, m = min(datas).year, min(datas).month
        while (y, m) <= (max(datas).year, max(datas).month):
            meses.append((y, m))
            y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    titulo(27, ["Mês", "Pedidos", "Valor das vendas", "Lucro"])
    for i, (y, m) in enumerate(meses):
        r = 28 + i
        y2, m2 = (y + 1, 1) if m == 12 else (y, m + 1)
        per = f'{V("B")},">="&DATE({y},{m},1),{V("B")},"<"&DATE({y2},{m2},1)'
        ws.cell(r, 1, f"{m:02d}/{y}")
        ws.cell(r, 2, f"=COUNTIFS({ok},{per})").number_format = "0"
        ws.cell(r, 3, f"=SUMIFS({V('H')},{ok},{per})").number_format = BRL
        ws.cell(r, 4, f"=SUMIFS({V('P')},{ok},{per})").number_format = BRL
        for c in range(1, 5):
            ws.cell(r, c).font = fonte()
    fim_mes = 27 + len(meses)

    # por produto
    ids = [c.value for c in wb["Produtos"]["A"][1:] if c.value not in (None, "")]
    ini_p = fim_mes + 3
    titulo(ini_p, ["Produto", "ID", "Unidades", "Valor das vendas", "Lucro"])
    for i, pid in enumerate(ids):
        r = ini_p + 1 + i
        ws.cell(r, 1, f'=IFERROR(INDEX(Produtos!$E:$E,MATCH(B{r},Produtos!$A:$A,0)),B{r})')
        ws.cell(r, 2, str(pid)).number_format = "@"
        ws.cell(r, 3, f"=SUMIFS({V('G')},{ok},{V('E')},$B{r})").number_format = "0"
        ws.cell(r, 4, f"=SUMIFS({V('H')},{ok},{V('E')},$B{r})").number_format = BRL
        ws.cell(r, 5, f"=SUMIFS({V('P')},{ok},{V('E')},$B{r})").number_format = BRL
        for c in range(1, 6):
            ws.cell(r, c).font = fonte()
    fim_p = ini_p + len(ids)

    def barra(titulo_, tipo, dados, cats, ancora, alt=8.5):
        ch = BarChart()
        ch.type = tipo
        ch.title = titulo_
        ch.add_data(dados, titles_from_data=True)
        ch.set_categories(cats)
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        ch.width, ch.height = 17, alt
        ws.add_chart(ch, ancora)

    if meses:
        barra("Valor das vendas e lucro por mês", "col",
              Reference(ws, min_col=3, max_col=4, min_row=27, max_row=fim_mes),
              Reference(ws, min_col=1, min_row=28, max_row=fim_mes), "G4")
    if ids:
        barra("Lucro por produto", "bar",
              Reference(ws, min_col=5, min_row=ini_p, max_row=fim_p),
              Reference(ws, min_col=1, min_row=ini_p + 1, max_row=fim_p), "G22", alt=max(8.5, 1.0 * len(ids) + 3))
    pie = PieChart()
    pie.title = "Para onde vai o dinheiro das vendas"
    pie.add_data(Reference(ws, min_col=2, min_row=22, max_row=25))
    pie.set_categories(Reference(ws, min_col=1, min_row=22, max_row=25))
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
    pie.width, pie.height = 17, 8.5
    ws.add_chart(pie, "G" + str(22 + max(18, int(max(8.5, 1.0 * len(ids) + 3) * 2) + 1)))


# ---------------------------------------------------------------------------
# Importação
# ---------------------------------------------------------------------------
def importar(arquivos, banco=BANCO):
    log = []
    wb = load_workbook(banco) if os.path.exists(banco) else criar_banco()
    vendas, prods, cods = wb["Vendas"], wb["Produtos"], wb["Códigos"]
    linhas = {str(vendas.cell(r, 1).value): r for r in range(2, vendas.max_row + 1)
              if vendas.cell(r, 1).value not in (None, "")}
    produtos = {str(prods.cell(r, 1).value): r for r in range(2, prods.max_row + 1)
                if prods.cell(r, 1).value not in (None, "")}
    for arq in arquivos:
        nome = os.path.basename(arq)
        try:
            lista = ler_relatorio(arq)
        except Exception as e:
            log.append(f"{nome}: ERRO - {e}")
            continue
        novas = ignoradas = atualizadas = 0
        novos_prod = []
        for s in lista:
            if s["id"] in linhas:
                r = linhas[s["id"]]
                if any(mudou(vendas.cell(r, c).value, s[campo]) for c, campo in CAMPOS_VENDA.items()):
                    gravar_venda(vendas, r, s)
                    atualizadas += 1
                else:
                    ignoradas += 1
                continue
            modelo, cor, estampa = identificar(s["titulo"], s["variacao"])
            pid = codigo(cods, "modelo", modelo) + codigo(cods, "cor", cor) + codigo(cods, "estampa", estampa)
            if pid not in produtos:
                r = max([1] + list(produtos.values())) + 1
                produtos[pid] = r
                gravar_produto(prods, r, pid, s)
                novos_prod.append(f"{pid} = {modelo} - {cor} - {estampa}")
            r = max([1] + list(linhas.values())) + 1
            linhas[s["id"]] = r
            gravar_venda(vendas, r, s, pid)
            novas += 1
        msg = f"{nome}: {novas} nova(s), {ignoradas} já registrada(s) (ignorada(s))"
        if atualizadas:
            msg += f", {atualizadas} atualizada(s) (status/valores mudaram)"
        log.append(msg)
        for p in novos_prod:
            log.append(f"   produto novo: {p}  -> preencha o custo de produção na aba Produtos")
    montar_resumo(wb)
    wb.calculation.fullCalcOnLoad = True
    try:
        wb.save(banco)
    except PermissionError:
        log.append("ERRO: feche o banco_vendas.xlsx no Excel e tente de novo (nada foi salvo).")
    return log


def janela():
    import tkinter as tk
    from tkinter import filedialog, scrolledtext

    root = tk.Tk()
    root.title("Importar vendas do Mercado Livre")
    root.geometry("720x380")
    saida = scrolledtext.ScrolledText(root, font=("Consolas", 10))

    def escolher():
        arqs = filedialog.askopenfilenames(title="Planilha(s) de vendas", filetypes=[("Excel", "*.xlsx")])
        if arqs:
            saida.insert("end", "\n".join(importar(arqs)) + "\n\n")
            saida.see("end")

    tk.Button(root, text="Escolher planilha(s) de vendas e importar", command=escolher,
              font=("Arial", 11), pady=6).pack(fill="x", padx=8, pady=8)
    saida.pack(fill="both", expand=True, padx=8, pady=(0, 8))
    saida.insert("end", f"Banco: {BANCO}\n\n")
    root.mainloop()


if __name__ == "__main__":
    if len(sys.argv) > 1:
        print("\n".join(importar(sys.argv[1:])))
    else:
        janela()