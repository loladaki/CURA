from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles.differential import DifferentialStyle
import datetime

# ── COLORS ───────────────────────────────────────────────
AMBER    = "C8792A"
AMBER_D  = "9E5E1E"
AMBER_L  = "F0C080"
CREAM    = "FAF6F0"
CREAM2   = "F5EDE0"
WHITE    = "FFFFFF"
DARK     = "2C1A0E"
BORDER_C = "E8D5B8"
OK_G     = "4A9E6B"
OK_GL    = "D1FAE5"
ERR_R    = "E05252"
ERR_RL   = "FEE2E2"
WARN_YL  = "FEF3C7"
INFO_BL  = "DBEAFE"

# ── HELPERS ───────────────────────────────────────────────
def sf(c): return PatternFill("solid", fgColor=c)

def fnt(bold=False, color=DARK, size=10, italic=False):
    return Font(name="Arial", bold=bold, color=color, size=size, italic=italic)

def bdr(c=BORDER_C):
    s = Side(style="thin", color=c)
    return Border(left=s, right=s, top=s, bottom=s)

def bdr_t(c=BORDER_C):
    s = Side(style="medium", color=c)
    return Border(top=s)

def aln(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

EUR   = '€#,##0.00'
PCT1  = '0.0%'
DTFMT = 'DD/MM/YYYY'
NUM3  = '#,##0.000'

def sc(ws, r, c, val=None, bold=False, col=DARK, size=10, bg=None,
       h="left", fmt=None, wrap=False, italic=False, border=True):
    cell = ws.cell(row=r, column=c, value=val)
    cell.font = fnt(bold=bold, color=col, size=size, italic=italic)
    cell.alignment = aln(h=h, wrap=wrap)
    if bg: cell.fill = sf(bg)
    if fmt: cell.number_format = fmt
    if border: cell.border = bdr()
    return cell

def row_bg(ws, r, c_start, c_end, color):
    for c in range(c_start, c_end + 1):
        ws.cell(r, c).fill = sf(color)

def title_row(ws, r, text, ncols):
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
    c = ws.cell(r, 1, text)
    c.font = fnt(bold=True, color=WHITE, size=13)
    c.fill = sf(AMBER)
    c.alignment = aln("left", "center")
    ws.row_dimensions[r].height = 30

def col_header(ws, r, cols, bgs=None):
    for i, h in enumerate(cols, 1):
        c = ws.cell(r, i, h)
        c.font = fnt(bold=True, color=WHITE, size=9)
        c.fill = sf(AMBER_D if bgs is None else bgs[i-1])
        c.alignment = aln("center", "center")
        c.border = bdr()
    ws.row_dimensions[r].height = 22

def data_dv(ws, formula, cells):
    dv = DataValidation(type="list", formula1=formula, allow_blank=True, showErrorMessage=False)
    ws.add_data_validation(dv)
    dv.sqref = cells

# ── WORKBOOK ─────────────────────────────────────────────
wb = Workbook()

# Create sheets
sh = {}
wb.active.title = "Dashboard"
sh["Dashboard"] = wb.active
for name in ["Inventário", "Calculadora", "Receitas", "Stock", "Vendas", "Gastos", "Listas"]:
    sh[name] = wb.create_sheet(name)

# Color tabs
for name, ws in sh.items():
    ws.sheet_properties.tabColor = AMBER if name != "Listas" else "999999"

# ═══════════════════════════════════════════════════════════
# SHEET: LISTAS (hidden — dropdown sources)
# ═══════════════════════════════════════════════════════════
ws = sh["Listas"]
ws.sheet_state = "hidden"
tipos_inv   = ["Cera","Essência","Pavio","Autocolante","Recipiente","Embalagem","Cimento","Gesso","Areia","Molde","Corante","Aditivo","Outro"]
unidades    = ["kg","g","L","ml","unidade(s)","caixa(s)","rolo(s)","saco(s)"]
estados_ord = ["Pendente","A Processar","Enviada","Entregue","Cancelada"]
pagamentos  = ["MB Way","Multibanco","Transferência","Numerário","PayPal"]
categorias  = ["Matérias-Primas","Embalagem","Marketing","Transporte","Equipamento","Outro"]

def write_list(ws, col, items, label):
    ws.cell(1, col, label).font = fnt(bold=True)
    for i, v in enumerate(items, 2):
        ws.cell(i, col, v)

write_list(ws, 1, tipos_inv,   "TiposInv")
write_list(ws, 2, unidades,    "Unidades")
write_list(ws, 3, estados_ord, "EstadosOrdem")
write_list(ws, 4, pagamentos,  "Pagamentos")
write_list(ws, 5, categorias,  "Categorias")

# ═══════════════════════════════════════════════════════════
# SHEET: INVENTÁRIO
# ═══════════════════════════════════════════════════════════
ws = sh["Inventário"]
ws.freeze_panes = "A3"

title_row(ws, 1, "📦  Inventário de Matérias-Primas", 11)
cols = ["ID","Nome","Tipo","Unidade","Quantidade","Preço/Unid.","Valor Total","Stock Mín.","Fornecedor","Estado","Notas"]
col_header(ws, 2, cols)

inv_data = [
    ["INV001","Cera de Soja","Cera","kg",10,4.50,"","5","AromaShop","","Soja natural 464"],
    ["INV002","Essência Lavanda","Essência","kg",2,28.00,"","0.5","AromaShop","",""],
    ["INV003","Copos de Vidro 20cl","Recipiente","unidade(s)",200,1.20,"","50","VidroCasa","",""],
    ["INV004","Pavios Algodão CD14","Pavio","unidade(s)",500,0.10,"","100","CandleShop","","CD14 para copos 7-8cm"],
    ["INV005","Cimento Branco","Cimento","kg",5,2.80,"","2","BricoMarkt","","Fino, para moldes"],
]

row_colors = [WHITE, CREAM2]
for i, d in enumerate(inv_data):
    r = i + 3
    bg = row_colors[i % 2]
    sc(ws, r, 1, d[0], col="666666", size=9, bg=bg, h="center")   # ID
    sc(ws, r, 2, d[1], bold=True, bg=bg)                           # Nome
    sc(ws, r, 3, d[2], bg=bg, h="center")                          # Tipo
    sc(ws, r, 4, d[3], bg=bg, h="center")                          # Unidade
    sc(ws, r, 5, d[4], bg=bg, h="right", fmt=NUM3)                 # Quantidade — blue (input)
    ws.cell(r, 5).font = fnt(color="0000FF")
    sc(ws, r, 6, d[5], bg=bg, h="right", fmt=EUR)                  # Preço — blue (input)
    ws.cell(r, 6).font = fnt(color="0000FF")
    sc(ws, r, 7, f"=E{r}*F{r}", col="000000", bg=bg, h="right", fmt=EUR)  # Valor Total — formula
    sc(ws, r, 8, float(d[7]) if d[7] else 0, bg=bg, h="right", fmt=NUM3)  # Stock min — blue
    ws.cell(r, 8).font = fnt(color="0000FF")
    sc(ws, r, 9, d[8], bg=bg)                                       # Fornecedor
    # Estado formula
    sc(ws, r, 10, f'=IF(H{r}>0,IF(E{r}<=H{r},"⚠️ Baixo","✅ OK"),"—")', col="000000", bg=bg, h="center")
    sc(ws, r, 11, d[10], bg=bg, italic=True, col="666666")          # Notas

# Data validation for Tipo and Unidade
data_dv(ws, "Listas!$A$2:$A$14", f"C3:C1000")
data_dv(ws, "Listas!$B$2:$B$9",  f"D3:D1000")

# Conditional formatting: ⚠️ Baixo
red_f = sf(ERR_RL)
red_fnt = fnt(color=ERR_R, bold=True)
ws.conditional_formatting.add(
    "J3:J1000",
    CellIsRule(operator="equal", formula=['"⚠️ Baixo"'],
               fill=red_f, font=red_fnt)
)

# Totals row
tr = len(inv_data) + 3
sc(ws, tr, 1, "TOTAL", bold=True, bg=AMBER_L, h="right")
ws.merge_cells(start_row=tr, start_column=1, end_row=tr, end_column=6)
ws.cell(tr, 1).alignment = aln("right")
sc(ws, tr, 7, f"=SUM(G3:G{tr-1})", bold=True, bg=AMBER_L, h="right", fmt=EUR, col="000000")
for c in range(8, 12):
    sc(ws, tr, c, "", bg=AMBER_L)

# Column widths
widths = [8, 28, 14, 13, 12, 13, 13, 11, 18, 12, 30]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

# ═══════════════════════════════════════════════════════════
# SHEET: CALCULADORA
# ═══════════════════════════════════════════════════════════
ws = sh["Calculadora"]

title_row(ws, 1, "🧮  Calculadora de Fabricação de Velas", 6)

def section_hdr(ws, r, text, ncols=6):
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
    c = ws.cell(r, 1, text)
    c.font = fnt(bold=True, color=WHITE, size=10)
    c.fill = sf(AMBER_D)
    c.alignment = aln("left", "center")
    c.border = bdr()
    ws.row_dimensions[r].height = 20

# INPUTS
section_hdr(ws, 2, "  ⚙️  Parâmetros de Produção (alterar células a azul)")
inputs = [
    ("Nº de Velas",              100,  "",       "B3",  ""),
    ("Peso por Vela (g)",        200,  "g",      "B4",  ""),
    ("% de Essência",            0.08, "%",      "B5",  PCT1),
    ("Preço da Cera (€/kg)",     4.50, "€/kg",   "B6",  EUR),
    ("Preço da Essência (€/kg)", 25.00,"€/kg",   "B7",  EUR),
    ("Custo Recipiente (€/un)",  1.20, "€/un",   "B8",  EUR),
    ("Custo Pavio (€/un)",       0.10, "€/un",   "B9",  EUR),
    ("Custo Etiqueta (€/un)",    0.08, "€/un",   "B10", EUR),
    ("Outros Custos Fixos (€)",  0.00, "€",      "B11", EUR),
    ("Margem de Lucro",          0.40, "%",      "B12", PCT1),
]
input_labels_col = [
    "Nº de Velas","Peso por Vela (g)","% de Essência",
    "Preço da Cera (€/kg)","Preço da Essência (€/kg)",
    "Custo Recipiente (€/un)","Custo Pavio (€/un)",
    "Custo Etiqueta (€/un)","Outros Custos Fixos (€)","Margem de Lucro (%)"
]
for i, (label, val, unit, ref, fmt) in enumerate(inputs):
    r = i + 3
    sc(ws, r, 1, label, bold=True, bg=CREAM2, border=False)
    cell = ws.cell(r, 2, val)
    cell.font = fnt(color="0000FF", bold=True, size=11)
    cell.fill = sf("FFFACD")
    cell.alignment = aln("right")
    cell.border = bdr("C8792A")
    if fmt: cell.number_format = fmt
    sc(ws, r, 3, unit, col="888888", italic=True, bg=CREAM, border=False)
    ws.row_dimensions[r].height = 20

# RESULTS
section_hdr(ws, 14, "  📋  Resultados Calculados")
results = [
    ("Peso Total do Lote (g)",       "=B3*B4",                 "g",   ""),
    ("Essência Necessária (g)",       "=B15*B5",                "g",   "#,##0.0"),
    ("Cera Necessária (kg)",          "=(B15-B16)/1000",        "kg",  "#,##0.000"),
    ("Essência Necessária (kg)",      "=B16/1000",              "kg",  "#,##0.000"),
    ("— Custo Cera (€)",              "=B17*B6",                "€",   EUR),
    ("— Custo Essência (€)",          "=B18*B7",                "€",   EUR),
    ("— Custo Recipientes (€)",       "=B3*B8",                 "€",   EUR),
    ("— Custo Pavios (€)",            "=B3*B9",                 "€",   EUR),
    ("— Custo Etiquetas (€)",         "=B3*B10",                "€",   EUR),
    ("— Outros Custos Fixos (€)",     "=B11",                   "€",   EUR),
]
for i, (label, formula, unit, fmt) in enumerate(results):
    r = i + 15
    sc(ws, r, 1, label, bold=False, bg=CREAM2, border=False)
    cell = ws.cell(r, 2, formula)
    cell.font = fnt(color="000000", bold=False)
    cell.fill = sf(CREAM)
    cell.alignment = aln("right")
    cell.border = bdr()
    if fmt: cell.number_format = fmt
    sc(ws, r, 3, unit, col="888888", italic=True, bg=CREAM, border=False)
    ws.row_dimensions[r].height = 20

# TOTALS — highlighted
totals_data = [
    ("💶  CUSTO TOTAL DO LOTE (€)", "=SUM(B19:B24)", EUR,  AMBER,    WHITE,  True, 13),
    ("💰  CUSTO POR VELA (€)",      "=IF(B3>0,B25/B3,0)", EUR, AMBER_D, WHITE, True, 11),
    ("🏷️  PVP SUGERIDO (€)",        "=B26*(1+B12)", EUR,  OK_G,    WHITE, True, 11),
]
for i, (label, formula, fmt, bg, tc, bold, sz) in enumerate(totals_data):
    r = i + 25
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=1)
    sc(ws, r, 1, label, bold=True, col=tc, bg=bg, size=sz, border=False)
    cell = ws.cell(r, 2, formula)
    cell.font = fnt(bold=True, color=tc, size=sz+1)
    cell.fill = sf(bg)
    cell.alignment = aln("right")
    cell.border = bdr()
    cell.number_format = fmt
    sc(ws, r, 3, "", bg=bg, border=False)
    ws.row_dimensions[r].height = 26

# Breakdown table for chart
section_hdr(ws, 29, "  📊  Distribuição de Custos do Lote", 3)
breakdown_items = [
    ("Cera",       "=B19"),
    ("Essência",   "=B20"),
    ("Recipiente", "=B21"),
    ("Pavio",      "=B22"),
    ("Etiqueta",   "=B23"),
    ("Outros",     "=B24"),
]
sc(ws, 30, 1, "Componente", bold=True, bg=AMBER_D, col=WHITE, h="center")
sc(ws, 30, 2, "Custo (€)",  bold=True, bg=AMBER_D, col=WHITE, h="center")
for i, (label, formula) in enumerate(breakdown_items):
    r = i + 31
    bg = WHITE if i % 2 == 0 else CREAM2
    sc(ws, r, 1, label, bg=bg)
    cell = ws.cell(r, 2, formula)
    cell.font = fnt(color="000000")
    cell.fill = sf(bg)
    cell.alignment = aln("right")
    cell.border = bdr()
    cell.number_format = EUR

# Pie chart
chart = DoughnutChart()
chart.title = "Distribuição de Custos"
chart.style = 10
chart.hole_size = 40
labels = Reference(ws, min_col=1, min_row=31, max_row=36)
data   = Reference(ws, min_col=2, min_row=31, max_row=36)
chart.add_data(data)
chart.set_categories(labels)
chart.width = 14
chart.height = 12
ws.add_chart(chart, "E14")

# Lotes guardados
section_hdr(ws, 40, "  📁  Lotes Guardados", 8)
lote_cols = ["Nome do Lote","Data","Nº Velas","Cera (kg)","Essência (g)","Custo Total","Custo/Vela","PVP Sugerido"]
col_header(ws, 41, lote_cols)
ws.row_dimensions[41].height = 22
# 3 empty rows
for i in range(3):
    r = 42 + i
    for c in range(1, 9):
        sc(ws, r, c, "", bg=WHITE if i%2==0 else CREAM2)

# Column widths
for col, w in [(1,32),(2,12),(3,10),(4,12),(5,13),(6,13),(7,12),(8,13)]:
    ws.column_dimensions[get_column_letter(col)].width = w

# ═══════════════════════════════════════════════════════════
# SHEET: RECEITAS
# ═══════════════════════════════════════════════════════════
ws = sh["Receitas"]
ws.freeze_panes = "A3"

title_row(ws, 1, "🧪  Receitas de Velas — Fórmulas e Combinações", 14)
rcols = ["ID","Nome","Tipo Cera","Peso (g)","% Essência","Temp. (°C)","Tipo Pavio",
         "Recipiente","Cura (h)","Fragrâncias","Corante","Avaliação ⭐","Notas","Data"]
col_header(ws, 2, rcols)

rec_data = [
    ["REC001","Lavanda & Baunilha","Soja",200,0.08,55,"CD14","Copo Vidro 20cl",48,
     "Lavanda 6% + Baunilha 2%","Roxo — 2 gotas/100g",5,"Burn test OK. Cheiro forte e duradouro.","2025-01-20"],
    ["REC002","Cimento Nórdico","Soja + Coco 80/20",180,0.07,52,"ECO4","Molde Cimento",72,
     "Eucalipto 5% + Alecrim 2%","Sem corante",4,"Tempo de cura mais longo. Acabamento rústico.","2025-02-10"],
]
for i, d in enumerate(rec_data):
    r = i + 3
    bg = WHITE if i % 2 == 0 else CREAM2
    sc(ws, r, 1,  d[0],  col="666666", size=9, bg=bg, h="center")
    sc(ws, r, 2,  d[1],  bold=True, bg=bg)
    sc(ws, r, 3,  d[2],  bg=bg, h="center")
    sc(ws, r, 4,  d[3],  bg=bg, h="right", fmt="#,##0")
    ws.cell(r, 4).font = fnt(color="0000FF")
    sc(ws, r, 5,  d[4],  bg=bg, h="right", fmt=PCT1)
    ws.cell(r, 5).font = fnt(color="0000FF")
    sc(ws, r, 6,  d[5],  bg=bg, h="center")
    ws.cell(r, 6).font = fnt(color="0000FF")
    sc(ws, r, 7,  d[6],  bg=bg, h="center")
    sc(ws, r, 8,  d[7],  bg=bg)
    sc(ws, r, 9,  d[8],  bg=bg, h="center")
    sc(ws, r, 10, d[9],  bg=bg, wrap=True)
    sc(ws, r, 11, d[10], bg=bg, wrap=True, italic=True, col="666666")
    sc(ws, r, 12, "⭐"*d[11], bg=bg, h="center")
    sc(ws, r, 13, d[12], bg=bg, wrap=True, col="555555")
    c14 = ws.cell(r, 14, d[13])
    c14.number_format = DTFMT
    c14.fill = sf(bg)
    c14.font = fnt(color="666666")
    c14.alignment = aln("center")
    c14.border = bdr()
    ws.row_dimensions[r].height = 30

# Widths
rw = [8,26,18,9,10,9,10,18,9,30,22,13,40,13]
for i, w in enumerate(rw, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

# ═══════════════════════════════════════════════════════════
# SHEET: STOCK
# ═══════════════════════════════════════════════════════════
ws = sh["Stock"]
ws.freeze_panes = "A3"

title_row(ws, 1, "🕯️  Stock de Velas Acabadas", 10)
scols = ["ID","Nome","Essência","Tamanho","Quantidade","Preço Custo","PVP","Valor Stock","Margem %","Estado"]
col_header(ws, 2, scols)

stock_data = [
    ["STK001","Vela Lavanda 200g","Lavanda","200g",45,2.85,9.90],
    ["STK002","Vela Baunilha 200g","Baunilha","200g",12,2.85,9.90],
    ["STK003","Vela Cimento Nórdico","Eucalipto+Alecrim","180g",4,3.20,14.90],
]
for i, d in enumerate(stock_data):
    r = i + 3
    bg = WHITE if i % 2 == 0 else CREAM2
    sc(ws, r, 1, d[0], col="666666", size=9, bg=bg, h="center")
    sc(ws, r, 2, d[1], bold=True, bg=bg)
    sc(ws, r, 3, d[2], bg=bg, h="center")
    sc(ws, r, 4, d[3], bg=bg, h="center")
    sc(ws, r, 5, d[4], bg=bg, h="right")
    ws.cell(r, 5).font = fnt(color="0000FF")
    sc(ws, r, 6, d[5], bg=bg, h="right", fmt=EUR)
    ws.cell(r, 6).font = fnt(color="0000FF")
    sc(ws, r, 7, d[6], bg=bg, h="right", fmt=EUR)
    ws.cell(r, 7).font = fnt(color="0000FF")
    sc(ws, r, 8, f"=E{r}*G{r}", col="000000", bg=bg, h="right", fmt=EUR)
    sc(ws, r, 9, f'=IF(F{r}>0,(G{r}-F{r})/F{r},0)', col="000000", bg=bg, h="right", fmt=PCT1)
    sc(ws, r, 10, f'=IF(E{r}=0,"❌ Esgotado",IF(E{r}<5,"⚠️ Baixo","✅ OK"))', col="000000", bg=bg, h="center")

# Conditional formatting Estado
ws.conditional_formatting.add("J3:J1000",
    CellIsRule(operator="equal", formula=['"❌ Esgotado"'], fill=sf(ERR_RL), font=fnt(color=ERR_R, bold=True)))
ws.conditional_formatting.add("J3:J1000",
    CellIsRule(operator="equal", formula=['"⚠️ Baixo"'], fill=sf(WARN_YL), font=fnt(color="92400E", bold=True)))
ws.conditional_formatting.add("J3:J1000",
    CellIsRule(operator="equal", formula=['"✅ OK"'], fill=sf(OK_GL), font=fnt(color=OK_G, bold=True)))

# Totals
nr = len(stock_data) + 3
sc(ws, nr, 1, "TOTAL", bold=True, bg=AMBER_L, h="right")
ws.merge_cells(start_row=nr, start_column=1, end_row=nr, end_column=4)
sc(ws, nr, 5, f"=SUM(E3:E{nr-1})", bold=True, bg=AMBER_L, h="right", col="000000")
sc(ws, nr, 6, "", bg=AMBER_L)
sc(ws, nr, 7, "", bg=AMBER_L)
sc(ws, nr, 8, f"=SUM(H3:H{nr-1})", bold=True, bg=AMBER_L, h="right", fmt=EUR, col="000000")
sc(ws, nr, 9, f"=IFERROR(AVERAGE(I3:I{nr-1}),0)", bold=True, bg=AMBER_L, h="right", fmt=PCT1, col="000000")
sc(ws, nr, 10, "", bg=AMBER_L)

sw = [8,26,18,11,11,13,13,14,11,13]
for i, w in enumerate(sw, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

# ═══════════════════════════════════════════════════════════
# SHEET: VENDAS
# ═══════════════════════════════════════════════════════════
ws = sh["Vendas"]
ws.freeze_panes = "A3"

title_row(ws, 1, "🛒  Vendas & Encomendas", 10)
vcols = ["ID","Data","Cliente","Produto","Quantidade","Preço/Un.","Total","Pagamento","Estado","Notas"]
col_header(ws, 2, vcols)

vendas_data = [
    ["VND001","2025-01-15","Maria Santos","Vela Lavanda 200g",3,9.90,"Entregue","MB Way",""],
    ["VND002","2025-02-03","João Ferreira","Vela Baunilha 200g",2,9.90,"Entregue","Transferência",""],
    ["VND003","2025-03-10","Ana Costa","Pack 3 Velas Sortidas",1,24.90,"Enviada","MB Way","Envio CTT"],
    ["VND004","2025-04-22","Carlos Lopes","Vela Cimento Nórdico",2,14.90,"Pendente","","Aguarda confirmação"],
    ["VND005","2025-05-01","Rita Oliveira","Vela Lavanda 200g",5,9.90,"A Processar","Numerário",""],
]
for i, d in enumerate(vendas_data):
    r = i + 3
    bg = WHITE if i % 2 == 0 else CREAM2
    sc(ws, r, 1, d[0], col="666666", size=9, bg=bg, h="center")
    dc = ws.cell(r, 2, d[1])
    dc.number_format = DTFMT; dc.fill = sf(bg); dc.font = fnt(); dc.alignment = aln("center"); dc.border = bdr()
    sc(ws, r, 3, d[2], bold=True, bg=bg)
    sc(ws, r, 4, d[3], bg=bg)
    sc(ws, r, 5, d[4], bg=bg, h="right")
    ws.cell(r, 5).font = fnt(color="0000FF")
    sc(ws, r, 6, d[5], bg=bg, h="right", fmt=EUR)
    ws.cell(r, 6).font = fnt(color="0000FF")
    sc(ws, r, 7, f"=E{r}*F{r}", col="000000", bg=bg, h="right", fmt=EUR)
    sc(ws, r, 8, d[7], bg=bg, h="center")
    sc(ws, r, 9, d[6], bg=bg, h="center")
    sc(ws, r, 10, d[8], bg=bg, italic=True, col="666666")

# Estado dropdown and conditional formatting
data_dv(ws, "Listas!$C$2:$C$6", "I3:I1000")
data_dv(ws, "Listas!$D$2:$D$6", "H3:H1000")

ws.conditional_formatting.add("I3:I1000",
    CellIsRule(operator="equal", formula=['"Pendente"'], fill=sf(WARN_YL), font=fnt(color="92400E")))
ws.conditional_formatting.add("I3:I1000",
    CellIsRule(operator="equal", formula=['"A Processar"'], fill=sf(INFO_BL), font=fnt(color="1e40af")))
ws.conditional_formatting.add("I3:I1000",
    CellIsRule(operator="equal", formula=['"Entregue"'], fill=sf(OK_GL), font=fnt(color=OK_G)))
ws.conditional_formatting.add("I3:I1000",
    CellIsRule(operator="equal", formula=['"Cancelada"'], fill=sf(ERR_RL), font=fnt(color=ERR_R)))

# Totals
nt = len(vendas_data) + 3
sc(ws, nt, 1, "TOTAL", bold=True, bg=AMBER_L, h="right")
ws.merge_cells(start_row=nt, start_column=1, end_row=nt, end_column=6)
sc(ws, nt, 7, f"=SUM(G3:G{nt-1})", bold=True, bg=AMBER_L, h="right", fmt=EUR, col="000000")
for c in range(8, 11):
    sc(ws, nt, c, "", bg=AMBER_L)

vw = [8,14,20,28,11,12,13,14,14,28]
for i, w in enumerate(vw, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

# ═══════════════════════════════════════════════════════════
# SHEET: GASTOS
# ═══════════════════════════════════════════════════════════
ws = sh["Gastos"]
ws.freeze_panes = "A3"

title_row(ws, 1, "💸  Gastos & Despesas", 7)
gcols = ["ID","Data","Descrição","Categoria","Valor (€)","Fornecedor","Notas"]
col_header(ws, 2, gcols)

gastos_data = [
    ["GST001","2025-01-10","Cera de Soja 20kg","Matérias-Primas",90.00,"AromaShop",""],
    ["GST002","2025-01-10","Essências variadas 2kg","Matérias-Primas",56.00,"AromaShop",""],
    ["GST003","2025-02-05","Copos de Vidro x200","Embalagem",240.00,"VidroCasa",""],
    ["GST004","2025-02-20","Fotografia de produtos","Marketing",80.00,"Studio Foto","Sessão produto"],
    ["GST005","2025-03-15","Envios CTT Março","Transporte",45.50,"CTT",""],
    ["GST006","2025-04-01","Pavios x1000","Matérias-Primas",85.00,"CandleShop",""],
]
for i, d in enumerate(gastos_data):
    r = i + 3
    bg = WHITE if i % 2 == 0 else CREAM2
    sc(ws, r, 1, d[0], col="666666", size=9, bg=bg, h="center")
    dc = ws.cell(r, 2, d[1])
    dc.number_format = DTFMT; dc.fill = sf(bg); dc.font = fnt(); dc.alignment = aln("center"); dc.border = bdr()
    sc(ws, r, 3, d[2], bold=True, bg=bg)
    sc(ws, r, 4, d[3], bg=bg, h="center")
    sc(ws, r, 5, d[4], bg=bg, h="right", fmt=EUR)
    ws.cell(r, 5).font = fnt(color="0000FF")
    sc(ws, r, 6, d[5], bg=bg)
    sc(ws, r, 7, d[6], bg=bg, italic=True, col="666666")

data_dv(ws, "Listas!$E$2:$E$7", "D3:D1000")

cat_colors = {
    "Matérias-Primas": ("DBEAFE", "1e40af"),
    "Embalagem":       ("F3F4F6", "374151"),
    "Marketing":       ("FEE2E2", ERR_R),
    "Transporte":      (WARN_YL, "92400E"),
    "Equipamento":     (OK_GL, OK_G),
    "Outro":           ("F5F3FF", "5b21b6"),
}
for cat, (bg_hex, fg_hex) in cat_colors.items():
    ws.conditional_formatting.add("D3:D1000",
        CellIsRule(operator="equal", formula=[f'"{cat}"'],
                   fill=sf(bg_hex), font=fnt(color=fg_hex, bold=True)))

ng = len(gastos_data) + 3
sc(ws, ng, 1, "TOTAL", bold=True, bg=AMBER_L, h="right")
ws.merge_cells(start_row=ng, start_column=1, end_row=ng, end_column=4)
sc(ws, ng, 5, f"=SUM(E3:E{ng-1})", bold=True, bg=AMBER_L, h="right", fmt=EUR, col="000000")
sc(ws, ng, 6, "", bg=AMBER_L); sc(ws, ng, 7, "", bg=AMBER_L)

# Summary by category (cols I–K)
sc(ws, 2, 9, "Categoria",    bold=True, bg=AMBER_D, col=WHITE, h="center")
sc(ws, 2, 10,"Total (€)",    bold=True, bg=AMBER_D, col=WHITE, h="center")
sc(ws, 2, 11,"% do Total",   bold=True, bg=AMBER_D, col=WHITE, h="center")
cats = list(cat_colors.keys())
for i, cat in enumerate(cats):
    r = i + 3
    sc(ws, r, 9,  cat, bg=WHITE)
    sc(ws, r, 10, f'=SUMIF(D$3:D$1000,I{r},E$3:E$1000)', col="008000", bg=WHITE, h="right", fmt=EUR)
    sc(ws, r, 11, f'=IFERROR(J{r}/J{ng+1},0)', col="008000", bg=WHITE, h="right", fmt=PCT1)

sc(ws, ng+1, 9,  "TOTAL", bold=True, bg=AMBER_L)
sc(ws, ng+1, 10, f"=SUM(J3:J{ng})", bold=True, bg=AMBER_L, h="right", fmt=EUR, col="000000")
sc(ws, ng+1, 11, "100%",  bold=True, bg=AMBER_L, h="right", fmt=PCT1)

gw = [8,14,30,18,13,18,28]
for i, w in enumerate(gw, 1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.column_dimensions["I"].width = 18
ws.column_dimensions["J"].width = 14
ws.column_dimensions["K"].width = 12

# ═══════════════════════════════════════════════════════════
# SHEET: DASHBOARD
# ═══════════════════════════════════════════════════════════
ws = sh["Dashboard"]

# Big title
ws.merge_cells("A1:J1")
ws.row_dimensions[1].height = 46
c = ws.cell(1, 1, "🕯️  VELAS — Dashboard de Gestão")
c.font = Font(name="Arial", bold=True, color=WHITE, size=18)
c.fill = sf(AMBER)
c.alignment = Alignment(horizontal="center", vertical="center")

# Date subtitle
ws.merge_cells("A2:J2")
ws.row_dimensions[2].height = 18
c2 = ws.cell(2, 1, f"Atualizado em {datetime.date.today().strftime('%d/%m/%Y')}")
c2.font = fnt(italic=True, color="888888", size=9)
c2.alignment = aln("right")

# ── KPI CARDS (row 4–8) ──
ws.row_dimensions[3].height = 10
ws.row_dimensions[4].height = 22
ws.row_dimensions[5].height = 32
ws.row_dimensions[6].height = 18
ws.row_dimensions[7].height = 10

def kpi_card(ws, col_start, label, formula, fmt, icon, bg_label, bg_val, col_val):
    cs = col_start
    # merge across 2 columns
    ws.merge_cells(start_row=4, start_column=cs, end_row=4, end_column=cs+1)
    ws.merge_cells(start_row=5, start_column=cs, end_row=5, end_column=cs+1)
    ws.merge_cells(start_row=6, start_column=cs, end_row=6, end_column=cs+1)
    lbl = ws.cell(4, cs, f"{icon}  {label}")
    lbl.font = fnt(bold=True, color=WHITE, size=9)
    lbl.fill = sf(bg_label); lbl.alignment = aln("center", "center"); lbl.border = bdr()
    val = ws.cell(5, cs, formula)
    val.font = Font(name="Arial", bold=True, color=col_val, size=18)
    val.fill = sf(bg_val); val.alignment = aln("center", "center"); val.number_format = fmt; val.border = bdr()
    sep = ws.cell(6, cs, "")
    sep.fill = sf(bg_label); sep.border = bdr()

kpi_cards = [
    (1,  "Vendas Totais",      '=SUMIFS(Vendas!G:G,Vendas!I:I,"<>Cancelada")',       EUR,  "📈", AMBER_D, CREAM,  AMBER_D),
    (3,  "Este Mês",           f'=SUMPRODUCT((MONTH(Vendas!B3:B1000)=MONTH(TODAY()))*(YEAR(Vendas!B3:B1000)=YEAR(TODAY()))*(Vendas!I3:I1000<>"Cancelada")*Vendas!G3:G1000)', EUR, "📅", AMBER, CREAM2, AMBER),
    (5,  "Lucro Estimado",     '=SUMIFS(Vendas!G:G,Vendas!I:I,"<>Cancelada")-SUM(Gastos!E:E)', EUR, "💰", OK_G,   OK_GL,  OK_G),
    (7,  "Pendentes",          '=COUNTIF(Vendas!I:I,"Pendente")',                     "#,##0","⏳", "E8A030", WARN_YL,"92400E"),
    (9,  "Valor Inventário",   "=SUMPRODUCT('Inventário'!E3:E1000,'Inventário'!F3:F1000)", EUR, "📦", AMBER_D, CREAM, AMBER_D),
    (11, "Stock Velas",        "=SUM(Stock!E:E)",                                    "#,##0 \"un\"", "🕯️", AMBER, CREAM2, AMBER),
]
for cs, label, formula, fmt, icon, bg_lbl, bg_val, col_val in kpi_cards:
    kpi_card(ws, cs, label, formula, fmt, icon, bg_lbl, bg_val, col_val)

# ── MONTHLY SALES TABLE ──
ws.row_dimensions[9].height = 10
sc(ws, 10, 1, "📅  Vendas Mensais 2025", bold=True, bg=AMBER_D, col=WHITE, size=11, h="left", border=False)
ws.merge_cells("A10:D10")

sc(ws, 11, 1, "Mês",          bold=True, bg=AMBER, col=WHITE, h="center")
sc(ws, 11, 2, "Vendas (€)",   bold=True, bg=AMBER, col=WHITE, h="center")
sc(ws, 11, 3, "Despesas (€)", bold=True, bg=AMBER, col=WHITE, h="center")
sc(ws, 11, 4, "Lucro (€)",    bold=True, bg=AMBER, col=WHITE, h="center")

months_pt = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]
for m in range(1, 13):
    r = m + 11
    bg = WHITE if m % 2 == 0 else CREAM2
    sc(ws, r, 1, months_pt[m-1]+" 2025", bold=True, bg=bg)
    # Sales for month m: SUMPRODUCT where month=m, year=2025, status<>Cancelada
    sales_f = f"=SUMPRODUCT((MONTH(Vendas!$B$3:$B$1000)={m})*(YEAR(Vendas!$B$3:$B$1000)=2025)*(Vendas!$I$3:$I$1000<>\"Cancelada\")*Vendas!$G$3:$G$1000)"
    exp_f   = f"=SUMPRODUCT((MONTH(Gastos!$B$3:$B$1000)={m})*(YEAR(Gastos!$B$3:$B$1000)=2025)*Gastos!$E$3:$E$1000)"
    sc(ws, r, 2, sales_f, col="008000", bg=bg, h="right", fmt=EUR)
    sc(ws, r, 3, exp_f,   col="008000", bg=bg, h="right", fmt=EUR)
    sc(ws, r, 4, f"=B{r}-C{r}", col="000000", bg=bg, h="right", fmt=EUR)
    ws.row_dimensions[r].height = 18

# Total row for monthly table
tr2 = 24
sc(ws, tr2, 1, "TOTAL", bold=True, bg=AMBER_L, h="right")
sc(ws, tr2, 2, "=SUM(B12:B23)", bold=True, bg=AMBER_L, h="right", fmt=EUR, col="000000")
sc(ws, tr2, 3, "=SUM(C12:C23)", bold=True, bg=AMBER_L, h="right", fmt=EUR, col="000000")
sc(ws, tr2, 4, "=SUM(D12:D23)", bold=True, bg=AMBER_L, h="right", fmt=EUR, col="000000")
ws.row_dimensions[tr2].height = 22

# ── EXPENSES BY CATEGORY TABLE ──
sc(ws, 10, 6, "💸  Despesas por Categoria", bold=True, bg=AMBER_D, col=WHITE, size=11, border=False)
ws.merge_cells("F10:H10")

sc(ws, 11, 6, "Categoria",  bold=True, bg=AMBER, col=WHITE, h="center")
sc(ws, 11, 7, "Total (€)",  bold=True, bg=AMBER, col=WHITE, h="center")
sc(ws, 11, 8, "% do Total", bold=True, bg=AMBER, col=WHITE, h="center")
total_exp_ref = "SUM(Gastos!$E$3:$E$1000)"
for i, cat in enumerate(cats):
    r = i + 12
    bg = WHITE if i % 2 == 0 else CREAM2
    sc(ws, r, 6, cat, bg=bg)
    sc(ws, r, 7, f'=SUMIF(Gastos!$D$3:$D$1000,F{r},Gastos!$E$3:$E$1000)', col="008000", bg=bg, h="right", fmt=EUR)
    sc(ws, r, 8, f'=IFERROR(G{r}/{total_exp_ref},0)', col="008000", bg=bg, h="right", fmt=PCT1)
    ws.row_dimensions[r].height = 18

trc = 18
sc(ws, trc, 6, "TOTAL", bold=True, bg=AMBER_L)
sc(ws, trc, 7, f"=SUM(G12:G{trc-1})", bold=True, bg=AMBER_L, h="right", fmt=EUR, col="000000")
sc(ws, trc, 8, "100%", bold=True, bg=AMBER_L, h="right", fmt=PCT1)
ws.row_dimensions[trc].height = 22

# ── CHART: Monthly Sales Bar ──
bar = BarChart()
bar.type = "col"
bar.title = "Vendas vs Despesas Mensais 2025"
bar.style = 10
bar.grouping = "clustered"
bar.y_axis.title = "€"
bar.x_axis.title = "Mês"

cats_ref = Reference(ws, min_col=1, min_row=12, max_row=23)
sales_ref = Reference(ws, min_col=2, min_row=11, max_row=23)
exp_ref   = Reference(ws, min_col=3, min_row=11, max_row=23)
bar.add_data(sales_ref, titles_from_data=True)
bar.add_data(exp_ref,   titles_from_data=True)
bar.set_categories(cats_ref)
bar.series[0].graphicalProperties.solidFill = AMBER
bar.series[1].graphicalProperties.solidFill = ERR_R
bar.width = 20; bar.height = 14
ws.add_chart(bar, "A26")

# ── CHART: Expenses Doughnut ──
donut = DoughnutChart()
donut.title = "Despesas por Categoria"
donut.style = 10
donut.hole_size = 40
d_labels = Reference(ws, min_col=6, min_row=12, max_row=17)
d_data   = Reference(ws, min_col=7, min_row=11, max_row=17)
donut.add_data(d_data, titles_from_data=True)
donut.set_categories(d_labels)
donut.width = 18; donut.height = 14
ws.add_chart(donut, "F20")

# Dashboard column widths
dw = [14,14,14,14,2,18,14,11,2,2,2,2]
for i, w in enumerate(dw, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

# ── SAVE ─────────────────────────────────────────────────
out = r"C:\Users\PC\Desktop\velas-dashboard\velas.xlsx"
wb.save(out)
print(f"Saved: {out}")
