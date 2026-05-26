"""
Excel Avançado + Power Query — Modelo Operacional
==================================================
Autor: Emanoel Cavalcante

Este script Python replica a lógica de transformação
que era feita manualmente via Power Query no Excel,
demonstrando domínio do processo de ETL e das fórmulas
avançadas utilizadas na operação CDDNPA e na Santos & Abreu.

Fórmulas Excel equivalentes documentadas nos comentários.
Saída: arquivo .xlsx com múltiplas abas formatadas.
"""

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import (PatternFill, Font, Alignment,
                              Border, Side, GradientFill)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import SeriesLabel
import warnings
warnings.filterwarnings('ignore')

np.random.seed(99)

print("=" * 60)
print("  EXCEL AVANÇADO + POWER QUERY — MODELO OPERACIONAL")
print("  Emanoel Cavalcante · emanoelinc.github.io")
print("=" * 60)

# ── Cores ──
TEAL_DARK  = '085041'
TEAL       = '1D9E75'
TEAL_LIGHT = 'E1F5EE'
GRAY_BG    = 'F8F7F3'
BORDER_C   = 'E0DDD8'
WHITE      = 'FFFFFF'
INK        = '1A1A18'
RED        = 'E8384D'
YELLOW     = 'F2C811'

def border(style='thin'):
    s = Side(style=style, color=BORDER_C)
    return Border(left=s, right=s, top=s, bottom=s)

def header_cell(ws, row, col, value, bg=TEAL_DARK, fg=WHITE, bold=True, size=10):
    c = ws.cell(row=row, column=col, value=value)
    c.fill = PatternFill('solid', fgColor=bg)
    c.font = Font(bold=bold, color=fg, size=size)
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    c.border = border()
    return c

def data_cell(ws, row, col, value, bg=WHITE, bold=False, fmt=None, align='center'):
    c = ws.cell(row=row, column=col, value=value)
    c.fill = PatternFill('solid', fgColor=bg)
    c.font = Font(bold=bold, color=INK, size=9)
    c.alignment = Alignment(horizontal=align, vertical='center')
    c.border = border()
    if fmt: c.number_format = fmt
    return c

# ──────────────────────────────────────────────
# DADOS
# ──────────────────────────────────────────────

meses = ['Jan/25','Fev/25','Mar/25','Abr/25','Mai/25']
categorias = ['Medicamentos Genéricos','Dermocosméticos','Higiene Pessoal',
              'Perfumaria','Med. de Marca','Suplementos','Fraldas','OTC / Vitaminas']

# Vendas por categoria e mês
np.random.seed(5)
vendas_mat = np.random.randint(15000, 80000, (len(categorias), len(meses))).astype(float)
cmv_pct    = np.random.uniform(0.40, 0.68, len(categorias))
ruptura    = np.random.uniform(2, 14, (len(categorias), len(meses)))

wb = Workbook()

# ──────────────────────────────────────────────
# ABA 1: RESUMO EXECUTIVO
# ──────────────────────────────────────────────

ws1 = wb.active; ws1.title = 'Resumo Executivo'
ws1.sheet_view.showGridLines = False
ws1.column_dimensions['A'].width = 28
for col in 'BCDEFG': ws1.column_dimensions[col].width = 16

# Título
ws1.merge_cells('A1:G1')
c = ws1['A1']; c.value = 'RESUMO EXECUTIVO — VAREJO FARMACÊUTICO'
c.fill = PatternFill('solid', fgColor=TEAL_DARK)
c.font = Font(bold=True, color=WHITE, size=13)
c.alignment = Alignment(horizontal='center', vertical='center')
ws1.row_dimensions[1].height = 32

ws1.merge_cells('A2:G2')
c = ws1['A2']; c.value = 'Santos & Abreu — Gestão de Drogarias  |  Jan–Mai 2025  |  Emanoel Cavalcante'
c.fill = PatternFill('solid', fgColor=TEAL)
c.font = Font(color=WHITE, size=9)
c.alignment = Alignment(horizontal='center', vertical='center')
ws1.row_dimensions[2].height = 18

# Headers
ws1.row_dimensions[4].height = 24
headers = ['Categoria', 'Jan/25', 'Fev/25', 'Mar/25', 'Abr/25', 'Mai/25', 'TOTAL']
for i, h in enumerate(headers, 1):
    header_cell(ws1, 4, i, h)

# Dados
total_geral = 0
for r, cat in enumerate(categorias, 5):
    bg = GRAY_BG if r % 2 == 0 else WHITE
    data_cell(ws1, r, 1, cat, bg=bg, align='left')
    total_linha = 0
    for c_idx, val in enumerate(vendas_mat[r-5], 2):
        data_cell(ws1, r, c_idx, val, bg=bg, fmt='R$ #,##0')
        total_linha += val
    data_cell(ws1, r, 7, total_linha, bg=TEAL_LIGHT, bold=True, fmt='R$ #,##0')
    total_geral += total_linha
    ws1.row_dimensions[r].height = 18

# Total geral
row_tot = len(categorias) + 5
ws1.row_dimensions[row_tot].height = 22
data_cell(ws1, row_tot, 1, 'TOTAL GERAL', bg=TEAL_DARK, bold=True, align='left').font = Font(bold=True, color=WHITE, size=10)
for c_idx in range(2, 7):
    data_cell(ws1, row_tot, c_idx, sum(vendas_mat[:, c_idx-2]),
              bg=TEAL_DARK, bold=True, fmt='R$ #,##0').font = Font(bold=True, color=WHITE, size=10)
data_cell(ws1, row_tot, 7, total_geral, bg=TEAL, bold=True, fmt='R$ #,##0').font = Font(bold=True, color=WHITE, size=11)

# ──────────────────────────────────────────────
# ABA 2: CMV & MARGEM
# ──────────────────────────────────────────────

ws2 = wb.create_sheet('CMV & Margem')
ws2.sheet_view.showGridLines = False
ws2.column_dimensions['A'].width = 28
for col in ['B','C','D','E','F']: ws2.column_dimensions[col].width = 16

ws2.merge_cells('A1:F1')
c = ws2['A1']; c.value = 'CMV & MARGEM BRUTA POR CATEGORIA'
c.fill = PatternFill('solid', fgColor=TEAL_DARK)
c.font = Font(bold=True, color=WHITE, size=13)
c.alignment = Alignment(horizontal='center', vertical='center')
ws2.row_dimensions[1].height = 32

headers2 = ['Categoria', 'Faturamento Total', 'CMV %', 'CMV R$', 'Margem R$', 'Margem %']
for i, h in enumerate(headers2, 1):
    header_cell(ws2, 3, i, h)
ws2.row_dimensions[3].height = 24

for r, (cat, pct) in enumerate(zip(categorias, cmv_pct), 4):
    fat_total = sum(vendas_mat[r-4])
    cmv_r     = fat_total * pct
    marg_r    = fat_total - cmv_r
    marg_pct  = marg_r / fat_total * 100
    bg = GRAY_BG if r % 2 == 0 else WHITE

    data_cell(ws2, r, 1, cat, bg=bg, align='left')
    data_cell(ws2, r, 2, fat_total, bg=bg, fmt='R$ #,##0')
    data_cell(ws2, r, 3, pct, bg=bg, fmt='0.0%')

    # Semáforo CMV
    cmv_bg = 'FFE5E5' if pct > 0.60 else 'FFF9E5' if pct > 0.50 else 'E5F9F0'
    c = ws2.cell(row=r, column=4, value=cmv_r)
    c.fill = PatternFill('solid', fgColor=cmv_bg)
    c.font = Font(size=9); c.alignment = Alignment(horizontal='center')
    c.number_format = 'R$ #,##0'; c.border = border()

    data_cell(ws2, r, 5, marg_r, bg=bg, fmt='R$ #,##0')

    # Semáforo margem
    mg_bg = 'E5F9F0' if marg_pct > 45 else 'FFF9E5' if marg_pct > 35 else 'FFE5E5'
    c = ws2.cell(row=r, column=6, value=marg_pct/100)
    c.fill = PatternFill('solid', fgColor=mg_bg)
    c.font = Font(bold=True, size=9)
    c.alignment = Alignment(horizontal='center')
    c.number_format = '0.0%'; c.border = border()
    ws2.row_dimensions[r].height = 18

# ──────────────────────────────────────────────
# ABA 3: RUPTURA
# ──────────────────────────────────────────────

ws3 = wb.create_sheet('Ruptura de Estoque')
ws3.sheet_view.showGridLines = False
ws3.column_dimensions['A'].width = 28
for col in 'BCDEFG': ws3.column_dimensions[col].width = 14

ws3.merge_cells('A1:G1')
c = ws3['A1']; c.value = 'RUPTURA DE ESTOQUE % POR CATEGORIA — Meta ≤ 5%'
c.fill = PatternFill('solid', fgColor=TEAL_DARK)
c.font = Font(bold=True, color=WHITE, size=13)
c.alignment = Alignment(horizontal='center', vertical='center')
ws3.row_dimensions[1].height = 32

for i, h in enumerate(['Categoria'] + meses + ['Média'], 1):
    header_cell(ws3, 3, i, h)
ws3.row_dimensions[3].height = 24

for r, cat in enumerate(categorias, 4):
    bg = GRAY_BG if r % 2 == 0 else WHITE
    data_cell(ws3, r, 1, cat, bg=bg, align='left')
    media = 0
    for c_idx, val in enumerate(ruptura[r-4], 2):
        # Semáforo ruptura
        if val > 10:   rb = 'FFE5E5'
        elif val > 5:  rb = 'FFF9E5'
        else:          rb = 'E5F9F0'
        cell = ws3.cell(row=r, column=c_idx, value=val/100)
        cell.fill = PatternFill('solid', fgColor=rb)
        cell.font = Font(size=9)
        cell.alignment = Alignment(horizontal='center')
        cell.number_format = '0.0%'; cell.border = border()
        media += val
    media /= len(meses)
    mb = 'FFE5E5' if media > 10 else 'FFF9E5' if media > 5 else 'E5F9F0'
    c2 = ws3.cell(row=r, column=7, value=media/100)
    c2.fill = PatternFill('solid', fgColor=mb)
    c2.font = Font(bold=True, size=9)
    c2.alignment = Alignment(horizontal='center')
    c2.number_format = '0.0%'; c2.border = border()
    ws3.row_dimensions[r].height = 18

# ──────────────────────────────────────────────
# ABA 4: DOCUMENTAÇÃO POWER QUERY
# ──────────────────────────────────────────────

ws4 = wb.create_sheet('Power Query — Lógica')
ws4.sheet_view.showGridLines = False
ws4.column_dimensions['A'].width = 35
ws4.column_dimensions['B'].width = 65

ws4.merge_cells('A1:B1')
c = ws4['A1']; c.value = 'DOCUMENTAÇÃO — LÓGICA POWER QUERY EQUIVALENTE'
c.fill = PatternFill('solid', fgColor=TEAL_DARK)
c.font = Font(bold=True, color=WHITE, size=12)
c.alignment = Alignment(horizontal='center', vertical='center')
ws4.row_dimensions[1].height = 28

header_cell(ws4, 2, 1, 'Etapa Power Query', bg=TEAL)
header_cell(ws4, 2, 2, 'Código M / Descrição', bg=TEAL)
ws4.row_dimensions[2].height = 20

etapas = [
    ('1. Fonte de dados', '= Excel.CurrentWorkbook(){[Name="TabelaVendas"]}[Content]'),
    ('2. Tipo das colunas', '= Table.TransformColumnTypes(Fonte, {{"Data", type date}, {"Valor", type number}})'),
    ('3. Filtrar período', '= Table.SelectRows(Tipado, each [Data] >= #date(2025,1,1))'),
    ('4. Coluna Mês-Ano', '= Table.AddColumn(Filtrado, "MesAno", each Date.ToText([Data], "MMM/yy"))'),
    ('5. CMV calculado', '= Table.AddColumn(t, "CMV", each [Faturamento] * [CMV_Pct])'),
    ('6. Margem bruta', '= Table.AddColumn(t, "Margem", each [Faturamento] - [CMV])'),
    ('7. Margem %', '= Table.AddColumn(t, "Margem_Pct", each [Margem] / [Faturamento])'),
    ('8. Curva ABC', '= Table.AddColumn(t, "Curva", each if [Pct_Acum] <= 0.70 then "A" else if [Pct_Acum] <= 0.90 then "B" else if [Pct_Acum] <= 0.99 then "C" else "D")'),
    ('9. Agrupar por mês', '= Table.Group(t, {"MesAno"}, {{"Total", each List.Sum([Faturamento]), type number}})'),
    ('10. Variação MoM', '= Table.AddColumn(t, "Var_MoM", each ([Faturamento] - [Fat_Anterior]) / [Fat_Anterior])'),
    ('11. Ruptura flag', '= Table.AddColumn(t, "Alerta", each if [Ruptura_Pct] > 0.05 then "⚠ Acima" else "✓ OK")'),
    ('12. Carregar no Excel', '= Table.SelectColumns(Final, {"MesAno","Categoria","Faturamento","CMV","Margem","Curva"})'),
]

for r, (etapa, cod) in enumerate(etapas, 3):
    bg = GRAY_BG if r % 2 == 0 else WHITE
    c1 = ws4.cell(row=r, column=1, value=etapa)
    c1.fill = PatternFill('solid', fgColor=bg)
    c1.font = Font(bold=True, size=9, color=TEAL_DARK)
    c1.alignment = Alignment(vertical='center', wrap_text=True)
    c1.border = border()
    c2 = ws4.cell(row=r, column=2, value=cod)
    c2.fill = PatternFill('solid', fgColor=bg)
    c2.font = Font(size=9, color=INK, name='Courier New')
    c2.alignment = Alignment(vertical='center', wrap_text=True)
    c2.border = border()
    ws4.row_dimensions[r].height = 22

wb.save('modelo_excel_avancado.xlsx')
print("\n✅ modelo_excel_avancado.xlsx gerado com sucesso!")
print("   Abas criadas: Resumo Executivo | CMV & Margem | Ruptura | Power Query — Lógica")
