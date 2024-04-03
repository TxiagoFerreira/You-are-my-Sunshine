from openpyxl import Workbook,load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Border, Side, PatternFill, GradientFill, Alignment

# Criar um novo workbook
wb = load_workbook(r"C:\Users\Alunos\Downloads\yha.xlsx")
WS = wb.active

WS['A1'] = "=DATE(2024,05,01)"

#mês
WS.merge_cells('D2:AG2')
d2 = WS['D2']
d2 = '=TEXT(A1,"mmmm")'
d2.alignment = Alignment(horizontal="center", vertical="center")
d2.font = Font(name="Calibri", size=48, b="true")
#d2.border = Border(top="double")

#nº de dias da semana
WS['C4'] = "=DAY(A1)"
WS['D4'] = "=C4+1"
WS['E4'] = "=D4+1"
WS['F4'] = "=E4+1"
WS['G4'] = "=F4+1"
WS['H4'] = "=G4+1"
WS['I4'] = "=H4+1"
WS['J4'] = "=I4+1"
WS['K4'] = "=J4+1"
WS['L4'] = "=K4+1"
WS['M4'] = "=L4+1"
WS['N4'] = "=M4+1"
WS['O4'] = "=N4+1"
WS['P4'] = "=O4+1"
WS['Q4'] = "=P4+1"
WS['R4'] = "=Q4+1"
WS['S4'] = "=R4+1"
WS['T4'] = "=S4+1"
WS['U4'] = "=T4+1"
WS['V4'] = "=U4+1"
WS['W4'] = "=V4+1"
WS['X4'] = "=W4+1"
WS['Y4'] = "=X4+1"
WS['Z4'] = "=Y4+1"
WS['AA4'] = "=Z4+1"
WS['AB4'] = "=AA4+1"
WS['AC4'] = "=AB4+1"    
WS['AD4'] = "=AC4+1"
WS['AE4'] = "=AD4+1"
WS['AF4'] = "=AE4+1"

#dias da semana
WS['C3'] = '=TEXT(A1, "ddd")'
WS['D3'] = '=TEXT(A1 + 1, "ddd")'
WS['E3'] = '=TEXT(D1 + 1, "ddd")'
WS['E3'] = '=TEXT(D1 + 1, "ddd")'


wb.save(r"C:\Users\Alunos\Downloads\yha.xlsx")