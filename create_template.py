import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Create workbook and sheet
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Computo Metrico"

# Define headers
headers = [
    "Categoria", "Codice", "Descrizione Sintetica", "Descrizione Estesa", 
    "U.M.", "N.", "Lung.", "Larg.", "Alt./Peso", 
    "Quantità", "Prezzo Unit.", "Prezzo Totale"
]

# Write headers
for col_num, header in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col_num)
    cell.value = header
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    cell.alignment = Alignment(horizontal="center", vertical="center")

# Add some dummy rows to demonstrate formulas
data = [
    ("Opere da Cap 02", "02.01.01", "Scavo di sbancamento", "Scavo di sbancamento in materie di qualsiasi natura...", "mc", 1, 10, 5, 2, 5.50),
    ("Opere da Cap 03", "03.05.12", "Calcestruzzo Rck 30", "Calcestruzzo per opere in c.a....", "mc", 2, 10, 0.5, 0.5, 120.00),
]

for i, row_data in enumerate(data):
    row_num = i + 2
    # Fill static data 
    ws.cell(row=row_num, column=1).value = row_data[0] # Categoria
    ws.cell(row=row_num, column=2).value = row_data[1] # Codice
    ws.cell(row=row_num, column=3).value = row_data[2] # Desc. Sint.
    ws.cell(row=row_num, column=4).value = row_data[3] # Desc. Est.
    ws.cell(row=row_num, column=5).value = row_data[4] # U.M.
    ws.cell(row=row_num, column=6).value = row_data[5] # N.
    ws.cell(row=row_num, column=7).value = row_data[6] # Lung.
    ws.cell(row=row_num, column=8).value = row_data[7] # Larg.
    ws.cell(row=row_num, column=9).value = row_data[8] # Alt.
    ws.cell(row=row_num, column=11).value = row_data[9] # Prezzo Unit.

    # Write FORMULAS
    # Quantity = N * Lung * Larg * Alt
    # Note: We use PRODUCT to handle empty cells nicely if needed, or simple multiplication logic
    # Actually, for simple excel, =F*G*H*I is best, but we need to ensure users put 1s if dim is missing or logic.
    # User asked for N*L*L*A. 
    # To avoid zeroing out if one is empty, user usually inputs dimensions. 
    # Let's put a robust formula: =PRODUCT(F2:I2) is safer for "empty as 1" ? No, empty is 0 in excel product usually? No, ignored.
    # Let's stick to explicit multiplication: =F2*G2*H2*I2
    
    qty_formula = f"=PRODUCT(F{row_num}:I{row_num})"
    ws.cell(row=row_num, column=10).value = qty_formula
    
    total_formula = f"=J{row_num}*K{row_num}"
    ws.cell(row=row_num, column=12).value = total_formula
    
    # Format Currency
    ws.cell(row=row_num, column=11).number_format = '#,##0.00 €'
    ws.cell(row=row_num, column=12).number_format = '#,##0.00 €'

# Adjust column widths
widths = [15, 12, 30, 40, 6, 5, 8, 8, 8, 12, 12, 15]
for i, w in enumerate(widths):
    ws.column_dimensions[get_column_letter(i+1)].width = w

# Save
wb.save("D:\\LIBRERIE\\ARCHITETTURA\\AI\\Modello_Computo_Umbria.xlsx")
print("Template created successfully.")
