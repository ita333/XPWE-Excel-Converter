import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import csv
import os

# --- Configuration ---
BASE_PATH = r"D:\LIBRERIE\ARCHITETTURA\PREZIARI\CSV\Umbria-2024-Diviso"

def find_item(filename, keywords, exclude=None):
    path = os.path.join(BASE_PATH, filename)
    best_match = None
    try:
        with open(path, 'r', encoding='latin-1') as f:
            reader = csv.reader(f, delimiter=';')
            for row in reader:
                if len(row) < 3: continue
                # Basic filter
                text = (row[1] + " " + row[2]).lower()
                if all(k.lower() in text for k in keywords):
                    if exclude and any(e.lower() in text for e in exclude):
                        continue
                    # Prefer shorter code or specific logic? 
                    # Just take the first good match for now, or refine
                    best_match = {
                        "code": row[0], 
                        "desc_short": row[1], 
                        "desc_long": row[2], 
                        "um": row[3] if len(row)>3 else "", 
                        "price": float(row[5].replace('.','').replace(',','.')) if len(row)>5 else 0.0
                    }
                    return best_match
    except Exception as e:
        print(f"Error reading {filename}: {e}")
    return best_match

# --- 1. FIND ITEMS ---
items = {}

# Demolitions
items["rimozione_sanitari"] = find_item("Cap_02_Scavi_Demolizioni.csv", ["rimozione", "apparecchi", "sanitari"])
items["dem_pavimenti"] = find_item("Cap_02_Scavi_Demolizioni.csv", ["demolizione", "pavimenti", "rivestimenti"])

# New Installations (Cap 14)
# Look for "Allaccio e montaggio" supply excluded
items["inst_lavabo"] = find_item("Cap_14_Idrico_Sanitario.csv", ["allaccio", "montaggio", "lavabo", "esclusi"], ["lavapiedi"])
items["inst_wc"] = find_item("Cap_14_Idrico_Sanitario.csv", ["allaccio", "montaggio", "vaso", "esclusi"])
items["inst_bidet"] = find_item("Cap_14_Idrico_Sanitario.csv", ["allaccio", "montaggio", "bidet", "esclusi"])
items["inst_doccia"] = find_item("Cap_14_Idrico_Sanitario.csv", ["allaccio", "montaggio", "doccia", "esclusi"])

# Flooring (Cap 06) matches "Posa" and "Gres" 
# Ideally "Solo posa" but hard to find. We'll look for "Pavimento ... gres" and verify price is low < 30
items["posa_gres"] = find_item("Cap_06_Intonaci_Pavimenti.csv", ["pavimento", "gres", "posa"]) 

# Electrical (Cap 15)
# Level 2 -> We'll just put points
items["punto_luce"] = find_item("Cap_15_Elettrico_Fotovoltaico.csv", ["punto luce", "in traccia"]) # Canalizzazione

# Fallbacks if None
default_item = {"code": "MISSING", "desc_short": "Voce non trovata", "desc_long": "", "um": "cad", "price": 0.0}
for k in items:
    if items[k] is None:
        print(f"Warning: Could not find match for {k}")
        items[k] = default_item

# --- 2. BUILD EXCEL ---
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Computo Ristrutturazione"

# Headers
headers = [
    "Categoria", "Codice", "Descrizione Sintetica", "Descrizione Estesa", 
    "U.M.", "N.", "Lung.", "Larg.", "Alt./Peso", 
    "Quantità", "Prezzo Unit.", "Prezzo Totale"
]

for col_num, header in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col_num)
    cell.value = header
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    cell.alignment = Alignment(horizontal="center", vertical="center")

# Rows Logic
rows = []

# Cat: Demolizioni
rows.append(("DEMOLIZIONI", items["rimozione_sanitari"], 1, 1, 1, 1)) # 1 block of removals (existing bath)
# Actually existing bath has: 1 Lav, 1 Vasca, 1 Wc, 1 Bidet = 4 pezzi
rows.append(("DEMOLIZIONI", items["rimozione_sanitari"], 4, 1, 1, 1)) # Correcting to 4 pieces

# Demolizione Pavimenti (Existing 1 bath ~ 6mq)
rows.append(("DEMOLIZIONI", items["dem_pavimenti"], 1, 3, 2, 1)) # 6 mq

# Cat: Nuovi Impianti Idrici (3 Bagni)
# 3 x (Doccia, Lavabo, WC, Bidet)
rows.append(("IMPIANTI IDRICI", items["inst_doccia"], 3, 1, 1, 1))
rows.append(("IMPIANTI IDRICI", items["inst_lavabo"], 3, 1, 1, 1))
rows.append(("IMPIANTI IDRICI", items["inst_wc"], 3, 1, 1, 1))
rows.append(("IMPIANTI IDRICI", items["inst_bidet"], 3, 1, 1, 1))

# Cat: Pavimenti e Rivestimenti
# 3 Bagni. Assume 2x2m = 4mq each. Total 12mq floors.
rows.append(("FINITURE", items["posa_gres"], 3, 2, 2, 1)) 
# Walls: Perimeter (2+2+2+2)=8m * h2.40 = 19.2 mq * 3 baths
posa_riv_item = items["posa_gres"] # Reuse or find specific
rows.append(("FINITURE", posa_riv_item, 3, 8, 2.4, 1, "Rivestimento Pareti (Similare Posa)"))

# Cat: Impianto Elettrico
# Level 2 standard for apartment ~ 80mq implies ~40-60 points total usually. 
# For just the bathrooms refit + general update? "Rifatto impianto elettrico" means WHOLE apt or just baths?
# "ristrutturazione di un appartamento... verra rifatto anche l'impianto elettrico" -> Whole apartment.
# Assuming 90mq apartment -> ~80 punti luce totali (Standard Liv 2).
rows.append(("IMPIANTI ELETTRICI", items["punto_luce"], 80, 1, 1, 1))

# Write Data
current_row = 2
for cat, item, n, l, w, h, *opt_desc in rows:
    ws.cell(row=current_row, column=1).value = cat
    ws.cell(row=current_row, column=2).value = item["code"]
    ws.cell(row=current_row, column=3).value = opt_desc[0] if opt_desc else item["desc_short"]
    ws.cell(row=current_row, column=4).value = item["desc_long"]
    ws.cell(row=current_row, column=5).value = item["um"]
    ws.cell(row=current_row, column=6).value = n
    ws.cell(row=current_row, column=7).value = l
    ws.cell(row=current_row, column=8).value = w
    ws.cell(row=current_row, column=9).value = h
    ws.cell(row=current_row, column=11).value = item["price"]
    
    # Formulas
    ws.cell(row=current_row, column=10).value = f"=F{current_row}*G{current_row}*H{current_row}*I{current_row}"
    ws.cell(row=current_row, column=12).value = f"=J{current_row}*K{current_row}"
    
    # Format
    ws.cell(row=current_row, column=11).number_format = '#,##0.00 €'
    ws.cell(row=current_row, column=12).number_format = '#,##0.00 €'
    
    current_row += 1

# Total Row
total_row = current_row + 1
ws.cell(row=total_row, column=11).value = "TOTALE GENERALE"
ws.cell(row=total_row, column=11).font = Font(bold=True)
ws.cell(row=total_row, column=12).value = f"=SUM(L2:L{current_row-1})"
ws.cell(row=total_row, column=12).font = Font(bold=True)
ws.cell(row=total_row, column=12).number_format = '#,##0.00 €'

# Output
out_path = r"D:\LIBRERIE\ARCHITETTURA\AI\Computo_Ristrutturazione.xlsx"
wb.save(out_path)
print(f"File saved to {out_path}")
