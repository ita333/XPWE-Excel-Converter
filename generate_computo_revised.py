import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import csv
import os

# --- Configuration ---
BASE_PATH = r"D:\LIBRERIE\ARCHITETTURA\PREZIARI\CSV\Umbria-2024-Diviso"

def find_item_exact_prefix(filename, prefix):
    path = os.path.join(BASE_PATH, filename)
    try:
        with open(path, 'r', encoding='latin-1') as f:
            reader = csv.reader(f, delimiter=';')
            for row in reader:
                if len(row) < 3: continue
                if row[0].startswith(prefix):
                    # Clean price
                    p_str = row[5].replace('.','').replace(',','.').strip()
                    try:
                        price = float(p_str)
                    except:
                        price = 0.0
                    return {
                        "code": row[0],
                        "desc_short": row[1],
                        "desc_long": row[2],
                        "um": row[3],
                        "price": price
                    }
    except Exception as e:
        print(f"Error reading {filename}: {e}")
    return None

def find_item_keywords(filename, keywords, exclude=None):
    path = os.path.join(BASE_PATH, filename)
    try:
        with open(path, 'r', encoding='latin-1') as f:
            reader = csv.reader(f, delimiter=';')
            for row in reader:
                if len(row) < 3: continue
                text = (row[1] + " " + row[2]).lower()
                if all(k.lower() in text for k in keywords):
                    if exclude and any(e.lower() in text for e in exclude):
                        continue
                     # Clean price
                    p_str = row[5].replace('.','').replace(',','.').strip()
                    try:
                        price = float(p_str)
                    except:
                        price = 0.0
                    return {
                        "code": row[0], 
                        "desc_short": row[1], 
                        "desc_long": row[2], 
                        "um": row[3], 
                        "price": price
                    }
    except:
        pass
    return None

# --- 1. GATHER DATA ---
data_map = {}

# A. DEMOLIZIONI
data_map["rimoz_sanitari"] = find_item_exact_prefix("Cap_02_Scavi_Demolizioni.csv", "02.04.0130") # Rimozione apparecchi
data_map["dem_pav"] = find_item_exact_prefix("Cap_02_Scavi_Demolizioni.csv", "02.04.0140") # Demolizione pavimenti (Generic search found this earlier or similar)
if not data_map["dem_pav"]:
    # Fallback search if exact code failed
    data_map["dem_pav"] = find_item_keywords("Cap_02_Scavi_Demolizioni.csv", ["demolizione", "pavimenti", "rivestimenti", "ceramica"])

# B. OPERE EDILI (New Floors)
data_map["posa_pav"] = find_item_keywords("Cap_06_Intonaci_Pavimenti.csv", ["pavimento", "gres", "posa"])
data_map["posa_riv"] = find_item_keywords("Cap_06_Intonaci_Pavimenti.csv", ["rivestimento", "ceramica", "posa"])

# C. IMPIANTI IDRICI (Supply Excluded)
# Exact prefix based on previous research
data_map["water_start"] = find_item_exact_prefix("Cap_14_Idrico_Sanitario.csv", "14.01.0020") # Base code check
# We need specific inputs or we reuse the item found. 
# Search specific "allaccio" items
data_map["wc_inst"] = find_item_keywords("Cap_14_Idrico_Sanitario.csv", ["allaccio", "montaggio", "vaso", "esclusi"])
data_map["bidet_inst"] = find_item_keywords("Cap_14_Idrico_Sanitario.csv", ["allaccio", "montaggio", "bidet", "esclusi"])
data_map["lavabo_inst"] = find_item_keywords("Cap_14_Idrico_Sanitario.csv", ["allaccio", "montaggio", "lavabo", "esclusi"])
data_map["doccia_inst"] = find_item_keywords("Cap_14_Idrico_Sanitario.csv", ["allaccio", "montaggio", "doccia", "esclusi"])

# D. IMPIANTI ELETTRICI (Composite)
data_map["ele_tube"] = find_item_exact_prefix("Cap_15_Elettrico_Fotovoltaico.csv", "15.01.0001") # Canalizzazione base
data_map["ele_wire"] = find_item_exact_prefix("Cap_15_Elettrico_Fotovoltaico.csv", "15.01.0004.001") # Filo 1.5mmq est.
data_map["ele_box"] = find_item_exact_prefix("Cap_15_Elettrico_Fotovoltaico.csv", "15.03.0010.001") # Scatola/Placca
data_map["ele_switch"] = find_item_keywords("Cap_15_Elettrico_Fotovoltaico.csv", ["interruttore", "unipolare", "10", "16"], ["magnetotermico"]) 
# If exact code 15.03.0003 exists use it
f_switch = find_item_exact_prefix("Cap_15_Elettrico_Fotovoltaico.csv", "15.03.0003")
if f_switch: data_map["ele_switch"] = f_switch

data_map["ele_board"] = find_item_keywords("Cap_15_Elettrico_Fotovoltaico.csv", ["centralino", "incasso", "moduli"]) # Small board

# Defaults
for k in data_map:
    if not data_map[k]:
        data_map[k] = {"code": "MISSING", "desc_short": f"Voce {k} non trovata", "desc_long": "", "um": "cad", "price": 0.0}

# --- 2. BUILD EXCEL ---
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Computo Ristrutturazione Rev"

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
# (Category, Item, N, L, W, H, Optional Note)
rows = []

# A.01 DEMOLIZIONI
rows.append(("A.01 - DEMOLIZIONI", data_map["rimoz_sanitari"], 4, 1, 1, 1, "Rimozione sanitari esistenti (Lav+Vasca+Wc+Bidet)"))
rows.append(("A.01 - DEMOLIZIONI", data_map["dem_pav"], 1, 3, 2, 1, "Demolizione pav. esistente (Bagno 6mq)"))
rows.append(("A.01 - DEMOLIZIONI", data_map["dem_pav"], 1, 8, 2.4, 1, "Demolizione riv. esistente (Pareti)"))

# B.01 OPERE EDILI
# 3 New Bathrooms. Assume 2x2. Wall Perim = 8m. Height 2.40m
rows.append(("B.01 - PAVIMENTI E RIVESTIMENTI", data_map["posa_pav"], 3, 2, 2, 1, "Posa Gres Nuovi Bagni"))
rows.append(("B.01 - PAVIMENTI E RIVESTIMENTI", data_map["posa_riv"], 3, 8, 2.4, 1, "Posa Rivestimento Nuovi Bagni"))

# C.01 IMPIANTI IDRICI
rows.append(("C.01 - IMPIANTI IDRICO-SANITARI", data_map["doccia_inst"], 3, 1, 1, 1))
rows.append(("C.01 - IMPIANTI IDRICO-SANITARI", data_map["lavabo_inst"], 3, 1, 1, 1))
rows.append(("C.01 - IMPIANTI IDRICO-SANITARI", data_map["wc_inst"], 3, 1, 1, 1))
rows.append(("C.01 - IMPIANTI IDRICO-SANITARI", data_map["bidet_inst"], 3, 1, 1, 1))

# D.01 IMPIANTI ELETTRICI (COMPOSIZIONE PUNTO LUCE STANDARD)
# Approx 80 points total
rows.append(("D.01 - IMPIANTI ELETTRICI", data_map["ele_tube"], 80, 1, 1, 1, "Canalizzazione Punti Luce"))
rows.append(("D.01 - IMPIANTI ELETTRICI", data_map["ele_wire"], 80, 1, 1, 1, "Infilaggio Cavi (quota parte)"))
rows.append(("D.01 - IMPIANTI ELETTRICI", data_map["ele_box"], 80, 1, 1, 1, "Scatole e Placche"))
rows.append(("D.01 - IMPIANTI ELETTRICI", data_map["ele_switch"], 80, 1, 1, 1, "Frutti (Interruttori/Prese)"))
rows.append(("D.01 - IMPIANTI ELETTRICI", data_map["ele_board"], 1, 1, 1, 1, "Quadro Elettrico Generale"))

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
    
    # Simple Formulas
    ws.cell(row=current_row, column=10).value = f"=F{current_row}*G{current_row}*H{current_row}*I{current_row}"
    ws.cell(row=current_row, column=12).value = f"=J{current_row}*K{current_row}"
    
    ws.cell(row=current_row, column=11).number_format = '#,##0.00 €'
    ws.cell(row=current_row, column=12).number_format = '#,##0.00 €'
    
    current_row += 1

# Total
total_row = current_row + 1
ws.cell(row=total_row, column=11).value = "TOTALE GENERALE"
ws.cell(row=total_row, column=11).font = Font(bold=True)
ws.cell(row=total_row, column=12).value = f"=SUM(L2:L{current_row-1})"
ws.cell(row=total_row, column=12).font = Font(bold=True)
ws.cell(row=total_row, column=12).number_format = '#,##0.00 €'

# Output
out_path = r"D:\LIBRERIE\ARCHITETTURA\AI\Computo_Ristrutturazione_Corretto.xlsx"
wb.save(out_path)
print(f"File saved to {out_path}")
