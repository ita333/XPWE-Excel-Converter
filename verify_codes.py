import csv
import os

codes = [
    "14.01.0020.001",
    "14.01.0020.003",
    "14.01.0020.005",
    "14.01.0020.007", # Guessing
    "14.01.0020.009", # Guessing
    "02.04.0140",
    "02.04.0130",   # Rimozione sanitari found earlier
    "06.09.0200",   # Check if supply included
    "06.09.0210",
    "15.01.0001"
]

base_path = r"D:\LIBRERIE\ARCHITETTURA\PREZIARI\CSV\Umbria-2024-Diviso"
files = ["Cap_14_Idrico_Sanitario.csv", "Cap_02_Scavi_Demolizioni.csv", "Cap_06_Intonaci_Pavimenti.csv", "Cap_15_Elettrico_Fotovoltaico.csv"]

found = {}

print("Verifying codes...")
for fname in files:
    fpath = os.path.join(base_path, fname)
    try:
        with open(fpath, 'r', encoding='latin-1') as f:
            reader = csv.reader(f, delimiter=';')
            for row in reader:
                if len(row) < 3: continue
                # key is row[0].strip()
                # Check if code starts with one of our targets (formatting can vary)
                for c in codes:
                    if row[0].startswith(c):
                        print(f"CODE: {row[0]}")
                        print(f"DESC: {row[1]}")
                        print(f"FULL: {row[2][:100]}...")
                        print(f"PRICE: {row[5]}")
                        print("-" * 30)
    except:
        pass
