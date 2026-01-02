import csv
import os

base_path = r"D:\LIBRERIE\ARCHITETTURA\PREZIARI\CSV\Umbria-2024-Diviso"

queries = [
    {
        "file": "Cap_14_Idrico_Sanitario.csv",
        "filter": lambda r: "esclusi: la fornitura" in r[2].lower() and "lavabo" in r[2].lower(),
        "name": "Lavabo (Montaggio)"
    },
    {
        "file": "Cap_14_Idrico_Sanitario.csv",
        "filter": lambda r: "esclusi: la fornitura" in r[2].lower() and "vaso" in r[2].lower(),
        "name": "WC (Montaggio)"
    },
    {
        "file": "Cap_14_Idrico_Sanitario.csv",
        "filter": lambda r: "esclusi: la fornitura" in r[2].lower() and "bidet" in r[2].lower(),
        "name": "Bidet (Montaggio)"
    },
    {
        "file": "Cap_14_Idrico_Sanitario.csv",
        "filter": lambda r: "esclusi: la fornitura" in r[2].lower() and ("doccia" in r[2].lower() or "piatto" in r[2].lower()),
        "name": "Doccia (Montaggio)"
    },
    {
        "file": "Cap_02_Scavi_Demolizioni.csv",
        "filter": lambda r: "pavimenti" in r[1].lower() and "demolizione" in r[1].lower(),
        "name": "Demolizione Pavimenti"
    },
    {
        "file": "Cap_02_Scavi_Demolizioni.csv",
        "filter": lambda r: "rivestimenti" in r[1].lower() and "demolizione" in r[1].lower(),
        "name": "Demolizione Rivestimenti"
    },
    {
        "file": "Cap_06_Intonaci_Pavimenti.csv",
        "filter": lambda r: "gres" in r[2].lower() and "posa" in r[1].lower(),
        "name": "Posa Pavimento Gres"
    },
    {
        "file": "Cap_06_Intonaci_Pavimenti.csv",
        "filter": lambda r: "ceramica" in r[2].lower() and "rivestiment" in r[1].lower(),
        "name": "Posa Rivestimento"
    },
    {
        "file": "Cap_15_Elettrico_Fotovoltaico.csv",
        "filter": lambda r: "punto luce" in r[1].lower(),
        "name": "Punto Luce"
    }
]

for q in queries:
    fpath = os.path.join(base_path, q["file"])
    print(f"Searching for {q['name']} in {q['file']}...")
    try:
        with open(fpath, 'r', encoding='latin-1') as f:
            reader = csv.reader(f, delimiter=';')
            cnt = 0
            for row in reader:
                if len(row) < 3: continue
                if q["filter"](row):
                    print(f"  MATCH: {row[0]} | {row[1][:40]}... | {row[5]} EUR")
                    cnt += 1
                    if cnt >= 3: break # Take first 3 matches
    except Exception as e:
        print(f"Error reading {q['file']}: {e}")
    print("-" * 30)
