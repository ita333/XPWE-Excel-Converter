import csv
import os

base_path = r"D:\LIBRERIE\ARCHITETTURA\PREZIARI\CSV\Umbria-2024-Diviso"

def search_code_prefix(filename, prefix, name):
    fpath = os.path.join(base_path, filename)
    print(f"--- Searching {name} ({prefix}) ---")
    try:
        with open(fpath, 'r', encoding='latin-1') as f:
            reader = csv.reader(f, delimiter=';')
            matches = []
            for row in reader:
                if len(row) < 3: continue
                if row[0].startswith(prefix):
                    # Save a few examples
                    print(f"MATCH: {row[0]} | {row[1][:60]}... | {row[5]} EUR")
                    matches.append(row)
                    if len(matches) > 3: break # just need a few to pick
    except Exception as e:
        print(e)

# Electrical Components based on memory
search_code_prefix("Cap_15_Elettrico_Fotovoltaico.csv", "15.01.0004", "Cavi / Infilaggio")
search_code_prefix("Cap_15_Elettrico_Fotovoltaico.csv", "15.03.0003", "Frutti / Interruttori")
search_code_prefix("Cap_15_Elettrico_Fotovoltaico.csv", "15.03.0010", "Placche")
