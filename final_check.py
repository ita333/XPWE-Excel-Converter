import csv
import os

base_path = r"D:\LIBRERIE\ARCHITETTURA\PREZIARI\CSV\Umbria-2024-Diviso"

def search(file, keywords, name):
    fpath = os.path.join(base_path, file)
    print(f"Searching {name}...")
    try:
        with open(fpath, 'r', encoding='latin-1') as f:
            reader = csv.reader(f, delimiter=';')
            for row in reader:
                if len(row) < 3: continue
                match = True
                for k in keywords:
                    if k.lower() not in row[1].lower() and k.lower() not in row[2].lower():
                        match = False
                        break
                if match:
                    print(f"MATCH: {row[0]} | Price: {row[5]}")
                    print(f"DESC: {row[1]}")
                    return row # Return first match
    except Exception as e:
        print(e)
    return None

# Electrical: Try to find a complete point or at least the box/cable
search("Cap_15_Elettrico_Fotovoltaico.csv", ["punto luce", "deviato", "completo"], "Punto Luce Completo")
search("Cap_15_Elettrico_Fotovoltaico.csv", ["punto luce", "semplice"], "Punto Luce Semplice")

# Flooring: Solo posa
search("Cap_06_Intonaci_Pavimenti.csv", ["pavimento", "gres", "colla", "posa"], "Posa Gres")
search("Cap_06_Intonaci_Pavimenti.csv", ["rivestimento", "ceramica", "posa", "colla"], "Posa Rivestimento")
