import csv
import sys

filename = r"D:\LIBRERIE\ARCHITETTURA\PREZIARI\CSV\Umbria-2024-Diviso\Cap_14_Idrico_Sanitario.csv"

search_terms = ["esclusi: la fornitura"]
targets = ["lavabo", "vaso", "bidet", "doccia"]

try:
    with open(filename, 'r', encoding='latin-1') as f: # standard encoding for these CSVs usually
        reader = csv.reader(f, delimiter=';')
        for row in reader:
            if len(row) < 3: continue
            code = row[0]
            desc_short = row[1]
            desc_long = row[2]
            
            # Check if supply is excluded
            if "esclusi: la fornitura" in desc_long.lower() or "esclusi: la fornitura" in desc_short.lower():
                # Check for target keywords
                hits = [t for t in targets if t in desc_short.lower() or t in desc_long.lower()]
                if hits:
                    print(f"FOUND: {code} | {desc_short} | {row[5]} EUR")
                    print(f"DESC: {desc_long[:100]}...")
                    print("-" * 50)
except Exception as e:
    print(f"Error: {e}")
