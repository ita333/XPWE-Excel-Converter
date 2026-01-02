import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
import os

def convert_xlsx_to_xpwe(xlsx_path, xpwe_path):
    # Carica il file Excel
    df = pd.read_excel(xlsx_path)
    
    # Pulisce i nomi delle colonne da eventuali spazi
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]

    # Root element
    pwe_doc = ET.Element("PweDocumento")
    ET.SubElement(pwe_doc, "CopyRight").text = "Copyright ACCA software S.p.A."
    ET.SubElement(pwe_doc, "TipoDocumento").text = "1"
    ET.SubElement(pwe_doc, "TipoFormato").text = "XMLPwe"
    ET.SubElement(pwe_doc, "Versione").text = "5.04"
    ET.SubElement(pwe_doc, "SourceVersione").text = "Gemini-CLI-Converter"
    ET.SubElement(pwe_doc, "SourceNome").text = "PriMus-DCF"

    # Dati Generali
    dati_gen = ET.SubElement(pwe_doc, "PweDatiGenerali")
    dg_prog = ET.SubElement(dati_gen, "PweDGProgetto")
    dg_dati = ET.SubElement(dg_prog, "PweDGDatiGenerali")
    ET.SubElement(dg_dati, "Comune").text = "Comune di Conversione"
    ET.SubElement(dg_dati, "Oggetto").text = "Computo convertito da XLSX"

    # Categorie
    dg_cap_cat = ET.SubElement(dati_gen, "PweDGCapitoliCategorie")
    dg_cat_list = ET.SubElement(dg_cap_cat, "PweDGCategorie")
    
    categories = df['Categoria'].unique()
    cat_map = {}
    for i, cat in enumerate(categories, 1):
        if pd.isna(cat): continue
        cat_item = ET.SubElement(dg_cat_list, "DGCategorieItem", ID=str(i))
        ET.SubElement(cat_item, "DesSintetica").text = str(cat)
        cat_map[cat] = i

    # Misurazioni wrapper
    misurazioni = ET.SubElement(pwe_doc, "PweMisurazioni")

    # Elenco Prezzi (EP)
    # Raccogliamo voci uniche basate sul Codice
    ep_list = ET.SubElement(misurazioni, "PweElencoPrezzi")
    unique_items = df.drop_duplicates(subset=['Codice'])
    ep_map = {} # Code -> ID
    
    for i, (_, row) in enumerate(unique_items.iterrows(), 1):
        if pd.isna(row['Codice']): continue
        
        ep_item = ET.SubElement(ep_list, "EPItem", ID=str(i))
        ET.SubElement(ep_item, "TipoEP").text = "0"
        ET.SubElement(ep_item, "Tariffa").text = str(row['Codice'])
        ET.SubElement(ep_item, "DesRidotta").text = str(row.get('Descrizione Sintetica', ''))
        ET.SubElement(ep_item, "DesEstesa").text = str(row.get('Descrizione Estesa', ''))
        
        um = str(row.get('U.M.', ''))
        ET.SubElement(ep_item, "UnMisura").text = um if um.lower() != "nan" else "cad"
        
        pz = str(row.get('Prezzo Unit.', '0')).replace(',', '.')
        ET.SubElement(ep_item, "Prezzo1").text = pz if pz.lower() != "nan" else "0"
        
        ep_map[row['Codice']] = i

    # Voci Computo (VC)
    vc_list = ET.SubElement(misurazioni, "PweVociComputo")
    
    # Raggruppiamo per Categoria e Codice per creare le voci di computo
    # Ogni riga dell'excel diventa un rigo di misura (RGItem)
    # Per semplicità, ogni blocco di righe uguali per Categoria/Codice/DescSintetica lo mettiamo in un VCItem
    
    current_vc = None
    last_key = None
    rg_id_counter = 1
    vc_id_counter = 1

    for _, row in df.iterrows():
        if pd.isna(row['Codice']): continue
        
        key = (row['Categoria'], row['Codice'], row.get('Descrizione Sintetica'))
        
        if key != last_key:
            # Crea nuova Voce di Computo
            current_vc = ET.SubElement(vc_list, "VCItem", ID=str(vc_id_counter))
            vc_id_counter += 1
            
            idep = ep_map.get(row['Codice'], 0)
            idcat = cat_map.get(row['Categoria'], 0)
            
            ET.SubElement(current_vc, "IDEP").text = str(idep)
            # Quantita is usually AFTER IDEP in VCItem
            vc_qty_elem = ET.SubElement(current_vc, "Quantita")
            vc_qty_elem.text = "0.0000"
            ET.SubElement(current_vc, "DataMis").text = "30/12/1899"
            ET.SubElement(current_vc, "Flags").text = "0"
            ET.SubElement(current_vc, "IDSpCat").text = "0"
            ET.SubElement(current_vc, "IDCat").text = str(idcat)
            ET.SubElement(current_vc, "IDSbCat").text = "0"
            ET.SubElement(current_vc, "CodiceWBS").text = ""

            misure_list = ET.SubElement(current_vc, "PweVCMisure")
            last_key = key
        else:
            misure_list = current_vc.find("PweVCMisure")

        # Aggiungi rigo di misura
        rg_item = ET.SubElement(misure_list, "RGItem", ID=str(rg_id_counter))
        rg_id_counter += 1
        
        ET.SubElement(rg_item, "IDVV").text = "-2"
        # Usiamo la descrizione sintetica come descrizione del rigo se non c'è altro
        desc_rigo = str(row.get('Descrizione Sintetica', 'Misura'))
        ET.SubElement(rg_item, "Descrizione").text = desc_rigo
        
        # Variabili - Preserve formulas if they start with =
        def fmt_val(val):
            if pd.isna(val) or val == "": return ""
            s = str(val).strip()
            if s.startswith("="): return s[1:].replace(',', '.') # Remove '=' and normalize decimal
            return s.replace(',', '.')

        ET.SubElement(rg_item, "PartiUguali").text = fmt_val(row.get('N.'))
        ET.SubElement(rg_item, "Lunghezza").text = fmt_val(row.get('Lung.'))
        ET.SubElement(rg_item, "Larghezza").text = fmt_val(row.get('Larg.'))
        ET.SubElement(rg_item, "HPeso").text = fmt_val(row.get('Alt./Peso'))
        
        # Calcolo Quantità
        def get_num(val):
            if pd.isna(val) or val == "": return 1.0
            s = str(val).replace(',', '.').strip()
            if s.startswith("="):
                try: 
                    # Rough math parser for basic formulas in XLSX
                    clean_s = re.sub(r'[^0-9\.\+\-\*\/\(\)]', '', s[1:])
                    return float(eval(clean_s))
                except: return 1.0
            try: return float(s)
            except: return 1.0

        try:
            # We only multiply factors that are not "empty" (represented by 1.0 in get_num but we need to check original)
            factors = []
            for k in ['N.', 'Lung.', 'Larg.', 'Alt./Peso']:
                v = row.get(k)
                if not pd.isna(v) and str(v).strip() != "":
                    factors.append(get_num(v))
            
            if not factors:
                q = 0.0
            else:
                q = 1.0
                for f in factors: q *= f
            
            ET.SubElement(rg_item, "Quantita").text = "{:.4f}".format(q)
            
            # Update Volume total in VCItem
            current_total = float(vc_qty_elem.text)
            vc_qty_elem.text = "{:.4f}".format(current_total + q)
        except:
            ET.SubElement(rg_item, "Quantita").text = "0.0000"
        
        ET.SubElement(rg_item, "Flags").text = "0"
    # Salvataggio
    xml_str = ET.tostring(pwe_doc, encoding='utf-8')
    # PriMus vuole l'header specifico
    header = b'<?mso-application progid="PriMus.Document.XPWE"?>'
    
    with open(xpwe_path, "wb") as f:
        f.write(header + xml_str)

if __name__ == "__main__":
    import sys
    xlsx_in = sys.argv[1] if len(sys.argv) > 1 else "EsportaComputo-v4.xlsx"
    xpwe_out = sys.argv[2] if len(sys.argv) > 2 else "strutture_esempio2_fixed.xpwe"
    
    if os.path.exists(xlsx_in):
        convert_xlsx_to_xpwe(xlsx_in, xpwe_out)
        print(f"Conversione completata: {xpwe_out}")
    else:
        print(f"Errore: Il file {xlsx_in} non esiste.")
