import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import xml.etree.ElementTree as ET
import os
import re
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# --- UTILS ---
def safe_eval_math(val):
    if pd.isna(val) or str(val).strip() == "" or str(val).lower() == "nan": return 0.0
    s_val = str(val).replace(',', '.').strip()
    if s_val.startswith("="): s_val = s_val[1:]
    try: return float(eval(s_val))
    except: return 0.0
    if not re.match(r'^[0-9\.\+\-\*\/\(\)\s]+$', s_val):
        try: return float(s_val)
        except: return 0.0
    try: return float(eval(s_val))
    except: return 0.0

def clean_val_for_excel(val):
    if pd.isna(val) or val == "": return None
    s = str(val).replace(',', '.').strip()
    try: return float(s)
    except: pass
    if any(c in s for c in "+-*/") and not any(c.isalpha() for c in s): return "=" + s
    return val

def safe_fmt_num(val, fmt="{:.4f}"):
    try:
        if pd.isna(val) or str(val).lower() == "nan": return ""
        v = float(str(val).replace(',', '.'))
        return fmt.format(v)
    except:
        return str(val)

def fmt_xpwe_num(val): return safe_fmt_num(val, "{:.4f}")
def fmt_qty(val): return "{:,.4f}".format(val).replace(',', 'X').replace('.', ',').replace('X', '.') if val!="" and val is not None else ""
def fmt_money(val): return "{:,.2f} €".format(val).replace(',', 'X').replace('.', ',').replace('X', '.') if val!="" and val is not None else ""

# --- XPWE STRICT HANDLER ---
class XPWEHandler:
    @staticmethod
    def read_xpwe(file_path):
        try:
            tree = ET.parse(file_path); root = tree.getroot()
            cat_map = {c.get("ID"): (c.find("DesSintetica").text or "Generale") for c in root.findall(".//DGCategorieItem")}
            ep_map = {}
            for item in root.findall(".//EPItem"):
                ep_map[item.get("ID")] = {
                    "Codice": item.find("Tariffa").text or "n.d.",
                    "Descrizione Sintetica": item.find("DesRidotta").text or "",
                    "Descrizione Estesa": item.find("DesEstesa").text or "",
                    "U.M.": item.find("UnMisura").text or "n.d.",
                    "Prezzo Unit.": item.find("Prezzo1").text or "0"
                }
            rows = []
            for vc in root.findall(".//VCItem"):
                idep = vc.find("IDEP").text
                idcat = vc.find("IDCat").text
                ep = ep_map.get(idep, {})
                cat = cat_map.get(idcat, "Generale")
                misure = vc.find("PweVCMisure")
                if misure is not None:
                    for rg in misure.findall("RGItem"):
                        rows.append({
                            "Categoria": cat,
                            "Codice": ep.get("Codice", ""),
                            "Descrizione Sintetica": ep.get("Descrizione Sintetica", ""),
                            "Descrizione Estesa": ep.get("Descrizione Estesa", ""),
                            "Descrizione Rigo": rg.find("Descrizione").text or "",
                            "U.M.": ep.get("U.M.", ""),
                            "N.": rg.find("PartiUguali").text or "",
                            "Lung.": rg.find("Lunghezza").text or "",
                            "Larg.": rg.find("Larghezza").text or "",
                            "Alt./Peso": rg.find("HPeso").text or "",
                            "Quantità": rg.find("Quantita").text or "",
                            "Prezzo Unit.": ep.get("Prezzo Unit.", "0")
                        })
            return pd.DataFrame(rows)
        except Exception as e:
            raise Exception(f"Errore lettura XPWE: {e}")

    @staticmethod
    def write_xpwe(df_raw, file_path):
        df = df_raw.copy(); df.fillna("", inplace=True)
        for c in ["Codice", "Categoria"]: 
             if c in df.columns: df[c] = df[c].astype(str).str.strip()
        df = df[df['Codice'] != ""]
        
        # ROOT
        pwe = ET.Element("PweDocumento")
        ET.SubElement(pwe, "CopyRight").text = "Software di conversione"
        ET.SubElement(pwe, "TipoDocumento").text = "1"
        ET.SubElement(pwe, "TipoFormato").text = "XMLPwe"
        ET.SubElement(pwe, "Versione").text = "5.04"
        ET.SubElement(pwe, "SourceVersione").text = "Gemini-CLI-Converter"
        ET.SubElement(pwe, "SourceNome").text = "PriMus-DCF"

        # DATI GENERALI
        dg = ET.SubElement(pwe, "PweDatiGenerali")
        dg_prog = ET.SubElement(dg, "PweDGProgetto")
        dg_dati = ET.SubElement(dg_prog, "PweDGDatiGenerali")
        ET.SubElement(dg_dati, "PercPrezzi").text = "0"
        ET.SubElement(dg_dati, "Oggetto").text = "Computo convertito da XLSX"

        # CATEGORIE
        dg_cat_wrapper = ET.SubElement(dg, "PweDGCapitoliCategorie")
        dg_cat = ET.SubElement(dg_cat_wrapper, "PweDGCategorie")
        
        # Gestione robusta della colonna Categoria
        if 'Categoria' not in df.columns: df['Categoria'] = ""
        cats = sorted(list(set(df['Categoria'].unique())))
        if not cats or (len(cats)==1 and cats[0]==""): cats = [""]
        
        cat_id_map = {}
        for i, cname in enumerate(cats, 1):
            cid = str(i)
            it = ET.SubElement(dg_cat, "DGCategorieItem", ID=cid)
            ET.SubElement(it, "DesSintetica").text = str(cname)
            ET.SubElement(it, "Codice").text = ""
            cat_id_map[str(cname)] = cid
        
        # --- MISSING SECTIONS (WBS, Moduli, Config) ---
        # WBS
        wbs = ET.SubElement(dg, "PweDGWBS")
        ET.SubElement(wbs, "DGWBSAttiva").text = "0"
        for i, (uid, tit) in enumerate([("{00000000-0000-0000-0000-000000000001}", "WBS"), ("{00000000-0000-0000-0000-000000000002}", "<NON assegnata>")], 2):
            wb_it = ET.SubElement(wbs, "DGWBSItem", ID=str(i))
            ET.SubElement(wb_it, "CU").text = uid
            ET.SubElement(wb_it, "CUParent").text = "{00000000-0000-0000-0000-000000000001}" if i==3 else ""
            ET.SubElement(wb_it, "TITOLO").text = tit
        # MISURAZIONI WRAPPER
        misurazioni = ET.SubElement(pwe, "PweMisurazioni")

        # ELENCO PREZZI
        ep_list = ET.SubElement(misurazioni, "PweElencoPrezzi")
        unique_codes = df.drop_duplicates(subset=['Codice'])
        ep_id_map = {}
        
        for i, (_, row) in enumerate(unique_codes.iterrows(), 1):
            eid = str(i)
            codice = str(row['Codice'])
            it = ET.SubElement(ep_list, "EPItem", ID=eid)
            ET.SubElement(it, "TipoEP").text = "0"
            ET.SubElement(it, "Tariffa").text = codice
            ET.SubElement(it, "DesRidotta").text = str(row.get('Descrizione Sintetica', ''))
            ET.SubElement(it, "DesEstesa").text = str(row.get('Descrizione Estesa', ''))
            
            um = str(row.get('U.M.', ''))
            ET.SubElement(it, "UnMisura").text = um if um.lower().strip() != "nan" and um.strip() != "" else "cad"
            
            pz_val = safe_eval_math(row.get('Prezzo Unit.', '0'))
            ET.SubElement(it, "Prezzo1").text = fmt_xpwe_num(pz_val)
            for k in range(2, 6): ET.SubElement(it, f"Prezzo{k}").text = "0"
            
            ep_id_map[codice] = eid

        # VOCI COMPUTO
        vc_list = ET.SubElement(misurazioni, "PweVociComputo")
        grouped = df.groupby(['Categoria', 'Codice'], sort=False)
        vc_counter = 1; rg_counter = 1
        
        for (cat, cod), grp in grouped:
            vc = ET.SubElement(vc_list, "VCItem", ID=str(vc_counter))
            ET.SubElement(vc, "IDEP").text = ep_id_map.get(str(cod), "0")
            
            vc_qty_elem = ET.SubElement(vc, "Quantita")
            vc_qty_elem.text = "0.0000"

            ET.SubElement(vc, "DataMis").text = "30/12/1899"
            ET.SubElement(vc, "Flags").text = "0"
            ET.SubElement(vc, "IDSpCat").text = "0"
            ET.SubElement(vc, "IDCat").text = cat_id_map.get(str(cat), "1")
            ET.SubElement(vc, "IDSbCat").text = "0"
            ET.SubElement(vc, "CodiceWBS").text = ""
            
            mis_tag = ET.SubElement(vc, "PweVCMisure")
            total_qty_voce = 0.0
            
            for _, row in grp.iterrows():
                # Calcolo quantità riga misurazione
                vn=safe_eval_math(row.get('N.')); vl=safe_eval_math(row.get('Lung.')); vw=safe_eval_math(row.get('Larg.')); 
                # Cerca Alt./Peso (nome interno standard)
                alt_val = row.get('Alt./Peso', 0)
                vh = safe_eval_math(alt_val)
                
                # Formula Primus: se tutti 0 -> usa quantità esplicita. Se alcuni >0, gli altri valgono 1.
                if vn==0 and vl==0 and vw==0 and vh==0:
                    q = safe_eval_math(row.get('Quantità', 0))
                else:
                    q = (vn if vn!=0 else 1) * (vl if vl!=0 else 1) * (vw if vw!=0 else 1) * (vh if vh!=0 else 1)
                
                total_qty_voce += q
                
                rg = ET.SubElement(mis_tag, "RGItem", ID=str(rg_counter))
                rg_counter += 1
                ET.SubElement(rg, "IDVV").text = "-2"
                ET.SubElement(rg, "Descrizione").text = str(row.get('Descrizione Rigo', '') or row.get('Descrizione Sintetica', ''))
                
                def fmt_val(v):
                    s_v = str(v).strip()
                    if s_v.startswith("="): return s_v[1:].replace(',', '.')
                    return fmt_xpwe_num(v) if safe_eval_math(v)!=0 else ""

                ET.SubElement(rg, "PartiUguali").text = fmt_val(row.get('N.'))
                ET.SubElement(rg, "Lunghezza").text = fmt_val(row.get('Lung.'))
                ET.SubElement(rg, "Larghezza").text = fmt_val(row.get('Larg.'))
                ET.SubElement(rg, "HPeso").text = fmt_val(alt_val)
                ET.SubElement(rg, "Quantita").text = "{:.4f}".format(q)
                ET.SubElement(rg, "Flags").text = "0"

            vc_qty_elem.text = "{:.4f}".format(total_qty_voce)
            vc_counter += 1

        # Salvataggio con header mso-application
        with open(file_path, "wb") as f:
            f.write(b'<?mso-application progid="PriMus.Document.XPWE"?>\n')
            ET.ElementTree(pwe).write(f, encoding='utf-8', xml_declaration=False)

# --- EXCEL HANDLER ---
class ExcelHandler:
    @staticmethod
    def write_excel_pro(df_raw, file_path):
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Computo"
        h_fill = PatternFill(start_color="4F81BD", fill_type="solid"); t_fill = PatternFill(start_color="F2F2F2", fill_type="solid"); g_fill = PatternFill(start_color="C0C0C0", fill_type="solid")
        border = Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin')); align = Alignment(horizontal="left", vertical="top", wrap_text=True)
        headers = ["N. Ord.", "Categoria", "Codice", "Descrizione Sintetica", "Descrizione Estesa", "Descrizione Rigo", "U.M.", "N.", "Lung.", "Larg.", "Alt.", "Quantità", "Prezzo Unit.", "Importo Totale"]
        for c, t in enumerate(headers, 1):
            cell = ws.cell(row=1, column=c, value=t); cell.font = Font(bold=True, color="FFFFFF"); cell.fill = h_fill; cell.alignment = Alignment(horizontal="center"); cell.border = border
        df = df_raw.copy(); df.fillna("", inplace=True)
        grouped = df.groupby(['Categoria', 'Codice'], sort=False)
        r = 2; vc = 1; imps = []
        for (cat, cod), grp in grouped:
            f = grp.iloc[0]
            ws.cell(r,1,vc).font=Font(bold=True); ws.cell(r,2,cat); ws.cell(r,3,cod).font=Font(bold=True); ws.cell(r,4,f.get('Descrizione Sintetica','')).font=Font(bold=True); ws.cell(r,5,f.get('Descrizione Estesa',''))
            for c in range(1,15): ws.cell(r,c).border=border; ws.cell(r,c).alignment=align
            r+=1; start=r
            for _,row in grp.iterrows():
                ws.cell(r,6,row.get('Descrizione Rigo','') or row.get('Descrizione Sintetica',''))
                for i,k in enumerate(['N.','Lung.','Larg.','Alt./Peso'],8): ws.cell(r,i,clean_val_for_excel(row.get(k,'')))
                ws.cell(r,12,f"=PRODUCT({get_column_letter(8)}{r}:{get_column_letter(11)}{r})").number_format='#,##0.00'
                for c in range(1,15): ws.cell(r,c).border=border; ws.cell(r,c).alignment=align
                r+=1
            ws.cell(r,6,"TOTALE VOCE").font=Font(bold=True); ws.cell(r,6).alignment=Alignment("right")
            ws.cell(r,7,f.get('U.M.','')).font=Font(bold=True)
            ws.cell(r,12,f"=SUM({get_column_letter(12)}{start}:{get_column_letter(12)}{r-1})").font=Font(bold=True); ws.cell(r,12).number_format='#,##0.00'
            ws.cell(r,13,clean_val_for_excel(f.get('Prezzo Unit.',0))).font=Font(bold=True); ws.cell(r,13).number_format='#,##0.00 €'
            ws.cell(r,14,f"={get_column_letter(12)}{r}*{get_column_letter(13)}{r}").font=Font(bold=True); ws.cell(r,14).number_format='#,##0.00 €'; ws.cell(r,14).fill=t_fill
            imps.append(f"{get_column_letter(14)}{r}"); r+=2; vc+=1
            for c in range(1,15): ws.cell(r-2,c).border=border
        if imps:
            ws.cell(r,4,"TOTALE GENERALE").font=Font(bold=True); ws.cell(r,4).alignment=Alignment("right")
            ws.cell(r,14,"="+"+".join(imps)).font=Font(bold=True); ws.cell(r,14).number_format='#,##0.00 €'; ws.cell(r,14).fill=g_fill; ws.cell(r,4).border=border; ws.cell(r,14).border=border
        for i,w in enumerate([6,10,12,25,40,25,6,6,6,6,6,10,10,12],1): ws.column_dimensions[get_column_letter(i)].width=w
        wb.save(file_path)

# --- MAPPING DIALOG ---
class ColumnMappingDialog(tk.Toplevel):
    def __init__(self, parent, excel_cols):
        super().__init__(parent)
        self.title("Mappatura Colonne Excel")
        self.geometry("500x600")
        self.excel_cols = [""] + excel_cols
        self.result = None
        
        # Campi Standard XPWE (Nomi Interni)
        self.fields = {
            "Codice": ["Codice", "Tariffa", "Articolo", "Codice Voce", "Voce"],
            "Categoria": ["Categoria", "WBS", "Capitolo"],
            "Descrizione Sintetica": ["Descrizione Sintetica", "DesRidotta", "Titolo"],
            "Descrizione Estesa": ["Descrizione Estesa", "DesEstesa", "Descrizione", "Voce di Capitolato"],
            "U.M.": ["U.M.", "Unità", "Unità di Misura", "UM"],
            "Prezzo Unit.": ["Prezzo Unit.", "Prezzo Unitario", "Prezzo 1", "Elenco Prezzi"],
            "Quantità": ["Quantità", "Quantita", "Totale", "Qta"],
            "Descrizione Rigo": ["Descrizione Rigo", "Descrizione Misurazione", "Rigo"],
            "N.": ["N.", "Parti Uguali", "N"],
            "Lung.": ["Lung.", "Lunghezza"],
            "Larg.": ["Larg.", "Larghezza"],
            "Alt./Peso": ["Alt./Peso", "Alt.", "Altezza", "Peso"]
        }
        
        self.mappings = {}
        
        main_frame = tk.Frame(self, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(main_frame, text="Associa le colonne del tuo Excel ai campi PriMus:", 
                 font=("Segoe UI", 11, "bold")).pack(pady=(0, 15))
        
        # Grid per le select
        grid_frame = tk.Frame(main_frame)
        grid_frame.pack(fill=tk.BOTH, expand=True)
        
        for i, (field, synonyms) in enumerate(self.fields.items()):
            tk.Label(grid_frame, text=field + ":", font=("Segoe UI", 10)).grid(row=i, column=0, sticky="w", pady=5)
            
            var = tk.StringVar()
            cb = ttk.Combobox(grid_frame, textvariable=var, values=self.excel_cols, width=30, state="readonly")
            cb.grid(row=i, column=1, padx=10, pady=5)
            
            # Heuristic auto-detection
            detected = ""
            for s in synonyms:
                for ecol in excel_cols:
                    if s.lower() in ecol.lower():
                        detected = ecol
                        break
                if detected: break
            
            if detected: var.set(detected)
            self.mappings[field] = var

        btn_frame = tk.Frame(main_frame, pady=20)
        btn_frame.pack(fill=tk.X)
        
        tk.Button(btn_frame, text="✅ Conferma Mappatura", bg="#4CAF50", fg="white", 
                  font=("Segoe UI", 10, "bold"), command=self.on_confirm, height=2).pack(fill=tk.X)
        
        self.transient(parent)
        self.grab_set()
        self.wait_window()

    def on_confirm(self):
        self.result = {k: v.get() for k, v in self.mappings.items() if v.get()}
        self.destroy()

# --- APP ---
class ConverterTab:
    def __init__(self, parent, title, mode, app):
        self.mode = mode # "XPWE->XLSX" or "XLSX->XPWE"
        self.app = app
        self.raw_df = None
        self.mapping = {}
        
        # Theme colors based on mode
        self.theme_color = "#217346" if mode == "XPWE->XLSX" else "#005a9e" # Excel Green vs PriMus Blue
        self.btn_color = "#e8f5e9" if mode == "XPWE->XLSX" else "#e1f5fe"
        
        self.frame = tk.Frame(parent, bg="#ffffff")
        self.frame.pack(fill=tk.BOTH, expand=True)

        # Visual Header
        header_bar = tk.Frame(self.frame, bg=self.theme_color, height=60)
        header_bar.pack(fill=tk.X)
        header_bar.pack_propagate(False)
        
        header_title = "ESTRAZIONE DATI: XPWE -> EXCEL" if mode == "XPWE->XLSX" else "CONVERSIONE DATI: EXCEL -> XPWE"
        tk.Label(header_bar, text=header_title, bg=self.theme_color, fg="#ffffff", 
                 font=("Segoe UI", 16, "bold")).pack(pady=12)

        # Controls
        ctrl_frame = tk.Frame(self.frame, bg="#f8f9fa", pady=15, padx=20)
        ctrl_frame.pack(fill=tk.X)
        
        btn_text_load = "📂 SELEZIONA FILE SORGENTE"
        tk.Button(ctrl_frame, text=btn_text_load, command=self.load_file, 
                  font=("Segoe UI", 10), width=30, height=2).pack(side=tk.LEFT)
        
        self.lbl_status = tk.Label(ctrl_frame, text="Pronto all'uso", bg="#f8f9fa", 
                                  font=("Segoe UI", 10, "italic"), fg="#666666")
        self.lbl_status.pack(side=tk.LEFT, padx=20)
        
        btn_text_export = "🚀 GENERA EXCEL" if mode == "XPWE->XLSX" else "🚀 GENERA XPWE"
        self.btn_export = tk.Button(ctrl_frame, text=btn_text_export, command=self.export_file, 
                                   bg=self.theme_color, fg="#ffffff", font=("Segoe UI", 11, "bold"), 
                                   width=25, height=2, activebackground="#333333", activeforeground="#ffffff")
        self.btn_export.pack(side=tk.RIGHT)

        # Treeview
        tree_frame = tk.Frame(self.frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # COLONNE ALLINEATE ALL'OUTPUT EXCEL
        self.cols = ["N. Ord.", "Categoria", "Codice", "Descrizione Sintetica", "Descrizione Estesa", "Descrizione Rigo", "U.M.", "N.", "Lung.", "Larg.", "Alt.", "Quantità", "Prezzo Unit.", "Importo Totale"]
        
        self.tree = ttk.Treeview(tree_frame, columns=self.cols, show="headings")
        
        # Tags per i colori
        self.tree.tag_configure("header", background="#4F81BD", foreground="white")  # Blu Intenso (Header Voce)
        self.tree.tag_configure("row", background="white", foreground="black")       # Righe Misure
        self.tree.tag_configure("subtotal", background="#F2F2F2", foreground="black", font=("Arial", 9, "bold")) # Grigio Chiaro (Totale Voce)
        self.tree.tag_configure("grandtotal", background="#C0C0C0", foreground="black", font=("Arial", 10, "bold")) # Grigio Scuro (Totale Generale)

        # Scrollbars
        vs = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        
        vs.pack(side=tk.RIGHT, fill=tk.Y)
        hs.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill=tk.BOTH, expand=True)

        # Configurazione Colonne
        widths = [60, 100, 80, 200, 150, 200, 40, 40, 50, 50, 50, 80, 80, 100]
        for c, w in zip(self.cols, widths):
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, stretch=False)

    def load_file(self):
        filetypes = [("XPWE Files", "*.xpwe")] if self.mode == "XPWE->XLSX" else [("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")]
        p = filedialog.askopenfilename(filetypes=filetypes)
        if not p: return
        
        try:
            self.lbl_status.config(text="Lettura in corso...")
            self.frame.update()
            
            ext = p.split('.')[-1].lower()
            if ext == 'xpwe': 
                raw = XPWEHandler.read_xpwe(p)
            else:
                # Caricamento parziale per rilevare le colonne
                raw_preview = pd.read_excel(p, nrows=5) if ext == 'xlsx' else pd.read_csv(p, nrows=5, sep=None, engine='python')
                cols = raw_preview.columns.tolist()
                
                # Mostra Finestra Mappatura
                diag = ColumnMappingDialog(self.app.root, cols)
                if not diag.result:
                    self.lbl_status.config(text="Caricamento annullato")
                    return
                
                self.mapping = diag.result
                # Caricamento completo
                if ext == 'xlsx': raw = pd.read_excel(p, dtype=str)
                else: raw = pd.read_csv(p, dtype=str, sep=None, engine='python')
                
                # Pulisce i nomi delle colonne prim'ancora di rinominare
                raw.columns = [str(c).strip() for c in raw.columns]
                
                # Applica Mappatura (Rinomina colonne Excel -> Nomi Interni)
                rename_map = {v: k for k, v in self.mapping.items()}
                raw.rename(columns=rename_map, inplace=True)
            
            self.raw_df = self.normalize_df(raw)
            self.populate_tree()
            self.lbl_status.config(text=f"Caricato: {os.path.basename(p)} ({len(self.raw_df)} righe)")
        except Exception as e:
            messagebox.showerror("Errore", str(e))
            self.lbl_status.config(text="Errore nel caricamento")

    def normalize_df(self, df):
        # PROPAGAZIONE DATI (Fill Forward)
        # Fondamentale per leggere l'Excel che ha la struttura "Intestazione" -> "Misure"
        std_cols = ["N. Ord.", "Categoria", "Codice", "Descrizione Sintetica", "Descrizione Estesa", "U.M.", "Prezzo Unit.", "Quantità"]
        present_cols = [c for c in std_cols if c in df.columns]
        if present_cols:
            df[present_cols] = df[present_cols].ffill()

        # Garantisce esistenza colonne minime (interni standard)
        for c in ["Categoria", "Codice", "N.", "Lung.", "Larg.", "Alt./Peso", "Quantità", "Descrizione Rigo", "Descrizione Estesa", "Descrizione Sintetica", "U.M.", "Prezzo Unit."]: 
            if c not in df.columns: df[c] = ""
        
        # Filtro righe vuote o di totale
        mask = pd.Series(True, index=df.index)
        for c in ["Descrizione Sintetica", "Descrizione Rigo"]:
            if c in df.columns: 
                mask &= ~df[c].astype(str).str.contains("TOTALE VOCE|TOTALE GENERALE", case=False, na=False)
        
        # Rimuove righe dove Codice è vuoto (dopo ffill dovrebbero esserci quasi sempre)
        if "Codice" in df.columns:
            mask &= (df["Codice"].astype(str).str.strip() != "")
        
        return df[mask].copy()

    def populate_tree(self):
        # Clear
        for i in self.tree.get_children(): self.tree.delete(i)
        if self.raw_df is None: return

        df = self.raw_df.copy(); df.fillna("", inplace=True)
        
        # Raggruppamento per replicare la struttura Excel
        grouped = df.groupby(['Categoria', 'Codice'], sort=False)
        
        vc = 1 # Contatore voci
        grand_total = 0.0
        
        for (cat, cod), grp in grouped:
            f = grp.iloc[0] # Dati della "Testata" voce
            
            # --- 1. RIGA INTESTATURA (Blu) ---
            head_vals = [
                str(vc), cat, cod, 
                f.get('Descrizione Sintetica',''), f.get('Descrizione Estesa',''), 
                "", "", "", "", "", "", "", "", ""
            ]
            self.tree.insert("", "end", values=head_vals, tags=("header",))
            
            # Calcolo Subtotali
            tot_qty_voce = 0.0
            
            # --- 2. RIGHE MISURE (Bianco) ---
            for _, row in grp.iterrows():
                n = safe_eval_math(row.get('N.',''))
                l = safe_eval_math(row.get('Lung.',''))
                w = safe_eval_math(row.get('Larg.',''))
                h = safe_eval_math(row.get('Alt./Peso',''))
                
                # Calcolo Quantità Riga
                items = [n,l,w,h]
                qty_val = 1.0
                has_dims = False
                for x in items:
                    if x != 0: 
                        qty_val *= x
                        has_dims = True
                
                if not has_dims: qty_val = 0.0
                
                # Override se c'è quantità esplicita (XPWE)
                q_txt = str(row.get('Quantità', '')).strip()
                if q_txt and q_txt not in ['','None','nan']:
                     # Se c'è un valore esplicito usiamo quello se non abbiamo dimensioni
                     # Oppure se le dimensioni darebbero 0?
                     # Per coerenza con XPWE, se ci sono dim usiamo dim, se no qta.
                     calc_chk = n*l*w*h if all([n,l,w,h]) else 0 # check semplice
                     if not has_dims:
                        res = safe_eval_math(q_txt)
                        if res != 0: qty_val = res

                tot_qty_voce += qty_val
                
                meas_vals = [
                    "", "", "", "", "", # Primi campi vuoti
                    row.get('Descrizione Rigo','') or row.get('Descrizione Sintetica',''),
                    "", # UM
                    clean_val_for_excel(row.get('N.','')),
                    clean_val_for_excel(row.get('Lung.','')),
                    clean_val_for_excel(row.get('Larg.','')),
                    clean_val_for_excel(row.get('Alt./Peso','')),
                    fmt_qty(qty_val),
                    "", "" # Prezzo e Tot non sulla riga misure
                ]
                self.tree.insert("", "end", values=meas_vals, tags=("row",))

            # --- 3. RIGA TOTALE VOCE (Grigio Chiaro) ---
            price = safe_eval_math(f.get('Prezzo Unit.', 0))
            tot_price_voce = tot_qty_voce * price
            grand_total += tot_price_voce
            
            subtot_vals = [
                "", "", "", "", "", 
                "TOTALE VOCE", 
                f.get('U.M.', ''), 
                "", "", "", "", 
                fmt_qty(tot_qty_voce), 
                fmt_money(price), 
                fmt_money(tot_price_voce)
            ]
            self.tree.insert("", "end", values=subtot_vals, tags=("subtotal",))
            
            vc += 1

        # --- 4. TOTALE GENERALE (Grigio Scuro) ---
        grand_vals = [
            "", "", "", 
            "TOTALE GENERALE", 
            "", "", "", "", "", "", "", "", "", 
            fmt_money(grand_total)
        ]
        self.tree.insert("", "end", values=grand_vals, tags=("grandtotal",))

    def export_file(self):
        if self.raw_df is None: return
        def_ext = ".xlsx" if self.mode == "XPWE->XLSX" else ".xpwe"
        p = filedialog.asksaveasfilename(defaultextension=def_ext)
        if not p: return
        
        try:
            self.lbl_status.config(text="Esportazione...")
            self.frame.update()
            
            if self.mode == "XPWE->XLSX":
                ExcelHandler.write_excel_pro(self.raw_df, p)
            else:
                XPWEHandler.write_xpwe(self.raw_df, p)
                
            self.lbl_status.config(text="Esportazione completata!")
            messagebox.showinfo("Successo", f"File salvato in:\n{p}")
        except Exception as e:
            messagebox.showerror("Errore Export", str(e))
            self.lbl_status.config(text="Errore export")

class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Convertitore Computi 18.0 (Tabs + Grid)")
        self.root.geometry("1400x800")
        
        # Tabs
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Tab 1: XPWE -> XLSX
        self.tab1 = ConverterTab(self.notebook, "📗 XPWE -> EXCEL", "XPWE->XLSX", self)
        self.notebook.add(self.tab1.frame, text="📗 XPWE -> EXCEL")
        
        # Tab 2: XLSX -> XPWE
        self.tab2 = ConverterTab(self.notebook, "📘 EXCEL -> XPWE", "XLSX->XPWE", self)
        self.notebook.add(self.tab2.frame, text="📘 EXCEL -> XPWE")
        
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_change)

    def on_tab_change(self, event):
        idx = self.notebook.index("current")
        # Sincronizza il colore del tab selezionato con la barra sottostante
        selected_color = "#217346" if idx == 0 else "#005a9e"
        style = ttk.Style()
        style.map("TNotebook.Tab", background=[("selected", selected_color), ("!selected", "#f8f9fa")])

if __name__ == "__main__":
    root = tk.Tk()
    
    # Modern Styling for Tabs
    style = ttk.Style()
    style.theme_use('clam') # Better style handling
    
    style.configure("Treeview", rowheight=28, font=("Segoe UI", 9))
    style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
    
    # Customize Notebook Tabs (Fixed Height & Symmetric)
    style.configure("TNotebook.Tab", font=("Segoe UI", 11, "bold"), padding=[20, 10], shiftrelief=0)
    style.map("TNotebook.Tab",
              background=[("selected", "#217346"), ("!selected", "#f8f9fa")],
              foreground=[("selected", "#ffffff"), ("!selected", "#888888")],
              padding=[("selected", [20, 10]), ("!selected", [20, 10])])
    
    app = ConverterApp(root)
    root.mainloop()