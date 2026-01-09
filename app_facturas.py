import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from tkinterdnd2 import DND_FILES, TkinterDnD
import pdfplumber
import re
import os
import platform
import struct
import win32clipboard
import webbrowser
import tempfile

# --- Clase ScrollableFrame (Sin cambios) ---
class ScrollableFrame(tk.Frame):
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # Bind scroll solo cuando el mouse está sobre el canvas
        self.canvas.bind("<Enter>", self._bind_mousewheel)
        self.canvas.bind("<Leave>", self._unbind_mousewheel)

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)
    
    def _bind_mousewheel(self, event):
        """Activa el scroll cuando el mouse entra al área"""
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel)
    
    def _unbind_mousewheel(self, event):
        """Desactiva el scroll cuando el mouse sale del área"""
        self.canvas.unbind_all("<MouseWheel>")
        self.canvas.unbind_all("<Button-4>")
        self.canvas.unbind_all("<Button-5>")

    def _on_mousewheel(self, event):
        if platform.system() == 'Windows':
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        elif platform.system() == 'Darwin':
            self.canvas.yview_scroll(int(-1*event.delta), "units")
        else:
            if event.num == 4:
                self.canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self.canvas.yview_scroll(1, "units")

# --- Clase Principal ---
class NativeInvoiceApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("Procesador de Facturas - PuntoBase")
        self.geometry("900x650")
        self.minsize(500, 400)  # Tamaño mínimo para garantizar usabilidad
        
        self.pdf_files = [] 
        self.parsed_data = []
        self.file_widgets = {}
        self.current_html = ""  # Para guardar el HTML generado
        self.grid_columns = 5   # Columnas para tarjetas de archivos

        # --- GUI SETUP con GRID para control total del espacio ---
        # Configurar grid principal: 3 filas (header, content, footer)
        self.grid_rowconfigure(0, weight=0)              # Header: tamaño fijo
        self.grid_rowconfigure(1, weight=1, minsize=150) # Content: expandible
        self.grid_rowconfigure(2, weight=0, minsize=90)  # Footer: GARANTIZADO
        self.grid_columnconfigure(0, weight=1)

        # === HEADER (row=0): Instrucciones y zona de drop ===
        header_frame = tk.Frame(self, bg="#f0f0f0", padx=15, pady=10)
        header_frame.grid(row=0, column=0, sticky="nsew")

        self.lbl_instruction = tk.Label(
            header_frame, 
            text="Arrastra tus Facturas (PDF) aquí", 
            font=("Segoe UI", 16, "bold"), bg="#f0f0f0"
        )
        self.lbl_instruction.pack(pady=(0, 8))

        self.drop_zone = tk.Label(
            header_frame,
            text="⬇️\nSuéltalos aquí\n(o haz clic para seleccionar)",
            relief="groove", borderwidth=3, width=50, height=4,
            fg="#555", bg="#ffffff", font=("Segoe UI", 11)
        )
        self.drop_zone.pack(fill=tk.X)
        
        self.drop_zone.drop_target_register(DND_FILES)
        self.drop_zone.dnd_bind('<<Drop>>', self.drop_files)
        self.drop_zone.dnd_bind('<<DragEnter>>', self.on_drag_enter)
        self.drop_zone.dnd_bind('<<DragLeave>>', self.on_drag_leave)
        self.drop_zone.bind("<Button-1>", self.open_file_dialog)

        # Label con contador de documentos
        self.lbl_files = tk.Label(header_frame, text="Documentos en cola: 0", 
                                  font=("Segoe UI", 10, "bold"), bg="#f0f0f0", anchor="w")
        self.lbl_files.pack(fill=tk.X, pady=(8, 0))

        # === CONTENT (row=1): Archivos y vista previa ===
        content_frame = tk.Frame(self, bg="#f0f0f0", padx=15)
        content_frame.grid(row=1, column=0, sticky="nsew")
        
        # Dividir content en dos: archivos (arriba) y texto (abajo)
        content_frame.grid_rowconfigure(0, weight=1, minsize=60)  # Archivos
        content_frame.grid_rowconfigure(1, weight=2, minsize=80)  # Texto
        content_frame.grid_columnconfigure(0, weight=1)

        # Área de archivos cargados
        self.files_container = ScrollableFrame(content_frame)
        self.files_container.grid(row=0, column=0, sticky="nsew", pady=(0, 5))
        
        # Vista previa del correo
        self.text_area = ScrolledText(content_frame, wrap=tk.WORD, font=("Arial", 10))
        self.text_area.grid(row=1, column=0, sticky="nsew", pady=(5, 0))
        self.text_area.insert(tk.END, "Cargue archivos PDF y genere el correo...")
        self.text_area.config(state=tk.DISABLED)

        # === FOOTER (row=2): Botones SIEMPRE VISIBLES ===
        footer_frame = tk.Frame(self, bg="#f0f0f0", padx=15, pady=10, height=90)
        footer_frame.grid(row=2, column=0, sticky="nsew")
        footer_frame.grid_propagate(False)  # CLAVE: No permite que se encoja
        
        # Configurar grid interno del footer
        footer_frame.grid_rowconfigure(0, weight=1)
        footer_frame.grid_rowconfigure(1, weight=1)
        footer_frame.grid_columnconfigure(0, weight=1)

        # Botón principal - ocupa todo el ancho
        self.btn_generate = tk.Button(
            footer_frame, text="GENERAR CORREO", command=self.generate_email,
            font=("Segoe UI", 11, "bold"), bg="#007aff", fg="white", 
            height=1, borderwidth=0, cursor="hand2"
        )
        self.btn_generate.grid(row=0, column=0, sticky="nsew", pady=(0, 5))

        # Frame para botones secundarios
        secondary_frame = tk.Frame(footer_frame, bg="#f0f0f0")
        secondary_frame.grid(row=1, column=0, sticky="nsew")
        secondary_frame.grid_columnconfigure(0, weight=1)
        secondary_frame.grid_columnconfigure(1, weight=1)
        secondary_frame.grid_columnconfigure(2, weight=1)
        secondary_frame.grid_columnconfigure(3, weight=1)

        self.btn_preview = ttk.Button(secondary_frame, text="👁️ Vista HTML", command=self.preview_html_in_browser)
        self.btn_preview.grid(row=0, column=0, sticky="nsew", padx=(0, 3))

        self.btn_copy = ttk.Button(secondary_frame, text="Copiar Texto", command=self.copy_to_clipboard)
        self.btn_copy.grid(row=0, column=1, sticky="nsew", padx=(3, 3))

        self.btn_copy_pdfs = ttk.Button(secondary_frame, text="Copiar PDFs", command=self.copy_pdfs_to_clipboard)
        self.btn_copy_pdfs.grid(row=0, column=2, sticky="nsew", padx=(3, 3))

        self.btn_clear = ttk.Button(secondary_frame, text="Limpiar Todo", command=self.clear_all)
        self.btn_clear.grid(row=0, column=3, sticky="nsew", padx=(3, 0))

        # Vincular evento de redimensionamiento
        self.bind("<Configure>", self._on_window_resize)

    def _on_window_resize(self, event):
        """Ajusta elementos dinámicamente según el tamaño de ventana"""
        # Solo procesar eventos de la ventana principal
        if event.widget != self:
            return
            
        width = self.winfo_width()
        
        # Ajustar texto de botones según ancho disponible
        if width < 550:
            self.btn_generate.config(text="GENERAR", font=("Segoe UI", 10, "bold"))
            self.btn_preview.config(text="👁️")
            self.btn_copy.config(text="Texto")
            self.btn_copy_pdfs.config(text="PDFs")
            self.btn_clear.config(text="Limpiar")
        elif width < 750:
            self.btn_generate.config(text="GENERAR CORREO", font=("Segoe UI", 10, "bold"))
            self.btn_preview.config(text="👁️ HTML")
            self.btn_copy.config(text="Copiar")
            self.btn_copy_pdfs.config(text="PDFs")
            self.btn_clear.config(text="Limpiar")
        else:
            self.btn_generate.config(text="GENERAR CORREO", font=("Segoe UI", 11, "bold"))
            self.btn_preview.config(text="👁️ Vista HTML")
            self.btn_copy.config(text="Copiar Texto")
            self.btn_copy_pdfs.config(text="Copiar PDFs")
            self.btn_clear.config(text="Limpiar Todo")
        
        # Ajustar número de columnas del grid de archivos
        if width < 500:
            new_cols = 3
        elif width < 700:
            new_cols = 5
        elif width < 900:
            new_cols = 7
        elif width < 1200:
            new_cols = 9
        else:
            new_cols = 11
            
        if new_cols != self.grid_columns:
            self.grid_columns = new_cols
            if self.pdf_files:
                self.refresh_grid()

    # --- LOGICA VISUAL ---
    def add_file_card(self, file_path):
        if file_path in self.file_widgets: return
        filename = os.path.basename(file_path)
        index = len(self.pdf_files)
        # Usar columnas dinámicas según ancho de ventana
        row = index // self.grid_columns
        col = index % self.grid_columns
        
        # Configurar la columna para que se expanda
        self.files_container.scrollable_frame.grid_columnconfigure(col, weight=1, uniform="cards")
        
        card = tk.Frame(
            self.files_container.scrollable_frame, 
            relief="raised", borderwidth=1, bg="white",
            height=115
        )
        card.grid(row=row, column=col, padx=4, pady=4, sticky="nsew")

        btn_del = tk.Label(card, text="×", fg="#999", bg="white", font=("Arial", 14), cursor="hand2")
        btn_del.place(relx=0.95, rely=0.0, anchor="ne")
        btn_del.bind("<Button-1>", lambda e, p=file_path: self.remove_file(p))
        btn_del.bind("<Enter>", lambda e: e.widget.config(fg="red"))
        btn_del.bind("<Leave>", lambda e: e.widget.config(fg="#999"))

        tk.Label(card, text="📄", font=("Arial", 28), bg="white").pack(pady=(12, 3))
        display_name = filename if len(filename) < 20 else filename[:17] + "..."
        tk.Label(card, text=display_name, font=("Segoe UI", 8), bg="white", wraplength=120).pack()
        self.pdf_files.append(file_path)
        self.file_widgets[file_path] = card
        self._update_file_count()

    def remove_file(self, file_path):
        if file_path in self.file_widgets:
            self.file_widgets[file_path].destroy()
            del self.file_widgets[file_path]
        if file_path in self.pdf_files:
            self.pdf_files.remove(file_path)
        self.refresh_grid()
        self._update_file_count()

    def refresh_grid(self):
        current_files = list(self.pdf_files)
        # Limpiar configuración de columnas anterior
        for i in range(20):  # Limpiar hasta 20 columnas posibles
            self.files_container.scrollable_frame.grid_columnconfigure(i, weight=0, uniform="")
        for widget in self.files_container.scrollable_frame.winfo_children():
            widget.destroy()
        self.pdf_files = []
        self.file_widgets = {}
        for f in current_files:
            self.add_file_card(f)

    # --- LOGICA INTERACCION (Sin cambios) ---
    def on_drag_enter(self, event):
        self.drop_zone.config(bg="#e1f5fe", text="⬇️\n¡SUELTA AHORA!")

    def on_drag_leave(self, event):
        self.drop_zone.config(bg="#ffffff", text="⬇️\nSuéltalos aquí\n(o haz clic para seleccionar)")

    def open_file_dialog(self, event=None):
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if files:
            for f in files:
                self.add_file_card(f)

    def drop_files(self, event):
        self.on_drag_leave(None)
        if event.data:
            raw_files = self.tk.splitlist(event.data)
            for file_path in raw_files:
                if file_path.lower().endswith('.pdf'):
                    self.add_file_card(file_path)

    def clear_all(self):
        self.pdf_files = []
        self.parsed_data = []
        self.file_widgets = {}
        for widget in self.files_container.scrollable_frame.winfo_children():
            widget.destroy()
        self.text_area.config(state=tk.NORMAL)
        self.text_area.delete(1.0, tk.END)
        self.text_area.insert(tk.END, "Cargue archivos PDF y genere el correo...")
        self.text_area.config(state=tk.DISABLED)
        self.current_html = ""
        self._update_file_count()

    def _update_file_count(self):
        """Actualiza el label con el número de documentos en cola"""
        count = len(self.pdf_files)
        self.lbl_files.config(text=f"Documentos en cola: {count}")

    # --- EXTRACCION PDF (Mejorado con múltiples estrategias) ---
    def extract_pdf_data(self, pdf_path):
        try:
            with pdfplumber.open(pdf_path) as pdf:
                page = pdf.pages[0]
                text = page.extract_text()
                
                # Si no hay texto, intentar OCR o informar error útil
                if not text or len(text.strip()) < 50:
                    filename = os.path.basename(pdf_path)
                    print(f"⚠️  PDF '{filename}' no contiene texto extraíble (puede ser imagen escaneada)")
                    return {
                        "emisor_nombre": f"[PDF sin texto: {filename}]",
                        "emisor_rut": "S/I",
                        "deudor_nombre": "S/I",
                        "deudor_rut": "S/I",
                        "folio": "S/I",
                        "monto": "0",
                        "fecha_emision": "S/I",
                        "valor_bruto": "0",
                    }

                # Funciones auxiliares
                def safe_search(patterns, text, default="S/I"):
                    """Intenta múltiples patrones hasta encontrar match"""
                    if not isinstance(patterns, list):
                        patterns = [patterns]
                    for pattern in patterns:
                        match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
                        if match:
                            result = match.group(1).strip() if match.lastindex >= 1 else match.group(0).strip()
                            # Limpiar saltos de línea múltiples
                            result = re.sub(r'\s+', ' ', result)
                            return result
                    return default

                def norm_rut(r):
                    """Normalizar espacios en RUTs"""
                    return re.sub(r'\s+', '', r)

                def format_fecha(raw):
                    """Convertir fecha a DD/MM/YYYY"""
                    raw = raw.strip()
                    if raw == "S/I" or not raw:
                        return "S/I"
                    
                    # Formato YYYY-MM-DD (ISO)
                    m = re.search(r'(\d{4})-(\d{1,2})-(\d{1,2})', raw)
                    if m:
                        y, mo, d = m.groups()
                        return f"{int(d):02d}/{int(mo):02d}/{int(y):04d}"
                    
                    # Formato DD/MM/YYYY or DD-MM-YYYY
                    m = re.search(r"(\d{1,2})\s*[/-]\s*(\d{1,2})\s*[/-]\s*(\d{2,4})", raw)
                    if m:
                        d, mo, y = m.groups()
                        y = y if len(y) == 4 else ("20" + y)
                        return f"{int(d):02d}/{int(mo):02d}/{int(y):04d}"
                    
                    # Textual español: '24 de Diciembre del 2025'
                    m2 = re.search(r"(\d{1,2})\s+de\s+([A-Za-záéíóúÁÉÍÓÚñÑ]+)\s+(?:del|de)\s+(\d{4})", raw, re.IGNORECASE)
                    if m2:
                        d, month_name, y = m2.groups()
                        months = {
                            'enero':1,'febrero':2,'marzo':3,'abril':4,'mayo':5,'junio':6,
                            'julio':7,'agosto':8,'septiembre':9,'setiembre':9,'octubre':10,'noviembre':11,'diciembre':12
                        }
                        mnum = months.get(month_name.lower())
                        if mnum:
                            return f"{int(d):02d}/{mnum:02d}/{int(y):04d}"
                    
                    return raw

                # Extraer todos los RUTs del documento
                rut_pattern = r"(?:R\.U\.T\.?:?\s*)?(\d{1,3}(?:\.\d{3}){1,2}-\s*[\dkK])"
                all_ruts = re.findall(rut_pattern, text, re.IGNORECASE)
                all_ruts = [norm_rut(r) for r in all_ruts if r]

                # ESTRATEGIA 1: Detectar emisor RUT (primero en el documento)
                emisor_rut = all_ruts[0] if all_ruts else "S/I"
                
                # ESTRATEGIA 2: Detectar deudor RUT (después de SEÑOR(ES) o segundo RUT)
                deudor_rut = "S/I"
                if len(all_ruts) >= 2:
                    señor_idx = text.upper().find('SEÑOR')
                    if señor_idx != -1:
                        for m in re.finditer(rut_pattern, text, re.IGNORECASE):
                            if m.start() > señor_idx:
                                deudor_rut = norm_rut(m.group(1))
                                break
                    if deudor_rut == "S/I":
                        deudor_rut = all_ruts[1]

                # Normalizar números: convertir comas a puntos para unificar formato
                def normalize_amount(amt_str):
                    """Convierte '32,567,147' o '32.567.147' a '32.567.147'"""
                    # Si tiene comas, asumir que son separadores de miles (formato anglosajón)
                    if ',' in amt_str:
                        return amt_str.replace(',', '.')
                    return amt_str

                # Extracción de campos con múltiples estrategias
                data = {
                    # Emisor nombre: múltiples patrones
                    "emisor_nombre": safe_search([
                        r"^([A-ZÁÉÍÓÚÑ][A-Z\sÁÉÍÓÚÑ\.\-]+?SPA)\s*\n",  # Buscar línea 1 que termine en SPA
                        r"(?:FACTURA\s+ELECTR[OÓ]NICA\s*\n.*?\n)([A-ZÁÉÍÓÚÑ][A-Z\sÁÉÍÓÚÑ\.\-]+?)\s*\n",  # Después de FACTURA ELECTRONICA
                        r"^([A-ZÁÉÍÓÚÑ][A-Z\sÁÉÍÓÚÑ\.\-]+?)\s*\n.*?Giro:",  # Línea antes de Giro:
                    ], text, default="EMISOR DESCONOCIDO"),
                    
                    "emisor_rut": emisor_rut,
                    
                    # Deudor nombre: múltiples patrones
                    "deudor_nombre": safe_search([
                        r"Señor\(es\):[^\n]*\n([A-Z][A-ZÁÉÍÓÚÑ\s\.\-]+?SOCIEDAD\s+ANONIMA)",  # Formato Factura1550: nombre en línea siguiente
                        r"Señor\(es\)([A-Z][A-ZÁÉÍÓÚÑ\s\.\-]+?)(?:Direcci[oó]n|RUT\s|R\.U\.T\.?:|\n.*?RUT\s)",  # Sin espacio después de ()
                        r"SEÑOR\(ES\)[:\s]*([A-Z][A-ZÁÉÍÓÚÑ\s\.\-]+?)(?:\s+R\.U\.T\.?:|Direcci[oó]n:|\n.*?R\.U\.T\.?:)",
                        r"Señor\(es\)[:\s]+([A-Z][A-ZÁÉÍÓÚÑ\s\.\-]+?)(?:\s+Giro\s*:|R\.U\.T\.?:|Direcci[oó]n:|\n.*?Giro\s*:)",
                    ], text, default="S/I"),
                    
                    "deudor_rut": deudor_rut,
                    
                    # Folio: múltiples formatos
                    "folio": safe_search([
                        r"(?:FACTURA\s+ELECTR[OÓ]NICA|ELECTRONICA)\s*\n\s*N[°º]?\s*(\d+)",  # Folio en línea separada
                        r"(?:N[°º]|Nº)\s*(\d+)",
                        r"Folio[:\s]*(\d+)",
                    ], text, default="0"),
                    
                    # Monto total: múltiples variantes (soporta comas y puntos)
                    "monto": None,  # Lo calculamos después
                    
                    # Fecha emisión
                    "fecha_emision": None,
                    
                    # Valor bruto (mismo que monto)
                    "valor_bruto": None,  # Lo calculamos después
                }
                
                # Extraer monto con soporte para comas y puntos
                raw_monto = safe_search([
                    r"Total\s+Final[\s\$]*:?\s*\$?\s*([\d\.,]+)",
                    r"TOTAL\s+FINAL[\s\$]*:?\s*\$?\s*([\d\.,]+)",
                    r"TOTAL[\s\$]*:?\s*\$?\s*([\d\.,]+)",
                    r"Total[\s\$]*:?\s*\$?\s*([\d\.,]+)",
                ], text, default="0")
                data['monto'] = normalize_amount(raw_monto)
                data['valor_bruto'] = data['monto']
                
                # Formatear fecha de emisión
                raw_fecha = safe_search([
                    r"Fecha\s+(?:de\s+)?Emisi[oó]n[:\s]*([^\n]+)",
                    r"Fecha:[^\n]*\n(\d{1,2}/\d{1,2}/\d{4})",  # Fecha en línea siguiente (Factura1550)
                    r"Fecha[:\s]+(\d{1,2}[/-]\d{1,2}[/-]\d{4})",  # Fecha en misma línea
                    r"Fecha[:\s]*(\d{4}-\d{1,2}-\d{1,2})",
                ], text, default="S/I")
                
                data['fecha_emision'] = format_fecha(raw_fecha)
                
                return data

        except Exception as e:
            print(f"❌ Error parsing {os.path.basename(pdf_path)}: {type(e).__name__}: {e}")
            import traceback
            traceback.print_exc()
            return None

    # --- GENERACION DEL CORREO (CAMBIO PRINCIPAL AQUÍ) ---
    def generate_email(self):
        if not self.pdf_files:
            messagebox.showwarning("Alerta", "Carga al menos un archivo PDF primero.")
            return

        self.parsed_data = []
        errors = 0
        
        for file_path in self.pdf_files:
            data = self.extract_pdf_data(file_path)
            if data:
                self.parsed_data.append(data)
            else:
                errors += 1

        if not self.parsed_data:
            messagebox.showerror("Error", "No se pudieron leer datos de los PDFs.")
            return

        # Tomamos los datos del deudor del primer PDF cargado para el encabezado
        header_data = self.parsed_data[0]
        deudor_full = f"{header_data['deudor_nombre']} {header_data['deudor_rut']}".upper()

        # --- FORMATO DE CORREO HTML PARA GMAIL ---
        # Genera HTML inline completo listo para copiar/pegar en Gmail
        
        # Nuevo orden de columnas solicitado:
        # Fecha Emisión, Rut Emisor, Nombre Emisor, N° Factura, Rut Deudor, Nombre Deudor, Valor Bruto Factura
        rows_html = ""
        for item in self.parsed_data:
            rows_html += f"""    <tr>
        <td style="border: 1px solid #cccccc; padding: 4px 6px; font-size: 10pt;">{item.get('fecha_emision','S/I')}</td>
        <td style="border: 1px solid #cccccc; padding: 4px 6px; font-size: 10pt;">{item.get('emisor_rut','S/I')}</td>
        <td style="border: 1px solid #cccccc; padding: 4px 6px; font-size: 10pt;">{item.get('emisor_nombre','S/I')}</td>
        <td style="border: 1px solid #cccccc; padding: 4px 6px; font-size: 10pt; text-align: center;">{item.get('folio','')}</td>
        <td style="border: 1px solid #cccccc; padding: 4px 6px; font-size: 10pt;">{item.get('deudor_rut','S/I')}</td>
        <td style="border: 1px solid #cccccc; padding: 4px 6px; font-size: 10pt;">{item.get('deudor_nombre','S/I')}</td>
        <td style="border: 1px solid #cccccc; padding: 4px 6px; font-size: 10pt; text-align: right;">{item.get('valor_bruto', item.get('monto','0'))}</td>
    </tr>
"""
        
        email_body = f"""<div style="font-family: Arial, sans-serif; font-size: 11pt; line-height: 1.5;">
    <p>Estimado:</p>
    
    <p>Junto con saludar, agradeceré a<br>
    <strong>{deudor_full}</strong><br>
    usted que pueda confirmar por este medio, la recepción y conformidad de las siguientes facturas electrónicas adjuntas, emitida por nuestro(s) cliente(s), las cuales están siendo cedidas a nuestro Factoring Punto Base Financiero Spa.</p>
    
    <div style="background-color: #e8e8e8; padding: 8px 12px; margin: 12px 0; font-weight: bold; font-size: 11pt;">
        {deudor_full}
    </div>
    
    <table style="border-collapse: collapse; margin: 12px 0;">
        <thead>
            <tr style="background-color: #4a6fa5;">
                <th style="border: 1px solid #cccccc; padding: 5px 8px; text-align: left; color: #ffffff; font-size: 10pt; font-weight: 600;">Fecha Emisión</th>
                <th style="border: 1px solid #cccccc; padding: 5px 8px; text-align: left; color: #ffffff; font-size: 10pt; font-weight: 600;">Rut Emisor</th>
                <th style="border: 1px solid #cccccc; padding: 5px 8px; text-align: left; color: #ffffff; font-size: 10pt; font-weight: 600;">Nombre Emisor</th>
                <th style="border: 1px solid #cccccc; padding: 5px 8px; text-align: center; color: #ffffff; font-size: 10pt; font-weight: 600;">N° Factura</th>
                <th style="border: 1px solid #cccccc; padding: 5px 8px; text-align: left; color: #ffffff; font-size: 10pt; font-weight: 600;">Rut Deudor</th>
                <th style="border: 1px solid #cccccc; padding: 5px 8px; text-align: left; color: #ffffff; font-size: 10pt; font-weight: 600;">Nombre Deudor</th>
                <th style="border: 1px solid #cccccc; padding: 5px 8px; text-align: right; color: #ffffff; font-size: 10pt; font-weight: 600;">Valor Bruto</th>
            </tr>
        </thead>
        <tbody>
{rows_html}        </tbody>
    </table>
    
    <p style="font-size: 10pt;">Favor ayudarnos con la siguiente información:<br>
    -Si mercaderías y/o productos se encuentran recibidos conformes sin observaciones?<br>
    -Si las facturas se encuentran recibidas?<br>
    -Existen notas de crédito o algún otro descuento que afecte el pago de estos documentos?<br>
    -Posible fecha de pago.</p>
</div>"""

        # Guardar HTML para copiar
        self.current_html = email_body
        # Mostrar vista previa como texto
        self.text_area.config(state=tk.NORMAL)
        self.text_area.delete(1.0, tk.END)
        self.text_area.insert(tk.END, self._html_to_preview_text())
        self.text_area.config(state=tk.DISABLED)
        
        if errors > 0:
            messagebox.showwarning("Atención", f"Se generó el correo, pero {errors} archivo(s) no pudieron ser leídos.")

    def copy_to_clipboard(self):
        # Usar el HTML guardado en lugar de obtenerlo del widget
        html = getattr(self, 'current_html', '')
        if len(html) < 5:
            messagebox.showwarning("Aviso", "No hay contenido para copiar. Genera el correo primero.")
            return
        
        # Copiar HTML al portapapeles de Windows para que Gmail lo reconozca
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            
            # Formato CF_HTML para que aplicaciones web lo reconozcan
            cf_html = self._generate_cf_html(html)
            win32clipboard.SetClipboardData(win32clipboard.RegisterClipboardFormat("HTML Format"), cf_html.encode('utf-8'))
            
            # También copiar como texto plano (fallback)
            win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, html)
            
            win32clipboard.CloseClipboard()
            messagebox.showinfo("¡Listo!", "Correo copiado al portapapeles (formato HTML).")
        except Exception as e:
            # Fallback a método estándar si falla
            self.clipboard_clear()
            self.clipboard_append(html)
            messagebox.showwarning("Copiado", f"Copiado como texto plano. {str(e)}")
    
    def _generate_cf_html(self, html):
        """Genera el formato CF_HTML requerido por Windows"""
        html_prefix = "<!DOCTYPE html><html><body>"
        html_suffix = "</body></html>"
        html_full = html_prefix + html + html_suffix
        
        # CF_HTML requiere encabezados específicos
        template = (
            "Version:0.9\r\n"
            "StartHTML:{start_html:09d}\r\n"
            "EndHTML:{end_html:09d}\r\n"
            "StartFragment:{start_frag:09d}\r\n"
            "EndFragment:{end_frag:09d}\r\n"
            "<html><body>\r\n"
            "<!--StartFragment-->{fragment}<!--EndFragment-->\r\n"
            "</body></html>"
        )
        
        fragment = html
        dummy = template.format(start_html=0, end_html=0, start_frag=0, end_frag=0, fragment=fragment)
        
        start_html = dummy.index("<html>")
        end_html = dummy.index("</html>") + 7
        start_frag = dummy.index("<!--StartFragment-->") + 20
        end_frag = dummy.index("<!--EndFragment-->")
        
        return template.format(
            start_html=start_html,
            end_html=end_html,
            start_frag=start_frag,
            end_frag=end_frag,
            fragment=fragment
        )

    def copy_pdfs_to_clipboard(self):
        """Copia los archivos PDF cargados al portapapeles de Windows (para pegar en Gmail)"""
        if not self.pdf_files:
            messagebox.showwarning("Aviso", "No hay archivos PDF cargados.")
            return
        
        try:
            # Formato CF_HDROP para copiar archivos al portapapeles
            # Estructura DROPFILES: https://docs.microsoft.com/en-us/windows/win32/api/shlobj_core/ns-shlobj_core-dropfiles
            
            # Crear lista de archivos con rutas absolutas terminadas en null
            files_str = "\0".join(os.path.abspath(f) for f in self.pdf_files) + "\0\0"
            files_bytes = files_str.encode('utf-16-le')
            
            # Estructura DROPFILES (20 bytes header + files)
            # pFiles: offset donde empiezan los archivos (20 bytes)
            # pt.x, pt.y: coordenadas (0, 0)
            # fNC: non-client area flag (False)
            # fWide: Unicode flag (True para UTF-16)
            dropfiles = struct.pack('IIIIi', 20, 0, 0, 0, 1) + files_bytes
            
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_HDROP, dropfiles)
            win32clipboard.CloseClipboard()
            
            count = len(self.pdf_files)
            messagebox.showinfo("¡Listo!", f"{count} archivo(s) PDF copiado(s) al portapapeles.\n\nAhora puedes pegarlos (Ctrl+V) en Gmail como adjuntos.")
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron copiar los archivos:\n{str(e)}")

    def preview_html_in_browser(self):
        """Abre el HTML generado en el navegador por defecto para vista previa renderizada"""
        html = getattr(self, 'current_html', '')
        if len(html) < 5:
            messagebox.showwarning("Aviso", "No hay contenido para previsualizar. Genera el correo primero.")
            return
        
        try:
            # Crear archivo temporal HTML
            with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
                # Escribir HTML completo con estilos
                full_html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Vista Previa - Correo de Facturas</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            padding: 20px;
            max-width: 900px;
            margin: 0 auto;
            background-color: #f5f5f5;
        }}
        .email-container {{
            background-color: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }}
    </style>
</head>
<body>
    <div class="email-container">
        {html}
    </div>
</body>
</html>"""
                f.write(full_html)
                temp_path = f.name
            
            # Abrir en navegador por defecto
            webbrowser.open('file://' + temp_path)
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir la vista previa:\n{str(e)}")

    def _html_to_preview_text(self):
        """Convierte el contenido a texto legible para la vista previa"""
        if not self.parsed_data:
            return ""

        header_data = self.parsed_data[0]
        deudor_full = f"{header_data['deudor_nombre']} {header_data['deudor_rut']}".upper()

        text = "Estimado:\n\n"
        text += f"Junto con saludar, agradeceré a\n{deudor_full}\nusted que pueda confirmar por este medio, la recepción y conformidad de las siguientes facturas electrónicas adjuntas, emitida por nuestro(s) cliente(s), las cuales están siendo cedidas a nuestro Factoring Punto Base Financiero Spa.\n\n"
        text += f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        text += f"  {deudor_full}\n"
        text += f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        text += "{:<20} | {:<15} | {:<30} | {:^10} | {:<15} | {:<30} | {:>18}\n".format(
            'Fecha Emisión', 'Rut Emisor', 'Nombre Emisor', 'N° Factura', 'Rut Deudor', 'Nombre Deudor', 'Valor Bruto Factura'
        )
        text += "{:-<20}-+-{:-<15}-+-{:-<30}-+-{:-<10}-+-{:-<15}-+-{:-<30}-+-{:-<18}\n".format('', '', '', '', '', '', '')

        for item in self.parsed_data:
            text += f"{item.get('fecha_emision','S/I'):<20} | {item.get('emisor_rut','S/I'):<15} | {item.get('emisor_nombre','S/I'):<30} | {item.get('folio',''):^10} | {item.get('deudor_rut','S/I'):<15} | {item.get('deudor_nombre','S/I'):<30} | {item.get('valor_bruto', item.get('monto','0')):>18}\n"

        text += "\n\nFavor ayudarnos con la siguiente información:\n"
        text += "-Si mercaderías y/o productos se encuentran recibidos conformes sin observaciones?\n"
        text += "-Si las facturas se encuentran recibidas?\n"
        text += "-Existen notas de crédito o algún otro descuento que afecte el pago de estos documentos?\n"
        text += "-Posible fecha de pago.\n"

        return text

if __name__ == "__main__":
    app = NativeInvoiceApp()
    app.mainloop()