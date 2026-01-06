import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from tkinterdnd2 import DND_FILES, TkinterDnD
import pdfplumber
import re
import os
import platform
import win32clipboard

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
        self.geometry("1100x850")
        
        self.pdf_files = [] 
        self.parsed_data = []
        self.file_widgets = {}
        self.current_html = ""  # Para guardar el HTML generado 

        # --- GUI SETUP (Sin cambios mayores) ---
        main_frame = tk.Frame(self, padx=20, pady=20, bg="#f0f0f0")
        main_frame.pack(fill=tk.BOTH, expand=True)

        lbl_instruction = tk.Label(
            main_frame, 
            text="Arrastra tus Facturas (PDF) aquí", 
            font=("Segoe UI", 16, "bold"), bg="#f0f0f0"
        )
        lbl_instruction.pack(pady=(0, 10))

        self.drop_zone = tk.Label(
            main_frame,
            text="⬇️\nSuéltalos aquí\n(o haz clic para seleccionar)",
            relief="groove", borderwidth=3, width=50, height=5,
            fg="#555", bg="#ffffff", font=("Segoe UI", 12)
        )
        self.drop_zone.pack(fill=tk.X, pady=10)
        
        self.drop_zone.drop_target_register(DND_FILES)
        self.drop_zone.dnd_bind('<<Drop>>', self.drop_files)
        self.drop_zone.dnd_bind('<<DragEnter>>', self.on_drag_enter)
        self.drop_zone.dnd_bind('<<DragLeave>>', self.on_drag_leave)
        self.drop_zone.bind("<Button-1>", self.open_file_dialog)

        lbl_files = tk.Label(main_frame, text="Documentos en cola:", font=("Segoe UI", 11, "bold"), bg="#f0f0f0", anchor="w")
        lbl_files.pack(fill=tk.X, pady=(10, 5))

        self.files_container = ScrollableFrame(main_frame)
        self.files_container.pack(fill=tk.BOTH, expand=True, pady=5)
        
        btn_frame = tk.Frame(main_frame, bg="#f0f0f0")
        btn_frame.pack(fill=tk.X, pady=15)

        self.btn_generate = tk.Button(
            btn_frame, text="GENERAR CORREO", command=self.generate_email,
            font=("Segoe UI", 12, "bold"), bg="#007aff", fg="white", 
            height=2, borderwidth=0, cursor="hand2"
        )
        self.btn_generate.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        self.btn_copy = ttk.Button(btn_frame, text="Copiar Texto", command=self.copy_to_clipboard)
        self.btn_copy.pack(side=tk.LEFT, padx=5)

        self.btn_clear = ttk.Button(btn_frame, text="Limpiar Todo", command=self.clear_all)
        self.btn_clear.pack(side=tk.RIGHT, padx=5)

        # Vista previa del correo
        self.text_area = ScrolledText(main_frame, wrap=tk.WORD, font=("Arial", 10))
        self.text_area.pack(fill=tk.BOTH, expand=True)
        self.text_area.insert(tk.END, "Cargue archivos PDF y genere el correo...")
        self.text_area.config(state=tk.DISABLED)

    # --- LOGICA VISUAL (Sin cambios) ---
    def add_file_card(self, file_path):
        if file_path in self.file_widgets: return
        filename = os.path.basename(file_path)
        index = len(self.pdf_files)
        row = index // 4
        col = index % 4
        card = tk.Frame(
            self.files_container.scrollable_frame, 
            relief="raised", borderwidth=1, bg="white",
            width=140, height=130
        )
        card.grid_propagate(False)
        card.grid(row=row, column=col, padx=8, pady=8)

        btn_del = tk.Label(card, text="×", fg="#999", bg="white", font=("Arial", 14), cursor="hand2")
        btn_del.place(relx=0.95, rely=0.0, anchor="ne")
        btn_del.bind("<Button-1>", lambda e, p=file_path: self.remove_file(p))
        btn_del.bind("<Enter>", lambda e: e.widget.config(fg="red"))
        btn_del.bind("<Leave>", lambda e: e.widget.config(fg="#999"))

        tk.Label(card, text="📄", font=("Arial", 32), bg="white").pack(pady=(15, 5))
        display_name = filename if len(filename) < 18 else filename[:15] + "..."
        tk.Label(card, text=display_name, font=("Segoe UI", 9), bg="white", wraplength=130).pack()
        self.pdf_files.append(file_path)
        self.file_widgets[file_path] = card

    def remove_file(self, file_path):
        if file_path in self.file_widgets:
            self.file_widgets[file_path].destroy()
            del self.file_widgets[file_path]
        if file_path in self.pdf_files:
            self.pdf_files.remove(file_path)
        self.refresh_grid()

    def refresh_grid(self):
        current_files = list(self.pdf_files)
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

    # --- EXTRACCION PDF (Con pequeñas mejoras) ---
    def extract_pdf_data(self, pdf_path):
        try:
            with pdfplumber.open(pdf_path) as pdf:
                page = pdf.pages[0]
                text = page.extract_text()
                
                if not text: raise ValueError("No text found")

                def safe_search(pattern, text, default="S/I"):
                    match = re.search(pattern, text, re.IGNORECASE)
                    return match.group(1).strip() if match else default

                # Normalizar espacios en RUTs
                def norm_rut(r):
                    return r.replace(" ", "")

                # Buscamos todos los RUTs presentes en el texto y aplicamos heurísticas
                rut_pattern = r"(\d{1,3}(?:\.\d{3})*-\s*[\dkK])"
                all_ruts = re.findall(rut_pattern, text, re.IGNORECASE)
                all_ruts = [norm_rut(r) for r in all_ruts]

                # heurística: emisor = primer RUT, deudor = RUT localizado cerca de 'SEÑOR(ES)'
                emisor_rut = all_ruts[0] if len(all_ruts) >= 1 else "S/I"
                deudor_rut = "S/I"
                if len(all_ruts) >= 2:
                    # intentamos seleccionar el RUT que aparece después de la etiqueta SEÑOR(ES)
                    señor_idx = text.upper().find('SEÑOR')
                    if señor_idx != -1:
                        # buscar primer match cuyo índice sea mayor que señor_idx
                        for m in re.finditer(rut_pattern, text, re.IGNORECASE):
                            if m.start() > señor_idx:
                                deudor_rut = norm_rut(m.group(1))
                                break
                        if deudor_rut == "S/I":
                            deudor_rut = all_ruts[1]
                    else:
                        deudor_rut = all_ruts[1]

                data = {
                    "emisor_nombre": safe_search(r"(.+?)\s*\n.*Giro:", text, default="EMISOR DESCONOCIDO"),
                    "emisor_rut": emisor_rut,
                    "deudor_nombre": safe_search(r"SEÑOR\(ES\):\s*(.+)", text, default="S/I"),
                    "deudor_rut": deudor_rut,
                    "folio": safe_search(r"Nº\s*(\d+)", text, default="0"),
                    "monto": safe_search(r"TOTAL[\s\$]*:?\s*([\d\.]+)", text, default="0"),
                    # Fecha de emisión (soporta 'Fecha Emision' o 'Fecha Emisión')
                        "fecha_emision": None,
                    # Valor bruto usar el mismo que 'monto' por solicitud
                    "valor_bruto": safe_search(r"TOTAL[\s\$]*:?\s*([\d\.]+)", text, default="0"),
                }
                # Formatear fecha_emision a DD/MM/YYYY si se puede
                raw_fecha = safe_search(r"Fecha\s+Emisi[oó]n[:\s]*([^\n]+)", text, default="S/I")

                def format_fecha(raw):
                    raw = raw.strip()
                    if raw == "S/I":
                        return raw
                    # dd/mm/yyyy or dd-mm-yyyy
                    m = re.search(r"(\d{1,2})\s*[/-]\s*(\d{1,2})\s*[/-]\s*(\d{2,4})", raw)
                    if m:
                        d, mo, y = m.groups()
                        y = y if len(y) == 4 else ("20" + y)
                        return f"{int(d):02d}/{int(mo):02d}/{int(y):04d}"
                    # textual spanish: '24 de Diciembre del 2025' or '24 de diciembre de 2025'
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
                    # fallback: return original trimmed
                    return raw

                data['fecha_emision'] = format_fecha(raw_fecha)
                return data

        except Exception as e:
            print(f"Error parsing {os.path.basename(pdf_path)}: {e}")
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
        <td style="border: 1px solid #000000; padding: 6px;">{item.get('fecha_emision','S/I')}</td>
        <td style="border: 1px solid #000000; padding: 6px;">{item.get('emisor_rut','S/I')}</td>
        <td style="border: 1px solid #000000; padding: 6px;">{item.get('emisor_nombre','S/I')}</td>
        <td style="border: 1px solid #000000; padding: 6px; text-align: center;">{item.get('folio','')}</td>
        <td style="border: 1px solid #000000; padding: 6px;">{item.get('deudor_rut','S/I')}</td>
        <td style="border: 1px solid #000000; padding: 6px;">{item.get('deudor_nombre','S/I')}</td>
        <td style="border: 1px solid #000000; padding: 6px; text-align: right;">{item.get('valor_bruto', item.get('monto','0'))}</td>
    </tr>
"""
        
        email_body = f"""<div style="font-family: Arial, sans-serif; font-size: 11pt; line-height: 1.6;">
    <p>Estimado:</p>
    
    <p>Junto con saludar, agradeceré a<br>
    <strong>{deudor_full}</strong><br>
    usted que pueda confirmar por este medio, la recepción y conformidad de las siguientes facturas electrónicas adjuntas, emitida por nuestro(s) cliente(s), las cuales están siendo cedidas a nuestro Factoring Punto Base Financiero Spa.</p>
    
    <div style="background-color: #d9d9d9; padding: 10px; margin: 15px 0; font-weight: bold;">
        {deudor_full}
    </div>
    
    <table style="border-collapse: collapse; width: 100%; margin: 15px 0;">
        <thead>
            <tr style="background-color: #f2f2f2;">
                <th style="border: 1px solid #000000; padding: 6px; text-align: left;">Fecha Emisión</th>
                <th style="border: 1px solid #000000; padding: 6px; text-align: left;">Rut Emisor</th>
                <th style="border: 1px solid #000000; padding: 6px; text-align: left;">Nombre Emisor</th>
                <th style="border: 1px solid #000000; padding: 6px; text-align: center;">N° Factura</th>
                <th style="border: 1px solid #000000; padding: 6px; text-align: left;">Rut Deudor</th>
                <th style="border: 1px solid #000000; padding: 6px; text-align: left;">Nombre Deudor</th>
                <th style="border: 1px solid #000000; padding: 6px; text-align: right;">Valor Bruto Factura</th>
            </tr>
        </thead>
        <tbody>
{rows_html}        </tbody>
    </table>
    
    <p>Favor ayudarnos con la siguiente información:<br>
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