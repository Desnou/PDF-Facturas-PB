from app_facturas import NativeInvoiceApp
p = r"N:\bastian\Desktop\Facturas\Factura N°855 - DVC (1).pdf"
app = NativeInvoiceApp()
try:
    app.withdraw()
except Exception:
    pass
app.add_file_card(p)
# ensure pdf_files populated
app.pdf_files = [p]
app.generate_email()
print(app.current_html)
try:
    app.destroy()
except Exception:
    pass
