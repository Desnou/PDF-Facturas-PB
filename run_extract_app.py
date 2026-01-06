from app_facturas import NativeInvoiceApp
import os
p = r"N:\bastian\Desktop\Facturas\Factura N°855 - DVC (1).pdf"
app = NativeInvoiceApp()
# evitar mostrar la ventana
try:
    app.withdraw()
except Exception:
    pass
res = app.extract_pdf_data(p)
print('Resultado extracción:')
print(res)
# destruir la app para salir limpio
try:
    app.destroy()
except Exception:
    pass
