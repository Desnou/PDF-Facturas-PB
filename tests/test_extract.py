import sys
from app_facturas import NativeInvoiceApp

PDF_PATH = r"N:\bastian\Desktop\Facturas\Factura N°855 - DVC (1).pdf"

def main():
    app = NativeInvoiceApp()
    try:
        app.withdraw()
    except Exception:
        pass
    data = app.extract_pdf_data(PDF_PATH)
    try:
        assert data is not None, "No se extrajeron datos"
        assert data.get('deudor_rut') and data['deudor_rut'] != 'S/I', "Deudor RUT no extraído"
        assert data.get('emisor_rut') and data['emisor_rut'] != 'S/I', "Emisor RUT no extraído"
        assert data.get('monto') and data['monto'] != '0', "Monto no extraído"
        assert 'fecha_emision' in data and data['fecha_emision'] != 'S/I', "Fecha emisión no extraída"
    except AssertionError as e:
        print('TEST FAILED:', e)
        print('Datos extraídos:', data)
        sys.exit(2)

    print('TEST PASSED')
    print('Datos extraídos:', data)
    try:
        app.destroy()
    except Exception:
        pass

if __name__ == '__main__':
    main()
