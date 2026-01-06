import pdfplumber
import re
import sys
p = r"N:\bastian\Desktop\Facturas\Factura N°855 - DVC (1).pdf"
print(f"Abriendo: {p}")
with pdfplumber.open(p) as pdf:
    for i, page in enumerate(pdf.pages[:3]):
        print('\n' + '='*40)
        print(f"Página {i+1}:")
        text = page.extract_text()
        if not text:
            print("(no se extrajo texto)")
            continue
        print(text)

        # Buscar patrones comunes de RUT
        patterns = [r"R\.U\.T\.?:?\s*([\d\.]+-\s*[\dkK])",
                    r"RUT:?\s*([\d\.]+-\s*[\dkK])",
                    r"R\.T\.?:?\s*([\d\.]+-\s*[\dkK])",
                    r"(\d{1,3}(?:\.\d{3})*-\d|\d+-[\dkKk])"]
        for pat in patterns:
            matches = re.findall(pat, text, re.IGNORECASE)
            if matches:
                print(f"Patrón: {pat} -> {matches}")

        # También mostrar líneas que contienen 'RUT' o 'R.U.T' o 'Rut' para inspección
        for ln in text.splitlines():
            if 'RUT' in ln.upper() or 'R.U.T' in ln.upper() or 'R.' in ln.upper():
                print("LINE:", ln)

print('\nHecho')
