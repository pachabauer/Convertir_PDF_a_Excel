import re
import pdfplumber
import pandas as pd

# Expresión regular para capturar los grupos
new_line = re.compile(
    r'(\d{2}/\d{2})\s*(D(?: \d{3})?)?\s*(.*?)\s*(-?\d{1,3}(?:\.\d{3})*,\d{2})?\s*(-?\d{1,3}(?:\.\d{3})*,\d{2})?\s*(-?\d{1,3}(?:\.\d{3})*,\d{2})')


def parse_amount(amount):
    if amount:
        amount = amount.replace('.', '').replace(',', '.')
        return float(amount)
    return None


def extract_data_from_pdf(pdf_path, pages):
    fechas = []
    origenes = []
    conceptos = []
    debitos = []
    creditos = []
    saldos = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_num in pages:
            page = pdf.pages[page_num - 1]
            text = page.extract_text()

            for line in text.split('\n'):
                match = new_line.match(line)
                if match:
                    fecha, origen, concepto, importe1, importe2, saldo = match.groups()

                    # Determinar si los importes son débitos o créditos
                    if importe1 and '-' in importe1:
                        debito = parse_amount(importe1)
                        credito = parse_amount(importe2 if importe2 and '-' not in importe2 else "")
                    else:
                        debito = parse_amount(importe2 if importe2 and '-' in importe2 else "")
                        credito = parse_amount(importe1 if importe1 and '-' not in importe1 else "")

                    fechas.append(fecha)
                    origenes.append(origen.strip() if origen else "")
                    conceptos.append(concepto.strip() if concepto else "")
                    debitos.append(debito)
                    creditos.append(credito)
                    saldos.append(parse_amount(saldo))

    return pd.DataFrame({
        'Fecha': fechas,
        'Origen': origenes,
        'Concepto': conceptos,
        'Débito': debitos,
        'Crédito': creditos,
        'Saldo': saldos
    })


# Ejemplo de uso directo del módulo
if __name__ == "__main__":
    # Parámetros de entrada
    pdf_path = r"C:\Users\esteb\Downloads\Resumen BBVA - Nortear - Marzo 2024.pdf"
    pages = [2, 3, 4, 5, 6]  # Páginas de la 2 a la 6

    # Extraer datos y crear el DataFrame
    df = extract_data_from_pdf(pdf_path, pages)

    # Guardar el DataFrame en un archivo Excel
    df.to_excel('extracto_bancario.xlsx', index=False)

    print("Archivo Excel creado con éxito.")
