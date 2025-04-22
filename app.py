import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook

# Función para extraer datos del PDF
def extraer_datos(pdf_file):
    cuentas_objetivo = [
        ("349", "TOTAL ACTIVOS CORRIENTES"),
        ("389", "TOTAL ACTIVOS INTANGIBLES"),
        ("599", "TOTAL DEL PASIVO"),
        ("698", "TOTAL PATRIMONIO NETO"),
        ("701", "UTILIDAD DEL EJERCICIO"),
        ("6999", "TOTAL INGRESOS"),
        ("7991", "TOTAL COSTOS"),
        ("314", "Locales"),
        ("316", "Locales"),
        ("318", "Locales"),
    ]

    resultado = []
    pattern = re.compile(r"^(.*?)\s+(\d{3,4})\s+([\d.,-]+)$")

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                for line in text.split("\n"):
                    match = pattern.match(line.strip())
                    if match:
                        descripcion = match.group(1).strip()
                        codigo = match.group(2)
                        valor = float(match.group(3).replace(",", ""))
                        for cod_ref, desc_ref in cuentas_objetivo:
                            if codigo == cod_ref and desc_ref.lower() in descripcion.lower():
                                resultado.append([descripcion, codigo, valor])
                                break

    df = pd.DataFrame(resultado, columns=["Descripción", "Código", "Valor"])
    
    # Cálculos adicionales
    valor_cuentas = df[df["Código"].isin(["314", "316", "318"])]["Valor"].sum()
    df_cxc = pd.DataFrame([["CUENTAS POR COBRAR", "CXC", valor_cuentas]], columns=df.columns)

    ingresos = df[df["Código"] == "6999"]["Valor"].sum()
    costos = df[df["Código"] == "7991"]["Valor"].sum()
    df_gb = pd.DataFrame([["GANANCIA BRUTA", "GB", ingresos - costos]], columns=df.columns)

    df_final = pd.concat([df, df_cxc, df_gb], ignore_index=True)
    return df_final

# Interfaz de Streamlit
st.title("📄 Convertidor PDF a Excel - SRI PERSONA NATURAL")

pdf_file = st.file_uploader("Sube tu archivo del SRI en PDF", type=["pdf"])

if pdf_file is not None:
    st.info("Procesando archivo...")

    # Extraer los datos
    df_final = extraer_datos(pdf_file)

    # Cargar la plantilla Excel
    plantilla_path = "Plantilla.xlsx"  # debe estar en el mismo directorio que app.py en GitHub
    wb = load_workbook(plantilla_path)

    # Si ya existe la hoja DATA-BRUTO, la borramos
    if "DATA-BRUTO" in wb.sheetnames:
        del wb["DATA-BRUTO"]

    # Crear nueva hoja DATA-BRUTO
    ws = wb.create_sheet("DATA-BRUTO")

    # Escribir encabezados
    for col_idx, col_name in enumerate(df_final.columns, start=1):
        ws.cell(row=1, column=col_idx, value=col_name)

    # Escribir los valores
    for row_idx, row in df_final.iterrows():
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row=row_idx + 2, column=col_idx, value=value)

    # Guardar en memoria
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.success("✅ Archivo generado con éxito a partir de la plantilla")
    st.download_button(
        label="📥 Descargar Excel",
        data=output,
        file_name="Slope Policy Output SRI PERSONA NATURAL.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
