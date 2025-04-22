import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.formatting import Rule
from openpyxl.styles import PatternFill
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.worksheet.datavalidation import DataValidation

# Función para extraer los datos del PDF
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

    # Agregar cálculos adicionales
    valor_cuentas = df[df["Código"].isin(["314", "316", "318"])]["Valor"].sum()
    df_cxc = pd.DataFrame([["CUENTAS POR COBRAR", "CXC", valor_cuentas]], columns=df.columns)

    ingresos = df[df["Código"] == "6999"]["Valor"].sum()
    costos = df[df["Código"] == "7991"]["Valor"].sum()
    df_gb = pd.DataFrame([["GANANCIA BRUTA", "GB", ingresos - costos]], columns=df.columns)

    df_final = pd.concat([df, df_cxc, df_gb], ignore_index=True)
    return df_final

# Interfaz Streamlit
st.title("📄 Convertidor PDF a Excel - SRI PERSONA NATURAL")

pdf_file = st.file_uploader("Sube tu archivo en PDF", type=["pdf"])

if pdf_file is not None:
    st.info("Procesando archivo...")

    # Extraer y preparar los datos
    df_final = extraer_datos(pdf_file)

    # Cargar plantilla base desde el mismo directorio
    plantilla_path = "Plantilla.xlsx"  # debe estar en el mismo repositorio que app.py
    wb = load_workbook(plantilla_path)

    # Acceder a la hoja "DATA-BRUTO" sin eliminarla
    if "DATA-BRUTO" not in wb.sheetnames:
        st.error("La hoja 'DATA-BRUTO' no existe en la plantilla.")
    else:
        ws = wb["DATA-BRUTO"]

        # Limpiar contenido desde fila 2 (dejamos encabezados en fila 1)
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=3):
            for cell in row:
                cell.value = None

        # Escribir los datos en la hoja
        for row_idx, row in df_final.iterrows():
            for col_idx, value in enumerate(row, start=1):
                ws.cell(row=row_idx + 2, column=col_idx, value=value)



        # Asegurar que la hoja "Decisioning" existe
        if "Decisioning" in wb.sheetnames:
            ws_decisioning = wb["Decisioning"]
        
            # Eliminar reglas previas si quieres
            ws_decisioning.conditional_formatting.clear()
        
            # Estilo para "Pass"
            fill_pass = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            dxf_pass = DifferentialStyle(fill=fill_pass)
            rule_pass = Rule(type="containsText", operator="containsText", text="Pass", dxf=dxf_pass)
            rule_pass.formula = ['NOT(ISERROR(SEARCH("Pass",F4)))']
            ws_decisioning.conditional_formatting.add("F4:F13", rule_pass)
        
            # Estilo para "Fail"
            fill_fail = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            dxf_fail = DifferentialStyle(fill=fill_fail)
            rule_fail = Rule(type="containsText", operator="containsText", text="Fail", dxf=dxf_fail)
            rule_fail.formula = ['NOT(ISERROR(SEARCH("Fail",F4)))']
            ws_decisioning.conditional_formatting.add("F4:F13", rule_fail)


        # Guardar en memoria
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        # Descargar
        st.success("✅ Archivo generado correctamente.")
        st.download_button(
            label="📥 Descargar Excel",
            data=output,
            file_name="Slope Policy Output SRI PERSONA NATURAL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )




