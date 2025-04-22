import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.formatting import Rule
from openpyxl.styles import PatternFill
from openpyxl.styles.differential import DifferentialStyle
import pyodbc

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

    # Agregar cálculos adicionales
    valor_cuentas = df[df["Código"].isin(["314", "316", "318"])]["Valor"].sum()
    df_cxc = pd.DataFrame([["CUENTAS POR COBRAR", "CXC", valor_cuentas]], columns=df.columns)
    ingresos = df[df["Código"] == "6999"]["Valor"].sum()
    costos = df[df["Código"] == "7991"]["Valor"].sum()
    df_gb = pd.DataFrame([["GANANCIA BRUTA", "GB", ingresos - costos]], columns=df.columns)

    df_final = pd.concat([df, df_cxc, df_gb], ignore_index=True)
    return df_final

# Función para consultar SQL Server
def obtener_datos_credito(cod_cliente):
    conn = pyodbc.connect(
        "DRIVER={ODBC Driver 17 for SQL Server};"
        "SERVER=tu_servidor;"
        "DATABASE=tu_base;"
        "UID=tu_usuario;"
        "PWD=tu_contraseña;"
    )
    query = """
    SELECT 
        COD_CUENTA_CLIENTE,
        NOMBRE_CLIENTE,
        CUPO_ACTUAL_CREDITO,
        CUPO_DISPONIBLE,
        ANIOS_ACT_ECONOMICA,
        LETRA_EQUIFAX,
        SCOREBURO_EQUIFAX,
        SCORE_CREDITO,
        DEUDA_TOTAL,
        LETRA_FM,
        ESTADO_CREDITO,
        IDENTIFICACION,
        BURO_EQUIFAX,
        INDEX_PYMES_EQUIFAX,
        app.KYC_Estado
    FROM BI_DIM_CLIENTE clt
    LEFT JOIN APP_FLUJO_CLIENTE_DETALLE app
        ON clt.COD_CUENTA_CLIENTE = app.Cod_Cliente_AX
    WHERE COD_CUENTA_CLIENTE = ?
    """
    df = pd.read_sql(query, conn, params=[cod_cliente])
    conn.close()
    return df

# Interfaz Streamlit
st.title("📄 Convertidor PDF a Excel con Datos de Crédito")

# Campos de entrada
codigo_cliente = st.text_input("🔢 Ingresa el código del cliente")
pdf_file = st.file_uploader("📄 Sube tu declaración en PDF", type=["pdf"])

if pdf_file is not None and codigo_cliente:
    st.info("Procesando archivo y consultando base de datos...")

    # Extraer datos del PDF
    df_final = extraer_datos(pdf_file)

    # Consultar datos de crédito
    df_credito = obtener_datos_credito(codigo_cliente)

    # Cargar plantilla
    plantilla_path = "Plantilla.xlsx"
    wb = load_workbook(plantilla_path)

    # Insertar en hoja DATA-BRUTO
    if "DATA-BRUTO" in wb.sheetnames:
        ws = wb["DATA-BRUTO"]
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=3):
            for cell in row:
                cell.value = None
        for row_idx, row in df_final.iterrows():
            for col_idx, value in enumerate(row, start=1):
                ws.cell(row=row_idx + 2, column=col_idx, value=value)

    # Insertar en hoja CREDITO
    if "CREDITO" in wb.sheetnames and not df_credito.empty:
        ws_cred = wb["CREDITO"]
        for row in ws_cred.iter_rows(min_row=2, max_row=ws_cred.max_row):
            for cell in row:
                cell.value = None
        for col_idx, col in enumerate(df_credito.columns, start=1):
            ws_cred.cell(row=1, column=col_idx, value=col)
        for row_idx, row in df_credito.iterrows():
            for col_idx, val in enumerate(row, start=1):
                ws_cred.cell(row=row_idx + 2, column=col_idx, value=val)

    # Formato condicional en hoja Decisioning
    if "Decisioning" in wb.sheetnames:
        ws2 = wb["Decisioning"]

        fill_pass = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        dxf_pass = DifferentialStyle(fill=fill_pass)
        rule_pass = Rule(type="containsText", operator="containsText", text="Pass", dxf=dxf_pass)
        rule_pass.formula = ['NOT(ISERROR(SEARCH("Pass",F4)))']
        ws2.conditional_formatting.add("F4:F13", rule_pass)

        fill_fail = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        dxf_fail = DifferentialStyle(fill=fill_fail)
        rule_fail = Rule(type="containsText", operator="containsText", text="Fail", dxf=dxf_fail)
        rule_fail.formula = ['NOT(ISERROR(SEARCH("Fail",F4)))']
        ws2.conditional_formatting.add("F4:F13", rule_fail)

    # Guardar archivo final
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.success("✅ Archivo generado correctamente.")
    st.download_button(
        label="📥 Descargar Excel",
        data=output,
        file_name="Slope Policy Output SRI PERSONA NATURAL.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
