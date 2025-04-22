import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

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
    valor_cuentas = df[df["Código"].isin(["314", "316", "318"])]["Valor"].sum()
    df_cxc = pd.DataFrame([["CUENTAS POR COBRAR", "CXC", valor_cuentas]], columns=df.columns)
    ingresos = df[df["Código"] == "6999"]["Valor"].sum()
    costos = df[df["Código"] == "7991"]["Valor"].sum()
    df_gb = pd.DataFrame([["GANANCIA BRUTA", "GB", ingresos - costos]], columns=df.columns)

    return pd.concat([df, df_cxc, df_gb], ignore_index=True), ingresos, valor_cuentas, costos

# Streamlit App
st.title("📄 Convertidor PDF a Excel - Declaración de Renta")

pdf_file = st.file_uploader("Sube tu declaración en PDF", type=["pdf"])

if pdf_file is not None:
    st.info("Procesando archivo...")

    df_final, ingresos, cuentas_por_cobrar, costos = extraer_datos(pdf_file)

    activos = df_final[df_final["Código"] == "349"]["Valor"].values[0]
    pasivo = df_final[df_final["Código"] == "599"]["Valor"].values[0]
    patrimonio = df_final[df_final["Código"] == "698"]["Valor"].values[0]
    intangibles = df_final[df_final["Código"] == "389"]["Valor"].values[0]
    utilidad = df_final[df_final["Código"] == "701"]["Valor"].values[0]
    ganancia = df_final[df_final["Código"] == "GB"]["Valor"].values[0]

    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # Hoja 1: DATA-BRUTO
        df_final.to_excel(writer, sheet_name="DATA-BRUTO", index=False)

        # Hoja 2: Decisioning
        workbook = writer.book
        worksheet = workbook.add_worksheet("Decisioning")

        title_format = workbook.add_format({"bold": True, "font_size": 14})
        header_format = workbook.add_format({"bold": True, "bg_color": "#4472C4", "font_color": "white", "align": "center"})
        normal_format = workbook.add_format({"align": "left"})
        
        # Título
        worksheet.write("B1", "Knockout Rules", title_format)

        headers = ["Sr No", "Parameters", "Criteria", "Actual Values", "Pass/Fail"]
        worksheet.write_row("B3", headers, header_format)

        data = [
            [1, "Minimum Annual revenue $5,000,000", ">=\$200,000", ingresos, '=SI(E4>=200000;"Pass";"Fail")'],
            [2, "Negative bank balance days in the last 6 months", "<=5", "No se cuenta con la información", ""],
            [3, "Liquidity Runway", ">=6 Months", round(activos / pasivo, 2), '=SI(E6>=6;"Pass";"Fail")'],
            [4, "If Tangible Net Worth is negative, business must be profitable", "N/A", patrimonio - intangibles, '=SI(E7>=0;"Pass";"Fail")'],
            [5, "Net Income Margin", "<-5%", round(utilidad / ingresos, 4), '=SI(E8>=-0.05;"Pass";"Fail")'],
            [6, "Current liabilities must not exceed 60% of annual revenue", "<=60%", round(pasivo / ingresos, 4), '=SI(E9<=0.6;"Pass";"Fail")'],
            [7, "Gross Margin", ">10%", round(ganancia / ingresos, 4), '=SI(E10>0.1;"Pass";"Fail")'],
            [8, "Minimum time in Business (In Years)", ">=3 Years", "", '=SI(E11>=3;"Pass";"Fail")'],
            [9, "Minimum Experian Intelliscore", "N/A", "", ""],
            [10, "Not delinquent on any Slope obligations or gone more than 15 days delinquent on any prior Slope obligations", "", "Aprobado Crédito", '=SI(E13="Aprobado Crédito";"Pass";"Fail")'],
        ]

        start_row = 3
        for i, row in enumerate(data):
            for j, val in enumerate(row):
                col = j + 1
                if isinstance(val, str) and val.startswith("=SI"):
                    worksheet.write_formula(start_row + i, col, val)
                else:
                    worksheet.write(start_row + i, col, val, normal_format)

        # Ajustar anchos
        worksheet.set_column("C:C", 62)
        worksheet.set_column("D:D", 15.71)
        worksheet.set_column("E:E", 31.57)

    output.seek(0)

    st.success("✅ Archivo procesado con éxito")
    st.download_button(
        label="📥 Descargar Excel",
        data=output,
        file_name="declaracion_convertida.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
