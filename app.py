import streamlit as st
import pandas as pd
from PIL import Image
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
import datetime, os

st.set_page_config(page_title="Plan Intervención", layout="wide")

st.title("Plan Personalizado de Intervención")

# Sidebar: upload CSV and logo
st.sidebar.header("Archivos")
uploaded_csv = st.sidebar.file_uploader("Sube un CSV (una columna 'Objetivo')", type=["csv"])
uploaded_logo = st.sidebar.file_uploader("Sube logo (opcional, reemplaza logo.png)", type=["png","jpg","jpeg"])
if uploaded_logo is not None:
    logo = Image.open(uploaded_logo)
else:
    # load bundled logo.png
    try:
        logo = Image.open("logo.png")
    except Exception:
        logo = None

# Auto-detect course year
now = datetime.datetime.now()
year = now.year
# define academic year spanning current-year to next-year
acad_default = f"{year}-{year+1}"
# allow manual edit
curso_acad = st.text_input("Curso académico", value=acad_default)

# Header inputs
col1, col2 = st.columns([1,3])
with col1:
    if logo is not None:
        st.image(logo, use_column_width=False, width=150)
with col2:
    st.subheader("Aula Audición y Lenguaje")
    st.markdown("**Plan Personalizado de Intervención**")

st.header("Datos del alumno")
with st.form("datos_alumno"):
    nombre = st.text_input("ALUMNO/A", value="")
    curso = st.text_input("CURSO", value="")
    grupo = st.text_input("GRUPO", value="")
    etapa = st.text_input("ETAPA", value="")
    tutor = st.text_input("TUTOR/A", value="")
    fecha_inicio = st.date_input("FECHA DE INICIO", value=datetime.date.today())
    diagnostico = st.text_area("DIAGNÓSTICO", height=80)
    situacion_actual = st.text_area("SITUACIÓN ACTUAL", height=80)
    atencion = st.text_area("ATENCIÓN EDUCATIVA ESPECÍFICA", height=80)
    numero_sesiones = st.text_input("NÚMERO SESIONES", value="")
    obs_superiores = st.text_area("OBSERVACIONES (arriba)", height=80)
    submitted = st.form_submit_button("Guardar datos")

# Load CSV or example
if uploaded_csv is not None:
    df_in = pd.read_csv(uploaded_csv)
else:
    try:
        df_in = pd.read_csv("ejemplo.csv")
    except Exception:
        df_in = pd.DataFrame({"Objetivo": []})

if "Objetivo" not in df_in.columns:
    st.warning("El CSV debe contener una columna llamada 'Objetivo'.")
    st.write(df_in.head())
else:
    objetivos = df_in["Objetivo"].astype(str).tolist()
    # build editable dataframe with columns: Objetivo + each block/trimestre
    columns = ["Objetivo"]
    blocks = ["SIN INICIAR", "NECESITA MEJORAR", "ESTÁ PROGRESANDO", "CONSEGUIDO"]
    trimestres = ["1ºT", "2ºT", "3ºT"]
    for b in blocks:
        for t in trimestres:
            columns.append(f"{b} {t}")
    data = []
    for obj in objetivos:
        row = [obj] + [""] * (len(columns)-1)
        data.append(row)
    df_table = pd.DataFrame(data, columns=columns)

    st.subheader("Tabla editable")
    edited = st.data_editor(df_table, num_rows="dynamic", use_container_width=True)

    st.write("Marcadores: escribe 'X' o cualquier texto en la celda. Se exportará tal cual.")

    obs_final = st.text_area("OBSERVACIONES (finales)", height=120, value="")
    firma = st.text_input("Firmado:", value="")

    if st.button("Generar Word"):
        # create Document
        doc = Document()
        section = doc.sections[0]
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)

        # header with logo and title
        table = doc.add_table(rows=1, cols=2)
        table.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        table.autofit = False
        if logo is not None:
            # save logo to temp BytesIO with smaller size
            bio = BytesIO()
            logo.save(bio, format="PNG")
            bio.seek(0)
            cell = table.cell(0,0)
            p = cell.paragraphs[0]
            run = p.add_run()
            run.add_picture(bio, width=Inches(1.5))
        # title on right cell
        cell = table.cell(0,1)
        p = cell.paragraphs[0]
        p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        run = p.add_run("PLAN PERSONALIZADO DE INTERVENCIÓN\n")
        run.bold = True
        run.font.size = Pt(14)
        run2 = p.add_run("Aula Audición y Lenguaje\n")
        run2.font.size = Pt(12)

        doc.add_paragraph("")

        # Student info as a table
        info_tbl = doc.add_table(rows=5, cols=4)
        info_tbl.autofit = True
        info_tbl.cell(0,0).text = "ALUMNO/A:"
        info_tbl.cell(0,1).text = nombre
        info_tbl.cell(0,2).text = "CURSO:"
        info_tbl.cell(0,3).text = curso

        info_tbl.cell(1,0).text = "GRUPO:"
        info_tbl.cell(1,1).text = grupo
        info_tbl.cell(1,2).text = "ETAPA:"
        info_tbl.cell(1,3).text = etapa

        info_tbl.cell(2,0).text = "TUTOR/A:"
        info_tbl.cell(2,1).text = tutor
        info_tbl.cell(2,2).text = "FECHA DE INICIO:"
        info_tbl.cell(2,3).text = fecha_inicio.strftime("%d/%m/%Y")

        info_tbl.cell(3,0).text = "DIAGNÓSTICO:"
        info_tbl.cell(3,1).text = diagnostico
        info_tbl.cell(3,2).text = "SITUACIÓN ACTUAL:"
        info_tbl.cell(3,3).text = situacion_actual

        info_tbl.cell(4,0).text = "ATENCIÓN EDUCATIVA ESPECÍFICA:"
        info_tbl.cell(4,1).text = atencion
        info_tbl.cell(4,2).text = "Nº SESIONES:"
        info_tbl.cell(4,3).text = numero_sesiones

        doc.add_paragraph("")

        # Observaciones superiores
        if obs_superiores.strip():
            doc.add_paragraph("OBSERVACIONES:")
            doc.add_paragraph(obs_superiores)

        doc.add_paragraph("")

        # Build the main tracking table
        n_rows = len(edited)
        n_cols = 1 + 4*3  # objective + 4 blocks * 3 trimesters
        tbl = doc.add_table(rows=2 + n_rows, cols=n_cols)
        tbl.style = 'Table Grid'
        # First header row: merge cells for blocks
        # First cell is for "Objetivo" label spanning two rows
        tbl.cell(0,0).text = ""
        tbl.cell(1,0).text = "OBJETIVO"
        # Merge first cell down
        tbl.cell(0,0).merge(tbl.cell(1,0))

        # Fill block headers across 3 columns each
        col_idx = 1
        for block in ["SIN INICIAR", "NECESITA MEJORAR", "ESTÁ PROGRESANDO", "CONSEGUIDO"]:
            start = col_idx
            end = col_idx + 3 - 1
            # set text in top row merged cell
            top_cell = tbl.cell(0, start)
            # merge horizontally across 3 cells
            merge_to = tbl.cell(0, end)
            top_cell.text = block
            top_cell.merge(merge_to)
            # fill trimester headers in second row
            for i, t in enumerate(["1ºT","2ºT","3ºT"]):
                tbl.cell(1, start+i).text = t
            col_idx += 3

        # Fill objective rows
        for i in range(n_rows):
            tbl.cell(2+i,0).text = str(edited.iloc[i]["Objetivo"])
            col = 1
            for block in ["SIN INICIAR", "NECESITA MEJORAR", "ESTÁ PROGRESANDO", "CONSEGUIDO"]:
                for t in ["1ºT","2ºT","3ºT"]:
                    val = edited.iloc[i][f"{block} {t}"]
                    # if value is a single 'X' or 'x' or '❌', use a black cross character 'X' else write text
                    if isinstance(val, str) and val.strip() in ["X","x","❌","✖","✕"]:
                        tbl.cell(2+i, col).text = "X"
                    elif val is None:
                        tbl.cell(2+i, col).text = ""
                    else:
                        tbl.cell(2+i, col).text = str(val)
                    col += 1

        doc.add_paragraph("")

        # Observaciones finales
        if obs_final.strip():
            doc.add_paragraph("OBSERVACIONES FINALES:")
            doc.add_paragraph(obs_final)

        doc.add_paragraph("")
        doc.add_paragraph(f"Firmado: {firma}")

        # Filename includes student name and curso académico
        safe_name = nombre.strip().replace(" ", "_") if nombre.strip() else "sin_nombre"
        filename = f"Plan_Intervención_{safe_name}_{curso_acad}.docx"
        doc.save(filename)

        with open(filename, "rb") as f:
            btn = st.download_button("Descargar documento Word", data=f, file_name=filename, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        st.success(f"Documento generado: {filename}")
