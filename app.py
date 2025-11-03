import streamlit as st
import pandas as pd
from PIL import Image
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import datetime, os

st.set_page_config(page_title="Plan Intervención v2", layout="wide")

st.title("Plan Personalizado de Intervención — v2")

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

# Auto-detect academic year default, editable
now = datetime.datetime.now()
year = now.year
acad_default = f"{year}-{year+1}"
curso_acad = st.text_input("Curso académico", value=acad_default)

# Header layout: logo left, course right, title centered below
col1, col2, col3 = st.columns([1,1,1])
with col1:
    if logo is not None:
        st.image(logo, use_column_width=False, width=140)
with col3:
    st.markdown(f"**Curso académico:** {curso_acad}", unsafe_allow_html=True)
# Title centered below
st.markdown("<h2 style='text-align:center; font-family:Calibri; font-size:16pt'>PLAN PERSONALIZADO DE INTERVENCIÓN</h2>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align:center; font-family:Calibri; font-size:12pt'>Aula de Audición y Lenguaje</h3>", unsafe_allow_html=True)

st.header("Rellena los datos")
with st.form("datos_alumno"):
    nombre = st.text_input("ALUMNO/A", value="")
    fecha_nac = st.date_input("Fecha de nacimiento", value=None)
    curso = st.text_input("CURSO", value="")
    grupo = st.text_input("GRUPO", value="")
    etapa = st.text_input("ETAPA", value="")
    tutor = st.text_input("TUTOR/A", value="")
    maestra_al = st.text_input("Maestra AL", value="Paula Bort Museros")
    diagnostico = st.text_area("DIAGNÓSTICO", height=80)
    fecha_inicio = st.date_input("FECHA DE INICIO", value=datetime.date.today())
    numero_sesiones = st.text_input("NÚMERO SESIONES", value="")
    situacion_actual = st.text_area("SITUACIÓN ACTUAL", height=120)
    atencion = st.text_area("ATENCIÓN EDUCATIVA ESPECÍFICA", height=120)
    observaciones = st.text_area("OBSERVACIONES", height=120)
    obs_trim1 = st.text_area("Observaciones 1º Trimestre", height=80)
    obs_trim2 = st.text_area("Observaciones 2º Trimestre", height=80)
    obs_trim3 = st.text_area("Observaciones 3º Trimestre", height=80)
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

    st.subheader("Tabla editable de objetivos")
    edited = st.data_editor(df_table, num_rows="dynamic", use_container_width=True)

    st.write("Marcadores: escribe 'X' o cualquier texto en la celda. Pulsa 'Vista previa' para ver la tabla con colores.")

    # Preview button shows colored HTML table (only the objectives table)
    if st.button("Vista previa de la tabla"):
        color_map = {
            "SIN INICIAR": "#D9D9D9",
            "NECESITA MEJORAR": "#FFD8B1",
            "ESTÁ PROGRESANDO": "#FFF9B1",
            "CONSEGUIDO": "#C7E7B3"
        }
        html = "<div style='font-family:Calibri; font-size:10pt; overflow:auto;'>"
        html += "<table style='border-collapse:collapse; width:100%;'>"
        html += "<tr>"
        html += "<th style='border:1px solid #000; padding:4px; text-align:left; width:35%;' rowspan='2'>OBJETIVO</th>"
        for block in blocks:
            html += f"<th style='border:1px solid #000; padding:4px; text-align:center; background:{color_map[block]};' colspan='3'>{block}</th>"
        html += "</tr>"
        html += "<tr>"
        for _ in blocks:
            for t in trimestres:
                html += f"<th style='border:1px solid #000; padding:3px; width:7%; text-align:center;'>{t}</th>"
        html += "</tr>"
        for i, row in edited.iterrows():
            html += "<tr>"
            html += f"<td style='border:1px solid #000; padding:3px; font-size:9pt;'>{row['Objetivo']}</td>"
            for block in blocks:
                for t in trimestres:
                    cell = row[f'{block} {t}']
                    display = ""
                    if isinstance(cell, str) and cell.strip().lower() in ['x','❌','✖','✕']:
                        display = "X"
                    elif pd.isna(cell) or (isinstance(cell, str) and cell.strip() == ""):
                        display = ""
                    else:
                        display = str(cell)
                    html += f"<td style='border:1px solid #000; padding:3px; text-align:center; font-size:9pt;'>{display}</td>"
            html += "</tr>"
        html += "</table></div>"
        st.markdown(html, unsafe_allow_html=True)

    obs_final = st.text_area("Observaciones finales (opcional)", height=120, value="")
    firma = st.text_input("Texto firma (se añadirá abajo, alineado a la derecha)", value="Paula Bort Museros\nMaestra de Audición y Lenguaje")

    if st.button("Generar Word"):
        def set_run_font(run, name="Calibri", size_pt=11, bold=False):
            run.font.name = name
            run.font.size = Pt(size_pt)
            run.bold = bold
            try:
                run._element.rPr.rFonts.set(qn('w:eastAsia'), name)
            except Exception:
                pass

        def set_cell_bg(cell, color_hex):
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:fill'), color_hex.replace('#',''))
            tcPr.append(shd)

        doc = Document()
        style = doc.styles['Normal']
        style.font.name = 'Calibri'
        style.font.size = Pt(11)

        section = doc.sections[0]
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)

        header_tbl = doc.add_table(rows=1, cols=2)
        header_tbl.autofit = False
        try:
            header_tbl.columns[0].width = Inches(1.8)
            header_tbl.columns[1].width = Inches(5.2)
        except Exception:
            pass
        if logo is not None:
            bio = BytesIO()
            logo.save(bio, format="PNG")
            bio.seek(0)
            cell0 = header_tbl.cell(0,0)
            p0 = cell0.paragraphs[0]
            run0 = p0.add_run()
            run0.add_picture(bio, width=Inches(1.6))
        cell1 = header_tbl.cell(0,1)
        p1 = cell1.paragraphs[0]
        p1.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
        run1 = p1.add_run(f"Curso académico: {curso_acad}")
        run1.bold = True
        set_run_font(run1, size_pt=11)

        doc.add_paragraph("")
        p_title = doc.add_paragraph()
        p_title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        rtitle = p_title.add_run("PLAN PERSONALIZADO DE INTERVENCIÓN\n")
        rtitle.bold = True
        set_run_font(rtitle, size_pt=14)
        rsub = p_title.add_run("Aula de Audición y Lenguaje\n")
        set_run_font(rsub, size_pt=12)

        doc.add_paragraph("")

        p = doc.add_paragraph()
        r = p.add_run("1. DATOS DEL ALUMNO")
        r.bold = True
        set_run_font(r, size_pt=12)
        doc.add_paragraph("")

        info_tbl = doc.add_table(rows=7, cols=2)
        info_tbl.autofit = False
        try:
            info_tbl.columns[0].width = Inches(2.0)
            info_tbl.columns[1].width = Inches(4.0)
        except Exception:
            pass
        labels = ["Alumno/a", "Fecha de nacimiento", "Curso", "Tutor/a", "Maestra AL", "Diagnóstico", "Fecha de inicio"]
        values = [nombre, fecha_nac.strftime("%d/%m/%Y") if fecha_nac is not None else "", curso, tutor, maestra_al, diagnostico, fecha_inicio.strftime("%d/%m/%Y")]
        for i, lab in enumerate(labels):
            cell_l = info_tbl.cell(i,0)
            cell_r = info_tbl.cell(i,1)
            cell_l.text = lab + ":"
            cell_r.text = values[i]

        doc.add_paragraph("")
        p = doc.add_paragraph()
        r = p.add_run("2. SITUACIÓN ACTUAL")
        r.bold = True
        set_run_font(r, size_pt=12)
        doc.add_paragraph(situacion_actual)

        doc.add_paragraph("")
        p = doc.add_paragraph()
        r = p.add_run("3. ATENCIÓN EDUCATIVA ESPECÍFICA")
        r.bold = True
        set_run_font(r, size_pt=12)
        doc.add_paragraph(atencion)

        doc.add_paragraph("")
        p = doc.add_paragraph()
        r = p.add_run("4. NÚMERO DE SESIONES")
        r.bold = True
        set_run_font(r, size_pt=12)
        doc.add_paragraph(numero_sesiones)

        doc.add_paragraph("")
        p = doc.add_paragraph()
        r = p.add_run("5. OBSERVACIONES")
        r.bold = True
        set_run_font(r, size_pt=12)
        doc.add_paragraph(observaciones)

        doc.add_paragraph("")
        p = doc.add_paragraph()
        r = p.add_run("6. OBJETIVOS A TRABAJAR")
        r.bold = True
        set_run_font(r, size_pt=12)
        doc.add_paragraph("")

        n_rows = len(edited)
        n_cols = 1 + 4*3
        tbl = doc.add_table(rows=2 + n_rows, cols=n_cols)
        tbl.style = 'Table Grid'
        tbl.autofit = False
        widths = [Inches(3.2)] + [Inches(0.6)] * (n_cols-1)
        for i, w in enumerate(widths):
            try:
                tbl.columns[i].width = w
            except Exception:
                for row in tbl.rows:
                    row.cells[i].width = w

        tbl.cell(0,0).text = ""
        tbl.cell(1,0).text = "OBJETIVO"
        tbl.cell(0,0).merge(tbl.cell(1,0))

        color_map = {
            "SIN INICIAR": "D9D9D9",
            "NECESITA MEJORAR": "FFD8B1",
            "ESTÁ PROGRESANDO": "FFF9B1",
            "CONSEGUIDO": "C7E7B3"
        }

        def set_cell_bg(cell, color_hex):
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:fill'), color_hex.replace('#',''))
            tcPr.append(shd)

        col_idx = 1
        for block in ["SIN INICIAR", "NECESITA MEJORAR", "ESTÁ PROGRESANDO", "CONSEGUIDO"]:
            start = col_idx
            end = col_idx + 3 - 1
            top_cell = tbl.cell(0, start)
            top_cell.text = block
            top_cell.merge(tbl.cell(0, end))
            set_cell_bg(top_cell, color_map[block])
            for i, t in enumerate(["1ºT","2ºT","3ºT"]):
                cell = tbl.cell(1, start+i)
                cell.text = t
                set_cell_bg(cell, color_map[block])
            col_idx += 3

        for i in range(n_rows):
            obj_text = str(edited.iloc[i]["Objetivo"])
            tbl.cell(2+i,0).text = obj_text
            col = 1
            for block in ["SIN INICIAR", "NECESITA MEJORAR", "ESTÁ PROGRESANDO", "CONSEGUIDO"]:
                for t in ["1ºT","2ºT","3ºT"]:
                    val = edited.iloc[i][f"{block} {t}"]
                    cell = tbl.cell(2+i, col)
                    set_cell_bg(cell, color_map[block])
                    if isinstance(val, str) and val.strip().lower() in ["x","❌","✖","✕"]:
                        display = "X"
                    elif pd.isna(val) or (isinstance(val, str) and val.strip() == ""):
                        display = ""
                    else:
                        display = str(val)
                    cell.text = display
                    col += 1

        doc.add_paragraph("")
        p = doc.add_paragraph()
        r = p.add_run("7. OBSERVACIONES TRIMESTRALES")
        r.bold = True
        set_run_font = None
        doc.add_paragraph("Observaciones 1º Trimestre:")
        doc.add_paragraph(obs_trim1)
        doc.add_paragraph("Observaciones 2º Trimestre:")
        doc.add_paragraph(obs_trim2)
        doc.add_paragraph("Observaciones 3º Trimestre:")
        doc.add_paragraph(obs_trim3)

        doc.add_paragraph("")
        if obs_final.strip():
            doc.add_paragraph("OBSERVACIONES FINALES:")
            doc.add_paragraph(obs_final)

        p_sig = doc.add_paragraph()
        p_sig.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
        r_sig = p_sig.add_run(f"Firmado: {firma.splitlines()[0] if firma else 'Paula Bort Museros'}\n{firma.splitlines()[1] if firma and len(firma.splitlines())>1 else 'Maestra de Audición y Lenguaje'}")
        # simple font for signature
        # save
        safe_name = nombre.strip().replace(" ", "_") if nombre.strip() else "sin_nombre"
        filename = f"Plan_Intervención_{safe_name}_{curso_acad}.docx"
        doc.save(filename)

        with open(filename, "rb") as f:
            st.download_button("Descargar documento Word", data=f, file_name=filename, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        st.success(f"Documento generado: {filename}")
