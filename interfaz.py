import streamlit as st
import pdfplumber
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill

# Configuración de la página
st.set_page_config(page_title="SENA AVAL", layout="centered")

# --- LÓGICA DE REINICIO ---
if 'uploader_key' not in st.session_state:
    st.session_state.uploader_key = 0

if 'datos_listos' not in st.session_state:
    st.session_state.datos_listos = None

def reiniciar_todo():
    # Cambiamos la llave del uploader para que se limpie el campo de archivos
    st.session_state.uploader_key += 1
    # Limpiamos los datos procesados
    st.session_state.datos_listos = None

# --------------------------

st.title("📄 Automatización de Pagos AVAL")
st.write("Sube los PDFs y presiona Ejecutar para generar el reporte.")

def limpiar_monto(texto):
    if not texto: return 0.0
    encontrado = re.search(r'([\d\.]+,\d{2})|(\d+)', texto)
    if encontrado:
        valor = encontrado.group().replace(".", "").replace(",", ".")
        try: return float(valor)
        except: return 0.0
    return 0.0

# 1. Zona de subida (con KEY dinámica para poder resetearla)
uploaded_files = st.file_uploader(
    "Paso 1: Arrastra los PDFs aquí", 
    type="pdf", 
    accept_multiple_files=True,
    key=f"uploader_{st.session_state.uploader_key}"
)

# 2. Botón Ejecutar
if uploaded_files:
    if st.button("Paso 2: EJECUTAR PROCESO"):
        base_datos = []
        with st.spinner('Procesando archivos...'):
            for uploaded_file in uploaded_files:
                with pdfplumber.open(uploaded_file) as pdf:
                    texto = "\n".join([p.extract_text() for p in pdf.pages])
                    lineas = texto.split('\n')
                    datos = {"USO": "A-02-02-02-009-002-09", "Mes": "Febrero", "Bruto": 0.0, "ICA": 0.0}
                    
                    for linea in lineas:
                        if "Compromiso SIIF" in linea:
                            m = re.search(r'SIIF\s+(\d+)', linea)
                            if m: datos["Compromiso"] = m.group(1)
                        if "Nombres y apellidos:" in linea:
                            datos["Nombre"] = linea.split("apellidos:")[1].split("Banco")[0].strip()
                        if "Cédula de Ciudadanía" in linea:
                            m = re.search(r'Ciudadanía\s+([\d\.]+)', linea)
                            if m: datos["ID"] = m.group(1).replace(".", "")
                        if "Valor Bruto Pago:" in linea:
                            datos["Bruto"] = limpiar_monto(linea.split("Pago:")[1].split("Saldo")[0])
                        if "Reteica - 8299" in linea:
                            datos["ICA"] = limpiar_monto(linea.split("MANIZALES")[1])

                    if datos.get("Nombre"): base_datos.append(datos)
            
            st.session_state.datos_listos = base_datos
            st.success("✅ Proceso completado. Ahora puedes descargar el Excel.")

# 3. Botón Descargar y Reset
if st.session_state.datos_listos:
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    
    # Encabezados
    ws.merge_cells('A1:B1'); ws['A1'] = "Nombre supervisor"
    ws.merge_cells('J1:K1'); ws['J1'] = "Número de identificación"
    ws.merge_cells('A2:B2'); ws['A2'] = "Dependencia"
    ws.merge_cells('C2:I2'); ws['C2'] = "Centro de Procesos Industriales y de la Construcción"
    ws.merge_cells('J2:K2'); ws['J2'] = "Fecha de la solicitud"
    ws.merge_cells('L2:N2'); ws['L2'] = "FEBRERO"

    headers = ["No. Consecutivo", "No. Compromiso ptal.", "USO", "Mes a pagar", "Nombre del contratista", "ID", "Número de Identificación", "Valor bruto", "Valor ICA", "Ret_F", "Ret_IVA", "Ret_Emb", "Otras", "Total"]
    ws.append(headers)

    for i, d in enumerate(st.session_state.datos_listos, start=1):
        idx = ws.max_row + 1
        formula_total = f"=H{idx}-I{idx}" 
        
        ws.append([
            i, d.get("Compromiso"), d.get("USO"), d.get("Mes"), d.get("Nombre"), "CC", d.get("ID"),
            d.get("Bruto"), d.get("ICA"), 0, 0, 0, 0, formula_total
        ])
        
        for col in range(8, 15): ws.cell(row=idx, column=col).number_format = '#,##0'
        ws[f'H{idx}'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        ws[f'I{idx}'].fill = PatternFill(start_color="00B0F0", end_color="00B0F0", fill_type="solid")
        ws[f'N{idx}'].fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")

    wb.save(output)
    
    st.download_button(
        label="Paso 3: 📥 DESCARGAR EXCEL Y LIMPIAR",
        data=output.getvalue(),
        file_name="RESULTADO_AVAL_FEBRERO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        on_click=reiniciar_todo  # <--- Esto borra los PDFs y reinicia la página
    )