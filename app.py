import streamlit as st
import pandas as pd
from datetime import datetime

# Configuración de página
st.set_page_config(
    page_title="Control de Órdenes de Trabajo - Corte y Repo",
    page_icon="🛠️",
    layout="wide"
)

st.title("🛠️ Control y Verificación de Órdenes de Trabajo")
st.write("Gestiona la verificación de direcciones ejecutadas y pendientes a partir del registro del servicio.")

# Cargar archivo Excel
uploaded_file = st.sidebar.file_uploader("Cargar archivo 'corte y repo.xlsx'", type=["xlsx"])

if uploaded_file is not None:
    @st.cache_data
    @st.cache_data
def load_data(file):
    # Leer el archivo Excel para inspeccionar los nombres de las hojas
    xls = pd.ExcelFile(file)
    
    # Si existe una pestaña llamada 'Hoja2', la usa; si no, toma la primera pestaña disponible
    sheet_to_use = 'Hoja2' if 'Hoja2' in xls.sheet_names else xls.sheet_names[0]
    
    # Cargar los datos
    df = pd.read_excel(xls, sheet_name=sheet_to_use)
    
    # Si la primera fila es un encabezado fuera de lugar, ajustamos la lectura
    if 'Número OT' not in df.columns:
        df = pd.read_excel(xls, sheet_name=sheet_to_use, header=1)
        
    # Limpieza de columnas vacías generadas por formato
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    
    # Inicializar columnas de verificación
    if 'Estado Verificación' not in df.columns:
        df['Estado Verificación'] = 'Pendiente'
    if 'Observación Verificador' not in df.columns:
        df['Observación Verificador'] = ''
    if 'Fecha Verificación' not in df.columns:
        df['Fecha Verificación'] = ''
        
    return df
    df = load_data(uploaded_file)

    # Persistencia en session_state para permitir modificaciones
    if 'data' not in st.session_state:
        st.session_state.data = df.copy()

    # --- BARRA LATERAL: FILTROS ---
    st.sidebar.header("🔍 Filtros de Búsqueda")
    
    # Filtro por Estado
    estado_filtro = st.sidebar.multiselect(
        "Estado de Verificación:",
        options=['Pendiente', 'Ejecutada'],
        default=['Pendiente', 'Ejecutada']
    )
    
    # Filtro por Tipo de OT
    tipos_ot = st.sidebar.multiselect(
        "Tipo de OT:",
        options=st.session_state.data['Tipo OT'].dropna().unique(),
        default=st.session_state.data['Tipo OT'].dropna().unique()
    )

    # Búsqueda por texto (Dirección, OT o Cliente)
    search_query = st.sidebar.text_input("Buscar por Dirección, OT o Cliente:")

    # Aplicar filtros
    filtered_df = st.session_state.data[
        (st.session_state.data['Estado Verificación'].isin(estado_filtro)) &
        (st.session_state.data['Tipo OT'].isin(tipos_ot))
    ]

    if search_query:
        query = search_query.lower()
        filtered_df = filtered_df[
            filtered_df['Dirección'].astype(str).str.lower().str.contains(query) |
            filtered_df['Número OT'].astype(str).str.lower().str.contains(query) |
            filtered_df['Cliente'].astype(str).str.lower().str.contains(query)
        ]

    # --- PANEL DE MÉTRICAS ---
    total_registros = len(st.session_state.data)
    ejecutadas = len(st.session_state.data[st.session_state.data['Estado Verificación'] == 'Ejecutada'])
    pendientes = total_registros - ejecutadas
    porcentaje = (ejecutadas / total_registros * 100) if total_registros > 0 else 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total de Órdenes", total_registros)
    col2.metric("Ejecutadas", ejecutadas)
    col3.metric("Pendientes", pendientes)
    col4.metric("Progreso", f"{porcentaje:.1f}%")

    st.markdown("---")

    # --- TABLA INTERACTIVA DE VERIFICACIÓN ---
    st.subheader("📋 Registro de Dirección y Estado de Órdenes")
    st.caption("Puedes cambiar el estado de **Pendiente** a **Ejecutada** e ingresar observaciones directamente en la tabla.")

    # Columnas principales a mostrar y editar
    columns_to_show = [
        'Número OT', 'Cliente', 'Dirección', 'Tipo OT', 
        'F. Culminación', 'Obs. del Operario', 
        'Estado Verificación', 'Observación Verificador'
    ]

    # Editor de datos interactivo
    edited_df = st.data_editor(
        filtered_df[columns_to_show],
        column_config={
            "Estado Verificación": st.column_config.SelectboxColumn(
                "Estado",
                options=["Pendiente", "Ejecutada"],
                required=True,
            ),
            "Observación Verificador": st.column_config.TextColumn(
                "Obs. Verificador",
                help="Escribe comentarios o aclaraciones de la verificación"
            ),
            "Número OT": st.column_config.TextColumn("N° OT", disabled=True),
            "Cliente": st.column_config.TextColumn("Cliente", disabled=True),
            "Dirección": st.column_config.TextColumn("Dirección", disabled=True),
            "Tipo OT": st.column_config.TextColumn("Tipo OT", disabled=True),
            "F. Culminación": st.column_config.DatetimeColumn("F. Culminación", disabled=True),
            "Obs. del Operario": st.column_config.TextColumn("Obs. Operario", disabled=True),
        },
        disabled=["Número OT", "Cliente", "Dirección", "Tipo OT", "F. Culminación", "Obs. del Operario"],
        hide_index=True,
        use_container_width=True,
        key="data_editor"
    )

    # Actualizar st.session_state cuando el usuario edita la tabla
    if st.button("💾 Guardar Cambios"):
        for idx in edited_df.index:
            st.session_state.data.loc[idx, 'Estado Verificación'] = edited_df.loc[idx, 'Estado Verificación']
            st.session_state.data.loc[idx, 'Observación Verificador'] = edited_df.loc[idx, 'Observación Verificador']
            if edited_df.loc[idx, 'Estado Verificación'] == 'Ejecutada':
                st.session_state.data.loc[idx, 'Fecha Verificación'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        st.success("¡Cambios guardados con éxito!")
        st.rerun()

    # --- DESCARGA DE RESULTADOS ---
    st.markdown("---")
    st.subheader("📥 Exportar Reporte")
    
    @st.cache_data
    def convert_df_to_excel(df_to_save):
        import io
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_to_save.to_excel(writer, sheet_name='Verificación_OT', index=False)
        return output.getvalue()

    excel_data = convert_df_to_excel(st.session_state.data)

    st.download_button(
        label="Download Reporte Actualizado (.xlsx)",
        data=excel_data,
        file_name=f"reporte_verificacion_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("👆 Por favor, carga el archivo `corte y repo.xlsx` usando el panel lateral de la izquierda para comenzar.")
