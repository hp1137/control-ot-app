import streamlit as st
import pandas as pd
from datetime import datetime

# Configuración de la página
st.set_page_config(
    page_title="Control de Órdenes de Trabajo - Corte y Repo",
    page_icon="🛠️",
    layout="wide"
)

st.title("🛠️ Control y Verificación de Órdenes de Trabajo")
st.write("Gestiona la verificación de direcciones ejecutadas y pendientes a partir del registro del servicio.")

# Cargar archivo Excel desde la barra lateral
uploaded_file = st.sidebar.file_uploader("Cargar archivo 'corte y repo.xlsx'", type=["xlsx"])

@st.cache_data
def load_data(file):
    xls = pd.ExcelFile(file)
    df = None
    
    # Buscar automáticamente la hoja y la fila que contiene los encabezados reales
    for sheet in xls.sheet_names:
        temp = pd.read_excel(xls, sheet_name=sheet, header=None)
        
        header_idx = None
        for idx, row in temp.iterrows():
            # Convertir toda la fila a un texto seguro en minúsculas
            row_text = " ".join([str(val).lower() for val in row.values if pd.notna(val)])
            if any(k in row_text for k in ['número ot', 'numero ot', 'dirección', 'direccion', 'cliente']):
                header_idx = idx
                break
                
        if header_idx is not None:
            df = pd.read_excel(xls, sheet_name=sheet, header=header_idx)
            break
            
    # Si no se encontró la palabra clave, cargar la primera pestaña normalmente
    if df is None:
        df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
        
    # Limpiar espacios extra en los nombres de las columnas
    df.columns = [str(c).strip() for c in df.columns]
    
    # Eliminar columnas vacías tipo 'Unnamed'
    df = df.loc[:, ~df.columns.str.contains('^Unnamed', case=False, na=False)]
    
    # Inicializar columnas de verificación si no existen
    if 'Estado Verificación' not in df.columns:
        df['Estado Verificación'] = 'Pendiente'
    if 'Observación Verificador' not in df.columns:
        df['Observación Verificador'] = ''
    if 'Fecha Verificación' not in df.columns:
        df['Fecha Verificación'] = ''
        
    return df

if uploaded_file is not None:
    df = load_data(uploaded_file)

    if 'data' not in st.session_state:
        st.session_state.data = df.copy()

    # --- BARRA LATERAL: FILTROS ---
    st.sidebar.header("🔍 Filtros de Búsqueda")
    
    estado_filtro = st.sidebar.multiselect(
        "Estado de Verificación:",
        options=['Pendiente', 'Ejecutada'],
        default=['Pendiente', 'Ejecutada']
    )
    
    col_tipo_ot = 'Tipo OT' if 'Tipo OT' in st.session_state.data.columns else ('Tipo de trabajo' if 'Tipo de trabajo' in st.session_state.data.columns else None)
    
    if col_tipo_ot:
        tipos_ot = st.sidebar.multiselect(
            "Tipo de OT:",
            options=st.session_state.data[col_tipo_ot].dropna().unique(),
            default=st.session_state.data[col_tipo_ot].dropna().unique()
        )
    else:
        tipos_ot = []

    search_query = st.sidebar.text_input("Buscar por Dirección, OT o Cliente:")

    # Aplicar filtros
    filtered_df = st.session_state.data[
        st.session_state.data['Estado Verificación'].isin(estado_filtro)
    ]
    
    if col_tipo_ot and tipos_ot:
        filtered_df = filtered_df[filtered_df[col_tipo_ot].isin(tipos_ot)]

    if search_query:
        query = search_query.lower()
        cond = pd.Series(False, index=filtered_df.index)
        for col in ['Dirección', 'Número OT', 'Cliente']:
            if col in filtered_df.columns:
                cond = cond | filtered_df[col].astype(str).str.lower().str.contains(query, na=False)
        filtered_df = filtered_df[cond]

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

    # --- TABLA INTERACTIVA ---
    st.subheader("📋 Registro de Dirección y Estado de Órdenes")
    st.caption("Puedes cambiar el estado de **Pendiente** a **Ejecutada** e ingresar observaciones directamente en la tabla.")

    expected_cols = [
        'Número OT', 'Cliente', 'Dirección', 'Tipo OT', 'Tipo de trabajo',
        'F. Culminación', 'Obs. del Operario', 
        'Estado Verificación', 'Observación Verificador'
    ]
    
    columns_to_show = [c for c in expected_cols if c in filtered_df.columns]

    config_dict = {
        "Estado Verificación": st.column_config.SelectboxColumn(
            "Estado",
            options=["Pendiente", "Ejecutada"],
            required=True,
        ),
        "Observación Verificador": st.column_config.TextColumn(
            "Obs. Verificador",
            help="Escribe comentarios o aclaraciones de la verificación"
        )
    }

    disabled_cols = [c for c in columns_to_show if c not in ['Estado Verificación', 'Observación Verificador']]

    edited_df = st.data_editor(
        filtered_df[columns_to_show],
        column_config=config_dict,
        disabled=disabled_cols,
        hide_index=True,
        use_container_width=True,
        key="data_editor"
    )

    if st.button("💾 Guardar Cambios"):
        for idx in edited_df.index:
            st.session_state.data.loc[idx, 'Estado Verificación'] = edited_df.loc[idx, 'Estado Verificación']
            st.session_state.data.loc[idx, 'Observación Verificador'] = edited_df.loc[idx, 'Observación Verificador']
            if edited_df.loc[idx, 'Estado Verificación'] == 'Ejecutada':
                st.session_state.data.loc[idx, 'Fecha Verificación'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        st.success("¡Cambios guardados con éxito!")
        st.rerun()

    # --- EXPORTAR REPORTE ---
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
