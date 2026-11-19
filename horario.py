import streamlit as st
import pandas as pd
import numpy as np
import io
import os
from datetime import datetime

st.set_page_config(page_title="Validación Uonline", layout="wide")
st.title("✅ HORARIOS INGRESADOS EN UONLINE")

# FECHA Y HORA DE CARGA DE ARCHIVOS
file_labs = 'PRACTICAS LAB CICLO ACTUAL DATA ONLY.xlsx'
file_horario = 'HORARIO POR AULA CICLO ACTUAL DATA ONLY.xlsx'

if os.path.exists(file_labs):
    timestamp = os.path.getmtime(file_labs)
    fecha_hora = datetime.fromtimestamp(timestamp).strftime("%d/%m/%Y - %H:%M:%S")
else:
    fecha_hora = "Archivo no encontrado"

col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    
    st.caption("📊 **FUENTE:** REPORTES DE HORARIO DE CLASES POR AULA Y PRÁCTICAS DE LABORATORIO UONLINE")

# ========== FUNCIÓN NORMALIZAR DÍAS ==========
def normalizar_dia(dia):
    """Normaliza abreviaturas de días a formato estándar"""
    if pd.isna(dia):
        return ''
    d = str(dia).strip().upper()
    
    dias_map = {
        'L': 'Lu', 'LU': 'Lu',
        'M': 'Ma', 'MA': 'Ma',
        'X': 'Mi', 'MI': 'Mi',
        'J': 'Ju', 'JU': 'Ju',
        'V': 'Vi', 'VI': 'Vi',
        'S': 'Sa', 'SA': 'Sa',
        'D': 'Do', 'DO': 'Do',
    }
    
    return dias_map.get(d, d)

# ========== FUNCIÓN EXTRAER FRANJA HORARIA ==========
def extraer_franja(horario):
    """Extrae la franja horaria del inicio (ej: '06:00 - 07:30' → '06:')"""
    if pd.isna(horario):
        return ''
    horario = str(horario).strip()
    inicio = horario.split(' - ')[0].strip()
    hora = inicio.split(':')[0]
    return hora + ':'

# ========== LABORATORIOS ==========
labs = pd.read_excel('PRACTICAS LAB CICLO ACTUAL DATA ONLY.xlsx', header=None, skiprows=1)
labs.columns = ['ID', 'ASIGNATURA', 'SECCION', 'CUPO', 'PARIDAD', 'HORARIO', 'DIA', 'AULA']

labs['PARIDAD'] = labs['PARIDAD'].str.strip().str.upper()
labs['HORARIO'] = labs['HORARIO'].str.replace('_x000D_ -', '', regex=False).str.strip()
labs['HORARIO'] = labs['HORARIO'].astype(str)
labs['DIA'] = labs['DIA'].str.strip().apply(normalizar_dia)
labs['AULA'] = labs['AULA'].str.strip()
labs['SECCION'] = labs['SECCION'].astype(str)
labs['CUPO'] = pd.to_numeric(labs['CUPO'], errors='coerce').fillna(0).astype(int)

labs['ASIGNATURA_LIMPIA'] = labs['ASIGNATURA'].apply(lambda x: x.rsplit(' - ', 1)[0] if isinstance(x, str) else x)

labs['TIPO'] = 'LABORATORIO'
labs['DOCENTE'] = ''

labs_consolidated = labs.groupby(['AULA', 'ASIGNATURA_LIMPIA', 'SECCION', 'HORARIO', 'DIA', 'PARIDAD']).agg({
    'CUPO': 'sum'
}).reset_index()

labs_consolidated.rename(columns={'ASIGNATURA_LIMPIA': 'ASIGNATURA'}, inplace=True)
labs_consolidated['TIPO'] = 'LABORATORIO'
labs_consolidated['DOCENTE'] = ''
labs_consolidated['INSCRITOS'] = 0

labs_norm = labs_consolidated[['AULA', 'ASIGNATURA', 'SECCION', 'CUPO', 'INSCRITOS', 'HORARIO', 'DIA', 'DOCENTE', 'PARIDAD', 'TIPO']].copy()

# ========== HORARIOS POR AULA ==========
df = pd.read_excel('HORARIO POR AULA CICLO ACTUAL DATA ONLY.xlsx', header=None)

clases = []
aula_actual = None

def separar_dia_horario(texto, aula):
    """Separa DIA y HORARIO según modalidad (presencial/virtual)"""
    if pd.isna(texto) or pd.isna(aula):
        return '', ''
    
    texto = str(texto).strip()
    aula = str(aula).strip().lower()
    
    if '/' not in texto:
        return '', texto
    
    # Dividir por "/"
    partes = texto.split('/')
    dia = partes[0].strip()
    
    # Detectar si es virtual o presencial
    es_virtual = 'virtual' in aula
    
    if es_virtual:
        # Aula virtual: tomar SEGUNDO horario (después del segundo "/")
        if len(partes) >= 3:
            horario = partes[2].strip()
        else:
            horario = partes[1].strip() if len(partes) > 1 else ''
    else:
        # Aula presencial: tomar PRIMER horario (después del primer "/")
        horario = partes[1].strip() if len(partes) > 1 else ''
    
    return dia, horario

for idx, row in df.iterrows():
    if pd.notna(row[0]) and ' - ' in str(row[0]) and row[1] == 'Cap.':
        aula_actual = str(row[0]).split(' - ')[-1].strip()
    elif aula_actual and pd.notna(row[0]) and row[0] not in ['Semestre', 'Regional']:
        try:
            cupo_val = pd.to_numeric(row[2], errors='coerce')
            cupo_final = int(cupo_val) if pd.notna(cupo_val) else 0
        except:
            cupo_final = 0
        
        dia, horario = separar_dia_horario(row[3], aula_actual)
        dia = normalizar_dia(dia)
        docente = str(row[4]) if pd.notna(row[4]) else ''
        
        clases.append({
            'AULA': aula_actual,
            'ASIGNATURA': str(row[0]),
            'SECCION': str(row[1]),
            'CUPO': 0,
            'INSCRITOS': cupo_final,
            'HORARIO': horario,
            'DIA': dia,
            'DOCENTE': docente,
            'PARIDAD': 'SEMANAL',
            'TIPO': 'TEÓRICA'
        })

clases_df = pd.DataFrame(clases)

# ========== CONSOLIDAR ==========
consolidada = pd.concat([labs_norm, clases_df], ignore_index=True)

# ========== DETECTAR CHOQUES ==========
def detectar_choques(df):
    """Detecta choques considerando AULA + DÍA + HORARIO + PARIDAD"""
    choques = []
    
    grupos = df.groupby(['AULA', 'DIA', 'HORARIO'])
    
    for (aula, dia, horario), grupo in grupos:
        semanal = len(grupo[grupo['PARIDAD'] == 'SEMANAL'])
        par = len(grupo[grupo['PARIDAD'] == 'PAR'])
        impar = len(grupo[grupo['PARIDAD'] == 'IMPAR'])
        
        asignaturas = ' | '.join(grupo['ASIGNATURA'].unique())
        
        if semanal > 1:
            choques.append({
                'AULA': aula,
                'DIA': dia,
                'HORARIO': horario,
                'ASIGNATURA': asignaturas,
                'PARIDAD': 'SEMANAL',
                'CANTIDAD': semanal,
                'TIPO_CHOQUE': '🔴 COINCICENCIA SEMANAL'
            })
        
        if par > 1:
            choques.append({
                'AULA': aula,
                'DIA': dia,
                'HORARIO': horario,
                'ASIGNATURA': asignaturas,
                'PARIDAD': 'PAR',
                'CANTIDAD': par,
                'TIPO_CHOQUE': '🔴 COINCICENCIA PAR'
            })
        
        if impar > 1:
            choques.append({
                'AULA': aula,
                'DIA': dia,
                'HORARIO': horario,
                'ASIGNATURA': asignaturas,
                'PARIDAD': 'IMPAR',
                'CANTIDAD': impar,
                'TIPO_CHOQUE': '🔴 COINCICENCIA IMPAR'
            })
        
        if semanal > 0 and (par > 0 or impar > 0):
            choques.append({
                'AULA': aula,
                'DIA': dia,
                'HORARIO': horario,
                'ASIGNATURA': asignaturas,
                'PARIDAD': 'SEMANAL vs PAR/IMPAR',
                'CANTIDAD': semanal + par + impar,
                'TIPO_CHOQUE': '🔴 CHOQUE PARIDAD'
            })
    
    return pd.DataFrame(choques) if choques else pd.DataFrame()

todos_para_choques = consolidada[consolidada['ASIGNATURA'].notna()].copy()
choques_labs = detectar_choques(todos_para_choques)

# ORDEN DE DÍAS
orden_dia_dict = {'Lu': 0, 'Ma': 1, 'Mi': 2, 'Ju': 3, 'Vi': 4, 'Sa': 5, 'Do': 6}

# MÉTRICAS GENERALES
col1, col2, col3, col4 = st.columns(4)
col1.metric("Total Clases", len(consolidada))
col2.metric("Total Aulas", consolidada['AULA'].nunique())
col3.metric("Total Cupo", consolidada['CUPO'].sum())
col4.metric("Total Inscritos", consolidada['INSCRITOS'].sum())

# TABS
tab1, tab2, tab3, tab4 = st.tabs(["Todas las Clases", "Teóricas", "Laboratorios", "Coincidencias horario"])

with tab1:
    st.subheader("📋 Consolidado (Laboratorios + Teóricas)")
    consolidada_sin_cupo = consolidada.drop(columns=['CUPO', 'INSCRITOS'])
    st.dataframe(consolidada_sin_cupo, use_container_width=True)
    
    st.subheader("📊 Resumen por Aula y Día")
    col1, col2 = st.columns(2)
    
    with col1:
        st.write("**Cantidad de Clases y Laboratorios por Aula**")
        tabla_aula = consolidada[consolidada['ASIGNATURA'].notna()].groupby('AULA').size().reset_index(name='CANTIDAD')
        tabla_aula = tabla_aula.sort_values('AULA')
        st.dataframe(tabla_aula, use_container_width=True)
    
    with col2:
        st.write("**Cantidad de Clases y Laboratorios por Día**")
        tabla_dia = consolidada[consolidada['ASIGNATURA'].notna()].groupby('DIA').size().reset_index(name='CANTIDAD')
        tabla_dia['orden'] = tabla_dia['DIA'].map(orden_dia_dict)
        tabla_dia = tabla_dia.sort_values('orden').drop('orden', axis=1)
        st.dataframe(tabla_dia, use_container_width=True)

with tab2:
    st.subheader("📚 Teóricas")
    teo_tab = consolidada[(consolidada['TIPO'] == 'TEÓRICA') & (consolidada['ASIGNATURA'].notna())].copy()
    teo_tab = teo_tab.reset_index(drop=True)
    teo_tab_mostrar = teo_tab.drop(columns=['CUPO'])
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Clases Teo", len(teo_tab))
    col2.metric("Aulas Teóricas", teo_tab['AULA'].nunique())
    col3.metric("Total Inscritos Teo", teo_tab['INSCRITOS'].sum())
    
    st.dataframe(teo_tab_mostrar, use_container_width=True)
    
    st.subheader("📊 Resumen por Aula y Día (Teóricas)")
    col1, col2 = st.columns(2)
    
    with col1:
        st.write("**Cantidad de Clases por Aula(Incluye virtual)**")
        tabla_aula_teo = teo_tab.groupby('AULA').size().reset_index(name='CANTIDAD')
        tabla_aula_teo = tabla_aula_teo.sort_values('AULA')
        st.dataframe(tabla_aula_teo, use_container_width=True)
    
    with col2:
        st.write("**Cantidad de Clases por Día(Incluye virtuales)**")
        tabla_dia_teo = teo_tab.groupby('DIA').size().reset_index(name='CANTIDAD')
        tabla_dia_teo['orden'] = tabla_dia_teo['DIA'].map(orden_dia_dict)
        tabla_dia_teo = tabla_dia_teo.sort_values('orden').drop('orden', axis=1)
        st.dataframe(tabla_dia_teo, use_container_width=True)
    
    st.subheader("📈 Matriz: Clases por Aula × Franja Horaria (Teóricas)")
    teo_tab_con_franja = teo_tab.copy()
    teo_tab_con_franja['FRANJA'] = teo_tab_con_franja['HORARIO'].apply(extraer_franja)
    matriz_teo_franja = teo_tab_con_franja.groupby(['AULA', 'FRANJA']).size().unstack(fill_value=0)
    matriz_teo_franja = matriz_teo_franja.sort_index()
    st.dataframe(matriz_teo_franja, use_container_width=True)
    
    st.subheader("📈 Matriz: Día × Hora (Aula)")
    teo_tab_filtrada = teo_tab[teo_tab['DIA'].notna() & (teo_tab['DIA'] != '')].copy()
    matriz_teo = teo_tab_filtrada.pivot_table(
        index=['DIA', 'HORARIO'],
        columns='AULA',
        values='ASIGNATURA',
        aggfunc=lambda x: ' | '.join(x)
    )
    
    indices_ordenados = sorted(
        matriz_teo.index,
        key=lambda x: (orden_dia_dict.get(x[0], 999), x[1])
    )
    matriz_teo = matriz_teo.loc[indices_ordenados]
    
    st.dataframe(matriz_teo, use_container_width=True)

with tab3:
    st.subheader("🧪 Laboratorios")
    labs_tab = consolidada[(consolidada['TIPO'] == 'LABORATORIO') & (consolidada['ASIGNATURA'].notna())].copy()
    labs_tab = labs_tab.reset_index(drop=True)
    labs_tab_mostrar = labs_tab.drop(columns=['INSCRITOS'])
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Clases Lab", len(labs_tab))
    col2.metric("Aulas Laboratorios", labs_tab['AULA'].nunique())
    col3.metric("Total Cupo Lab", labs_tab['CUPO'].sum())
    
    st.dataframe(labs_tab_mostrar, use_container_width=True)
    
    st.subheader("📊 Resumen por Aula y Día (Laboratorios)")
    col1, col2 = st.columns(2)
    
    with col1:
        st.write("**Cantidad de Clases por Aula**")
        tabla_aula_labs = labs_tab.groupby('AULA').size().reset_index(name='CANTIDAD')
        tabla_aula_labs = tabla_aula_labs.sort_values('AULA')
        st.dataframe(tabla_aula_labs, use_container_width=True)
    
    with col2:
        st.write("**Cantidad de Clases por Día**")
        tabla_dia_labs = labs_tab.groupby('DIA').size().reset_index(name='CANTIDAD')
        tabla_dia_labs['orden'] = tabla_dia_labs['DIA'].map(orden_dia_dict)
        tabla_dia_labs = tabla_dia_labs.sort_values('orden').drop('orden', axis=1)
        st.dataframe(tabla_dia_labs, use_container_width=True)
    
    st.subheader("📈 Matriz: Clases por Aula × Franja Horaria (Laboratorios)")
    labs_tab_con_franja = labs_tab.copy()
    labs_tab_con_franja['FRANJA'] = labs_tab_con_franja['HORARIO'].apply(extraer_franja)
    matriz_labs_franja = labs_tab_con_franja.groupby(['AULA', 'FRANJA']).size().unstack(fill_value=0)
    matriz_labs_franja = matriz_labs_franja.sort_index()
    st.dataframe(matriz_labs_franja, use_container_width=True)
    
    st.subheader("📈 Matriz: Día × Hora (Aula + Paridad)")
    labs_tab_filtrada = labs_tab[labs_tab['DIA'].notna() & (labs_tab['DIA'] != '')].copy()
    matriz_labs = labs_tab_filtrada.pivot_table(
        index=['DIA', 'HORARIO'],
        columns=['AULA', 'PARIDAD'],
        values='ASIGNATURA',
        aggfunc=lambda x: ' | '.join(x)
    )
    
    indices_ordenados = sorted(
        matriz_labs.index,
        key=lambda x: (orden_dia_dict.get(x[0], 999), x[1])
    )
    matriz_labs = matriz_labs.loc[indices_ordenados]
    
    st.dataframe(matriz_labs, use_container_width=True)

with tab4:
    st.subheader("⚠️ COINCIDENCIAS DETECTADAS (LABORATORIOS + TEÓRICAS)")
    
    if len(choques_labs) > 0:
        st.error(f"Se encontraron {len(choques_labs)} coincidencias")
        st.dataframe(choques_labs, use_container_width=True)
    else:
        st.success("No hay conflictos de horarios")

# DESCARGAR
st.subheader("📥 Descargar")

col1, col2 = st.columns(2)

with col1:
    csv = consolidada.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="Descargar Consolidado (CSV)",
        data=csv,
        file_name="validacion_uonline.csv",
        mime="text/csv"
    )

with col2:
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        consolidada.to_excel(writer, sheet_name='Consolidado', index=False)
        labs_tab_mostrar.to_excel(writer, sheet_name='Laboratorios', index=False)
        teo_tab_mostrar.to_excel(writer, sheet_name='Teóricas', index=False)
        if len(choques_labs) > 0:
            choques_labs.to_excel(writer, sheet_name='Coincidencias', index=False)
        tabla_aula.to_excel(writer, sheet_name='Resumen Aulas Total', index=False)
        tabla_aula_labs.to_excel(writer, sheet_name='Resumen Aulas Labs', index=False)
        tabla_aula_teo.to_excel(writer, sheet_name='Resumen Aulas Teo', index=False)
        matriz_labs.to_excel(writer, sheet_name='Matriz Labs Dia-Hora-Paridad')
        matriz_teo.to_excel(writer, sheet_name='Matriz Teo Dia-Hora')
    
    excel_buffer.seek(0)
    st.download_button(
        label="Descargar Consolidado (Excel)",
        data=excel_buffer.getvalue(),
        file_name="validacion_uonline.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
