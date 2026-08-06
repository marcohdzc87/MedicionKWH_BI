import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, date, timedelta
import calendar
import io
from supabase import create_client

st.set_page_config(page_title="Gestor Energético CFE & Solar", page_icon="⚡", layout="wide")

@st.cache_resource
def init_supabase():
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])

try:
    supabase = init_supabase()
except Exception as e:
    st.error("Error al conectar con Supabase. Verifica tus credenciales en los secretos de Streamlit.")

# -----------------------------------------------------------------------------
# Lógica de Cálculo de Días y Fechas de Ciclo
# -----------------------------------------------------------------------------
def obtener_fechas_ciclo(fecha_actual, dia_corte, es_bimestral=True):
    """
    Calcula el inicio y fin del ciclo de facturación según el día de corte y periodicidad.
    """
    año = fecha_actual.year
    mes = fecha_actual.month
    
    # Ajustar día de corte si el mes actual tiene menos días
    dias_en_mes = calendar.monthrange(año, mes)[1]
    dia_efectivo = min(dia_corte, dias_en_mes)
    
    if fecha_actual.day >= dia_efectivo:
        # El ciclo actual comenzó este mes en el día de corte
        inicio_ciclo = date(año, mes, dia_efectivo)
        meses_a_sumar = 2 if es_bimestral else 1
        
        # Calcular fecha fin
        mes_fin = mes + meses_a_sumar
        año_fin = año
        if mes_fin > 12:
            mes_fin -= 12
            año_fin += 1
            
        dias_en_mes_fin = calendar.monthrange(año_fin, mes_fin)[1]
        fin_ciclo = date(año_fin, mes_fin, min(dia_corte, dias_en_mes_fin))
    else:
        # El ciclo actual comenzó en un periodo anterior
        meses_a_restar = 2 if es_bimestral else 1
        mes_inicio = mes - meses_a_restar
        año_inicio = año
        if mes_inicio < 1:
            mes_inicio += 12
            año_inicio -= 1
            
        dias_en_mes_inicio = calendar.monthrange(año_inicio, mes_inicio)[1]
        inicio_ciclo = date(año_inicio, mes_inicio, min(dia_corte, dias_en_mes_inicio))
        fin_ciclo = date(año, mes, dia_efectivo)
        
    return inicio_ciclo, fin_ciclo

# -----------------------------------------------------------------------------
# Lógica de Cálculo Dinámico de Tarifa
# -----------------------------------------------------------------------------
def calcular_detalle_factura(kwh_netos_periodo, credito_disponible, tarifa):
    balance_kwh = kwh_netos_periodo - credito_disponible
    
    if balance_kwh <= 0:
        kwh_facturables = 0.0
        nuevo_credito = abs(balance_kwh)
        subtotal_energia = float(tarifa.get("cargo_fijo", 0.0))
        desglose_rangos = {"Básico": 0.0, "Intermedio": 0.0, "Excedente": 0.0}
    else:
        kwh_facturables = balance_kwh
        nuevo_credito = 0.0
        
        l_basico = float(tarifa.get("limite_basico", 150.0))
        l_inter = float(tarifa.get("limite_intermedio", 280.0))
        
        p_basico = float(tarifa.get("precio_basico", 1.087))
        p_inter = float(tarifa.get("precio_intermedio", 1.320))
        p_exced = float(tarifa.get("precio_excedente", 3.861))
        
        costo_b, costo_i, costo_e = 0.0, 0.0, 0.0
        
        if kwh_facturables <= l_basico:
            costo_b = kwh_facturables * p_basico
        elif kwh_facturables <= l_inter:
            costo_b = l_basico * p_basico
            costo_i = (kwh_facturables - l_basico) * p_inter
        else:
            costo_b = l_basico * p_basico
            costo_i = (l_inter - l_basico) * p_inter
            costo_e = (kwh_facturables - l_inter) * p_exced
            
        subtotal_energia = float(tarifa.get("cargo_fijo", 0.0)) + costo_b + costo_i + costo_e
        desglose_rangos = {
            "Básico": round(costo_b, 2),
            "Intermedio": round(costo_i, 2),
            "Excedente": round(costo_e, 2)
        }

    dap = float(tarifa.get("cuota_fija_dap", 0.0)) + (subtotal_energia * (float(tarifa.get("porcentaje_dap", 0.0)) / 100.0))
    subtotal_con_dap = subtotal_energia + dap
    iva = subtotal_con_dap * (float(tarifa.get("porcentaje_iva", 16.0)) / 100.0)
    total_factura = subtotal_con_dap + iva
    
    return {
        "kwh_facturables": round(kwh_facturables, 1),
        "nuevo_credito": round(nuevo_credito, 1),
        "subtotal_energia": round(subtotal_energia, 2),
        "desglose_rangos": desglose_rangos,
        "dap": round(dap, 2),
        "iva": round(iva, 2),
        "total": round(total_factura, 2)
    }

def generar_excel(df_export):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Historial_Lecturas')
    return output.getvalue()

# -----------------------------------------------------------------------------
# Interfaz de Usuario
# -----------------------------------------------------------------------------
st.title("⚡ Gestor Energético Bidireccional CFE & Solar")

pestaña1, pestaña2, pestaña3, pestaña4 = st.tabs([
    "📊 Dashboard & Proyección Detallada", 
    "➕ Capturar Lectura Libre", 
    "📥 Descargar Exportación",
    "⚙️ Administración de Tarifas y Ciclo"
])

# --- PESTAÑA 1: DASHBOARD ---
with pestaña1:
    try:
        res_lecturas = supabase.table("lecturas").select("*, tarifas(*)").order("fecha_corte", desc=False).execute()
        df = pd.DataFrame(res_lecturas.data) if res_lecturas.data else pd.DataFrame()
    except Exception:
        df = pd.DataFrame()

    if not df.empty and len(df) >= 2:
        df["fecha_corte"] = pd.to_datetime(df["fecha_corte"]).dt.date
        lectura_anterior = df.iloc[-2]
        lectura_actual = df.iloc[-1]
        
        fecha_act = lectura_actual["fecha_corte"]
        tarifa_act = lectura_actual.get("tarifas", {})
        
        if isinstance(tarifa_act, dict) and tarifa_act:
            # Datos de configuración del ciclo CFE
            dia_corte_cfe = int(tarifa_act.get("dia_corte_cfe", 15))
            es_bimestral = tarifa_act.get("es_bimestral", True)
            
            # Cálculo de fechas de ciclo
            inicio_ciclo, fin_ciclo = obtener_fechas_ciclo(fecha_act, dia_corte_cfe, es_bimestral)
            dias_totales_ciclo = (fin_ciclo - inicio_ciclo).days
            dias_transcurridos = (fecha_act - lectura_anterior["fecha_corte"]).days
            dias_transcurridos = max(1, dias_transcurridos)
            
            dias_restantes = (fin_ciclo - fecha_act).days
            dias_restantes = max(0, dias_restantes)
            
            # Consumos y promedios
            cons_medido = lectura_actual["lectura_cons_kwh"] - lectura_anterior["lectura_cons_kwh"]
            inyec_medida = lectura_actual["lectura_inyec_kwh"] - lectura_anterior["lectura_inyec_kwh"]
            neto_medido = cons_medido - inyec_medida
            
            gen_solar = lectura_actual["generacion_solar_total_kwh"] - lectura_anterior["generacion_solar_total_kwh"]
            
            cons_diario = cons_medido / dias_transcurridos
            inyec_diaria = inyec_medida / dias_transcurridos
            neto_diario = neto_medido / dias_transcurridos
            
            # Proyecciones
            kwh_restantes_proyectados = neto_diario * dias_restantes
            neto_proyectado_total = neto_medido + kwh_restantes_proyectados
            
            credito_previo = lectura_actual.get("credito_anterior_kwh", 0.0)
            
            calc_actual = calcular_detalle_factura(neto_medido, credito_previo, tarifa_act)
            calc_proyectado = calcular_detalle_factura(neto_proyectado_total, credito_previo, tarifa_act)
            costo_estimado_dias_restantes = max(0.0, calc_proyectado["total"] - calc_actual["total"])
            
            # Ahorro solar estimado
            # Consumo total real del hogar (Autoconsumo + Consumo CFE)
            autoconsumo_solar = max(0.0, gen_solar - inyec_medida)
            consumo_real_total = cons_medido + autoconsumo_solar
            calc_sin_paneles = calcular_detalle_factura(consumo_real_total, 0.0, tarifa_act)
            ahorro_solar = max(0.0, calc_sin_paneles["total"] - calc_actual["total"])
            
            # --- INTERFAZ DEL DASHBOARD ---
            st.subheader(f"📅 Estado al {fecha_act.strftime('%d/%m/%Y')}")
            
            # Progreso del periodo
            pct_tiempo = min(1.0, (dias_totales_ciclo - dias_restantes) / dias_totales_ciclo)
            st.progress(pct_tiempo, text=f"Progreso del Ciclo CFE: **{dias_totales_ciclo - dias_restantes} de {dias_totales_ciclo} días transcurridos** (Próximo corte: {fin_ciclo.strftime('%d/%m/%Y')})")
            
            st.divider()
            
            # Fila 1: Métricas de Consumo e Inyección
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Tomado de CFE", f"{cons_medido:,.1f} kWh", f"{cons_diario:.1f} kWh/día")
            col2.metric("Inyectado a CFE", f"{inyec_medida:,.1f} kWh", f"{inyec_diaria:.1f} kWh/día")
            col3.metric("kWh Netos a Pagar Hoy", f"{calc_actual['kwh_facturables']:,.1f} kWh", help="Consumo neto descontando inyección y crédito acumulado")
            col4.metric("Recibo Estimado Al Día", f"${calc_actual['total']:,.2f} MXN")
            
            st.divider()
            
            # Fila 2: Métricas Proyectadas ("Lo que falta")
            st.markdown("### 🔮 Proyección al Cierre de Corte CFE")
            p1, p2, p3, p4 = st.columns(4)
            p1.metric("Días Faltantes para Corte", f"{dias_restantes} días")
            p2.metric("kWh Estimados por Consumir", f"{kwh_restantes_proyectados:,.1f} kWh", f"{neto_diario:.1f} kWh/día neto")
            p3.metric("Costo Estimado Días Restantes", f"${costo_estimado_dias_restantes:,.2f} MXN")
            p4.metric("PROYECCIÓN TOTAL RECIBO", f"${calc_proyectado['total']:,.2f} MXN")
            
            st.divider()
            
            col_graf, col_info = st.columns([2, 1])
            
            with col_graf:
                st.subheader("Historial de Lecturas del Medidor")
                fig = px.line(
                    df, 
                    x="fecha_corte", 
                    y=["lectura_cons_kwh", "lectura_inyec_kwh"],
                    markers=True,
                    labels={"value": "Lectura Acumulada kWh", "variable": "Concepto", "fecha_corte": "Fecha"},
                    title="Evolución de Lecturas"
                )
                st.plotly_chart(fig, use_container_width=True)
                
            with col_info:
                st.subheader("Desglose Financiero y Ahorro")
                st.write(f"• **Crédito en Bolsa Previo:** {credito_previo:.1f} kWh")
                
                if calc_actual['nuevo_credito'] > 0:
                    st.success(f"🎉 Tienes **{calc_actual['nuevo_credito']} kWh** de saldo a favor acumulado en la bolsa.")
                
                st.write("---")
                st.markdown("#### **Desglose Actual Al Día**")
                st.write(f"• Energía Subtotal: ${calc_actual['subtotal_energia']:,.2f}")
                st.write(f"  - Escalón Básico: ${calc_actual['desglose_rangos']['Básico']:,.2f}")
                st.write(f"  - Escalón Intermedio: ${calc_actual['desglose_rangos']['Intermedio']:,.2f}")
                st.write(f"  - Escalón Excedente: ${calc_actual['desglose_rangos']['Excedente']:,.2f}")
                st.write(f"• DAP: ${calc_actual['dap']:,.2f}")
                st.write(f"• IVA: ${calc_actual['iva']:,.2f}")
                st.markdown(f"### **Total Actual: ${calc_actual['total']:,.2f} MXN**")
                
                st.write("---")
                st.info(f"☀️ **Ahorro Solar del Periodo:** Gracias a tus paneles has ahorrado **${ahorro_solar:,.2f} MXN** comparado con la tarifa normal sin solar.")
        else:
            st.warning("La última lectura registrada no tiene asociada una tarifa válida.")

    elif not df.empty and len(df) == 1:
        st.info("Has registrado tu primera lectura inicial. Agrega una segunda lectura para ver tendencias y proyecciones de días restantes.")
    else:
        st.info("👋 **¡Bienvenido!** Dirígete a la pestaña **⚙️ Administración de Tarifas y Ciclo** para configurar tus parámetros de corte.")

# --- PESTAÑA 2: CAPTURAR LECTURA ---
with pestaña2:
    st.subheader("➕ Capturar Lectura del Medidor (Cualquier día)")
    
    try:
        res_tarifas = supabase.table("tarifas").select("*").execute()
        tarifas_list = res_tarifas.data if res_tarifas.data else []
    except Exception:
        tarifas_list = []
        
    tarifas_dict = {t["nombre"]: t["id"] for t in tarifas_list} if tarifas_list else {}
    
    if tarifas_dict:
        try:
            res_ultima = supabase.table("lecturas").select("*").order("fecha_corte", desc=True).limit(1).execute()
            credito_anterior = res_ultima.data[0]["credito_remanente_kwh"] if res_ultima.data else 0.0
        except Exception:
            res_ultima = None
            credito_anterior = 0.0
        
        with st.form("form_registro_libre", clear_on_submit=True):
            fecha = st.date_input("Fecha de Toma de Lectura", value=datetime.now())
            
            c_a, c_b = st.columns(2)
            cons_kwh = c_a.number_input("Lectura Medidor - Consumo (kWh / Código 1.8.0)", min_value=0.0, step=0.1)
            inyec_kwh = c_b.number_input("Lectura Medidor - Inyección (kWh / Código 2.8.0)", min_value=0.0, step=0.1)
            
            solar_kwh = st.number_input("Generación Solar Acumulada Inversor (kWh)", min_value=0.0, step=0.1)
            tarifa_sel = st.selectbox("Tarifa Aplicable", list(tarifas_dict.keys()))
            
            es_cierre = st.checkbox("Marcar como Cierre Oficial de Bimestre / Recibo CFE")
            notas = st.text_input("Notas (Ej. 'Lectura semanal', 'Corte oficial CFE')")
            
            if st.form_submit_button("Guardar Registro"):
                if res_ultima and res_ultima.data:
                    u = res_ultima.data[0]
                    c_periodo = cons_kwh - u["lectura_cons_kwh"]
                    i_periodo = inyec_kwh - u["lectura_inyec_kwh"]
                    neto = c_periodo - i_periodo
                    
                    t_obj = [t for t in tarifas_list if t["id"] == tarifas_dict[tarifa_sel]][0]
                    res_calc = calcular_detalle_factura(neto, credito_anterior, t_obj)
                    
                    nuevo_rem = res_calc["nuevo_credito"] if es_cierre else credito_anterior
                else:
                    nuevo_rem = 0.0
                
                data = {
                    "fecha_corte": str(fecha),
                    "lectura_cons_kwh": cons_kwh,
                    "lectura_inyec_kwh": inyec_kwh,
                    "generacion_solar_total_kwh": solar_kwh,
                    "tarifa_id": tarifas_dict[tarifa_sel],
                    "credito_anterior_kwh": credito_anterior,
                    "credito_remanente_kwh": nuevo_rem,
                    "es_cierre_periodo": es_cierre,
                    "notas": notas
                }
                supabase.table("lecturas").insert(data).execute()
                st.success(f"¡Lectura registrada para el {fecha}!")
                st.rerun()
    else:
        st.warning("Primero debes registrar al menos una tarifa.")

# --- PESTAÑA 3: EXPORTACIÓN ---
with pestaña3:
    st.subheader("📥 Exportación de Lecturas e Historial de Crédito")
    try:
        res_export = supabase.table("lecturas").select("*, tarifas(nombre)").order("fecha_corte", desc=False).execute()
        df_exp = pd.DataFrame(res_export.data) if res_export.data else pd.DataFrame()
    except Exception:
        df_exp = pd.DataFrame()
    
    if not df_exp.empty:
        if "tarifas" in df_exp.columns:
            df_exp["nombre_tarifa"] = df_exp["tarifas"].apply(lambda x: x.get("nombre") if isinstance(x, dict) else "")
            df_exp.drop(columns=["tarifas"], inplace=True)
            
        st.dataframe(df_exp, use_container_width=True)
        col_csv, col_excel = st.columns(2)
        
        csv_data = df_exp.to_csv(index=False).encode('utf-8')
        col_csv.download_button(
            label="📄 Descargar en formato CSV",
            data=csv_data,
            file_name=f"historial_lecturas_energia_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True
        )
        
        try:
            excel_data = generar_excel(df_exp)
            col_excel.download_button(
                label="📊 Descargar en formato Excel (.xlsx)",
                data=excel_data,
                file_name=f"historial_lecturas_energia_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        except Exception:
            col_excel.info("Asegúrate de incluir openpyxl en requirements.txt para descargar en formato Excel.")
    else:
        st.info("No hay lecturas registradas para exportar.")

# --- PESTAÑA 4: ADMINISTRACIÓN DE TARIFAS Y CICLO ---
with pestaña4:
    st.subheader("⚙️ Configurar Tarifas y Ciclo de Facturación CFE")
    
    try:
        res_t = supabase.table("tarifas").select("*").order("id", desc=False).execute()
        lista_tarifas = res_t.data if res_t.data else []
    except Exception:
        lista_tarifas = []
        
    subtab1, subtab2 = st.tabs(["✏️ Tarifas Guardadas (Ver / Editar)", "➕ Crear Nueva Tarifa"])
    
    # 1. VER / EDITAR TARIFAS Y CICLO
    with subtab1:
        if lista_tarifas:
            opciones_t = {t["nombre"]: t for t in lista_tarifas}
            tarifa_seleccionada_nombre = st.selectbox("Selecciona una tarifa para editar:", list(opciones_t.keys()))
            t_sel = opciones_t[tarifa_seleccionada_nombre]
            
            with st.form("form_edit_tarifa"):
                st.markdown(f"#### Editando: **{t_sel['nombre']}**")
                
                edit_nombre = st.text_input("Nombre de la Tarifa", value=t_sel["nombre"])
                
                st.markdown("#### Configuración de Corte CFE")
                c_corte1, c_corte2 = st.columns(2)
                edit_dia_corte = c_corte1.number_input("Día de Corte en el Mes (1-31)", min_value=1, max_value=31, value=int(t_sel.get("dia_corte_cfe", 15)))
                edit_es_bimestral = c_corte2.selectbox("Periodo de Facturación", ["Bimestral (60 días)", "Mensual (30 días)"], index=0 if t_sel.get("es_bimestral", True) else 1)
                
                c1, c2 = st.columns(2)
                edit_cargo = c1.number_input("Cargo Fijo ($)", min_value=0.0, value=float(t_sel.get("cargo_fijo", 0.0)))
                edit_iva = c2.number_input("IVA (%)", min_value=0.0, value=float(t_sel.get("porcentaje_iva", 16.0)))
                
                st.markdown("#### Rangos de Consumo (kWh) y Precios ($/kWh)")
                r1_col, r2_col, r3_col = st.columns(3)
                edit_lim_b = r1_col.number_input("Límite Básico (kWh)", min_value=1.0, value=float(t_sel.get("limite_basico", 150.0)))
                edit_p_b = r1_col.number_input("Precio Básico ($)", min_value=0.0, value=float(t_sel.get("precio_basico", 1.087)))
                
                edit_lim_i = r2_col.number_input("Límite Intermedio Acumulado (kWh)", min_value=edit_lim_b, value=float(t_sel.get("limite_intermedio", 280.0)))
                edit_p_i = r2_col.number_input("Precio Intermedio ($)", min_value=0.0, value=float(t_sel.get("precio_intermedio", 1.320)))
                
                edit_p_e = r3_col.number_input("Precio Excedente ($)", min_value=0.0, value=float(t_sel.get("precio_excedente", 3.861)))
                
                st.markdown("#### Alumbrado Público (DAP)")
                dap_col1, dap_col2 = st.columns(2)
                edit_dap_pct = dap_col1.number_input("DAP en Porcentaje (%)", min_value=0.0, value=float(t_sel.get("porcentaje_dap", 0.0)))
                edit_dap_fijo = dap_col2.number_input("DAP en Cuota Fija ($)", min_value=0.0, value=float(t_sel.get("cuota_fija_dap", 0.0)))
                
                if st.form_submit_button("💾 Guardar Cambios"):
                    data_update = {
                        "nombre": edit_nombre,
                        "dia_corte_cfe": edit_dia_corte,
                        "es_bimestral": True if edit_es_bimestral == "Bimestral (60 días)" else False,
                        "cargo_fijo": edit_cargo,
                        "limite_basico": edit_lim_b,
                        "precio_basico": edit_p_b,
                        "limite_intermedio": edit_lim_i,
                        "precio_intermedio": edit_p_i,
                        "precio_excedente": edit_p_e,
                        "porcentaje_dap": edit_dap_pct,
                        "cuota_fija_dap": edit_dap_fijo,
                        "porcentaje_iva": edit_iva
                    }
                    try:
                        supabase.table("tarifas").update(data_update).eq("id", t_sel["id"]).execute()
                        st.success(f"¡Tarifa '{edit_nombre}' actualizada correctamente!")
                        st.rerun()
                    except Exception as err:
                        st.error(f"Error al actualizar la tarifa: {err}")
        else:
            st.info("No hay tarifas guardadas para editar.")

    # 2. CREAR NUEVA TARIFA
    with subtab2:
        with st.form("form_nueva_tarifa", clear_on_submit=True):
            nombre_tarifa = st.text_input("Nombre de la Tarifa", placeholder="Ej. Tarifa Residencial / Verano")
            
            st.markdown("#### Configuración de Corte CFE")
            c_corte1, c_corte2 = st.columns(2)
            dia_corte = c_corte1.number_input("Día de Corte en el Mes (1-31)", min_value=1, max_value=31, value=15)
            es_bimestral = c_corte2.selectbox("Periodo de Facturación", ["Bimestral (60 días)", "Mensual (30 días)"], index=0)
            
            c1, c2 = st.columns(2)
            cargo_fijo = c1.number_input("Cargo Fijo ($)", min_value=0.0, value=0.0)
            iva_pct = c2.number_input("IVA (%)", min_value=0.0, value=16.0)
            
            st.markdown("#### Rangos de Consumo (kWh) y Precios ($/kWh)")
            r1_col, r2_col, r3_col = st.columns(3)
            lim_basico = r1_col.number_input("Límite Básico (kWh)", min_value=1.0, value=150.0)
            p_basico = r1_col.number_input("Precio Básico ($)", min_value=0.0, value=1.087)
            
            lim_inter = r2_col.number_input("Límite Intermedio Acumulado (kWh)", min_value=lim_basico, value=280.0)
            p_inter = r2_col.number_input("Precio Intermedio ($)", min_value=0.0, value=1.320)
            
            p_exced = r3_col.number_input("Precio Excedente ($)", min_value=0.0, value=3.861)
            
            dap_col1, dap_col2 = st.columns(2)
            dap_pct = dap_col1.number_input("DAP en Porcentaje (%)", min_value=0.0, value=0.0)
            dap_fijo = dap_col2.number_input("DAP en Cuota Fija ($)", min_value=0.0, value=0.0)
            
            if st.form_submit_button("➕ Registrar Nueva Tarifa"):
                data_tarifa = {
                    "nombre": nombre_tarifa,
                    "dia_corte_cfe": int(dia_corte),
                    "es_bimestral": True if es_bimestral == "Bimestral (60 días)" else False,
                    "cargo_fijo": float(cargo_fijo),
                    "limite_basico": float(lim_basico),
                    "precio_basico": float(p_basico),
                    "limite_intermedio": float(lim_inter),
                    "precio_intermedio": float(p_inter),
                    "precio_excedente": float(p_exced),
                    "porcentaje_dap": float(dap_pct),
                    "cuota_fija_dap": float(dap_fijo),
                    "porcentaje_iva": float(iva_pct)
                }
                try:
                    supabase.table("tarifas").insert(data_tarifa).execute()
                    st.success(f"Tarifa '{nombre_tarifa}' guardada con éxito.")
                    st.rerun()
                except Exception as err:
                    st.error(f"Error al guardar la tarifa: {err}")
