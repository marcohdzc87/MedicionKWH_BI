import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import io
from supabase import create_client

st.set_page_config(page_title="Gestor Energético CFE & Solar", page_icon="⚡", layout="wide")

@st.cache_resource
def init_supabase():
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])

supabase = init_supabase()

# -----------------------------------------------------------------------------
# Lógica de Cálculo de Tarifa y Proyección
# -----------------------------------------------------------------------------
def calcular_detalle_factura(kwh_netos_periodo, credito_disponible, tarifa):
    balance_kwh = kwh_netos_periodo - credito_disponible
    
    if balance_kwh <= 0:
        kwh_facturables = 0.0
        nuevo_credito = abs(balance_kwh)
        subtotal_energia = tarifa["cargo_fijo"]
        desglose_rangos = {"Básico": 0.0, "Intermedio": 0.0, "Excedente": 0.0}
    else:
        kwh_facturables = balance_kwh
        nuevo_credito = 0.0
        
        l_basico = tarifa["limite_basico"]
        l_inter = tarifa["limite_intermedio"]
        p_basico = tarifa["precio_basico"]
        p_inter = tarifa["precio_intermedio"]
        p_exced = tarifa["precio_excedente"]
        
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
            
        subtotal_energia = tarifa["cargo_fijo"] + costo_b + costo_i + costo_e
        desglose_rangos = {
            "Básico": round(costo_b, 2),
            "Intermedio": round(costo_i, 2),
            "Excedente": round(costo_e, 2)
        }

    dap = tarifa["cuota_fija_dap"] + (subtotal_energia * (tarifa["porcentaje_dap"] / 100.0))
    subtotal_con_dap = subtotal_energia + dap
    iva = subtotal_con_dap * (tarifa["porcentaje_iva"] / 100.0)
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
    processed_data = output.getvalue()
    return processed_data

# -----------------------------------------------------------------------------
# Interfaz de Usuario
# -----------------------------------------------------------------------------
st.title("⚡ Gestor Energético Bidireccional CFE & Solar")

pestaña1, pestaña2, pestaña3, pestaña4 = st.tabs([
    "📊 Dashboard & Proyección", 
    "➕ Capturar Lectura Libre", 
    "📥 Descargar Exportación",
    "⚙️ Tarifas"
])

# --- PESTAÑA 1: DASHBOARD ---
    with pestaña1:

    # Reemplaza la línea 92 por este bloque protegido:
    try:
    res_lecturas = supabase.table("lecturas").select("*, tarifas(*)").order("fecha_corte", desc=False).execute()
    df = pd.DataFrame(res_lecturas.data)
    except Exception:
    df = pd.DataFrame()
    
    df = pd.DataFrame(res_lecturas.data)
    
    if len(df) >= 2:
        df["fecha_corte"] = pd.to_datetime(df["fecha_corte"])
        lectura_anterior = df.iloc[-2]
        lectura_actual = df.iloc[-1]
        
        dias_transcurridos = (lectura_actual["fecha_corte"] - lectura_anterior["fecha_corte"]).days
        dias_transcurridos = max(1, dias_transcurridos)
        
        cons_medido = lectura_actual["lectura_cons_kwh"] - lectura_anterior["lectura_cons_kwh"]
        inyec_medida = lectura_actual["lectura_inyec_kwh"] - lectura_anterior["lectura_inyec_kwh"]
        neto_medido = cons_medido - inyec_medida
        
        cons_diario = cons_medido / dias_transcurridos
        inyec_diaria = inyec_medida / dias_transcurridos
        neto_diario = neto_medido / dias_transcurridos
        
        DIAS_PERIODO = 60
        neto_proyectado = neto_diario * DIAS_PERIODO
        
        credito_previo = lectura_actual.get("credito_anterior_kwh", 0.0)
        tarifa_act = lectura_actual["tarifas"]
        
        calc_actual = calcular_detalle_factura(neto_medido, credito_previo, tarifa_act)
        calc_proyectado = calcular_detalle_factura(neto_proyectado, credito_previo, tarifa_act)
        
        st.subheader(f"📅 Estado al {lectura_actual['fecha_corte'].strftime('%d/%m/%Y')} ({dias_transcurridos} días medidos)")
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Consumido CFE", f"{cons_medido:,.1f} kWh", f"{cons_diario:.1f} kWh/día")
        c2.metric("Inyectado CFE", f"{inyec_medida:,.1f} kWh", f"{inyec_diaria:.1f} kWh/día")
        c3.metric("Monto Actual Al Día", f"${calc_actual['total']:,.2f} MXN")
        c4.metric("Proyección Fin de Bimestre", f"${calc_proyectado['total']:,.2f} MXN")
        
        st.divider()
        
        col_graf, col_info = st.columns([2, 1])
        
        with col_graf:
            st.subheader("Historial de Lecturas de Medidor")
            fig = px.line(
                df, 
                x="fecha_corte", 
                y=["lectura_cons_kwh", "lectura_inyec_kwh"],
                markers=True,
                labels={"value": "Lectura Acumulada kWh", "variable": "Concepto", "fecha_corte": "Fecha"},
                title="Evolución de Lecturas del Medidor Bidireccional"
            )
            st.plotly_chart(fig, use_container_width=True)
            
        with col_info:
            st.subheader("Detalle del Periodo Medido")
            st.write(f"• **Saldo Crédito Anterior:** {credito_previo:.1f} kWh")
            if calc_actual['nuevo_credito'] > 0:
                st.success(f"A la fecha tienes **{calc_actual['nuevo_credito']} kWh** a favor en tu bolsa de crédito.")
            else:
                st.warning(f"Llevas **{calc_actual['kwh_facturables']} kWh** netos a pagar.")
                
            st.write("---")
            st.markdown("### **Desglose Actual**")
            st.write(f"• Energía: ${calc_actual['subtotal_energia']:,.2f}")
            st.write(f"• DAP: ${calc_actual['dap']:,.2f}")
            st.write(f"• IVA: ${calc_actual['iva']:,.2f}")
            st.markdown(f"### **Total hoy: ${calc_actual['total']:,.2f} MXN**")

    else:
        st.info("Ingresa al menos 2 lecturas para calcular consumos por día y estimaciones.")

# --- PESTAÑA 2: CAPTURAR LECTURA ---
with pestaña2:
    st.subheader("➕ Capturar Lectura del Medidor (Cualquier día)")
    res_tarifas = supabase.table("tarifas").select("*").execute()
    tarifas_dict = {t["nombre"]: t["id"] for t in res_tarifas.data} if res_tarifas.data else {}
    
    if tarifas_dict:
        res_ultima = supabase.table("lecturas").select("*").order("fecha_corte", desc=True).limit(1).execute()
        credito_anterior = res_ultima.data[0]["credito_remanente_kwh"] if res_ultima.data else 0.0
        
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
                if res_ultima.data:
                    u = res_ultima.data[0]
                    c_periodo = cons_kwh - u["lectura_cons_kwh"]
                    i_periodo = inyec_kwh - u["lectura_inyec_kwh"]
                    neto = c_periodo - i_periodo
                    
                    t_obj = [t for t in res_tarifas.data if t["id"] == tarifas_dict[tarifa_sel]][0]
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

# --- PESTAÑA 3: DESCARGAR Y EXPORTAR DATOS ---
with pestaña3:
    st.subheader("📥 Exportación de Lecturas e Historial de Crédito")
    st.write("Descarga una copia completa de tus registros para llevar un respaldo local o analizarlo en Excel.")
    
    res_export = supabase.table("lecturas").select("*, tarifas(nombre)").order("fecha_corte", desc=False).execute()
    df_exp = pd.DataFrame(res_export.data)
    
    if not df_exp.empty:
        if "tarifas" in df_exp.columns:
            df_exp["nombre_tarifa"] = df_exp["tarifas"].apply(lambda x: x.get("nombre") if isinstance(x, dict) else "")
            df_exp.drop(columns=["tarifas"], inplace=True)
            
        st.dataframe(df_exp, use_container_width=True)
        
        col_csv, col_excel = st.columns(2)
        
        # Botón para descargar CSV
        csv_data = df_exp.to_csv(index=False).encode('utf-8')
        col_csv.download_button(
            label="📄 Descargar en formato CSV",
            data=csv_data,
            file_name=f"historial_lecturas_energia_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True
        )
        
        # Botón para descargar Excel
        try:
            excel_data = generar_excel(df_exp)
            col_excel.download_button(
                label="📊 Descargar en formato Excel (.xlsx)",
                data=excel_data,
                file_name=f"historial_lecturas_energia_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        except Exception as e:
            col_excel.info("Para activar la descarga en Excel, asegúrate de tener instalada la librería `openpyxl` en tu requirements.txt")
    else:
        st.info("No hay lecturas registradas para exportar.")

# --- PESTAÑA 4: TARIFAS ---
with pestaña4:
    st.subheader("Configuración de Tarifas")
    with st.form("form_tarifa", clear_on_submit=True):
        nombre_tarifa = st.text_input("Nombre de la Tarifa", placeholder="Ej. Mi Tarifa Personalizada / Tarifa 1")
        
        c1, c2 = st.columns(2)
        cargo_fijo = c1.number_input("Cargo Fijo ($)", min_value=0.0, value=0.0)
        iva_pct = c2.number_input("IVA (%)", min_value=0.0, value=16.0)
        
        r1_col, r2_col, r3_col = st.columns(3)
        lim_basico = r1_col.number_input("Límite Básico (kWh)", min_value=1.0, value=130.0)
        p_basico = r1_col.number_input("Precio Básico ($)", min_value=0.0, value=1.15)
        
        lim_inter = r2_col.number_input("Límite Intermedio (kWh)", min_value=lim_basico, value=200.0)
        p_inter = r2_col.number_input("Precio Intermedio ($)", min_value=0.0, value=2.50)
        
        p_exced = r3_col.number_input("Precio Excedente ($)", min_value=0.0, value=5.00)
        
        dap_col1, dap_col2 = st.columns(2)
        dap_pct = dap_col1.number_input("DAP en Porcentaje (%)", min_value=0.0, value=0.0)
        dap_fijo = dap_col2.number_input("DAP en Cuota Fija ($)", min_value=0.0, value=0.0)
        
        if st.form_submit_button("Guardar Tarifa"):
            data_tarifa = {
                "nombre": nombre_tarifa,
                "cargo_fijo": cargo_fijo,
                "limite_basico": lim_basico,
                "precio_basico": p_basico,
                "limite_intermedio": lim_inter,
                "precio_intermedio": p_inter,
                "precio_excedente": p_exced,
                "porcentaje_dap": dap_pct,
                "cuota_fija_dap": dap_fijo,
                "porcentaje_iva": iva_pct
            }
            supabase.table("tarifas").insert(data_tarifa).execute()
            st.success(f"Tarifa '{nombre_tarifa}' creada con éxito.")
            st.rerun()
