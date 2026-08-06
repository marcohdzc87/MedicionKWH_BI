import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
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
    año = fecha_actual.year
    mes = fecha_actual.month
    
    dias_en_mes = calendar.monthrange(año, mes)[1]
    dia_efectivo = min(dia_corte, dias_en_mes)
    
    if fecha_actual.day >= dia_efectivo:
        inicio_ciclo = date(año, mes, dia_efectivo)
        meses_a_sumar = 2 if es_bimestral else 1
        
        mes_fin = mes + meses_a_sumar
        año_fin = año
        if mes_fin > 12:
            mes_fin -= 12
            año_fin += 1
            
        dias_en_mes_fin = calendar.monthrange(año_fin, mes_fin)[1]
        fin_ciclo = date(año_fin, mes_fin, min(dia_corte, dias_en_mes_fin))
    else:
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
        kwh_facturables = 0
        nuevo_credito = int(round(abs(balance_kwh)))
        subtotal_energia = float(tarifa.get("cargo_fijo", 0.0))
        desglose_rangos = {"Básico": 0.0, "Intermedio": 0.0, "Excedente": 0.0}
    else:
        kwh_facturables = int(round(balance_kwh))
        nuevo_credito = 0
        
        l_basico = int(round(float(tarifa.get("limite_basico", 150))))
        l_inter = int(round(float(tarifa.get("limite_intermedio", 280))))
        
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
        "kwh_facturables": kwh_facturables,
        "nuevo_credito": nuevo_credito,
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
    "📊 Dashboard & Proyección", 
    "➕ Capturar y Editar Lecturas", 
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
            dia_corte_cfe = int(tarifa_act.get("dia_corte_cfe", 15))
            es_bimestral = tarifa_act.get("es_bimestral", True)
            
            inicio_ciclo, fin_ciclo = obtener_fechas_ciclo(fecha_act, dia_corte_cfe, es_bimestral)
            dias_totales_ciclo = (fin_ciclo - inicio_ciclo).days
            dias_transcurridos = (fecha_act - lectura_anterior["fecha_corte"]).days
            dias_transcurridos = max(1, dias_transcurridos)
            
            dias_restantes = (fin_ciclo - fecha_act).days
            dias_restantes = max(0, dias_restantes)
            
            # Consumos e inyecciones en enteros estrictos
            cons_medido = max(0, int(round(lectura_actual["lectura_cons_kwh"] - lectura_anterior["lectura_cons_kwh"])))
            inyec_medida = max(0, int(round(lectura_actual["lectura_inyec_kwh"] - lectura_anterior["lectura_inyec_kwh"])))
            neto_medido = cons_medido - inyec_medida
            
            gen_solar = max(0, int(round(lectura_actual["generacion_solar_total_kwh"] - lectura_anterior["generacion_solar_total_kwh"])))
            gen_solar_acumulada = int(round(lectura_actual["generacion_solar_total_kwh"]))
            
            cons_diario = cons_medido / dias_transcurridos
            inyec_diaria = inyec_medida / dias_transcurridos
            neto_diario = neto_medido / dias_transcurridos
            gen_solar_diaria = gen_solar / dias_transcurridos
            
            kwh_restantes_proyectados = int(round(neto_diario * dias_restantes))
            neto_proyectado_total = neto_medido + kwh_restantes_proyectados
            
            credito_previo = int(round(lectura_actual.get("credito_anterior_kwh", 0)))
            
            calc_actual = calcular_detalle_factura(neto_medido, credito_previo, tarifa_act)
            calc_proyectado = calcular_detalle_factura(neto_proyectado_total, credito_previo, tarifa_act)
            costo_estimado_dias_restantes = max(0.0, calc_proyectado["total"] - calc_actual["total"])
            
            autoconsumo_solar = max(0, gen_solar - inyec_medida)
            consumo_real_total = cons_medido + autoconsumo_solar
            calc_sin_paneles = calcular_detalle_factura(consumo_real_total, 0, tarifa_act)
            ahorro_solar = max(0.0, calc_sin_paneles["total"] - calc_actual["total"])
            
            st.subheader(f"📅 Estado al {fecha_act.strftime('%d/%m/%Y')}")
            
            pct_tiempo = min(1.0, (dias_totales_ciclo - dias_restantes) / dias_totales_ciclo)
            st.progress(pct_tiempo, text=f"Progreso del Ciclo CFE: **{dias_totales_ciclo - dias_restantes} de {dias_totales_ciclo} días transcurridos** (Próximo corte: {fin_ciclo.strftime('%d/%m/%Y')})")
            
            st.divider()
            
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Tomado de CFE", f"{cons_medido:,} kWh", f"{cons_diario:.1f} kWh/día")
            col2.metric("Generación Solar Periodo", f"{gen_solar:,} kWh", f"Total Acum: {gen_solar_acumulada:,} kWh", help="Producción en el intervalo")
            col3.metric("Inyectado a CFE", f"{inyec_medida:,} kWh", f"{inyec_diaria:.1f} kWh/día")
            col4.metric("Recibo Estimado Al Día", f"${calc_actual['total']:,.2f} MXN", f"Neto: {calc_actual['kwh_facturables']:,} kWh")
            
            st.divider()
            
            st.markdown("### 🔮 Proyección al Cierre de Corte CFE")
            p1, p2, p3, p4 = st.columns(4)
            p1.metric("Días Faltantes para Corte", f"{dias_restantes} días")
            p2.metric("kWh Estimados por Consumir", f"{kwh_restantes_proyectados:,} kWh", f"{neto_diario:.1f} kWh/día neto")
            p3.metric("Costo Estimado Días Restantes", f"${costo_estimado_dias_restantes:,.2f} MXN")
            p4.metric("PROYECCIÓN TOTAL RECIBO", f"${calc_proyectado['total']:,.2f} MXN")
            
            st.divider()
            
            col_graf, col_info = st.columns([2, 1])
            
            with col_graf:
                st.subheader("Visualización del Historial de Energía")
                
                df_grafica = df.sort_values("fecha_corte").copy()
                
                df_grafica["Consumido CFE"] = df_grafica["lectura_cons_kwh"].diff().fillna(0).apply(lambda x: max(0, int(round(x))))
                df_grafica["Inyectado CFE"] = df_grafica["lectura_inyec_kwh"].diff().fillna(0).apply(lambda x: max(0, int(round(x))))
                df_grafica["Generación Solar"] = df_grafica["generacion_solar_total_kwh"].diff().fillna(0).apply(lambda x: max(0, int(round(x))))
                
                df_grafica_clean = df_grafica.iloc[1:].copy() if len(df_grafica) > 1 else df_grafica.copy()
                df_grafica_clean["Periodo"] = df_grafica_clean["fecha_corte"].astype(str)
                
                v_graf1, v_graf2, v_graf3 = st.tabs([
                    "📊 Barras Agrupadas (Intervalos)", 
                    "🌊 Gráfica de Área (3 Medidas)", 
                    "📈 Odómetros Acumulados"
                ])
                
                # OPCIÓN 1: Barras Agrupadas
                with v_graf1:
                    if not df_grafica_clean.empty:
                        df_melted = df_grafica_clean.melt(
                            id_vars=["Periodo"], 
                            value_vars=["Consumido CFE", "Inyectado CFE", "Generación Solar"],
                            var_name="Concepto", 
                            value_name="kWh"
                        )
                        fig1 = px.bar(
                            df_melted, 
                            x="Periodo", 
                            y="kWh", 
                            color="Concepto",
                            barmode="group",
                            text_auto=True,
                            color_discrete_map={
                                "Consumido CFE": "#1f77b4",
                                "Inyectado CFE": "#2ca02c",
                                "Generación Solar": "#ff7f0e"
                            },
                            title="Energía del Intervalo (kWh Reales por Toma)"
                        )
                        fig1.update_layout(yaxis_title="kWh", xaxis_title="Fecha de Corte")
                        st.plotly_chart(fig1, use_container_width=True)
                    else:
                        st.info("Registra al menos 2 lecturas para visualizar barras agrupadas.")
                        
                # OPCIÓN 2: GRÁFICA DE ÁREA ROBUSTA (GRAPH OBJECTS)
                with v_graf2:
                    if not df_grafica_clean.empty:
                        fig2 = go.Figure()
                        
                        # Área 1: Consumido CFE (Azul)
                        fig2.add_trace(go.Scatter(
                            x=df_grafica_clean["Periodo"],
                            y=df_grafica_clean["Consumido CFE"],
                            name="Consumido CFE",
                            mode="lines+markers",
                            fill="tozeroy",
                            line=dict(color="#1f77b4", width=2),
                            fillcolor="rgba(31, 119, 180, 0.3)"
                        ))
                        
                        # Área 2: Inyectado CFE (Verde)
                        fig2.add_trace(go.Scatter(
                            x=df_grafica_clean["Periodo"],
                            y=df_grafica_clean["Inyectado CFE"],
                            name="Inyectado CFE",
                            mode="lines+markers",
                            fill="tozeroy",
                            line=dict(color="#2ca02c", width=2),
                            fillcolor="rgba(44, 160, 44, 0.3)"
                        ))
                        
                        # Área 3: Generación Solar (Naranja)
                        fig2.add_trace(go.Scatter(
                            x=df_grafica_clean["Periodo"],
                            y=df_grafica_clean["Generación Solar"],
                            name="Generación Solar",
                            mode="lines+markers",
                            fill="tozeroy",
                            line=dict(color="#ff7f0e", width=2),
                            fillcolor="rgba(255, 127, 14, 0.3)"
                        ))
                        
                        fig2.update_layout(
                            title="Comparativa de Volúmenes Energéticos (Áreas Superpuestas)",
                            xaxis_title="Fecha de Corte",
                            yaxis_title="kWh del Intervalo",
                            hovermode="x unified",
                            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                        )
                        st.plotly_chart(fig2, use_container_width=True)
                    else:
                        st.info("Registra al menos 2 lecturas para calcular la gráfica de área.")
                        
                # OPCIÓN 3: Odómetros Acumulados
                with v_graf3:
                    fig3 = px.line(
                        df, 
                        x="fecha_corte", 
                        y=["lectura_cons_kwh", "lectura_inyec_kwh", "generacion_solar_total_kwh"],
                        markers=True,
                        labels={"value": "Lectura Acumulada (kWh)", "variable": "Concepto", "fecha_corte": "Fecha"},
                        title="Evolución Continua de Contadores Brutos"
                    )
                    st.plotly_chart(fig3, use_container_width=True)
                
            with col_info:
                st.subheader("Desglose Financiero y Ahorro")
                st.write(f"• **Crédito en Bolsa Previo:** {credito_previo} kWh")
                
                if calc_actual['nuevo_credito'] > 0:
                    st.success(f"🎉 Tienes **{calc_actual['nuevo_credito']} kWh** de saldo a favor en la bolsa.")
                
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
                st.info(f"☀️ **Ahorro Solar del Periodo:** Has ahorrado **${ahorro_solar:,.2f} MXN** gracias a tus paneles.")
        else:
            st.warning("La última lectura registrada no tiene asociada una tarifa válida.")

    elif not df.empty and len(df) == 1:
        st.info("Has registrado tu primera lectura inicial. Agrega una segunda lectura para activar proyecciones.")
    else:
        st.info("👋 **¡Bienvenido!** Dirígete a la pestaña **⚙️ Administración de Tarifas y Ciclo** para configurar tus parámetros de corte.")

# --- PESTAÑA 2: CAPTURAR Y EDITAR LECTURAS ---
with pestaña2:
    try:
        res_tarifas = supabase.table("tarifas").select("*").execute()
        tarifas_list = res_tarifas.data if res_tarifas.data else []
    except Exception:
        tarifas_list = []
        
    tarifas_dict = {t["nombre"]: t["id"] for t in tarifas_list} if tarifas_list else {}
    
    subtab_lec1, subtab_lec2 = st.tabs(["➕ Capturar Nueva Lectura", "✏️ Editar / Eliminar Lecturas Guardadas"])
    
    # 1. CAPTURAR NUEVA LECTURA
    with subtab_lec1:
        st.subheader("Ingresar nueva captura del medidor")
        if tarifas_dict:
            try:
                res_ultima = supabase.table("lecturas").select("*").order("fecha_corte", desc=True).limit(1).execute()
                ult_reg = res_ultima.data[0] if res_ultima.data else None
                credito_anterior = int(round(ult_reg["credito_remanente_kwh"])) if ult_reg else 0
            except Exception:
                ult_reg = None
                credito_anterior = 0
            
            if ult_reg:
                min_cons = int(round(ult_reg["lectura_cons_kwh"]))
                min_inyec = int(round(ult_reg["lectura_inyec_kwh"]))
                min_solar = int(round(ult_reg["generacion_solar_total_kwh"]))
                st.info(f"📌 **Valores mínimos esperados (Lectura del {ult_reg['fecha_corte']}):**  \n"
                        f"• Consumo Red $\ge$ **{min_cons:,} kWh** | "
                        f"• Inyección Red $\ge$ **{min_inyec:,} kWh** | "
                        f"• Generación Solar $\ge$ **{min_solar:,} kWh**")
            else:
                min_cons, min_inyec, min_solar = 0, 0, 0
            
            with st.form("form_registro_libre", clear_on_submit=False):
                fecha = st.date_input("Fecha de Toma de Lectura", value=datetime.now())
                
                c_a, c_b = st.columns(2)
                cons_kwh = c_a.number_input("Lectura Medidor - Consumo (kWh / Código 1.8.0)", min_value=0, step=1, value=min_cons)
                inyec_kwh = c_b.number_input("Lectura Medidor - Inyección (kWh / Código 2.8.0)", min_value=0, step=1, value=min_inyec)
                
                solar_kwh = st.number_input("Generación Solar Acumulada Inversor (kWh)", min_value=0, step=1, value=min_solar)
                tarifa_sel = st.selectbox("Tarifa Aplicable", list(tarifas_dict.keys()))
                
                es_cierre = st.checkbox("Marcar como Cierre Oficial de Bimestre / Recibo CFE")
                notas = st.text_input("Notas (Ej. 'Lectura semanal', 'Corte oficial CFE')")
                
                if st.form_submit_button("Guardar Registro"):
                    errores_val = []
                    if ult_reg:
                        if cons_kwh < min_cons:
                            errores_val.append(f"El consumo ({cons_kwh:,} kWh) no puede ser menor a la lectura previa ({min_cons:,} kWh).")
                        if inyec_kwh < min_inyec:
                            errores_val.append(f"La inyección ({inyec_kwh:,} kWh) no puede ser menor a la lectura previa ({min_inyec:,} kWh).")
                        if solar_kwh < min_solar:
                            errores_val.append(f"La generación solar ({solar_kwh:,} kWh) no puede ser menor a la lectura previa ({min_solar:,} kWh).")
                    
                    if errores_val:
                        for err in errores_val:
                            st.error(f"❌ {err}")
                    else:
                        if ult_reg:
                            c_periodo = cons_kwh - min_cons
                            i_periodo = inyec_kwh - min_inyec
                            neto = c_periodo - i_periodo
                            
                            t_obj = [t for t in tarifas_list if t["id"] == tarifas_dict[tarifa_sel]][0]
                            res_calc = calcular_detalle_factura(neto, credito_anterior, t_obj)
                            
                            nuevo_rem = res_calc["nuevo_credito"] if es_cierre else credito_anterior
                        else:
                            nuevo_rem = 0
                        
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

    # 2. EDITAR / ELIMINAR LECTURAS
    with subtab_lec2:
        st.subheader("Consultar o modificar capturas previas")
        try:
            res_todas_l = supabase.table("lecturas").select("*, tarifas(nombre)").order("fecha_corte", desc=True).execute()
            lecturas_list = res_todas_l.data if res_todas_l.data else []
        except Exception:
            lecturas_list = []
            
        if lecturas_list:
            opciones_lec = {f"{l['fecha_corte']} - Consumo: {int(round(l['lectura_cons_kwh']))} kWh / Inyección: {int(round(l['lectura_inyec_kwh']))} kWh / Solar: {int(round(l['generacion_solar_total_kwh']))} kWh": l for l in lecturas_list}
            sel_lec_label = st.selectbox("Selecciona una lectura registrada para editar o eliminar:", list(opciones_lec.keys()))
            lec_sel = opciones_lec[sel_lec_label]
            
            with st.form("form_edit_lectura"):
                st.markdown(f"#### Modificando Registro ID: {lec_sel['id']}")
                edit_fecha = st.date_input("Fecha de Lectura", value=datetime.strptime(lec_sel["fecha_corte"], "%Y-%m-%d").date())
                
                ec1, ec2 = st.columns(2)
                edit_cons = ec1.number_input("Lectura Consumo (kWh)", min_value=0, step=1, value=int(round(lec_sel["lectura_cons_kwh"])))
                edit_inyec = ec2.number_input("Lectura Inyección (kWh)", min_value=0, step=1, value=int(round(lec_sel["lectura_inyec_kwh"])))
                
                edit_solar = st.number_input("Generación Solar Total (kWh)", min_value=0, step=1, value=int(round(lec_sel["generacion_solar_total_kwh"])))
                
                idx_tar = 0
                if tarifas_list and lec_sel.get("tarifa_id"):
                    ids_t = [t["id"] for t in tarifas_list]
                    if lec_sel["tarifa_id"] in ids_t:
                        idx_tar = ids_t.index(lec_sel["tarifa_id"])
                        
                edit_tarifa = st.selectbox("Tarifa Aplicada", list(tarifas_dict.keys()), index=idx_tar)
                edit_cierre = st.checkbox("Es Cierre Oficial de Bimestre", value=bool(lec_sel.get("es_cierre_periodo", False)))
                edit_notas = st.text_input("Notas", value=lec_sel.get("notas", "") or "")
                
                btn_save, btn_del = st.columns(2)
                
                if btn_save.form_submit_button("💾 Guardar Cambios"):
                    data_edit_lec = {
                        "fecha_corte": str(edit_fecha),
                        "lectura_cons_kwh": edit_cons,
                        "lectura_inyec_kwh": edit_inyec,
                        "generacion_solar_total_kwh": edit_solar,
                        "tarifa_id": tarifas_dict[edit_tarifa],
                        "es_cierre_periodo": edit_cierre,
                        "notas": edit_notas
                    }
                    supabase.table("lecturas").update(data_edit_lec).eq("id", lec_sel["id"]).execute()
                    st.success("¡Lectura modificada correctamente!")
                    st.rerun()
                    
                if btn_del.form_submit_button("🗑️ Eliminar Registro"):
                    supabase.table("lecturas").delete().eq("id", lec_sel["id"]).execute()
                    st.success("Lectura eliminada.")
                    st.rerun()
        else:
            st.info("No hay lecturas registradas para editar.")

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
    
    # 1. VER / EDITAR TARIFAS
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
                edit_cargo = c1.number_input("Cargo Fijo ($)", min_value=0.0, value=float(t_sel.get("cargo_fijo", 0.0)), step=0.001, format="%.3f")
                edit_iva = c2.number_input("IVA (%)", min_value=0.0, value=float(t_sel.get("porcentaje_iva", 16.0)), step=0.001, format="%.3f")
                
                st.markdown("#### Rangos de Consumo (kWh Enteros) y Precios ($ con 3 decimales)")
                r1_col, r2_col, r3_col = st.columns(3)
                edit_lim_b = r1_col.number_input("Límite Básico (kWh)", min_value=1, value=int(round(float(t_sel.get("limite_basico", 150)))), step=1)
                edit_p_b = r1_col.number_input("Precio Básico ($/kWh)", min_value=0.0, value=float(t_sel.get("precio_basico", 1.087)), step=0.001, format="%.3f")
                
                edit_lim_i = r2_col.number_input("Límite Intermedio Acumulado (kWh)", min_value=edit_lim_b, value=int(round(float(t_sel.get("limite_intermedio", 280)))), step=1)
                edit_p_i = r2_col.number_input("Precio Intermedio ($/kWh)", min_value=0.0, value=float(t_sel.get("precio_intermedio", 1.320)), step=0.001, format="%.3f")
                
                edit_p_e = r3_col.number_input("Precio Excedente ($/kWh)", min_value=0.0, value=float(t_sel.get("precio_excedente", 3.861)), step=0.001, format="%.3f")
                
                st.markdown("#### Alumbrado Público (DAP)")
                dap_col1, dap_col2 = st.columns(2)
                edit_dap_pct = dap_col1.number_input("DAP en Porcentaje (%)", min_value=0.0, value=float(t_sel.get("porcentaje_dap", 0.0)), step=0.001, format="%.3f")
                edit_dap_fijo = dap_col2.number_input("DAP en Cuota Fija ($)", min_value=0.0, value=float(t_sel.get("cuota_fija_dap", 0.0)), step=0.001, format="%.3f")
                
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
                    supabase.table("tarifas").update(data_update).eq("id", t_sel["id"]).execute()
                    st.success(f"¡Tarifa '{edit_nombre}' actualizada correctamente!")
                    st.rerun()
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
            cargo_fijo = c1.number_input("Cargo Fijo ($)", min_value=0.0, value=0.0, step=0.001, format="%.3f")
            iva_pct = c2.number_input("IVA (%)", min_value=0.0, value=16.0, step=0.001, format="%.3f")
            
            st.markdown("#### Rangos de Consumo (kWh Enteros) y Precios ($ con 3 decimales)")
            r1_col, r2_col, r3_col = st.columns(3)
            lim_basico = r1_col.number_input("Límite Básico (kWh)", min_value=1, value=150, step=1)
            p_basico = r1_col.number_input("Precio Básico ($/kWh)", min_value=0.0, value=1.087, step=0.001, format="%.3f")
            
            lim_inter = r2_col.number_input("Límite Intermedio Acumulado (kWh)", min_value=lim_basico, value=280, step=1)
            p_inter = r2_col.number_input("Precio Intermedio ($/kWh)", min_value=0.0, value=1.320, step=0.001, format="%.3f")
            
            p_exced = r3_col.number_input("Precio Excedente ($/kWh)", min_value=0.0, value=3.861, step=0.001, format="%.3f")
            
            dap_col1, dap_col2 = st.columns(2)
            dap_pct = dap_col1.number_input("DAP en Porcentaje (%)", min_value=0.0, value=0.0, step=0.001, format="%.3f")
            dap_fijo = dap_col2.number_input("DAP en Cuota Fija ($)", min_value=0.0, value=0.0, step=0.001, format="%.3f")
            
            if st.form_submit_button("➕ Registrar Nueva Tarifa"):
                data_tarifa = {
                    "nombre": nombre_tarifa,
                    "dia_corte_cfe": int(dia_corte),
                    "es_bimestral": True if es_bimestral == "Bimestral (60 días)" else False,
                    "cargo_fijo": float(cargo_fijo),
                    "limite_basico": int(lim_basico),
                    "precio_basico": float(p_basico),
                    "limite_intermedio": int(lim_inter),
                    "precio_intermedio": float(p_inter),
                    "precio_excedente": float(p_exced),
                    "porcentaje_dap": float(dap_pct),
                    "cuota_fija_dap": float(dap_fijo),
                    "porcentaje_iva": float(iva_pct)
                }
                supabase.table("tarifas").insert(data_tarifa).execute()
                st.success(f"Tarifa '{nombre_tarifa}' guardada con éxito.")
                st.rerun()
