import streamlit as st
import pandas as pd
import re
import os

st.set_page_config(page_title="Reporte de DCPD", layout="wide")
st.title("📊 Reporte de DCPD (valores Acreditados y Rechazados)")

uploaded_file = st.file_uploader("📂 Subí tu archivo Excel ValoresAcreditados", type=["xls", "xlsx", "xlsb", "csv"])

# -----------------------------
# Lectura robusta del archivo 
# -----------------------------
def leer_archivo_robusto(file):
    ext = os.path.splitext(file.name)[1].lower().lstrip(".")
    file.seek(0)
    try:
        if ext == "xlsx":
            return pd.read_excel(file, engine="openpyxl")
        elif ext == "xls":
            try:
                return pd.read_excel(file, engine="xlrd")
            except Exception:
                try:
                    return pd.read_excel(file, engine="pyxlsb")
                except Exception:
                    file.seek(0)
                    try:
                        return pd.read_csv(file, sep="\t", encoding="latin1", quotechar='"', engine="python")
                    except Exception:
                        file.seek(0)
                        try:
                            return pd.read_csv(file, sep=";", encoding="latin1", quotechar='"', engine="python")
                        except Exception:
                            file.seek(0)
                            return pd.read_csv(file, sep=",", encoding="latin1", quotechar='"', engine="python")
        elif ext == "xlsb":
            return pd.read_excel(file, engine="pyxlsb")
        elif ext == "csv":
            file.seek(0)
            try:
                return pd.read_csv(file, sep=";", encoding="latin1", quotechar='"', engine="python")
            except Exception:
                file.seek(0)
                return pd.read_csv(file, sep=",", encoding="latin1", quotechar='"', engine="python")
        else:
            raise ValueError("Formato no soportado")
    except Exception as e:
        raise ValueError(f"No se pudo leer el archivo: {e}")

# -----------------------------
# Función para parsear montos
# -----------------------------
def parse_amount_from_text(text):
    if pd.isna(text):
        return 0.0
    s = str(text).upper().strip()
    m = re.search(r"[-+]?[0-9\.,]+", s)
    if not m:
        return 0.0
    token = m.group(0).replace('"', '').strip()
    if token.count(".") > 0 and token.count(",") > 0:
        if token.rfind(",") > token.rfind("."):
            token = token.replace(".", "").replace(",", ".")
        else:
            token = token.replace(",", "")
    elif token.count(",") > 0 and token.count(".") == 0:
        part_after = token.split(",")[-1]
        if 1 <= len(part_after) <= 2:
            token = token.replace(".", "").replace(",", ".")
        else:
            token = token.replace(",", "")
    else:
        token = token.replace(",", "")
    try:
        return float(token)
    except:
        return 0.0

# -----------------------------
# Formateo visual de montos
# -----------------------------
def fmt_monto(x):
    try:
        return f"$ {x:,.0f}".replace(",", ".")
    except:
        return "$ 0"

# -----------------------------
# Mostrar tabla (Fuente ampliada + Ancho forzado "large")
# -----------------------------
def mostrar_tabla_estilizada(df_to_show):
    df_to_show = df_to_show.copy()
    df_to_show.index = range(1, len(df_to_show) + 1)
    styled = df_to_show.style.set_properties(**{'font-size': '16px'}).set_table_styles([{'selector': 'th', 'props': [('font-size', '16px')]}])
    altura_dinamica = min(400, 45 + (len(df_to_show) * 36))
    st.dataframe(
        styled, 
        height=altura_dinamica, 
        use_container_width=False, 
        column_config={
            "Den. Firmante": st.column_config.TextColumn("Den. Firmante", width="large"), 
            "Motivo del rechazo": st.column_config.TextColumn("Motivo del rechazo", width="large")
        }
    )

# -----------------------------
# Main
# -----------------------------
if uploaded_file:
    try:
        df = leer_archivo_robusto(uploaded_file)
    except Exception as e:
        st.error(f"No se pudo leer el archivo: {e}"); st.stop()

    df.columns = df.columns.astype(str).str.strip().str.replace('"', '')
    required_cols = ["Tipo Op.", "Monto Acreditado / Rechazado", "Den. Firmante", "Fecha Acreditación"]
    if any(c not in df.columns for c in required_cols):
        st.error("Faltan columnas necesarias en el archivo."); st.stop()

    df["Tipo Op."] = df["Tipo Op."].astype(str).str.strip().str.replace('"', '')
    df = df[df["Tipo Op."] == "CO"].copy()

    if "Motivo Rechazo" not in df.columns: df["Motivo Rechazo"] = ""
    df["Motivo Rechazo"] = df["Motivo Rechazo"].astype(str)
    df["Den. Firmante"] = df["Den. Firmante"].astype(str).str.strip().str.replace('"', '')

    df["_monto_texto"] = df["Monto Acreditado / Rechazado"].astype(str)
    df["Estado"] = df["_monto_texto"].str.upper().apply(lambda x: "ACREDITADO" if "ACREDITADO" in x else ("RECHAZADO" if "RECHAZADO" in x else "OTRO"))
    df["Monto"] = df["_monto_texto"].apply(parse_amount_from_text)

    # --- LÓGICA DE RECHAZOS FINANCIEROS/JUDICIALES (R01, R02, R10, R21) ---
    mask_rechazo_finan = (df["Estado"] == "RECHAZADO") & df["Motivo Rechazo"].str.contains(r"R01|R02|R10|R21", na=False, regex=True)
    
    total_acreditado = df.loc[df["Estado"] == "ACREDITADO", "Monto"].sum()
    rechazados_finan = df.loc[mask_rechazo_finan, "Monto"].sum()
    total_operado = total_acreditado + rechazados_finan

    pct_acreditado = (total_acreditado / total_operado * 100) if total_operado > 0 else 0.0
    pct_rechazado_finan = (rechazados_finan / total_operado * 100) if total_operado > 0 else 0.0

    # Cabecera Socio
    try:
        fechas = pd.to_datetime(df["Fecha Acreditación"], errors="coerce")
        socio = df["Den. Socio"].dropna().unique()[0] if "Den. Socio" in df.columns else ""
        if fechas.notna().any():
            st.markdown(f'<div style="font-size:30px; font-weight:bold; color:green; margin-bottom:10px;">👥 {socio}</div>', unsafe_allow_html=True)
            st.markdown(f'<div style="font-size:20px; font-weight:bold; color:#444;">📅 Detalle DCPD entre <span>{fechas.min().strftime("%d/%m/%Y")}</span> y <span>{fechas.max().strftime("%d/%m/%Y")}</span></div>', unsafe_allow_html=True)
    except: pass

    # Métricas principales
    cant_total_operado = len(df[(df["Estado"] == "ACREDITADO") | mask_rechazo_finan])
    col1, col2, col3 = st.columns(3)
    with col1: st.metric("📦 Total Operado", fmt_monto(total_operado)); st.markdown(f"Cant: **{cant_total_operado}**")
    with col2: st.metric("💰 Total Acreditado", fmt_monto(total_acreditado)); st.markdown(f"Cant: **{len(df[df['Estado'] == 'ACREDITADO'])}**")
    with col3: st.metric("❌ Total Rechazados", fmt_monto(rechazados_finan)); st.markdown(f"Cant: **{len(df[mask_rechazo_finan])}**")

    colA, colB = st.columns(2)
    colA.markdown(f"<div style='font-size:26px; font-weight:bold; color:green;'>✅ % Acreditado: {pct_acreditado:.2f}%</div>", unsafe_allow_html=True)
    colB.markdown(f"<div style='font-size:26px; font-weight:bold; color:red;'>❌ % Rechazados: {pct_rechazado_finan:.2f}%</div>", unsafe_allow_html=True)

    # Top Firmantes
    df_firmantes = df[(df["Estado"] == "ACREDITADO") | mask_rechazo_finan].copy()
    df_firmantes["Tipo"] = df_firmantes.apply(lambda row: "ACREDITADO" if row["Estado"] == "ACREDITADO" else "RECHAZADO", axis=1)
    firmantes = df_firmantes.groupby(["Den. Firmante", "Tipo"])["Monto"].sum().unstack(fill_value=0).reset_index()
    for c in ["ACREDITADO", "RECHAZADO"]: 
        if c not in firmantes.columns: firmantes[c] = 0
    firmantes["Total_Firmante"] = firmantes["ACREDITADO"] + firmantes["RECHAZADO"]
    firmantes["% Concentración"] = (firmantes["Total_Firmante"] / total_operado * 100) if total_operado > 0 else 0
    firmantes = firmantes.sort_values("Total_Firmante", ascending=False).reset_index(drop=True)
    
    st.subheader("👤 Top 10 Firmantes (sobre total operado)")
    mostrar_tabla_estilizada(firmantes.assign(ACREDITADO=firmantes["ACREDITADO"].apply(fmt_monto), RECHAZADO=firmantes["RECHAZADO"].apply(fmt_monto), Total_Firmante=firmantes["Total_Firmante"].apply(fmt_monto), **{"% Concentración": firmantes["% Concentración"].apply(lambda x: f"{x:.2f}%")}))

    # FRASE RESUMEN GLOBAL
    if rechazados_finan == 0:
        st.info(f"**Descontó {cant_total_operado} valores por un total de {fmt_monto(total_operado)} sin registrar rechazos.**")
    else:
        st.info(f"**Descontó {cant_total_operado} valores por un total de {fmt_monto(total_operado)} con un margen de rechazos del {pct_rechazado_finan:.2f}%.**")

    # Tabla Rechazados por Problemas Financieros/Judiciales
    st.subheader("👤 Totales Rechazados por Firmante (solo rechazos por problemas financieros)")
    firmantes_r_finan = df[mask_rechazo_finan].groupby("Den. Firmante").agg(Monto=("Monto", "sum"), Motivo_Rechazo=("Motivo Rechazo", lambda x: " | ".join(sorted(set(x.dropna().astype(str).str.strip()))))).reset_index().sort_values("Monto", ascending=False)
    if not firmantes_r_finan.empty:
        firmantes_r_finan["% Concentración"] = firmantes_r_finan["Monto"] / rechazados_finan * 100
        mostrar_tabla_estilizada(firmantes_r_finan.assign(Monto=firmantes_r_finan["Monto"].apply(fmt_monto), **{"% Concentración": firmantes_r_finan["% Concentración"].apply(lambda x: f"{x:.2f}%")}).rename(columns={"Motivo_Rechazo": "Motivo del rechazo"}))
    else:
        st.write("No se registran rechazos financieros o judiciales (R01, R02, R10, R21).")

    # -----------------------------
    # SECCIÓN: ANÁLISIS DE LOS ÚLTIMOS 4 MESES 
    # -----------------------------
    if rechazados_finan > 0:
        st.markdown("---")
        fecha_actual = pd.Timestamp.today().normalize()
        min_date_4m = fecha_actual - pd.DateOffset(months=4)
        fechas_dt = pd.to_datetime(df["Fecha Acreditación"], errors="coerce")
        df_4m = df[(fechas_dt >= min_date_4m) & (fechas_dt <= fecha_actual)].copy()
        
        mask_r_finan_4m = (df_4m["Estado"] == "RECHAZADO") & df_4m["Motivo Rechazo"].str.contains(r"R01|R02|R10|R21", na=False, regex=True)
        total_acreditado_4m = df_4m.loc[df_4m["Estado"] == "ACREDITADO", "Monto"].sum()
        rechazados_finan_4m = df_4m.loc[mask_r_finan_4m, "Monto"].sum()
        total_operado_4m = total_acreditado_4m + rechazados_finan_4m
        pct_r_finan_4m = (rechazados_finan_4m / total_operado_4m * 100) if total_operado_4m > 0 else 0.0

        st.markdown(f'<div style="font-size:24px; font-weight:bold; color:#d62728; margin-bottom:10px;">🔎 Foco de Riesgo: Últimos 4 Meses (Desde {min_date_4m.strftime("%d/%m/%Y")} hasta Hoy)</div>', unsafe_allow_html=True)

        if total_operado_4m == 0:
            st.info("Durante los últimos 4 meses no ha registrado operatoria en DCPD")
        elif rechazados_finan_4m == 0:
            st.info(f"Durante los últimos 4 meses la operatoria en DCPD totalizó **{fmt_monto(total_operado_4m)}**, sin registrar rechazos.")
        else:
            df_4m["Mes_Anio"] = pd.to_datetime(df_4m["Fecha Acreditación"], errors="coerce").dt.strftime('%m-%Y')
            rechazos_mes = df_4m[mask_r_finan_4m].groupby("Mes_Anio")["Monto"].sum()
            meses_pct = (rechazos_mes / rechazados_finan_4m * 100).round().astype(int).sort_values(ascending=False)
            str_meses = ", ".join([f"{mes} ({pct}%)" for mes, pct in meses_pct.items()])
            st.info(f"**Durante los últimos 4 meses la operatoria en DCPD totalizó {fmt_monto(total_operado_4m)}, con un margen de rechazos del {pct_r_finan_4m:.2f}%, concentrados en los meses de {str_meses}.**")
            
            # Tabla 4M
            st.subheader("👤 Top Firmantes - Últimos 4 Meses")
            df_firm_4m = df_4m[(df_4m["Estado"] == "ACREDITADO") | mask_r_finan_4m].copy()
            df_firm_4m["Tipo"] = df_firm_4m.apply(lambda row: "ACREDITADO" if row["Estado"] == "ACREDITADO" else "RECHAZADO", axis=1)
            firm_4m_res = df_firm_4m.groupby(["Den. Firmante", "Tipo"])["Monto"].sum().unstack(fill_value=0).reset_index()
            for c in ["ACREDITADO", "RECHAZADO"]: 
                if c not in firm_4m_res.columns: firm_4m_res[c] = 0
            firm_4m_res["Total"] = firm_4m_res["ACREDITADO"] + firm_4m_res["RECHAZADO"]
            mostrar_tabla_estilizada(firm_4m_res.sort_values("Total", ascending=False).assign(ACREDITADO=firm_4m_res["ACREDITADO"].apply(fmt_monto), RECHAZADO=firm_4m_res["RECHAZADO"].apply(fmt_monto), Total=firm_4m_res["Total"].apply(fmt_monto)))
