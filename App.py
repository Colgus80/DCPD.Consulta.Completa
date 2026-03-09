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
# Mostrar tabla estática HTML (Solución para todas las resoluciones)
# -----------------------------
def mostrar_tabla_estilizada(df_to_show):
    df_to_show = df_to_show.copy()
    df_to_show.index = range(1, len(df_to_show) + 1)
    
    # Configuramos el CSS para que la tabla sea inamovible frente a la resolución
    styled = df_to_show.style.set_properties(**{
        'font-size': '16px',
        'white-space': 'nowrap', # Fuerza a que el texto no baje de renglón
        'padding': '10px'        # Da aire a las celdas
    }).set_table_styles([
        {'selector': 'th', 'props': [('font-size', '17px'), ('text-align', 'left'), ('background-color', '#f0f2f6')]}
    ])
    
    # Usamos st.table() en vez de st.dataframe() para anular la grilla interactiva
    st.table(styled)

# -----------------------------
# Preparador para Datos Crudos
# -----------------------------
def preparar_datos_crudos(df_in):
    mapeo_columnas = {
        "Den.Socio": "Den. Socio", "Den. Socio": "Den. Socio",
        "Tipo op.": "Tipo Op.", "Tipo Op.": "Tipo Op.",
        "CUI": "CUIT", "CUIT": "CUIT",
        "Den.Firmante": "Den. Firmante", "Den. Firmante": "Den. Firmante",
        "Monto": "Monto", "Fecha Acreditación": "Fecha Acreditación",
        "Estado": "Estado", "Motivo Rechazo": "Motivo Rechazo"
    }
    cols_encontradas = [col for col in df_in.columns if col in mapeo_columnas]
    df_out = df_in[cols_encontradas].copy().rename(columns={c: mapeo_columnas[c] for c in cols_encontradas})
    if "Fecha Acreditación" in df_out.columns:
        df_out["Fecha Acreditación"] = pd.to_datetime(df_out["Fecha Acreditación"], errors='coerce').dt.strftime('%d/%m/%Y')
    if "Monto" in df_out.columns:
        df_out["Monto"] = df_out["Monto"].apply(fmt_monto)
    orden_ideal = ["Den. Socio", "Tipo Op.", "CUIT", "Den. Firmante", "Monto", "Fecha Acreditación", "Estado", "Motivo Rechazo"]
    df_final = df_out[[col for col in orden_ideal if col in df_out.columns]]
    df_final.index = range(1, len(df_final) + 1)
    return df_final

# -----------------------------
# Lógica Principal
# -----------------------------
if uploaded_file:
    try:
        df = leer_archivo_robusto(uploaded_file)
    except Exception as e:
        st.error(f"Error al leer archivo: {e}"); st.stop()

    df.columns = df.columns.astype(str).str.strip().str.replace('"', '')
    required_cols = ["Tipo Op.", "Monto Acreditado / Rechazado", "Den. Firmante", "Fecha Acreditación"]
    if any(c not in df.columns for c in required_cols):
        st.error("Archivo incompatible. Faltan columnas necesarias."); st.stop()

    df["Tipo Op."] = df["Tipo Op."].astype(str).str.strip().str.replace('"', '')
    df = df[df["Tipo Op."] == "CO"].copy()
    df["Motivo Rechazo"] = df.get("Motivo Rechazo", "").astype(str)
    df["Den. Firmante"] = df["Den. Firmante"].astype(str).str.strip().str.replace('"', '')
    df["_monto_texto"] = df["Monto Acreditado / Rechazado"].astype(str)
    df["Estado"] = df["_monto_texto"].str.upper().apply(lambda x: "ACREDITADO" if "ACREDITADO" in x else ("RECHAZADO" if "RECHAZADO" in x else "OTRO"))
    df["Monto"] = df["_monto_texto"].apply(parse_amount_from_text)

    # Lógica de Rechazos Financieros (R01, R02, R10, R21)
    mask_rechazo_finan = (df["Estado"] == "RECHAZADO") & df["Motivo Rechazo"].str.contains(r"R01|R02|R10|R21", na=False, regex=True)
    
    total_acreditado = df.loc[df["Estado"] == "ACREDITADO", "Monto"].sum()
    rechazados_finan = df.loc[mask_rechazo_finan, "Monto"].sum()
    total_operado = total_acreditado + rechazados_finan
    pct_rechazado_finan = (rechazados_finan / total_operado * 100) if total_operado > 0 else 0.0

    # Cabecera
    try:
        fechas = pd.to_datetime(df["Fecha Acreditación"], errors="coerce")
        socio = df["Den. Socio"].dropna().unique()[0] if "Den. Socio" in df.columns else "Sin Nombre"
        st.markdown(f'<div style="font-size:30px; font-weight:bold; color:green; margin-bottom:10px;">👥 {socio}</div>', unsafe_allow_html=True)
        st.markdown(f'<div style="font-size:20px; font-weight:bold; color:#444;">📅 Detalle DCPD entre {fechas.min().strftime("%d/%m/%Y")} y {fechas.max().strftime("%d/%m/%Y")}</div>', unsafe_allow_html=True)
    except: pass

    # Métricas
    cant_total_operado = len(df[(df["Estado"] == "ACREDITADO") | mask_rechazo_finan])
    c1, c2, c3 = st.columns(3)
    c1.metric("📦 Total Operado", fmt_monto(total_operado))
    c1.markdown(f"Cantidad de cheques: **{cant_total_operado}**")
    c2.metric("💰 Total Acreditado", fmt_monto(total_acreditado))
    c2.markdown(f"Cantidad de cheques: **{len(df[df['Estado'] == 'ACREDITADO'])}**")
    c3.metric("❌ Rechazados", fmt_monto(rechazados_finan))
    c3.markdown(f"Cantidad de cheques: **{len(df[mask_rechazo_finan])}**")

    st.columns(2)[0].markdown(f"<div style='font-size:26px; font-weight:bold; color:green;'>✅ % Acreditado: {(total_acreditado/total_operado*100 if total_operado>0 else 0):.2f}%</div>", unsafe_allow_html=True)
    st.columns(2)[1].markdown(f"<div style='font-size:26px; font-weight:bold; color:green;'>❌ % Rechazados: {pct_rechazado_finan:.2f}%</div>", unsafe_allow_html=True)

    # Tabla Top Firmantes
    df_firmantes = df[(df["Estado"] == "ACREDITADO") | mask_rechazo_finan].copy()
    df_firmantes["Tipo"] = df_firmantes["Estado"].apply(lambda x: "ACREDITADO" if x == "ACREDITADO" else "RECHAZADO")
    firm_res = df_firmantes.groupby(["Den. Firmante", "Tipo"])["Monto"].sum().unstack(fill_value=0).reset_index()
    for c in ["ACREDITADO", "RECHAZADO"]: 
        if c not in firm_res.columns: firm_res[c] = 0
    firm_res["Total_Firmante"] = firm_res["ACREDITADO"] + firm_res["RECHAZADO"]
    firm_res["% Concentración"] = (firm_res["Total_Firmante"] / total_operado * 100) if total_operado > 0 else 0
    
    st.subheader("👤 Top 10 Firmantes (sobre total operado)")
    mostrar_tabla_estilizada(firm_res.sort_values("Total_Firmante", ascending=False).head(10).assign(ACREDITADO=lambda x: x["ACREDITADO"].apply(fmt_monto), RECHAZADO=lambda x: x["RECHAZADO"].apply(fmt_monto), Total_Firmante=lambda x: x["Total_Firmante"].apply(fmt_monto), **{"% Concentración": lambda x: x["% Concentración"].apply(lambda v: f"{v:.2f}%")}))

    if rechazados_finan == 0:
        st.info(f"**Descontó {cant_total_operado} valores por un total de {fmt_monto(total_operado)} sin registrar rechazos.**")
    else:
        st.info(f"**Descontó {cant_total_operado} valores por un total de {fmt_monto(total_operado)} con un margen de rechazos del {pct_rechazado_finan:.2f}%.**")

    # Tabla Solo Rechazados
    st.subheader("👤 Totales Rechazados por Firmante (solo rechazos por problemas financieros)")
    firm_r = df[mask_rechazo_finan].groupby("Den. Firmante").agg(Monto=("Monto", "sum"), Motivo_Rechazo=("Motivo Rechazo", lambda x: " | ".join(sorted(set(x.dropna().astype(str)))))).reset_index().sort_values("Monto", ascending=False)
    if not firm_r.empty:
        firm_r["% Concentración"] = (firm_r["Monto"] / rechazados_finan * 100)
        mostrar_tabla_estilizada(firm_r.assign(Monto=lambda x: x["Monto"].apply(fmt_monto), **{"% Concentración": lambda x: x["% Concentración"].apply(lambda v: f"{v:.2f}%")}).rename(columns={"Motivo_Rechazo": "Motivo del rechazo"}))

    # Visor (También actualizado a st.table)
    st.markdown("---"); st.subheader("🔍 Visor Rápido de Cheques por Firmante")
    f_sel = st.selectbox("Elegí un Firmante:", ["-- Seleccionar --"] + sorted(df_firmantes["Den. Firmante"].unique().tolist()))
    if f_sel != "-- Seleccionar --":
        datos_visor = preparar_datos_crudos(df_firmantes[df_firmantes["Den. Firmante"] == f_sel])
        styled_visor = datos_visor.style.set_properties(**{'font-size': '16px', 'white-space': 'nowrap'}).set_table_styles([{'selector': 'th', 'props': [('font-size', '16px')]}])
        st.table(styled_visor)

    # Sección 4 Meses
    if rechazados_finan > 0:
        st.markdown("---")
        min_date_4m = pd.Timestamp.today().normalize() - pd.DateOffset(months=4)
        df_4m = df[(pd.to_datetime(df["Fecha Acreditación"], errors="coerce") >= min_date_4m)].copy()
        mask_4m = (df_4m["Estado"] == "RECHAZADO") & df_4m["Motivo Rechazo"].str.contains(r"R01|R02|R10|R21", na=False, regex=True)
        t_4m = df_4m.loc[df_4m["Estado"]=="ACREDITADO", "Monto"].sum() + df_4m.loc[mask_4m, "Monto"].sum()
        r_4m = df_4m.loc[mask_4m, "Monto"].sum()

        st.markdown(f'<div style="font-size:24px; font-weight:bold; color:#d62728;">🔎 Foco de Riesgo: Últimos 4 Meses (Desde {min_date_4m.strftime("%d/%m/%Y")} hasta Hoy)</div>', unsafe_allow_html=True)
        if t_4m == 0:
            st.info("Durante los últimos 4 meses no ha registrado operatoria en DCPD")
        elif r_4m == 0:
            st.info(f"Durante los últimos 4 meses la operatoria en DCPD totalizó **{fmt_monto(t_4m)}**, sin registrar rechazos.")
        else:
            pct_4m = (r_4m/t_4m*100)
            st.info(f"Durante los últimos 4 meses la operatoria totalizó {fmt_monto(t_4m)} con un margen de rechazos del {pct_4m:.2f}%.")
            
            # Tabla 4M
            df_firm_4m = df_4m[(df_4m["Estado"]=="ACREDITADO") | mask_4m].copy()
            df_firm_4m["Tipo"] = df_firm_4m["Estado"].apply(lambda x: "ACREDITADO" if x == "ACREDITADO" else "RECHAZADO")
            res_4m = df_firm_4m.groupby(["Den. Firmante", "Tipo"])["Monto"].sum().unstack(fill_value=0).reset_index()
            res_4m["Total"] = res_4m.get("ACREDITADO",0) + res_4m.get("RECHAZADO",0)
            st.subheader("👤 Top Firmantes - Últimos 4 Meses")
            mostrar_tabla_estilizada(res_4m.sort_values("Total", ascending=False).assign(Total=lambda x: x["Total"].apply(fmt_monto)))
