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
# Mostrar tabla (Fuente grande + Índice 1 en adelante + Scroll altura fija)
# -----------------------------
def mostrar_tabla_estilizada(df_to_show, height=400):
    df_to_show = df_to_show.copy()
    
    # Asignar explícitamente el índice del 1 en adelante
    df_to_show.index = range(1, len(df_to_show) + 1)
            
    # Agrandar la fuente usando propiedades nativas
    styled = df_to_show.style.set_properties(**{
        'font-size': '15px'
    }).set_table_styles([
        {'selector': 'th', 'props': [('font-size', '15px')]}
    ])
    
    # Pasamos el height para limitar lo que se ve y forzar el scroll
    st.dataframe(styled, use_container_width=True, height=height)

# -----------------------------
# Filtro y preparador para Datos Crudos
# -----------------------------
def preparar_datos_crudos(df_in):
    mapeo_columnas = {
        "Den.Socio": "Den. Socio",
        "Den. Socio": "Den. Socio",
        "Tipo op.": "Tipo Op.",
        "Tipo Op.": "Tipo Op.",
        "CUI": "CUIT",
        "CUIT": "CUIT",
        "Den.Firmante": "Den. Firmante",
        "Den. Firmante": "Den. Firmante",
        "Monto": "Monto",
        "Fecha Acreditación": "Fecha Acreditación",
        "Estado": "Estado",
        "Motivo Rechazo": "Motivo Rechazo"
    }
    
    cols_encontradas = []
    renombres = {}
    
    for col_original in df_in.columns:
        if col_original in mapeo_columnas:
            cols_encontradas.append(col_original)
            if col_original != mapeo_columnas[col_original]:
                renombres[col_original] = mapeo_columnas[col_original]
                
    df_out = df_in[cols_encontradas].copy()
    df_out.rename(columns=renombres, inplace=True)
    
    if "Fecha Acreditación" in df_out.columns:
        df_out["Fecha Acreditación"] = pd.to_datetime(df_out["Fecha Acreditación"], errors='coerce').dt.strftime('%d/%m/%Y')
    if "Monto" in df_out.columns:
        df_out["Monto"] = df_out["Monto"].apply(fmt_monto)
        
    orden_ideal = ["Den. Socio", "Tipo Op.", "CUIT", "Den. Firmante", "Monto", "Fecha Acreditación", "Estado", "Motivo Rechazo"]
    orden_final = [col for col in orden_ideal if col in df_out.columns]
    
    df_final = df_out[orden_final]
    df_final.index = range(1, len(df_final) + 1)
    
    return df_final


# -----------------------------
# Main
# -----------------------------
if uploaded_file:
    try:
        df = leer_archivo_robusto(uploaded_file)
    except Exception as e:
        st.error(f"No se pudo leer el archivo: {e}")
        st.stop()

    df.columns = df.columns.astype(str).str.strip().str.replace('"', '')

    required_cols = ["Tipo Op.", "Monto Acreditado / Rechazado", "Den. Firmante", "Fecha Acreditación"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"Faltan columnas necesarias: {missing}")
        st.stop()

    df["Tipo Op."] = df["Tipo Op."].astype(str).str.strip().str.replace('"', '')
    df = df[df["Tipo Op."] == "CO"].copy()

    if "Motivo Rechazo" not in df.columns:
        df["Motivo Rechazo"] = ""
    df["Motivo Rechazo"] = df["Motivo Rechazo"].astype(str)
    df["Den. Firmante"] = df["Den. Firmante"].astype(str).str.strip().str.replace('"', '')

    df["_monto_texto"] = df["Monto Acreditado / Rechazado"].astype(str)
    df["Estado"] = df["_monto_texto"].str.upper().apply(
        lambda x: "ACREDITADO" if "ACREDITADO" in x else ("RECHAZADO" if "RECHAZADO" in x else "OTRO")
    )
    df["Monto"] = df["_monto_texto"].apply(parse_amount_from_text)

    # --- Totales ---
    total_acreditado = df.loc[df["Estado"] == "ACREDITADO", "Monto"].sum()
    mask_rechazo = df["Estado"] == "RECHAZADO"
    
    # Buscar tanto R10 como R21
    mask_r10_r21 = mask_rechazo & df["Motivo Rechazo"].str.contains(r"R10|R21", na=False, regex=True)
    rechazados_r10_r21 = df.loc[mask_r10_r21, "Monto"].sum()

    # Total operado = acreditado + rechazado R10/R21
    total_operado = total_acreditado + rechazados_r10_r21

    pct_acreditado = (total_acreditado / total_operado * 100) if total_operado > 0 else 0.0
    pct_r10_r21 = (rechazados_r10_r21 / total_operado * 100) if total_operado > 0 else 0.0

    # -----------------------------
    # Socio + Lapso temporal
    # -----------------------------
    try:
        fechas = pd.to_datetime(df["Fecha Acreditación"], errors="coerce")
        socio = df["Den. Socio"].dropna().unique()[0] if "Den. Socio" in df.columns else ""
        if fechas.notna().any():
            st.markdown(
                f"""
                <div style="font-size:30px; font-weight:bold; color:green; margin-bottom:10px;">
                👥 {socio}
                </div>
                <div style="font-size:20px; font-weight:bold; color:#444;">
                📅 Detalle de cheques de pago diferido descontados (DCPD) con vencimiento operado entre 
                <span>{fechas.min().strftime('%d/%m/%Y')}</span> y 
                <span>{fechas.max().strftime('%d/%m/%Y')}</span>
                </div>
                """,
                unsafe_allow_html=True
            )
    except Exception as e:
        st.warning(f"No se pudo procesar fechas o socio: {e}")

    # -----------------------------
    # Totales + Cantidad de cheques
    # -----------------------------
    cant_total_operado = len(df[(df["Estado"] == "ACREDITADO") | mask_r10_r21])
    cant_acreditados = len(df[df["Estado"] == "ACREDITADO"])
    cant_r10_r21 = len(df[mask_r10_r21])

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("📦 Total Operado", fmt_monto(total_operado))
        st.markdown(f"<div style='font-size:14px; color:gray;'>Cantidad de cheques: <b>{cant_total_operado}</b></div>", unsafe_allow_html=True)

    with col2:
        st.metric("💰 Total Acreditado", fmt_monto(total_acreditado))
        st.markdown(f"<div style='font-size:14px; color:gray;'>Cantidad de cheques: <b>{cant_acreditados}</b></div>", unsafe_allow_html=True)

    with col3:
        st.metric("❌ Rechazados (R10/R21)", fmt_monto(rechazados_r10_r21))
        st.markdown(f"<div style='font-size:14px; color:gray;'>Cantidad de cheques: <b>{cant_r10_r21}</b></div>", unsafe_allow_html=True)

    # -----------------------------
    # % Acreditado y Rechazado
    # -----------------------------
    colA, colB = st.columns(2)
    colA.markdown(f"<div style='font-size:26px; font-weight:bold; color:green;'>✅ % Acreditado: {pct_acreditado:.2f}%</div>", unsafe_allow_html=True)
    colB.markdown(f"<div style='font-size:26px; font-weight:bold; color:green;'>❌ % Rechazados (R10/R21): {pct_r10_r21:.2f}%</div>", unsafe_allow_html=True)

    # -----------------------------
    # Tabla de firmantes (Acreditados + R10/R21)
    # -----------------------------
    df_firmantes = df[(df["Estado"] == "ACREDITADO") | (mask_r10_r21)].copy()
    df_firmantes["Tipo"] = df_firmantes.apply(
        lambda row: "ACREDITADO" if row["Estado"] == "ACREDITADO" else "RECHAZADO R10/R21", axis=1
    )

    firmantes = (
        df_firmantes.groupby(["Den. Firmante", "Tipo"])["Monto"]
        .sum()
        .unstack(fill_value=0)
        .reset_index()
    )

    if "ACREDITADO" not in firmantes.columns: firmantes["ACREDITADO"] = 0
    if "RECHAZADO R10/R21" not in firmantes.columns: firmantes["RECHAZADO R10/R21"] = 0

    firmantes["Total_Firmante"] = firmantes["ACREDITADO"] + firmantes["RECHAZADO R10/R21"]
    firmantes["% Concentración"] = firmantes["Total_Firmante"] / total_operado * 100

    firmantes = firmantes.sort_values("Total_Firmante", ascending=False).reset_index(drop=True)

    firmantes["ACREDITADO"] = firmantes["ACREDITADO"].apply(fmt_monto)
    firmantes["RECHAZADO R10/R21"] = firmantes["RECHAZADO R10/R21"].apply(fmt_monto)
    firmantes["Total_Firmante"] = firmantes["Total_Firmante"].apply(fmt_monto)
    firmantes["% Concentración"] = firmantes["% Concentración"].apply(lambda x: f"{x:.2f}%")

    # MUESTRA TODOS LOS FIRMANTES GLOBALES CON SCROLL (TÍTULO SOLICITADO MANTENIDO)
    st.subheader("👤 Top 10 Firmantes (sobre total operado)")
    mostrar_tabla_estilizada(firmantes, height=400)

    st.download_button(
        "⬇️ Descargar reporte firmantes (ACR + R10/R21) CSV Completo",
        firmantes.to_csv(index=False).encode("utf-8"),
        "reporte_firmantes_total.csv",
        "text/csv"
    )

    # -----------------------------
    # Tabla de firmantes SOLO R10/R21 (Agregado Motivo)
    # -----------------------------
    firmantes_r10_r21 = (
        df[mask_r10_r21].groupby("Den. Firmante")
        .agg(
            Monto=("Monto", "sum"),
            Motivo_Rechazo=("Motivo Rechazo", lambda x: " | ".join(sorted(set(x.dropna().astype(str).str.strip()))))
        )
        .reset_index()
        .rename(columns={"Motivo_Rechazo": "Motivo del rechazo"})
        .sort_values("Monto", ascending=False)
    )

    firmantes_r10_r21["% Concentración"] = firmantes_r10_r21["Monto"] / rechazados_r10_r21 * 100 if rechazados_r10_r21 > 0 else 0
    firmantes_r10_r21["Monto"] = firmantes_r10_r21["Monto"].apply(fmt_monto)
    firmantes_r10_r21["% Concentración"] = firmantes_r10_r21["% Concentración"].apply(lambda x: f"{x:.2f}%")
    
    firmantes_r10_r21 = firmantes_r10_r21[["Den. Firmante", "Monto", "% Concentración", "Motivo del rechazo"]]

    # MUESTRA TODOS LOS RECHAZOS GLOBALES CON SCROLL
    st.subheader("👤 Totales por Firmante (SOLO Rechazados R10 y R21)")
    mostrar_tabla_estilizada(firmantes_r10_r21, height=400)

    st.download_button(
        "⬇️ Descargar reporte firmantes SOLO R10/R21 CSV Completo",
        firmantes_r10_r21.to_csv(index=False).encode("utf-8"),
        "reporte_firmantes_r10_r21.csv",
        "text/csv"
    )

    # -----------------------------
    # Datos crudos filtrados (GLOBAL)
    # -----------------------------
    with st.expander("🗂️ Ver datos crudos filtrados (Tipo Op. = CO, ACR + R10/R21)"):
        st.dataframe(preparar_datos_crudos(df_firmantes), use_container_width=True)

    # =========================================================================
    # SECCIÓN: ANÁLISIS DE LOS ÚLTIMOS 4 MESES 
    # =========================================================================
    if rechazados_r10_r21 > 0:
        st.markdown("---")
        
        fechas_dt = pd.to_datetime(df["Fecha Acreditación"], errors="coerce")
        fecha_actual = pd.Timestamp.today().normalize()
        min_date_4m = fecha_actual - pd.DateOffset(months=4)
        
        mask_fechas_4m = (fechas_dt >= min_date_4m) & (fechas_dt <= fecha_actual)
        df_4m = df[mask_fechas_4m].copy()
        
        mask_r10_r21_4m = (df_4m["Estado"] == "RECHAZADO") & df_4m["Motivo Rechazo"].str.contains(r"R10|R21", na=False, regex=True)
        total_acreditado_4m = df_4m.loc[df_4m["Estado"] == "ACREDITADO", "Monto"].sum()
        rechazados_r10_r21_4m = df_4m.loc[mask_r10_r21_4m, "Monto"].sum()
        total_operado_4m = total_acreditado_4m + rechazados_r10_r21_4m
        
        pct_acreditado_4m = (total_acreditado_4m / total_operado_4m * 100) if total_operado_4m > 0 else 0.0
        pct_r10_r21_4m = (rechazados_r10_r21_4m / total_operado_4m * 100) if total_operado_4m > 0 else 0.0

        if total_operado_4m > 0:
            st.markdown(
                f"""
                <div style="font-size:24px; font-weight:bold; color:#d62728; margin-bottom:10px;">
                🔎 Foco de Riesgo: Últimos 4 Meses (Desde {min_date_4m.strftime('%d/%m/%Y')} hasta Hoy)
                </div>
                """, unsafe_allow_html=True
            )
            
            df_4m_fechas = df_4m.copy()
            df_4m_fechas["Fecha Acreditación"] = pd.to_datetime(df_4m_fechas["Fecha Acreditación"], errors="coerce")
            df_4m_fechas["Mes_Anio"] = df_4m_fechas["Fecha Acreditación"].dt.strftime('%m-%Y')
            rechazos_por_mes = df_4m_fechas[mask_r10_r21_4m].groupby("Mes_Anio")["Monto"].sum()
            
            if not rechazos_por_mes.empty and rechazados_r10_r21_4m > 0:
                meses_pct = (rechazos_por_mes / rechazados_r10_r21_4m * 100).sort_values(ascending=False)
                meses_pct_int = meses_pct.round().astype(int)
                
                diferencia = 100 - meses_pct_int.sum()
                if diferencia != 0 and len(meses_pct_int) > 0:
                    meses_pct_int.iloc[0] += diferencia

                str_meses = ", ".join([f"{mes} ({pct}%)" for mes, pct in meses_pct_int.items()])
            else:
                str_meses = "Ninguno (0%)"

            st.info(f"**Durante los últimos 4 meses la operatoria en DCPD totalizó {fmt_monto(total_operado_4m)}, con un margen de rechazos del {pct_r10_r21_4m:.2f}%, concentrados en los meses de {str_meses}.**")

            cant_total_operado_4m = len(df_4m[(df_4m["Estado"] == "ACREDITADO") | mask_r10_r21_4m])
            cant_acreditados_4m = len(df_4m[df_4m["Estado"] == "ACREDITADO"])
            cant_r10_r21_4m = len(df_4m[mask_r10_r21_4m])

            col1_4m, col2_4m, col3_4m = st.columns(3)
            with col1_4m:
                st.metric("📦 Total Operado (4M)", fmt_monto(total_operado_4m))
                st.markdown(f"<div style='font-size:14px; color:gray;'>Cantidad de cheques: <b>{cant_total_operado_4m}</b></div>", unsafe_allow_html=True)
            with col2_4m:
                st.metric("💰 Total Acreditado (4M)", fmt_monto(total_acreditado_4m))
                st.markdown(f"<div style='font-size:14px; color:gray;'>Cantidad de cheques: <b>{cant_acreditados_4m}</b></div>", unsafe_allow_html=True)
            with col3_4m:
                st.metric("❌ Rechazados (R10/R21) (4M)", fmt_monto(rechazados_r10_r21_4m))
                st.markdown(f"<div style='font-size:14px; color:gray;'>Cantidad de cheques: <b>{cant_r10_r21_4m}</b></div>", unsafe_allow_html=True)

            colA_4m, colB_4m = st.columns(2)
            colA_4m.markdown(f"<div style='font-size:26px; font-weight:bold; color:green;'>✅ % Acreditado: {pct_acreditado_4m:.2f}%</div>", unsafe_allow_html=True)
            colB_4m.markdown(f"<div style='font-size:26px; font-weight:bold; color:green;'>❌ % Rechazados (R10/R21): {pct_r10_r21_4m:.2f}%</div>", unsafe_allow_html=True)

            # Tabla de firmantes (Acreditados + R10/R21) - 4M
            df_firmantes_4m = df_4m[(df_4m["Estado"] == "ACREDITADO") | mask_r10_r21_4m].copy()
            df_firmantes_4m["Tipo"] = df_firmantes_4m.apply(
                lambda row: "ACREDITADO" if row["Estado"] == "ACREDITADO" else "RECHAZADO R10/R21", axis=1
            )

            firmantes_4m = df_firmantes_4m.groupby(["Den. Firmante", "Tipo"])["Monto"].sum().unstack(fill_value=0).reset_index()
            if "ACREDITADO" not in firmantes_4m.columns: firmantes_4m["ACREDITADO"] = 0
            if "RECHAZADO R10/R21" not in firmantes_4m.columns: firmantes_4m["RECHAZADO R10/R21"] = 0

            firmantes_4m["Total_Firmante"] = firmantes_4m["ACREDITADO"] + firmantes_4m["RECHAZADO R10/R21"]
            firmantes_4m["% Concentración"] = firmantes_4m["Total_Firmante"] / total_operado_4m * 100
            firmantes_4m = firmantes_4m.sort_values("Total_Firmante", ascending=False).reset_index(drop=True)

            firmantes_4m_disp = firmantes_4m.copy()
            firmantes_4m_disp["ACREDITADO"] = firmantes_4m_disp["ACREDITADO"].apply(fmt_monto)
            firmantes_4m_disp["RECHAZADO R10/R21"] = firmantes_4m_disp["RECHAZADO R10/R21"].apply(fmt_monto)
            firmantes_4m_disp["Total_Firmante"] = firmantes_4m_disp["Total_Firmante"].apply(fmt_monto)
            firmantes_4m_disp["% Concentración"] = firmantes_4m_disp["% Concentración"].apply(lambda x: f"{x:.2f}%")

            # MUESTRA TODOS LOS FIRMANTES 4 MESES CON SCROLL (TÍTULO SOLICITADO MANTENIDO)
            st.subheader("👤 Top 10 Firmantes (sobre total operado) - Últimos 4 Meses")
            mostrar_tabla_estilizada(firmantes_4m_disp, height=400)

            # Tabla de firmantes SOLO R10/R21 - 4M (Agregado Motivo)
            firmantes_r10_r21_4m = (
                df_4m[mask_r10_r21_4m].groupby("Den. Firmante")
                .agg(
                    Monto=("Monto", "sum"),
                    Motivo_Rechazo=("Motivo Rechazo", lambda x: " | ".join(sorted(set(x.dropna().astype(str).str.strip()))))
                )
                .reset_index()
                .rename(columns={"Motivo_Rechazo": "Motivo del rechazo"})
                .sort_values("Monto", ascending=False)
            )

            if not firmantes_r10_r21_4m.empty:
                firmantes_r10_r21_4m["% Concentración"] = firmantes_r10_r21_4m["Monto"] / rechazados_r10_r21_4m * 100 if rechazados_r10_r21_4m > 0 else 0
                firmantes_r10_r21_4m["Monto"] = firmantes_r10_r21_4m["Monto"].apply(fmt_monto)
                firmantes_r10_r21_4m["% Concentración"] = firmantes_r10_r21_4m["% Concentración"].apply(lambda x: f"{x:.2f}%")
                
                firmantes_r10_r21_4m = firmantes_r10_r21_4m[["Den. Firmante", "Monto", "% Concentración", "Motivo del rechazo"]]

                # MUESTRA TODOS LOS RECHAZOS 4 MESES CON SCROLL
                st.subheader("👤 Totales por Firmante (SOLO Rechazados R10 y R21) - Últimos 4 Meses")
                mostrar_tabla_estilizada(firmantes_r10_r21_4m, height=400)
            else:
                st.success("No hay rechazos R10 ni R21 en los últimos 4 meses.")

            # -----------------------------
            # Datos crudos filtrados (ÚLTIMOS 4 MESES)
            # -----------------------------
            with st.expander("🗂️ Ver datos crudos filtrados (Tipo Op. = CO, ACR + R10/R21) - Últimos 4 Meses"):
                st.dataframe(preparar_datos_crudos(df_firmantes_4m), use_container_width=True)
