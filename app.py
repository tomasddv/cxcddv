from __future__ import annotations

import json
import os
import re
import urllib.request
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


APP_DIR = Path(__file__).parent
DATA_PATH = APP_DIR / "data" / "dashboard-data.json"
XLSX_PATH = APP_DIR / "data" / "CXC_BEESCARE_DelValle_analisis.xlsx"
PPTX_PATH = APP_DIR / "data" / "Capacitacion_JDV_SPV_CXC_BEESCARE_GALAXIA_DelValle.pptx"
DEFAULT_CACHE_SECONDS = 300

COLORS = {
    "violet": "#7c3aed",
    "magenta": "#ec4899",
    "cyan": "#06b6d4",
    "green": "#22c55e",
    "lime": "#a3e635",
    "orange": "#f97316",
    "red": "#ef4444",
    "yellow": "#fde68a",
    "ink": "#172033",
}


st.set_page_config(
    page_title="CXC / BEESCARE - Distribuidora del Valle",
    page_icon=":bar_chart:",
    layout="wide",
    initial_sidebar_state="expanded",
)


def inject_style() -> None:
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
        html, body, [class*="css"] { font-family: Inter, Segoe UI, sans-serif; }
        .stApp {
          background:
            radial-gradient(circle at 8% 8%, rgba(124,58,237,.20), transparent 28%),
            radial-gradient(circle at 88% 4%, rgba(6,182,212,.18), transparent 30%),
            radial-gradient(circle at 55% 92%, rgba(236,72,153,.14), transparent 24%),
            linear-gradient(135deg, #f8fbff 0%, #f7f3ff 46%, #f7fffb 100%);
          color: #172033;
        }
        .hero {
          padding: 22px 26px;
          border: 1px solid rgba(255,255,255,.68);
          background: rgba(255,255,255,.62);
          backdrop-filter: blur(18px);
          border-radius: 22px;
          box-shadow: 0 24px 60px rgba(31,41,55,.12);
          margin-bottom: 18px;
        }
        .eyebrow {
          color: #7c3aed;
          font-weight: 800;
          letter-spacing: .06em;
          text-transform: uppercase;
          font-size: 12px;
        }
        .title {
          color: #172033;
          font-size: 34px;
          font-weight: 850;
          line-height: 1.05;
          margin-top: 4px;
        }
        .subtitle { color: #64748b; font-size: 15px; margin-top: 8px; }
        .kpi-card {
          border: 1px solid rgba(255,255,255,.72);
          background: rgba(255,255,255,.68);
          backdrop-filter: blur(18px);
          border-radius: 18px;
          padding: 18px 18px 16px;
          min-height: 116px;
          box-shadow: 0 16px 45px rgba(31,41,55,.10);
          transition: transform .18s ease, box-shadow .18s ease;
        }
        .kpi-card:hover { transform: translateY(-2px); box-shadow: 0 20px 55px rgba(31,41,55,.16); }
        .kpi-label { color: #64748b; font-weight: 700; font-size: 13px; }
        .kpi-value { color: #111827; font-size: 32px; font-weight: 850; margin-top: 7px; }
        .kpi-note { color: #64748b; font-size: 12px; margin-top: 4px; }
        .soft-panel {
          border: 1px solid rgba(255,255,255,.70);
          background: rgba(255,255,255,.58);
          backdrop-filter: blur(16px);
          border-radius: 20px;
          padding: 18px;
          box-shadow: 0 18px 50px rgba(31,41,55,.10);
        }
        .section-title { font-size: 20px; font-weight: 850; color: #172033; margin-bottom: 8px; }
        .badge {
          display: inline-flex;
          align-items: center;
          gap: 6px;
          border-radius: 999px;
          padding: 6px 10px;
          font-size: 12px;
          font-weight: 800;
          background: rgba(124,58,237,.10);
          color: #5b21b6;
          border: 1px solid rgba(124,58,237,.18);
        }
        .footer-signature {
          text-align: right;
          color: rgba(15,23,42,.38);
          font-weight: 800;
          padding: 24px 4px 4px;
          transition: color .2s ease;
        }
        .footer-signature:hover { color: rgba(124,58,237,.72); }
        div[data-testid="stMetric"] {
          border-radius: 16px;
          padding: 12px;
          background: rgba(255,255,255,.52);
          border: 1px solid rgba(255,255,255,.70);
        }
        .stDownloadButton button, .stButton button {
          border-radius: 12px;
          border: 0;
          background: linear-gradient(135deg, #7c3aed, #ec4899);
          color: white;
          font-weight: 800;
          box-shadow: 0 12px 28px rgba(124,58,237,.22);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def secret_or_env(name: str, default: str = "") -> str:
    try:
        value = st.secrets.get(name, "")
    except Exception:
        value = ""
    return str(value or os.environ.get(name, default) or "").strip()


def normalize_drive_url(url: str) -> str:
    url = (url or "").strip()
    if not url:
        return ""
    if "drive.google.com/drive/folders/" in url:
        raise ValueError("DATA_URL debe ser el link del archivo dashboard-data.json, no el link de la carpeta de Drive.")
    match = re.search(r"/d/([A-Za-z0-9_-]+)", url) or re.search(r"[?&]id=([A-Za-z0-9_-]+)", url)
    if "drive.google.com" in url and match:
        return f"https://drive.google.com/uc?export=download&id={match.group(1)}"
    return url


@st.cache_data(show_spinner=False, ttl=DEFAULT_CACHE_SECONDS)
def load_remote_data(source_url: str) -> dict:
    request = urllib.request.Request(
        normalize_drive_url(source_url),
        headers={"User-Agent": "Mozilla/5.0 Streamlit CXC"},
    )
    with urllib.request.urlopen(request, timeout=25) as response:
        payload = response.read().decode("utf-8-sig")
    return json.loads(payload)


@st.cache_data(show_spinner=False, ttl=60)
def load_local_data(file_mtime: float) -> dict:
    with DATA_PATH.open("r", encoding="utf-8") as fh:
        return json.load(fh)


def load_data() -> tuple[dict, str]:
    source_url = secret_or_env("DATA_URL")
    if source_url:
        try:
            return load_remote_data(source_url), "Google Drive"
        except Exception as exc:
            st.warning(f"No pude leer DATA_URL desde Drive. Uso la copia local. Detalle: {exc}")
    return load_local_data(DATA_PATH.stat().st_mtime), "Archivo local GitHub"


def as_df(data: dict, key: str) -> pd.DataFrame:
    return pd.DataFrame(data.get(key, []))


def pct(value: float | int | None, decimals: int = 1) -> str:
    value = 0 if value is None else float(value)
    return f"{value:.{decimals}f}%".replace(".", ",")


def status_for(row: pd.Series) -> str:
    if bool(row.get("Cerrado")) and bool(row.get("DentroSLA")):
        return "Cerrado dentro SLA"
    if bool(row.get("Cerrado")) and bool(row.get("FueraSLA")):
        return "Cerrado fuera SLA"
    if bool(row.get("PendienteVencido")):
        return "Pendiente vencido"
    return "Pendiente dentro SLA"


def csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, sep=";").encode("utf-8-sig")


def make_kpi(label: str, value: str, note: str = "", color: str = "#7c3aed") -> None:
    st.markdown(
        f"""
        <div class="kpi-card" style="border-top: 4px solid {color};">
          <div class="kpi-label">{label}</div>
          <div class="kpi-value">{value}</div>
          <div class="kpi-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def filter_tickets(tickets: pd.DataFrame, source_label: str) -> pd.DataFrame:
    with st.sidebar:
        st.markdown("### Datos")
        st.caption(f"Fuente actual: {source_label}")
        if st.button("Actualizar / limpiar cache", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
        st.divider()
        st.markdown("### Filtros")
        months = ["Todos"] + sorted(tickets["Mes"].dropna().unique().tolist())
        selected_month = st.selectbox("Mes", months)

        statuses = ["Todos"] + sorted(tickets["EstadoOperativo"].dropna().unique().tolist())
        selected_status = st.selectbox("Estado SLA", statuses)

        scopes = ["Todos"] + sorted(tickets["Alcance"].dropna().unique().tolist())
        selected_scope = st.selectbox("Alcance", scopes)

        reasons = ["Todos"] + sorted(tickets["Motivo"].dropna().unique().tolist())
        selected_reason = st.selectbox("Motivo", reasons)

        search = st.text_input("Buscar cliente / ticket / submotivo", "")

    out = tickets.copy()
    if selected_month != "Todos":
        out = out[out["Mes"] == selected_month]
    if selected_status != "Todos":
        out = out[out["EstadoOperativo"] == selected_status]
    if selected_scope != "Todos":
        out = out[out["Alcance"] == selected_scope]
    if selected_reason != "Todos":
        out = out[out["Motivo"] == selected_reason]
    if search.strip():
        haystack = out[["Ticket", "ClienteId", "Motivo", "Submotivo", "Estado"]].fillna("").astype(str).agg(" ".join, axis=1)
        out = out[haystack.str.lower().str.contains(search.strip().lower(), regex=False)]
    return out


def plot_monthly(monthly: pd.DataFrame) -> None:
    monthly = monthly.copy()
    monthly["ON TIME %"] = monthly["OnTime"] * 100
    monthly["Adopcion CXC %"] = monthly["AdopcionPct"] * 100

    fig = go.Figure()
    fig.add_trace(go.Bar(x=monthly["Mes"], y=monthly["Total"], name="Tickets", marker_color=COLORS["cyan"], opacity=.45))
    fig.add_trace(go.Scatter(x=monthly["Mes"], y=monthly["ON TIME %"], name="ON TIME", mode="lines+markers", line=dict(color=COLORS["violet"], width=4)))
    fig.add_trace(go.Scatter(x=monthly["Mes"], y=monthly["Adopcion CXC %"], name="Adopcion CXC", mode="lines+markers", line=dict(color=COLORS["green"], width=4)))
    fig.update_layout(
        height=390,
        margin=dict(l=10, r=10, t=30, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(255,255,255,.35)",
        legend=dict(orientation="h", y=1.10),
        yaxis=dict(title_text="Tickets / porcentaje"),
    )
    st.plotly_chart(fig, use_container_width=True)


def plot_bar(df: pd.DataFrame, name_col: str, value_col: str, title: str, colors: list[str]) -> None:
    if df.empty:
        st.info("Sin datos para mostrar con los filtros actuales.")
        return
    fig = px.bar(df, x=value_col, y=name_col, orientation="h", title=title, color=name_col, color_discrete_sequence=colors)
    fig.update_layout(
        height=390,
        margin=dict(l=10, r=10, t=45, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(255,255,255,.35)",
        showlegend=False,
        yaxis=dict(autorange="reversed"),
    )
    st.plotly_chart(fig, use_container_width=True)


def action_plan_editor(plan_clientes: pd.DataFrame, top5: pd.DataFrame) -> pd.DataFrame:
    base = top5.copy()
    if base.empty:
        base = plan_clientes.copy()
    wanted = [
        "Cliente",
        "Nombre",
        "Mes",
        "Motivo",
        "Submotivo",
        "Prioridad",
        "AccionSugerida",
        "Responsable",
        "FechaCompromiso",
        "AccionRealizada",
        "ComentarioSeguimiento",
        "Estado",
        "ProximoSeguimiento",
    ]
    for col in wanted:
        if col not in base.columns:
            base[col] = ""
    if "EstadoSugerido" in base.columns:
        base["Estado"] = base["Estado"].where(base["Estado"].astype(str).str.len() > 0, base["EstadoSugerido"])
    if "Responsable" not in base.columns or base["Responsable"].astype(str).eq("").all():
        base["Responsable"] = "JDV / SPV"

    edited = st.data_editor(
        base[wanted],
        use_container_width=True,
        height=430,
        num_rows="dynamic",
        column_config={
            "AccionSugerida": st.column_config.TextColumn("Accion sugerida", width="large"),
            "AccionRealizada": st.column_config.TextColumn("Accion realizada", width="large"),
            "ComentarioSeguimiento": st.column_config.TextColumn("Comentario seguimiento", width="large"),
            "Estado": st.column_config.SelectboxColumn("Estado", options=["Pendiente", "En curso", "Cerrado", "Requiere seguimiento"]),
            "FechaCompromiso": st.column_config.DateColumn("Fecha compromiso"),
            "ProximoSeguimiento": st.column_config.DateColumn("Proximo seguimiento"),
        },
        key="planes_accion",
    )
    return edited


def main() -> None:
    inject_style()
    data, source_label = load_data()

    tickets = as_df(data, "tickets")
    monthly = as_df(data, "monthly")
    plan_motivos = as_df(data, "planMotivos")
    plan_clientes = as_df(data, "planClientes")
    top5 = as_df(data, "top5Criticos")
    checklist = as_df(data, "auditChecklist")
    riesgo = as_df(data, "riesgoTickets")

    tickets["EstadoOperativo"] = tickets.apply(status_for, axis=1)
    filtered = filter_tickets(tickets, source_label)

    st.markdown(
        f"""
        <div class="hero">
          <div class="eyebrow">Manual GALAXIA · Nivel 1 · {data.get("distribuidor", "")}</div>
          <div class="title">Dashboard CXC / BEESCARE</div>
          <div class="subtitle">Periodo analizado: <b>{data.get("periodo", "")}</b> · Generado: {data.get("generado", "")}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    kpis = data.get("kpis", {})
    closed = filtered[filtered["Cerrado"] == True]
    inside = closed[closed["DentroSLA"] == True]
    on_time_filtered = (len(inside) / len(closed) * 100) if len(closed) else 0

    cols = st.columns(6)
    with cols[0]:
        make_kpi("Tickets filtrados", f"{len(filtered):,}".replace(",", "."), f"Total base: {kpis.get('totalTickets', 0)}", COLORS["cyan"])
    with cols[1]:
        make_kpi("ON TIME", pct(on_time_filtered), f"Acumulado: {pct(kpis.get('onTimeAcumulado', 0))}", COLORS["green"])
    with cols[2]:
        make_kpi("Dentro SLA", str(int(filtered["DentroSLA"].sum())), "Cerrados dentro SLA", COLORS["lime"])
    with cols[3]:
        make_kpi("Fuera SLA", str(int(filtered["FueraSLA"].sum())), "Prioridad media", COLORS["red"])
    with cols[4]:
        make_kpi("Pendientes", str(int(filtered["Pendiente"].sum())), "Dentro o vencidos", COLORS["yellow"])
    with cols[5]:
        make_kpi("+10 dias", str(int(filtered["RiesgoMasivo"].sum())), "Riesgo cierre masivo", COLORS["orange"])

    tab_resumen, tab_tickets, tab_criticos, tab_planes, tab_auditoria, tab_descargas = st.tabs(
        ["Resumen", "Tickets", "Criticos", "Planes de accion", "Auditoria 100%", "Descargas"]
    )

    with tab_resumen:
        c1, c2 = st.columns([1.35, 1])
        with c1:
            st.markdown('<div class="soft-panel"><div class="section-title">Evolucion mensual</div>', unsafe_allow_html=True)
            plot_monthly(monthly)
            st.markdown("</div>", unsafe_allow_html=True)
        with c2:
            status_counts = filtered["EstadoOperativo"].value_counts().reset_index()
            status_counts.columns = ["Estado", "Cantidad"]
            st.markdown('<div class="soft-panel"><div class="section-title">Estado operativo SLA</div>', unsafe_allow_html=True)
            fig = px.pie(status_counts, names="Estado", values="Cantidad", hole=.55, color_discrete_sequence=[COLORS["green"], COLORS["red"], COLORS["yellow"], COLORS["orange"]])
            fig.update_layout(height=390, margin=dict(l=8, r=8, t=20, b=8), paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)

        c3, c4 = st.columns(2)
        with c3:
            reason_counts = filtered["Motivo"].fillna("Sin dato").value_counts().head(10).reset_index()
            reason_counts.columns = ["Motivo", "Cantidad"]
            plot_bar(reason_counts, "Motivo", "Cantidad", "Top motivos", [COLORS["violet"], COLORS["magenta"], COLORS["cyan"], COLORS["orange"]])
        with c4:
            scope_counts = filtered["Alcance"].fillna("Sin dato").value_counts().reset_index()
            scope_counts.columns = ["Alcance", "Cantidad"]
            plot_bar(scope_counts, "Alcance", "Cantidad", "Corresponde / No corresponde", [COLORS["green"], COLORS["orange"], COLORS["red"]])

    with tab_tickets:
        st.markdown('<div class="section-title">Tickets filtrados</div>', unsafe_allow_html=True)
        st.dataframe(filtered, use_container_width=True, height=560)
        st.download_button("Descargar tickets filtrados CSV", csv_bytes(filtered), "tickets_filtrados_cxc.csv", "text/csv")

    with tab_criticos:
        c1, c2, c3 = st.columns(3)
        c1.metric("Criticos marzo", kpis.get("criticosMarzo", 0))
        c2.metric("Criticos abril", kpis.get("criticosAbril", 0))
        c3.metric("Criticos mayo", kpis.get("criticosMayo", 0))
        c4, c5, c6 = st.columns(3)
        c4.metric(f"Recurrentes {kpis.get('mesCriticoAnterior', 'Abril')}/{kpis.get('mesCriticoVigente', 'Mayo')}", kpis.get("recurrentes", 0))
        c5.metric(f"Nuevos {kpis.get('mesCriticoVigente', 'Mayo')}", kpis.get("nuevosMayo", 0))
        c6.metric(f"Recuperados {kpis.get('mesCriticoVigente', 'Mayo')}", kpis.get("recuperados", 0))
        st.markdown('<div class="section-title">Top clientes criticos para seguimiento</div>', unsafe_allow_html=True)
        st.dataframe(top5, use_container_width=True, height=320)
        st.markdown('<div class="section-title">Tickets en riesgo de cierre masivo</div>', unsafe_allow_html=True)
        st.dataframe(riesgo, use_container_width=True, height=260)

    with tab_planes:
        st.markdown('<span class="badge">Evidencia editable para JDV / SPV</span>', unsafe_allow_html=True)
        st.markdown("#### Planes de accion para clientes criticos")
        edited = action_plan_editor(plan_clientes, top5)
        st.download_button("Descargar planes completados CSV", csv_bytes(edited), "planes_accion_cxc_completados.csv", "text/csv")
        st.markdown("#### Planes sugeridos por motivo")
        st.dataframe(plan_motivos, use_container_width=True, height=360)

    with tab_auditoria:
        st.markdown('<div class="section-title">Checklist Auditoria 100%</div>', unsafe_allow_html=True)
        st.dataframe(checklist, use_container_width=True, height=420)
        st.success("Resultado objetivo: CUMPLE NIVEL 1 - ALCANCE 100% con capacitacion, indicadores, alcance, clientes criticos y planes de accion documentados.")

    with tab_descargas:
        st.markdown('<div class="section-title">Archivos de evidencia</div>', unsafe_allow_html=True)
        st.download_button("Descargar JSON de datos", DATA_PATH.read_bytes(), "dashboard-data.json", "application/json")
        if XLSX_PATH.exists():
            st.download_button("Descargar Excel de analisis", XLSX_PATH.read_bytes(), XLSX_PATH.name)
        if PPTX_PATH.exists():
            st.download_button("Descargar PPT capacitacion", PPTX_PATH.read_bytes(), PPTX_PATH.name)
        st.info("Para actualizar la app en GitHub, reemplaza `data/dashboard-data.json` con la ultima version generada.")

    st.markdown('<div class="footer-signature">by QπU</div>', unsafe_allow_html=True)


if __name__ == "__main__":
    main()
