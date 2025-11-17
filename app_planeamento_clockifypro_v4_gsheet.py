"""
App Streamlit: Planeamento — Clockify Pro v4 (CSV + Google Sheets)

• Editar horas na vista Calendário
• Botão robusto "💾 Guardar alterações" que grava em storage persistente
• Linha "Total (All)" fixa (pinned), não editável e dinâmica
• Totais por grupo, linha e coluna sempre corretos
• Vista Calendário/Lista, navegação semanal, intervalo personalizado e filtros
• Páginas: Adicionar Registos, Planeamento Semanal, Gerir Workers, Gerir Obras, Export/Backup

Storage:
- Local: CSV (workers.csv, obras.csv, records.csv)
- Online (Streamlit Cloud): Google Sheets se existirem credenciais em st.secrets

Requisitos:
    pip install streamlit pandas streamlit-aggrid gspread google-auth

Correr local:
    python -m streamlit run app_planeamento_clockifypro_v4_gsheet.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import os
from datetime import date, timedelta
from typing import List

from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, JsCode, DataReturnMode

# ==== TENTAR ATIVAR GOOGLE SHEETS ====
USE_GSHEETS = False
try:
    from google.oauth2.service_account import Credentials
    import gspread
    if "gcp_service_account" in st.secrets and "gsheet" in st.secrets and "spreadsheet_name" in st.secrets["gsheet"]:
        USE_GSHEETS = True
except Exception:
    USE_GSHEETS = False

st.set_page_config(page_title="Planeamento - Clockify Pro v4", layout="wide")
if not hasattr(st, "toast"):
    def _fallback_toast(msg, icon=None):
        st.success(msg)
    st.toast = _fallback_toast  # type: ignore

# ===== Constantes =====
DAY_COL_W = 80
TEXT_COL_W = 240
EN_WEEKDAYS = ["Mo","Tu","We","Th","Fr","Sa","Su"]
EN_MONTH = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

WORKERS_CSV = "workers.csv"
OBRAS_CSV   = "obras.csv"
RECORDS_CSV = "records.csv"

WORKERS_SHEET = "workers"
OBRAS_SHEET   = "obras"
RECORDS_SHEET = "records"

# ================= STORAGE LAYER =================

# ---------- CSV BACKEND ----------

def ensure_csv(path, cols):
    if not os.path.exists(path):
        pd.DataFrame(columns=cols).to_csv(path, index=False)

def csv_init_storage():
    ensure_csv(WORKERS_CSV, ["worker"])
    ensure_csv(OBRAS_CSV,   ["obra"])
    ensure_csv(RECORDS_CSV, ["data", "worker", "obra", "horas"])

def csv_load_workers() -> pd.DataFrame:
    if not os.path.exists(WORKERS_CSV):
        ensure_csv(WORKERS_CSV, ["worker"])
    df = pd.read_csv(WORKERS_CSV)
    if "worker" not in df.columns:
        df = pd.DataFrame(columns=["worker"])
    return df

def csv_save_workers(df: pd.DataFrame):
    out = df.copy()
    if "worker" not in out.columns:
        out["worker"] = pd.NA
    out = out[["worker"]]
    out.to_csv(WORKERS_CSV, index=False)

def csv_load_obras() -> pd.DataFrame:
    if not os.path.exists(OBRAS_CSV):
        ensure_csv(OBRAS_CSV, ["obra"])
    df = pd.read_csv(OBRAS_CSV)
    if "obra" not in df.columns:
        df = pd.DataFrame(columns=["obra"])
    return df

def csv_save_obras(df: pd.DataFrame):
    out = df.copy()
    if "obra" not in out.columns:
        out["obra"] = pd.NA
    out = out[["obra"]]
    out.to_csv(OBRAS_CSV, index=False)

def csv_read_records() -> pd.DataFrame:
    if not os.path.exists(RECORDS_CSV):
        ensure_csv(RECORDS_CSV, ["data", "worker", "obra", "horas"])
    df = pd.read_csv(RECORDS_CSV)
    for c in ["data", "worker", "obra", "horas"]:
        if c not in df.columns:
            df[c] = pd.NA
    if df.empty:
        df["date_parsed"], df["data_display"] = pd.NaT, pd.NA
        return df
    try:
        df["date_parsed"] = pd.to_datetime(df["data"], dayfirst=True, errors="coerce")
    except Exception:
        df["date_parsed"] = pd.to_datetime(df["data"], errors="coerce")
    df["horas"] = pd.to_numeric(df["horas"], errors="coerce")
    df = df.dropna(subset=["date_parsed", "worker", "obra", "horas"]).copy()
    df["data_display"] = df["date_parsed"].dt.strftime("%d/%m/%Y")
    return df

def csv_write_records(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "data_display" in out.columns:
        out["data"] = out["data_display"]
    elif "date_parsed" in out.columns:
        out["data"] = pd.to_datetime(out["date_parsed"]).dt.strftime("%d/%m/%Y")
    for c in ["data", "worker", "obra", "horas"]:
        if c not in out.columns:
            out[c] = pd.NA
    out = out[["data", "worker", "obra", "horas"]]
    out.to_csv(RECORDS_CSV, index=False)
    return csv_read_records()

# ---------- GOOGLE SHEETS BACKEND ----------

GSHEET_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

def get_gs_client():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=GSHEET_SCOPES,
    )
    client = gspread.authorize(creds)
    return client

def gs_get_sheet(title: str, header: List[str]):
    client = st.session_state.gs_client
    sh = client.open(st.secrets["gsheet"]["spreadsheet_name"])
    try:
        ws = sh.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=1000, cols=len(header))
        ws.update("A1", [header])
        return ws

    # garantir header na linha 1
    existing_header = ws.row_values(1)
    if existing_header != header:
        ws.update("A1", [header])
    return ws

def gs_load_workers() -> pd.DataFrame:
    ws = gs_get_sheet(WORKERS_SHEET, ["worker"])
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    if df.empty:
        df = pd.DataFrame(columns=["worker"])
    return df

def gs_save_workers(df: pd.DataFrame):
    ws = gs_get_sheet(WORKERS_SHEET, ["worker"])
    out = df.copy()
    if "worker" not in out.columns:
        out["worker"] = pd.NA
    out = out[["worker"]]
    ws.clear()
    rows = [out.columns.tolist()] + out.astype(str).fillna("").values.tolist()
    ws.update("A1", rows)

def gs_load_obras() -> pd.DataFrame:
    ws = gs_get_sheet(OBRAS_SHEET, ["obra"])
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    if df.empty:
        df = pd.DataFrame(columns=["obra"])
    return df

def gs_save_obras(df: pd.DataFrame):
    ws = gs_get_sheet(OBRAS_SHEET, ["obra"])
    out = df.copy()
    if "obra" not in out.columns:
        out["obra"] = pd.NA
    out = out[["obra"]]
    ws.clear()
    rows = [out.columns.tolist()] + out.astype(str).fillna("").values.tolist()
    ws.update("A1", rows)

def gs_read_records() -> pd.DataFrame:
    ws = gs_get_sheet(RECORDS_SHEET, ["data", "worker", "obra", "horas"])
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    for c in ["data", "worker", "obra", "horas"]:
        if c not in df.columns:
            df[c] = pd.NA
    if df.empty:
        df["date_parsed"], df["data_display"] = pd.NaT, pd.NA
        return df
    df["date_parsed"] = pd.to_datetime(df["data"], dayfirst=True, errors="coerce")
    df["horas"] = pd.to_numeric(df["horas"], errors="coerce")
    df = df.dropna(subset=["date_parsed", "worker", "obra", "horas"]).copy()
    df["data_display"] = df["date_parsed"].dt.strftime("%d/%m/%Y")
    return df

def gs_write_records(df: pd.DataFrame) -> pd.DataFrame:
    ws = gs_get_sheet(RECORDS_SHEET, ["data", "worker", "obra", "horas"])
    out = df.copy()
    if "data_display" in out.columns:
        out["data"] = out["data_display"]
    elif "date_parsed" in out.columns:
        out["data"] = pd.to_datetime(out["date_parsed"]).dt.strftime("%d/%m/%Y")
    for c in ["data", "worker", "obra", "horas"]:
        if c not in out.columns:
            out[c] = pd.NA
    out = out[["data", "worker", "obra", "horas"]]
    ws.clear()
    rows = [out.columns.tolist()] + out.astype(str).fillna("").values.tolist()
    ws.update("A1", rows)
    return gs_read_records()

# ---------- Dispatcher (CSV ou Sheets) ----------

def init_storage():
    if USE_GSHEETS:
        if "gs_client" not in st.session_state:
            st.session_state.gs_client = get_gs_client()
        # garante que as sheets existem
        _ = gs_get_sheet(WORKERS_SHEET, ["worker"])
        _ = gs_get_sheet(OBRAS_SHEET,   ["obra"])
        _ = gs_get_sheet(RECORDS_SHEET, ["data", "worker", "obra", "horas"])
        st.session_state.records_df = gs_read_records()
        st.session_state.workers_df = gs_load_workers()
        st.session_state.obras_df   = gs_load_obras()
    else:
        csv_init_storage()
        st.session_state.records_df = csv_read_records()
        st.session_state.workers_df = csv_load_workers()
        st.session_state.obras_df   = csv_load_obras()

def load_workers() -> pd.DataFrame:
    return st.session_state.workers_df

def save_workers(df: pd.DataFrame):
    if USE_GSHEETS:
        gs_save_workers(df)
        st.session_state.workers_df = gs_load_workers()
    else:
        csv_save_workers(df)
        st.session_state.workers_df = csv_load_workers()

def load_obras() -> pd.DataFrame:
    return st.session_state.obras_df

def save_obras(df: pd.DataFrame):
    if USE_GSHEETS:
        gs_save_obras(df)
        st.session_state.obras_df = gs_load_obras()
    else:
        csv_save_obras(df)
        st.session_state.obras_df = csv_load_obras()

def read_records() -> pd.DataFrame:
    return st.session_state.records_df

def write_records(df: pd.DataFrame):
    if USE_GSHEETS:
        st.session_state.records_df = gs_write_records(df)
    else:
        st.session_state.records_df = csv_write_records(df)

# ===== Datas =====

def monday_of_week(d: date) -> date:
    return d - timedelta(days=d.weekday())

def sunday_of_week(d: date) -> date:
    return monday_of_week(d) + timedelta(days=6)

def is_full_mon_sun(a: date, b: date) -> bool:
    return (b - a).days == 6 and a.weekday() == 0

def period_label(a: date, b: date) -> str:
    today = date.today()
    mo, su = monday_of_week(today), sunday_of_week(today)
    if is_full_mon_sun(a,b):
        if a == mo and b == su: return "This Week"
        if a == mo - timedelta(days=7): return "Last Week"
        if a == mo + timedelta(days=7): return "Next Week"
    def fmt(d):
        return f"{EN_MONTH[d.month-1]} {d.day}, {d.year}"
    return f"{fmt(a)} – {fmt(b)}"

# ===== JS helpers =====

HOURS_VALUE_FORMATTER = JsCode(
    """
    function(params){
      if (params.value == null || isNaN(params.value)) return '';
      const mins = Math.round(params.value * 60);
      const h = Math.floor(mins / 60);
      const m = mins % 60;
      return h + ':' + (m < 10 ? ('0'+m) : m);
    }
    """
)

GROUP_ROW_AGG = JsCode(
    """
    function groupRowAggNodes(nodes) {
      const result = {};
      if (!nodes || nodes.length===0) { return result; }
      nodes.forEach(n => {
        const d = n.data || {};
        for (const k in d) {
          if (typeof d[k] === 'number') {
            result[k] = (result[k] || 0) + d[k];
          }
        }
      });
      return result;
    }
    """
)

CELL_STYLE_RULES = JsCode(
    """
    function(params){
      if (params.node && params.node.rowPinned === 'bottom') {
        return {backgroundColor:'#d0d0d0', fontWeight:'700'};
      }
      if (params.node && params.node.group === true) {
        return {backgroundColor:'#ff8c0050', fontWeight:'600'};
      }
      return null;
    }
    """
)

EDITABLE_FN = JsCode(
    """
    function(params){
      if (params.node && (params.node.group === true || params.node.rowPinned === 'bottom')) return false;
      if (params.colDef && params.colDef.field === 'Total') return false;
      return true;
    }
    """
)

# ===== CSS =====
st.markdown(
    """
    <style>
      .block-container {padding-left: 1rem; padding-right: 1rem;}
      .ag-root-wrapper, .ag-root-wrapper-body, .ag-center-cols-viewport { width: 100% !important; }
      .ag-theme-balham .ag-header-cell-label { justify-content: center; }
      .ag-theme-balham .ag-pinned-bottom .ag-cell { font-weight: 700; background: #d0d0d0; }
      .ag-theme-balham .ag-row-group { font-weight: 600; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ===== Calendar DF =====

def build_calendar_dataframe(records: pd.DataFrame, start_cal: date, end_cal: date, group_by: str,
                             filter_workers: List[str], filter_obras: List[str]):
    rec = records[(records["date_parsed"] >= pd.to_datetime(start_cal)) & (records["date_parsed"] <= pd.to_datetime(end_cal))].copy()
    if filter_workers:
        rec = rec[rec["worker"].isin(filter_workers)]
    if filter_obras:
        rec = rec[rec["obra"].isin(filter_obras)]

    dates = pd.date_range(start=start_cal, end=end_cal)
    col_keys   = [d.strftime("%d/%m/%Y") for d in dates]
    col_labels = [f"{EN_WEEKDAYS[d.weekday()]}, {EN_MONTH[d.month-1]} {d.day}" for d in dates]
    rename_map = {k: l for k, l in zip(col_keys, col_labels)}
    rev_rename = {v:k for k,v in rename_map.items()}  # label -> dd/mm/yyyy

    if rec.empty:
        cols = (["Obra","Worker"] if group_by=="Project" else ["Worker","Obra"]) + col_labels + ["Total"]
        empty = pd.DataFrame(columns=cols)
        empty["row_key"] = []
        keyA, keyB = ("Obra","Worker") if group_by=="Project" else ("Worker","Obra")
        pinned = {keyA:"Total (All)", keyB:"", "row_key":"__PINNED__", **{c:0.0 for c in col_labels + ["Total"]}}
        return empty, col_labels, [pinned], rev_rename

    if group_by == "Project":
        detail = rec.pivot_table(index=["obra","worker"], columns="data_display", values="horas", aggfunc="sum").reindex(columns=col_keys, fill_value=0)
        detail = detail.rename(columns=rename_map).reset_index().rename(columns={"obra":"Obra","worker":"Worker"})
        keyA, keyB = "Obra", "Worker"
    else:
        detail = rec.pivot_table(index=["worker","obra"], columns="data_display", values="horas", aggfunc="sum").reindex(columns=col_keys, fill_value=0)
        detail = detail.rename(columns=rename_map).reset_index().rename(columns={"worker":"Worker","obra":"Obra"})
        keyA, keyB = "Worker", "Obra"

    detail[col_labels] = detail[col_labels].apply(pd.to_numeric, errors="coerce").fillna(0)
    detail["Total"] = detail[col_labels].sum(axis=1)
    detail["row_key"] = detail[keyA].astype(str) + "||" + detail[keyB].astype(str)

    detail = detail.sort_values([keyA, keyB]).reset_index(drop=True)

    totals_all = detail.drop(columns=[keyA, keyB, "row_key"]).sum(numeric_only=True)
    pinned = {keyA:"Total (All)", keyB:"", "row_key":"__PINNED__"}
    for lbl in col_labels + ["Total"]:
        pinned[lbl] = float(totals_all.get(lbl, 0))

    return detail, col_labels, [pinned], rev_rename

# ===== AgGrid renderer =====

def aggrid_calendar(df: pd.DataFrame, group_by: str, date_labels: list, pinned_bottom_rows: list,
                    expand_all: bool, editable: bool=True):
    gb = GridOptionsBuilder.from_dataframe(df)
    keyA, keyB = ("Obra","Worker") if group_by=="Project" else ("Worker","Obra")

    gb.configure_default_column(resizable=False, headerClass='header-center', cellClass='cell-right', editable=False)
    gb.configure_column(keyA, rowGroup=True, hide=True, width=TEXT_COL_W)
    gb.configure_column(keyB, width=TEXT_COL_W, editable=False)
    gb.configure_column("row_key", hide=True)

    for c in date_labels:
        gb.configure_column(
            c,
            type=["numericColumn","rightAligned"],
            aggFunc="sum",
            enableValue=True,
            width=DAY_COL_W,
            valueFormatter=HOURS_VALUE_FORMATTER,
            cellStyle=CELL_STYLE_RULES,
            editable=EDITABLE_FN
        )
    gb.configure_column(
        "Total",
        type=["numericColumn","rightAligned"],
        aggFunc="sum",
        enableValue=True,
        width=DAY_COL_W,
        pinned="right",
        valueFormatter=HOURS_VALUE_FORMATTER,
        cellStyle=CELL_STYLE_RULES,
        editable=False
    )

    gb.configure_grid_options(
        groupDisplayType="multipleColumns",
        groupRowAggNodes=GROUP_ROW_AGG,
        groupDefaultExpanded=(0 if not expand_all else -1),
        suppressAggFuncInHeader=True,
        animateRows=False,
        enableRangeSelection=True,
        groupIncludeFooter=True,
        groupIncludeTotalFooter=False,
        suppressColumnMove=True,
        suppressDragLeaveHidesColumns=True,
        suppressScrollOnNewData=True,
        domLayout="autoHeight",
        pinnedBottomRowData=pinned_bottom_rows,
        autoGroupColumnDef={
            "headerName": keyA,
            "minWidth": TEXT_COL_W,
            "cellStyle": {"fontWeight": "600", "backgroundColor": "#ff8c0050", "textAlign":"left"},
            "editable": False,
        },
    )

    return AgGrid(
        df,
        gridOptions=gb.build(),
        update_mode=GridUpdateMode.VALUE_CHANGED,
        data_return_mode=DataReturnMode.AS_INPUT,
        enable_enterprise_modules=True,
        fit_columns_on_grid_load=False,
        allow_unsafe_jscode=True,
        theme="balham",
    )

# ===== App =====
init_storage()
workers_df = load_workers()
obras_df   = load_obras()

st.sidebar.title("Menu")
menu = ["Adicionar Registos","Planeamento Semanal","Gerir Workers","Gerir Obras","Export / Backup"]
page = st.sidebar.selectbox("Ir para", menu, index=0)

# === Adicionar Registos ===
if page == "Adicionar Registos":
    st.header("Adicionar registos em massa")
    if workers_df.empty or obras_df.empty:
        st.warning("Adiciona pelo menos um worker e uma obra antes de registar horas.")
    else:
        c1, c2 = st.columns(2)
        with c1:
            sel_workers = st.multiselect("Selecionar workers", options=workers_df["worker"].tolist())
            sel_obra = st.selectbox("Selecionar obra", options=obras_df["obra"].tolist())
        with c2:
            horas = st.number_input("Horas por dia", min_value=0.0, value=8.0, step=0.25)
            start_dt = st.date_input("Data início", value=date.today())
            end_dt = st.date_input("Data fim", value=date.today())
            dias = st.multiselect(
                "Dias da semana",
                ["Segunda","Terça","Quarta","Quinta","Sexta","Sábado","Domingo"],
                default=["Segunda","Terça","Quarta","Quinta","Sexta"],
            )
        if st.button("Adicionar registos"):
            if not sel_workers:
                st.warning("Seleciona pelo menos um worker.")
            elif end_dt < start_dt:
                st.error("Data fim anterior à data início.")
            elif not dias:
                st.warning("Seleciona pelo menos um dia da semana.")
            else:
                dias_map = {"Segunda":0, "Terça":1, "Quarta":2, "Quinta":3, "Sexta":4, "Sábado":5, "Domingo":6}
                weekday_idx = [dias_map[d] for d in dias]
                rows = []
                for d in pd.date_range(start_dt, end_dt):
                    if d.weekday() in weekday_idx:
                        for w in sel_workers:
                            rows.append({
                                "data_display": d.strftime("%d/%m/%Y"),
                                "worker": w,
                                "obra": sel_obra,
                                "horas": float(horas),
                            })
                if rows:
                    new_df = pd.concat([read_records(), pd.DataFrame(rows)], ignore_index=True)
                    write_records(new_df)
                    st.success("Registos adicionados com sucesso.")
                    st.session_state.records_df = read_records()

# === Planeamento Semanal ===
elif page == "Planeamento Semanal":
    st.header("Planeamento Semanal")

    if "anchor_date" not in st.session_state:
        st.session_state.anchor_date = date.today()
    if "range" not in st.session_state:
        mo, su = monday_of_week(st.session_state.anchor_date), sunday_of_week(st.session_state.anchor_date)
        st.session_state.range = (mo, su)

    # Intervalo personalizado
    dr = st.date_input("Intervalo personalizado", value=st.session_state.range)
    if isinstance(dr, tuple) and len(dr) == 2 and dr != st.session_state.range:
        st.session_state.range = dr
        st.rerun()
    start_cal, end_cal = st.session_state.range

    # Navegação
    l, c, r = st.columns([1,6,1])
    with l:
        if st.button("◀", use_container_width=True):
            st.session_state.anchor_date -= timedelta(days=7)
            st.session_state.range = (monday_of_week(st.session_state.anchor_date), sunday_of_week(st.session_state.anchor_date))
            st.rerun()
    with c:
        st.markdown(f"### {period_label(start_cal, end_cal)}")
    with r:
        if st.button("▶", use_container_width=True):
            st.session_state.anchor_date += timedelta(days=7)
            st.session_state.range = (monday_of_week(st.session_state.anchor_date), sunday_of_week(st.session_state.anchor_date))
            st.rerun()

    # Controles
    c1, c2, c3, c4 = st.columns([2,2,2,2])
    with c1:
        group_by = st.selectbox("Agrupar por", ["Project","User"], index=1)
    with c2:
        view_mode = st.selectbox("Modo", ["Calendário","Lista"], index=0)
    with c3:
        filter_workers = st.multiselect("Workers", options=sorted(read_records()["worker"].dropna().unique().tolist()))
    with c4:
        filter_obras = st.multiselect("Obras", options=sorted(read_records()["obra"].dropna().unique().tolist()))

    rec0 = read_records().copy()
    msk = (rec0["date_parsed"] >= pd.to_datetime(start_cal)) & (rec0["date_parsed"] <= pd.to_datetime(end_cal))
    rec0 = rec0.loc[msk].copy()
    if filter_workers:
        rec0 = rec0[rec0["worker"].isin(filter_workers)]
    if filter_obras:
        rec0 = rec0[rec0["obra"].isin(filter_obras)]

    if view_mode == "Lista":
        st.dataframe(
            rec0[["data_display","worker","obra","horas"]]
              .rename(columns={"data_display":"Data","worker":"Worker","obra":"Obra","horas":"Horas"})
              .sort_values(["Data","Worker","Obra"]),
            use_container_width=True,
        )
    else:
        base_df, col_labels, pinned_bottom, label_to_date = build_calendar_dataframe(
            rec0, start_cal, end_cal, group_by, filter_workers, filter_obras
        )

        # Garantir Totais por linha
        for c in col_labels:
            base_df[c] = pd.to_numeric(base_df[c], errors="coerce").fillna(0)
        base_df["Total"] = base_df[col_labels].sum(axis=1)

        keyA = ("Obra" if group_by == "Project" else "Worker")
        keyB = ("Worker" if group_by == "Project" else "Obra")

        totals_all = base_df.drop(columns=[keyA, keyB, "row_key"]).sum(numeric_only=True)
        pinned_now = [{
            keyA: "Total (All)",
            keyB: "",
            "row_key": "__PINNED__",
            **{lbl: float(totals_all.get(lbl, 0)) for lbl in col_labels + ["Total"]},
        }]

        expand_all = st.checkbox("Expandir tudo", value=False)
        grid = aggrid_calendar(base_df, group_by, col_labels, pinned_now, expand_all, editable=True)

        # Ler dados atuais da grelha
        detail_df = None
        grid_data = None
        try:
            if isinstance(grid, dict):
                grid_data = grid.get("data", None)
            else:
                grid_data = getattr(grid, "data", None)
        except Exception:
            grid_data = None

        if grid_data is not None:
            df_grid = pd.DataFrame(grid_data).copy()
            if "row_key" in df_grid.columns:
                df_grid = df_grid[df_grid["row_key"] != "__PINNED__"].copy()
            if "Worker" in df_grid.columns and "Obra" in df_grid.columns:
                df_grid = df_grid[df_grid["Worker"].notna() & df_grid["Obra"].notna()].copy()
            for c in col_labels:
                df_grid[c] = pd.to_numeric(df_grid[c], errors="coerce").fillna(0)
            df_grid["Total"] = df_grid[col_labels].sum(axis=1)
            detail_df = df_grid

        st.markdown("---")

        if st.button("💾 Guardar alterações", type="primary"):
            if detail_df is None or detail_df.empty:
                st.warning("Não há alterações para guardar.")
            else:
                all_df = read_records().copy()
                if "date_parsed" not in all_df.columns:
                    all_df["date_parsed"] = pd.to_datetime(
                        all_df["data"], dayfirst=True, errors="coerce"
                    )
                week_mask = (
                    (all_df["date_parsed"] >= pd.to_datetime(start_cal)) &
                    (all_df["date_parsed"] <= pd.to_datetime(end_cal))
                )
                outside_week = all_df.loc[~week_mask].copy()

                dates = pd.date_range(start=start_cal, end=end_cal)
                lbl_map = {
                    f"{EN_WEEKDAYS[d.weekday()]}, {EN_MONTH[d.month-1]} {d.day}": d.strftime("%d/%m/%Y")
                    for d in dates
                }

                new_rows = []
                for _, row in detail_df.iterrows():
                    worker = str(row["Worker"])
                    obra   = str(row["Obra"])
                    for lbl in col_labels:
                        ddmmyyyy = lbl_map.get(lbl)
                        if not ddmmyyyy:
                            continue
                        val = float(row.get(lbl, 0) or 0)
                        if val > 0:
                            new_rows.append(
                                {
                                    "data_display": ddmmyyyy,
                                    "worker": worker,
                                    "obra": obra,
                                    "horas": val,
                                }
                            )

                if new_rows:
                    week_df = pd.DataFrame(new_rows)
                    combined = pd.concat([outside_week, week_df], ignore_index=True)
                else:
                    combined = outside_week

                write_records(combined)
                st.toast("✅ Alterações guardadas com sucesso.")
                st.rerun()

# === Gerir Workers ===
elif page == "Gerir Workers":
    st.header("Gerir Workers")
    c1, c2 = st.columns([3,1])
    with c1:
        novo = st.text_input("Nome do trabalhador")
    with c2:
        if st.button("Adicionar"):
            if novo:
                dfw = load_workers()
                dfw = pd.concat([dfw, pd.DataFrame([{ "worker": novo }])], ignore_index=True)
                dfw.drop_duplicates(subset=["worker"], inplace=True)
                save_workers(dfw)
                st.success(f"Worker '{novo}' adicionado.")
                workers_df = load_workers()
    st.subheader("Workers existentes")
    st.dataframe(load_workers())

# === Gerir Obras ===
elif page == "Gerir Obras":
    st.header("Gerir Obras")
    c1, c2 = st.columns([3,1])
    with c1:
        nova = st.text_input("Nome da obra")
    with c2:
        if st.button("Adicionar"):
            if nova:
                dfo = load_obras()
                dfo = pd.concat([dfo, pd.DataFrame([{ "obra": nova }])], ignore_index=True)
                dfo.drop_duplicates(subset=["obra"], inplace=True)
                save_obras(dfo)
                st.success(f"Obra '{nova}' adicionada.")
                obras_df = load_obras()
    st.subheader("Obras existentes")
    st.dataframe(load_obras())

# === Export / Backup ===
else:
    st.header("Export / Backup")
    st.write("Descarrega os dados atuais:")

    workers = load_workers()
    obras   = load_obras()
    records = read_records()

    w_csv = workers.to_csv(index=False).encode("utf-8-sig")
    o_csv = obras.to_csv(index=False).encode("utf-8-sig")
    r_csv = records[["data_display","worker","obra","horas"]].rename(columns={"data_display":"data"}).to_csv(index=False).encode("utf-8-sig")

    st.download_button("Descarregar workers.csv", data=w_csv, file_name="workers.csv", mime="text/csv")
    st.download_button("Descarregar obras.csv",   data=o_csv, file_name="obras.csv",   mime="text/csv")
    st.download_button("Descarregar records.csv", data=r_csv, file_name="records.csv", mime="text/csv")
