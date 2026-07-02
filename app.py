from __future__ import annotations

from datetime import date
from io import BytesIO
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from utils.analise_eventos_incorretos import (
    REQUESTED_RANKINGS,
    build_event_excel_bytes,
    build_event_rankings,
    detect_event_columns,
    event_insights,
    filter_event_dataframe,
    format_ranking_table,
    list_excel_sheets,
    read_excel_sheet,
)


APP_DIR = Path(__file__).resolve().parent
LOGO_PATH = APP_DIR / "assets" / "valenet_logo.png"


st.set_page_config(
    page_title="Analista Field",
    page_icon=str(LOGO_PATH) if LOGO_PATH.exists() else ":bar_chart:",
    layout="wide",
)


st.markdown(
    """
    <style>
    .stApp { background: #071426; color: #f8fafc; }
    .main .block-container { padding-top: 1.4rem; max-width: 1480px; }
    h1, h2, h3, h4, label, .stMarkdown { color: #f8fafc; }
    div[data-testid="stMetric"] {
        background: #0b111a;
        border: 1px solid #1f2937;
        border-radius: 8px;
        padding: 14px 16px;
    }
    div[data-testid="stMetric"] label, div[data-testid="stMetric"] [data-testid="stMetricValue"] {
        color: #f8fafc;
    }
    .event-panel {
        background: #0b111a;
        border: 1px solid #1f2937;
        border-radius: 8px;
        padding: 14px 16px;
        min-height: 190px;
    }
    .event-panel h4 { color: #f8fafc; margin: 0 0 8px 0; }
    .event-panel li, .event-panel p { color: #aab4c3; }
    .event-panel strong { color: #6ee7b7; }
    div[data-testid="stDataFrame"] {
        border: 1px solid #1f2937;
        border-radius: 8px;
    }
    .signature-footer {
        margin-top: 2.5rem;
        padding-top: .8rem;
        border-top: 1px solid rgba(148, 163, 184, .18);
        color: #64748b;
        font-size: .78rem;
        text-align: center;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


EVENT_DARK_LAYOUT = {
    "paper_bgcolor": "#0b111a",
    "plot_bgcolor": "#0b111a",
    "font": {"color": "#f8fafc"},
    "margin": {"l": 45, "r": 45, "t": 70, "b": 45},
    "legend": {"font": {"color": "#f8fafc"}},
}

MONTH_NAMES_PT_BR = {
    1: "janeiro",
    2: "fevereiro",
    3: "março",
    4: "abril",
    5: "maio",
    6: "junho",
    7: "julho",
    8: "agosto",
    9: "setembro",
    10: "outubro",
    11: "novembro",
    12: "dezembro",
}


@st.cache_data(show_spinner=False)
def _event_sheet_names(raw: bytes) -> list[str]:
    return list_excel_sheets(raw)


@st.cache_data(show_spinner=False)
def _event_sheet_dataframe(raw: bytes, sheet_name: str) -> tuple[pd.DataFrame, int, int]:
    sheet_data = read_excel_sheet(raw, sheet_name)
    return sheet_data.dataframe, sheet_data.header_row, sheet_data.raw_rows


def _select_index(options: list[str], selected: str | None) -> int:
    if selected and selected in options:
        return options.index(selected)
    return 0


def _event_filter_options(df: pd.DataFrame, column: str | None, limit: int = 250) -> list[str]:
    if not column or column not in df.columns:
        return []
    values = df[column].astype("string").str.strip()
    values = values[values.notna() & ~values.str.lower().isin(["", "nan", "none"])]
    return [str(value) for value in values.value_counts().head(limit).index.tolist()]


def _bar_colors(size: int) -> list[str]:
    if size <= 1:
        return ["#3b82f6"]
    return px.colors.sample_colorscale("Blues", [index / (size - 1) for index in range(size)])


def _event_donut_chart(ranking: pd.DataFrame, label_col: str, top_n: int) -> go.Figure:
    plot_df = ranking.head(top_n).copy()
    fig = px.pie(
        plot_df,
        names=label_col,
        values="Quantidade",
        hole=0.42,
        color_discrete_sequence=px.colors.qualitative.Set3,
    )
    fig.update_traces(
        textinfo="percent",
        textposition="inside",
        hovertemplate="%{label}<br>%{value} registros<br>%{percent}<extra></extra>",
    )
    fig.update_layout(**EVENT_DARK_LAYOUT)
    fig.update_layout(title=f"Participacao por {label_col}", title_x=0.02)
    return fig


def _event_pareto_chart(ranking: pd.DataFrame, label_col: str, top_n: int) -> go.Figure:
    plot_df = ranking.head(top_n).copy()
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(
        go.Bar(
            x=plot_df[label_col],
            y=plot_df["Quantidade"],
            text=plot_df["Quantidade"],
            textposition="outside",
            marker_color=_bar_colors(len(plot_df)),
            name="Quantidade",
            hovertemplate="%{x}<br>%{y} registros<extra></extra>",
        ),
        secondary_y=False,
    )
    fig.add_trace(
        go.Scatter(
            x=plot_df[label_col],
            y=plot_df["% acumulado"],
            mode="lines+markers",
            line={"color": "#f97316", "width": 3},
            marker={"size": 7},
            name="% acumulado",
            hovertemplate="%{x}<br>%{y:.1f}% acumulado<extra></extra>",
        ),
        secondary_y=True,
    )
    fig.update_layout(**EVENT_DARK_LAYOUT)
    fig.update_layout(title=f"Top {len(plot_df)} categorias - {label_col}", title_x=0.02)
    fig.update_yaxes(title_text="Quantidade", gridcolor="#273244", secondary_y=False)
    fig.update_yaxes(title_text="% acumulado", range=[0, 105], gridcolor="#273244", secondary_y=True)
    fig.update_xaxes(title_text=label_col, tickangle=0)
    return fig


def _render_event_insights(title: str, insights: list[str]) -> None:
    items = "".join(f"<li>{item}</li>" for item in insights)
    st.markdown(
        f"""
        <div class="event-panel">
          <h4>{title}</h4>
          <ul>{items}</ul>
          <p>A pizza mostra participacao por quantidade. A barra mostra ranking/Pareto com percentual acumulado.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _render_footer() -> None:
    st.markdown(
        '<div class="signature-footer">© 2026 Walace.gorino. Todos os direitos reservados.</div>',
        unsafe_allow_html=True,
    )


def _pdf_table_from_dataframe(df: pd.DataFrame, max_rows: int = 12) -> Table:
    visible = df.head(max_rows).copy()
    for column in visible.columns:
        visible[column] = visible[column].astype(str)
    data = [list(visible.columns)] + visible.values.tolist()
    table = Table(data, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b3d75")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 7.5),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#d5dbe5")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f3f6fb")]),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    return table


def _format_filter_summary(filters: dict[str, list[str]], date_range: tuple[pd.Timestamp, pd.Timestamp] | None) -> str:
    parts: list[str] = []
    for key, values in filters.items():
        if values:
            parts.append(f"{key}: {', '.join(values[:6])}{'...' if len(values) > 6 else ''}")
    if date_range:
        start, end = date_range
        parts.append(f"período: {start.strftime('%d/%m/%Y')} a {end.strftime('%d/%m/%Y')}")
    return "; ".join(parts) if parts else "Sem filtros adicionais."


def _build_pdf_report(
    *,
    selected_sheet: str,
    header_row: int,
    raw_rows: int,
    filtered_rows: int,
    available_dimensions: list[tuple[str, str]],
    selected_dimension_label: str,
    selected_ranking: pd.DataFrame,
    top_n: int,
    rankings: dict[str, pd.DataFrame],
    filters: dict[str, list[str]],
    date_range: tuple[pd.Timestamp, pd.Timestamp] | None,
    detected: dict[str, str | None],
) -> bytes:
    output = BytesIO()
    doc = SimpleDocTemplate(
        output,
        pagesize=A4,
        rightMargin=1.2 * cm,
        leftMargin=1.2 * cm,
        topMargin=1.1 * cm,
        bottomMargin=1.1 * cm,
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="CenterTitle", parent=styles["Title"], alignment=TA_CENTER, textColor=colors.HexColor("#08213f")))
    styles.add(ParagraphStyle(name="SmallMuted", parent=styles["BodyText"], fontSize=8, textColor=colors.HexColor("#4b5563"), leading=10))
    styles.add(ParagraphStyle(name="Section", parent=styles["Heading2"], textColor=colors.HexColor("#0b3d75"), spaceBefore=10, spaceAfter=6))

    story: list = []
    if LOGO_PATH.exists():
        story.append(Image(str(LOGO_PATH), width=4.2 * cm, height=1.4 * cm, kind="proportional"))
        story.append(Spacer(1, 0.15 * cm))
    story.append(Paragraph("Analista Field", styles["CenterTitle"]))
    story.append(Paragraph("Relatorio de eventos de campo", styles["SmallMuted"]))
    story.append(Spacer(1, 0.35 * cm))

    summary_data = [
        ["Item", "Valor"],
        ["Aba analisada", selected_sheet],
        ["Linha de cabecalho", str(header_row)],
        ["Linhas carregadas", f"{raw_rows:,}".replace(",", ".")],
        ["Registros filtrados", f"{filtered_rows:,}".replace(",", ".")],
        ["Dimensoes detectadas", str(len(available_dimensions))],
        ["Dimensao explorada", selected_dimension_label],
        ["Filtro aplicado", _format_filter_summary(filters, date_range)],
        ["Coluna de regional", detected.get("regional") or "Nao detectada"],
        ["Coluna de problema", detected.get("problema") or "Nao detectada"],
        ["Coluna de data", detected.get("data") or "Nao detectada"],
    ]
    summary_table = Table(summary_data, colWidths=[5 * cm, 12.2 * cm])
    summary_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b3d75")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#d5dbe5")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f3f6fb")]),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )
    story.append(summary_table)

    story.append(Paragraph("Leitura executiva", styles["Section"]))
    for insight in event_insights(selected_ranking, selected_dimension_label, top_n):
        story.append(Paragraph(f"- {insight}", styles["BodyText"]))

    story.append(Paragraph(f"Top {min(top_n, len(selected_ranking))} - {selected_dimension_label}", styles["Section"]))
    story.append(_pdf_table_from_dataframe(format_ranking_table(selected_ranking), max_rows=top_n))

    story.append(PageBreak())
    story.append(Paragraph("Rankings principais", styles["Section"]))
    for key, label in REQUESTED_RANKINGS:
        table = rankings.get(key, pd.DataFrame())
        if table.empty:
            continue
        story.append(Paragraph(label, styles["Heading3"]))
        story.append(_pdf_table_from_dataframe(format_ranking_table(table), max_rows=12))
        story.append(Spacer(1, 0.25 * cm))

    doc.build(story)
    return output.getvalue()


def _last_day_of_month(year: int, month: int) -> int:
    return int(pd.Timestamp(year=year, month=month, day=1).days_in_month)


def _set_date_selector_defaults(prefix: str, value: date) -> None:
    st.session_state[f"{prefix}_day"] = value.day
    st.session_state[f"{prefix}_month"] = value.month
    st.session_state[f"{prefix}_year"] = value.year


def _date_selector(prefix: str, label: str, years: list[int]) -> date:
    current_year = int(st.session_state.get(f"{prefix}_year", years[0]))
    if current_year not in years:
        current_year = years[0]
        st.session_state[f"{prefix}_year"] = current_year
    current_month = int(st.session_state.get(f"{prefix}_month", 1))
    if current_month not in MONTH_NAMES_PT_BR:
        current_month = 1
        st.session_state[f"{prefix}_month"] = current_month
    max_day = _last_day_of_month(current_year, current_month)
    current_day = min(int(st.session_state.get(f"{prefix}_day", 1)), max_day)
    st.session_state[f"{prefix}_day"] = current_day

    st.caption(label)
    day_col, month_col, year_col = st.columns([1, 2, 1])
    with day_col:
        day = st.selectbox(
            "Dia",
            list(range(1, max_day + 1)),
            index=current_day - 1,
            key=f"{prefix}_day",
        )
    with month_col:
        month = st.selectbox(
            "Mês",
            list(MONTH_NAMES_PT_BR.keys()),
            format_func=lambda value: MONTH_NAMES_PT_BR[value],
            index=current_month - 1,
            key=f"{prefix}_month",
        )
    with year_col:
        year = st.selectbox(
            "Ano",
            years,
            index=years.index(current_year) if current_year in years else 0,
            key=f"{prefix}_year",
        )
    return date(int(year), int(month), min(int(day), _last_day_of_month(int(year), int(month))))


def _date_range_control(df: pd.DataFrame, date_col: str | None) -> tuple[pd.Timestamp, pd.Timestamp] | None:
    if not date_col or date_col not in df.columns:
        return None
    dates = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)
    valid_dates = dates.dropna()
    if valid_dates.empty:
        return None
    start_default = valid_dates.min().date()
    end_default = valid_dates.max().date()

    bounds_key = (date_col, start_default.isoformat(), end_default.isoformat())
    if st.session_state.get("event_date_bounds") != bounds_key:
        st.session_state["event_date_bounds"] = bounds_key
        st.session_state["event_applied_date_range"] = (pd.Timestamp(start_default), pd.Timestamp(end_default))
        _set_date_selector_defaults("event_start", start_default)
        _set_date_selector_defaults("event_end", end_default)

    st.caption("Período")
    years = list(range(start_default.year, end_default.year + 1))
    picker_col, button_col = st.columns([5, 1])
    with picker_col:
        start_col, end_col = st.columns(2)
        with start_col:
            selected_start = _date_selector("event_start", "De", years)
        with end_col:
            selected_end = _date_selector("event_end", "Até", years)
    with button_col:
        st.write("")
        st.write("")
        apply_clicked = st.button("Aplicar período", key="event_apply_period", use_container_width=True)

    if apply_clicked:
        start = max(pd.Timestamp(selected_start), pd.Timestamp(start_default))
        end = min(pd.Timestamp(selected_end), pd.Timestamp(end_default))
        if start > end:
            st.warning("A data inicial não pode ser maior que a data final.")
        else:
            st.session_state["event_applied_date_range"] = (start, end)

    applied = st.session_state.get("event_applied_date_range")
    if not applied:
        return None
    start, end = applied
    st.caption(f"Período aplicado: {start.strftime('%d/%m/%Y')} a {end.strftime('%d/%m/%Y')}")
    return start, end + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1)


def _mapping_controls(df: pd.DataFrame, detected: dict[str, str | None]) -> dict[str, str | None]:
    columns = [str(col) for col in df.columns]
    options = [""] + columns
    with st.expander("Mapeamento automatico de colunas", expanded=False):
        cols = st.columns(4)
        mapping_labels = [
            ("instalador", "Instalador"),
            ("cidade", "Cidade"),
            ("bairro", "Bairro"),
            ("regional", "Regional"),
            ("evento", "Tipo de evento"),
            ("problema", "Problema/observacao"),
            ("data", "Data"),
        ]
        for index, (key, label) in enumerate(mapping_labels):
            with cols[index % 4]:
                detected[key] = st.selectbox(
                    label,
                    options,
                    index=_select_index(options, detected.get(key)),
                    key=f"event_map_{key}",
                ) or None
    return detected


def _filter_controls(
    df: pd.DataFrame,
    detected: dict[str, str | None],
) -> tuple[pd.DataFrame, tuple[pd.Timestamp, pd.Timestamp] | None, dict[str, list[str]]]:
    st.subheader("Filtros")
    selected_filters: dict[str, list[str]] = {}
    cols = st.columns(4)
    filter_labels = [
        ("regional", "Regional"),
        ("evento", "Tipo de evento"),
        ("problema", "Problema"),
        ("instalador", "Instalador"),
        ("cidade", "Cidade"),
        ("bairro", "Bairro"),
    ]
    for index, (key, label) in enumerate(filter_labels):
        values = _event_filter_options(df, detected.get(key))
        if not values:
            continue
        with cols[index % 4]:
            selected_filters[key] = st.multiselect(
                label,
                values,
                default=[],
                key=f"event_filter_{key}",
                help="Sem selecao, o filtro fica aberto.",
            )
    with cols[2]:
        date_range = _date_range_control(df, detected.get("data"))
    return filter_event_dataframe(df, detected, selected_filters, date_range), date_range, selected_filters


def main() -> None:
    header_cols = st.columns([1, 5])
    with header_cols[0]:
        if LOGO_PATH.exists():
            st.image(str(LOGO_PATH), width=150)
    with header_cols[1]:
        st.title("Analista Field")

    uploaded_file = st.file_uploader(
        "Planilha XLSX",
        type=["xlsx", "xlsm"],
        key="upload_eventos_incorretos",
    )
    if uploaded_file is None:
        st.info("Envie a planilha para escolher a aba e iniciar a analise.")
        _render_footer()
        return

    raw = uploaded_file.getvalue()
    try:
        sheets = _event_sheet_names(raw)
    except Exception as exc:  # noqa: BLE001 - friendly app error.
        st.error(f"Nao foi possivel ler as abas da planilha: {exc}")
        _render_footer()
        return
    if not sheets:
        st.warning("A planilha nao possui abas legiveis.")
        _render_footer()
        return

    top_cols = st.columns([2, 2, 2])
    with top_cols[0]:
        selected_sheet = st.selectbox("Aba para analisar", sheets, key="event_sheet")
    try:
        with st.spinner("Lendo a aba selecionada..."):
            df, header_row, raw_rows = _event_sheet_dataframe(raw, selected_sheet)
    except Exception as exc:  # noqa: BLE001 - friendly app error.
        st.error(f"Nao foi possivel ler a aba selecionada: {exc}")
        _render_footer()
        return
    if df.empty:
        st.warning("A aba selecionada nao possui dados tabulares suficientes.")
        _render_footer()
        return

    detected = _mapping_controls(df, detect_event_columns(df))
    columns = [str(col) for col in df.columns]
    temporal_options = [""] + columns
    with top_cols[2]:
        temporal_col = st.selectbox(
            "Coluna temporal",
            temporal_options,
            index=_select_index(temporal_options, detected.get("data")),
            key="event_temporal",
        )
    detected["data"] = temporal_col or None
    with top_cols[1]:
        st.metric("Linhas carregadas", f"{raw_rows:,}".replace(",", "."))

    st.caption(f"Aba: {selected_sheet} | Cabecalho identificado na linha {header_row}")

    filtered_df, date_range, selected_filters = _filter_controls(df, detected)
    if filtered_df.empty:
        st.warning("Nenhuma linha encontrada com os filtros selecionados.")
        _render_footer()
        return

    rankings = build_event_rankings(filtered_df, detected)
    available_dimensions = [
        (key, label)
        for key, label in REQUESTED_RANKINGS
        if not rankings.get(key, pd.DataFrame()).empty
    ]
    if not available_dimensions:
        st.warning("Nao encontrei colunas suficientes para montar os rankings solicitados.")
        _render_footer()
        return

    metric_items = [
        ("Registros filtrados", f"{len(filtered_df):,}".replace(",", ".")),
        ("Dimensoes detectadas", str(len(available_dimensions))),
        ("Regional", detected.get("regional") or "Nao detectada"),
        ("Problema", detected.get("problema") or "Nao detectado"),
    ]
    metric_cols = st.columns(4)
    for col, (label, value) in zip(metric_cols, metric_items):
        col.metric(label, value)

    st.subheader("Explorador visual interativo")
    control_cols = st.columns([2, 2, 2])
    dimension_label_map = {label: key for key, label in available_dimensions}
    with control_cols[0]:
        selected_dimension_label = st.selectbox(
            "Dimensao categorica",
            list(dimension_label_map.keys()),
            key="event_dimension",
        )
    selected_key = dimension_label_map[selected_dimension_label]
    selected_ranking = rankings[selected_key]
    max_top = min(30, max(3, len(selected_ranking)))
    with control_cols[1]:
        top_n = st.slider("Top categorias", min_value=3, max_value=max_top, value=min(10, max_top), key="event_top_n")
    with control_cols[2]:
        st.metric("Categorias na dimensao", f"{len(selected_ranking):,}".replace(",", "."))

    chart_cols = st.columns(2)
    with chart_cols[0]:
        st.plotly_chart(_event_donut_chart(selected_ranking, selected_dimension_label, top_n), use_container_width=True)
    with chart_cols[1]:
        st.plotly_chart(_event_pareto_chart(selected_ranking, selected_dimension_label, top_n), use_container_width=True)

    reading_cols = st.columns(2)
    insights = event_insights(selected_ranking, selected_dimension_label, top_n)
    with reading_cols[0]:
        _render_event_insights(f"Leitura do grafico - Participacao por {selected_dimension_label}", insights)
    with reading_cols[1]:
        _render_event_insights(f"Leitura do grafico - Top categorias - {selected_dimension_label}", insights)

    st.subheader("Rankings principais")
    tabs = st.tabs([label for _key, label in REQUESTED_RANKINGS])
    for tab, (key, label) in zip(tabs, REQUESTED_RANKINGS):
        with tab:
            table = rankings.get(key, pd.DataFrame())
            if table.empty:
                st.info(f"Nao foi encontrada coluna suficiente para o ranking de {label.lower()}.")
                continue
            st.dataframe(format_ranking_table(table), use_container_width=True, hide_index=True)

    st.subheader("Downloads e base")
    download_cols = st.columns([1, 3])
    with download_cols[0]:
        st.download_button(
            "Baixar Excel da analise",
            data=build_event_excel_bytes(rankings, filtered_df),
            file_name="analise_eventos_incorretos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.download_button(
            "Baixar relatorio PDF",
            data=_build_pdf_report(
                selected_sheet=selected_sheet,
                header_row=header_row,
                raw_rows=raw_rows,
                filtered_rows=len(filtered_df),
                available_dimensions=available_dimensions,
                selected_dimension_label=selected_dimension_label,
                selected_ranking=selected_ranking,
                top_n=top_n,
                rankings=rankings,
                filters=selected_filters,
                date_range=date_range,
                detected=detected,
            ),
            file_name="relatorio_analista_field.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
    with download_cols[1]:
        st.caption("A exportacao inclui os rankings solicitados e a base filtrada da aba escolhida.")

    with st.expander("Previa da base filtrada"):
        st.dataframe(filtered_df.head(1000), use_container_width=True, hide_index=True)

    _render_footer()


if __name__ == "__main__":
    main()
