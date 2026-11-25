from __future__ import annotations

import io
import re
import unicodedata
import zipfile
from pathlib import Path
from typing import Any

import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException


DEFAULT_WORKBOOK = Path(__file__).with_name("Anexo_N°_4_Presupuesto_APU.xlsx")

CATEGORY_NAMES = {
    "materiales": "Materiales",
    "maquinaria": "Maquinaria y equipos",
    "mano de obra": "Mano de obra",
    "subcontratos": "Subcontratos y otros",
}


def normalize_text(value: Any) -> str:
    text = "" if value is None else str(value)
    text = unicodedata.normalize("NFKD", text)
    return "".join(char for char in text if not unicodedata.combining(char)).strip().lower()


def cell_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def parse_number(value: Any) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip().replace("\u00a0", "").replace(" ", "")
    if not text:
        return None
    text = text.replace("$", "").replace("CLP", "").replace("clp", "")
    text = re.sub(r"[^\d,.\-]", "", text)
    if not text or text in {"-", ".", ","}:
        return None

    if "," in text and "." in text:
        decimal_separator = "," if text.rfind(",") > text.rfind(".") else "."
        grouping_separator = "." if decimal_separator == "," else ","
        text = text.replace(grouping_separator, "")
        text = text.replace(decimal_separator, ".")
    elif "," in text:
        pieces = text.split(",")
        if len(pieces) > 2 or (len(pieces[-1]) == 3 and pieces[0] != "0"):
            text = "".join(pieces)
        else:
            text = "".join(pieces[:-1]) + "." + pieces[-1]
    elif "." in text:
        pieces = text.split(".")
        if len(pieces) > 2 or (
            len(pieces[-1]) == 3 and len(pieces[0]) >= 1 and pieces[0] != "0"
        ):
            text = "".join(pieces)

    try:
        return float(text)
    except ValueError:
        return None


def find_budget_header(workbook: Any) -> tuple[Any, int]:
    for worksheet in workbook.worksheets:
        for row_number, row in enumerate(
            worksheet.iter_rows(min_row=1, max_row=min(worksheet.max_row, 40), values_only=True),
            start=1,
        ):
            values = [normalize_text(value) for value in row[:14]]
            if "item" in values and any("designacion" in value or "descripcion" in value for value in values):
                if any("cant" in value for value in values):
                    return worksheet, row_number
    raise ValueError("No se encontró la tabla del presupuesto (encabezados ITEM, DESIGNACIÓN y CANT).")


def column_for(values: list[str], predicates: tuple[str, ...], fallback: int) -> int:
    for index, value in enumerate(values, start=1):
        if any(predicate in value for predicate in predicates):
            return index
    return fallback


def read_budget(workbook: Any) -> pd.DataFrame:
    worksheet, header_row = find_budget_header(workbook)
    rows = worksheet.iter_rows(values_only=True)
    headers = [normalize_text(value) for value in next(rows)]
    for _ in range(header_row - 1):
        headers = [normalize_text(value) for value in next(rows)]
    columns = {
        "item": column_for(headers, ("item",), 2),
        "descripcion": column_for(headers, ("designacion", "descripcion"), 3),
        "unidad": column_for(headers, ("unidad",), 4),
        "cantidad": column_for(headers, ("cant",), 5),
        "precio_unitario": column_for(headers, ("valor", "precio unitario", "costo unitario"), 6),
        "total": column_for(headers, ("total",), 7),
    }

    records = []
    for row in rows:
        item = cell_text(row[columns["item"] - 1] if len(row) >= columns["item"] else None)
        description = cell_text(
            row[columns["descripcion"] - 1] if len(row) >= columns["descripcion"] else None
        )
        unit = cell_text(row[columns["unidad"] - 1] if len(row) >= columns["unidad"] else None)
        quantity = parse_number(
            row[columns["cantidad"] - 1] if len(row) >= columns["cantidad"] else None
        )
        if not item or not description or not unit or quantity is None:
            continue

        unit_price = parse_number(
            row[columns["precio_unitario"] - 1]
            if len(row) >= columns["precio_unitario"]
            else None
        )
        total = parse_number(
            row[columns["total"] - 1] if len(row) >= columns["total"] else None
        )
        if total is None and unit_price is not None:
            total = quantity * unit_price
        records.append(
            {
                "partida_codigo": item,
                "partida": description,
                "unidad_partida": unit,
                "cantidad_partida": quantity,
                "precio_unitario": unit_price,
                "total": total,
            }
        )

    if not records:
        raise ValueError("La hoja del presupuesto no contiene partidas con unidad y cantidad.")
    return pd.DataFrame(records)


def category_from_heading(value: Any) -> str | None:
    text = normalize_text(value)
    if "material" in text:
        return CATEGORY_NAMES["materiales"]
    if "maquina" in text or "equipo" in text:
        return CATEGORY_NAMES["maquinaria"]
    if "mano de obra" in text:
        return CATEGORY_NAMES["mano de obra"]
    if "subcontrato" in text or "otros" in text:
        return CATEGORY_NAMES["subcontratos"]
    return None


def canonical_unit(value: Any) -> str:
    unit = normalize_text(value).replace(".", "").strip()
    if unit in {"dia", "dias", "d"}:
        return "días"
    if unit in {"hr", "hrs", "hora", "horas", "h"}:
        return "horas"
    if unit in {"jornada", "jornadas"}:
        return "jornadas"
    return cell_text(value)


def read_apu_sheet(
    worksheet: Any, budget_quantities: dict[str, float]
) -> list[dict[str, Any]]:
    header_row = None
    resources = []
    category = None
    item = ""
    activity = ""
    quantity = 0.0
    for row_number, row in enumerate(worksheet.iter_rows(values_only=True), start=1):
        code_cell = row[1] if len(row) > 1 else None
        description_cell = row[2] if len(row) > 2 else None
        if header_row is None:
            if row_number <= 20 and normalize_text(code_cell) == "item" and normalize_text(
                description_cell
            ) == "actividad":
                header_row = row_number
            continue
        if row_number == header_row + 1:
            item = cell_text(code_cell)
            if item not in budget_quantities:
                return []
            activity = cell_text(description_cell)
            quantity = budget_quantities[item]
            continue
        if row_number <= header_row + 1:
            continue

        heading_category = category_from_heading(code_cell)
        if heading_category:
            category = heading_category
            continue
        if category is None:
            continue

        code = cell_text(code_cell)
        description = cell_text(description_cell)
        yield_per_unit = parse_number(row[3] if len(row) > 3 else None)
        unit = cell_text(row[4] if len(row) > 4 else None)
        if not code or not description or yield_per_unit is None or not unit:
            continue
        if category == CATEGORY_NAMES["materiales"] and not code.upper().startswith("MAT-"):
            continue
        if category == CATEGORY_NAMES["maquinaria"] and not code.upper().startswith("MAQ-"):
            continue
        if category == CATEGORY_NAMES["mano de obra"] and not code.upper().startswith("MO-"):
            continue

        normalized_unit = canonical_unit(unit)
        is_time = category in {
            CATEGORY_NAMES["mano de obra"],
            CATEGORY_NAMES["maquinaria"],
        } and normalized_unit in {"días", "horas", "jornadas"}
        resources.append(
            {
                "partida_codigo": item,
                "partida": activity,
                "cantidad_partida": quantity,
                "tipo_recurso": category,
                "codigo_recurso": code,
                "recurso": description,
                "rendimiento_unitario": yield_per_unit,
                "unidad_recurso": unit,
                "unidad_normalizada": normalized_unit,
                "cantidad_total": yield_per_unit * quantity,
                "es_tiempo": is_time,
                "costo_unitario_recurso": parse_number(row[6] if len(row) > 6 else None),
                "costo_total_recurso": parse_number(row[7] if len(row) > 7 else None),
            }
        )
    return resources


@st.cache_data(show_spinner=False)
def load_project(file_bytes: bytes) -> tuple[pd.DataFrame, pd.DataFrame]:
    workbook = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    budget = read_budget(workbook)
    budget_quantities = dict(zip(budget["partida_codigo"], budget["cantidad_partida"]))
    resource_records = []
    for worksheet in workbook.worksheets:
        resource_records.extend(read_apu_sheet(worksheet, budget_quantities))
    workbook.close()
    resources = pd.DataFrame(resource_records)
    return budget, resources


def fmt_money(value: float) -> str:
    return "$ " + f"{value:,.0f}".replace(",", ".")


def fmt_quantity(value: float) -> str:
    return f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def time_summary(resources: pd.DataFrame, category: str) -> str:
    if resources.empty:
        return "0 días · 0 horas"
    subset = resources[
        (resources["tipo_recurso"] == category)
        & resources["es_tiempo"]
    ]
    if subset.empty:
        return "0 días · 0 horas"
    totals = subset.groupby("unidad_normalizada")["cantidad_total"].sum()
    parts = []
    for unit in ("días", "horas", "jornadas"):
        if unit in totals.index:
            parts.append(f"{fmt_quantity(float(totals[unit]))} {unit}")
    return " · ".join(parts) if parts else "0 días · 0 horas"


def excel_download(budget: pd.DataFrame, resources: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        budget.to_excel(writer, sheet_name="Presupuesto", index=False)
        resources.to_excel(writer, sheet_name="Recursos APU", index=False)
    return output.getvalue()


def render_kpi(label: str, value: str, detail: str, accent: str) -> None:
    st.markdown(
        f"""
        <div class="kpi-card" style="--accent:{accent}">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-detail">{detail}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def main() -> None:
    st.set_page_config(page_title="Control de Presupuesto y APU", page_icon="🏗️", layout="wide")
    st.markdown(
        """
        <style>
        .stApp { background: #f4f6f3; }
        .block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 1500px; }
        .hero { padding: 1.5rem 1.8rem; border-radius: 18px; color: #f8faf8;
                background: linear-gradient(115deg, #173b32, #28634e); margin-bottom: 1.2rem; }
        .hero h1 { margin: 0 0 .35rem; font-size: 2rem; }
        .hero p { margin: 0; color: #d3e4da; }
        .kpi-card { background: #fff; border: 1px solid #e3e9e3; border-top: 4px solid var(--accent);
                    border-radius: 14px; padding: 1rem 1.1rem; min-height: 120px;
                    box-shadow: 0 5px 16px rgba(27, 55, 42, .05); }
        .kpi-label { color: #68756c; font-size: .82rem; font-weight: 650; text-transform: uppercase; }
        .kpi-value { color: #183d32; font-size: 1.42rem; font-weight: 750; margin: .45rem 0 .2rem; }
        .kpi-detail { color: #7d8880; font-size: .78rem; }
        h2, h3 { color: #183d32; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="hero"><h1>Control de presupuesto y APU</h1>'
        '<p>Costos, recursos y ocupación estimada · Centro Comunitario Pedro del Río Zañartu</p></div>',
        unsafe_allow_html=True,
    )

    uploaded_file = st.sidebar.file_uploader("Cargar otro archivo Excel", type=["xlsx"])
    if uploaded_file is not None:
        file_bytes = uploaded_file.getvalue()
        workbook_name = uploaded_file.name
    elif DEFAULT_WORKBOOK.exists():
        file_bytes = DEFAULT_WORKBOOK.read_bytes()
        workbook_name = DEFAULT_WORKBOOK.name
    else:
        st.info("Sube el archivo Anexo_N°_4_Presupuesto_APU.xlsx para comenzar.")
        st.stop()

    try:
        budget, resources = load_project(file_bytes)
    except (InvalidFileException, zipfile.BadZipFile, OSError, ValueError) as exc:
        st.error(f"No fue posible procesar el archivo «{workbook_name}»: {exc}")
        st.stop()

    st.sidebar.caption(f"Fuente: {workbook_name}")
    item_options = budget.apply(
        lambda row: f"{row['partida_codigo']} · {row['partida']}", axis=1
    ).tolist()
    label_to_code = dict(zip(item_options, budget["partida_codigo"]))
    selected_items = st.sidebar.multiselect("Partida", item_options, help="Vacío = todas las partidas")
    resource_types = sorted(resources["tipo_recurso"].unique().tolist()) if not resources.empty else []
    selected_types = st.sidebar.multiselect(
        "Tipo de recurso", resource_types, default=resource_types, help="Vacío = sin recursos"
    )
    selected_codes = [label_to_code[label] for label in selected_items]

    filtered_budget = budget if not selected_codes else budget[budget["partida_codigo"].isin(selected_codes)]
    filtered_resources = resources
    if not resources.empty:
        if selected_codes:
            filtered_resources = filtered_resources[
                filtered_resources["partida_codigo"].isin(selected_codes)
            ]
        filtered_resources = filtered_resources[
            filtered_resources["tipo_recurso"].isin(selected_types)
        ]

    selected_total = float(filtered_budget["total"].fillna(0).sum())
    project_total = float(budget["total"].fillna(0).sum())
    st.caption(f"Archivo analizado: **{workbook_name}** · {len(budget):,} partidas con cantidad")

    kpi_columns = st.columns(4)
    with kpi_columns[0]:
        render_kpi("Costo del presupuesto", fmt_money(selected_total), "Selección actual", "#3b7d5b")
    with kpi_columns[1]:
        render_kpi("Partidas", f"{len(filtered_budget):,}", f"de {len(budget):,} en el archivo", "#d18b36")
    with kpi_columns[2]:
        render_kpi(
            "Mano de obra · ocupación",
            time_summary(filtered_resources, CATEGORY_NAMES["mano de obra"]),
            "Días, horas o jornadas según APU",
            "#477ea8",
        )
    with kpi_columns[3]:
        render_kpi(
            "Maquinaria · ocupación",
            time_summary(filtered_resources, CATEGORY_NAMES["maquinaria"]),
            "Días u horas según APU",
            "#8b68a6",
        )

    st.markdown("### Costos e incidencia por partida")
    if filtered_budget.empty:
        st.info("No hay partidas para los filtros seleccionados.")
    else:
        cost_chart = filtered_budget.copy()
        cost_chart["incidencia"] = cost_chart["total"].fillna(0) / project_total * 100 if project_total else 0
        cost_chart["etiqueta"] = cost_chart["partida_codigo"] + " · " + cost_chart["partida"]
        cost_chart = cost_chart.sort_values("total", ascending=True)
        figure = px.bar(
            cost_chart,
            x="total",
            y="etiqueta",
            orientation="h",
            text=cost_chart["incidencia"].map(lambda value: f"{value:.2f}%"),
            hover_data={
                "partida_codigo": True,
                "partida": True,
                "cantidad_partida": ":,.2f",
                "unidad_partida": True,
                "precio_unitario": ":,.0f",
                "total": ":,.0f",
                "incidencia": ":.2f",
                "etiqueta": False,
            },
            labels={"total": "Costo total ($)", "etiqueta": "Partida"},
            color_discrete_sequence=["#367456"],
        )
        figure.update_layout(
            height=max(380, min(780, 26 * len(cost_chart))),
            yaxis={"categoryorder": "array", "categoryarray": cost_chart["etiqueta"].tolist()},
            margin={"l": 10, "r": 15, "t": 15, "b": 10},
        )
        figure.update_traces(textposition="outside", cliponaxis=False)
        st.plotly_chart(figure, use_container_width=True)

    st.markdown("### Ocupación de recursos por partida")
    temporal_resources = filtered_resources[
        filtered_resources["es_tiempo"]
    ] if not filtered_resources.empty else pd.DataFrame()
    available_units = sorted(temporal_resources["unidad_normalizada"].unique().tolist()) if not temporal_resources.empty else []
    if not available_units:
        st.info("No se encontraron recursos de mano de obra o maquinaria con unidades temporales para esta selección.")
    else:
        chart_unit = st.selectbox("Unidad de tiempo", available_units)
        time_data = temporal_resources[
            temporal_resources["unidad_normalizada"] == chart_unit
        ].copy()
        time_data["etiqueta_partida"] = time_data["partida_codigo"] + " · " + time_data["partida"]
        part_chart, resource_chart = st.columns(2)
        by_item = (
            time_data.groupby(["etiqueta_partida", "recurso", "tipo_recurso"], as_index=False)[
                "cantidad_total"
            ]
            .sum()
        )
        with part_chart:
            st.markdown("#### Por partida y recurso")
            figure = px.bar(
                by_item,
                x="etiqueta_partida",
                y="cantidad_total",
                color="recurso",
                hover_data=["tipo_recurso"],
                labels={
                    "etiqueta_partida": "Partida",
                    "cantidad_total": f"Ocupación ({chart_unit})",
                    "recurso": "Recurso",
                },
            )
            figure.update_layout(
                barmode="stack",
                height=460,
                xaxis_tickangle=-55,
                margin={"l": 10, "r": 10, "t": 20, "b": 100},
                legend_title_text="Recurso",
            )
            st.plotly_chart(figure, use_container_width=True)

        with resource_chart:
            st.markdown("#### Total por recurso")
            by_resource = (
                time_data.groupby(["recurso", "tipo_recurso"], as_index=False)["cantidad_total"]
                .sum()
                .nlargest(25, "cantidad_total")
                .sort_values("cantidad_total", ascending=True)
            )
            figure = px.bar(
                by_resource,
                x="cantidad_total",
                y="recurso",
                orientation="h",
                color="tipo_recurso",
                labels={
                    "cantidad_total": f"Ocupación ({chart_unit})",
                    "recurso": "Recurso",
                    "tipo_recurso": "Tipo",
                },
            )
            figure.update_layout(height=max(420, 24 * len(by_resource)), margin={"l": 10, "r": 10, "t": 20, "b": 10})
            st.plotly_chart(figure, use_container_width=True)

    st.markdown("### Tablas de detalle")
    budget_view = filtered_budget.rename(
        columns={
            "partida_codigo": "Ítem",
            "partida": "Descripción",
            "unidad_partida": "Unidad",
            "cantidad_partida": "Cantidad",
            "precio_unitario": "Precio unitario ($)",
            "total": "Total ($)",
        }
    )
    resources_view = filtered_resources.rename(
        columns={
            "partida_codigo": "Ítem",
            "partida": "Partida",
            "tipo_recurso": "Tipo de recurso",
            "codigo_recurso": "Código",
            "recurso": "Recurso",
            "rendimiento_unitario": "Rendimiento unitario",
            "unidad_recurso": "Unidad",
            "cantidad_total": "Cantidad total",
            "costo_unitario_recurso": "Costo unitario ($)",
            "costo_total_recurso": "Costo total ($)",
        }
    )
    budget_columns = [
        "Ítem", "Descripción", "Unidad", "Cantidad", "Precio unitario ($)", "Total ($)"
    ]
    resource_columns = [
        "Ítem",
        "Partida",
        "Tipo de recurso",
        "Código",
        "Recurso",
        "Rendimiento unitario",
        "Unidad",
        "Cantidad total",
        "Costo unitario ($)",
        "Costo total ($)",
    ]
    tab_budget, tab_resources = st.tabs(["Presupuesto", "Recursos APU"])
    with tab_budget:
        st.dataframe(budget_view[budget_columns], use_container_width=True, hide_index=True)
    with tab_resources:
        if resources_view.empty:
            st.info("No hay recursos APU para los filtros seleccionados.")
        else:
            st.dataframe(resources_view[resource_columns], use_container_width=True, hide_index=True)

    csv_bytes = budget_view[budget_columns].to_csv(index=False).encode("utf-8-sig")
    excel_bytes = excel_download(
        filtered_budget,
        filtered_resources,
    )
    download_columns = st.columns(2)
    with download_columns[0]:
        st.download_button(
            "Descargar tablas en Excel",
            data=excel_bytes,
            file_name="resumen_presupuesto_apu.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with download_columns[1]:
        st.download_button(
            "Descargar presupuesto en CSV",
            data=csv_bytes,
            file_name="presupuesto.csv",
            mime="text/csv",
        )

if __name__ == "__main__":
    main()
