import base64
import io
import logging
import math
from collections import Counter
from datetime import date
from pathlib import Path

import matplotlib
import matplotlib.pyplot as plt
import plotly.graph_objects as go

matplotlib.use("Agg")

logger = logging.getLogger("viktor")

REQUIRED_TYPE_NAME = "45 Degree"


def to_float(value) -> float:
    if value is None or value == "":
        return float("nan")
    try:
        return float(value)
    except (TypeError, ValueError):
        return float("nan")


def format_pressure_drop(value: float) -> str:
    if value is None or math.isnan(value):
        return "NaN"
    return f"{value:.2f} Pa"


def safe_text(value) -> str:
    return str(value or "").strip()


def normalize_rows(rows: list) -> list[dict]:
    normalized = []
    for row in rows:
        family_name = safe_text(getattr(row, "family_name", None) if not isinstance(row, dict) else row.get("family_name"))
        pressure_drop = to_float(getattr(row, "pressure_drop", None) if not isinstance(row, dict) else row.get("pressure_drop"))
        normalized.append(
            {
                "system_name": safe_text(getattr(row, "system_name", None) if not isinstance(row, dict) else row.get("system_name")),
                "revit_element_id": safe_text(getattr(row, "revit_element_id", None) if not isinstance(row, dict) else row.get("revit_element_id")),
                "family_name": family_name,
                "type_name": safe_text(getattr(row, "type_name", None) if not isinstance(row, dict) else row.get("type_name")),
                "pressure_drop": pressure_drop,
            }
        )
    return normalized


def analyze_fittings(rows: list[dict]) -> dict:
    results = []
    for row in rows:
        type_name = row["type_name"]
        type_name_lower = type_name.lower()
        requires_type_change = "30°" in type_name_lower or "30" in type_name_lower
        pressure_issue = False
        is_compliant = not requires_type_change

        results.append(
            {
                **row,
                "requires_type_change": requires_type_change,
                "recommended_type": REQUIRED_TYPE_NAME if requires_type_change else "OK",
                "pressure_issue": pressure_issue,
                "pressure_issue_reason": "OK",
                "type_issue_reason": "Type contains 30 and should be changed to 45 Degree" if requires_type_change else "OK",
                "pressure_drop_label": format_pressure_drop(pressure_drop),
                "short_revit_id": row["revit_element_id"][-3:] if len(row["revit_element_id"]) >= 3 else row["revit_element_id"],
                "is_compliant": is_compliant,
            }
        )

    requires_change = [row for row in results if row["requires_type_change"]]
    pressure_issues = []
    compliant_rows = [row for row in results if row["is_compliant"]]
    non_compliant_rows = [row for row in results if not row["is_compliant"]]
    family_counter = Counter(row["family_name"] or "Unknown" for row in results)

    summary = {
        "total": len(results),
        "compliant_count": len(compliant_rows),
        "non_compliant_count": len(non_compliant_rows),
        "requires_change_count": len(requires_change),
        "pressure_issue_count": 0,
    }

    logger.info(
        "Fitting analysis: total=%s compliant=%s requires_type_change=%s",
        summary["total"],
        summary["compliant_count"],
        summary["requires_change_count"],
    )

    return {
        "rows": results,
        "requires_change": requires_change,
        "pressure_issues": pressure_issues,
        "compliant_rows": compliant_rows,
        "non_compliant_rows": non_compliant_rows,
        "family_counter": family_counter,
        "summary": summary,
    }


def fig_to_base64(fig: plt.Figure) -> str:
    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", bbox_inches="tight", dpi=140)
    buffer.seek(0)
    encoded = base64.b64encode(buffer.read()).decode("utf-8")
    plt.close(fig)
    return encoded


def build_pressure_drop_bar_base64(rows: list[dict]) -> str:
    labels = [row["short_revit_id"] or "-" for row in rows]
    values = [0 if math.isnan(row["pressure_drop"]) else row["pressure_drop"] for row in rows]
    colors = ["#E88D7B" if row["pressure_issue"] else "#1E90FF" for row in rows]

    fig, ax = plt.subplots(figsize=(11, 4.8))
    fig.patch.set_facecolor("#ffffff")
    ax.set_facecolor("#ffffff")
    ax.bar(labels, values, color=colors, width=0.68)
    ax.set_title("Pressure Drop per Revit ID", fontsize=13, fontweight="bold", color="#1E90FF")
    ax.set_xlabel("Last 3 characters of Revit Element ID", fontsize=10, color="#2C3E50")
    ax.set_ylabel("Pressure Drop (Pa)", fontsize=10, color="#2C3E50")
    ax.grid(axis="y", linestyle="--", linewidth=0.6, color="#C9E5FF")
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#A8D0FF")
    ax.spines["bottom"].set_color("#A8D0FF")
    ax.tick_params(colors="#2C3E50")
    fig.tight_layout()
    return fig_to_base64(fig)


def build_family_pie_base64(family_counter: Counter) -> str:
    labels = list(family_counter.keys())
    values = list(family_counter.values())
    colors = ["#1E90FF", "#4DA6FF", "#70B8FF", "#94C9FF", "#E88D7B", "#B8DAFF"]

    fig, ax = plt.subplots(figsize=(7, 5))
    fig.patch.set_facecolor("#ffffff")
    ax.pie(values, labels=labels, autopct="%1.0f%%", startangle=120, colors=colors[: len(values)], textprops={"fontsize": 8, "color": "#2C3E50"})
    ax.set_title("Family Count Breakdown", fontsize=13, fontweight="bold", color="#1E90FF")
    fig.tight_layout()
    return fig_to_base64(fig)


def build_plotly_pressure_drop_bar(rows: list[dict]) -> str:
    labels = [row["short_revit_id"] or "-" for row in rows]
    values = [0 if math.isnan(row["pressure_drop"]) else row["pressure_drop"] for row in rows]
    colors = ["#E88D7B" if row["pressure_issue"] else "#1E90FF" for row in rows]
    hover_texts = [
        f"<b>Revit ID:</b> ...{row['short_revit_id']}<br>"
        f"<b>Full ID:</b> {row['revit_element_id']}<br>"
        f"<b>Pressure Drop:</b> {row['pressure_drop_label']}<br>"
        f"<b>Status:</b> {row['pressure_issue_reason']}"
        for row in rows
    ]

    fig = go.Figure(
        data=[
            go.Bar(
                x=labels,
                y=values,
                marker=dict(color=colors, line=dict(color="#A8D0FF", width=1)),
                hovertemplate="%{customdata}<extra></extra>",
                customdata=hover_texts,
            )
        ]
    )

    fig.update_layout(
        title=dict(
            text="Pressure Drop per Revit ID",
            font=dict(size=16, color="#1E90FF", family="Times New Roman"),
            x=0.5,
            xanchor="center",
        ),
        xaxis=dict(
            title="Last 3 characters of Revit Element ID",
            titlefont=dict(size=12, color="#2C3E50", family="Times New Roman"),
            tickfont=dict(size=10, color="#2C3E50"),
            gridcolor="#C9E5FF",
            showgrid=False,
            linecolor="#A8D0FF",
        ),
        yaxis=dict(
            title="Pressure Drop (Pa)",
            titlefont=dict(size=12, color="#2C3E50", family="Times New Roman"),
            tickfont=dict(size=10, color="#2C3E50"),
            gridcolor="#C9E5FF",
            showgrid=True,
            linecolor="#A8D0FF",
        ),
        plot_bgcolor="white",
        paper_bgcolor="white",
        hovermode="closest",
        margin=dict(l=60, r=20, t=60, b=60),
        height=400,
    )

    return fig.to_html(include_plotlyjs="cdn", div_id="pressure_drop_chart", config={"displayModeBar": True, "displaylogo": False})


def build_plotly_family_pie(family_counter: Counter) -> str:
    labels = list(family_counter.keys())
    values = list(family_counter.values())
    colors = ["#1E90FF", "#4DA6FF", "#70B8FF", "#94C9FF", "#E88D7B", "#B8DAFF"]

    fig = go.Figure(
        data=[
            go.Pie(
                labels=labels,
                values=values,
                marker=dict(colors=colors[: len(values)], line=dict(color="white", width=2)),
                hovertemplate="<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent}<extra></extra>",
                textfont=dict(size=12, color="white", family="Times New Roman"),
                textposition="inside",
            )
        ]
    )

    fig.update_layout(
        title=dict(
            text="Family Count Breakdown",
            font=dict(size=16, color="#1E90FF", family="Times New Roman"),
            x=0.5,
            xanchor="center",
        ),
        showlegend=True,
        legend=dict(
            font=dict(size=10, color="#2C3E50", family="Times New Roman"),
            orientation="h",
            yanchor="top",
            y=-0.15,
            xanchor="center",
            x=0.5,
        ),
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(l=20, r=20, t=60, b=80),
        height=450,
    )

    return fig.to_html(include_plotlyjs=False, div_id="family_pie_chart", config={"displayModeBar": True, "displaylogo": False})


def build_html_report(analysis: dict) -> str:
    summary = analysis["summary"]
    rows = analysis["rows"]
    requires_change = analysis["requires_change"]
    pressure_issues = analysis["pressure_issues"]

    summary_rows_html = "".join(
        [
            f"<tr><td>Total fittings</td><td><strong>{summary['total']}</strong></td></tr>",
            f"<tr><td>Compliant fittings</td><td><strong>{summary['compliant_count']}</strong></td></tr>",
            f"<tr><td>Non compliant fittings</td><td><strong>{summary['non_compliant_count']}</strong></td></tr>",
            f"<tr><td>Type changes required</td><td><strong>{summary['requires_change_count']}</strong></td></tr>",
        ]
    )

    input_rows_html = ""
    for row in rows:
        row_class = "row-danger" if not row["is_compliant"] else ""
        input_rows_html += (
            f"<tr class='{row_class}'>"
            f"<td>{row['system_name']}</td>"
            f"<td>{row['revit_element_id']}</td>"
            f"<td>{row['family_name']}</td>"
            f"<td>{row['type_name'] or 'N/A'}</td>"
            f"<td>{row['pressure_drop_label']}</td>"
            f"</tr>"
        )

    pressure_rows_html = ""
    for row in rows:
        row_class = "row-danger" if row["requires_type_change"] else ""
        pressure_rows_html += (
            f"<tr class='{row_class}'>"
            f"<td>{row['revit_element_id']}</td>"
            f"<td>{row['family_name']}</td>"
            f"<td>{row['type_name'] or 'N/A'}</td>"
            f"<td>{row['type_issue_reason']}</td>"
            f"</tr>"
        )

    wrong_family_rows_html = ""
    for row in rows:
        row_class = "row-danger" if row in requires_change else ""
        wrong_family_rows_html += (
            f"<tr class='{row_class}'>"
            f"<td>{row['revit_element_id']}</td>"
            f"<td>{row['family_name']}</td>"
            f"<td>{row['type_name'] or 'N/A'}</td>"
            f"<td>{row['recommended_type']}</td>"
            f"</tr>"
        )

    html = (Path(__file__).parent / "report_template.html").read_text(encoding="utf-8")
    html = html.replace("__DATE__", date.today().isoformat())
    html = html.replace("__SUMMARY_ROWS__", summary_rows_html)
    html = html.replace("__INPUT_ROWS__", input_rows_html or "<tr><td colspan='5'>No input rows</td></tr>")
    html = html.replace(
        "__PRESSURE_ROWS__",
        pressure_rows_html or "<tr><td colspan='4'>No type changes required</td></tr>",
    )
    html = html.replace(
        "__WRONG_FAMILY_ROWS__",
        wrong_family_rows_html or "<tr><td colspan='4'>No family changes required</td></tr>",
    )
    html = html.replace("__PLOTLY_BAR_CHART__", build_plotly_pressure_drop_bar(rows))
    html = html.replace("__PLOTLY_PIE_CHART__", build_plotly_family_pie(analysis["family_counter"]))
    html = html.replace("__TOTAL_COUNT__", str(summary["total"]))
    html = html.replace("__WRONG_FAMILY_COUNT__", str(len(requires_change)))
    html = html.replace("__PRESSURE_ISSUE_COUNT__", str(len(requires_change)))
    return html


def image_from_base64(encoded: str):
    image_bytes = base64.b64decode(encoded)
    return io.BytesIO(image_bytes)


def build_pdf_report(analysis: dict) -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=12 * mm,
        rightMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("title", parent=styles["Heading1"], fontSize=16, leading=20, spaceAfter=6, textColor=colors.HexColor("#1E90FF"))
    h2_style = ParagraphStyle("h2", parent=styles["Heading2"], fontSize=11, leading=14, spaceBefore=10, spaceAfter=5, textColor=colors.HexColor("#1E90FF"))
    body_style = ParagraphStyle("body", parent=styles["BodyText"], fontSize=8, leading=11, textColor=colors.HexColor("#2C3E50"))
    small_style = ParagraphStyle("small", parent=styles["BodyText"], fontSize=7, leading=9, textColor=colors.HexColor("#2C3E50"))
    intro_style = ParagraphStyle("intro", parent=styles["BodyText"], fontSize=9, leading=12, spaceAfter=8, textColor=colors.HexColor("#2C3E50"))

    today_str = date.today().strftime("%B %d, %Y")

    story = [
        Paragraph("QA/QC Fittings Check Report", title_style),
        Paragraph(f"Engineering Report - Generated on {today_str}", body_style),
        Spacer(1, 3 * mm),
        Paragraph(
            "This report provides a comprehensive analysis of duct fitting elements to ensure the fitting type names "
            "meet the project QA QC rule. The analysis identifies fittings whose type contains 30 or 30 degrees and "
            "therefore should be changed to 45 Degree.",
            intro_style
        ),
        Paragraph(
            "The quality assurance checks performed in this report verify that non compliant type names are flagged "
            "clearly and grouped for direct corrective action in Revit.",
            intro_style
        ),
        Spacer(1, 4 * mm),
    ]

    summary = analysis["summary"]
    summary_table = Table(
        [
            ["Metric", "Value"],
            ["Total fittings", str(summary["total"])],
            ["Compliant fittings", str(summary["compliant_count"])],
            ["Non compliant fittings", str(summary["non_compliant_count"])],
            ["Type changes required", str(summary["requires_change_count"])],
        ],
        colWidths=[70 * mm, 35 * mm],
    )
    summary_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E90FF")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#A8D0FF")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#E6F3FF")]),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
            ]
        )
    )
    story.append(Paragraph("1. Executive Summary", h2_style))
    story.append(
        Paragraph(
            "The following summary provides key metrics for the analyzed fitting elements. This includes total counts, "
            "compliance status, and items requiring corrective action.",
            body_style
        )
    )
    story.append(Spacer(1, 2 * mm))
    story.append(summary_table)
    story.append(Spacer(1, 2 * mm))

    compliance_rate = (summary["compliant_count"] / summary["total"] * 100) if summary["total"] > 0 else 0
    status_text = (
        f"<b>Compliance Rate:</b> {compliance_rate:.1f}% ({summary['compliant_count']} of {summary['total']} fittings). "
    )
    if summary["non_compliant_count"] > 0:
        status_text += (
            f"<b>Action Required:</b> {summary['requires_change_count']} fitting(s) require a type change to {REQUIRED_TYPE_NAME}."
        )
    else:
        status_text += "<b>Status:</b> All fittings meet compliance requirements."

    story.append(Paragraph(status_text, body_style))
    story.append(Spacer(1, 4 * mm))

    bar_image = Image(image_from_base64(build_pressure_drop_bar_base64(analysis["rows"])), width=120 * mm, height=52 * mm)
    pie_image = Image(image_from_base64(build_family_pie_base64(analysis["family_counter"])), width=70 * mm, height=52 * mm)
    story.append(Paragraph("2. Visual Analysis", h2_style))
    story.append(
        Paragraph(
            "<b>Pressure Drop Plot (Left):</b> Bar chart showing pressure drop values for each fitting element. "
            "This stays in the report for engineering context only and does not drive the QA QC rule.",
            body_style
        )
    )
    story.append(Spacer(1, 2 * mm))
    story.append(
        Paragraph(
            "<b>Family Distribution (Right):</b> Pie chart illustrating the distribution of Revit family types across "
            "all analyzed fittings. This provides insight into which family types are most commonly used in the system.",
            body_style
        )
    )
    story.append(Spacer(1, 3 * mm))
    story.append(Table([[bar_image, pie_image]], colWidths=[130 * mm, 75 * mm]))
    story.append(Spacer(1, 4 * mm))

    story.append(Paragraph("3. Detailed Fitting Inventory", h2_style))
    story.append(
        Paragraph(
            "Complete listing of all analyzed duct fitting elements with their associated properties. "
            "Rows highlighted in red indicate non-compliant fittings that require attention. "
            "This table serves as the master record for tracking and corrective action.",
            body_style
        )
    )
    story.append(Spacer(1, 2 * mm))
    input_rows = [["System Name", "Revit Element ID", "Family", "Type", "Pressure Drop"]]
    for row in analysis["rows"]:
        input_rows.append(
            [
                row["system_name"],
                row["revit_element_id"],
                Paragraph(row["family_name"], small_style),
                Paragraph(row["type_name"] or "N/A", small_style),
                row["pressure_drop_label"],
            ]
        )
    input_table = Table(input_rows, colWidths=[42 * mm, 35 * mm, 72 * mm, 48 * mm, 28 * mm], repeatRows=1)
    input_style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E90FF")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#A8D0FF")),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]
    for idx, row in enumerate(analysis["rows"], start=1):
        if not row["is_compliant"]:
            input_style.append(("BACKGROUND", (0, idx), (-1, idx), colors.HexColor("#FFE5E0")))
    input_table.setStyle(TableStyle(input_style))
    story.append(input_table)
    story.append(Spacer(1, 4 * mm))

    story.append(Paragraph("4. Type Angle Validation", h2_style))
    story.append(
        Paragraph(
            "This section validates the fitting type names. "
            "Any fitting whose type contains 30 or 30 degrees is flagged and should be updated to 45 Degree.",
            body_style
        )
    )
    story.append(Spacer(1, 2 * mm))
    story.append(
        Paragraph(
            f"<b>Results:</b> {summary['compliant_count']} of {summary['total']} fittings are compliant. "
            f"{summary['requires_change_count']} fitting(s) require a type change.",
            body_style
        )
    )
    story.append(Spacer(1, 2 * mm))
    pressure_rows = [["Revit Element ID", "Family", "Type", "Status"]]
    for row in analysis["rows"]:
        pressure_rows.append(
            [
                row["revit_element_id"],
                Paragraph(row["family_name"], small_style),
                row["type_name"] or "N/A",
                row["type_issue_reason"],
            ]
        )
    pressure_table = Table(pressure_rows, colWidths=[38 * mm, 82 * mm, 30 * mm, 55 * mm], repeatRows=1)
    pressure_style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E90FF")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#A8D0FF")),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
    ]
    for idx, row in enumerate(analysis["rows"], start=1):
        if row["requires_type_change"]:
            pressure_style.append(("BACKGROUND", (0, idx), (-1, idx), colors.HexColor("#FFE5E0")))
    pressure_table.setStyle(TableStyle(pressure_style))
    story.append(pressure_table)
    story.append(Spacer(1, 4 * mm))

    story.append(Paragraph("5. Type Change Review", h2_style))
    story.append(
        Paragraph(
            "This section identifies fitting elements whose type names contain 30 or 30 degrees. "
            f"Those fittings should be updated to <b>{REQUIRED_TYPE_NAME}</b>.",
            body_style
        )
    )
    story.append(Spacer(1, 2 * mm))
    story.append(
        Paragraph(
            f"<b>Results:</b> {summary['compliant_count']} of {summary['total']} fittings are compliant. "
            f"{summary['requires_change_count']} fitting(s) require a type change to {REQUIRED_TYPE_NAME}.",
            body_style
        )
    )
    story.append(Spacer(1, 2 * mm))
    family_rows = [["Revit Element ID", "Current Family", "Current Type", "Type To Assign"]]
    for row in analysis["rows"]:
        family_rows.append(
            [
                row["revit_element_id"],
                Paragraph(row["family_name"], small_style),
                Paragraph(row["type_name"] or "N/A", small_style),
                Paragraph(row["recommended_type"], small_style),
            ]
        )
    family_table = Table(family_rows, colWidths=[38 * mm, 75 * mm, 40 * mm, 75 * mm], repeatRows=1)
    family_style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E90FF")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#A8D0FF")),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
    ]
    for idx, row in enumerate(analysis["rows"], start=1):
        if row in analysis["requires_change"]:
            family_style.append(("BACKGROUND", (0, idx), (-1, idx), colors.HexColor("#FFE5E0")))
    family_table.setStyle(TableStyle(family_style))
    story.append(family_table)
    story.append(Spacer(1, 6 * mm))

    story.append(Paragraph("6. Recommendations and Next Steps", h2_style))
    story.append(Spacer(1, 2 * mm))

    recommendations = []
    if summary["requires_change_count"] > 0:
        recommendations.append(
            f"• Update {summary['requires_change_count']} fitting element(s) to use the type "
            f"{REQUIRED_TYPE_NAME} as identified in Section 5."
        )
    if summary["compliant_count"] == summary["total"]:
        recommendations.append(
            "• All fittings are compliant with the type angle rule. No corrective action required at this time."
        )
    else:
        recommendations.append(
            "• After making corrections in Revit, re-export the fitting data and regenerate this report to verify compliance."
        )
        recommendations.append(
            "• Maintain consistent use of 45 Degree fitting types where the design intent requires that angle."
        )

    for rec in recommendations:
        story.append(Paragraph(rec, body_style))
        story.append(Spacer(1, 2 * mm))

    story.append(Spacer(1, 4 * mm))
    story.append(
        Paragraph(
            "<b>Note:</b> This report was automatically generated by the QA/QC Fittings Check application. "
            "For questions or assistance, please contact your BIM coordinator or project manager.",
            body_style
        )
    )

    doc.build(story)
    buffer.seek(0)
    return buffer.read()
