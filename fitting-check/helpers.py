import base64
import io
import logging
import math
from collections import Counter
from datetime import date
from pathlib import Path

import matplotlib
import matplotlib.pyplot as plt

matplotlib.use("Agg")

logger = logging.getLogger("viktor")

REQUIRED_FAMILY_NAME = "M_Rectangular to Round Duct Transsition - Angle"
COMPANY_STANDARD_FAMILIES = {
    REQUIRED_FAMILY_NAME,
    "M_Rectangular Duct Tee",
    "M_Rectangular Elbow",
}
WRONG_FAMILY_TOKENS = ("m_round duct transition - angle", "m_round")


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
        family_name = row["family_name"]
        family_name_lower = family_name.lower()
        is_company_standard = family_name in COMPANY_STANDARD_FAMILIES
        is_wrong_family = any(token in family_name_lower for token in WRONG_FAMILY_TOKENS)
        pressure_drop = row["pressure_drop"]
        pressure_missing = math.isnan(pressure_drop)
        pressure_zero = not pressure_missing and pressure_drop == 0
        pressure_issue = pressure_missing or pressure_zero
        is_compliant = is_company_standard and not is_wrong_family and not pressure_issue

        results.append(
            {
                **row,
                "is_company_standard": is_company_standard,
                "is_wrong_family": is_wrong_family,
                "recommended_family": REQUIRED_FAMILY_NAME if is_wrong_family else "OK",
                "pressure_issue": pressure_issue,
                "pressure_issue_reason": (
                    "Pressure drop is NaN"
                    if pressure_missing
                    else "Pressure drop is 0"
                    if pressure_zero
                    else "OK"
                ),
                "family_issue_reason": (
                    "Family should be changed to company standard transition family"
                    if is_wrong_family
                    else "Family is not in company standard"
                    if not is_company_standard
                    else "OK"
                ),
                "pressure_drop_label": format_pressure_drop(pressure_drop),
                "short_revit_id": row["revit_element_id"][-3:] if len(row["revit_element_id"]) >= 3 else row["revit_element_id"],
                "is_compliant": is_compliant,
            }
        )

    requires_change = [row for row in results if not row["is_company_standard"] or row["is_wrong_family"]]
    wrong_family = [row for row in results if row["is_wrong_family"]]
    pressure_issues = [row for row in results if row["pressure_issue"]]
    compliant_rows = [row for row in results if row["is_compliant"]]
    non_compliant_rows = [row for row in results if not row["is_compliant"]]
    family_counter = Counter(row["family_name"] or "Unknown" for row in results)

    summary = {
        "total": len(results),
        "compliant_count": len(compliant_rows),
        "non_compliant_count": len(non_compliant_rows),
        "requires_change_count": len(requires_change),
        "pressure_issue_count": len(pressure_issues),
    }

    logger.info(
        "Fitting analysis: total=%s compliant=%s requires_change=%s pressure_issues=%s",
        summary["total"],
        summary["compliant_count"],
        summary["requires_change_count"],
        summary["pressure_issue_count"],
    )

    return {
        "rows": results,
        "requires_change": requires_change,
        "wrong_family": wrong_family,
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
    colors = ["#8B0000" if row["pressure_issue"] else "#111111" for row in rows]

    fig, ax = plt.subplots(figsize=(11, 4.8))
    fig.patch.set_facecolor("#ffffff")
    ax.set_facecolor("#ffffff")
    ax.bar(labels, values, color=colors, width=0.68)
    ax.set_title("Pressure Drop per Revit ID", fontsize=13, fontweight="bold")
    ax.set_xlabel("Last 3 characters of Revit Element ID", fontsize=10)
    ax.set_ylabel("Pressure Drop (Pa)", fontsize=10)
    ax.grid(axis="y", linestyle="--", linewidth=0.6, color="#d0d0d0")
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()
    return fig_to_base64(fig)


def build_family_pie_base64(family_counter: Counter) -> str:
    labels = list(family_counter.keys())
    values = list(family_counter.values())
    colors = ["#111111", "#5f5f5f", "#9a9a9a", "#c8c8c8", "#8B0000", "#404040"]

    fig, ax = plt.subplots(figsize=(7, 5))
    fig.patch.set_facecolor("#ffffff")
    ax.pie(values, labels=labels, autopct="%1.0f%%", startangle=120, colors=colors[: len(values)], textprops={"fontsize": 8})
    ax.set_title("Family Count Breakdown", fontsize=13, fontweight="bold")
    fig.tight_layout()
    return fig_to_base64(fig)


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
            f"<tr><td>Family changes required</td><td><strong>{summary['requires_change_count']}</strong></td></tr>",
            f"<tr><td>Pressure drop issues</td><td><strong>{summary['pressure_issue_count']}</strong></td></tr>",
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
        row_class = "row-danger" if row["pressure_issue"] else ""
        pressure_rows_html += (
            f"<tr class='{row_class}'>"
            f"<td>{row['revit_element_id']}</td>"
            f"<td>{row['family_name']}</td>"
            f"<td>{row['pressure_drop_label']}</td>"
            f"<td>{row['pressure_issue_reason']}</td>"
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
            f"<td>{row['recommended_family']}</td>"
            f"</tr>"
        )

    html = (Path(__file__).parent / "report_template.html").read_text(encoding="utf-8")
    html = html.replace("__DATE__", date.today().isoformat())
    html = html.replace("__SUMMARY_ROWS__", summary_rows_html)
    html = html.replace("__INPUT_ROWS__", input_rows_html or "<tr><td colspan='5'>No input rows</td></tr>")
    html = html.replace(
        "__PRESSURE_ROWS__",
        pressure_rows_html or "<tr><td colspan='4'>No pressure drop issues found</td></tr>",
    )
    html = html.replace(
        "__WRONG_FAMILY_ROWS__",
        wrong_family_rows_html or "<tr><td colspan='4'>No family changes required</td></tr>",
    )
    html = html.replace("__BAR_CHART_B64__", build_pressure_drop_bar_base64(rows))
    html = html.replace("__PIE_CHART_B64__", build_family_pie_base64(analysis["family_counter"]))
    html = html.replace("__TOTAL_COUNT__", str(summary["total"]))
    html = html.replace("__WRONG_FAMILY_COUNT__", str(len(requires_change)))
    html = html.replace("__PRESSURE_ISSUE_COUNT__", str(len(pressure_issues)))
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
    title_style = ParagraphStyle("title", parent=styles["Heading1"], fontSize=16, leading=20, spaceAfter=6)
    h2_style = ParagraphStyle("h2", parent=styles["Heading2"], fontSize=11, leading=14, spaceBefore=10, spaceAfter=5)
    body_style = ParagraphStyle("body", parent=styles["BodyText"], fontSize=8, leading=11)
    small_style = ParagraphStyle("small", parent=styles["BodyText"], fontSize=7, leading=9)

    story = [
        Paragraph("QA/QC Fittings Check Report", title_style),
        Paragraph("Engineering report for fitting family validation and pressure drop review", body_style),
        Spacer(1, 4 * mm),
    ]

    summary = analysis["summary"]
    summary_table = Table(
        [
            ["Metric", "Value"],
            ["Total fittings", str(summary["total"])],
            ["Compliant fittings", str(summary["compliant_count"])],
            ["Non compliant fittings", str(summary["non_compliant_count"])],
            ["Family changes required", str(summary["requires_change_count"])],
            ["Pressure drop issues", str(summary["pressure_issue_count"])],
        ],
        colWidths=[70 * mm, 35 * mm],
    )
    summary_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.black),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#9a9a9a")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f1f1f1")]),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
            ]
        )
    )
    story.append(Paragraph("1. Summary", h2_style))
    story.append(summary_table)
    story.append(Spacer(1, 4 * mm))

    bar_image = Image(image_from_base64(build_pressure_drop_bar_base64(analysis["rows"])), width=120 * mm, height=52 * mm)
    pie_image = Image(image_from_base64(build_family_pie_base64(analysis["family_counter"])), width=70 * mm, height=52 * mm)
    story.append(Paragraph("2. Plots", h2_style))
    story.append(Table([[bar_image, pie_image]], colWidths=[130 * mm, 75 * mm]))
    story.append(Spacer(1, 4 * mm))

    story.append(Paragraph("3. Input Table", h2_style))
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
        ("BACKGROUND", (0, 0), (-1, 0), colors.black),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#c0c0c0")),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]
    for idx, row in enumerate(analysis["rows"], start=1):
        if not row["is_compliant"]:
            input_style.append(("BACKGROUND", (0, idx), (-1, idx), colors.HexColor("#f4cccc")))
    input_table.setStyle(TableStyle(input_style))
    story.append(input_table)
    story.append(Spacer(1, 4 * mm))

    story.append(Paragraph("4. Pressure Drop Check", h2_style))
    pressure_rows = [["Revit Element ID", "Family", "Pressure Drop", "Status"]]
    for row in analysis["rows"]:
        pressure_rows.append(
            [
                row["revit_element_id"],
                Paragraph(row["family_name"], small_style),
                row["pressure_drop_label"],
                row["pressure_issue_reason"],
            ]
        )
    pressure_table = Table(pressure_rows, colWidths=[38 * mm, 82 * mm, 30 * mm, 55 * mm], repeatRows=1)
    pressure_style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.black),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#c0c0c0")),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
    ]
    for idx, row in enumerate(analysis["rows"], start=1):
        if row["pressure_issue"]:
            pressure_style.append(("BACKGROUND", (0, idx), (-1, idx), colors.HexColor("#f4cccc")))
    pressure_table.setStyle(TableStyle(pressure_style))
    story.append(pressure_table)
    story.append(Spacer(1, 4 * mm))

    story.append(Paragraph("5. Wrong Family Assignments", h2_style))
    family_rows = [["Revit Element ID", "Current Family", "Type", "Family To Assign"]]
    for row in analysis["rows"]:
        family_rows.append(
            [
                row["revit_element_id"],
                Paragraph(row["family_name"], small_style),
                Paragraph(row["type_name"] or "N/A", small_style),
                Paragraph(row["recommended_family"], small_style),
            ]
        )
    family_table = Table(family_rows, colWidths=[38 * mm, 75 * mm, 40 * mm, 75 * mm], repeatRows=1)
    family_style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.black),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#c0c0c0")),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
    ]
    for idx, row in enumerate(analysis["rows"], start=1):
        if row in analysis["requires_change"]:
            family_style.append(("BACKGROUND", (0, idx), (-1, idx), colors.HexColor("#f4cccc")))
    family_table.setStyle(TableStyle(family_style))
    story.append(family_table)

    doc.build(story)
    buffer.seek(0)
    return buffer.read()
