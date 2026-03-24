import logging
import math
import shutil
import subprocess
from datetime import date
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory

logger = logging.getLogger("viktor")

AIR_DENSITY = 1.2
FRICTION_FACTOR_RIGID = 0.02
FRICTION_FACTOR_FLEX = 0.035
DEFAULT_DUCT_LENGTH = 1.0
FLEX_DUCT_SIZES = [75, 100, 125, 150, 200, 225, 250, 300, 350]
ROUND_RIGID_DUCT_SIZES = [100, 125, 150, 160, 200, 250, 315, 400, 500]
RECTANGULAR_DUCT_SIZES = [200, 225, 250, 275, 300, 325, 350, 375, 400, 425, 475, 500, 600, 650]
# For sizes above 650, add increments of 50: 700, 750, 800, etc.
RECTANGULAR_DUCT_SIZES.extend(range(700, 1500, 50))
VELOCITY_INCONSISTENCY_THRESHOLD = 0.5
AIR_VISCOSITY = 1.81e-5

REFERENCE_VELOCITY_LIMITS: list[dict] = [
    {"duct_type": "Flex Duct", "system_type": "Supply", "v_max": 4.0, "nc_limit": 35},
    {"duct_type": "Flex Duct", "system_type": "Return", "v_max": 3.0, "nc_limit": 35},
    {"duct_type": "Flex Duct", "system_type": "Exhaust", "v_max": 3.5, "nc_limit": 40},
    {"duct_type": "Rigid Duct (Round)", "system_type": "Supply", "v_max": 6.0, "nc_limit": 35},
    {"duct_type": "Rigid Duct (Round)", "system_type": "Return", "v_max": 5.0, "nc_limit": 35},
    {"duct_type": "Rigid Duct (Round)", "system_type": "Exhaust", "v_max": 5.5, "nc_limit": 40},
    {"duct_type": "Rigid Duct (Rectangular)", "system_type": "Supply", "v_max": 6.0, "nc_limit": 35},
    {"duct_type": "Rigid Duct (Rectangular)", "system_type": "Return", "v_max": 4.0, "nc_limit": 35},
    {"duct_type": "Rigid Duct (Rectangular)", "system_type": "Exhaust", "v_max": 5.0, "nc_limit": 40},
]

VELOCITY_LOOKUP: dict[tuple[str, str], tuple[float, int]] = {
    (row["duct_type"].lower(), row["system_type"].lower()): (row["v_max"], row["nc_limit"])
    for row in REFERENCE_VELOCITY_LIMITS
}


def is_flex(duct_type: str) -> bool:
    return "flex" in duct_type.lower()


def is_rectangular(duct_type: str) -> bool:
    return "rectangular" in duct_type.lower()


def friction_factor(duct_type: str) -> float:
    return FRICTION_FACTOR_FLEX if is_flex(duct_type) else FRICTION_FACTOR_RIGID


def cross_section_area(duct_type: str, diameter_mm: float, height_mm: float) -> float:
    if is_rectangular(duct_type):
        return (diameter_mm / 1000.0) * (height_mm / 1000.0)
    return math.pi * (diameter_mm / 2000.0) ** 2


def calc_velocity(q_ls: float, area_m2: float) -> float:
    if area_m2 <= 0:
        return 0.0
    return (q_ls / 1000.0) / area_m2


def pressure_drop(duct_type: str, diameter_mm: float, v_measured: float) -> float:
    if diameter_mm <= 0:
        return 0.0
    lam = friction_factor(duct_type)
    d_m = diameter_mm / 1000.0
    return lam * (DEFAULT_DUCT_LENGTH / d_m) * (AIR_DENSITY * v_measured ** 2 / 2.0)


def next_standard_size(d_min_mm: float, duct_type: str) -> int | None:
    """Find the next standard size for round or flex ducts."""
    if is_flex(duct_type):
        sizes = FLEX_DUCT_SIZES
    else:
        sizes = ROUND_RIGID_DUCT_SIZES

    for size in sizes:
        if size >= d_min_mm:
            return size
    return None


def d_min(q_ls: float, v_max: float) -> float:
    if v_max <= 0:
        return 0.0
    return math.sqrt(4.0 * (q_ls / 1000.0) / (math.pi * v_max)) * 1000.0


def find_optimal_rectangular_size(q_ls: float, v_max: float) -> tuple[int | None, int | None]:
    """Find the optimal width x height for rectangular duct from standard sizes."""
    if v_max <= 0:
        return None, None

    # Calculate minimum required area
    min_area_m2 = (q_ls / 1000.0) / v_max

    # Try to find the smallest rectangular duct that meets the requirement
    best_width = None
    best_height = None
    best_area = float('inf')

    for width in RECTANGULAR_DUCT_SIZES:
        for height in RECTANGULAR_DUCT_SIZES:
            # Calculate area in m²
            area_m2 = (width / 1000.0) * (height / 1000.0)

            # Check if this meets the minimum area requirement
            if area_m2 >= min_area_m2:
                # Prefer sizes closer to the minimum (more optimized)
                # and prefer aspect ratios closer to 1:1 (more balanced)
                aspect_ratio = max(width, height) / min(width, height)
                score = area_m2 + (aspect_ratio - 1.0) * 0.01  # Slight preference for balanced ratios

                if best_width is None or score < best_area:
                    best_width = width
                    best_height = height
                    best_area = score

    return best_width, best_height


def format_size_label(duct_type: str, primary_mm: float | int, height_mm: float | int = 0) -> str:
    primary = int(primary_mm or 0)
    height = int(height_mm or 0)
    if primary <= 0:
        return "—"
    if is_rectangular(duct_type):
        return f"{primary} x {height} mm" if height > 0 else f"{primary} mm"
    return f"Ø{primary} mm"


def process_ducts(duct_rows: list, system_type: str = "Supply") -> list[dict]:
    results = []

    for row in duct_rows:
        system_name = (row.get("system_name") or "").strip()
        revit_id = (row.get("revit_id") or "").strip()
        duct_type = (row.get("duct_type") or "").strip()
        q_ls = row.get("airflow_q") or 0.0
        diam_mm = row.get("duct_diameter") or 0.0
        height_mm = row.get("duct_height") or 0.0
        v_measured = row.get("measured_velocity") or 0.0

        # Force height to 0 for flex and round ducts (they are circular)
        if is_flex(duct_type) or ("round" in duct_type.lower() and not is_rectangular(duct_type)):
            height_mm = 0.0

        lookup_key = (duct_type.lower(), system_type.lower())
        match = VELOCITY_LOOKUP.get(lookup_key)
        v_max = match[0] if match else None
        system_type_found = system_type if match else "—"

        if v_max is None:
            results.append({
                "system_name": system_name,
                "revit_id": revit_id,
                "duct_type": duct_type,
                "system_type": "N/A",
                "airflow_q_ls": q_ls,
                "v_measured": v_measured,
                "v_max": "—",
                "v_calc": "—",
                "delta_p": "—",
                "delta_v": "—",
                "status": "NO REF",
                "optimization": "—",
                "proposed_size": None,
                "proposed_diameter_mm": None,
                "proposed_width_mm": None,
                "proposed_height_mm": None,
                "v_new": None,
                "current_size_label": format_size_label(duct_type, diam_mm, height_mm),
                "proposed_size_label": "—",
            })
            continue

        area = cross_section_area(duct_type, diam_mm, height_mm)
        v_calc = calc_velocity(q_ls, area)
        delta_p = pressure_drop(duct_type, diam_mm, v_measured)
        delta_v = abs(v_measured - v_calc)

        if delta_v > VELOCITY_INCONSISTENCY_THRESHOLD:
            status = "INCONSISTENT"
        elif v_measured > v_max:
            status = "OVERSIZE"
        else:
            status = "COMPLIANT"

        # Calculate optimization based on duct type
        v_new = None
        optimization = "—"
        proposed_size = None
        proposed_diameter_mm = None
        proposed_width_mm = None
        proposed_height_mm = None

        if is_rectangular(duct_type):
            # Find optimal rectangular size from standard dimensions
            proposed_width_mm, proposed_height_mm = find_optimal_rectangular_size(q_ls, v_max)
            if proposed_width_mm and proposed_height_mm:
                new_area = (proposed_width_mm / 1000.0) * (proposed_height_mm / 1000.0)
                v_new = calc_velocity(q_ls, new_area)
                optimization = f"→ {proposed_width_mm} x {proposed_height_mm} mm"
                proposed_size = proposed_width_mm  # Store width as main size for reference
        else:
            # Round and flex ducts
            d_min_mm = d_min(q_ls, v_max)
            proposed_size = next_standard_size(d_min_mm, duct_type)
            if proposed_size:
                new_area = cross_section_area(duct_type, proposed_size, 0)
                v_new = calc_velocity(q_ls, new_area)
                proposed_diameter_mm = proposed_size
                optimization = f"→ Ø{proposed_size} mm"

        results.append({
            "system_name": system_name,
            "revit_id": revit_id,
            "duct_type": duct_type,
            "system_type": system_type_found,
            "airflow_q_ls": q_ls,
            "v_measured": v_measured,
            "v_max": v_max,
            "v_calc": round(v_calc, 2),
            "delta_p": round(delta_p, 2),
            "delta_v": round(delta_v, 2),
            "status": status,
            "optimization": optimization,
            "proposed_size": proposed_size,
            "proposed_diameter_mm": proposed_diameter_mm,
            "proposed_width_mm": proposed_width_mm,
            "proposed_height_mm": proposed_height_mm,
            "v_new": round(v_new, 2) if v_new is not None else None,
            "diam_mm": diam_mm,
            "height_mm": height_mm,
            "current_size_label": format_size_label(duct_type, diam_mm, height_mm),
            "proposed_size_label": format_size_label(duct_type, proposed_size, proposed_size),
        })

    logger.info(f"Processed {len(results)} duct rows")
    return results


def build_html(results: list[dict]) -> str:
    total = len(results)
    compliant = sum(1 for r in results if r["status"] == "COMPLIANT")
    oversize = sum(1 for r in results if r["status"] == "OVERSIZE")
    incons = sum(1 for r in results if r["status"] == "INCONSISTENT")

    stat_boxes = f"""
    <div style="display:flex;gap:0;margin-bottom:32px;border:1.5px solid #000;">
      <div style="flex:1;padding:18px 20px;text-align:center;border-right:1px solid #000;">
        <div style="font-size:10px;letter-spacing:2px;text-transform:uppercase;margin-bottom:6px;color:#000;">Total Checked</div>
        <div style="font-size:38px;font-weight:700;color:#000;font-family:'Times New Roman',serif;">{total}</div>
      </div>
      <div style="flex:1;padding:18px 20px;text-align:center;border-right:1px solid #000;">
        <div style="font-size:10px;letter-spacing:2px;text-transform:uppercase;margin-bottom:6px;color:#000;">Compliant</div>
        <div style="font-size:38px;font-weight:700;color:#000;font-family:'Times New Roman',serif;">{compliant}</div>
      </div>
      <div style="flex:1;padding:18px 20px;text-align:center;border-right:1px solid #000;">
        <div style="font-size:10px;letter-spacing:2px;text-transform:uppercase;margin-bottom:6px;color:#000;">Oversize</div>
        <div style="font-size:38px;font-weight:700;color:#000;font-family:'Times New Roman',serif;">{oversize}</div>
      </div>
      <div style="flex:1;padding:18px 20px;text-align:center;">
        <div style="font-size:10px;letter-spacing:2px;text-transform:uppercase;margin-bottom:6px;color:#000;">Inconsistent</div>
        <div style="font-size:38px;font-weight:700;color:#000;font-family:'Times New Roman',serif;">{incons}</div>
      </div>
    </div>
    """

    formula_block = """
    <div style="border:1.5px solid #000;padding:24px 28px;margin-bottom:32px;background:#fff;">
      <div style="font-size:11px;letter-spacing:2px;text-transform:uppercase;border-bottom:1px solid #000;
                  padding-bottom:8px;margin-bottom:20px;font-family:'Courier New',monospace;">
        Engineering Calculation Basis
      </div>
      <div style="margin-bottom:18px;">
        <span style="font-size:11px;letter-spacing:1px;text-transform:uppercase;font-family:'Courier New',monospace;">
          ① Velocity Compliance Check
        </span>
        <div style="margin-top:8px;font-size:15px;">
          \\( V_{\\text{measured}} \\leq V_{\\text{max}} \\quad \\Rightarrow \\quad \\textbf{PASS} \\quad / \\quad \\textbf{FAIL} \\)
        </div>
      </div>
      <div style="margin-bottom:18px;">
        <span style="font-size:11px;letter-spacing:1px;text-transform:uppercase;font-family:'Courier New',monospace;">
          ② Calculated Velocity from Geometry
        </span>
        <div style="margin-top:8px;font-size:15px;">
          \\( V_{\\text{calc}} = \\dfrac{Q}{A} \\quad \\text{where} \\quad
             A = \\begin{cases}
               \\pi \\left(\\dfrac{D}{2}\\right)^2 & \\text{round / flex} \\\\[6pt]
               W \\times H & \\text{rectangular}
             \\end{cases} \\)
        </div>
        <div style="margin-top:10px;font-size:15px;">
          \\( \\text{Flag INCONSISTENT if} \\quad \\left| V_{\\text{measured}} - V_{\\text{calc}} \\right| > 0.5 \\ \\text{m/s} \\)
        </div>
      </div>
      <div style="margin-bottom:18px;">
        <span style="font-size:11px;letter-spacing:1px;text-transform:uppercase;font-family:'Courier New',monospace;">
          ③ Pressure Drop — Darcy-Weisbach
        </span>
        <div style="margin-top:8px;font-size:15px;">
          \\( \\Delta P = \\lambda \\cdot \\dfrac{L}{D} \\cdot \\dfrac{\\rho \\, V^2}{2}
             \\qquad \\lambda = \\begin{cases} 0.020 & \\text{rigid} \\\\ 0.035 & \\text{flex} \\end{cases}
             \\qquad \\rho = 1.2 \\ \\text{kg/m}^3 \\qquad L = 1 \\ \\text{m} \\)
        </div>
      </div>
      <div>
        <span style="font-size:11px;letter-spacing:1px;text-transform:uppercase;font-family:'Courier New',monospace;">
          ④ Duct Size Optimization
        </span>
        <div style="margin-top:8px;font-size:15px;">
          \\( D_{\\min} = \\sqrt{\\dfrac{4 \\, Q}{\\pi \\, V_{\\max}}} \\quad \\text{(for round ducts)} \\)
        </div>
        <div style="margin-top:10px;font-size:13px;">
          <b>Flex ducts:</b> \\(\\{75, 100, 125, 150, 200, 225, 250, 300, 350\\} \\ \\text{mm}\\)<br>
          <b>Round rigid:</b> \\(\\{100, 125, 150, 160, 200, 250, 315, 400, 500\\} \\ \\text{mm}\\)<br>
          <b>Rectangular:</b> Width × Height from standard dimensions \\(\\{75, 90, 100, ...\\}\\)
        </div>
      </div>
      <div style="margin-top:20px;border-top:1px solid #000;padding-top:12px;">
        <span style="font-size:11px;letter-spacing:1px;text-transform:uppercase;font-family:'Courier New',monospace;">
          ⑤ Pressure Drop Reduction after Resizing
        </span>
        <div style="margin-top:8px;font-size:15px;">
          \\( \\frac{\\Delta P_{\\text{new}}}{\\Delta P_{\\text{old}}} =
             \\left( \\frac{D_{\\text{old}}}{D_{\\text{new}}} \\right)^5
             \\quad \\Rightarrow \\quad
             \\% \\text{ reduction} = \\left(1 - \\frac{1}{\\left(D_{\\text{new}}/D_{\\text{old}}\\right)^5}\\right) \\times 100 \\)
        </div>
      </div>
      <div style="margin-top:20px;border-top:1px solid #000;padding-top:12px;">
        <span style="font-size:11px;letter-spacing:1px;text-transform:uppercase;font-family:'Courier New',monospace;">
          ⑥ Reynolds Number (Turbulence Check)
        </span>
        <div style="margin-top:8px;font-size:15px;">
          \\( Re = \\dfrac{\\rho \\, V \\, D}{\\mu}
             \\qquad \\mu_{\\text{air}} = 1.81 \\times 10^{-5} \\ \\text{Pa·s}
             \\qquad Re > 4000 \\Rightarrow \\text{turbulent flow assumed} \\)
        </div>
      </div>
    </div>
    """

    def badge(status: str) -> str:
        if status == "COMPLIANT":
            return '<span style="border:1px solid #000;padding:2px 8px;font-size:10px;font-weight:700;letter-spacing:1px;font-family:\'Courier New\',monospace;">COMPLIANT</span>'
        if status == "OVERSIZE":
            return '<span style="border:1.5px solid #000;padding:2px 8px;font-size:10px;font-weight:700;letter-spacing:1px;font-family:\'Courier New\',monospace;background:#000;color:#fff;">OVERSIZE</span>'
        if status == "INCONSISTENT":
            return '<span style="border:1px solid #000;padding:2px 8px;font-size:10px;font-weight:700;letter-spacing:1px;font-family:\'Courier New\',monospace;">WARN</span>'
        return f'<span style="border:1px dashed #000;padding:2px 8px;font-size:10px;font-family:\'Courier New\',monospace;">{status}</span>'

    def row_bg(status: str) -> str:
        if status == "COMPLIANT":
            return "#f0fdf4"
        if status == "OVERSIZE":
            return "#fef2f2"
        if status == "INCONSISTENT":
            return "#fefce8"
        return "#ffffff"

    def opt_cell(result: dict) -> str:
        if result["optimization"] != "—":
            return f'<span style="font-weight:700;font-family:\'Courier New\',monospace;">{result["optimization"]}</span>'
        return result["optimization"]

    rows_html = ""
    for result in results:
        bg = row_bg(result["status"])
        v_max_str = f'{result["v_max"]:.2f}' if isinstance(result["v_max"], float) else result["v_max"]
        v_calc_str = f'{result["v_calc"]:.2f}' if isinstance(result["v_calc"], float) else result["v_calc"]
        dp_str = f'{result["delta_p"]:.2f}' if isinstance(result["delta_p"], float) else result["delta_p"]
        dv_str = f'{result["delta_v"]:.2f}' if isinstance(result["delta_v"], float) else result["delta_v"]
        rows_html += f"""
        <tr style="background:{bg};">
          <td class="td">{result["system_name"] or "—"}</td>
          <td class="td">{result["revit_id"]}</td>
          <td class="td">{result["duct_type"]}</td>
          <td class="td">{result["system_type"]}</td>
          <td class="td tdr">{result["v_measured"]:.2f}</td>
          <td class="td tdr">{v_max_str}</td>
          <td class="td tdr">{v_calc_str}</td>
          <td class="td tdr">{dp_str}</td>
          <td class="td tdr">{dv_str}</td>
          <td class="td" style="text-align:center;">{badge(result["status"])}</td>
          <td class="td">{opt_cell(result)}</td>
        </tr>"""

    results_table = f"""
    <div style="overflow-x:auto;margin-bottom:36px;">
      <table class="eng-table">
        <thead>
          <tr>
            <th class="th">System Name</th>
            <th class="th">Revit ID</th>
            <th class="th">Duct Type</th>
            <th class="th">System</th>
            <th class="th thr">V<sub>meas</sub> (m/s)</th>
            <th class="th thr">V<sub>max</sub> (m/s)</th>
            <th class="th thr">V<sub>calc</sub> (m/s)</th>
            <th class="th thr">&Delta;P (Pa/m)</th>
            <th class="th thr">&Delta;v (m/s)</th>
            <th class="th" style="text-align:center;">Status</th>
            <th class="th">Optimization</th>
          </tr>
        </thead>
        <tbody>{rows_html}</tbody>
      </table>
    </div>
    """

    sized_rows = [result for result in results if result["proposed_size"]]
    if sized_rows:
        opt_rows_html = ""
        for result in sized_rows:
            current_size = result.get("diam_mm", 0) or 0
            proposed_size = result["proposed_size"]
            dp_reduction = "—"
            if current_size > 0 and proposed_size > 0:
                ratio = (proposed_size / current_size) ** 5
                dp_reduction = f"~{(1 - 1 / ratio) * 100:.0f} %" if ratio > 1 else "—"
            re_current = (AIR_DENSITY * result["v_measured"] * (current_size / 1000.0)) / AIR_VISCOSITY if current_size > 0 else 0
            re_new = (AIR_DENSITY * (result["v_new"] or 0) * (proposed_size / 1000.0)) / AIR_VISCOSITY if proposed_size > 0 else 0
            row_color = "#fef2f2" if result["status"] == "OVERSIZE" else "#f0fdf4"
            opt_rows_html += f"""
            <tr style="background:{row_color};">
              <td class="td">{result["system_name"] or "—"}</td>
              <td class="td">{result["revit_id"]}</td>
              <td class="td">{result["duct_type"]}</td>
              <td class="td" style="text-align:center;">{badge(result["status"])}</td>
              <td class="td tdr">{result["current_size_label"]}</td>
              <td class="td tdr"><b>{result["proposed_size_label"]}</b></td>
              <td class="td tdr">{result["v_measured"]:.2f} m/s</td>
              <td class="td tdr"><b>{result["v_new"]:.2f} m/s</b></td>
              <td class="td tdr">{dp_reduction}</td>
              <td class="td tdr">{re_current:,.0f}</td>
              <td class="td tdr">{re_new:,.0f}</td>
            </tr>"""
        opt_summary = f"""
        <div style="margin-bottom:36px;">
          <div class="section-title">Sizing Proposals — All Ducts</div>
          <div style="overflow-x:auto;">
            <table class="eng-table">
              <thead>
                <tr>
                  <th class="th">System Name</th>
                  <th class="th">Revit ID</th>
                  <th class="th">Duct Type</th>
                  <th class="th" style="text-align:center;">Status</th>
                  <th class="th thr">Current Size</th>
                  <th class="th thr">Proposed Size</th>
                  <th class="th thr">Current V (m/s)</th>
                  <th class="th thr">Est. New V (m/s)</th>
                  <th class="th thr">&Delta;P Reduction</th>
                  <th class="th thr">Re (current)</th>
                  <th class="th thr">Re (proposed)</th>
                </tr>
              </thead>
              <tbody>{opt_rows_html}</tbody>
            </table>
          </div>
        </div>
        """
    else:
        opt_summary = """
        <div style="border:1px solid #000;padding:14px 20px;margin-bottom:32px;
                    font-family:'Courier New',monospace;font-size:12px;letter-spacing:0.5px;">
          &#10003;&nbsp; No sizing data available &mdash; check duct input table.
        </div>
        """

    today = date.today().strftime("%d %B %Y")
    html = (Path(__file__).parent / "report_template.html").read_text(encoding="utf-8")
    html = html.replace("__TODAY__", today)
    html = html.replace("__STAT_BOXES__", stat_boxes)
    html = html.replace("__FORMULA_BLOCK__", formula_block)
    html = html.replace("__RESULTS_TABLE__", results_table)
    html = html.replace("__OPT_SUMMARY__", opt_summary)
    return html


def build_pdf(results: list[dict]) -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=15 * mm,
        rightMargin=15 * mm,
        topMargin=15 * mm,
        bottomMargin=15 * mm,
    )
    styles = getSampleStyleSheet()
    story = []

    navy = colors.HexColor("#1a1a2e")
    green_bg = colors.HexColor("#dcfce7")
    red_bg = colors.HexColor("#fee2e2")
    amber_bg = colors.HexColor("#fef3c7")
    white = colors.white

    header_style = ParagraphStyle("header", fontSize=16, textColor=white, fontName="Helvetica-Bold", spaceAfter=4)
    sub_style = ParagraphStyle("sub", fontSize=9, textColor=colors.HexColor("#94a3b8"), fontName="Helvetica")
    today = date.today().strftime("%d %B %Y")

    header_data = [[Paragraph("QA/QC – Duct Velocity & Optimization Checker", header_style), Paragraph(f"Generated: {today}", sub_style)]]
    header_table = Table(header_data, colWidths=["70%", "30%"])
    header_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), navy),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 10 * mm))

    total = len(results)
    compliant = sum(1 for r in results if r["status"] == "COMPLIANT")
    oversize = sum(1 for r in results if r["status"] == "OVERSIZE")
    incons = sum(1 for r in results if r["status"] == "INCONSISTENT")

    summary_data = [
        ["Metric", "Count"],
        ["Total Ducts Checked", str(total)],
        ["Compliant", str(compliant)],
        ["Oversize", str(oversize)],
        ["Data Inconsistency", str(incons)],
    ]
    summary_table = Table(summary_data, colWidths=[80 * mm, 40 * mm])
    summary_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), navy),
        ("TEXTCOLOR", (0, 0), (-1, 0), white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor("#f8fafc"), white]),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#e2e8f0")),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(Paragraph("Compliance Summary", styles["Heading2"]))
    story.append(Spacer(1, 3 * mm))
    story.append(summary_table)
    story.append(Spacer(1, 8 * mm))

    detail_headers = [
        "System Name", "Revit ID", "Duct Type", "System",
        "V meas\n(m/s)", "V max\n(m/s)", "V calc\n(m/s)",
        "ΔP\n(Pa/m)", "Δv (m/s)", "Status", "Optimization",
    ]
    detail_data = [detail_headers]
    for result in results:
        v_max_str = f'{result["v_max"]:.2f}' if isinstance(result["v_max"], float) else str(result["v_max"])
        v_calc_str = f'{result["v_calc"]:.2f}' if isinstance(result["v_calc"], float) else str(result["v_calc"])
        dp_str = f'{result["delta_p"]:.2f}' if isinstance(result["delta_p"], float) else str(result["delta_p"])
        dv_str = f'{result["delta_v"]:.2f}' if isinstance(result["delta_v"], float) else str(result["delta_v"])
        detail_data.append([
            result["system_name"] or "—", result["revit_id"], result["duct_type"], result["system_type"],
            f'{result["v_measured"]:.2f}', v_max_str, v_calc_str, dp_str, dv_str, result["status"], result["optimization"],
        ])

    col_widths = [28 * mm, 18 * mm, 34 * mm, 16 * mm, 16 * mm, 16 * mm, 16 * mm, 16 * mm, 16 * mm, 18 * mm, 24 * mm]
    detail_table = Table(detail_data, colWidths=col_widths, repeatRows=1)
    detail_style = [
        ("BACKGROUND", (0, 0), (-1, 0), navy),
        ("TEXTCOLOR", (0, 0), (-1, 0), white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#e2e8f0")),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (4, 1), (8, -1), "RIGHT"),
        ("ALIGN", (9, 1), (9, -1), "CENTER"),
    ]
    for idx, result in enumerate(results, start=1):
        if result["status"] == "COMPLIANT":
            detail_style.append(("BACKGROUND", (0, idx), (-1, idx), green_bg))
        elif result["status"] == "OVERSIZE":
            detail_style.append(("BACKGROUND", (0, idx), (-1, idx), red_bg))
        elif result["status"] == "INCONSISTENT":
            detail_style.append(("BACKGROUND", (0, idx), (-1, idx), amber_bg))

    detail_table.setStyle(TableStyle(detail_style))
    story.append(Paragraph("Full Detail Results", styles["Heading2"]))
    story.append(Spacer(1, 3 * mm))
    story.append(detail_table)
    story.append(Spacer(1, 8 * mm))

    resized_rows = [result for result in results if result["proposed_size"]]
    if resized_rows:
        opt_headers = [
            "System Name", "Revit ID", "Duct Type", "Status", "Current Size",
            "Proposed Size", "Current V (m/s)", "Expected V after resize (m/s)", "ΔP reduction est.",
        ]
        opt_data = [opt_headers]
        for result in resized_rows:
            current_size = result.get("diam_mm", 0) or 0
            proposed_size = result["proposed_size"]
            dp_reduction = "—"
            if current_size > 0 and proposed_size > 0:
                ratio = (proposed_size / current_size) ** 5
                dp_reduction = f"~{(1 - 1 / ratio) * 100:.0f}% reduction" if ratio > 1 else "—"
            opt_data.append([
                result["system_name"] or "—",
                result["revit_id"],
                result["duct_type"],
                result["status"],
                result["current_size_label"],
                result["proposed_size_label"],
                f'{result["v_measured"]:.2f}',
                f'{result["v_new"]:.2f}' if result["v_new"] is not None else "—",
                dp_reduction,
            ])
        opt_col_widths = [34 * mm, 20 * mm, 34 * mm, 18 * mm, 28 * mm, 28 * mm, 24 * mm, 46 * mm, 34 * mm]
        opt_table = Table(opt_data, colWidths=opt_col_widths, repeatRows=1)
        opt_style = [
            ("BACKGROUND", (0, 0), (-1, 0), navy),
            ("TEXTCOLOR", (0, 0), (-1, 0), white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#e2e8f0")),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("ALIGN", (3, 1), (3, -1), "CENTER"),
        ]
        for idx, result in enumerate(resized_rows, start=1):
            if result["status"] == "OVERSIZE":
                opt_style.append(("BACKGROUND", (0, idx), (-1, idx), red_bg))
            elif result["status"] == "INCONSISTENT":
                opt_style.append(("BACKGROUND", (0, idx), (-1, idx), amber_bg))
            else:
                opt_style.append(("BACKGROUND", (0, idx), (-1, idx), green_bg))

        opt_table.setStyle(TableStyle(opt_style))
        story.append(Paragraph("Sizing Proposals — All Ducts", styles["Heading2"]))
        story.append(Spacer(1, 3 * mm))
        story.append(opt_table)
        story.append(Spacer(1, 8 * mm))

    footer_style = ParagraphStyle("footer", fontSize=7, textColor=colors.HexColor("#94a3b8"), fontName="Helvetica")
    story.append(Paragraph(
        "Calculation basis: ASHRAE Handbook – HVAC Systems &amp; Equipment | "
        "SMACNA HVAC Duct Construction Standards | "
        "Darcy-Weisbach friction model | λ = 0.02 (rigid), 0.035 (flex) | ρ = 1.2 kg/m³",
        footer_style,
    ))

    doc.build(story)
    return buffer.getvalue()


def build_pdf_from_html(results: list[dict]) -> bytes:
    wkhtmltopdf = shutil.which("wkhtmltopdf")
    if not wkhtmltopdf:
        raise RuntimeError("wkhtmltopdf is not available on PATH")

    html = build_html(results)
    with TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        html_path = temp_path / "duct_velocity_report.html"
        pdf_path = temp_path / "duct_velocity_report.pdf"
        html_path.write_text(html, encoding="utf-8")

        cmd = [
            wkhtmltopdf,
            "--quiet",
            "--encoding", "utf-8",
            "--page-size", "A4",
            "--orientation", "Landscape",
            "--margin-top", "15mm",
            "--margin-right", "15mm",
            "--margin-bottom", "15mm",
            "--margin-left", "15mm",
            "--enable-local-file-access",
            "--javascript-delay", "2000",
            "--print-media-type",
            str(html_path),
            str(pdf_path),
        ]
        completed = subprocess.run(cmd, capture_output=True, text=True, check=False)
        if completed.returncode != 0 or not pdf_path.exists():
            stderr = (completed.stderr or "").strip()
            raise RuntimeError(f"wkhtmltopdf failed: {stderr or 'unknown error'}")

        return pdf_path.read_bytes()
