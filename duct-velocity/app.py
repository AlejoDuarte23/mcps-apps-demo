import logging
from datetime import date

import viktor as vkt

from helpers import REFERENCE_VELOCITY_LIMITS, build_html, build_pdf, build_pdf_from_html, process_ducts

logger = logging.getLogger("viktor")


# ---------------------------------------------------------------------------
# Parametrization
# ---------------------------------------------------------------------------

class Parametrization(vkt.Parametrization):

    # ── 1. App header ──────────────────────────────────────────────────────
    header_text = vkt.Text(
        "# QA/QC – Duct Velocity & Optimization Checker\n\n"
        "This tool validates HVAC duct air velocities against **ASHRAE Handbook HVAC Systems & Equipment** "
        "and **SMACNA HVAC Duct Construction Standards** maximum allowable limits. "
        "For each duct, the app computes the calculated velocity from geometry, estimates the pressure drop "
        "using the Darcy-Weisbach equation, and proposes an optimised standard duct size for any non-compliant duct."
    )

    # ── 2. Duct input table ────────────────────────────────────────────────
    duct_table = vkt.Table(
        "Duct Input Table",
        default=[
            {"system_name": "Supply Air 01", "revit_id": "D-001", "duct_type": "Flex Duct",                "airflow_q": 100,  "duct_diameter": 200, "duct_height": 0,   "measured_velocity": 3.18},
            {"system_name": "Supply Air 01", "revit_id": "D-002", "duct_type": "Flex Duct",                "airflow_q": 150,  "duct_diameter": 250, "duct_height": 0,   "measured_velocity": 3.06},
            {"system_name": "Supply Air 01", "revit_id": "D-003", "duct_type": "Rigid Duct (Round)",       "airflow_q": 200,  "duct_diameter": 250, "duct_height": 0,   "measured_velocity": 4.07},
            {"system_name": "Supply Air 01", "revit_id": "D-004", "duct_type": "Rigid Duct (Round)",       "airflow_q": 300,  "duct_diameter": 315, "duct_height": 0,   "measured_velocity": 3.85},
            {"system_name": "Supply Air 01", "revit_id": "D-005", "duct_type": "Rigid Duct (Rectangular)", "airflow_q": 500,  "duct_diameter": 400, "duct_height": 300, "measured_velocity": 4.17},
            {"system_name": "Supply Air 01", "revit_id": "D-006", "duct_type": "Rigid Duct (Rectangular)", "airflow_q": 1000, "duct_diameter": 500, "duct_height": 400, "measured_velocity": 5.00},
            {"system_name": "Supply Air 02", "revit_id": "D-007", "duct_type": "Flex Duct",                "airflow_q": 50,   "duct_diameter": 125, "duct_height": 0,   "measured_velocity": 4.07},
            {"system_name": "Supply Air 02", "revit_id": "D-008", "duct_type": "Rigid Duct (Round)",       "airflow_q": 600,  "duct_diameter": 400, "duct_height": 0,   "measured_velocity": 4.77},
            {"system_name": "Supply Air 02", "revit_id": "D-009", "duct_type": "Rigid Duct (Rectangular)", "airflow_q": 1500, "duct_diameter": 600, "duct_height": 400, "measured_velocity": 6.25},
            {"system_name": "Supply Air 02", "revit_id": "D-010", "duct_type": "Flex Duct",                "airflow_q": 250,  "duct_diameter": 315, "duct_height": 0,   "measured_velocity": 3.22},
        ],
    )
    duct_table.system_name       = vkt.TextField("System Name")
    duct_table.revit_id          = vkt.TextField("Revit ID")
    duct_table.duct_type         = vkt.OptionField("Duct Type", options=["Flex Duct", "Rigid Duct (Round)", "Rigid Duct (Rectangular)"])
    duct_table.airflow_q         = vkt.NumberField("Flow (L/s)")
    duct_table.duct_diameter     = vkt.NumberField("Diameter or Width (mm)")
    duct_table.duct_height       = vkt.NumberField("Duct Height (mm)")
    duct_table.measured_velocity = vkt.NumberField("Measured Velocity (m/s)")

    # ── 3a. Text below the table — context note ────────────────────────────
    table_note = vkt.Text(
        "The **System Name** must be provided for every row. "
        "The **Duct Height** column is only required for **Rigid Duct (Rectangular)**. "
        "For round and flex ducts, leave it as `0`. "
        "Use **Flow (L/s)** and **Measured Velocity (m/s)**. "
        "All ducts in this table will be validated against the same **System Type** selected below."
    )

    # ── 3b. Text above the system type selector ────────────────────────────
    system_type_label = vkt.Text(
        "## System Type\n\n"
        "Select the HVAC system type that applies to **all ducts** in the table above. "
        "This determines which ASHRAE / SMACNA maximum velocity limit is used for each duct type:\n\n"
        "- **Supply** — air delivered from AHU to spaces (highest velocity limits)\n"
        "- **Return** — air drawn back to AHU from spaces (lower limits for acoustic comfort)\n"
        "- **Exhaust** — air expelled from the building (intermediate limits)"
    )

    # ── 3c. System type option field ───────────────────────────────────────
    system_type = vkt.OptionField(
        "System Type",
        options=["Supply", "Return", "Exhaust"],
        default="Supply",
        description="Applied to all ducts in the table above when looking up ASHRAE/SMACNA velocity limits.",
    )

    # ── 3. Engineering calculations explanation ────────────────────────────
    calc_text = vkt.Text(
        "## Engineering Checks Performed\n\n"
        "Velocity limits are sourced from the internal **ASHRAE / SMACNA reference dataset** "
        "(view the *Reference Velocity Limits* tab). "
        "The **Duct Type** column in the input table and the **System Type** selected below "
        "are combined to retrieve the correct V_max for each duct — for example, a Supply Flex Duct "
        "has a different limit than a Return Flex Duct.\n\n"
        "For each duct row, the following four checks are executed:\n\n"
        "**① Velocity Check:** `V_measured ≤ V_max` → **COMPLIANT** / **OVERSIZE**\n\n"
        "**② Calculated Velocity from Q and cross-section:**\n"
        "- `V_calc = Q / A` (Q converted from L/s to m³/s)\n"
        "- `A = π(D/2)²` for round/flex ducts (D in metres)\n"
        "- `A = W × H` for rectangular ducts (W, H in metres)\n"
        "- Flagged as **DATA INCONSISTENCY** if `|V_measured − V_calc| > 0.5 m/s`\n\n"
        "**③ Pressure Drop Estimate (Darcy-Weisbach):**\n"
        "- `ΔP = λ × (L/D) × (ρ × V² / 2)`\n"
        "- Default duct length L = 1 m, air density ρ = 1.2 kg/m³\n"
        "- λ = 0.02 for rigid ducts, λ = 0.035 for flex ducts\n\n"
        "**④ Optimization Proposal (all ducts):**\n"
        "- Back-calculate minimum required diameter: `D_min = √(4Q / (π × V_max))`\n"
        "- Suggest the next standard duct size from: `[100, 125, 150, 160, 200, 250, 315, 400, 500]` mm\n"
        "- For rigid rectangular ducts the demo proposes a square target size such as `200 x 200 mm`"
    )

    # ── 6. Export explanation ──────────────────────────────────────────────
    export_text = vkt.Text(
        "## PDF Report Export\n\n"
        "The downloaded PDF report contains:\n"
        "- **Dark navy header** with app title and generation date\n"
        "- **Compliance Summary Stats Table** (total / compliant / non-compliant / inconsistent)\n"
        "- **Full Detail Table** with green/red/amber row colouring per duct result\n"
        "- **Optimization Proposals Table** with current size, proposed size, current velocity, "
        "expected velocity after resizing, and estimated ΔP reduction\n"
        "- **System Name** shown in the detailed tables for Autodesk writeback traceability\n"
        "- **Footer** citing ASHRAE + SMACNA as the calculation basis"
    )

    # ── 7. Download button ─────────────────────────────────────────────────
    download_btn = vkt.DownloadButton("Download PDF Report", method="download_pdf")


# ---------------------------------------------------------------------------
# Controller
# ---------------------------------------------------------------------------

class Controller(vkt.Controller):
    parametrization = Parametrization

    @vkt.WebView("QA/QC Results", duration_guess=3)
    def view_results(self, params, **kwargs) -> vkt.WebResult:
        """Compute all engineering checks and render the HTML results view."""
        duct_rows = list(params.duct_table) if params.duct_table else []
        system_type = params.system_type or "Supply"

        logger.info(f"Duct rows received: {len(duct_rows)}, system type: {system_type}")

        if not duct_rows:
            html = "<html><body style='font-family:Courier New;padding:32px;'>" \
                   "<h2>No duct data entered yet.</h2></body></html>"
            return vkt.WebResult(html=html)

        results = process_ducts(duct_rows, system_type=system_type)
        html = build_html(results)
        return vkt.WebResult(html=html)

    @vkt.TableView("Reference Velocity Limits (ASHRAE / SMACNA)")
    def view_reference_table(self, params, **kwargs) -> vkt.TableResult:
        """
        Display the internal ASHRAE / SMACNA reference velocity limits as a
        read-only TableView. Users cannot edit this data — it is hardcoded
        from the standards and used automatically during QA/QC checks.
        """
        headers = [
            vkt.TableHeader("Duct Type",           align="left"),
            vkt.TableHeader("System Type",          align="center"),
            vkt.TableHeader("Max Velocity (m/s)",   align="right"),
            vkt.TableHeader("NC Limit",             align="right"),
            vkt.TableHeader("Standard",             align="center"),
        ]

        # Map duct type to the relevant standard citation
        standard_map = {
            "flex duct":                "SMACNA §2 / ASHRAE 2020",
            "rigid duct (round)":       "SMACNA §3 / ASHRAE 2020",
            "rigid duct (rectangular)": "SMACNA §4 / ASHRAE 2020",
        }

        rows = []
        for entry in REFERENCE_VELOCITY_LIMITS:
            dt  = entry["duct_type"]
            st  = entry["system_type"]
            vm  = entry["v_max"]
            nc  = entry["nc_limit"]
            std = standard_map.get(dt.lower(), "ASHRAE / SMACNA")

            # Colour-code system type cell for quick visual scanning
            if st == "Supply":
                st_cell = vkt.TableCell(st, background_color=vkt.Color(220, 242, 255))
            elif st == "Return":
                st_cell = vkt.TableCell(st, background_color=vkt.Color(240, 253, 244))
            else:  # Exhaust
                st_cell = vkt.TableCell(st, background_color=vkt.Color(254, 252, 232))

            rows.append([dt, st_cell, vm, nc, std])

        logger.info(f"Reference table rendered: {len(rows)} entries")
        return vkt.TableResult(rows, column_headers=headers)

    @vkt.DataView("Flex Duct — Sections to Change")
    def qa_qc_flex_duct_checks(self, params, **kwargs) -> vkt.DataResult:
        """
        DataView showing only Flex Duct sections that require changes:
        FAIL (velocity exceeded) or INCONSISTENT (data mismatch).
        Groups results by issue type with full engineering detail per duct.
        """
        duct_rows = list(params.duct_table) if params.duct_table else []
        system_type = params.system_type or "Supply"
        results = process_ducts(duct_rows, system_type=system_type)

        # Filter: flex ducts that need attention
        flex_fail   = [r for r in results if "flex" in r["duct_type"].lower() and r["status"] == "OVERSIZE"]
        flex_incons = [r for r in results if "flex" in r["duct_type"].lower() and r["status"] == "INCONSISTENT"]
        flex_pass   = [r for r in results if "flex" in r["duct_type"].lower() and r["status"] == "COMPLIANT"]

        logger.info(f"Flex ducts — OVERSIZE: {len(flex_fail)}, INCONSISTENT: {len(flex_incons)}, COMPLIANT: {len(flex_pass)}")

        def duct_item(r: dict) -> vkt.DataItem:
            """Build a DataItem for a single duct with its key metrics as subgroup."""
            v_max_str  = f'{r["v_max"]:.2f} m/s'  if isinstance(r["v_max"],  float) else str(r["v_max"])
            v_calc_str = f'{r["v_calc"]:.2f} m/s'  if isinstance(r["v_calc"], float) else str(r["v_calc"])
            dp_str     = f'{r["delta_p"]:.2f} Pa/m' if isinstance(r["delta_p"],float) else str(r["delta_p"])
            opt_str    = r["optimization"] if r["optimization"] != "—" else "No standard size available"

            return vkt.DataItem(
                r["revit_id"],
                r["status"],
                subgroup=vkt.DataGroup(
                    system_name = vkt.DataItem("System Name",         r["system_name"] or "—"),
                    flow_q      = vkt.DataItem("Flow",                r["airflow_q_ls"], suffix="L/s"),
                    size_now    = vkt.DataItem("Current Diameter",    r["current_size_label"]),
                    v_measured  = vkt.DataItem("Measured Velocity",   r["v_measured"],  suffix="m/s"),
                    v_max      = vkt.DataItem("Max Allowed (ASHRAE)", v_max_str),
                    v_calc     = vkt.DataItem("Calculated Velocity",  v_calc_str),
                    delta_p    = vkt.DataItem("Pressure Drop",        dp_str),
                    proposal   = vkt.DataItem("Proposed Diameter",    opt_str),
                )
            )

        # ── OVERSIZE group ────────────────────────────────────────────────
        if flex_fail:
            fail_kwargs = {f"duct_{i}": duct_item(r) for i, r in enumerate(flex_fail)}
            fail_group = vkt.DataGroup(**fail_kwargs)
            fail_item = vkt.DataItem(
                f"Velocity Exceeded — {len(flex_fail)} duct(s) must be resized",
                "",
                subgroup=fail_group,
                status=vkt.DataStatus.ERROR,
            )
        else:
            fail_item = vkt.DataItem(
                "Velocity Exceeded",
                "None — all flex ducts comply",
                status=vkt.DataStatus.SUCCESS,
            )

        # ── INCONSISTENT group ────────────────────────────────────────────
        if flex_incons:
            incons_kwargs = {f"duct_{i}": duct_item(r) for i, r in enumerate(flex_incons)}
            incons_group = vkt.DataGroup(**incons_kwargs)
            incons_item = vkt.DataItem(
                f"Data Inconsistency — {len(flex_incons)} duct(s) need field verification",
                "",
                subgroup=incons_group,
                status=vkt.DataStatus.WARNING,
            )
        else:
            incons_item = vkt.DataItem(
                "Data Inconsistency",
                "None — all flex duct measurements are consistent",
                status=vkt.DataStatus.SUCCESS,
            )

        # ── Summary header ────────────────────────────────────────────────
        total_flex    = len([r for r in results if "flex" in r["duct_type"].lower()])
        needs_change  = len(flex_fail) + len(flex_incons)
        summary_item  = vkt.DataItem(
            "Flex Duct Summary",
            f"{needs_change} of {total_flex} section(s) require action",
            status=vkt.DataStatus.ERROR if needs_change > 0 else vkt.DataStatus.SUCCESS,
        )

        data = vkt.DataGroup(
            summary  = summary_item,
            failures = fail_item,
            warnings = incons_item,
        )
        return vkt.DataResult(data)

    @vkt.DataView("Rigid Duct — Sections to Change")
    def qa_qc_rigid_duct_checks(self, params, **kwargs) -> vkt.DataResult:
        """
        DataView showing only Rigid Duct sections (Round + Rectangular) that
        require changes: FAIL (velocity exceeded) or INCONSISTENT (data mismatch).
        Splits results by sub-type: Round vs Rectangular.
        """
        duct_rows = list(params.duct_table) if params.duct_table else []
        system_type = params.system_type or "Supply"
        results = process_ducts(duct_rows, system_type=system_type)

        # Filter: rigid ducts only (round + rectangular)
        rigid_all   = [r for r in results if "rigid" in r["duct_type"].lower()]
        round_fail  = [r for r in rigid_all if "round" in r["duct_type"].lower() and r["status"] == "OVERSIZE"]
        round_incon = [r for r in rigid_all if "round" in r["duct_type"].lower() and r["status"] == "INCONSISTENT"]
        rect_fail   = [r for r in rigid_all if "rectangular" in r["duct_type"].lower() and r["status"] == "OVERSIZE"]
        rect_incon  = [r for r in rigid_all if "rectangular" in r["duct_type"].lower() and r["status"] == "INCONSISTENT"]

        logger.info(
            f"Rigid ducts — Round OVERSIZE: {len(round_fail)}, Round INCON: {len(round_incon)}, "
            f"Rect OVERSIZE: {len(rect_fail)}, Rect INCON: {len(rect_incon)}"
        )

        def duct_item(r: dict) -> vkt.DataItem:
            """Build a DataItem for a single rigid duct with engineering detail."""
            v_max_str  = f'{r["v_max"]:.2f} m/s'   if isinstance(r["v_max"],  float) else str(r["v_max"])
            v_calc_str = f'{r["v_calc"]:.2f} m/s'   if isinstance(r["v_calc"], float) else str(r["v_calc"])
            dp_str     = f'{r["delta_p"]:.2f} Pa/m'  if isinstance(r["delta_p"],float) else str(r["delta_p"])
            opt_str    = r["optimization"] if r["optimization"] != "—" else "No standard size available"
            d_curr     = int(r.get("diam_mm", 0) or 0)
            d_prop     = r.get("proposed_size")
            dp_red     = "—"
            if d_curr > 0 and d_prop and d_prop > d_curr:
                ratio  = (d_prop / d_curr) ** 5
                dp_red = f"~{(1 - 1/ratio)*100:.0f}% reduction"

            return vkt.DataItem(
                r["revit_id"],
                r["status"],
                subgroup=vkt.DataGroup(
                    system_name = vkt.DataItem("System Name",         r["system_name"] or "—"),
                    flow_q      = vkt.DataItem("Flow",                r["airflow_q_ls"], suffix="L/s"),
                    size_now    = vkt.DataItem("Current Size",        r["current_size_label"]),
                    v_measured = vkt.DataItem("Measured Velocity",    r["v_measured"], suffix="m/s"),
                    v_max      = vkt.DataItem("Max Allowed (ASHRAE)", v_max_str),
                    v_calc     = vkt.DataItem("Calculated Velocity",  v_calc_str),
                    delta_p    = vkt.DataItem("Pressure Drop",        dp_str),
                    proposal   = vkt.DataItem("Proposed Size",        opt_str),
                    dp_saving  = vkt.DataItem("Est. ΔP Reduction",    dp_red),
                )
            )

        def make_group_item(label: str, rows: list, status_val) -> vkt.DataItem:
            """Build a collapsible DataItem group for a list of duct results."""
            if rows:
                kwargs_dict = {f"duct_{i}": duct_item(r) for i, r in enumerate(rows)}
                return vkt.DataItem(label, "", subgroup=vkt.DataGroup(**kwargs_dict), status=status_val)
            return vkt.DataItem(label, "None — compliant", status=vkt.DataStatus.SUCCESS)

        # ── Round sub-section ─────────────────────────────────────────────
        round_fail_item  = make_group_item(
            f"Round — Velocity Exceeded ({len(round_fail)} duct(s))", round_fail,  vkt.DataStatus.ERROR)
        round_incon_item = make_group_item(
            f"Round — Data Inconsistency ({len(round_incon)} duct(s))", round_incon, vkt.DataStatus.WARNING)

        round_section = vkt.DataItem(
            "Rigid Duct (Round)",
            f"{len(round_fail) + len(round_incon)} section(s) need action",
            subgroup=vkt.DataGroup(
                fail  = round_fail_item,
                incon = round_incon_item,
            ),
            status=vkt.DataStatus.ERROR if (round_fail or round_incon) else vkt.DataStatus.SUCCESS,
        )

        # ── Rectangular sub-section ───────────────────────────────────────
        rect_fail_item  = make_group_item(
            f"Rectangular — Velocity Exceeded ({len(rect_fail)} duct(s))", rect_fail,  vkt.DataStatus.ERROR)
        rect_incon_item = make_group_item(
            f"Rectangular — Data Inconsistency ({len(rect_incon)} duct(s))", rect_incon, vkt.DataStatus.WARNING)

        rect_section = vkt.DataItem(
            "Rigid Duct (Rectangular)",
            f"{len(rect_fail) + len(rect_incon)} section(s) need action",
            subgroup=vkt.DataGroup(
                fail  = rect_fail_item,
                incon = rect_incon_item,
            ),
            status=vkt.DataStatus.ERROR if (rect_fail or rect_incon) else vkt.DataStatus.SUCCESS,
        )

        # ── Overall summary ───────────────────────────────────────────────
        total_rigid   = len(rigid_all)
        needs_change  = len(round_fail) + len(round_incon) + len(rect_fail) + len(rect_incon)
        summary_item  = vkt.DataItem(
            "Rigid Duct Summary",
            f"{needs_change} of {total_rigid} section(s) require action",
            status=vkt.DataStatus.ERROR if needs_change > 0 else vkt.DataStatus.SUCCESS,
        )

        data = vkt.DataGroup(
            summary      = summary_item,
            round_ducts  = round_section,
            rect_ducts   = rect_section,
        )
        return vkt.DataResult(data)

    def download_pdf(self, params, **kwargs) -> vkt.DownloadResult:
        """Generate and return the PDF QA/QC report."""
        duct_rows = list(params.duct_table) if params.duct_table else []
        system_type = params.system_type or "Supply"

        logger.info(f"Generating PDF for {len(duct_rows)} duct rows, system type: {system_type}")

        if not duct_rows:
            raise vkt.UserError("Please enter at least one duct row before downloading the report.")

        results = process_ducts(duct_rows, system_type=system_type)
        try:
            pdf_bytes = build_pdf_from_html(results)
            logger.info("PDF generated from HTML via wkhtmltopdf")
        except Exception as exc:
            logger.warning(f"wkhtmltopdf render failed, falling back to ReportLab PDF: {exc}")
            pdf_bytes = build_pdf(results)
        today = date.today().strftime("%Y-%m-%d")
        return vkt.DownloadResult(
            file_content=pdf_bytes,
            file_name=f"duct_velocity_qaqc_{today}.pdf",
        )
