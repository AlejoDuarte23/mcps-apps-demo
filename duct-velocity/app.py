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

    @vkt.DataView("Flex Duct — Compliance Summary")
    def qa_qc_flex_duct_checks(self, params, **kwargs) -> vkt.DataResult:
        """
        DataView showing Flex Duct compliance status.
        """
        duct_rows = list(params.duct_table) if params.duct_table else []
        system_type = params.system_type or "Supply"
        results = process_ducts(duct_rows, system_type=system_type)

        # Filter: flex ducts
        flex_all = [r for r in results if "flex" in r["duct_type"].lower()]
        flex_compliant = [r for r in flex_all if r["status"] == "COMPLIANT"]
        flex_non_compliant = [r for r in flex_all if r["status"] in ("OVERSIZE", "INCONSISTENT")]

        logger.info(f"Flex ducts — Compliant: {len(flex_compliant)}, Non-compliant: {len(flex_non_compliant)}")

        # Compliant group
        compliant_group = (
            vkt.DataGroup(
                **{
                    f"compliant_{i}": vkt.DataItem(
                        r["revit_id"],
                        "Compliant",
                        status=vkt.DataStatus.SUCCESS,
                    )
                    for i, r in enumerate(flex_compliant)
                }
            )
            if flex_compliant
            else vkt.DataGroup(
                empty=vkt.DataItem("Compliant flex ducts", "None", status=vkt.DataStatus.WARNING)
            )
        )

        # Non-compliant group
        non_compliant_group = (
            vkt.DataGroup(
                **{
                    f"non_compliant_{i}": vkt.DataItem(
                        r["revit_id"],
                        r["status"],
                        status=vkt.DataStatus.ERROR if r["status"] == "OVERSIZE" else vkt.DataStatus.WARNING,
                        subgroup=vkt.DataGroup(
                            type=vkt.DataItem("Type", r["duct_type"]),
                            current_size=vkt.DataItem("Current Size (Diameter)", r["current_size_label"]),
                            proposed_size=vkt.DataItem("Proposed Size (Diameter)", r["optimization"] if r["optimization"] != "—" else "No standard size available"),
                        ),
                    )
                    for i, r in enumerate(flex_non_compliant)
                }
            )
            if flex_non_compliant
            else vkt.DataGroup(
                empty=vkt.DataItem("Non-compliant flex ducts", "None — all comply", status=vkt.DataStatus.SUCCESS)
            )
        )

        data = vkt.DataGroup(
            compliant=vkt.DataItem(
                "Compliant Flex Ducts",
                f"{len(flex_compliant)} duct(s)",
                status=vkt.DataStatus.SUCCESS if flex_compliant else vkt.DataStatus.WARNING,
                subgroup=compliant_group,
            ),
            non_compliant=vkt.DataItem(
                "Requires Resizing",
                f"{len(flex_non_compliant)} duct(s)",
                status=vkt.DataStatus.ERROR if flex_non_compliant else vkt.DataStatus.SUCCESS,
                subgroup=non_compliant_group,
            ),
        )
        return vkt.DataResult(data)

    @vkt.DataView("Rigid Duct — Compliance Summary")
    def qa_qc_rigid_duct_checks(self, params, **kwargs) -> vkt.DataResult:
        """
        DataView showing Rigid Duct compliance status.
        """
        duct_rows = list(params.duct_table) if params.duct_table else []
        system_type = params.system_type or "Supply"
        results = process_ducts(duct_rows, system_type=system_type)

        # Filter: rigid ducts (round + rectangular)
        rigid_all = [r for r in results if "rigid" in r["duct_type"].lower()]
        rigid_compliant = [r for r in rigid_all if r["status"] == "COMPLIANT"]
        rigid_non_compliant = [r for r in rigid_all if r["status"] in ("OVERSIZE", "INCONSISTENT")]

        logger.info(f"Rigid ducts — Compliant: {len(rigid_compliant)}, Non-compliant: {len(rigid_non_compliant)}")

        # Compliant group
        compliant_group = (
            vkt.DataGroup(
                **{
                    f"compliant_{i}": vkt.DataItem(
                        r["revit_id"],
                        "Compliant",
                        status=vkt.DataStatus.SUCCESS,
                    )
                    for i, r in enumerate(rigid_compliant)
                }
            )
            if rigid_compliant
            else vkt.DataGroup(
                empty=vkt.DataItem("Compliant rigid ducts", "None", status=vkt.DataStatus.WARNING)
            )
        )

        # Non-compliant group (with dynamic labels based on duct shape)
        def make_non_compliant_item(i, r):
            is_round = "round" in r["duct_type"].lower()
            current_label = "Current Size (Diameter)" if is_round else "Current Size (Width x Height)"
            proposed_label = "Proposed Size (Diameter)" if is_round else "Proposed Size (Width x Height)"

            return vkt.DataItem(
                r["revit_id"],
                r["status"],
                status=vkt.DataStatus.ERROR if r["status"] == "OVERSIZE" else vkt.DataStatus.WARNING,
                subgroup=vkt.DataGroup(
                    type=vkt.DataItem("Type", r["duct_type"]),
                    current_size=vkt.DataItem(current_label, r["current_size_label"]),
                    proposed_size=vkt.DataItem(proposed_label, r["optimization"] if r["optimization"] != "—" else "No standard size available"),
                ),
            )

        non_compliant_group = (
            vkt.DataGroup(
                **{f"non_compliant_{i}": make_non_compliant_item(i, r) for i, r in enumerate(rigid_non_compliant)}
            )
            if rigid_non_compliant
            else vkt.DataGroup(
                empty=vkt.DataItem("Non-compliant rigid ducts", "None — all comply", status=vkt.DataStatus.SUCCESS)
            )
        )

        data = vkt.DataGroup(
            compliant=vkt.DataItem(
                "Compliant Rigid Ducts",
                f"{len(rigid_compliant)} duct(s)",
                status=vkt.DataStatus.SUCCESS if rigid_compliant else vkt.DataStatus.WARNING,
                subgroup=compliant_group,
            ),
            non_compliant=vkt.DataItem(
                "Requires Resizing",
                f"{len(rigid_non_compliant)} duct(s)",
                status=vkt.DataStatus.ERROR if rigid_non_compliant else vkt.DataStatus.SUCCESS,
                subgroup=non_compliant_group,
            ),
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
