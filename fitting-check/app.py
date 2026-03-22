import logging
from datetime import date

import viktor as vkt

from helpers import (
    REQUIRED_FAMILY_NAME,
    analyze_fittings,
    build_html_report,
    build_pdf_report,
    normalize_rows,
)

logger = logging.getLogger("viktor")


class Parametrization(vkt.Parametrization):
    intro = vkt.Text(
        "# QA/QC – Fittings Check\n\n"
        "This app checks if duct fittings use company standard Revit families and if all fittings have valid pressure drop values.\n"
        "It returns compliant elements and fittings that require a family update to the company standard transition family."
    )

    fittings_table = vkt.Table(
        "Fittings Input Table",
        default=[
            {
                "system_name": "Mechanical Supply Air 17",
                "revit_element_id": "8452101",
                "family_name": "M_Round Duct Transition - Angle",
                "type_name": "45 Degree",
                "pressure_drop": 11.4,
            },
            {
                "system_name": "Mechanical Supply Air 17",
                "revit_element_id": "8452102",
                "family_name": "M_Rectangular Duct Tee",
                "type_name": "Standard",
                "pressure_drop": 0.0,
            },
            {
                "system_name": "Mechanical Supply Air 17",
                "revit_element_id": "8452103",
                "family_name": "M_Rectangular to Round Duct Transsition - Angle",
                "type_name": "45 Degree",
                "pressure_drop": 8.9,
            },
            {
                "system_name": "Mechanical Supply Air 17",
                "revit_element_id": "8452104",
                "family_name": "M_Round",
                "type_name": "Tap",
                "pressure_drop": 0.0,
            },
        ],
    )
    fittings_table.system_name = vkt.TextField("System Name")
    fittings_table.revit_element_id = vkt.TextField("Revit Element ID")
    fittings_table.family_name = vkt.TextField("Family Name")
    fittings_table.type_name = vkt.TextField("Type Name")
    fittings_table.pressure_drop = vkt.NumberField("Pressure Drop (Pa)")

    input_note = vkt.Text(
        "## Input note\n\n"
        "Provide one row per duct fitting.\n"
        "The pressure drop input is checked for all fittings.\n"
        f"The company standard transition family used in this demo is `{REQUIRED_FAMILY_NAME}`.\n"
        "Type Name is included for report traceability and Autodesk writeback."
    )

    checks_note = vkt.Text(
        "## Checks in this app\n\n"
        "1. Company standard family check.\n"
        "2. Pressure drop check for all fitting rows.\n"
        "3. Web report with input table, plots, pressure drop review, and wrong family review.\n"
        "4. One grouped data view with compliant fittings and fittings that require a family change."
    )

    download_note = vkt.Text(
        "## PDF report\n\n"
        "The PDF contains the input table, pressure drop plot, family breakdown pie chart, pressure drop review, and wrong family assignment table."
    )
    download_button = vkt.DownloadButton("Download PDF Report", method="download_pdf")


class Controller(vkt.Controller):
    parametrization = Parametrization

    def _analysis(self, params) -> dict:
        rows = normalize_rows(list(params.fittings_table) if params.fittings_table else [])
        return analyze_fittings(rows)

    @vkt.WebView("Engineering Report", duration_guess=2)
    def engineering_report(self, params, **kwargs) -> vkt.WebResult:
        analysis = self._analysis(params)
        if not analysis["rows"]:
            return vkt.WebResult(html="<html><body style='font-family:Arial;padding:24px;'>No fitting data entered yet.</body></html>")
        html = build_html_report(analysis)
        return vkt.WebResult(html=html)

    @vkt.DataView("Compliance Summary")
    def compliance_summary(self, params, **kwargs) -> vkt.DataResult:
        analysis = self._analysis(params)
        compliant_rows = analysis["compliant_rows"]
        requires_change_rows = analysis["requires_change"]

        compliant_group = (
            vkt.DataGroup(
                **{
                    f"compliant_{index}": vkt.DataItem(
                        row["revit_element_id"],
                        "Compliant",
                        status=vkt.DataStatus.SUCCESS,
                    )
                    for index, row in enumerate(compliant_rows)
                }
            )
            if compliant_rows
            else vkt.DataGroup(
                empty=vkt.DataItem("Compliant fittings", "None", status=vkt.DataStatus.WARNING)
            )
        )

        requires_change_group = (
            vkt.DataGroup(
                **{
                    f"requires_change_{index}": vkt.DataItem(
                        row["revit_element_id"],
                        row["family_name"],
                        status=vkt.DataStatus.ERROR,
                        subgroup=vkt.DataGroup(
                            current_family=vkt.DataItem("Current Family", row["family_name"]),
                            family_to_assign=vkt.DataItem("Family To Assign", REQUIRED_FAMILY_NAME),
                        ),
                    )
                    for index, row in enumerate(requires_change_rows)
                }
            )
            if requires_change_rows
            else vkt.DataGroup(
                empty=vkt.DataItem("Fittings requiring family change", "None", status=vkt.DataStatus.SUCCESS)
            )
        )

        data = vkt.DataGroup(
            compliant=vkt.DataItem(
                "Compliant Revit IDs",
                f"{len(compliant_rows)} fitting(s)",
                status=vkt.DataStatus.SUCCESS if compliant_rows else vkt.DataStatus.WARNING,
                subgroup=compliant_group,
            ),
            requires_change=vkt.DataItem(
                "Requires Family Change",
                f"{len(requires_change_rows)} fitting(s)",
                status=vkt.DataStatus.ERROR if requires_change_rows else vkt.DataStatus.SUCCESS,
                subgroup=requires_change_group,
            ),
        )
        return vkt.DataResult(data)

    def download_pdf(self, params, **kwargs) -> vkt.DownloadResult:
        analysis = self._analysis(params)
        pdf_bytes = build_pdf_report(analysis)
        today = date.today().strftime("%Y-%m-%d")
        return vkt.DownloadResult(
            file_content=pdf_bytes,
            file_name=f"fittings_check_report_{today}.pdf",
        )
