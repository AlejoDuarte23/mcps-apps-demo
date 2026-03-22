import logging
import math
from datetime import date

import plotly.graph_objects as go
from plotly.subplots import make_subplots
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

    @vkt.PlotlyView("Fittings Charts", duration_guess=1)
    def fittings_chart(self, params, **kwargs) -> vkt.PlotlyResult:
        analysis = self._analysis(params)
        if not analysis["rows"]:
            fig = go.Figure()
            fig.add_annotation(
                text="No fitting data entered yet.",
                xref="paper",
                yref="paper",
                x=0.5,
                y=0.5,
                showarrow=False,
                font=dict(size=16, color="#5B8FA3"),
            )
            fig.update_layout(
                xaxis=dict(visible=False),
                yaxis=dict(visible=False),
                plot_bgcolor="white",
            )
            return vkt.PlotlyResult(fig)

        rows = analysis["rows"]
        family_counter = analysis["family_counter"]

        fig = make_subplots(
            rows=1,
            cols=2,
            subplot_titles=("Pressure Drop per Revit ID", "Family Count Breakdown"),
            specs=[[{"type": "bar"}, {"type": "pie"}]],
            column_widths=[0.6, 0.4],
        )

        labels = [row["short_revit_id"] or "-" for row in rows]
        values = [0 if math.isnan(row["pressure_drop"]) else row["pressure_drop"] for row in rows]
        colors = ["#E88D7B" if row["pressure_issue"] else "#5B8FA3" for row in rows]

        fig.add_trace(
            go.Bar(
                x=labels,
                y=values,
                marker=dict(color=colors),
                name="Pressure Drop",
                hovertemplate="<b>Revit ID:</b> %{x}<br><b>Pressure Drop:</b> %{y:.2f} Pa<extra></extra>",
            ),
            row=1,
            col=1,
        )

        pie_labels = list(family_counter.keys())
        pie_values = list(family_counter.values())
        pie_colors = ["#5B8FA3", "#8FB8C9", "#A8C5D6", "#B8D4E3", "#E88D7B", "#9BAEC4"]

        fig.add_trace(
            go.Pie(
                labels=pie_labels,
                values=pie_values,
                marker=dict(colors=pie_colors[: len(pie_values)]),
                name="Family Count",
                hovertemplate="<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent}<extra></extra>",
            ),
            row=1,
            col=2,
        )

        fig.update_xaxes(
            title_text="Last 3 characters of Revit Element ID",
            row=1,
            col=1,
            gridcolor="#E0EEF5",
            title_font=dict(color="#2C3E50"),
        )
        fig.update_yaxes(
            title_text="Pressure Drop (Pa)",
            row=1,
            col=1,
            gridcolor="#E0EEF5",
            title_font=dict(color="#2C3E50"),
        )

        fig.update_layout(
            showlegend=False,
            plot_bgcolor="white",
            paper_bgcolor="white",
            font=dict(color="#2C3E50", size=11),
            title_font=dict(color="#2C5F7D", size=14),
            height=500,
        )

        return vkt.PlotlyResult(fig)

    @vkt.DataView("Compliance Summary")
    def qa_qc_fittings_check(self, params, **kwargs) -> vkt.DataResult:
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
