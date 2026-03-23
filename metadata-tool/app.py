import base64
import logging
from collections import Counter
from io import BytesIO

import viktor as vkt

from helpers import build_bar_chart, build_html_report, build_pie_chart, build_plotly_pie_chart, enrich_element, get_unique_types

logger = logging.getLogger("viktor")



METADATA_DATABASE = {
    "Flex Duct Round|Flex - Round": {
        "familyName": "Flex Duct Round", "typeName": "Flex - Round", "category": "OST_FlexDuctCurves",
        "metadataToApply": {"Manufacturer": "DemoFlex", "Model": "FLEX-RND-STD", "Keynote": "23 31 00", "Description": "Round flexible duct for supply air routing", "Assembly Code": "23.31.13.00", "Type Mark": "FD-RND-STD", "Cost": "18.50"},
    },
    "M_Rectangular to Round Duct Transition - Angle|45 Degree": {
        "familyName": "M_Rectangular to Round Duct Transition - Angle", "typeName": "45 Degree", "category": "OST_DuctFitting",
        "metadataToApply": {"Manufacturer": "DemoFab", "Model": "RTR-45", "Keynote": "23 33 00", "Description": "Rectangular to round duct transition with 45 degree angle", "Assembly Code": "23.33.13.00", "Type Mark": "DF-RTR-45", "Cost": "78.00"},
    },
    "Rectangular Duct|Mitered Elbows / Taps": {
        "familyName": "Rectangular Duct", "typeName": "Mitered Elbows / Taps", "category": "OST_DuctCurves",
        "metadataToApply": {"Manufacturer": "DemoSheetMetal", "Model": "RECT-MITER-TAP", "Keynote": "23 31 13", "Description": "Rectangular duct system type using mitered elbows and taps", "Assembly Code": "23.31.13.13", "Type Mark": "RD-MIT-TAP", "Cost": "42.00"},
    },
    "Rectangular Duct|Mitered Elbows / Tees": {
        "familyName": "Rectangular Duct", "typeName": "Mitered Elbows / Tees", "category": "OST_DuctCurves",
        "metadataToApply": {"Manufacturer": "DemoSheetMetal", "Model": "RECT-MITER-TEE", "Keynote": "23 31 13", "Description": "Rectangular duct system type using mitered elbows and tees", "Assembly Code": "23.31.13.13", "Type Mark": "RD-MIT-TEE", "Cost": "44.00"},
    },
    "M_Rectangular Duct Transition - Angle|45 Degree": {
        "familyName": "M_Rectangular Duct Transition - Angle", "typeName": "45 Degree", "category": "OST_DuctFitting",
        "metadataToApply": {"Manufacturer": "DemoFab", "Model": "RDTA-45", "Keynote": "23 33 00", "Description": "Rectangular duct transition angle fitting", "Assembly Code": "23.33.13.00", "Type Mark": "DF-RECT-TRANS-45", "Cost": "72.00"},
    },
    "M_Rectangular Duct Takeoff|Standard": {
        "familyName": "M_Rectangular Duct Takeoff", "typeName": "Standard", "category": "OST_DuctFitting",
        "metadataToApply": {"Manufacturer": "DemoFab", "Model": "RDTO-STD", "Keynote": "23 33 00", "Description": "Rectangular duct takeoff fitting", "Assembly Code": "23.33.13.00", "Type Mark": "DF-TAKEOFF-STD", "Cost": "61.00"},
    },
    "M_Rectangular Duct Tee|Standard": {
        "familyName": "M_Rectangular Duct Tee", "typeName": "Standard", "category": "OST_DuctFitting",
        "metadataToApply": {"Manufacturer": "DemoFab", "Model": "RDT-STD", "Keynote": "23 33 00", "Description": "Standard rectangular duct tee fitting", "Assembly Code": "23.33.13.00", "Type Mark": "DF-TEE-RECT-STD", "Cost": "65.00"},
    },
    "M_Supply Diffuser - Rectangular Face Round Neck|600x600 - 250 Neck": {
        "familyName": "M_Supply Diffuser - Rectangular Face Round Neck", "typeName": "600x600 - 250 Neck", "category": "OST_DuctTerminal",
        "metadataToApply": {"Manufacturer": "DemoAir", "Model": "SQR-600-250", "Keynote": "23 37 13", "Description": "Rectangular face supply diffuser 600x600 with 250 round neck", "Assembly Code": "23.37.13.00", "Type Mark": "AT-SQR-600-250", "Cost": "185.00"},
    },
}


class Parametrization(vkt.Parametrization):
    title = vkt.Text("""# QA/QC Elements Metadata App
**Automated metadata lookup for Revit MEP elements.** Enter elements below (System Name, Element ID, Category, Family & Type Name) to match against the internal database — enriching each with **Manufacturer, Model, Keynote, Assembly Code, Type Mark and Cost**. Elements not found are highlighted in **amber**.
""")

    elements = vkt.Table(
        "Revit Elements",
        default=[
            {"system_name": "Mechanical Supply Air 17", "revit_element_id": "120001", "category": "OST_FlexDuctCurves",  "family_name": "Flex Duct Round",                                  "type_name": "Flex - Round"},
            {"system_name": "Mechanical Supply Air 17", "revit_element_id": "120006", "category": "OST_FlexDuctCurves",  "family_name": "Flex Duct Round",                                  "type_name": "Flex - Round"},
            {"system_name": "Mechanical Supply Air 17", "revit_element_id": "120002", "category": "OST_DuctFitting",     "family_name": "M_Rectangular to Round Duct Transition - Angle",    "type_name": "45 Degree"},
            {"system_name": "Mechanical Supply Air 17", "revit_element_id": "120007", "category": "OST_DuctFitting",     "family_name": "M_Rectangular to Round Duct Transition - Angle",    "type_name": "45 Degree"},
            {"system_name": "Mechanical Supply Air 17", "revit_element_id": "120003", "category": "OST_DuctCurves",      "family_name": "Rectangular Duct",                                  "type_name": "Mitered Elbows / Taps"},
            {"system_name": "Mechanical Supply Air 17", "revit_element_id": "120008", "category": "OST_DuctCurves",      "family_name": "Rectangular Duct",                                  "type_name": "Mitered Elbows / Taps"},
            {"system_name": "Mechanical Supply Air 17", "revit_element_id": "120004", "category": "OST_DuctFitting",     "family_name": "M_Rectangular Duct Tee",                            "type_name": "Standard"},
            {"system_name": "Mechanical Supply Air 17", "revit_element_id": "120009", "category": "OST_DuctFitting",     "family_name": "M_Rectangular Duct Tee",                            "type_name": "Standard"},
            {"system_name": "Mechanical Supply Air 17", "revit_element_id": "120010", "category": "OST_DuctTerminal",    "family_name": "M_Supply Diffuser - Rectangular Face Round Neck",   "type_name": "600x600 - 250 Neck"},
        ],
    )
    elements.system_name       = vkt.TextField("System Name")
    elements.revit_element_id  = vkt.TextField("Revit Element ID")
    elements.category          = vkt.TextField("Category")
    elements.family_name       = vkt.TextField("Family Name")
    elements.type_name         = vkt.TextField("Type Name")

    download_title = vkt.Text("""# Download Report
**Export the full Metadata Report as a PDF** — includes a **project summary**, **all element instances** with enriched metadata, **analytical charts**, and a **unique types summary**.
""")
    download_report = vkt.DownloadButton("Download PDF Report", method="download_pdf_report")



class Controller(vkt.Controller):
    parametrization = Parametrization

    def get_input_elements(self, params) -> list[dict]:
        """Convert DynamicArray rows from params into a list of element dicts."""
        elements = []
        for row in params.elements:
            elements.append({
                "systemName":      row.system_name      or "",
                "revitElementId":  row.revit_element_id or "",
                "category":        row.category         or "",
                "familyName":      row.family_name      or "",
                "typeName":        row.type_name        or "",
            })
        return elements

    @vkt.WebView("Engineering Report")
    def report_view(self, params, **kwargs):
        """Generate a full HTML engineering report with tables and charts."""
        enriched = [enrich_element(el, METADATA_DATABASE) for el in self.get_input_elements(params)]
        unique_types = get_unique_types(enriched)
        compliant_count = len([el for el in enriched if el["status"] == "Compliant"])

        metadata_found_count = len([el for el in enriched if el["status"] == "Metadata Found"])
        logger.info(f"WebView: total={len(enriched)}, unique={len(unique_types)}, metadata_found={metadata_found_count}")

        # --- Chart 2: Plotly Pie chart – distribution by type name ---
        type_counter = Counter(el["typeName"] for el in enriched)
        type_labels = list(type_counter.keys())
        type_counts = list(type_counter.values())
        chart2_plotly = build_plotly_pie_chart(type_labels, type_counts, "Element Distribution by Type")

        # --- Chart 4: Plotly Pie chart – distribution by family name ---
        family_counter = Counter(el["familyName"] for el in enriched)
        family_labels = list(family_counter.keys())
        family_counts = list(family_counter.values())
        chart4_plotly = build_plotly_pie_chart(family_labels, family_counts, "Element Distribution by Family Name")

        # --- Build HTML ---
        html = build_html_report(
            enriched=enriched,
            unique_types=unique_types,
            metadata_found_count=metadata_found_count,
            chart2_plotly=chart2_plotly,
            chart4_plotly=chart4_plotly,
        )
        return vkt.WebResult(html=html)

    @vkt.DataView("Metadata to Assign Summary")
    def qa_qc_metadata_to_assign(self, params, **kwargs):
        enriched = [enrich_element(el, METADATA_DATABASE) for el in self.get_input_elements(params)]
        unique_types = get_unique_types(enriched)

        logger.info(f"UniqueTypesView: {len(unique_types)} unique types from {len(enriched)} elements")

        # One top-level DataItem per unique type, labelled by typeName only
        data = vkt.DataGroup()
        for i, ut in enumerate(unique_types):
            status = ut["status"]
            ds = vkt.DataStatus.SUCCESS if status == "Metadata Found" else vkt.DataStatus.WARNING

            type_group = vkt.DataGroup(
                family_name=vkt.DataItem("Family Name", ut["familyName"]),
                type_name=vkt.DataItem("Type Name", ut["typeName"]),
                category=vkt.DataItem("Category", ut["category"]),
                manufacturer=vkt.DataItem("Manufacturer", ut.get("Manufacturer", "N/A")),
                model=vkt.DataItem("Model", ut.get("Model", "N/A")),
                keynote=vkt.DataItem("Keynote", ut.get("Keynote", "N/A")),
                assembly_code=vkt.DataItem("Assembly Code", ut.get("Assembly Code", "N/A")),
                type_mark=vkt.DataItem("Type Mark", ut.get("Type Mark", "N/A")),
                cost=vkt.DataItem("Cost", ut.get("Cost", "N/A"), prefix="$"),
                status_item=vkt.DataItem("Status", status, status=ds),
            )
            data.add(vkt.DataItem(ut["typeName"], "", subgroup=type_group))

        return vkt.DataResult(data)

    def download_pdf_report(self, params, **kwargs):
        """Generate and download the engineering report as a PDF file."""
        enriched = [enrich_element(el, METADATA_DATABASE) for el in self.get_input_elements(params)]
        unique_types = get_unique_types(enriched)

        metadata_found_count = len([el for el in enriched if el["status"] == "Metadata Found"])
        logger.info(f"📄 PDF Download: total={len(enriched)}, unique={len(unique_types)}, metadata_found={metadata_found_count}")

        # Build charts
        type_counter = Counter(el["typeName"] for el in enriched)
        type_labels = list(type_counter.keys())
        type_counts = list(type_counter.values())
        chart1_b64 = build_bar_chart(type_labels, type_counts, "Element Count by Type", "Type Name")
        chart2_b64 = build_pie_chart(type_labels, type_counts, "Element Distribution by Type")

        family_counter = Counter(el["familyName"] for el in enriched)
        family_labels = list(family_counter.keys())
        family_counts = list(family_counter.values())
        chart3_b64 = build_bar_chart(family_labels, family_counts, "Element Count by Family Name", "Family Name")
        chart4_b64 = build_pie_chart(family_labels, family_counts, "Element Distribution by Family Name")

        # Convert to PDF using reportlab
        from io import BytesIO
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import mm
        from reportlab.platypus import (
            Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
        )

        pdf_buffer = BytesIO()
        doc = SimpleDocTemplate(
            pdf_buffer,
            pagesize=landscape(A4),
            leftMargin=15 * mm, rightMargin=15 * mm,
            topMargin=15 * mm, bottomMargin=15 * mm,
        )
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle("title", parent=styles["Heading1"], fontSize=16, spaceAfter=6, textColor=colors.HexColor("#1E90FF"))
        h2_style = ParagraphStyle("h2", parent=styles["Heading2"], fontSize=12, spaceBefore=14, spaceAfter=4, textColor=colors.HexColor("#1E90FF"))
        small_style = ParagraphStyle("small", parent=styles["Normal"], fontSize=8, textColor=colors.HexColor("#2C3E50"))
        body_style = ParagraphStyle("body", parent=styles["Normal"], fontSize=9, spaceAfter=6, textColor=colors.HexColor("#2C3E50"))

        story = []

        # --- Title ---
        story.append(Paragraph("Revit Elements Metadata Report", title_style))
        story.append(Paragraph("Revit MEP Elements — Automated Metadata Lookup from Type Database", body_style))
        story.append(Spacer(1, 3 * mm))
        story.append(Paragraph(
            "This report provides a comprehensive analysis of Revit MEP element instances and their associated "
            "metadata properties. The tool automatically enriches element data by matching family and type names "
            "against a centralized metadata database containing manufacturer information, model numbers, keynotes, "
            "assembly codes, and cost data.",
            body_style
        ))
        story.append(Paragraph(
            "Elements that cannot be matched to the database are flagged for manual review and metadata assignment. "
            "This automated approach ensures consistent metadata standards across the project and facilitates BIM "
            "coordination, quantity takeoffs, and construction documentation workflows.",
            body_style
        ))
        story.append(Spacer(1, 6 * mm))

        # --- Section 1: Project Summary ---
        story.append(Paragraph("1. Project Summary", h2_style))
        story.append(Paragraph(
            "The following metrics summarize the metadata enrichment results for all analyzed Revit elements. "
            "Elements are matched against the database using the combination of family name and type name as a unique identifier.",
            body_style
        ))
        story.append(Spacer(1, 2 * mm))
        not_in_db_count = len(enriched) - metadata_found_count
        summary_data = [
            ["Metric", "Value"],
            ["Total Elements", str(len(enriched))],
            ["Unique Types", str(len(unique_types))],
            ["Types with Metadata Found", str(metadata_found_count)],
            ["Types Not in Database", str(not_in_db_count)],
        ]
        summary_table = Table(summary_data, colWidths=[80 * mm, 40 * mm])
        summary_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E90FF")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#E6F3FF")]),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#A8D0FF")),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(summary_table)
        story.append(Spacer(1, 2 * mm))

        coverage_rate = (metadata_found_count / len(enriched) * 100) if len(enriched) > 0 else 0
        summary_text = (
            f"<b>Database Coverage Rate:</b> {coverage_rate:.1f}% ({metadata_found_count} of {len(enriched)} elements matched). "
        )
        if not_in_db_count > 0:
            summary_text += (
                f"<b>Action Required:</b> {not_in_db_count} element(s) are not in the database and require "
                "manual metadata assignment or database updates."
            )
        else:
            summary_text += "<b>Status:</b> All elements successfully matched to metadata database."

        story.append(Paragraph(summary_text, body_style))
        story.append(Spacer(1, 6 * mm))

        # --- Section 2: Element Instance Metadata Table ---
        story.append(Paragraph("2. Element Instance Metadata", h2_style))
        story.append(Paragraph(
            "Complete listing of all analyzed Revit element instances with enriched metadata properties. "
            "Rows highlighted in coral indicate elements that could not be matched to the database and are missing metadata. "
            "This table includes system assignment, category classification, and all available metadata fields.",
            body_style
        ))
        story.append(Spacer(1, 2 * mm))
        inst_headers = ["Element ID", "System", "Category", "Family Name", "Type Name",
                        "Manufacturer", "Model", "Keynote", "Assembly Code", "Type Mark", "Cost", "Status"]
        inst_rows = [inst_headers]
        for el in enriched:
            inst_rows.append([
                el["revitElementId"], el["systemName"], el["category"],
                Paragraph(el["familyName"], small_style),
                el["typeName"],
                el.get("Manufacturer", "N/A"), el.get("Model", "N/A"),
                el.get("Keynote", "N/A"), el.get("Assembly Code", "N/A"),
                el.get("Type Mark", "N/A"), el.get("Cost", "N/A"),
                el["status"],
            ])
        page_w = landscape(A4)[0] - 30 * mm
        col_widths = [22, 38, 30, 60, 38, 28, 22, 20, 28, 22, 16, 30]
        col_widths = [w * mm * page_w / sum(w * mm for w in col_widths) for w in col_widths]
        inst_table = Table(inst_rows, colWidths=col_widths, repeatRows=1)
        inst_style = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E90FF")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 7),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#E6F3FF")]),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#A8D0FF")),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]
        # Highlight "Not in Database" rows in soft coral
        for row_idx, el in enumerate(enriched, start=1):
            if el["status"] == "Not in Database":
                inst_style.append(("BACKGROUND", (0, row_idx), (-1, row_idx), colors.HexColor("#FFE5E0")))
        inst_table.setStyle(TableStyle(inst_style))
        story.append(inst_table)
        story.append(Spacer(1, 6 * mm))

        # --- Section 3: Charts (2x2 grid using a table of images) ---
        story.append(Paragraph("3. Visual Analysis", h2_style))
        story.append(Paragraph(
            "The following charts provide visual representations of element distribution and composition. "
            "These visualizations help identify the most common element types and families in the project, "
            "supporting quality control and coordination efforts.",
            body_style
        ))
        story.append(Spacer(1, 2 * mm))
        chart_w = (page_w / 2) - 5 * mm
        chart_h = chart_w * 0.55

        def b64_to_image(b64_str: str, w, h) -> Image:
            img_buf = BytesIO(base64.b64decode(b64_str))
            return Image(img_buf, width=w, height=h)

        charts_table = Table(
            [
                [b64_to_image(chart1_b64, chart_w, chart_h), b64_to_image(chart2_b64, chart_w, chart_h)],
                [b64_to_image(chart3_b64, chart_w, chart_h), b64_to_image(chart4_b64, chart_w, chart_h)],
            ],
            colWidths=[chart_w + 5 * mm, chart_w + 5 * mm],
        )
        charts_table.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 1, colors.HexColor("#A8D0FF")),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(charts_table)
        story.append(Spacer(1, 6 * mm))

        # --- Section 4: Unique Types Metadata Summary ---
        story.append(Paragraph("4. Unique Types Metadata Summary", h2_style))
        story.append(Paragraph(
            "This section consolidates metadata for all unique family-type combinations found in the project. "
            "Each unique type represents a distinct element configuration that may appear multiple times in the model. "
            "Rows highlighted in coral indicate types that require database entries or manual metadata assignment.",
            body_style
        ))
        story.append(Spacer(1, 2 * mm))
        ut_headers = ["Family Name", "Type Name", "Category", "Manufacturer",
                      "Model", "Keynote", "Assembly Code", "Type Mark", "Cost", "Status"]
        ut_rows = [ut_headers]
        for ut in unique_types:
            ut_rows.append([
                Paragraph(ut["familyName"], small_style),
                ut["typeName"], ut["category"],
                ut.get("Manufacturer", "N/A"), ut.get("Model", "N/A"),
                ut.get("Keynote", "N/A"), ut.get("Assembly Code", "N/A"),
                ut.get("Type Mark", "N/A"), ut.get("Cost", "N/A"),
                ut["status"],
            ])
        ut_col_widths_raw = [60, 38, 30, 28, 22, 20, 28, 22, 16, 30]
        ut_col_widths = [w * mm * page_w / sum(w * mm for w in ut_col_widths_raw) for w in ut_col_widths_raw]
        ut_table = Table(ut_rows, colWidths=ut_col_widths, repeatRows=1)
        ut_style = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E90FF")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 7),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#E6F3FF")]),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#A8D0FF")),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]
        for row_idx, ut in enumerate(unique_types, start=1):
            if ut["status"] == "Not in Database":
                ut_style.append(("BACKGROUND", (0, row_idx), (-1, row_idx), colors.HexColor("#FFE5E0")))
        ut_table.setStyle(TableStyle(ut_style))
        story.append(ut_table)
        story.append(Spacer(1, 6 * mm))

        # --- Section 5: Recommendations ---
        story.append(Paragraph("5. Recommendations and Next Steps", h2_style))
        story.append(Spacer(1, 2 * mm))

        recommendations = []
        if not_in_db_count > 0:
            recommendations.append(
                f"<b>Database Updates:</b> Add {not_in_db_count} missing family-type combination(s) to the metadata "
                "database with complete manufacturer, model, keynote, assembly code, type mark, and cost information."
            )
            recommendations.append(
                "<b>Manual Assignment:</b> For elements that cannot be added to the database, manually assign metadata "
                "properties directly in Revit using the Type Properties dialog."
            )

        if metadata_found_count > 0:
            recommendations.append(
                f"<b>Metadata Verification:</b> Review the {metadata_found_count} element(s) with auto-assigned metadata "
                "to ensure accuracy and completeness before proceeding with construction documentation."
            )

        if not_in_db_count == 0:
            recommendations.append(
                "<b>Success:</b> All element types are present in the database with complete metadata. "
                "No corrective action required at this time."
            )
        else:
            recommendations.append(
                "<b>Re-validation:</b> After updating the database or manually assigning metadata in Revit, "
                "regenerate this report to verify that all elements have complete metadata."
            )

        recommendations.append(
            "<b>Data Governance:</b> Maintain the metadata database as a centralized resource for consistent "
            "BIM data standards across all projects. Establish workflows for adding new families and types."
        )

        for rec in recommendations:
            story.append(Paragraph(rec, body_style))
            story.append(Spacer(1, 2 * mm))

        story.append(Spacer(1, 4 * mm))
        story.append(Paragraph(
            "<b>Note:</b> This report was automatically generated by the Revit Elements Metadata Tool. "
            "For questions about metadata standards or database updates, contact your BIM manager or project coordinator.",
            body_style
        ))

        # Build PDF
        doc.build(story)
        pdf_buffer.seek(0)
        logger.info("✅ PDF generated successfully via reportlab")

        return vkt.DownloadResult(pdf_buffer, file_name="revit_metadata_report.pdf")

    @vkt.TableView("Metadata Database Registry")
    def data_view(self, params, **kwargs):
        """Display all entries from METADATA_DATABASE as a full registry/catalogue table."""
        logger.info(f"📋 Metadata Database Registry: showing {len(METADATA_DATABASE)} registered types")

        # Build one row per entry in the metadata database
        rows = []
        for key, entry in METADATA_DATABASE.items():
            meta = entry["metadataToApply"]
            rows.append([
                entry["familyName"],
                entry["typeName"],
                entry["category"],
                meta.get("Manufacturer", "N/A"),
                meta.get("Model", "N/A"),
                meta.get("Keynote", "N/A"),
                meta.get("Description", "N/A"),
                meta.get("Assembly Code", "N/A"),
                meta.get("Type Mark", "N/A"),
                meta.get("Cost", "N/A"),
            ])

        headers = [
            vkt.TableHeader("Family Name", align="left"),
            vkt.TableHeader("Type Name", align="left"),
            vkt.TableHeader("Category", align="left"),
            vkt.TableHeader("Manufacturer", align="left"),
            vkt.TableHeader("Model", align="left"),
            vkt.TableHeader("Keynote", align="left"),
            vkt.TableHeader("Description", align="left"),
            vkt.TableHeader("Assembly Code", align="left"),
            vkt.TableHeader("Type Mark", align="left"),
            vkt.TableHeader("Cost", align="right"),
        ]

        return vkt.TableResult(rows, column_headers=headers)
