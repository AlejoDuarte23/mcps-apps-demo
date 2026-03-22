import base64
import io
from pathlib import Path

import matplotlib
import matplotlib.pyplot as plt

matplotlib.use("Agg")


def get_lookup_key(element: dict) -> str:
    return f"{element['familyName']}|{element['typeName']}"


def enrich_element(element: dict, metadata_database: dict) -> dict:
    key = get_lookup_key(element)
    db_entry = metadata_database.get(key)
    enriched = dict(element)
    if db_entry:
        enriched.update(db_entry["metadataToApply"])
        enriched["status"] = "Metadata Found"
    else:
        for field in ("Manufacturer", "Model", "Keynote", "Description", "Assembly Code", "Type Mark", "Cost"):
            enriched[field] = "N/A"
        enriched["status"] = "Not in Database"
    return enriched


def get_unique_types(enriched_elements: list[dict]) -> list[dict]:
    seen: set[str] = set()
    unique: list[dict] = []
    for element in enriched_elements:
        key = get_lookup_key(element)
        if key not in seen:
            seen.add(key)
            unique.append(element)
    return unique


def fig_to_base64(fig: plt.Figure) -> str:
    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", bbox_inches="tight", dpi=120)
    buffer.seek(0)
    encoded = base64.b64encode(buffer.read()).decode("utf-8")
    plt.close(fig)
    return encoded


def build_bar_chart(labels: list[str], counts: list[int], title: str, xlabel: str) -> str:
    short_labels = [label if len(label) <= 35 else label[:32] + "..." for label in labels]
    paired = sorted(zip(counts, short_labels), key=lambda item: item[0])
    sorted_counts, sorted_labels = zip(*paired) if paired else ([], [])

    bar_height = 0.55
    fig_height = max(4, len(sorted_labels) * 0.85 + 1.5)
    fig, ax = plt.subplots(figsize=(10, fig_height))
    fig.patch.set_facecolor("#ffffff")
    ax.set_facecolor("#f8f9fa")

    max_val = max(sorted_counts) if sorted_counts else 1
    bar_colors = [plt.cm.Blues(0.45 + 0.45 * (value / max_val)) for value in sorted_counts]

    bars = ax.barh(sorted_labels, sorted_counts, height=bar_height, color=bar_colors, edgecolor="#ffffff", linewidth=0.8)
    ax.set_axisbelow(True)
    ax.xaxis.grid(True, color="#dddddd", linestyle="--", linewidth=0.7)
    ax.set_xlabel("Number of Elements", fontsize=11, fontfamily="sans-serif", labelpad=8, color="#333333")
    ax.set_title(title, fontsize=13, fontweight="bold", fontfamily="sans-serif", pad=14, color="#111111")
    ax.tick_params(axis="y", labelsize=10, colors="#333333", length=0)
    ax.tick_params(axis="x", labelsize=9, colors="#555555")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.spines["bottom"].set_color("#cccccc")

    for bar, value in zip(bars, sorted_counts):
        ax.text(bar.get_width() + max_val * 0.02, bar.get_y() + bar.get_height() / 2, str(value), va="center", ha="left", fontsize=10, fontweight="bold", color="#222222")

    ax.set_xlim(0, max_val + max_val * 0.25 + 0.5)
    fig.tight_layout(pad=1.5)
    return fig_to_base64(fig)


def build_pie_chart(labels: list[str], counts: list[int], title: str) -> str:
    short_labels = [label if len(label) <= 25 else label[:22] + "..." for label in labels]
    colors = ["#4472C4", "#ED7D31", "#A9D18E", "#FF0000", "#FFC000", "#5B9BD5", "#70AD47"]
    fig, ax = plt.subplots(figsize=(7, 5))
    ax.pie(counts, labels=short_labels, autopct="%1.1f%%", colors=colors[:len(labels)], startangle=140, textprops={"fontsize": 9, "fontfamily": "sans-serif"})
    ax.set_title(title, fontsize=13, fontweight="bold", fontfamily="sans-serif")
    fig.tight_layout()
    return fig_to_base64(fig)


def build_html_report(
    enriched: list[dict],
    unique_types: list[dict],
    metadata_found_count: int,
    chart1_b64: str,
    chart2_b64: str,
    chart3_b64: str,
    chart4_b64: str,
) -> str:
    not_in_db_count = len(enriched) - metadata_found_count
    summary_rows = [
        ("Total Elements", len(enriched)),
        ("Unique Types", len(unique_types)),
        ("Types with Metadata Found", metadata_found_count),
        ("Types Not in Database", not_in_db_count),
    ]
    summary_html = "".join(f"<tr><td>{label}</td><td><strong>{value}</strong></td></tr>" for label, value in summary_rows)

    instance_rows_html = ""
    for element in enriched:
        row_style = ' style="background-color: #fff3cd;"' if element["status"] == "Not in Database" else ""
        instance_rows_html += (
            f"<tr{row_style}>"
            f"<td>{element['revitElementId']}</td>"
            f"<td>{element['systemName']}</td>"
            f"<td>{element['category']}</td>"
            f"<td>{element['familyName']}</td>"
            f"<td>{element['typeName']}</td>"
            f"<td>{element.get('Manufacturer', 'N/A')}</td>"
            f"<td>{element.get('Model', 'N/A')}</td>"
            f"<td>{element.get('Keynote', 'N/A')}</td>"
            f"<td>{element.get('Description', 'N/A')}</td>"
            f"<td>{element.get('Assembly Code', 'N/A')}</td>"
            f"<td>{element.get('Type Mark', 'N/A')}</td>"
            f"<td>{element.get('Cost', 'N/A')}</td>"
            f"<td>{element['status']}</td>"
            f"</tr>"
        )

    unique_rows_html = ""
    for unique_type in unique_types:
        row_style = ' style="background-color: #fff3cd;"' if unique_type["status"] == "Not in Database" else ""
        unique_rows_html += (
            f"<tr{row_style}>"
            f"<td>{unique_type['familyName']}</td>"
            f"<td>{unique_type['typeName']}</td>"
            f"<td>{unique_type['category']}</td>"
            f"<td>{unique_type.get('Manufacturer', 'N/A')}</td>"
            f"<td>{unique_type.get('Model', 'N/A')}</td>"
            f"<td>{unique_type.get('Keynote', 'N/A')}</td>"
            f"<td>{unique_type.get('Description', 'N/A')}</td>"
            f"<td>{unique_type.get('Assembly Code', 'N/A')}</td>"
            f"<td>{unique_type.get('Type Mark', 'N/A')}</td>"
            f"<td>{unique_type.get('Cost', 'N/A')}</td>"
            f"<td>{unique_type['status']}</td>"
            f"</tr>"
        )

    html = (Path(__file__).parent / "report_template.html").read_text(encoding="utf-8")
    html = html.replace("__SUMMARY_HTML__", summary_html)
    html = html.replace("__INSTANCE_ROWS_HTML__", instance_rows_html)
    html = html.replace("__CHART1_B64__", chart1_b64)
    html = html.replace("__CHART2_B64__", chart2_b64)
    html = html.replace("__CHART3_B64__", chart3_b64)
    html = html.replace("__CHART4_B64__", chart4_b64)
    html = html.replace("__UNIQUE_ROWS_HTML__", unique_rows_html)
    return html
