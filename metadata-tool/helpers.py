import base64
import io
from pathlib import Path

import matplotlib
import matplotlib.pyplot as plt
import plotly.graph_objects as go

matplotlib.use("Agg")


def normalize_family_name(family_name: str) -> str:
    family_name = (family_name or "").strip()
    if family_name.startswith("System Family: "):
        return family_name.replace("System Family: ", "", 1).strip()
    return family_name


def get_lookup_key(element: dict) -> str:
    return f"{normalize_family_name(element['familyName'])}|{element['typeName']}"


def build_fallback_metadata(element: dict) -> dict:
    family_name = normalize_family_name(element["familyName"])
    type_name = element["typeName"] or "Standard"

    return {
        "Manufacturer": "Demo Company",
        "Model": f"{family_name.upper().replace(' ', '-')[:24]}-{type_name.upper().replace(' ', '-')[:16]}",
        "Keynote": "23 00 00",
        "Description": f"Fallback metadata for {family_name} {type_name}",
        "Assembly Code": "23.00.00.00",
        "Type Mark": f"FB-{type_name.upper().replace(' ', '-')[:20]}",
        "Cost": "50.00",
    }


def enrich_element(element: dict, metadata_database: dict) -> dict:
    key = get_lookup_key(element)
    db_entry = metadata_database.get(key)
    enriched = dict(element)
    if db_entry:
        enriched.update(db_entry["metadataToApply"])
        enriched["familyName"] = normalize_family_name(enriched["familyName"])
        enriched["status"] = "Metadata to add"
        enriched["_from_database"] = True
    else:
        enriched["familyName"] = normalize_family_name(enriched["familyName"])
        enriched.update(build_fallback_metadata(enriched))
        enriched["status"] = "Metadata to add"
        enriched["_from_database"] = False
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
    ax.set_facecolor("#E6F3FF")

    max_val = max(sorted_counts) if sorted_counts else 1
    bar_colors = ["#1E90FF" if value == max_val else "#4DA6FF" if value >= max_val * 0.7 else "#70B8FF" if value >= max_val * 0.4 else "#94C9FF" for value in sorted_counts]

    bars = ax.barh(sorted_labels, sorted_counts, height=bar_height, color=bar_colors, edgecolor="#ffffff", linewidth=0.8)
    ax.set_axisbelow(True)
    ax.xaxis.grid(True, color="#C9E5FF", linestyle="--", linewidth=0.7)
    ax.set_xlabel("Number of Elements", fontsize=11, fontfamily="sans-serif", labelpad=8, color="#1E90FF")
    ax.set_title(title, fontsize=13, fontweight="bold", fontfamily="sans-serif", pad=14, color="#1E90FF")
    ax.tick_params(axis="y", labelsize=10, colors="#2C3E50", length=0)
    ax.tick_params(axis="x", labelsize=9, colors="#1E90FF")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.spines["bottom"].set_color("#A8D0FF")

    for bar, value in zip(bars, sorted_counts):
        ax.text(bar.get_width() + max_val * 0.02, bar.get_y() + bar.get_height() / 2, str(value), va="center", ha="left", fontsize=10, fontweight="bold", color="#1E90FF")

    ax.set_xlim(0, max_val + max_val * 0.25 + 0.5)
    fig.tight_layout(pad=1.5)
    return fig_to_base64(fig)


def build_pie_chart(labels: list[str], counts: list[int], title: str) -> str:
    short_labels = [label if len(label) <= 25 else label[:22] + "..." for label in labels]
    colors = ["#1E90FF", "#4DA6FF", "#70B8FF", "#94C9FF", "#E88D7B", "#B8DAFF", "#A8D0FF"]
    fig, ax = plt.subplots(figsize=(7, 5))
    fig.patch.set_facecolor("#ffffff")
    ax.pie(counts, labels=short_labels, autopct="%1.1f%%", colors=colors[:len(labels)], startangle=140, textprops={"fontsize": 9, "fontfamily": "sans-serif", "color": "#2C3E50"})
    ax.set_title(title, fontsize=13, fontweight="bold", fontfamily="sans-serif", color="#1E90FF")
    fig.tight_layout()
    return fig_to_base64(fig)


def build_plotly_pie_chart(labels: list[str], counts: list[int], title: str) -> str:
    colors = ["#1E90FF", "#4DA6FF", "#70B8FF", "#94C9FF", "#E88D7B", "#B8DAFF", "#A8D0FF"]

    fig = go.Figure(
        data=[
            go.Pie(
                labels=labels,
                values=counts,
                marker=dict(colors=colors[:len(labels)], line=dict(color="white", width=2)),
                hovertemplate="<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent}<extra></extra>",
                textfont=dict(size=12, color="white"),
                textposition="inside",
            )
        ]
    )

    fig.update_layout(
        title=dict(
            text=title,
            font=dict(size=16, color="#1E90FF", family="Arial"),
            x=0.5,
            xanchor="center",
        ),
        showlegend=True,
        legend=dict(
            font=dict(size=11, color="#2C3E50", family="Arial"),
            orientation="h",
            yanchor="top",
            y=-0.1,
            xanchor="center",
            x=0.5,
        ),
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(l=20, r=20, t=60, b=80),
        height=450,
    )

    return fig.to_html(include_plotlyjs="cdn", div_id=f"pie_chart_{title.replace(' ', '_')}", config={"displayModeBar": True, "displaylogo": False})


def build_html_report(
    enriched: list[dict],
    unique_types: list[dict],
    metadata_found_count: int,
    chart2_plotly: str,
    chart4_plotly: str,
) -> str:
    not_in_db_count = len(enriched) - metadata_found_count
    summary_rows = [
        ("Total Elements", len(enriched)),
        ("Unique Types", len(unique_types)),
        ("Types with Metadata Found", metadata_found_count),
    ]
    summary_html = "".join(f"<tr><td>{label}</td><td><strong>{value}</strong></td></tr>" for label, value in summary_rows)

    instance_rows_html = ""
    for element in enriched:
        row_style = ""
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
        row_style = ' style="background-color: #FFE5E0;"' if unique_type["status"] == "Not in Database" else ""
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
    html = html.replace("__CHART2_PLOTLY__", chart2_plotly)
    html = html.replace("__CHART4_PLOTLY__", chart4_plotly)
    html = html.replace("__UNIQUE_ROWS_HTML__", unique_rows_html)
    return html
