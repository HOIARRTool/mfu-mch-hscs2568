
import re
import textwrap
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import openpyxl


st.set_page_config(
    page_title="HSCS Dashboard",
    page_icon="📊",
    layout="wide"
)

# =========================================================
# Shared thresholds / colors
# =========================================================
# Quadrant colors
Q_RED_BG = "#8A1538"       # แดงเลือดนก
Q_ORANGE_BG = "#EF6C00"    # ส้มแก่
Q_YELLOW_BG = "#F3E58A"    # เหลืองนวลตา
Q_GREEN_BG = "#2E7D32"     # เขียวเข้ม

# Heatmap colors
H_RED_BG = "#FF2B2B"       # แดงสด
H_ORANGE_BG = "#EF6C00"
H_YELLOW_BG = "#F3E58A"
H_GREEN_BG = "#2E7D32"


def classify_score_quadrant(score: float) -> tuple[str, str]:
    if score < 60:
        return "ควรพัฒนาด่วน", "แดง"
    elif 60 <= score <= 70:
        return "เร่งพัฒนา", "ส้ม"
    elif 70 < score <= 80:
        return "ควรพัฒนาต่อเนื่อง", "เหลือง"
    else:
        return "ควรส่งเสริม", "เขียว"


def heatmap_bg_color(score) -> str:
    if pd.isna(score):
        return "#FFFFFF"
    score = float(score)
    if score < 60:
        return H_RED_BG
    elif 60 <= score <= 70:
        return H_ORANGE_BG
    elif 70 < score <= 80:
        return H_YELLOW_BG
    return H_GREEN_BG


def heatmap_font_color(score) -> str:
    if pd.isna(score):
        return "#000000"
    score = float(score)
    if score < 60:
        return "#FFFFFF"
    elif 60 <= score <= 70:
        return "#FFFFFF"
    elif 70 < score <= 80:
        return "#111111"
    return "#FFFFFF"


def get_dimension_colors() -> dict:
    return {
        "1. งานทำงานเป็นทีม": "#4472C4",
        "2. บุคลากรและพื้นที่การทำงาน": "#ED7D31",
        "3. การเป็นองค์กรแห่งการเรียนรู้ มีการพัฒนาอย่างต่อเนื่อง": "#70AD47",
        "4. การตอบสนองต่อความคลาดเคลื่อน": "#C00000",
        "5. การสนับสนุนของหัวหน้างาน, ผู้จัดการ หรือทีมนำทางคลินิกในเรื่องความปลอดภัย": "#7030A0",
        "6. การสื่อสารเรื่องความคลาดเคลื่อน": "#5B9BD5",
        "7. การสื่อสารที่เปิดกว้าง": "#A5A5A5",
        "8. การรายงานเหตุการณ์ความปลอดภัยของผู้ป่วย/ผู้รับบริการ": "#FFC000",
        "9. การสนับสนุนจากทีมผู้บริหารของสถานพยาบาลในเรื่องความปลอดภัยของผู้ป่วย /ผู้รับบริการ": "#2F5597",
        "10. การส่งต่องานและแลกเปลี่ยนข้อมูล ในการเปลี่ยนผ่านระหว่างหน่วยงานหรือเวร": "#00B0F0",
    }


def dedupe_labels(labels):
    seen = {}
    out = []
    for lab in labels:
        if lab not in seen:
            seen[lab] = 1
            out.append(lab)
        else:
            seen[lab] += 1
            out.append(f"{lab} ({seen[lab]})")
    return out




def wrap_tick_label(label: str, width: int = 18) -> str:
    """
    ทำ wrap text สำหรับชื่อหน่วยงานบนแกน X
    รองรับข้อความไทยที่ยาวและข้อความที่มี /
    """
    if label is None:
        return ""
    s = str(label).strip()
    s = s.replace(" / ", " /| ")
    s = s.replace("/", " / ")
    parts = [p.strip() for p in s.split("|") if p.strip()]
    wrapped_parts = []
    for part in parts:
        if len(part) <= width:
            wrapped_parts.append(part)
            continue
        if " " in part:
            wrapped_parts.append("<br>".join(textwrap.wrap(part, width=width, break_long_words=False)))
        else:
            chunks = [part[i:i+width] for i in range(0, len(part), width)]
            wrapped_parts.append("<br>".join(chunks))
    return "<br>".join(wrapped_parts)


def get_plotly_config_for_heatmap(unit_count: int) -> dict:
    """
    ถ้าหน่วยงานน้อยกว่า 3 คอลัมน์ ให้เอาปุ่ม autoscale ออก
    """
    cfg = {"responsive": True}
    if unit_count < 3:
        cfg["modeBarButtonsToRemove"] = ["autoScale2d"]
    return cfg




def get_heatmap_display_mode(unit_count: int) -> dict:
    """
    กลุ่มที่มีหน่วยงานน้อยกว่า 3 คอลัมน์ ไม่ควรยืดเต็มจอ
    """
    if unit_count <= 1:
        return {"compact": True, "width": 760}
    elif unit_count == 2:
        return {"compact": True, "width": 920}
    return {"compact": False, "width": None}

# =========================================================
# Quadrant page
# =========================================================
@st.cache_data(show_spinner=False)
def load_quadrant_excel(file_obj, sheet_name: str = "HSCS2568 (2)") -> pd.DataFrame:
    raw = pd.read_excel(file_obj, sheet_name=sheet_name, header=None)

    records = []
    current_dimension = None

    for i in range(3, len(raw)):  # เริ่มจากแถว 4 ของ Excel
        dim = raw.iloc[i, 0] if raw.shape[1] > 0 else None
        sub = raw.iloc[i, 1] if raw.shape[1] > 1 else None
        score = raw.iloc[i, 2] if raw.shape[1] > 2 else None

        if pd.notna(dim):
            current_dimension = str(dim).strip()

        if pd.notna(sub) and pd.notna(score):
            sub = str(sub).strip()
            code_match = re.match(r"^([A-Z]\d+)\.\s*", sub)
            code = code_match.group(1) if code_match else ""
            full_name = re.sub(r"^[A-Z]\d+\.\s*", "", sub).strip()

            try:
                score = float(score)
            except Exception:
                continue

            records.append(
                {
                    "dimension": current_dimension,
                    "sub_code": code,
                    "sub_name": full_name,
                    "sub_raw": sub,
                    "sub_score": score,
                }
            )

    df = pd.DataFrame(records)
    if df.empty:
        raise ValueError("ไม่พบข้อมูลมิติย่อยในชีตที่เลือก")

    dim_avg = (
        df.groupby("dimension", dropna=False)["sub_score"]
        .mean()
        .rename("dimension_avg")
        .reset_index()
    )
    df = df.merge(dim_avg, on="dimension", how="left")
    return df


def apply_quadrant_logic(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    out = df["sub_score"].apply(classify_score_quadrant)
    df["quadrant"] = out.apply(lambda x: x[0])
    df["quadrant_color"] = out.apply(lambda x: x[1])
    return df


def score_to_y(score: float, quadrant: str) -> float:
    if quadrant == "ควรพัฒนาด่วน":
        low, high, y0, y1 = 0.0, 60.0, 5.0, 39.0
    elif quadrant == "เร่งพัฒนา":
        low, high, y0, y1 = 60.0, 70.0, 5.0, 39.0
    elif quadrant == "ควรพัฒนาต่อเนื่อง":
        low, high, y0, y1 = 70.0, 80.0, 52.0, 74.0
    else:
        low, high, y0, y1 = 80.0, 100.0, 82.0, 99.0

    ratio = (score - low) / (high - low) if high > low else 0.5
    ratio = max(0.0, min(1.0, ratio))
    ratio = ratio ** 1.8
    return y0 + ratio * (y1 - y0)


def assign_positions_by_quadrant(df: pd.DataFrame) -> pd.DataFrame:
    """
    จัดตำแหน่งจุดแบบกันชนเพื่อลดการซ้อนของ marker และ label
    พร้อมคงหลักการคะแนนสูงอยู่สูงกว่า และโซนเขียวสูงกว่าเหลืองอย่างชัดเจน
    """
    df = df.copy()
    df["plot_x"] = np.nan
    df["plot_y"] = np.nan

    quadrant_meta = {
        "ควรพัฒนาด่วน": {"center_x": 25.0, "half_width": 18.0, "ymin": 5.0, "ymax": 39.0},
        "เร่งพัฒนา": {"center_x": 75.0, "half_width": 18.0, "ymin": 5.0, "ymax": 39.0},
        "ควรพัฒนาต่อเนื่อง": {"center_x": 25.0, "half_width": 18.0, "ymin": 52.0, "ymax": 74.0},
        "ควรส่งเสริม": {"center_x": 75.0, "half_width": 18.0, "ymin": 82.0, "ymax": 99.0},
    }

    all_dims = [d for d in df["dimension"].dropna().unique().tolist()]
    dim_index = {d: i for i, d in enumerate(all_dims)}
    dim_count = max(len(all_dims), 1)

    def clamp(val, low, high):
        return max(low, min(high, val))

    min_dx = 3.6
    min_dy = 3.1

    for quadrant, grp in df.groupby("quadrant", sort=False):
        meta = quadrant_meta[quadrant]
        center_x = meta["center_x"]
        half_width = meta["half_width"]
        ymin = meta["ymin"]
        ymax = meta["ymax"]

        grp_sorted = grp.sort_values(
            ["sub_score", "dimension", "sub_name"],
            ascending=[False, True, True]
        ).copy()
        grp_sorted["rank_in_dim"] = grp_sorted.groupby("dimension").cumcount()

        placed = []
        for idx, row in grp_sorted.iterrows():
            score = float(row["sub_score"])
            dim = row["dimension"]
            rank_in_dim = int(row["rank_in_dim"])
            d_idx = dim_index.get(dim, 0)

            if dim_count == 1:
                dim_offset = 0.0
            else:
                dim_offset = ((d_idx / (dim_count - 1)) - 0.5) * (half_width * 1.35)

            micro_offset = ((rank_in_dim % 5) - 2) * 1.15
            score_offset = ((score * 10) % 7 - 3) * 0.16

            x = center_x + dim_offset + micro_offset + score_offset
            y = score_to_y(score, quadrant)

            x = clamp(x, center_x - half_width, center_x + half_width)
            y = clamp(y, ymin, ymax)

            placed.append({
                "idx": idx,
                "x": x,
                "y": y,
                "score": score,
            })

        for _ in range(120):
            moved = False
            placed.sort(key=lambda p: (-p["score"], p["x"]))
            for i in range(len(placed)):
                for j in range(i + 1, len(placed)):
                    a = placed[i]
                    b = placed[j]
                    dx = b["x"] - a["x"]
                    dy = b["y"] - a["y"]

                    if abs(dx) < min_dx and abs(dy) < min_dy:
                        overlap_x = min_dx - abs(dx)
                        overlap_y = min_dy - abs(dy)

                        push_x = overlap_x / 2 + 0.05
                        if dx >= 0:
                            a["x"] -= push_x
                            b["x"] += push_x
                        else:
                            a["x"] += push_x
                            b["x"] -= push_x

                        if abs(dy) < min_dy * 0.7:
                            push_y = overlap_y / 2 + 0.03
                            if a["score"] >= b["score"]:
                                a["y"] += push_y * 0.35
                                b["y"] -= push_y * 0.85
                            else:
                                a["y"] -= push_y * 0.85
                                b["y"] += push_y * 0.35

                        a["x"] = clamp(a["x"], center_x - half_width, center_x + half_width)
                        b["x"] = clamp(b["x"], center_x - half_width, center_x + half_width)
                        a["y"] = clamp(a["y"], ymin, ymax)
                        b["y"] = clamp(b["y"], ymin, ymax)
                        moved = True

            if not moved:
                break

        for p in placed:
            df.loc[p["idx"], "plot_x"] = p["x"]
            df.loc[p["idx"], "plot_y"] = p["y"]

    return df


def build_quadrant_figure(df: pd.DataFrame) -> go.Figure:
    dim_colors = get_dimension_colors()
    fig = go.Figure()

    fig.add_shape(type="rect", x0=0,  x1=50, y0=50, y1=100, fillcolor=Q_YELLOW_BG, opacity=0.88, line=dict(width=0), layer="below")
    fig.add_shape(type="rect", x0=50, x1=100, y0=50, y1=100, fillcolor=Q_GREEN_BG, opacity=0.88, line=dict(width=0), layer="below")
    fig.add_shape(type="rect", x0=0,  x1=50, y0=0,  y1=50, fillcolor=Q_RED_BG, opacity=0.88, line=dict(width=0), layer="below")
    fig.add_shape(type="rect", x0=50, x1=100, y0=0,  y1=50, fillcolor=Q_ORANGE_BG, opacity=0.88, line=dict(width=0), layer="below")

    fig.add_vline(x=50, line_dash="dash", line_width=2.0, line_color="rgba(255,255,255,0.95)")
    fig.add_hline(y=50, line_dash="dash", line_width=2.0, line_color="rgba(255,255,255,0.95)")

    for dim in df["dimension"].dropna().unique():
        sub = df[df["dimension"] == dim].copy()
        fig.add_trace(
            go.Scatter(
                x=sub["plot_x"],
                y=sub["plot_y"],
                mode="markers+text",
                name=dim,
                legendgroup=dim,
                marker=dict(
                    size=18,
                    color=dim_colors.get(dim, "#4F81BD"),
                    symbol="circle",
                    line=dict(color="white", width=2.3),
                ),
                text=sub["sub_code"].fillna(""),
                textposition="middle center",
                textfont=dict(size=9, color="white"),
                customdata=sub[
                    ["sub_code", "sub_name", "dimension", "sub_score", "dimension_avg", "quadrant", "quadrant_color"]
                ].values,
                hovertemplate=(
                    "<b>%{customdata[0]}</b><br>"
                    "มิติย่อย: %{customdata[1]}<br>"
                    "มิติหลัก: %{customdata[2]}<br>"
                    "คะแนนมิติย่อย: %{customdata[3]:.1f}%<br>"
                    "ค่าเฉลี่ยมิติหลัก: %{customdata[4]:.1f}%<br>"
                    "Quadrant: %{customdata[5]}<br>"
                    "เกณฑ์สี Quadrant: %{customdata[6]}"
                    "<extra></extra>"
                ),
            )
        )

    counts = df["quadrant"].value_counts().to_dict()
    fig.add_annotation(
        x=3.5, y=96.5, xanchor="left", yanchor="top",
        text=f"<b>ควรพัฒนาต่อเนื่อง</b><br>% positive response 70.1–80<br>{counts.get('ควรพัฒนาต่อเนื่อง', 0)} ข้อ",
        showarrow=False, align="left",
        font=dict(size=14, color="#4E4300"),
        bgcolor="rgba(255,255,255,0.34)"
    )
    fig.add_annotation(
        x=53.5, y=96.5, xanchor="left", yanchor="top",
        text=f"<b>ควรส่งเสริม</b><br>% positive response > 80<br>{counts.get('ควรส่งเสริม', 0)} ข้อ",
        showarrow=False, align="left",
        font=dict(size=14, color="white"),
        bgcolor="rgba(0,0,0,0.14)"
    )
    fig.add_annotation(
        x=3.5, y=46.5, xanchor="left", yanchor="top",
        text=f"<b>ควรพัฒนาด่วน</b><br>% positive response < 60<br>{counts.get('ควรพัฒนาด่วน', 0)} ข้อ",
        showarrow=False, align="left",
        font=dict(size=14, color="white"),
        bgcolor="rgba(0,0,0,0.14)"
    )
    fig.add_annotation(
        x=53.5, y=46.5, xanchor="left", yanchor="top",
        text=f"<b>เร่งพัฒนา</b><br>% positive response 60–70<br>{counts.get('เร่งพัฒนา', 0)} ข้อ",
        showarrow=False, align="left",
        font=dict(size=14, color="white"),
        bgcolor="rgba(0,0,0,0.14)"
    )

    fig.update_layout(
        title=dict(text="Quadrant Graph แบบ Interactive Infographic", x=0.5, font=dict(size=28, color="#173B71")),
        paper_bgcolor="#F8FBFF",
        plot_bgcolor="#F8FBFF",
        hoverlabel=dict(bgcolor="white", font_size=14, font_family="Arial"),
        legend=dict(
            title="สีของจุด = มิติหลัก",
            orientation="h",
            yanchor="bottom",
            y=-0.28,
            xanchor="center",
            x=0.5,
            bgcolor="rgba(255,255,255,0.92)"
        ),
        margin=dict(l=40, r=30, t=80, b=145),
        height=790,
    )

    fig.update_xaxes(range=[0, 100], showgrid=False, showticklabels=False, zeroline=False)
    fig.update_yaxes(range=[0, 100], showgrid=False, showticklabels=False, zeroline=False, scaleanchor="x", scaleratio=1)
    return fig


def quadrant_summary(df: pd.DataFrame) -> pd.DataFrame:
    order = ["ควรพัฒนาด่วน", "เร่งพัฒนา", "ควรพัฒนาต่อเนื่อง", "ควรส่งเสริม"]
    out = (
        df.groupby(["quadrant", "quadrant_color"])
        .agg(
            จำนวนข้อ=("sub_name", "count"),
            ค่าเฉลี่ยคะแนนมิติย่อย=("sub_score", "mean")
        )
        .reset_index()
    )
    out["order"] = out["quadrant"].apply(lambda x: order.index(x) if x in order else 999)
    out = out.sort_values("order").drop(columns=["order"])
    out = out.rename(columns={"quadrant": "Quadrant", "quadrant_color": "สี"})
    return out


# =========================================================
# Heatmap page
# =========================================================
def _resolve_header_value(ws, merge_map, row_num, col_num):
    v = ws.cell(row_num, col_num).value
    if v is None and (row_num, col_num) in merge_map:
        v = merge_map[(row_num, col_num)]
    return v


@st.cache_data(show_spinner=False)
def load_heatmap_excel(file_obj, sheet_name: str = "HSCS2568") -> tuple[pd.DataFrame, list[str]]:
    """
    อ่านไฟล์ heatmap โดยใช้ merged cells จริงของ Excel
    เพื่อป้องกันการ forward-fill ผิดคอลัมน์ เช่นคอลัมน์ 'ภาพรวม'
    """
    raw = pd.read_excel(file_obj, sheet_name=sheet_name, header=None)

    if hasattr(file_obj, "seek"):
        file_obj.seek(0)

    wb = openpyxl.load_workbook(file_obj, data_only=True)
    ws = wb[sheet_name]

    merge_map = {}
    for mr in ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = mr.bounds
        if min_row <= 3:
            top_val = ws.cell(min_row, min_col).value
            for r in range(min_row, max_row + 1):
                for c in range(min_col, max_col + 1):
                    merge_map[(r, c)] = top_val

    data_rows = []
    current_dimension = None

    for r in range(3, len(raw)):
        dim = raw.iloc[r, 0] if raw.shape[1] > 0 else None
        sub = raw.iloc[r, 1] if raw.shape[1] > 1 else None

        if pd.notna(dim):
            current_dimension = str(dim).strip()

        numeric_found = False
        for c in range(2, raw.shape[1]):
            val = raw.iloc[r, c]
            if pd.notna(val):
                try:
                    float(val)
                    numeric_found = True
                    break
                except Exception:
                    pass

        if pd.notna(sub) and numeric_found:
            code_match = re.match(r"^([A-Z]\d+)\.\s*", str(sub).strip())
            code = code_match.group(1) if code_match else ""
            full_name = re.sub(r"^[A-Z]\d+\.\s*", "", str(sub).strip()).strip()
            data_rows.append((r, {"dimension": current_dimension, "sub_code": code, "sub_name": full_name}))

    if not data_rows:
        raise ValueError("ไม่พบข้อมูล heatmap ในชีตที่เลือก")

    row_indices = [r for r, _ in data_rows]

    score_cols = []
    for c in range(2, raw.shape[1]):
        any_numeric = False
        for r in row_indices:
            val = raw.iloc[r, c]
            if pd.notna(val):
                try:
                    float(val)
                    any_numeric = True
                    break
                except Exception:
                    pass
        if any_numeric:
            score_cols.append(c)

    if not score_cols:
        raise ValueError("ไม่พบคอลัมน์คะแนนในชีตที่เลือก")

    records = []
    groups_found = []

    for r, base in data_rows:
        for c in score_cols:
            # pandas 0-based -> openpyxl 1-based
            col_num = c + 1

            top_group = _resolve_header_value(ws, merge_map, 1, col_num)
            division = _resolve_header_value(ws, merge_map, 2, col_num)
            unit = _resolve_header_value(ws, merge_map, 3, col_num)

            top_group = str(top_group).replace("\n", " ").strip() if top_group is not None else ""
            division = str(division).replace("\n", " ").strip() if division is not None else ""
            unit = str(unit).replace("\n", " ").strip() if unit is not None else ""

            if not unit:
                unit = division if division else top_group

            groups_found.append(top_group)

            val = raw.iloc[r, c]
            score = np.nan
            if pd.notna(val):
                try:
                    score = float(val)
                except Exception:
                    score = np.nan

            records.append(
                {
                    "group": top_group,
                    "division": division,
                    "unit": unit,
                    "dimension": base["dimension"],
                    "sub_code": base["sub_code"],
                    "sub_name": base["sub_name"],
                    "score": score,
                    "col_index": c,
                }
            )

    long_df = pd.DataFrame(records)

    ordered_groups = []
    for g in groups_found:
        if g and g not in ordered_groups:
            ordered_groups.append(g)

    return long_df, ordered_groups


def build_heatmap_figure(long_df: pd.DataFrame, title_text: str = "") -> go.Figure:
    df = long_df.copy()

    row_order = df[["sub_code", "sub_name", "dimension"]].drop_duplicates()
    row_order["row_label"] = row_order["sub_code"].replace("", np.nan).fillna("NA")

    col_order = (
        df[["col_index", "unit", "division", "group"]]
        .drop_duplicates()
        .sort_values("col_index")
        .reset_index(drop=True)
    )
    col_order["col_label"] = dedupe_labels(col_order["unit"].tolist())

    df = df.merge(col_order[["col_index", "col_label"]], on="col_index", how="left")

    row_labels = row_order["row_label"].tolist()
    col_labels = col_order["col_label"].tolist()

    pivot = (
        df.assign(
            row_label=df["sub_code"].replace("", np.nan).fillna("NA")
        )
        .pivot_table(index="row_label", columns="col_label", values="score", aggfunc="mean")
        .reindex(index=row_labels, columns=col_labels)
    )

    row_meta = row_order.set_index("row_label")[["sub_code", "sub_name", "dimension"]]
    col_meta = col_order.set_index("col_label")[["unit", "division", "group"]]

    customdata = []
    text_x = []
    text_y = []
    text_values = []
    text_colors = []

    for rlab in pivot.index:
        row_cd = []
        for clab in pivot.columns:
            score = pivot.loc[rlab, clab]
            row_cd.append([
                row_meta.loc[rlab, "sub_code"],
                row_meta.loc[rlab, "sub_name"],
                row_meta.loc[rlab, "dimension"],
                col_meta.loc[clab, "unit"],
                col_meta.loc[clab, "division"],
                col_meta.loc[clab, "group"],
            ])

            if pd.notna(score):
                text_x.append(clab)
                text_y.append(rlab)
                text_values.append(f"{score:.1f}")
                text_colors.append(heatmap_font_color(score))

        customdata.append(row_cd)

    z = pivot.values.astype(float)

    colorscale = [
        [0.0, H_RED_BG], [0.599999, H_RED_BG],
        [0.6, H_ORANGE_BG], [0.7, H_ORANGE_BG],
        [0.700001, H_YELLOW_BG], [0.8, H_YELLOW_BG],
        [0.800001, H_GREEN_BG], [1.0, H_GREEN_BG],
    ]

    fig = go.Figure()

    fig.add_trace(
        go.Heatmap(
            z=z,
            x=col_labels,
            y=row_labels,
            zmin=0,
            zmax=100,
            colorscale=colorscale,
            showscale=False,
            customdata=customdata,
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "มิติย่อย: %{customdata[1]}<br>"
                "มิติหลัก: %{customdata[2]}<br>"
                "หน่วยงาน: %{customdata[3]}<br>"
                "ฝ่าย/งาน: %{customdata[4]}<br>"
                "กลุ่มงาน: %{customdata[5]}<br>"
                "คะแนน: %{z:.1f}%<extra></extra>"
            ),
            xgap=1,
            ygap=1,
        )
    )

    fig.add_trace(
        go.Scatter(
            x=text_x,
            y=text_y,
            mode="text",
            text=text_values,
            textfont=dict(size=10, color=text_colors),
            hoverinfo="skip",
            showlegend=False,
        )
    )

    unit_count = len(col_labels)
    display_mode = get_heatmap_display_mode(unit_count)

    fig.update_layout(
        title=None,
        paper_bgcolor="#F8FBFF",
        plot_bgcolor="#F8FBFF",
        margin=dict(l=6, r=10, t=30, b=20),
        height=max(720, 28 * len(row_labels) + 180),
        width=display_mode["width"],
    )

    fig.update_xaxes(side="top", tickangle=-35, showgrid=False, tickfont=dict(size=11), automargin=True)
    fig.update_yaxes(autorange="reversed", showgrid=False, tickfont=dict(size=10), automargin=True)
    return fig


def style_heatmap_table(df: pd.DataFrame):
    def style_cell(val):
        if pd.isna(val):
            return ""
        bg = heatmap_bg_color(val)
        fg = heatmap_font_color(val)
        return f"background-color: {bg}; color: {fg}; font-weight: 700; text-align: center;"

    return (
        df.style
        .format(lambda v: "" if pd.isna(v) else f"{v:.1f}")
        .applymap(style_cell)
    )


def render_heatmap_page(heatmap_source, heatmap_sheet: str, selected_page: str):
    long_df, groups = load_heatmap_excel(heatmap_source, sheet_name=heatmap_sheet)

    if selected_page == "Heatmap: ภาพรวมทุกกลุ่ม":
        filtered = long_df.copy()
        page_title = "Heatmap: ภาพรวมทุกกลุ่ม"
        page_desc = "Heatmap แยกตามกลุ่มงานจากแถวบนสุด"
    else:
        target_group = selected_page.replace("Heatmap: ", "", 1)
        filtered = long_df[long_df["group"] == target_group].copy()
        page_title = f"Heatmap: {target_group}"
        page_desc = "Heatmap แยกตามกลุ่มงานจากแถวบนสุด"

    st.title(page_title)
    st.markdown(page_desc)

    if filtered.empty:
        st.warning("ไม่มีข้อมูลสำหรับหน้านี้")
        return

    all_dims = filtered["dimension"].dropna().unique().tolist()
    all_units = filtered["unit"].dropna().unique().tolist()

    with st.sidebar.expander("ตัวกรอง Heatmap", expanded=True):
        dim_filter = st.multiselect(
            "เลือกมิติหลัก",
            options=all_dims,
            default=all_dims,
            key=f"hm_dim_{selected_page}"
        )
        unit_filter = st.multiselect(
            "เลือกหน่วยงาน/คอลัมน์",
            options=all_units,
            default=all_units,
            key=f"hm_unit_{selected_page}"
        )
        show_table = st.checkbox("แสดงตารางสีด้านล่าง", value=True, key=f"hm_table_{selected_page}")

    filtered = filtered[
        filtered["dimension"].isin(dim_filter) &
        filtered["unit"].isin(unit_filter)
    ].copy()

    if filtered.empty:
        st.warning("ไม่มีข้อมูลหลังจากกรอง")
        return

    c1, c2, c3 = st.columns(3)
    c1.metric("จำนวนมิติย่อย", f"{filtered[['sub_code','sub_name']].drop_duplicates().shape[0]:,}")
    c2.metric("จำนวนหน่วยงาน", f"{filtered['unit'].nunique():,}")
    c3.metric("จำนวน cell ที่มีคะแนน", f"{filtered['score'].notna().sum():,}")

    fig = build_heatmap_figure(filtered, title_text="")
    display_mode = get_heatmap_display_mode(filtered["unit"].nunique())
    st.plotly_chart(fig, use_container_width=not display_mode["compact"])

    with st.expander("ดูคำอธิบายรหัสมิติย่อย", expanded=False):
        show_map = (
            filtered[["sub_code", "dimension", "sub_name"]]
            .drop_duplicates()
            .sort_values(["dimension", "sub_code", "sub_name"])
            .rename(columns={"sub_code": "รหัส", "dimension": "มิติหลัก", "sub_name": "ชื่อข้อย่อย"})
        )
        st.dataframe(show_map, use_container_width=True, hide_index=True)

    if show_table:
        pivot = (
            filtered.assign(
                row_label=filtered["sub_code"].replace("", np.nan).fillna("NA")
            )
            .pivot(index="row_label", columns="unit", values="score")
            .sort_index()
        )
        st.markdown("### ตาราง Heatmap แบบเติมสีแต่ละตัวเลข")
        st.dataframe(style_heatmap_table(pivot), use_container_width=False)


def render_quadrant_page(quad_source, quad_sheet: str):
    st.title("📊 Quadrant Graph แบบ Interactive Infographic")
    st.markdown(
        "หน้าเดิมของ Quadrant พร้อม logic ล่าสุดที่คุณอนุมัติแล้ว  \n"
        "คำอธิบายด้านบนถูกเปลี่ยนเป็น **% positive response** แล้ว"
    )

    df = load_quadrant_excel(quad_source, sheet_name=quad_sheet)
    df = apply_quadrant_logic(df)
    df = assign_positions_by_quadrant(df)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("จำนวนมิติย่อย", f"{len(df):,}")
    c2.metric("Quadrant แดง", f"{(df['quadrant'] == 'ควรพัฒนาด่วน').sum():,}")
    c3.metric("Quadrant ส้ม", f"{(df['quadrant'] == 'เร่งพัฒนา').sum():,}")
    c4.metric("Quadrant เขียว", f"{(df['quadrant'] == 'ควรส่งเสริม').sum():,}")

    fig = build_quadrant_figure(df)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("### สรุปตาม Quadrant")
    st.dataframe(quadrant_summary(df), use_container_width=True, hide_index=True)

    st.markdown("### รายการมิติย่อยทั้งหมด")
    quadrant_filter = st.multiselect(
        "กรองตาม Quadrant",
        options=["ควรพัฒนาด่วน", "เร่งพัฒนา", "ควรพัฒนาต่อเนื่อง", "ควรส่งเสริม"],
        default=["ควรพัฒนาด่วน", "เร่งพัฒนา", "ควรพัฒนาต่อเนื่อง", "ควรส่งเสริม"],
        key="quad_filter"
    )

    dim_filter = st.multiselect(
        "กรองตามมิติหลัก",
        options=list(df["dimension"].dropna().unique()),
        default=list(df["dimension"].dropna().unique()),
        key="quad_dim_filter"
    )

    view_df = df[df["quadrant"].isin(quadrant_filter) & df["dimension"].isin(dim_filter)].copy()
    view_df = view_df.rename(
        columns={
            "dimension": "มิติหลัก",
            "sub_code": "รหัส",
            "sub_name": "ชื่อมิติย่อย",
            "sub_score": "คะแนนมิติย่อย (%)",
            "dimension_avg": "ค่าเฉลี่ยมิติหลัก (%)",
            "quadrant": "Quadrant",
            "quadrant_color": "สีของ Quadrant",
        }
    )

    st.dataframe(
        view_df[
            ["มิติหลัก", "รหัส", "ชื่อมิติย่อย", "คะแนนมิติย่อย (%)", "ค่าเฉลี่ยมิติหลัก (%)", "Quadrant", "สีของ Quadrant"]
        ],
        use_container_width=True,
        hide_index=True
    )




REPORT_URL = "https://sites.google.com/view/mch-hscs67-68/%E0%B8%A0%E0%B8%B2%E0%B8%9E%E0%B8%A3%E0%B8%A7%E0%B8%A1?authuser=0"
REPORT_PREVIEW_IMAGE = Path("hscs_report_preview.png")


def render_full_report_page():
    st.title("📘 รายงาน HSCS ฉบับสมบูรณ์")
    st.markdown("Preview card ของ Google Sites พร้อมปุ่มเปิดรายงานฉบับเต็ม")

    card_html = """
    <div style="
        background: linear-gradient(180deg, #ffffff 0%, #f8fbff 100%);
        border: 1px solid #dce6f2;
        border-radius: 18px;
        padding: 18px 18px 12px 18px;
        box-shadow: 0 6px 20px rgba(23,59,113,0.06);
        margin-bottom: 14px;">
        <div style="font-size: 24px; font-weight: 700; color: #173B71; margin-bottom: 6px;">
            Hospital Safety Culture Survey (HSCS)
        </div>
        <div style="font-size: 15px; color: #4a678f; line-height: 1.6;">
            รายงานฉบับสมบูรณ์บน Google Sites ของ HSCS ปี 2567–2568
            หน้านี้ใช้ <b>preview card</b> แทนการฝังเว็บตรง เพื่อให้แสดงผลเสถียรกว่า
        </div>
    </div>
    """
    st.markdown(card_html, unsafe_allow_html=True)

    if REPORT_PREVIEW_IMAGE.exists():
        st.image(str(REPORT_PREVIEW_IMAGE), use_container_width=True)
    else:
        st.info("ไม่พบภาพ preview ในแพ็กเกจ แต่ยังเปิดรายงานฉบับเต็มได้ตามปกติ")

    c1, c2 = st.columns([1, 2])
    with c1:
        st.link_button("🔗 เปิด Google Sites ฉบับเต็ม", REPORT_URL, use_container_width=True)
    with c2:
        st.caption("ถ้าต้องการดูรายละเอียดทั้งหมด แนะนำให้เปิดในแท็บใหม่เพื่อการใช้งานที่ครบถ้วนที่สุด")


# =========================================================
# App shell
# =========================================================
default_quad_file = Path("plotgraph_quadrant_infographic.xlsx")
default_heatmap_file = Path("HSCS2568_interac.xlsx")

st.sidebar.title("HSCS Web Service")

uploaded_quad = st.sidebar.file_uploader(
    "อัปโหลดไฟล์ Quadrant Excel",
    type=["xlsx"],
    key="quad_uploader"
)
quad_sheet = st.sidebar.text_input("ชื่อชีต Quadrant", value="HSCS2568 (2)")

uploaded_heatmap = st.sidebar.file_uploader(
    "อัปโหลดไฟล์ Heatmap Excel",
    type=["xlsx"],
    key="heatmap_uploader"
)
heatmap_sheet = st.sidebar.text_input("ชื่อชีต Heatmap", value="HSCS2568")

quad_source = uploaded_quad if uploaded_quad is not None else (default_quad_file if default_quad_file.exists() else None)
heatmap_source = uploaded_heatmap if uploaded_heatmap is not None else (default_heatmap_file if default_heatmap_file.exists() else None)

heatmap_pages = ["Heatmap: ภาพรวมทุกกลุ่ม"]
if heatmap_source is not None:
    try:
        if hasattr(heatmap_source, "seek"):
            heatmap_source.seek(0)
        _, group_names = load_heatmap_excel(heatmap_source, sheet_name=heatmap_sheet)
        heatmap_pages += [f"Heatmap: {g}" for g in group_names]
        if hasattr(heatmap_source, "seek"):
            heatmap_source.seek(0)
    except Exception:
        pass

page_options = ["Quadrant 4 Quadrants"] + heatmap_pages + ["รายงาน HSCS ฉบับสมบูรณ์"]

page = st.sidebar.radio(
    "เลือกหน้าที่ต้องการดู",
    page_options,
    index=0
)

st.sidebar.markdown("---")
st.sidebar.markdown(
    """
**เกณฑ์สีที่ใช้ร่วมกัน**
- 🔴 แดง: คะแนน < 60
- 🟠 ส้ม: คะแนน 60–70
- 🟡 เหลือง: คะแนน 70.1–80
- 🟢 เขียว: คะแนน > 80
"""
)

if page == "Quadrant 4 Quadrants":
    if quad_source is None:
        st.warning("กรุณาอัปโหลดไฟล์ Quadrant Excel ก่อน")
        st.stop()
    render_quadrant_page(quad_source, quad_sheet)
elif page == "รายงาน HSCS ฉบับสมบูรณ์":
    render_full_report_page()
else:
    if heatmap_source is None:
        st.warning("กรุณาอัปโหลดไฟล์ Heatmap Excel ก่อน")
        st.stop()
    if hasattr(heatmap_source, "seek"):
        heatmap_source.seek(0)
    render_heatmap_page(heatmap_source, heatmap_sheet, page)
