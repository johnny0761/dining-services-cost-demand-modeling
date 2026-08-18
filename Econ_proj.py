import pandas as pd
import numpy as np
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from matplotlib.ticker import FuncFormatter
import plotly.graph_objects as go
import seaborn as sns
import statsmodels.api as sm
from scipy.stats import linregress
import statsmodels.formula.api as smf
import statsmodels.stats.api as sms
import statsmodels.stats.proportion as sp
import statsmodels.stats.contingency_tables as ct
import statsmodels.stats.multitest as mt
import statsmodels.stats.multicomp as mc
from pathlib import Path
import sys



file = "A.5_Non-Perishables.xlsx"
sheet_name = "Data"

# Optional override: python3 Econ_proj.py path/to/file.xlsx [sheet_name]
if len(sys.argv) >= 2:
    file = sys.argv[1]
if len(sys.argv) >= 3:
    sheet_name = sys.argv[2]

file_path = Path(file).expanduser()
if not file_path.exists():
    raise FileNotFoundError(
        f"Excel file not found: {file_path}\n"
        f"Tip: run `python3 Econ_proj.py \"{file}\"` from this folder, or pass the correct path."
    )

# load the main sheet
try:
    df = pd.read_excel(file_path, sheet_name=sheet_name)
except Exception as e:
    try:
        sheets = pd.ExcelFile(file_path).sheet_names
    except Exception:
        sheets = None
    msg = f"Failed to read {file_path} (sheet={sheet_name!r})."
    if sheets:
        msg += f" Available sheets: {sheets!r}"
    raise RuntimeError(msg) from e

# clean column names
df.columns = df.columns.str.strip()

# count how many unique items and how often each appears (purchased)
item_id_col = "Item ID"
item_name_col = "Item Name"

if item_id_col not in df.columns:
    raise KeyError(f"Expected column {item_id_col!r} not found. Available columns: {df.columns.tolist()!r}")

base_cols = [item_id_col] + ([item_name_col] if item_name_col in df.columns else [])
items = df[base_cols].copy()

# normalize to reduce accidental duplicates from whitespace/case
if item_name_col in items.columns:
    items[item_name_col] = items[item_name_col].astype("string").str.strip()
items[item_id_col] = items[item_id_col].astype("string").str.strip()

items = items.dropna(subset=[item_id_col])
unique_item_count = items[item_id_col].nunique(dropna=True)

item_counts = (
    items.groupby(base_cols, dropna=False)
    .size()
    .reset_index(name="times_purchased")
    .sort_values(["times_purchased"] + base_cols, ascending=[False] + [True] * len(base_cols))
)

print("\nUnique items (by Item ID):", unique_item_count)
print("\nCounts per item (times purchased):")
print(item_counts.to_string(index=False))

out_path = Path("item_purchase_counts.csv")
item_counts.to_csv(out_path, index=False)
print(f"\nSaved: {out_path.resolve()}")

# Total spend per calendar week (sum of all purchases in that week; x = week end)
date_col = "Period End"
cost_col = "Total"

if date_col not in df.columns:
    raise KeyError(f"Expected column {date_col!r} not found. Available columns: {df.columns.tolist()!r}")
if cost_col not in df.columns:
    raise KeyError(f"Expected column {cost_col!r} not found. Available columns: {df.columns.tolist()!r}")

week_spend = df[[date_col, cost_col]].dropna(subset=[date_col, cost_col]).copy()
week_spend[date_col] = pd.to_datetime(week_spend[date_col], errors="coerce")
week_spend = week_spend.dropna(subset=[date_col])
week_series = week_spend.set_index(date_col)[cost_col].sort_index()
# Weeks end Sunday; label each bucket by that Sunday (end of week).
weekly = week_series.resample("W-SUN", label="right", closed="right").sum()

weekly_csv = Path("weekly_total_cost.csv")
weekly.to_frame(name="total_cost").rename_axis("week_ending").to_csv(weekly_csv)
print(f"\nSaved: {weekly_csv.resolve()}")

plt.figure(figsize=(11, 5))
plt.plot(weekly.index, weekly.values, linewidth=2, marker="o", markersize=4)
plt.title("Total cost per week (sum of purchases in each week; week ending Sunday)")
plt.xlabel("Week ending (Sunday)")
plt.ylabel("Total cost")
plt.tight_layout()

weekly_plot_path = Path("weekly_costs.png")
plt.savefig(weekly_plot_path, dpi=200)
plt.close()
print(f"Saved: {weekly_plot_path.resolve()}")

# cost per unit over time — one point per purchase (row)
qty_col = "Quantity"
if qty_col not in df.columns:
    raise KeyError(f"Expected column {qty_col!r} not found. Available columns: {df.columns.tolist()!r}")

purchase_cols = [date_col, cost_col, qty_col, item_id_col]
if item_name_col in df.columns:
    purchase_cols.append(item_name_col)
purchases = df[purchase_cols].copy()
purchases[date_col] = pd.to_datetime(purchases[date_col], errors="coerce")
purchases[item_id_col] = purchases[item_id_col].astype("string").str.strip()
if item_name_col in purchases.columns:
    purchases[item_name_col] = purchases[item_name_col].astype("string").str.strip()
purchases = purchases.dropna(subset=[date_col, cost_col, qty_col, item_id_col])
purchases = purchases[purchases[qty_col] > 0]
purchases["cost_per_unit"] = purchases[cost_col] / purchases[qty_col]

# Legend uses item name; fall back to ID if name missing
if item_name_col in purchases.columns:
    purchases["series_label"] = purchases[item_name_col].fillna(purchases[item_id_col])
else:
    purchases["series_label"] = purchases[item_id_col]

purchases = purchases.sort_values(["series_label", date_col])
purchases_csv = Path("cost_per_unit_per_purchase.csv")
purchases.to_csv(purchases_csv, index=False)
print(f"Saved: {purchases_csv.resolve()}")


def purchase_line_by_item(plot_df: pd.DataFrame, y_col: str, ylabel: str, title: str, out_path: Path) -> None:
    """Same layout for any per-purchase series (matches total-cost chart style)."""
    plt.figure(figsize=(12, 7))
    sns.lineplot(
        data=plot_df,
        x=date_col,
        y=y_col,
        hue="series_label",
        marker="o",
        dashes=False,
        linewidth=1.6,
        markersize=5,
        legend="full",
    )
    plt.title(title)
    plt.xlabel("Purchase date")
    plt.ylabel(ylabel)
    plt.legend(title="Item", bbox_to_anchor=(1.02, 1), loc="upper left", borderaxespad=0, fontsize=8)
    plt.tight_layout()
    plt.savefig(out_path, dpi=200, bbox_inches="tight")
    plt.close()
    print(f"Saved: {out_path.resolve()}")


# Same chart pattern as total cost: one line per item, markers, item names in legend
purchase_line_by_item(
    purchases,
    "cost_per_unit",
    "Cost per unit (Total ÷ Quantity)",
    "Cost per unit per purchase over time (by item)",
    Path("cost_per_unit_over_time.png"),
)

purchase_line_by_item(
    purchases,
    cost_col,
    "Total cost",
    "Total cost per purchase over time (by item)",
    Path("total_cost_per_item_over_time.png"),
)

# Meal plan swipes: total swipes used per month (combining Com/Res + commuter/on-campus)
meal_sheet_name = "Meal Plan "
try:
    meal_raw = pd.read_excel(file_path, sheet_name=meal_sheet_name, header=None)
except Exception as e:
    raise RuntimeError(f"Failed to read meal plan sheet {meal_sheet_name!r} from {file_path}.") from e

# Find the key rows (Dates, Description) that define the weekly blocks.
first_col = meal_raw.iloc[:, 0].astype("string").str.strip()
dates_row_idx = first_col[first_col == "Dates"].index
desc_row_idx = first_col[first_col == "Description"].index
if len(dates_row_idx) == 0 or len(desc_row_idx) == 0:
    raise RuntimeError(
        "Could not parse meal plan sheet layout (expected rows labeled 'Dates' and 'Description')."
    )
dates_row_idx = int(dates_row_idx[0])
desc_row_idx = int(desc_row_idx[0])

dates_row = meal_raw.iloc[dates_row_idx]
desc_row = meal_raw.iloc[desc_row_idx].astype("string").str.strip()

# Weekly blocks repeat: [Unknown, Breakfast, Brunch, Lunch, Dinner, Total]
total_cols = [int(i) for i, v in enumerate(desc_row) if str(v).strip().lower() == "total"]
if not total_cols:
    raise RuntimeError("Could not find 'Total' columns in meal plan sheet.")

plan_rows = meal_raw.iloc[desc_row_idx + 1 :].copy()
plan_rows = plan_rows.rename(columns={0: "plan"})
plan_rows["plan"] = plan_rows["plan"].astype("string").str.strip()
plan_rows = plan_rows.dropna(subset=["plan"])

def _normalize_plan_name(s: str) -> str:
    t = str(s).strip()
    # Rule 1: Block 150 com/res are the same
    t = t.replace("Blk150-Com", "Blk150").replace("Blk150-Res", "Blk150")
    # Rule 2: combine commuter and on campus (remove explicit Com/Res tags)
    t = t.replace("-Com", "").replace("-Res", "")
    return t

plan_rows["plan_norm"] = plan_rows["plan"].map(_normalize_plan_name)

def _parse_meal_week_end(val):
    # Excel dates may be strings ("February 05, 2022") or serials (e.g., 44647).
    if isinstance(val, (int, float)) and not pd.isna(val):
        # Excel's day 0 is 1899-12-30 in pandas' convention.
        return pd.to_datetime(val, unit="D", origin="1899-12-30", errors="coerce")
    return pd.to_datetime(val, errors="coerce")

weekly_totals = []
for total_col in total_cols:
    # End-of-week date is the 'To <date>' cell immediately before the Total column.
    end_date_cell = dates_row.iloc[total_col - 1] if total_col - 1 >= 0 else None
    week_end = _parse_meal_week_end(end_date_cell)
    if pd.isna(week_end):
        # Some weeks may be blank; skip them.
        continue
    swipes = pd.to_numeric(plan_rows.iloc[:, total_col], errors="coerce").fillna(0)
    weekly_totals.append({"week_ending": week_end, "swipes_total": float(swipes.sum())})

weekly_swipes = (
    pd.DataFrame(weekly_totals)
    .dropna(subset=["week_ending"])
    .sort_values("week_ending")
    .drop_duplicates(subset=["week_ending"], keep="last")
    .set_index("week_ending")["swipes_total"]
)

monthly_swipes = weekly_swipes.resample("ME").sum()

meal_monthly_csv = Path("meal_swipes_per_month.csv")
monthly_swipes.to_frame(name="total_swipes").rename_axis("month_end").to_csv(meal_monthly_csv)
print(f"Saved: {meal_monthly_csv.resolve()}")

plt.figure(figsize=(11, 5))
plt.plot(monthly_swipes.index, monthly_swipes.values, linewidth=2, marker="o", markersize=4)
plt.title("Total meal swipes used per month")
plt.xlabel("Month")
plt.ylabel("Total swipes")
plt.tight_layout()

meal_monthly_plot = Path("meal_swipes_per_month.png")
plt.savefig(meal_monthly_plot, dpi=200, bbox_inches="tight")
plt.close()
print(f"Saved: {meal_monthly_plot.resolve()}")

# Meal swipe demand units: all meal swipe totals collapsed by calendar month, ignoring year.
month_numbers = list(range(1, 13))
month_labels = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
calendar_month_swipes = pd.Series(0.0, index=month_numbers, name="swipes")
for from_col, value in dates_row.items():
    if str(value).strip().lower() != "from":
        continue
    if from_col + 5 >= meal_raw.shape[1] or from_col + 1 >= meal_raw.shape[1]:
        continue
    week_start = _parse_meal_week_end(dates_row.iloc[from_col + 1])
    if pd.isna(week_start):
        continue
    swipes = pd.to_numeric(plan_rows.iloc[:, from_col + 5], errors="coerce").fillna(0).sum()
    calendar_month_swipes.loc[int(week_start.month)] += float(swipes)

peak_calendar_month_swipes = calendar_month_swipes.max()
demand_index = (calendar_month_swipes / peak_calendar_month_swipes).rename("demand_index")
demand_index_df = pd.DataFrame(
    {
        "month": month_labels,
        "swipes": calendar_month_swipes.values,
        "demand_index": demand_index.values,
    },
    index=month_numbers,
)
demand_index_out = Path("meal_swipes_demand_index_by_calendar_month.csv")
demand_index_df.rename_axis("month_number").to_csv(demand_index_out)
print(f"Saved: {demand_index_out.resolve()}")

# Item alignment visual: X = actual meal swipe units, Y = each item's normalized ordering profile.
item_label_map = {
    "222-33756": "Mix Original",
    "2716959": "Alfredo Sauce",
    "6362412": "Oil Canola Olive Virgin 75/25 Blend",
    "9008475": "Sauce Spaghetti Tomato CND SS",
    "9332313": "Black Beans",
    "7335839": "Chipotle Mayo",
    "2554244": "Hummus",
    "6367866": "Canned Peaches",
    "9422965": "Granola Cereal",
    "0052332": "Rice Sushi",
    "8018475": "Hash Browns",
    "4198719": "Ketchup",
    "2167948": "Mango Fruit",
    "7832888": "Kettle Chips",
    "2170200": "Cheese Sauce",
}

alignment_rows = []
unit_regression_rows = []
item_monthly_unit_rows = []
PURCHASE_MONTH_QUANTITY_DIVISOR = 2
item_monthly_quantity = (
    purchases.assign(calendar_month=purchases[date_col].dt.month)
    .groupby([item_id_col, "series_label", "calendar_month"], dropna=False)
    .agg(
        quantity_total=(qty_col, "sum"),
        spending_total=(cost_col, "sum"),
    )
    .reset_index()
)

for (item_id, item_name), group in item_monthly_quantity.groupby([item_id_col, "series_label"], dropna=False):
    monthly_ordering = (
        group.set_index("calendar_month")["quantity_total"]
        .reindex(month_numbers, fill_value=0)
        .astype(float)
        / PURCHASE_MONTH_QUANTITY_DIVISOR
    )
    monthly_spending = (
        group.set_index("calendar_month")["spending_total"]
        .reindex(month_numbers, fill_value=0)
        .astype(float)
        / PURCHASE_MONTH_QUANTITY_DIVISOR
    )
    peak_ordering = monthly_ordering.max()
    if peak_ordering <= 0:
        continue

    y = (monthly_ordering / peak_ordering).values
    x = calendar_month_swipes.values.astype(float)
    regression = linregress(x, y)
    fitted_y = regression.intercept + regression.slope * x
    residual_sigma = float(np.std(y - fitted_y, ddof=1))
    r_value = float(regression.rvalue)

    actual_unit_regression = linregress(x, monthly_ordering.values)
    actual_unit_fitted = actual_unit_regression.intercept + actual_unit_regression.slope * x
    actual_unit_residual_sigma = float(np.std(monthly_ordering.values - actual_unit_fitted, ddof=1))
    quantity_per_1000_swipes = float(actual_unit_regression.slope * 1000)
    meal_swipes_per_unit = (
        float(1 / actual_unit_regression.slope)
        if actual_unit_regression.slope > 0
        else np.nan
    )

    if r_value >= 0.82:
        category = "Good"
    elif r_value >= 0.70:
        category = "Moderate"
    else:
        category = "Poor"

    largest_spending_month_number = int(monthly_spending.idxmax())
    lowest_spending_month_number = int(monthly_spending.idxmin())
    largest_spending_month = month_labels[largest_spending_month_number - 1]
    lowest_spending_month = month_labels[lowest_spending_month_number - 1]
    total_cases = float(monthly_ordering.sum())
    total_spending = float(monthly_spending.sum())
    cost_per_case = total_spending / total_cases if total_cases > 0 else np.nan
    std_error_per_1000_swipes = float(actual_unit_regression.stderr * 1000)
    uncertainty_pct = (
        abs(std_error_per_1000_swipes / quantity_per_1000_swipes) * 100
        if quantity_per_1000_swipes != 0
        else np.nan
    )

    for month_number, month_label, swipes, quantity, quantity_normalized in zip(
        month_numbers,
        month_labels,
        x,
        monthly_ordering.values,
        y,
    ):
        item_monthly_unit_rows.append(
            {
                "item_id": item_id,
                "item_name": item_name,
                "display_name": item_label_map.get(str(item_id), str(item_name)),
                "month_number": month_number,
                "month": month_label,
                "meal_swipes": swipes,
                "quantity": float(quantity),
                "quantity_unit": "cases",
                "quantity_month_adjustment": f"calendar-month purchase total divided by {PURCHASE_MONTH_QUANTITY_DIVISOR}",
                "spending": float(monthly_spending.loc[month_number]),
                "spending_unit": "dollars",
                "item_peak_quantity": float(peak_ordering),
                "quantity_normalized_to_item_peak": float(quantity_normalized),
            }
        )

    unit_regression_rows.append(
        {
            "item_id": item_id,
            "item_name": item_name,
            "display_name": item_label_map.get(str(item_id), str(item_name)),
            "slope_quantity_per_meal_swipe": float(actual_unit_regression.slope),
            "quantity_per_1000_meal_swipes": quantity_per_1000_swipes,
            "meal_swipes_per_1_quantity": meal_swipes_per_unit,
            "intercept_quantity": float(actual_unit_regression.intercept),
            "r": float(actual_unit_regression.rvalue),
            "r_squared": float(actual_unit_regression.rvalue**2),
            "p_value": float(actual_unit_regression.pvalue),
            "std_err": float(actual_unit_regression.stderr),
            "std_error_per_1000_swipes": std_error_per_1000_swipes,
            "uncertainty_pct": uncertainty_pct,
            "residual_sigma_quantity": actual_unit_residual_sigma,
            "item_peak_quantity": float(peak_ordering),
            "cost_per_case": cost_per_case,
            "largest_spending_month": largest_spending_month,
            "largest_spending": float(monthly_spending.loc[largest_spending_month_number]),
            "largest_spending_quantity": float(monthly_ordering.loc[largest_spending_month_number]),
            "lowest_spending_month": lowest_spending_month,
            "lowest_spending": float(monthly_spending.loc[lowest_spending_month_number]),
            "lowest_spending_quantity": float(monthly_ordering.loc[lowest_spending_month_number]),
        }
    )

    alignment_rows.append(
        {
            "item_id": item_id,
            "item_name": item_name,
            "display_name": item_label_map.get(str(item_id), str(item_name)),
            "slope_normalized_quantity_per_meal_swipe": float(regression.slope),
            "normalized_quantity_per_1000_meal_swipes": float(regression.slope * 1000),
            "intercept": float(regression.intercept),
            "r": r_value,
            "r_squared": r_value**2,
            "p_value": float(regression.pvalue),
            "std_err": float(regression.stderr),
            "residual_sigma": residual_sigma,
            "category": category,
            "ordering_profile": y,
        }
    )

unit_regression_df = pd.DataFrame(unit_regression_rows).sort_values("r", ascending=True).reset_index(drop=True)
unit_regression_out = Path("item_meal_swipe_unit_regression.csv")
unit_regression_df.to_csv(unit_regression_out, index=False)
print(f"Saved: {unit_regression_out.resolve()}")

item_monthly_units_df = pd.DataFrame(item_monthly_unit_rows).sort_values(["display_name", "month_number"])
item_monthly_units_out = Path("item_monthly_quantity_by_meal_swipes.csv")
item_monthly_units_df.to_csv(item_monthly_units_out, index=False)
print(f"Saved: {item_monthly_units_out.resolve()}")

alignment_df = pd.DataFrame(alignment_rows).sort_values("r", ascending=True).reset_index(drop=True)
alignment_scores_out = Path("item_alignment_scores.csv")
alignment_df.drop(columns=["ordering_profile"]).to_csv(alignment_scores_out, index=False)
print(f"Saved: {alignment_scores_out.resolve()}")

category_colors = {"Good": "#35b96f", "Moderate": "#ff8c2a", "Poor": "#d94b3d"}
fig, axes = plt.subplots(3, 5, figsize=(22, 15), sharex=True, sharey=True)
fig.patch.set_facecolor("#f3f4f6")
swipe_tick_formatter = FuncFormatter(lambda value, _: f"{value / 1000:,.0f}k")
fig.suptitle(
    "Item Alignment Score - Linear Regression of Ordering vs. Meal Swipe Units\n"
    "Each dot label includes calendar month and meal swipes  |  Y-axis normalized within each item",
    fontsize=18,
    fontweight="bold",
    color="#1f2937",
    y=0.985,
)

for ax, row in zip(axes.flat, alignment_df.itertuples(index=False)):
    color = category_colors[row.category]
    y = row.ordering_profile
    x = calendar_month_swipes.values.astype(float)
    ax.scatter(x, y, s=34, color=color, edgecolor="white", linewidth=0.6, zorder=3)
    fit_x = np.linspace(0, peak_calendar_month_swipes, 100)
    fit_y = row.intercept + row.slope_normalized_quantity_per_meal_swipe * fit_x
    lower_band = np.clip(fit_y - row.residual_sigma, 0, 1.2)
    upper_band = np.clip(fit_y + row.residual_sigma, 0, 1.2)
    ax.fill_between(fit_x, lower_band, upper_band, color=color, alpha=0.14, linewidth=0, zorder=1)
    ax.plot(fit_x, fit_y, color=color, linewidth=2.0, zorder=2)
    ax.plot([0, peak_calendar_month_swipes], [0, 1], color="#d1d5db", linestyle="--", linewidth=1.1, zorder=1)

    for month_label, xi, yi in zip(month_labels, x, y):
        swipe_label = f"{month_label}\n{xi / 1000:,.0f}k"
        ax.text(xi + 5000, yi + 0.008, swipe_label, fontsize=5.4, color="#6b7280")

    ax.text(
        0.04,
        1.14,
        f"r = {row.r:.3f}\nR2 = {row.r_squared:.3f}\n{row.category}",
        ha="left",
        va="top",
        fontsize=8,
        color=color,
        fontweight="bold",
        bbox={"boxstyle": "round,pad=0.25", "facecolor": "white", "edgecolor": color, "linewidth": 1},
        transform=ax.transAxes,
    )
    ax.set_title(row.display_name, fontsize=10, fontweight="bold", color="#374151")
    ax.set_xlim(-0.05 * peak_calendar_month_swipes, 1.1 * peak_calendar_month_swipes)
    ax.set_ylim(-0.05, 1.2)
    ax.grid(True, color="#e5e7eb", linestyle="--", linewidth=0.8, alpha=0.65)
    ax.tick_params(labelsize=8, colors="#4b5563", labelbottom=True)
    ax.xaxis.set_major_formatter(swipe_tick_formatter)
    ax.set_xlabel("Meal swipes (thousands)", fontsize=8, color="#4b5563")
    ax.set_ylabel("Ordering (normalized to item peak)", fontsize=8, color="#4b5563")

for ax in axes.flat[len(alignment_df) :]:
    ax.axis("off")

legend_handles = [
    Line2D([0], [0], marker="s", color="w", markerfacecolor=category_colors["Good"], markersize=10, label="Good (r >= 0.82)"),
    Line2D([0], [0], marker="s", color="w", markerfacecolor=category_colors["Moderate"], markersize=10, label="Moderate (0.70 <= r < 0.82)"),
    Line2D([0], [0], marker="s", color="w", markerfacecolor=category_colors["Poor"], markersize=10, label="Poor (r < 0.70)"),
    Line2D([0], [0], color="#6b7280", linewidth=6, alpha=0.22, label="+/- 1 sigma residual band"),
    Line2D([0], [0], color="#d1d5db", linestyle="--", linewidth=1.4, label="Perfect alignment (diagonal)"),
]
fig.legend(handles=legend_handles, loc="lower center", ncol=5, frameon=False, fontsize=11)
fig.tight_layout(rect=[0.02, 0.055, 0.98, 0.955])

alignment_plot = Path("item_alignment_score_regression.png")
fig.savefig(alignment_plot, dpi=200, bbox_inches="tight")
plt.close(fig)
print(f"Saved: {alignment_plot.resolve()}")

# Interactive product-level unit regression:
# X = meal swipes, Y = actual product quantity, with hoverable regression predictions.
interactive_fig = go.Figure()
interactive_items = unit_regression_df.sort_values("r", ascending=False).reset_index(drop=True)
interactive_x_line = np.linspace(0, peak_calendar_month_swipes, 160)

def _interactive_category(item_id: str) -> str:
    matches = alignment_df.loc[alignment_df["item_id"].astype(str).eq(str(item_id)), "category"]
    return str(matches.iloc[0]) if len(matches) else "Unknown"

def _hex_to_rgba(hex_color: str, alpha: float) -> str:
    color = str(hex_color).lstrip("#")
    if len(color) != 6:
        return f"rgba(107, 114, 128, {alpha})"
    red = int(color[0:2], 16)
    green = int(color[2:4], 16)
    blue = int(color[4:6], 16)
    return f"rgba({red}, {green}, {blue}, {alpha})"

def _stats_annotation(item_row: pd.Series) -> dict:
    category = _interactive_category(str(item_row["item_id"]))
    return {
        "xref": "paper",
        "yref": "paper",
        "x": 0.82,
        "y": 0.72,
        "xanchor": "left",
        "yanchor": "top",
        "align": "left",
        "showarrow": False,
        "bordercolor": "#d1d5db",
        "borderwidth": 1,
        "borderpad": 12,
        "bgcolor": "#f9fafb",
        "font": {"size": 13, "color": "#111827"},
        "text": (
            f"<b>{item_row['display_name']}</b><br>"
            f"Category: <b>{category}</b><br><br>"
            f"r: {item_row['r']:.3f}<br>"
            f"R2: {item_row['r_squared']:.3f}<br>"
            f"p-value: {item_row['p_value']:.4f}<br>"
            f"Std err: {item_row['std_err']:.8f}<br>"
            f"Std err / 1,000 swipes: {item_row['std_error_per_1000_swipes']:,.3f} cases<br>"
            f"Uncertainty: {item_row['uncertainty_pct']:,.1f}% of cases / 1,000 swipes<br>"
            f"1 sigma: {item_row['residual_sigma_quantity']:,.2f} cases<br><br>"
            f"Cases / 1,000 swipes: <b>{item_row['quantity_per_1000_meal_swipes']:,.3f}</b><br>"
            f"Swipes / 1 case: {item_row['meal_swipes_per_1_quantity']:,.0f}<br>"
            f"Intercept cases: {item_row['intercept_quantity']:,.2f}<br>"
            f"Peak monthly cases: {item_row['item_peak_quantity']:,.2f}<br><br>"
            f"Cost per case: <b>${item_row['cost_per_case']:,.2f}</b><br><br>"
            f"Largest spending month: <b>{item_row['largest_spending_month']}</b><br>"
            f"${item_row['largest_spending']:,.2f} | {item_row['largest_spending_quantity']:,.2f} cases<br>"
            f"Lowest spending month: <b>{item_row['lowest_spending_month']}</b><br>"
            f"${item_row['lowest_spending']:,.2f} | {item_row['lowest_spending_quantity']:,.2f} cases"
        ),
    }

metric_summary_annotation = {
    "xref": "paper",
    "yref": "paper",
    "x": 0.82,
    "y": 0.98,
    "xanchor": "left",
    "yanchor": "top",
    "align": "left",
    "showarrow": False,
    "bordercolor": "#d1d5db",
    "borderwidth": 1,
    "borderpad": 12,
    "bgcolor": "#ffffff",
    "font": {"size": 12, "color": "#111827"},
    "text": (
        "<b>Metric guide</b><br>"
        "r: how closely cases and swipes move together.<br>"
        "R2: how much of product purchasing is explained by swipes.<br>"
        "p-value: whether the relationship is likely real.<br>"
        "std error: how precise the estimated cases-per-swipe relationship is."
    ),
}

for item_index, item_row in interactive_items.iterrows():
    product_points = item_monthly_units_df[
        item_monthly_units_df["item_id"].astype(str).eq(str(item_row["item_id"]))
    ].sort_values("month_number")
    visible = item_index == 0
    predicted_quantity_line = (
        item_row["intercept_quantity"]
        + item_row["slope_quantity_per_meal_swipe"] * interactive_x_line
    )
    lower_sigma_line = np.clip(predicted_quantity_line - item_row["residual_sigma_quantity"], 0, None)
    upper_sigma_line = predicted_quantity_line + item_row["residual_sigma_quantity"]
    item_color = category_colors.get(_interactive_category(str(item_row["item_id"])), "#2563eb")
    one_to_one_line = item_row["item_peak_quantity"] * (interactive_x_line / peak_calendar_month_swipes)

    interactive_fig.add_trace(
        go.Scatter(
            x=np.concatenate([interactive_x_line, interactive_x_line[::-1]]),
            y=np.concatenate([upper_sigma_line, lower_sigma_line[::-1]]),
            mode="lines",
            name="+/- 1 sigma band",
            visible=visible,
            line={"color": "rgba(17, 24, 39, 0)"},
            fill="toself",
            fillcolor=_hex_to_rgba(item_color, 0.16),
            hoverinfo="skip",
            showlegend=True,
        )
    )

    interactive_fig.add_trace(
        go.Scatter(
            x=interactive_x_line,
            y=one_to_one_line,
            mode="lines",
            name="1:1 normalized comparison",
            visible=visible,
            line={"color": "rgba(107, 114, 128, 0.45)", "width": 2, "dash": "dash"},
            hovertemplate=(
                "<b>1:1 normalized comparison</b><br>"
                "Meal swipes: %{x:,.0f}<br>"
                "Comparison cases: %{y:,.2f}"
                "<extra></extra>"
            ),
        )
    )

    interactive_fig.add_trace(
        go.Scatter(
            x=product_points["meal_swipes"],
            y=product_points["quantity"],
            mode="markers+text",
            name=f"{item_row['display_name']} monthly points",
            visible=visible,
            marker={
                "size": 10,
                "color": item_color,
                "line": {"color": "white", "width": 1},
            },
            text=product_points["month"],
            textposition="top center",
            customdata=np.column_stack(
                [
                    product_points["month"],
                    product_points["quantity_normalized_to_item_peak"],
                    product_points["item_peak_quantity"],
                ]
            ),
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Meal swipes: %{x:,.0f}<br>"
                "Actual cases: %{y:,.2f}<br>"
                "Normalized to item peak: %{customdata[1]:.3f}<br>"
                "Item peak cases: %{customdata[2]:,.2f}"
                "<extra></extra>"
            ),
        )
    )

    interactive_fig.add_trace(
        go.Scatter(
            x=interactive_x_line,
            y=predicted_quantity_line,
            mode="lines",
            name=f"{item_row['display_name']} regression line",
            visible=visible,
            line={"color": "#111827", "width": 3},
            hovertemplate=(
                "<b>Regression prediction</b><br>"
                "Meal swipes: %{x:,.0f}<br>"
                "Predicted cases: %{y:,.2f}<br>"
                f"Cases per 1,000 swipes: {item_row['quantity_per_1000_meal_swipes']:,.3f}<br>"
                f"r: {item_row['r']:.3f} | R2: {item_row['r_squared']:.3f}"
                "<extra></extra>"
            ),
        )
    )

dropdown_buttons = []
trace_count = len(interactive_fig.data)
for item_index, item_row in interactive_items.iterrows():
    visible = [False] * trace_count
    visible[item_index * 4] = True
    visible[item_index * 4 + 1] = True
    visible[item_index * 4 + 2] = True
    visible[item_index * 4 + 3] = True
    y_max = max(
        float(
            item_monthly_units_df[
                item_monthly_units_df["item_id"].astype(str).eq(str(item_row["item_id"]))
            ]["quantity"].max()
        ),
        float(
            item_row["intercept_quantity"]
            + item_row["slope_quantity_per_meal_swipe"] * peak_calendar_month_swipes
            + item_row["residual_sigma_quantity"]
        ),
    )
    dropdown_buttons.append(
        {
            "label": str(item_row["display_name"]),
            "method": "update",
            "args": [
                {"visible": visible},
                {
                    "title": {
                        "text": (
                            f"{item_row['display_name']}: Quantity vs Meal Swipes"
                            f"<br><sup>Hover points for actual monthly cases; hover the line for predicted cases.</sup>"
                        ),
                        "x": 0.37,
                        "y": 0.96,
                        "xanchor": "center",
                        "yanchor": "top",
                    },
                    "yaxis": {
                        "range": [0, y_max * 1.18],
                        "title": {
                            "text": "Product quantity purchased (cases)",
                            "standoff": 24,
                        },
                        "automargin": True,
                    },
                    "annotations": [metric_summary_annotation, _stats_annotation(item_row)],
                },
            ],
        }
    )

first_item = interactive_items.iloc[0]
first_item_points = item_monthly_units_df[
    item_monthly_units_df["item_id"].astype(str).eq(str(first_item["item_id"]))
]
first_y_max = max(
    float(first_item_points["quantity"].max()),
    float(
        first_item["intercept_quantity"]
        + first_item["slope_quantity_per_meal_swipe"] * peak_calendar_month_swipes
        + first_item["residual_sigma_quantity"]
    ),
)

interactive_fig.update_layout(
    title={
        "text": (
            f"{first_item['display_name']}: Quantity vs Meal Swipes"
            "<br><sup>Hover points for actual monthly cases; hover the line for predicted cases.</sup>"
        ),
        "x": 0.37,
        "y": 0.96,
        "xanchor": "center",
        "yanchor": "top",
    },
    template="plotly_white",
    width=1350,
    height=720,
    hovermode="closest",
    xaxis={
        "title": "Meal swipes",
        "domain": [0, 0.64],
        "range": [-0.05 * peak_calendar_month_swipes, 1.08 * peak_calendar_month_swipes],
        "tickformat": ",.0f",
        "showspikes": True,
        "spikemode": "across",
        "spikesnap": "cursor",
        "spikethickness": 1,
    },
    yaxis={
        "title": {
            "text": "Product quantity purchased (cases)",
            "standoff": 24,
        },
        "range": [0, first_y_max * 1.18],
        "showspikes": True,
        "spikemode": "across",
        "spikesnap": "cursor",
        "spikethickness": 1,
        "automargin": True,
    },
    updatemenus=[
        {
            "buttons": dropdown_buttons,
            "direction": "down",
            "showactive": True,
            "x": 0,
            "xanchor": "left",
            "y": 1.08,
            "yanchor": "top",
            "pad": {"t": 8, "b": 8},
        }
    ],
    annotations=[metric_summary_annotation, _stats_annotation(first_item)],
    legend={"orientation": "h", "y": -0.18},
    margin={"l": 135, "r": 430, "t": 170, "b": 95},
)

interactive_alignment_plot = Path("item_meal_swipe_interactive_regression.html")
interactive_fig.write_html(interactive_alignment_plot, include_plotlyjs=True, full_html=True)
print(f"Saved: {interactive_alignment_plot.resolve()}")

# Headcount regression diagnostic + seasonal meal swipe forecast into 2025.
# Forecast rule: copy the latest actual year's monthly pattern forward and grow it 4% per year.
MEAL_SWIPES_FORECAST_GROWTH_RATE = 0.04

headcount_sheet_name = "Headcount"
try:
    headcount_df = pd.read_excel(file_path, sheet_name=headcount_sheet_name)
except Exception as e:
    raise RuntimeError(f"Failed to read headcount sheet {headcount_sheet_name!r} from {file_path}.") from e

headcount_df.columns = headcount_df.columns.astype("string").str.strip()
if "Year" not in headcount_df.columns or "Headcount" not in headcount_df.columns:
    raise RuntimeError(
        f"Headcount sheet must contain 'Year' and 'Headcount' columns. Found: {headcount_df.columns.tolist()!r}"
    )

headcount_df = headcount_df[["Year", "Headcount"]].copy()
headcount_df["Year"] = pd.to_numeric(headcount_df["Year"], errors="coerce").astype("Int64")
headcount_df["Headcount"] = pd.to_numeric(headcount_df["Headcount"], errors="coerce")
headcount_df = headcount_df.dropna(subset=["Year", "Headcount"]).astype({"Year": "int64"})

headcount_by_year = headcount_df.drop_duplicates(subset=["Year"], keep="last").set_index("Year")["Headcount"].sort_index()

# Regress: monthly_swipes (dependent) ~ headcount (independent) for overlapping years
swipes_with_year = monthly_swipes.to_frame(name="meal_swipes").copy()
swipes_with_year["Year"] = swipes_with_year.index.year
swipes_with_year["Headcount"] = swipes_with_year["Year"].map(headcount_by_year)
swipes_with_year = swipes_with_year.dropna(subset=["meal_swipes", "Headcount"])

if len(swipes_with_year) >= 3:
    X_hc = sm.add_constant(swipes_with_year["Headcount"])
    y_sw = swipes_with_year["meal_swipes"]
    model_swipes_hc = sm.OLS(y_sw, X_hc).fit()

    swipes_hc_out = Path("meal_swipes_vs_headcount_regression.txt")
    swipes_hc_out.write_text(model_swipes_hc.summary().as_text(), encoding="utf-8")
    print(f"Saved: {swipes_hc_out.resolve()}")

    latest_headcount_year = int(headcount_by_year.index.max())
    baseline_headcount = float(headcount_by_year.loc[latest_headcount_year])
    increased_headcount = baseline_headcount * (1 + MEAL_SWIPES_FORECAST_GROWTH_RATE)
    headcount_increase = increased_headcount - baseline_headcount
    intercept = float(model_swipes_hc.params.get("const", 0.0))
    meal_swipes_per_person_per_month = float(model_swipes_hc.params.get("Headcount", 0.0))
    monthly_swipe_increase = meal_swipes_per_person_per_month * headcount_increase
    annual_swipe_increase = monthly_swipe_increase * 12
    baseline_predicted_monthly_swipes = intercept + (meal_swipes_per_person_per_month * baseline_headcount)
    increased_predicted_monthly_swipes = intercept + (meal_swipes_per_person_per_month * increased_headcount)

    headcount_impact_out = Path("meal_swipes_headcount_4pct_impact.txt")
    headcount_impact_out.write_text(
        "\n".join(
            [
                "Estimated meal swipe impact of a 4% headcount increase",
                "======================================================",
                "",
                f"Method: OLS regression from meal_swipes_vs_headcount_regression.txt",
                f"Regression coefficient: {meal_swipes_per_person_per_month:,.4f} meal swipes per additional person per month",
                f"Regression R-squared: {model_swipes_hc.rsquared:.4f}",
                "",
                f"Baseline headcount year: {latest_headcount_year}",
                f"Baseline headcount: {baseline_headcount:,.0f}",
                f"Headcount after 4% increase: {increased_headcount:,.2f}",
                f"Additional people: {headcount_increase:,.2f}",
                "",
                f"Predicted monthly swipes at baseline headcount: {baseline_predicted_monthly_swipes:,.2f}",
                f"Predicted monthly swipes after 4% increase: {increased_predicted_monthly_swipes:,.2f}",
                f"Estimated additional meal swipes per month: {monthly_swipe_increase:,.2f}",
                f"Estimated additional meal swipes per year: {annual_swipe_increase:,.2f}",
                "",
                "Interpretation note: this uses the old headcount regression relationship only.",
                "The form-fitting forecast file copies the latest actual monthly pattern forward and grows each point by 4%.",
            ]
        )
        + "\n",
        encoding="utf-8",
    )
    print(f"Saved: {headcount_impact_out.resolve()}")

    plt.figure(figsize=(8, 6))
    sns.regplot(
        data=swipes_with_year,
        x="Headcount",
        y="meal_swipes",
        scatter_kws={"s": 45, "alpha": 0.8},
        line_kws={"linewidth": 2},
    )
    plt.title("Monthly meal swipes vs headcount (OLS fit)")
    plt.xlabel("Headcount (yearly, mapped to each month)")
    plt.ylabel("Meal swipes per month")
    plt.tight_layout()
    swipes_hc_plot = Path("meal_swipes_vs_headcount.png")
    plt.savefig(swipes_hc_plot, dpi=200, bbox_inches="tight")
    plt.close()
    print(f"Saved: {swipes_hc_plot.resolve()}")
else:
    model_swipes_hc = None
    swipes_hc_note = Path("meal_swipes_vs_headcount_NOTE.txt")
    swipes_hc_note.write_text(
        "Not enough overlapping months between meal swipes and headcount to run regression.\n",
        encoding="utf-8",
    )
    print(f"Saved: {swipes_hc_note.resolve()}")
    headcount_impact_out = Path("meal_swipes_headcount_4pct_impact.txt")
    headcount_impact_out.write_text(
        "Not enough overlapping months between meal swipes and headcount to estimate the impact of a 4% headcount increase.\n",
        encoding="utf-8",
    )
    print(f"Saved: {headcount_impact_out.resolve()}")

# Forecast swipes into future years by copying actual monthly points forward.
forecast_end_year = 2025
latest_actual_year = int(monthly_swipes.index.year.max())
latest_actual_year_swipes = monthly_swipes[monthly_swipes.index.year == latest_actual_year].copy()

forecast_rows = []
for target_year in range(latest_actual_year + 1, forecast_end_year + 1):
    years_ahead = target_year - latest_actual_year
    growth_multiplier = (1 + MEAL_SWIPES_FORECAST_GROWTH_RATE) ** years_ahead
    for source_month_end, source_swipes in latest_actual_year_swipes.items():
        target_month_end = source_month_end + pd.DateOffset(years=years_ahead)
        target_month_end = target_month_end + pd.offsets.MonthEnd(0)
        forecast_swipes = source_swipes * growth_multiplier
        forecast_rows.append(
            {
                "month_end": target_month_end,
                "source_month_end": source_month_end,
                "actual_swipes_source": source_swipes,
                "growth_multiplier": growth_multiplier,
                "forecast_meal_swipes": forecast_swipes,
                "forecast_swipes_increase_from_source": forecast_swipes - source_swipes,
            }
        )

forecast_df = pd.DataFrame(forecast_rows).set_index("month_end").sort_index()
forecast_swipes_monthly = forecast_df["forecast_meal_swipes"].rename("forecast_meal_swipes")

forecast_out = Path("meal_swipes_forecast_2024_2025.csv")
forecast_df.to_csv(forecast_out)
print(f"Saved: {forecast_out.resolve()}")

forecast_yearly_summary = (
    forecast_df.assign(year=forecast_df.index.year)
    .groupby("year")
    .agg(
        actual_swipes_source_total=("actual_swipes_source", "sum"),
        forecast_meal_swipes_total=("forecast_meal_swipes", "sum"),
        forecast_swipes_increase_from_source=("forecast_swipes_increase_from_source", "sum"),
    )
)
forecast_yearly_summary["growth_multiplier"] = (
    forecast_yearly_summary["forecast_meal_swipes_total"]
    / forecast_yearly_summary["actual_swipes_source_total"].replace(0, np.nan)
)

forecast_yearly_summary_out = Path("meal_swipes_forecast_yearly_summary.csv")
forecast_yearly_summary.to_csv(forecast_yearly_summary_out)
print(f"Saved: {forecast_yearly_summary_out.resolve()}")

plt.figure(figsize=(11, 5))
plt.plot(monthly_swipes.index, monthly_swipes.values, linewidth=2, marker="o", markersize=4, label="Actual (from meal plan)")
if forecast_swipes_monthly.notna().any():
    plt.plot(forecast_swipes_monthly.index, forecast_swipes_monthly.values, linewidth=2, marker="o", markersize=4, label="Forecast (actual pattern + 4%)")
plt.title("Meal swipes per month: actual + seasonal forecast")
plt.xlabel("Month")
plt.ylabel("Meal swipes")
plt.legend()
plt.tight_layout()
swipes_forecast_plot = Path("meal_swipes_actual_vs_forecast.png")
plt.savefig(swipes_forecast_plot, dpi=200, bbox_inches="tight")
plt.close()
print(f"Saved: {swipes_forecast_plot.resolve()}")

# Linear regression: monthly meal swipes vs monthly quantity purchased
monthly_quantity = (
    purchases[[date_col, qty_col]]
    .dropna(subset=[date_col, qty_col])
    .set_index(date_col)[qty_col]
    .sort_index()
    .resample("ME")
    .sum()
)

reg_df = (
    pd.concat(
        [
            forecast_swipes_monthly.rename("meal_swipes"),
            monthly_quantity.rename("total_quantity"),
        ],
        axis=1,
    )
    .dropna()
)

merged_monthly_path = Path("meal_swipes_and_quantity_by_month.csv")
reg_df.to_csv(merged_monthly_path, index=True)
print(f"Saved: {merged_monthly_path.resolve()}")

if len(reg_df) < 3:
    note = (
        "Not enough overlapping months to run a linear regression.\n"
        f"Overlapping months found: {len(reg_df)}.\n\n"
        "This usually means your meal plan data and purchase data cover different date ranges.\n"
        "Check `meal_swipes_per_month.csv` vs `weekly_total_cost.csv`/purchase dates.\n"
    )
    no_reg_path = Path("meal_swipes_vs_quantity_regression_NOTE.txt")
    no_reg_path.write_text(note, encoding="utf-8")
    print(f"Saved: {no_reg_path.resolve()}")
else:
    reg_df["meal_swipes"] = pd.to_numeric(reg_df["meal_swipes"], errors="coerce")
    reg_df["total_quantity"] = pd.to_numeric(reg_df["total_quantity"], errors="coerce")
    reg_df = reg_df.dropna(subset=["meal_swipes", "total_quantity"])

    X = sm.add_constant(reg_df["meal_swipes"].astype(float))
    y = reg_df["total_quantity"].astype(float)
    model = sm.OLS(y, X).fit()

    reg_out = Path("meal_swipes_vs_quantity_regression.txt")
    reg_out.write_text(model.summary().as_text(), encoding="utf-8")
    print(f"Saved: {reg_out.resolve()}")

    influence = model.get_influence()
    diagnostics_df = reg_df.copy()
    diagnostics_df["predicted_quantity"] = model.predict(X)
    diagnostics_df["residual"] = diagnostics_df["total_quantity"] - diagnostics_df["predicted_quantity"]
    diagnostics_df["studentized_residual"] = influence.resid_studentized_external
    diagnostics_df["cooks_distance"] = influence.cooks_distance[0]
    diagnostics_df["leverage"] = influence.hat_matrix_diag
    diagnostics_df["is_zero_swipe_month"] = diagnostics_df["meal_swipes"].eq(0)

    diagnostics_out = Path("meal_swipes_vs_quantity_regression_diagnostics.csv")
    diagnostics_df.to_csv(diagnostics_out, index=True)
    print(f"Saved: {diagnostics_out.resolve()}")

    nonzero_swipe_reg_df = reg_df[reg_df["meal_swipes"] > 0].copy()
    nonzero_reg_out = Path("meal_swipes_vs_quantity_regression_nonzero_swipes.txt")
    if len(nonzero_swipe_reg_df) >= 3:
        X_nonzero = sm.add_constant(nonzero_swipe_reg_df["meal_swipes"].astype(float))
        y_nonzero = nonzero_swipe_reg_df["total_quantity"].astype(float)
        model_nonzero = sm.OLS(y_nonzero, X_nonzero).fit()
        nonzero_reg_out.write_text(model_nonzero.summary().as_text(), encoding="utf-8")
    else:
        nonzero_reg_out.write_text(
            "Not enough nonzero meal swipe months to run the filtered regression.\n",
            encoding="utf-8",
        )
    print(f"Saved: {nonzero_reg_out.resolve()}")

    visual_df = diagnostics_df.reset_index().rename(columns={"index": "month_end"})
    visual_df["month_label"] = visual_df["month_end"].dt.strftime("%b %Y")
    largest_residuals = visual_df.reindex(visual_df["studentized_residual"].abs().sort_values(ascending=False).head(5).index)
    x_line = np.linspace(visual_df["meal_swipes"].min(), visual_df["meal_swipes"].max(), 100)
    y_line = float(model.params["const"]) + float(model.params["meal_swipes"]) * x_line

    plt.figure(figsize=(11, 7))
    sns.scatterplot(
        data=visual_df,
        x="meal_swipes",
        y="total_quantity",
        hue="is_zero_swipe_month",
        style="is_zero_swipe_month",
        palette={False: "#2563eb", True: "#dc2626"},
        markers={False: "o", True: "X"},
        s=95,
        edgecolor="white",
        linewidth=0.8,
    )
    plt.plot(x_line, y_line, color="#111827", linewidth=2, label="OLS fit")
    plt.axvline(0, color="#dc2626", linewidth=1.2, linestyle="--", alpha=0.65)

    for i, row in enumerate(largest_residuals.itertuples(index=False)):
        x_offset = 7000 if i % 2 == 0 else -42000
        y_offset = 22 if row.residual >= 0 else -34
        plt.annotate(
            row.month_label,
            xy=(row.meal_swipes, row.total_quantity),
            xytext=(row.meal_swipes + x_offset, row.total_quantity + y_offset),
            arrowprops={"arrowstyle": "->", "color": "#374151", "lw": 0.9},
            fontsize=9,
            color="#111827",
        )

    plt.title("Monthly quantity purchased vs meal swipes: outlier diagnostic")
    plt.xlabel("Forecasted meal swipes per month")
    plt.ylabel("Total quantity purchased per month")
    legend_handles = [
        Line2D([0], [0], color="#111827", lw=2, label="OLS fit"),
        Line2D([0], [0], marker="o", color="w", markerfacecolor="#2563eb", markeredgecolor="white", markersize=9, label="Nonzero-swipe month"),
        Line2D([0], [0], marker="X", color="w", markerfacecolor="#dc2626", markeredgecolor="white", markersize=9, label="Zero-swipe month"),
    ]
    plt.legend(handles=legend_handles, loc="upper left", frameon=True)
    plt.tight_layout()

    diagnostic_plot = Path("meal_swipes_vs_quantity_outlier_diagnostic.png")
    plt.savefig(diagnostic_plot, dpi=200, bbox_inches="tight")
    plt.close()
    print(f"Saved: {diagnostic_plot.resolve()}")

    plt.figure(figsize=(8, 6))
    sns.regplot(
        data=reg_df.reset_index(),
        x="meal_swipes",
        y="total_quantity",
        scatter_kws={"s": 45, "alpha": 0.8},
        line_kws={"linewidth": 2},
    )
    plt.title("Monthly quantity purchased vs meal swipes (OLS fit)")
    plt.xlabel("Total meal swipes per month")
    plt.ylabel("Total quantity purchased per month")
    plt.tight_layout()

    reg_plot = Path("meal_swipes_vs_quantity.png")
    plt.savefig(reg_plot, dpi=200, bbox_inches="tight")
    plt.close()
    print(f"Saved: {reg_plot.resolve()}")

# Quantity and total cost plotted separately (same x-axis time), one line per item
fig, (ax1, ax2) = plt.subplots(nrows=2, ncols=1, figsize=(12, 9), sharex=True)

sns.lineplot(
    data=purchases,
    x=date_col,
    y=qty_col,
    hue="series_label",
    marker="o",
    dashes=False,
    linewidth=1.4,
    markersize=4,
    legend="full",
    ax=ax1,
)
ax1.set_title("Quantity purchased over time (by item)")
ax1.set_xlabel("")
ax1.set_ylabel("Quantity")

sns.lineplot(
    data=purchases,
    x=date_col,
    y=cost_col,
    hue="series_label",
    marker="o",
    dashes=False,
    linewidth=1.4,
    markersize=4,
    legend=False,  # keep a single legend (top plot)
    ax=ax2,
)
ax2.set_title("Total cost per purchase over time (by item)")
ax2.set_xlabel("Purchase date")
ax2.set_ylabel("Total cost")

# Put the legend on the top subplot
ax1.legend(title="Item", bbox_to_anchor=(1.02, 1), loc="upper left", borderaxespad=0, fontsize=8)
fig.tight_layout()

qty_cost_over_time_path = Path("quantity_and_total_cost_over_time.png")
fig.savefig(qty_cost_over_time_path, dpi=200, bbox_inches="tight")
plt.close(fig)
print(f"Saved: {qty_cost_over_time_path.resolve()}")

# look at the first rows
print(df.head())

# inspect columns
print(df.columns.tolist())
print(df.info())
