import pandas as pd
import re

file = "A.5_Non-Perishables.xlsx"
df = pd.read_excel(file, sheet_name="Data ")
df.columns = df.columns.str.strip()

# clean numeric columns
for col in ["Quantity", "Price (Jan. Avg)", "Total"]:
    df[col] = (
        df[col]
        .astype(str)
        .str.replace("$", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.strip()
    )
    df[col] = pd.to_numeric(df[col], errors="coerce")

df["Unit"] = df["Unit"].astype(str).str.strip()

def parse_unit(unit_str):
    """
    Parse strings like:
    - Case (4/80 oz)
    - Case (8/3.75 lb)
    - Case (6/1 gal)
    - Case (50 lb)
    - Case (6/#10)
    """
    result = {
        "case_count": None,          # number of inner packs per case
        "pack_size": None,           # amount per inner pack
        "measure_unit": None,        # oz, lb, gal, #10, etc
        "total_oz_per_case": None,
        "total_lb_per_case": None,
        "total_gal_per_case": None,
        "total_each_per_case": None, # for #10 cans or count-based units
    }

    if pd.isna(unit_str):
        return result

    text = str(unit_str).strip()

    # pull out text inside parentheses if present
    m = re.search(r"\((.*?)\)", text)
    inside = m.group(1).strip() if m else text

    # pattern: 4/80 oz, 8/3.75 lb, 6/1 gal, 6/#10
    m = re.match(r"^\s*(\d+)\s*/\s*([#]?\d*\.?\d+)\s*([A-Za-z#0-9]+)?\s*$", inside)
    if m:
        case_count = float(m.group(1))
        pack_size_raw = m.group(2)
        unit = m.group(3)

        result["case_count"] = case_count

        if pack_size_raw.startswith("#"):
            # count-style can notation like #10
            result["pack_size"] = float(pack_size_raw.replace("#", ""))
            result["measure_unit"] = "#can"
            result["total_each_per_case"] = case_count
        else:
            pack_size = float(pack_size_raw)
            result["pack_size"] = pack_size
            result["measure_unit"] = unit.lower() if unit else None

            if result["measure_unit"] == "oz":
                result["total_oz_per_case"] = case_count * pack_size
            elif result["measure_unit"] == "lb":
                result["total_lb_per_case"] = case_count * pack_size
                result["total_oz_per_case"] = case_count * pack_size * 16
            elif result["measure_unit"] == "gal":
                result["total_gal_per_case"] = case_count * pack_size

        return result

    # pattern: 50 lb
    m = re.match(r"^\s*(\d*\.?\d+)\s*([A-Za-z]+)\s*$", inside)
    if m:
        size = float(m.group(1))
        unit = m.group(2).lower()

        result["case_count"] = 1
        result["pack_size"] = size
        result["measure_unit"] = unit

        if unit == "oz":
            result["total_oz_per_case"] = size
        elif unit == "lb":
            result["total_lb_per_case"] = size
            result["total_oz_per_case"] = size * 16
        elif unit == "gal":
            result["total_gal_per_case"] = size

        return result

    return result

parsed = df["Unit"].apply(parse_unit).apply(pd.Series)
df = pd.concat([df, parsed], axis=1)

# pricing metrics
df["cost_per_case"] = df["Total"] / df["Quantity"]              # same as price per purchased case
df["cost_per_inner_pack"] = df["cost_per_case"] / df["case_count"]

df["cost_per_oz"] = df["cost_per_case"] / df["total_oz_per_case"]
df["cost_per_lb"] = df["cost_per_case"] / df["total_lb_per_case"]
df["cost_per_gal"] = df["cost_per_case"] / df["total_gal_per_case"]
df["cost_per_each"] = df["cost_per_case"] / df["total_each_per_case"]

# inspect result
cols = [
    "Period End", "Item Name", "Unit", "Quantity", "Price (Jan. Avg)", "Total",
    "case_count", "pack_size", "measure_unit",
    "cost_per_case", "cost_per_inner_pack",
    "cost_per_oz", "cost_per_lb", "cost_per_gal", "cost_per_each"
]
print(df[cols].head(20))

# cheapest cost per lb items
print(
    df[df["cost_per_lb"].notna()]
    .sort_values("cost_per_lb")[["Item Name", "Unit", "cost_per_case", "cost_per_lb"]]
    .head(20)
)

# cheapest cost per oz items
print(
    df[df["cost_per_oz"].notna()]
    .sort_values("cost_per_oz")[["Item Name", "Unit", "cost_per_case", "cost_per_oz"]]
    .head(20)
)

# cheapest cost per gallon items
print(
    df[df["cost_per_gal"].notna()]
    .sort_values("cost_per_gal")[["Item Name", "Unit", "cost_per_case", "cost_per_gal"]]
    .head(20)
)