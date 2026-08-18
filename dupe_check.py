import pandas as pd

df = pd.read_excel("A.5_Non-Perishables.xlsx", sheet_name="Data ")
df.columns = df.columns.str.strip()

# exact duplicate rows across all columns
exact_dupes = df[df.duplicated(keep=False)]
print("Exact duplicate rows:", len(exact_dupes))
print(exact_dupes)

# duplicates ignoring Period End
dupes_ignoring_date = df[df.drop(columns=["Period End"]).duplicated(keep=False)]
print("Duplicates ignoring Period End:", len(dupes_ignoring_date))
print(dupes_ignoring_date.sort_values(["Item ID", "Period End"]))