import pandas as pd
import numpy as np

INPUT_XLSX = "player_club_prices_raw.xlsx"
OUTPUT_XLSX = "player_club_prices_log_steps.xlsx"

# ---- helpers ----
def to_num(s):
    """Convert strings like '12,345' or '12 345' to numbers safely."""
    if pd.api.types.is_numeric_dtype(s):
        return s
    return (
        s.astype(str)
         .str.replace(",", "", regex=False)
         .str.replace(" ", "", regex=False)
         .replace({"nan": np.nan, "None": np.nan, "": np.nan})
         .pipe(pd.to_numeric, errors="coerce")
    )

def safe_ln(x):
    """Natural log, but only for strictly positive values; else NaN."""
    x = pd.to_numeric(x, errors="coerce")
    return np.where(x > 0, np.log(x), np.nan)

# ---- main ----
df = pd.read_excel(INPUT_XLSX)

# Try to auto-detect the "price before" and "price after" columns.
# If your sheet uses different names, edit these two variables.
before_col_candidates = ["price_before", "before_price", "price_before_ps", "ps_price_before", "price_before_coins"]
after_col_candidates  = ["price_after",  "after_price",  "price_after_ps",  "ps_price_after",  "price_after_coins"]

def pick_col(cands):
    for c in cands:
        if c in df.columns:
            return c
    return None

before_col = pick_col(before_col_candidates)
after_col = pick_col(after_col_candidates)

if before_col is None or after_col is None:
    raise KeyError(
        "Couldn't find the before/after price columns.\n"
        f"Columns in file: {list(df.columns)}\n"
        "Edit before_col_candidates / after_col_candidates in the script to match your headers."
    )

# Ensure numeric
df[before_col] = to_num(df[before_col])
df[after_col] = to_num(df[after_col])

# Step 1: ln(price_before)
df["ln_price_before"] = safe_ln(df[before_col])

# Step 2: ln(price_after)
df["ln_price_after"] = safe_ln(df[after_col])

# Step 3: subtract after from before (ln_before - ln_after)
df["ln_before_minus_ln_after"] = df["ln_price_before"] - df["ln_price_after"]

# Step 4: ln(of the final result)  -> ln( ln_before_minus_ln_after )
# NOTE: This is only defined when (ln_before_minus_ln_after) > 0; otherwise NaN.
df["ln_of_ln_diff"] = safe_ln(df["ln_before_minus_ln_after"])

# Save
df.to_excel(OUTPUT_XLSX, index=False)
print(f"Saved: {OUTPUT_XLSX}")