import pandas as pd
import numpy as np
import statsmodels.api as sm

INPUT_XLSX = "player_club_prices_log_steps.xlsx"

# Columns to regress (X) against y
X_COLS = [
    "diff_Altitude(m)",
    "diff_Possesion",
    "diff_Win %",
    "diff_annual_avg_temp_c",
    "diff_annual_avg_humidity_pct",
    "diff_annual_avg_pressure_hpa",
    "diff_annual_avg_cloudcover_pct",
    "diff_annual_total_precip_mm",
]
Y_COL = "ln_before_minus_ln_after"


def to_numeric_series(s: pd.Series) -> pd.Series:
    """Convert messy numeric formats to float."""
    if pd.api.types.is_numeric_dtype(s):
        return s.astype(float)
    return (
        s.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("%", "", regex=False)
        .str.strip()
        .replace({"nan": np.nan, "None": np.nan, "": np.nan})
        .pipe(pd.to_numeric, errors="coerce")
        .astype(float)
    )


def main():
    df = pd.read_excel(INPUT_XLSX)

    missing = [c for c in X_COLS + [Y_COL] if c not in df.columns]
    if missing:
        raise KeyError(f"Missing columns: {missing}\nAvailable columns: {list(df.columns)}")

    # Convert all relevant columns to numeric
    for c in X_COLS + [Y_COL]:
        df[c] = to_numeric_series(df[c])

    # Drop rows with NaNs
    model_df = df[X_COLS + [Y_COL]].dropna()

    if len(model_df) < 5:
        raise ValueError(f"Not enough usable rows after dropping NaNs: {len(model_df)}")

    y = model_df[Y_COL]
    X = sm.add_constant(model_df[X_COLS], has_constant="add")

    # Fit OLS
    model = sm.OLS(y, X).fit(cov_type="HC3")

    # ---- Explicit R² reporting ----
    print("\n=== Model Fit Statistics ===")
    print(f"R-squared:          {model.rsquared:.6f}")
    print(f"Adjusted R-squared: {model.rsquared_adj:.6f}")
    print(f"Observations:       {int(model.nobs)}")

    # ---- Full regression output ----
    print("\n=== OLS Regression Results ===")
    print(model.summary())

    # ---- Compact coefficient table ----
    coef_table = pd.DataFrame({
        "coef": model.params,
        "std_err": model.bse,
        "t_stat": model.tvalues,
        "p_value": model.pvalues,
    })

    print("\n=== Coefficients (compact) ===")
    print(coef_table)


if __name__ == "__main__":
    main()
