import pandas as pd
import re

IN_PATH = "concatenated_player_data.xlsx"
OUT_PATH = "filtered_overall_within_2.xlsx"

df = pd.read_excel(IN_PATH)

# Columns you want to keep
keep_cols = [
    "full_name",
    "club_name_before",
    "club_name_after",
    "player_id",
    "club_league_name",
    "overall_rating",
    "overall",
    "league_name",
]

# --- helper: extract the last integer from a cell (handles "84 | 85") ---
def extract_last_int(x):
    if pd.isna(x):
        return None
    s = str(x)
    nums = re.findall(r"-?\d+", s)
    return int(nums[-1]) if nums else None

# Create numeric versions for comparison
df["overall_rating_num"] = df["overall_rating"].apply(extract_last_int)
df["overall_num"] = df["overall"].apply(extract_last_int)

# Filter: within +/- 2 (and both numbers present)
filtered = df[
    df["overall_rating_num"].notna()
    & df["overall_num"].notna()
    & ((df["overall_rating_num"] - df["overall_num"]).abs() <= 2)
].copy()

# Keep only requested columns that actually exist in the file
existing_keep_cols = [c for c in keep_cols if c in filtered.columns]
missing = [c for c in keep_cols if c not in filtered.columns]
if missing:
    print("Warning: these columns were not found and will be skipped:", missing)

result = filtered[existing_keep_cols].copy()

# Save
result.to_excel(OUT_PATH, index=False)
print(f"Saved {len(result)} rows to: {OUT_PATH}")
