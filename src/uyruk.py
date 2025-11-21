import pandas as pd
import warnings
import difflib
warnings.filterwarnings('ignore')

FILE_PATH = "input-data/uyruk/uyruk.xlsx"
OUTPUT_PATH = "output-data/uyruk_perfomans.xlsx"
UNIQUE_PAIRS_PATH = "./const/unique_region_country.xlsx"

print("🚀 Парсинг: Month → Country → Agency → Region (region from unique_region_country.xlsx)")

# ------------------------------
# 0. LOAD REGION–COUNTRY DATASET
# ------------------------------
pairs = pd.read_excel(UNIQUE_PAIRS_PATH)

pairs.columns = pairs.columns.astype(str).str.strip().str.title()
pairs['Country_norm'] = pairs['Country'].astype(str).str.upper().str.strip()
pairs['Country_title'] = pairs['Country'].astype(str).str.title().str.strip()
pairs['Region_title'] = pairs['Region'].astype(str).str.title().str.strip()

country_to_region = dict(zip(pairs['Country_norm'], pairs['Region_title']))

country_list_upper = list(country_to_region.keys())
region_list_title = set(pairs['Region_title'].unique())

print(f"✔ Эталон: {len(country_list_upper)} стран → {len(region_list_title)} регионов")

# ------------------------------
# 1. READ RAW EXCEL
# ------------------------------
raw = pd.read_excel(FILE_PATH, header=None)
header_row = raw[raw.iloc[:, 0].astype(str).str.contains("AgencyGroup", na=False)].index[0]
df = pd.read_excel(FILE_PATH, header=header_row)

df = df.dropna(how='all').dropna(axis=1, how='all').reset_index(drop=True)

df.columns = (
    df.columns.astype(str)
    .str.strip()
    .str.replace(r'\s+', ' ', regex=True)
    .str.replace(' ', '_')
    .str.lower()
)

# ------------------------------
# 2. MONTH DETECT
# ------------------------------
df['raw'] = df['agencygroup'].astype(str).str.strip()
month_mask = df['raw'].str.match(r'^\d{2}-', na=False)

df['Month'] = pd.NA
cur_month = None

for i in df.index:
    if month_mask[i]:
        cur_month = df.at[i, 'raw'].split('-', 1)[1].strip().title()
    df.at[i, 'Month'] = cur_month

df = df[~month_mask].reset_index(drop=True)

# ------------------------------
# 3. PARSE COUNTRY / AGENCY
# ------------------------------
df['Country'] = pd.NA
df['Agency'] = pd.NA
df['Region'] = pd.NA

cur_country = None
cur_region = None

for i in df.index:
    raw_val = str(df.at[i, 'raw']).strip()
    upper = raw_val.upper()

    if not raw_val:
        continue

    # 1) Exact country match
    if upper in country_to_region:
        cur_country = pairs.loc[pairs['Country_norm'] == upper, 'Country_title'].iloc[0]
        cur_region  = country_to_region[upper]
        continue

    # 2) Region
    if raw_val.title() in region_list_title:
        cur_region = raw_val.title()
        cur_country = None
        continue

    # 3) Skip TOTAL-like rows
    if "TOTAL" in upper or "USER" in upper or "UTOPIA" in upper:
        continue

    # 4) Fuzzy country detection
    cand = difflib.get_close_matches(upper, country_list_upper, n=1, cutoff=0.8)
    if cand:
        upper_match = cand[0]
        cur_country = pairs.loc[pairs['Country_norm'] == upper_match, 'Country_title'].iloc[0]
        cur_region  = country_to_region[upper_match]
        continue

    # 5) Agency
    if cur_region:
        df.at[i, 'Country'] = cur_country
        df.at[i, 'Agency'] = raw_val
        df.at[i, 'Region'] = cur_region

# ------------------------------
# 4. CLEAN VALID ONLY
# ------------------------------
df_clean = df.dropna(subset=['Agency']).copy()
df_clean = df_clean.dropna(subset=['Country', 'Region'])

df_clean = df_clean[
    ~df_clean['Agency'].str.upper().str.contains(r'TOTAL|USER|GRAND', na=False)
]

# ------------------------------
# NUMERIC FIX
# ------------------------------
num_cols = [
    c for c in df_clean.columns
    if c not in ['raw', 'agencygroup', 'Month', 'Country', 'Agency', 'Region']
]

for col in num_cols:
    df_clean[col] = pd.to_numeric(
        df_clean[col].astype(str)
        .str.replace('.', '', regex=False)
        .str.replace(',', '.', regex=False)
        .str.replace(r'[^0-9.-]', '', regex=True),
        errors='coerce'
    ).fillna(0)

# ------------------------------
# 4.5 → UPPERCASE all TEXT COLUMNS
# ------------------------------
for col in df_clean.select_dtypes(include=['object']).columns:
    df_clean[col] = df_clean[col].astype(str).str.upper().str.strip()

# ------------------------------
# 5. FINAL ORDER
# ------------------------------
final_cols = ['Month', 'Country', 'Agency', 'Region'] + num_cols

result = df_clean[final_cols].sort_values(
    ['Month', 'Country', 'Agency']
).reset_index(drop=True)

# ------------------------------
# SAVE
# ------------------------------
result.to_excel(OUTPUT_PATH, index=False)

print("\n✅ ГОТОВО! Итог сохранён →", OUTPUT_PATH)
print(result.head(10))
