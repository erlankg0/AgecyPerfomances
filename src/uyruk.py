import pandas as pd
import warnings
warnings.filterwarnings('ignore')

FILE_PATH = "input-data/uyruk/uyruk.xlsx"
OUTPUT_PATH = "output-data/uyruk_perfomans.xlsx"

print("🚀 Парсинг: Месяц → Регион → Страна → Агентство")

# ------------------------------
# 1. Загрузка Excel
# ------------------------------
raw = pd.read_excel(FILE_PATH, header=None)
header_row = raw[raw.iloc[:, 0].astype(str).str.contains("AgencyGroup", na=False)].index[0]
df = pd.read_excel(FILE_PATH, header=header_row)
df = df.dropna(how='all').dropna(axis=1, how='all').reset_index(drop=True)

# Нормализация колонок
df.columns = (df.columns.astype(str)
              .str.strip()
              .str.replace(r'\s+', ' ', regex=True)
              .str.replace(' ', '_')
              .str.lower())

# ------------------------------
# 2. Определение месяца
# ------------------------------
df['raw'] = df['agencygroup'].astype(str).str.strip()
month_mask = df['raw'].str.match(r'^\d{2}-', na=False)

current_month = None
df['Month'] = pd.NA

for i in df.index:
    if month_mask[i]:
        current_month = df.at[i, 'raw'].split('-', 1)[1].strip().title()
    df.at[i, 'Month'] = current_month

df = df[~month_mask].reset_index(drop=True)

# ------------------------------
# 3. Регион → Страна → Агентство
# ------------------------------
df['Region'] = pd.NA
df['Country'] = pd.NA
df['Agency'] = pd.NA

current_country = None
current_region = None

# Маппинг стран в регионы
nation_to_region = {
    'ARMENIA': 'CIS',
    'AZERBAIJAN': 'CIS',
    'BELARUS': 'CIS',
    'RUSSIA': 'CIS',
    'GEORGIA': 'CIS',
    'KAZAKHSTAN': 'CIS',
    'UKRAINE': 'CIS',
    'GERMANY': 'Europe',
    'ITALY': 'Europe',
    'FRANCE': 'Europe',
    'SPAIN': 'Europe',
    'UNITED KINGDOM': 'Europe',
    'TURKEY': 'Domestic',
    'IRAN': 'Middle East',
    'IRAQ': 'Middle East',
    'CHINA': 'Far East',
    'JAPAN': 'Far East',
    'KOREA': 'Far East',
    'USA': 'Other',
    'CANADA': 'Other',
    'AUSTRALIA': 'Other'
}

for i in df.index:
    val = str(df.at[i, 'raw']).strip()
    upper = val.upper()

    # Строка-нация
    if upper in nation_to_region:
        current_country = val.title()
        current_region = nation_to_region[upper]
        continue

    # Пропускаем TOTAL строки
    if 'TOTAL' in upper or val == current_country:
        continue

    # Всё остальное — агентство
    if current_country:
        df.at[i, 'Country'] = current_country
        df.at[i, 'Region'] = current_region
        df.at[i, 'Agency'] = val

# ------------------------------
# 4. Чистка
# ------------------------------
df_clean = df.dropna(subset=['Agency']).copy()

# ------------------------------
# 5. Преобразование числовых колонок
# ------------------------------
num_cols = [col for col in df_clean.columns if col not in ['raw','agencygroup','Month','Region','Country','Agency']]
for col in num_cols:
    df_clean[col] = pd.to_numeric(
        df_clean[col].astype(str)
        .str.replace('.', '', regex=False)      # убираем точки тысяч
        .str.replace(',', '.', regex=False)     # заменяем запятую на точку
        .str.replace(r'[^0-9.-]', '', regex=True),
        errors='coerce'
    ).fillna(0)

# ------------------------------
# 6. Итоговая таблица
# ------------------------------
final_cols = ['Month', 'Region', 'Country', 'Agency'] + num_cols
result = df_clean[final_cols].sort_values(['Month','Region','Country','Agency']).reset_index(drop=True)

# ------------------------------
# 7. Сохранение Excel
# ------------------------------
result.to_excel(OUTPUT_PATH, index=False)
print(f"\n✅ Сохранено → {OUTPUT_PATH}")
print(result.head(10))
