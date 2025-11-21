from PySide6.QtCore import QObject, Signal, Slot

class Worker(QObject):
    finished = Signal(str)
    error = Signal(str)
    log = Signal(str)

    def __init__(self, input_file, pairs_file, output_file):
        super().__init__()
        self.input_file = input_file
        self.pairs_file = pairs_file
        self.output_file = output_file

    @Slot()
    def run(self):
        try:
            import pandas as pd
            import warnings
            import difflib
            warnings.filterwarnings('ignore')

            self.log.emit("📥 VERİ YÜKLENİYOR...")

            # --- Загружаем пары стран и регионов ---
            pairs = pd.read_excel(self.pairs_file)
            pairs.columns = pairs.columns.astype(str).str.strip().str.title()
            pairs['Country_norm'] = pairs['Country'].astype(str).str.upper().str.strip()
            pairs['Country_title'] = pairs['Country'].astype(str).str.title().str.strip()
            pairs['Region_title'] = pairs['Region'].astype(str).str.title().str.strip()

            country_to_region = dict(zip(pairs['Country_norm'], pairs['Region_title']))
            region_list_title = set(pairs['Region_title'].unique())
            country_list_upper = list(country_to_region.keys())

            # --- Загружаем основной файл ---
            raw = pd.read_excel(self.input_file, header=None)
            header_row = raw[raw.iloc[:, 0].astype(str).str.contains("AgencyGroup", na=False)].index[0]
            df = pd.read_excel(self.input_file, header=header_row)

            df = df.dropna(how='all').dropna(axis=1, how='all').reset_index(drop=True)
            df.columns = (
                df.columns.astype(str)
                .str.strip()
                .str.replace(r"\s+", "_", regex=True)
                .str.lower()
            )

            # --- Определяем месяц ---
            df['raw'] = df['agencygroup'].astype(str).str.strip()
            month_mask = df['raw'].str.match(r'^\d{2}-', na=False)
            df['Month'] = pd.NA
            cur_month = None

            for i in df.index:
                if month_mask[i]:
                    cur_month = df.at[i, 'raw'].split('-', 1)[1].strip().title()
                df.at[i, 'Month'] = cur_month

            df = df[~month_mask].reset_index(drop=True)

            # --- Присвоение стран, регионов и агентств ---
            df['Country'] = pd.NA
            df['Agency'] = pd.NA
            df['Region'] = pd.NA
            df['YIL'] = 2026
            cur_country = None
            cur_region = None

            for i in df.index:
                raw_val = str(df.at[i, 'raw']).strip()
                upper = raw_val.upper()

                if not raw_val:
                    continue

                # Строка — страна
                if upper in country_to_region:
                    cur_country = pairs.loc[pairs['Country_norm'] == upper, 'Country_title'].iloc[0]
                    cur_region = country_to_region[upper]
                    continue

                # Строка — регион
                if raw_val.title() in region_list_title:
                    cur_region = raw_val.title()
                    cur_country = None
                    continue

                # Игнорируем TOTAL, USER, UTOPIA
                if any(x in upper for x in ['TOTAL', 'USER', 'UTOPIA']):
                    continue

                # Попытка fuzzy match к стране
                cand = difflib.get_close_matches(upper, country_list_upper, n=1, cutoff=0.8)
                if cand:
                    match = cand[0]
                    cur_country = pairs.loc[pairs['Country_norm'] == match, 'Country_title'].iloc[0]
                    cur_region = country_to_region[match]
                    continue

                # Присваиваем агентство с текущими страной и регионом
                if cur_country and cur_region:
                    df.at[i, 'Country'] = cur_country
                    df.at[i, 'Region'] = cur_region
                    df.at[i, 'Agency'] = raw_val

            # --- Очищаем данные ---
            df_clean = df.dropna(subset=['Agency']).copy()
            df_clean = df_clean.dropna(subset=['Country', 'Region'])
            df_clean = df_clean[
                ~df_clean['Agency'].str.upper().str.contains(r'TOTAL|USER|GRAND', na=False)
            ]

            # --- Объединяем финальный DataFrame ---
            num_cols = [c for c in df_clean.columns if c not in ['raw','agencygroup','Month','Country','Agency','Region']]
            final_cols = ['YIL','Month', 'Region', 'Country', 'Agency'] + num_cols
            result = df_clean[final_cols].sort_values(['Month','Country','Agency']).reset_index(drop=True)

            # --- Сохраняем результат ---
            result.to_excel(self.output_file, index=False)
            self.finished.emit(self.output_file)

        except Exception as e:
            self.error.emit(str(e))
