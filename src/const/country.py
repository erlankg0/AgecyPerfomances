import pandas as pd

# Загружаем файл
df = pd.read_excel("./county.xlsx")

# Оставляем только нужные колонки
df = df[['region', 'country']]

# Удаляем дубликаты пар region–country
unique_pairs = df.drop_duplicates().sort_values(by=['region', 'country'])

# Сохраняем новый Excel
unique_pairs.to_excel("unique_region_country.xlsx", index=False)

print("Готово! Сохранён файл unique_region_country.xlsx")
