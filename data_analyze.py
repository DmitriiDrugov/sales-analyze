import pandas as pd

file_path = './data/data_base.xlsx'
sheets = pd.read_excel(file_path, sheet_name=None)
cleaned_sheets = {}
for sheet_name, df in sheets.items():
    print(f"Обрабатываю лист: {sheet_name}")

    # Удаляем дубликаты
    df = df.drop_duplicates()

    # Убираем пропуски в ключевых полях (если такие есть в листе)
    critical_columns = ['Сумма продажи', 'Себестоимость', 'Дата']
    existing_critical = [col for col in critical_columns if col in df.columns]
    if existing_critical:
        df = df.dropna(subset=existing_critical)

    # Приводим типы данных
    if 'Дата' in df.columns:
        df['Дата'] = pd.to_datetime(df['Дата'], errors='coerce')

    numeric_columns = ['Сумма продажи', 'Себестоимость', 'Количество заказов', 'Логистика']
    for col in numeric_columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    # Убираем пробелы в текстовых полях
    text_columns = ['Клиент ID', 'Менеджер', 'Категория', 'Подкатегория', 'Номенклатура', 'Регион']
    for col in text_columns:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    # Добавляем новые поля, если возможно
    if all(x in df.columns for x in ['Сумма продажи', 'Себестоимость']):
        df['Прибыль'] = df['Сумма продажи'] - df['Себестоимость']
    if 'Сумма продажи' in df.columns and 'Количество заказов' in df.columns:
        df['Средний чек'] = df['Сумма продажи'] / df['Количество заказов']
    if 'Прибыль' in df.columns and 'Сумма продажи' in df.columns:
        df['Рентабельность'] = df['Прибыль'] / df['Сумма продажи']

    # Сохраняем очищенный датафрейм в словарь
    cleaned_sheets[sheet_name] = df

# 3. Сохраняем все очищенные листы в новый Excel-файл
clean_file_path = './data/data_base_cleaned.xlsx'
with pd.ExcelWriter(clean_file_path) as writer:
    for sheet_name, df in cleaned_sheets.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)