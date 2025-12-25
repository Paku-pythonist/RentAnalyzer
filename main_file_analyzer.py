import os
import pandas as pd
import numpy as np
from datetime import datetime, date
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

from settings import AppConfig


# _____________________________________


# _____________________________________ Тестирование ___________________________________________________________________
def test_normalize_data():
    # Создаём тестовый DataFrame
    result = normalize_data(AppConfig.TEST_DF)  # имитация
    assert 'Доход' in result.columns
    assert result['Доход'].iloc[0] == AppConfig.TEST_RESULT


# ______________________________________________________________________________________________________________________


# _____________________________________ Нормализация Данных из Excel-файла _____________________________________________
def calculate_income(row: pd.DataFrame) -> pd.DataFrame:
    source = row['Источник']
    for rate_str, sources in AppConfig.COMMISSION_RATES.items():
        if source in sources:
            return row['Сумма'] * float(rate_str)
    return row['Сумма']  # Без комиссии


# _____________________________________


# _____________________________________
def validate_dates(df: pd.DataFrame) -> pd.DataFrame:  # Проверка корректности дат и вывод статистики ошибок
    invalid_mask = df['Заезд'].isna() | df['Выезд'].isna()
    if invalid_mask.any():
        logger.warning(f" Пропущено {invalid_mask.sum()} строк с некорректными датами.")

    return df[~invalid_mask]  # Фильтрация ошибки, возврат списка проблемных строк


# _____________________________________


# _____________________________________
def read_excel_file(excel_path: str) -> pd.DataFrame:
    # 1. Проверка существования файла
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Файл не найден: {excel_path}")

    # 2. Чтение Excel-файла
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        raise Exception(f"Ошибка при чтении Excel-файла: {e}")

    return df


# _____________________________________


# _____________________________________
def normalize_data(df: pd.DataFrame) -> pd.DataFrame:
    # 1. Проверка наличия нужных столбцов
    missing = [col for col in AppConfig.REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(f"В файле отсутствуют столбцы: {missing}")

    # 2. Формируем графу Доход
    df['Доход'] = df.apply(calculate_income, axis=1)

    # 3. Приведение дат к типу datetime
    df['Заезд'] = pd.to_datetime(df['Заезд'], dayfirst=True, errors='coerce')
    df['Выезд'] = pd.to_datetime(df['Выезд'], dayfirst=True, errors='coerce')

    # 4. Проверка на пустые даты
    df = validate_dates(df)

    # 5. Извлекаем месяц и год для группировки
    df['Месяц'] = df['Заезд'].dt.to_period('M')  # формат: YYYY-MM

    # 6. Сортировка по Объекту и дате
    df = df.sort_values(['Объект', 'Заезд'])

    # 7. Вычисление количества дней проживания
    df['Дни_проживания'] = (df['Выезд'] - df['Заезд']).dt.days

    # 8. Удаляем строки с отрицательными или нулевыми днями (ошибки ввода)
    df = df[df['Дни_проживания'] > 0]

    return df


# ______________________________________________________________________________________________________________________


# _____________________________________ Статистика по Доходу и Загруженности ___________________________________________
def analyze_stay_days(df: pd.DataFrame) -> pd.DataFrame:
    # 1. Группировка по Объекту и Месяцу, а также суммирование дней проживания
    grouped_df = df.groupby(['Объект', 'Месяц'])[['Дни_проживания', 'Доход']].sum().reset_index()

    # # Преобразуем 'Месяц' обратно в читаемый формат (например, 'Январь 2025')
    # result['Месяц'] = result['Месяц'].dt.strftime('%B %Y')

    # 2. Сортировка по Объекту и дате
    grouped_df = grouped_df.sort_values(['Объект', 'Месяц'])

    # 3. Выбор данных по Дате
    today = date.today()
    target_date = f"{today.year}-{today.month}"

    filtered_df = grouped_df[grouped_df['Месяц'] == target_date]

    return filtered_df


# ______________________________________________________________________________________________________________________


# _____________________________________ Сохранение отчета ______________________________________________________________
def normalize_df_dates(df: pd.DataFrame) -> pd.DataFrame:
    # Приведение Дат к нормальному виду
    df['Заезд'] = df['Заезд'].dt.date
    df['Выезд'] = df['Выезд'].dt.date
    return df


# _____________________________________


# _____________________________________
def save2xlsx(df, sheet_name, file_name, mode='a'):
    df_copy = df.copy()

    if 'Заезд' in df_copy.columns and 'Выезд' in df_copy.columns:
        df_copy = normalize_df_dates(df_copy)

    # Проверяем существование файла при mode='a'
    if mode == 'a' and not os.path.exists(file_name):
        mode = 'w'  # Если файла нет, создаём заново
        print(f"! Нарушена логика записи листа {sheet_name} в итоговом файле")

    try:
        with pd.ExcelWriter(file_name, engine='openpyxl', mode=mode) as writer:
            df_copy.to_excel(writer, sheet_name=sheet_name.replace('/', '.'), index=False)
        print(f'Saved: {sheet_name}')
    except Exception as e:
        raise IOError(f"Ошибка записи в Excel: {e}")


# ______________________________________________________________________________________________________________________


# ______________________________________________________________________________________________________________________
if __name__ == "__main__":
    start_file = "Для Анализа.xlsx"  # Исходный файл формата *.xlsx
    result_file = 'Список Бронирований.xlsx'  # Результат

    try:
        # Тестирование
        test_normalize_data()

        # Приведение данных к рабочему формату
        df = read_excel_file(start_file)
        normal_df = normalize_data(df)
        save2xlsx(normal_df, result_file.split('.')[0], result_file, mode='w')

        # Формирование общей статистики
        stay_days_df = analyze_stay_days(normal_df)
        save2xlsx(stay_days_df, 'Анализ прод-ти пребывания', result_file)

        # Формирование отчетов по каждому объекту
        unique_object_list = sorted(normal_df['Объект'].unique())
        for unique_object in unique_object_list:
            filtered_df = normal_df[normal_df['Объект'] == unique_object].copy()
            save2xlsx(filtered_df, unique_object, result_file)

    except Exception as e:
        print(f"!Ошибка: {e}")
# ______________________________________________________________________________________________________________________
