"""
Модуль розрахунку номерів БР/БН за датами
"""
from datetime import datetime, timedelta
from typing import List, Tuple

def get_br_number(date: datetime) -> str:
    """
    Повертає номер БР для заданої дати
    
    Логіка:
    - Кожен день року = 1 БР
    - Формат: №[номер_дня_року] від [дата_попереднього_дня]
    - Приклад: 01.05.2025 → №121 від 30.04.2025
    
    Args:
        date: Дата для якої генерується БР
        
    Returns:
        str: Номер БР у форматі "121 від 30.04.2025"
    """
    year = date.year
    day_of_year = date.timetuple().tm_yday
    prev_date = date - timedelta(days=1)
    
    return f"№{day_of_year} від {prev_date.strftime('%d.%m.%Y')}"

def get_br_numbers_for_dates(dates: List[datetime]) -> List[str]:
    """
    Повертає список номерів БР для списку дат
    
    Args:
        dates: Список дат
        
    Returns:
        List[str]: Список номерів БР
    """
    return [get_br_number(date) for date in dates]

def get_br_numbers_for_period(start_date: datetime, end_date: datetime) -> List[str]:
    """
    Повертає список номерів БР для періоду
    
    Args:
        start_date: Початкова дата
        end_date: Кінцева дата
        
    Returns:
        List[str]: Список номерів БР для всіх днів у періоді
    """
    dates = []
    current_date = start_date
    
    while current_date <= end_date:
        dates.append(current_date)
        current_date += timedelta(days=1)
    
    return get_br_numbers_for_dates(dates)

def format_br_list(br_numbers: List[str]) -> str:
    """
    Форматує список номерів БР у рядок для вставки в документ
    
    Args:
        br_numbers: Список номерів БР
        
    Returns:
        str: Відформатований рядок
    """
    if not br_numbers:
        return ""
    
    return ", ".join(br_numbers)

def parse_date_from_excel_cell(cell_value) -> datetime:
    """
    Парсить дату з Excel комірки
    
    Args:
        cell_value: Значення комірки Excel
        
    Returns:
        datetime: Об'єкт дати
    """
    if isinstance(cell_value, datetime):
        return cell_value
    elif isinstance(cell_value, str):
        # Спробуємо різні формати дат
        formats = ["%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d"]
        for fmt in formats:
            try:
                return datetime.strptime(cell_value, fmt)
            except ValueError:
                continue
        raise ValueError(f"Неможливо розпарсити дату: {cell_value}")
    else:
        raise ValueError(f"Невідомий тип даних для дати: {type(cell_value)}")

def get_day_column_for_date(date: datetime, base_year: int, base_month: int) -> int:
    """
    Повертає номер стовпця Excel для конкретної дати
    
    Args:
        date: Дата
        base_year: Базовий рік
        base_month: Базовий місяць
        
    Returns:
        int: Номер стовпця (7 = G, 8 = H, тощо)
    """
    if date.year != base_year or date.month != base_month:
        raise ValueError(f"Дата {date} не належить до місяця {base_month}.{base_year}")
    
    day = date.day
    return 6 + day  # Стовпець G = 7, тому 6 + day

# Приклади використання
if __name__ == "__main__":
    # Тестування
    test_date = datetime(2025, 5, 1)
    print(f"БР для {test_date.strftime('%d.%m.%Y')}: {get_br_number(test_date)}")
    
    test_date2 = datetime(2025, 6, 9)
    print(f"БР для {test_date2.strftime('%d.%m.%Y')}: {get_br_number(test_date2)}")
    
    # Тест періоду
    start = datetime(2025, 5, 1)
    end = datetime(2025, 5, 3)
    br_list = get_br_numbers_for_period(start, end)
    print(f"БР для періоду {start.strftime('%d.%m.%Y')} - {end.strftime('%d.%m.%Y')}:")
    for br in br_list:
        print(f"  {br}")
    
    print(f"Відформатований список: {format_br_list(br_list)}")

