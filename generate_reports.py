"""
Головний скрипт генерації рапортів з меню вибору місяця та типу рапорту
"""
import os
import sys
from datetime import datetime
from month_utils import get_available_months
from excel_processor import TabelReader
from word_generator import WordReportGenerator
from excel_reports import ExcelReportGenerator

class ReportGenerator:
    """Альварес-AI для генерації всіх типів рапортів"""
    
    def __init__(self, excel_file: str = "Табель_Багатомісячний.xlsx"):
        self.excel_file = excel_file
        self.reader = TabelReader(excel_file)
        self.word_generator = WordReportGenerator()
        self.excel_generator = ExcelReportGenerator()
        
        # Автоматично визначаємо доступні місяці з аркушів Excel
        from path_utils import get_app_dir
        app_dir = get_app_dir()
        excel_path = os.path.join(app_dir, excel_file) if not os.path.isabs(excel_file) else excel_file
        self.available_months = get_available_months(excel_path)
        
        # Оберіть тип необхідних даних:
        self.report_types = {
            "1": "ДГВ 100к (Excel)",
            "2": "ДГВ 30к (Excel)", 
            "3": "Підтвердження 100к (Word)",
            "4": "Підтвердження 30к (Word)",
            "5": "ДГВ 0к (Excel)",
            "6": "Створити всі типи за обраний місяць"
        }
    
    def run(self):
        """Запускає головне меню"""
        print("=" * 60)
        print("АЛЬВАРЕС AI — СИСТЕМА ОБЛІКУ ОСОБОВОГО СКЛАДУ 12 ШТУРМОВОЇ РОТИ")
        print("=" * 60)
        
        while True:
            try:
                # Вибір місяця
                month = self._select_month()
                if month is None:
                    break
                
                # Вибір типу рапорту
                report_type = self._select_report_type()
                if report_type is None:
                    continue
                
                # Генерація рапорту
                self._generate_report(month, report_type)
                
                # Продовжити?
                if not self._ask_continue():
                    break
                    
            except KeyboardInterrupt:
                print("\n\nПрограма завершена користувачем.")
                break
            except Exception as e:
                print(f"\nПомилка: {e}")
                input("Натисніть Enter для продовження...")
    
    def _select_month(self):
        """Вибір місяця"""
        print("\nОберіть потрібний місяць:")
        for i, month in enumerate(self.available_months, 1):
            print(f"{i}. {month}")
        print("0. Вихід")
        
        while True:
            try:
                choice = input("\nВаш вибір (номер): ").strip()
                
                if choice == "0":
                    return None
                
                choice_num = int(choice)
                if 1 <= choice_num <= len(self.available_months):
                    selected_month = self.available_months[choice_num - 1]
                    print(f"Обрано: {selected_month}")
                    return selected_month
                else:
                    print("Ти походу щось попутав, спробуй обрати ще раз")
                    
            except ValueError:
                print("Введіть число.")
    
    def _select_report_type(self):
        """Вибір типу рапорту"""
        print("\nОберіть тип необхідних даних:")
        for key, description in self.report_types.items():
            print(f"{key}. {description}")
        print("0. Назад")
        
        while True:
            try:
                choice = input("\nВаш вибір (номер): ").strip()
                
                if choice == "0":
                    return None
                
                if choice in self.report_types:
                    print(f"Обрано: {self.report_types[choice]}")
                    return choice
                else:
                    print("Вась, не тупи, спробуй обрати ще раз")
                    
            except ValueError:
                print("Введіть число.")
    
    def _generate_report(self, month: str, report_type: str):
        """Генерує обраний рапорт"""
        print(f"\nВитя Альварес розпочав генерацію даних за {month}...")
        
        try:
            # Завантажуємо дані
            self.reader.load_workbook()
            soldiers = self.reader.read_month_data(month)
            
            if not soldiers:
                print("Не знайдено даних для цього місяця, походу всі були вихідні")
                return
            
            print(f"Знайдено та оброблено {len(soldiers)} військовослужбовців 12 штурмової роти")
            
            # Формуємо назву місяця для файлів
            month_display = month.replace("_", " ").lower()
            
            if report_type == "1":
                # ДГВ 100к
                soldiers_100 = self.reader.get_soldiers_by_category(soldiers, "100", include_no_payment=False)
                if soldiers_100:
                    filename = f"ДГВ_100к_{month_display}.xlsx"
                    self.excel_generator.create_dgv_report(soldiers_100, month_display, "100", filename)
                else:
                    print("Не знайдено військовослужбовців 12 штурмової роти на 100 (без 'не виплачувати')")
            
            elif report_type == "2":
                # ДГВ 30к
                soldiers_30 = self.reader.get_soldiers_by_category(soldiers, "30", include_no_payment=False)
                if soldiers_30:
                    filename = f"ДГВ_30к_{month_display}.xlsx"
                    self.excel_generator.create_dgv_report(soldiers_30, month_display, "30", filename)
                else:
                    print("Не знайдено військовослужбовців 12 штурмової роти на 30 (без 'не виплачувати')")
            
            elif report_type == "3":
                # Підтвердження 100к
                soldiers_100 = self.reader.get_soldiers_by_category(soldiers, "100", include_no_payment=True)
                if soldiers_100:
                    filename = f"Підтвердження_100к_{month_display}.docx"
                    self.word_generator.create_confirmation_report(soldiers_100, month_display, "100", filename)
                else:
                    print("Не знайдено військовослужбовців 12 штурмової роти на 100")
            
            elif report_type == "4":
                # Підтвердження 30к
                soldiers_30 = self.reader.get_soldiers_by_category(soldiers, "30", include_no_payment=True)
                if soldiers_30:
                    filename = f"Підтвердження_30к_{month_display}.docx"
                    self.word_generator.create_confirmation_report(soldiers_30, month_display, "30", filename)
                else:
                    print("Не знайдено військовослужбовців 12 штурмової роти на 30")
            
            elif report_type == "5":
                # ДГВ 0к
                soldiers_0 = self.reader.get_soldiers_by_category(soldiers, "0", include_no_payment=False)
                if soldiers_0:
                    filename = f"ДГВ_0к_{month_display}.xlsx"
                    self.excel_generator.create_dgv_report(soldiers_0, month_display, "0", filename)
                else:
                    print("Не знайдено військовослужбовців 12 штурмової роти на 0 (без 'не виплачувати')")
            
            elif report_type == "6":
                # Всі рапорти
                self._generate_all_reports(soldiers, month_display)
            
            print("Вітя Альварес роботу завершив — необхідні дані успішно створено!")
            
        except Exception as e:
            print(f"Бляяяя, Вітя рубає окуня — помилка при генерації необхідних даних: {e}")
            raise
    
    def _generate_all_reports(self, soldiers, month_display: str):
        """Вітя Альварес генерує всі типи рапортів за місяць"""
        print('"Працюю, як завжди швидко" © Вітя Альварес')
        
        # ДГВ 100к
        soldiers_100 = self.reader.get_soldiers_by_category(soldiers, "100", include_no_payment=False)
        if soldiers_100:
            filename = f"ДГВ_100к_{month_display}.xlsx"
            self.excel_generator.create_dgv_report(soldiers_100, month_display, "100", filename)
        
        # ДГВ 30к
        soldiers_30 = self.reader.get_soldiers_by_category(soldiers, "30", include_no_payment=False)
        if soldiers_30:
            filename = f"ДГВ_30к_{month_display}.xlsx"
            self.excel_generator.create_dgv_report(soldiers_30, month_display, "30", filename)
        
        # Підтвердження 100к
        soldiers_100_all = self.reader.get_soldiers_by_category(soldiers, "100", include_no_payment=True)
        if soldiers_100_all:
            filename = f"Підтвердження_100к_{month_display}.docx"
            self.word_generator.create_confirmation_report(soldiers_100_all, month_display, "100", filename)
        
        # Підтвердження 30к
        soldiers_30_all = self.reader.get_soldiers_by_category(soldiers, "30", include_no_payment=True)
        if soldiers_30_all:
            filename = f"Підтвердження_30к_{month_display}.docx"
            self.word_generator.create_confirmation_report(soldiers_30_all, month_display, "30", filename)
        
        # ДГВ 0к
        soldiers_0 = self.reader.get_soldiers_by_category(soldiers, "0", include_no_payment=False)
        if soldiers_0:
            filename = f"ДГВ_0к_{month_display}.xlsx"
            self.excel_generator.create_dgv_report(soldiers_0, month_display, "0", filename)
        
        print("Вітя Альварес роботу завершив — всі дані створено!")
    
    def _ask_continue(self) -> bool:
        """Питає чи продовжити роботу"""
        while True:
            choice = input("\nБажаєте підгрузити роботою Альвареса ще раз? (y/n): ").strip().lower()
            if choice in ['y', 'yes', 'так', 'т']:
                return True
            elif choice in ['n', 'no', 'ні', 'н']:
                return False
            else:
                print("Введіть 'y' або 'n'")

def main():
    """Головна функція"""
    print("АЛЬВАРЕС AI — зробить все! (якщо не забуде)")

    # Визначаємо директорію додатку
    from path_utils import get_app_dir
    excel_file = os.path.join(get_app_dir(), "Табель_Багатомісячний.xlsx")

    # Перевіряємо наявність файлу
    if not os.path.exists(excel_file):
        print(f"Помилка: Вітя не може знайти файл табелю")
        print(f"Шукав тут: {excel_file}")
        print("Переконайся, що файл 'Табель_Багатомісячний.xlsx' знаходиться в папці зі скриптом.")
        return
    
    # Запускаємо генератор
    generator = ReportGenerator(excel_file)
    generator.run()
    
    print("\nДякуємо за використання АЛЬВАРЕС AI!")
    print("\nАвтор: Володимир Барт, діловод 12 штурмової роти 4 штурмового батальйону")

if __name__ == "__main__":
    main()

