"""
Графічний інтерфейс для системи генерації рапортів ДГВ
"""
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import os
import openpyxl
from datetime import datetime
from typing import Optional

try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

from generate_reports import ReportGenerator
from br_calculator import get_br_number
from month_utils import (get_available_months, parse_month_sheet_name, get_source_filename,
                         build_month_sheet_name, MONTH_NAMES_UK_REVERSE)
from tabel_filler import fill_single_month, fill_tabel_months
from data.database import (init_db, get_all_personnel, get_all_roles,
                           set_personnel_role)
from core.br_roles import (auto_assign_all_roles, import_personnel_from_tabel,
                           build_composition_for_date, generate_br_word)
from path_utils import get_base_path, get_app_dir


# Палітра темної теми
_BG = "#1e1e2e"          # основний фон
_BG_CARD = "#2a2a3d"     # фон карток/фреймів
_BG_INPUT = "#33334d"    # фон полів введення / логів
_FG = "#cdd6f4"          # основний текст
_FG_DIM = "#7f849c"      # приглушений текст
_HEADER_BG = "#181825"   # заголовок
_HEADER_FG = "#cdd6f4"
_STATUS_BG = "#11111b"
_STATUS_FG = "#a6adc8"


class ReportGUI:
    """Графічний інтерфейс для генерації рапортів"""

    def __init__(self, root):
        self.root = root
        self.root.title("АЛЬВАРЕС AI — Система обліку особового складу")
        self.root.geometry("800x700")
        self.root.resizable(True, True)
        self.root.configure(bg=_BG)

        # Налаштовуємо темний стиль для ttk
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(".", background=_BG, foreground=_FG, fieldbackground=_BG_INPUT)
        style.configure("TCombobox", fieldbackground=_BG_INPUT, background=_BG_CARD,
                         foreground=_FG, arrowcolor=_FG)
        style.configure("Treeview", background=_BG_INPUT, foreground=_FG,
                         fieldbackground=_BG_INPUT, rowheight=25)
        style.configure("Treeview.Heading", background=_HEADER_BG, foreground=_FG,
                         font=("Arial", 10, "bold"))
        style.map("Treeview", background=[("selected", "#45475a")],
                  foreground=[("selected", "#cdd6f4")])
        style.configure("TScrollbar", background=_BG_CARD, troughcolor=_BG,
                         arrowcolor=_FG)

        # Шляхи: base_path — bundled ресурси, app_dir — робочі файли
        base_path = get_base_path()
        app_dir = get_app_dir()

        # Ініціалізуємо генератор
        self.generator = None
        self.excel_file = os.path.join(app_dir, "Табель_Багатомісячний.xlsx")
        
        # Змінні для вибору
        self.selected_month = tk.StringVar()
        self.report_type_var = tk.StringVar(value="6")
        
        # Змінні для БР
        self.start_date_var = tk.StringVar()
        self.end_date_var = tk.StringVar()
        self.order_number_var = tk.StringVar()
        # Фон вимкнено; використовуємо лого у заголовку
        self.bg_image = None
        self.canvas = None
        self.content_window_id = None
        self.logo_image = None
        
        # Поточний екран
        self.current_screen = None
        
        # Ініціалізуємо базу даних ролей
        init_db()

        self.template_path = os.path.join(base_path, "templates", "rozp_template.docx")
        self.br_4shb_file = os.path.join(app_dir, "BR_4ShB.xlsx")
        self.output_dir = os.path.join(app_dir, "output")

        # Створюємо інтерфейс
        self._create_main_menu()
        
        # Перевіряємо наявність файлів
        self._check_files()
    
    def _load_background_image(self):
        """Завантажує та створює напівпрозорий фон з emblem.png"""
        if not PIL_AVAILABLE:
            return None

        emblem_path = os.path.join(get_base_path(), "emblem.png")
        if not os.path.exists(emblem_path):
            return None
        
        try:
            # Завантажуємо зображення
            img = Image.open(emblem_path)
            
            # Отримуємо розміри вікна
            width = self.root.winfo_width() if self.root.winfo_width() > 1 else 800
            height = self.root.winfo_height() if self.root.winfo_height() > 1 else 700
            
            # Масштабуємо зображення до розміру вікна
            img = img.resize((width, height), Image.Resampling.LANCZOS)
            
            # Створюємо напівпрозоре зображення (alpha = 0.3)
            # Конвертуємо в RGBA якщо ще не
            if img.mode != 'RGBA':
                img = img.convert('RGBA')
            
            # Зменшуємо прозорість
            alpha = img.split()[3]
            alpha = alpha.point(lambda p: int(p * 0.3))  # 30% непрозорості
            img.putalpha(alpha)
            
            # Конвертуємо в PhotoImage
            return ImageTk.PhotoImage(img)
        except Exception as e:
            print(f"Помилка завантаження зображення: {e}")
            return None
    
    def _create_background_canvas(self, parent):
        """Створює Canvas з фоном для батьківського віджета"""
        if self.canvas:
            self.canvas.destroy()
        
        self.canvas = tk.Canvas(parent, highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        # Оновлюємо фон
        self._update_background()
        
        return self.canvas
    
    def _update_background(self):
        """Оновлює фонове зображення"""
        if not self.canvas:
            return
        
        # Отримуємо розміри canvas
        self.canvas.update_idletasks()
        width = self.canvas.winfo_width()
        height = self.canvas.winfo_height()
        
        if width > 1 and height > 1:
            bg_image = self._load_background_image()
            if bg_image:
                self.bg_image = bg_image  # Зберігаємо посилання
                # Оновлюємо тільки фон, не видаляючи вміст
                self.canvas.delete('bg')
                self.canvas.create_image(0, 0, anchor=tk.NW, image=bg_image, tags='bg')
                self.canvas.tag_lower('bg')
        
        # Плануємо повторне оновлення при зміні розміру
        self.root.after(100, self._update_background_on_resize)
    
    def _update_background_on_resize(self):
        """Оновлює фон при зміні розміру вікна"""
        if self.canvas and self.current_screen == "main":
            self._update_background()
    
    def _get_logo(self):
        """Повертає зменшене зображення для логотипу (ліворуч вгорі)"""
        if not PIL_AVAILABLE:
            return None
        if self.logo_image:
            return self.logo_image
        emblem_path = os.path.join(get_base_path(), "emblem.png")
        if not os.path.exists(emblem_path):
            return None
        try:
            img = Image.open(emblem_path)
            # Масштабуємо по висоті ~44px з збереженням пропорцій
            target_h = 44
            w, h = img.size
            target_w = max(1, int(w * (target_h / float(h))))
            img = img.resize((target_w, target_h), Image.Resampling.LANCZOS)
            self.logo_image = ImageTk.PhotoImage(img)
            return self.logo_image
        except Exception:
            return None
    
    def _clear_screen(self):
        """Очищає поточний екран"""
        for widget in self.root.winfo_children():
            widget.destroy()
        self.canvas = None
    
    def _create_main_menu(self):
        """Створює головне меню"""
        self.current_screen = "main"
        self._clear_screen()
        
        # Контент без фону
        content_frame = tk.Frame(self.root, bg=_BG, relief=tk.FLAT)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Заголовок
        title_frame = tk.Frame(content_frame, bg=_HEADER_BG, height=80)
        title_frame.pack(fill=tk.X, pady=(0, 30))
        title_frame.pack_propagate(False)
        
        # Лого зліва
        logo = self._get_logo()
        if logo:
            logo_lbl = tk.Label(title_frame, image=logo, bg=_HEADER_BG)
            logo_lbl.image = logo
            logo_lbl.pack(side=tk.LEFT, padx=12)

        title_label = tk.Label(
            title_frame,
            text="АЛЬВАРЕС AI",
            font=("Arial", 20, "bold"),
            bg=_HEADER_BG,
            fg="white"
        )
        title_label.pack(pady=(15, 5))
        
        subtitle_label = tk.Label(
            title_frame,
            text="Система обліку особового складу 12 штурмової роти",
            font=("Arial", 12),
            bg=_HEADER_BG,
            fg=_FG_DIM
        )
        subtitle_label.pack()
        
        # Кнопки функцій
        buttons_frame = tk.Frame(content_frame, bg=_BG)
        buttons_frame.pack(fill=tk.BOTH, expand=True, padx=50, pady=20)
        
        # Кнопка 1: Створити документи за місяць
        btn1 = tk.Button(
            buttons_frame,
            text="📄 Створити документи за місяць",
            command=self._show_reports_screen,
            font=("Arial", 14, "bold"),
            bg="#3498db",
            fg="white",
            relief=tk.FLAT,
            padx=30,
            pady=20,
            cursor="hand2",
            width=40,
            height=2
        )
        btn1.pack(pady=15)
        btn1.bind("<Enter>", lambda e: btn1.config(bg="#2980b9"))
        btn1.bind("<Leave>", lambda e: btn1.config(bg="#3498db"))
        
        # Кнопка 2: Створити БР
        btn3 = tk.Button(
            buttons_frame,
            text="📋 Створити БР",
            command=self._show_br_create_screen,
            font=("Arial", 14, "bold"),
            bg="#e67e22",
            fg="white",
            relief=tk.FLAT,
            padx=30,
            pady=20,
            cursor="hand2",
            width=40,
            height=2
        )
        btn3.pack(pady=15)
        btn3.bind("<Enter>", lambda e: btn3.config(bg="#d35400"))
        btn3.bind("<Leave>", lambda e: btn3.config(bg="#e67e22"))

        # Кнопка 3.5: Ролі для БР
        btn_roles = tk.Button(
            buttons_frame,
            text="🧩 Ролі для БР",
            command=self._show_roles_screen,
            font=("Arial", 14, "bold"),
            bg="#c0392b",
            fg="white",
            relief=tk.FLAT,
            padx=30,
            pady=20,
            cursor="hand2",
            width=40,
            height=2
        )
        btn_roles.pack(pady=15)
        btn_roles.bind("<Enter>", lambda e: btn_roles.config(bg="#a93226"))
        btn_roles.bind("<Leave>", lambda e: btn_roles.config(bg="#c0392b"))

        # Кнопка 4: Заповнити табель
        btn4 = tk.Button(
            buttons_frame,
            text="Заповнити табель",
            command=self._show_tabel_filler_screen,
            font=("Arial", 14, "bold"),
            bg="#8e44ad",
            fg="white",
            relief=tk.FLAT,
            padx=30,
            pady=20,
            cursor="hand2",
            width=40,
            height=2
        )
        btn4.pack(pady=15)
        btn4.bind("<Enter>", lambda e: btn4.config(bg="#7d3c98"))
        btn4.bind("<Leave>", lambda e: btn4.config(bg="#8e44ad"))

        # Кнопка 5: Додати місяць
        btn5 = tk.Button(
            buttons_frame,
            text="Додати місяць",
            command=self._show_add_month_dialog,
            font=("Arial", 14, "bold"),
            bg="#16a085",
            fg="white",
            relief=tk.FLAT,
            padx=30,
            pady=20,
            cursor="hand2",
            width=40,
            height=2
        )
        btn5.pack(pady=15)
        btn5.bind("<Enter>", lambda e: btn5.config(bg="#138d75"))
        btn5.bind("<Leave>", lambda e: btn5.config(bg="#16a085"))

        # Статус-бар
        self.status_label = tk.Label(
            content_frame,
            text="Готово до роботи",
            relief=tk.SUNKEN,
            anchor=tk.W,
            font=("Arial", 9),
            bg=_STATUS_BG
        )
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)
    
    def _show_reports_screen(self):
        """Показує екран створення документів"""
        self.current_screen = "reports"
        self._clear_screen()
        
        content_frame = tk.Frame(self.root, bg=_BG, relief=tk.FLAT)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Заголовок
        header_frame = tk.Frame(content_frame, bg=_HEADER_BG, height=60)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        header_frame.pack_propagate(False)
        
        logo = self._get_logo()
        if logo:
            logo_lbl = tk.Label(header_frame, image=logo, bg=_HEADER_BG)
            logo_lbl.image = logo
            logo_lbl.pack(side=tk.LEFT, padx=10)

        title_label = tk.Label(
            header_frame,
            text="Створити документи за місяць",
            font=("Arial", 16, "bold"),
            bg=_HEADER_BG,
            fg="white"
        )
        title_label.pack(pady=15)
        
        # Основний контент
        main_frame = tk.Frame(content_frame, bg=_BG, padx=30, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Вибір місяця
        month_frame = tk.LabelFrame(main_frame, text="Оберіть місяць", font=("Arial", 11, "bold"), padx=15, pady=15, bg=_BG, fg=_FG)
        month_frame.pack(fill=tk.X, pady=(0, 20))
        
        months = self.generator.available_months if self.generator else get_available_months(self.excel_file)
        
        month_combo = ttk.Combobox(month_frame, textvariable=self.selected_month, values=months, 
                                   state="readonly", width=40, font=("Arial", 11))
        month_combo.pack(pady=10)
        if months:
            month_combo.set(months[0])
        
        # Вибір типу рапорту
        report_frame = tk.LabelFrame(main_frame, text="Оберіть тип рапорту", font=("Arial", 11, "bold"), padx=15, pady=15, bg=_BG, fg=_FG)
        report_frame.pack(fill=tk.X, pady=(0, 20))
        
        report_types = {
            "1": "ДГВ 100к (Excel)",
            "2": "ДГВ 30к (Excel)",
            "3": "Підтвердження 100к (Word)",
            "4": "Підтвердження 30к (Word)",
            "5": "ДГВ 0к (Excel)",
            "6": "Створити всі типи за обраний місяць"
        }
        
        for key, description in report_types.items():
            rb = tk.Radiobutton(
                report_frame,
                text=description,
                variable=self.report_type_var,
                value=key,
                font=("Arial", 10),
                anchor="w",
                bg=_BG,
                fg=_FG,
                selectcolor=_BG_CARD,
                activebackground=_BG,
                activeforeground=_FG
            )
            rb.pack(anchor="w", pady=3)
        
        # Кнопки
        button_frame = tk.Frame(main_frame, bg=_BG)
        button_frame.pack(fill=tk.X, pady=(0, 20))
        
        self.generate_btn = tk.Button(
            button_frame,
            text="📄 Створити рапорти",
            command=self._generate_reports,
            font=("Arial", 12, "bold"),
            bg="#27ae60",
            fg="white",
            relief=tk.FLAT,
            padx=20,
            pady=10,
            cursor="hand2"
        )
        self.generate_btn.pack(side=tk.LEFT, padx=5)
        
        back_btn = tk.Button(
            button_frame,
            text="← Назад",
            command=self._create_main_menu,
            font=("Arial", 11),
            bg="#95a5a6",
            fg="white",
            relief=tk.FLAT,
            padx=20,
            pady=10,
            cursor="hand2"
        )
        back_btn.pack(side=tk.RIGHT, padx=5)
        
        # Поле для логів
        log_frame = tk.LabelFrame(main_frame, text="Статус виконання", font=("Arial", 11, "bold"), padx=15, pady=15, bg=_BG, fg=_FG)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            height=8,
            font=("Consolas", 9),
            wrap=tk.WORD,
            bg=_BG_INPUT,
            relief=tk.FLAT
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Статус-бар
        self.status_label = tk.Label(
            content_frame,
            text="Готово до роботи",
            relief=tk.SUNKEN,
            anchor=tk.W,
            font=("Arial", 9),
            bg=_STATUS_BG
        )
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)
    
    def _show_br_create_screen(self):
        """Показує екран створення БР"""
        self.current_screen = "br_create"
        self._clear_screen()
        
        content_frame = tk.Frame(self.root, bg=_BG, relief=tk.FLAT)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Заголовок
        header_frame = tk.Frame(content_frame, bg=_HEADER_BG, height=60)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        header_frame.pack_propagate(False)
        
        logo = self._get_logo()
        if logo:
            logo_lbl = tk.Label(header_frame, image=logo, bg=_HEADER_BG)
            logo_lbl.image = logo
            logo_lbl.pack(side=tk.LEFT, padx=10)

        title_label = tk.Label(
            header_frame,
            text="Створити БР",
            font=("Arial", 16, "bold"),
            bg=_HEADER_BG,
            fg="white"
        )
        title_label.pack(pady=15)
        
        # Основний контент
        main_frame = tk.Frame(content_frame, bg=_BG, padx=30, pady=30)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Вибір дат
        dates_frame = tk.LabelFrame(main_frame, text="Оберіть період", font=("Arial", 11, "bold"), padx=15, pady=15, bg=_BG, fg=_FG)
        dates_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Початкова дата
        start_frame = tk.Frame(dates_frame, bg=_BG)
        start_frame.pack(fill=tk.X, pady=5)
        tk.Label(start_frame, text="Початкова дата (ДД.ММ.РРРР):", font=("Arial", 10), bg=_BG, fg=_FG).pack(side=tk.LEFT, padx=5)
        start_entry = tk.Entry(start_frame, textvariable=self.start_date_var, font=("Arial", 11), width=20)
        start_entry.pack(side=tk.LEFT, padx=5)
        start_entry.bind('<KeyRelease>', lambda e: self._update_order_number())
        
        # Кінцева дата
        end_frame = tk.Frame(dates_frame, bg=_BG)
        end_frame.pack(fill=tk.X, pady=5)
        tk.Label(end_frame, text="Кінцева дата (ДД.ММ.РРРР):", font=("Arial", 10), bg=_BG, fg=_FG).pack(side=tk.LEFT, padx=5)
        end_entry = tk.Entry(end_frame, textvariable=self.end_date_var, font=("Arial", 11), width=20)
        end_entry.pack(side=tk.LEFT, padx=5)
        
        # Номер наказу
        order_frame = tk.LabelFrame(main_frame, text="Початковий номер наказу", font=("Arial", 11, "bold"), padx=15, pady=15, bg=_BG, fg=_FG)
        order_frame.pack(fill=tk.X, pady=(0, 20))
        
        order_display = tk.Label(
            order_frame,
            textvariable=self.order_number_var,
            font=("Arial", 12, "bold"),
            bg=_BG_INPUT,
            fg=_FG,
            relief=tk.SUNKEN,
            padx=10,
            pady=5
        )
        order_display.pack(fill=tk.X)
        
        tk.Label(
            order_frame,
            text="(Номер розраховується автоматично)",
            font=("Arial", 9),
            fg=_FG_DIM,
            bg=_BG
        ).pack(pady=(5, 0))
        
        # Кнопки
        button_frame = tk.Frame(main_frame, bg=_BG)
        button_frame.pack(fill=tk.X, pady=(0, 20))

        preview_btn = tk.Button(
            button_frame,
            text="Сформувати склад",
            command=self._preview_composition,
            font=("Arial", 11, "bold"),
            bg="#2980b9",
            fg="white",
            relief=tk.FLAT,
            padx=15,
            pady=8,
            cursor="hand2"
        )
        preview_btn.pack(side=tk.LEFT, padx=5)
        preview_btn.bind("<Enter>", lambda e: preview_btn.config(bg="#2471a3"))
        preview_btn.bind("<Leave>", lambda e: preview_btn.config(bg="#2980b9"))

        gen_word_btn = tk.Button(
            button_frame,
            text="Згенерувати Word БР",
            command=self._generate_word_br,
            font=("Arial", 11, "bold"),
            bg="#27ae60",
            fg="white",
            relief=tk.FLAT,
            padx=15,
            pady=8,
            cursor="hand2"
        )
        gen_word_btn.pack(side=tk.LEFT, padx=5)
        gen_word_btn.bind("<Enter>", lambda e: gen_word_btn.config(bg="#229954"))
        gen_word_btn.bind("<Leave>", lambda e: gen_word_btn.config(bg="#27ae60"))

        back_btn = tk.Button(
            button_frame,
            text="<- Назад",
            command=self._create_main_menu,
            font=("Arial", 11),
            bg="#95a5a6",
            fg="white",
            relief=tk.FLAT,
            padx=20,
            pady=10,
            cursor="hand2"
        )
        back_btn.pack(side=tk.RIGHT, padx=5)
        
        # Поле для логів
        log_frame = tk.LabelFrame(main_frame, text="Статус виконання", font=("Arial", 11, "bold"), padx=15, pady=15, bg=_BG, fg=_FG)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            height=8,
            font=("Consolas", 9),
            wrap=tk.WORD,
            bg=_BG_INPUT,
            relief=tk.FLAT
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Статус-бар
        self.status_label = tk.Label(
            content_frame,
            text="Введіть дати для автоматичного розрахунку номера наказу",
            relief=tk.SUNKEN,
            anchor=tk.W,
            font=("Arial", 9),
            bg=_STATUS_BG
        )
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Встановлюємо поточну дату за замовчуванням
        today = datetime.now()
        self.start_date_var.set(today.strftime("%d.%m.%Y"))
        self.end_date_var.set(today.strftime("%d.%m.%Y"))
        self._update_order_number()
    
    def _update_order_number(self):
        """Оновлює номер наказу на основі початкової дати"""
        date_str = self.start_date_var.get().strip()
        if not date_str:
            self.order_number_var.set("Введіть початкову дату")
            return

        try:
            date_obj = datetime.strptime(date_str, "%d.%m.%Y")
            from br_updater import get_tabel_date
            tabel_date = get_tabel_date(date_obj)
            self.order_number_var.set(get_br_number(tabel_date))
        except ValueError:
            self.order_number_var.set("Некоректний формат дати (використовуйте ДД.ММ.РРРР)")
    
    def _check_files(self):
        """Перевіряє наявність необхідних файлів"""
        if not os.path.exists(self.excel_file):
            messagebox.showerror(
                "Помилка",
                f"Не знайдено файл: {self.excel_file}\n\n"
                "Переконайтеся, що файл знаходиться в поточній папці."
            )
        else:
            try:
                self.generator = ReportGenerator(self.excel_file)
            except Exception as e:
                messagebox.showerror("Помилка", f"Не вдалося завантажити файл:\n{str(e)}")
    
    def _log(self, message):
        """Додає повідомлення в лог"""
        if hasattr(self, 'log_text'):
            self.log_text.insert(tk.END, message + "\n")
            self.log_text.see(tk.END)
            self.root.update_idletasks()
    
    def _update_status(self, message):
        """Оновлює статус-бар"""
        if hasattr(self, 'status_label'):
            self.status_label.config(text=message)
            self.root.update_idletasks()
    
    def _generate_reports(self):
        """Запускає генерацію рапортів"""
        if not self.selected_month.get():
            messagebox.showwarning("Увага", "Оберіть місяць!")
            return
        
        month = self.selected_month.get()
        report_type = self.report_type_var.get()
        
        report_name = {
            "1": "ДГВ 100к",
            "2": "ДГВ 30к",
            "3": "Підтвердження 100к",
            "4": "Підтвердження 30к",
            "5": "ДГВ 0к",
            "6": "Всі типи рапортів"
        }.get(report_type, "Рапорт")
        
        if not messagebox.askyesno("Підтвердження", 
                                   f"Створити {report_name} за {month}?"):
            return
        
        self.generate_btn.config(state=tk.DISABLED, text="⏳ Генерація...")
        self.log_text.delete(1.0, tk.END)
        
        thread = threading.Thread(target=self._do_generate_reports, args=(month, report_type))
        thread.daemon = True
        thread.start()
    
    def _do_generate_reports(self, month, report_type):
        """Виконує генерацію рапортів"""
        try:
            self._update_status("Генерація рапортів...")
            self._log("=" * 60)
            self._log(f"Вітя Альварес розпочав генерацію даних за {month}...")
            self._log("=" * 60)
            
            self.generator.reader.load_workbook()
            soldiers = self.generator.reader.read_month_data(month)
            
            if not soldiers:
                self._log("❌ Не знайдено даних для цього місяця")
                self._update_status("Помилка: дані не знайдено")
                return
            
            self._log(f"✓ Знайдено та оброблено {len(soldiers)} військовослужбовців")
            
            month_display = month.replace("_", " ").lower()
            
            if report_type == "6":
                self._generate_all_reports(soldiers, month_display)
            else:
                self.generator._generate_report(month, report_type)
                self._log("✓ Рапорт успішно створено!")
            
            self._log("=" * 60)
            self._log("✓ Вітя Альварес роботу завершив — дані успішно створено!")
            self._log("=" * 60)
            self._update_status("Готово!")
            
            self.root.after(0, lambda: messagebox.showinfo(
                "Успіх",
                f"Рапорти успішно створено!\n\nПеревірте файли в поточній папці."
            ))
            
        except Exception as e:
            error_msg = f"Помилка: {str(e)}"
            self._log(f"❌ {error_msg}")
            self._update_status("Помилка!")
            self.root.after(0, lambda: messagebox.showerror("Помилка", error_msg))
        finally:
            self.root.after(0, lambda: self.generate_btn.config(
                state=tk.NORMAL, text="📄 Створити рапорти"
            ))
    
    def _generate_all_reports(self, soldiers, month_display):
        """Генерує всі типи рапортів"""
        self._log('"Працюю, як завжди швидко" © Вітя Альварес\n')
        
        reports = []
        
        soldiers_100 = self.generator.reader.get_soldiers_by_category(soldiers, "100", include_no_payment=False)
        if soldiers_100:
            filename = f"ДГВ_100к_{month_display}.xlsx"
            self.generator.excel_generator.create_dgv_report(soldiers_100, month_display, "100", filename)
            reports.append(filename)
            self._log(f"✓ Створено: {filename}")
        
        soldiers_30 = self.generator.reader.get_soldiers_by_category(soldiers, "30", include_no_payment=False)
        if soldiers_30:
            filename = f"ДГВ_30к_{month_display}.xlsx"
            self.generator.excel_generator.create_dgv_report(soldiers_30, month_display, "30", filename)
            reports.append(filename)
            self._log(f"✓ Створено: {filename}")
        
        soldiers_100_all = self.generator.reader.get_soldiers_by_category(soldiers, "100", include_no_payment=True)
        if soldiers_100_all:
            filename = f"Підтвердження_100к_{month_display}.docx"
            self.generator.word_generator.create_confirmation_report(soldiers_100_all, month_display, "100", filename)
            reports.append(filename)
            self._log(f"✓ Створено: {filename}")
        
        soldiers_30_all = self.generator.reader.get_soldiers_by_category(soldiers, "30", include_no_payment=True)
        if soldiers_30_all:
            filename = f"Підтвердження_30к_{month_display}.docx"
            self.generator.word_generator.create_confirmation_report(soldiers_30_all, month_display, "30", filename)
            reports.append(filename)
            self._log(f"✓ Створено: {filename}")
        
        soldiers_0 = self.generator.reader.get_soldiers_by_category(soldiers, "0", include_no_payment=False)
        if soldiers_0:
            filename = f"ДГВ_0к_{month_display}.xlsx"
            self.generator.excel_generator.create_dgv_report(soldiers_0, month_display, "0", filename)
            reports.append(filename)
            self._log(f"✓ Створено: {filename}")
        
        self._log(f"\n✓ Всього створено файлів: {len(reports)}")
    
    # ===== Екран заповнення табелю =====

    def _show_tabel_filler_screen(self):
        """Показує екран заповнення табелю з місячних файлів"""
        self.current_screen = "tabel_filler"
        self._clear_screen()

        content_frame = tk.Frame(self.root, bg=_BG, relief=tk.FLAT)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # Заголовок
        header_frame = tk.Frame(content_frame, bg=_HEADER_BG, height=60)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        header_frame.pack_propagate(False)

        logo = self._get_logo()
        if logo:
            logo_lbl = tk.Label(header_frame, image=logo, bg=_HEADER_BG)
            logo_lbl.image = logo
            logo_lbl.pack(side=tk.LEFT, padx=10)

        title_label = tk.Label(
            header_frame,
            text="Заповнити табель з місячних файлів",
            font=("Arial", 16, "bold"),
            bg=_HEADER_BG,
            fg="white"
        )
        title_label.pack(pady=15)

        # Основний контент
        main_frame = tk.Frame(content_frame, bg=_BG, padx=30, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Вибір місяця
        month_frame = tk.LabelFrame(main_frame, text="Оберіть місяць", font=("Arial", 11, "bold"), padx=15, pady=15, bg=_BG, fg=_FG)
        month_frame.pack(fill=tk.X, pady=(0, 15))

        months = self.generator.available_months if self.generator else get_available_months(self.excel_file)
        month_values = months + ["-- Всі місяці --"]

        self.tabel_month_var = tk.StringVar()
        tabel_month_combo = ttk.Combobox(month_frame, textvariable=self.tabel_month_var, values=month_values,
                                         state="readonly", width=40, font=("Arial", 11))
        tabel_month_combo.pack(pady=10)
        if months:
            tabel_month_combo.set(months[-1])

        # Статус файлів-джерел
        files_frame = tk.LabelFrame(main_frame, text="Файли-джерела", font=("Arial", 11, "bold"), padx=15, pady=15, bg=_BG, fg=_FG)
        files_frame.pack(fill=tk.X, pady=(0, 15))

        self.files_status_text = tk.Text(
            files_frame, height=4, font=("Consolas", 10), wrap=tk.WORD,
            bg=_BG_INPUT, fg=_FG, relief=tk.FLAT, insertbackground=_FG, padx=10, pady=10
        )
        self.files_status_text.pack(fill=tk.X)

        def on_month_changed(event=None):
            self._update_source_files_status()

        tabel_month_combo.bind("<<ComboboxSelected>>", on_month_changed)
        self._update_source_files_status()

        # Кнопки
        button_frame = tk.Frame(main_frame, bg=_BG)
        button_frame.pack(fill=tk.X, pady=(0, 15))

        self.fill_tabel_btn = tk.Button(
            button_frame,
            text="Заповнити табель",
            command=self._fill_tabel,
            font=("Arial", 12, "bold"),
            bg="#8e44ad",
            fg="white",
            relief=tk.FLAT,
            padx=20,
            pady=10,
            cursor="hand2"
        )
        self.fill_tabel_btn.pack(side=tk.LEFT, padx=5)

        back_btn = tk.Button(
            button_frame,
            text="<- Назад",
            command=self._create_main_menu,
            font=("Arial", 11),
            bg="#95a5a6",
            fg="white",
            relief=tk.FLAT,
            padx=20,
            pady=10,
            cursor="hand2"
        )
        back_btn.pack(side=tk.RIGHT, padx=5)

        # Поле для логів
        log_frame = tk.LabelFrame(main_frame, text="Статус виконання", font=("Arial", 11, "bold"), padx=15, pady=15, bg=_BG, fg=_FG)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(
            log_frame, height=8, font=("Consolas", 9), wrap=tk.WORD,
            bg=_BG_INPUT, fg=_FG, relief=tk.FLAT, insertbackground=_FG
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Статус-бар
        self.status_label = tk.Label(
            content_frame, text="Готово до роботи",
            relief=tk.SUNKEN, anchor=tk.W, font=("Arial", 9), bg=_STATUS_BG, fg=_STATUS_FG
        )
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

    def _update_source_files_status(self):
        """Оновлює відображення статусу файлів-джерел"""
        if not hasattr(self, 'files_status_text') or not hasattr(self, 'tabel_month_var'):
            return

        self.files_status_text.config(state=tk.NORMAL)
        self.files_status_text.delete("1.0", tk.END)

        selected = self.tabel_month_var.get()
        if selected == "-- Всі місяці --":
            self.files_status_text.insert("1.0", "Буде оброблено всі доступні місяці")
            self.files_status_text.config(state=tk.DISABLED)
            return

        if not selected:
            self.files_status_text.config(state=tk.DISABLED)
            return

        source_file = get_source_filename(selected)
        source_path = os.path.join(get_app_dir(), source_file)
        exists = os.path.exists(source_path)
        status = "[+] знайдено" if exists else "[-] не знайдено"
        self.files_status_text.insert("1.0", f"{source_file}  {status}")
        self.files_status_text.config(state=tk.DISABLED)

    def _fill_tabel(self):
        """Запускає заповнення табелю"""
        selected = self.tabel_month_var.get()
        if not selected:
            messagebox.showwarning("Увага", "Оберіть місяць!")
            return

        if not messagebox.askyesno("Підтвердження",
                                    f"Заповнити табель за {selected}?"):
            return

        self.fill_tabel_btn.config(state=tk.DISABLED, text="Заповнення...")
        self.log_text.delete(1.0, tk.END)

        thread = threading.Thread(target=self._do_fill_tabel, args=(selected,))
        thread.daemon = True
        thread.start()

    def _do_fill_tabel(self, selected_month):
        """Виконує заповнення табелю"""
        try:
            self._update_status("Заповнення табелю...")
            self._log("=" * 60)
            self._log(f"Заповнення табелю: {selected_month}")
            self._log("=" * 60)

            import sys
            from io import StringIO
            old_stdout = sys.stdout
            sys.stdout = StringIO()

            try:
                if selected_month == "-- Всі місяці --":
                    fill_tabel_months(self.excel_file)
                else:
                    parsed = parse_month_sheet_name(selected_month)
                    if not parsed:
                        raise ValueError(f"Не вдалося розпарсити назву: {selected_month}")
                    year, month_num = parsed
                    source_file = get_source_filename(selected_month)
                    source_path = os.path.join(get_app_dir(), source_file)
                    fill_single_month(selected_month, source_path, year, month_num, self.excel_file)

                output = sys.stdout.getvalue()
                self._log(output)
            finally:
                sys.stdout = old_stdout

            self._log("=" * 60)
            self._log("Заповнення завершено!")
            self._update_status("Готово!")

            self.root.after(0, lambda: messagebox.showinfo(
                "Успіх", "Табель успішно заповнено!"
            ))

        except Exception as e:
            error_msg = f"Помилка: {str(e)}"
            self._log(f"{error_msg}")
            self._update_status("Помилка!")
            self.root.after(0, lambda: messagebox.showerror("Помилка", error_msg))
        finally:
            self.root.after(0, lambda: self.fill_tabel_btn.config(
                state=tk.NORMAL, text="Заповнити табель"
            ))

    # ===== Діалог додавання місяця =====

    def _show_add_month_dialog(self):
        """Показує діалог додавання нового місяця"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Додати новий місяць")
        dialog.geometry("400x280")
        dialog.resizable(False, False)
        dialog.configure(bg=_BG)

        dialog.transient(self.root)
        dialog.grab_set()

        # Заголовок
        header = tk.Label(dialog, text="Додати новий місяць", font=("Arial", 14, "bold"),
                          bg=_HEADER_BG, fg="white", pady=10)
        header.pack(fill=tk.X)

        form_frame = tk.Frame(dialog, bg=_BG, padx=30, pady=20)
        form_frame.pack(fill=tk.BOTH, expand=True)

        # Рік
        year_frame = tk.Frame(form_frame, bg=_BG)
        year_frame.pack(fill=tk.X, pady=10)
        tk.Label(year_frame, text="Рік:", font=("Arial", 11), bg=_BG, fg=_FG, width=10, anchor="w").pack(side=tk.LEFT)
        year_var = tk.StringVar(value=str(datetime.now().year))
        year_spin = tk.Spinbox(year_frame, from_=2025, to=2030, textvariable=year_var,
                               font=("Arial", 11), width=10)
        year_spin.pack(side=tk.LEFT, padx=10)

        # Місяць
        month_frame = tk.Frame(form_frame, bg=_BG)
        month_frame.pack(fill=tk.X, pady=10)
        tk.Label(month_frame, text="Місяць:", font=("Arial", 11), bg=_BG, fg=_FG, width=10, anchor="w").pack(side=tk.LEFT)
        month_names_list = [MONTH_NAMES_UK_REVERSE[i] for i in range(1, 13)]
        month_var = tk.StringVar()
        month_combo = ttk.Combobox(month_frame, textvariable=month_var, values=month_names_list,
                                   state="readonly", font=("Arial", 11), width=15)
        month_combo.pack(side=tk.LEFT, padx=10)

        # Підказка поточного місяця
        current_month = datetime.now().month
        month_combo.set(MONTH_NAMES_UK_REVERSE[current_month])

        # Кнопка створення
        def do_create():
            if not month_var.get():
                messagebox.showwarning("Увага", "Оберіть місяць!", parent=dialog)
                return

            year = int(year_var.get())
            month_name = month_var.get()
            sheet_name = f"{month_name}_{year}"

            try:
                wb = openpyxl.load_workbook(self.excel_file)

                if sheet_name in wb.sheetnames:
                    messagebox.showwarning("Увага", f"Аркуш '{sheet_name}' вже існує!", parent=dialog)
                    wb.close()
                    return

                # Копіюємо останній аркуш як шаблон
                available = [s for s in wb.sheetnames if parse_month_sheet_name(s)]
                if available:
                    template_sheet = wb[available[-1]]
                    new_sheet = wb.copy_worksheet(template_sheet)
                    new_sheet.title = sheet_name

                    # Очищаємо дані (рядки з 9-го)
                    for row in range(9, new_sheet.max_row + 1):
                        for col in range(1, new_sheet.max_column + 1):
                            new_sheet.cell(row, col).value = None
                else:
                    wb.create_sheet(sheet_name)

                wb.save(self.excel_file)
                wb.close()

                messagebox.showinfo("Успіх", f"Аркуш '{sheet_name}' створено!", parent=dialog)

                # Оновлюємо генератор щоб підтягнути новий місяць
                try:
                    self.generator = ReportGenerator(self.excel_file)
                except Exception:
                    pass

                dialog.destroy()

            except Exception as e:
                messagebox.showerror("Помилка", f"Не вдалося створити аркуш:\n{str(e)}", parent=dialog)

        btn_frame = tk.Frame(form_frame, bg=_BG)
        btn_frame.pack(fill=tk.X, pady=20)

        create_btn = tk.Button(
            btn_frame, text="Створити", command=do_create,
            font=("Arial", 12, "bold"), bg="#16a085", fg="white",
            relief=tk.FLAT, padx=20, pady=8, cursor="hand2"
        )
        create_btn.pack(side=tk.LEFT, padx=5)

        cancel_btn = tk.Button(
            btn_frame, text="Скасувати", command=dialog.destroy,
            font=("Arial", 11), bg="#95a5a6", fg="white",
            relief=tk.FLAT, padx=20, pady=8, cursor="hand2"
        )
        cancel_btn.pack(side=tk.RIGHT, padx=5)



    # ==================== ЕКРАН РОЛЕЙ ====================

    def _show_roles_screen(self):
        """Показує екран призначення ролей для БР"""
        self.current_screen = "roles"
        self._clear_screen()

        content_frame = tk.Frame(self.root, bg=_BG, relief=tk.FLAT)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # Заголовок
        header_frame = tk.Frame(content_frame, bg=_HEADER_BG, height=60)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        header_frame.pack_propagate(False)

        logo = self._get_logo()
        if logo:
            logo_lbl = tk.Label(header_frame, image=logo, bg=_HEADER_BG)
            logo_lbl.image = logo
            logo_lbl.pack(side=tk.LEFT, padx=10)

        tk.Label(
            header_frame, text="Ролі для БР",
            font=("Arial", 16, "bold"), bg=_HEADER_BG, fg="white"
        ).pack(pady=15)

        # Панель дій
        actions_frame = tk.Frame(content_frame, bg=_BG, padx=15, pady=5)
        actions_frame.pack(fill=tk.X)

        # Вибір місяця + імпорт
        import_frame = tk.Frame(actions_frame, bg=_BG)
        import_frame.pack(fill=tk.X, pady=5)

        tk.Label(import_frame, text="Місяць:", font=("Arial", 10), bg=_BG, fg=_FG).pack(side=tk.LEFT, padx=(0, 5))

        self.roles_month_var = tk.StringVar()
        months_combo = ttk.Combobox(
            import_frame, textvariable=self.roles_month_var,
            state="readonly", width=25, font=("Arial", 10)
        )
        try:
            months = get_available_months(self.excel_file)
            months_combo['values'] = months
            if months:
                months_combo.current(len(months) - 1)
        except Exception:
            months_combo['values'] = []
        months_combo.pack(side=tk.LEFT, padx=5)

        import_btn = tk.Button(
            import_frame, text="📥 Імпорт з табеля",
            command=self._import_from_tabel_action,
            font=("Arial", 10, "bold"), bg="#3498db", fg="white",
            relief=tk.FLAT, padx=10, pady=5, cursor="hand2"
        )
        import_btn.pack(side=tk.LEFT, padx=10)

        auto_btn = tk.Button(
            import_frame, text="⚙ Автопризначити ролі",
            command=self._auto_assign_roles_action,
            font=("Arial", 10, "bold"), bg="#e67e22", fg="white",
            relief=tk.FLAT, padx=10, pady=5, cursor="hand2"
        )
        auto_btn.pack(side=tk.LEFT, padx=5)

        back_btn = tk.Button(
            import_frame, text="← Назад",
            command=self._create_main_menu,
            font=("Arial", 10), bg="#95a5a6", fg="white",
            relief=tk.FLAT, padx=10, pady=5, cursor="hand2"
        )
        back_btn.pack(side=tk.RIGHT, padx=5)

        # Treeview з ролями
        tree_frame = tk.Frame(content_frame, bg=_BG, padx=15, pady=5)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        columns = ("num", "pib", "rank", "position", "role")
        self.roles_tree = ttk.Treeview(
            tree_frame, columns=columns, show="headings", height=18
        )
        self.roles_tree.heading("num", text="#")
        self.roles_tree.heading("pib", text="ПІБ")
        self.roles_tree.heading("rank", text="Звання")
        self.roles_tree.heading("position", text="Посада")
        self.roles_tree.heading("role", text="Роль")

        self.roles_tree.column("num", width=40, stretch=False)
        self.roles_tree.column("pib", width=220)
        self.roles_tree.column("rank", width=120)
        self.roles_tree.column("position", width=180)
        self.roles_tree.column("role", width=200)

        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.roles_tree.yview)
        self.roles_tree.configure(yscrollcommand=scrollbar.set)

        self.roles_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Подвійний клік для зміни ролі
        self.roles_tree.bind("<Double-1>", self._on_role_cell_click)

        # Зберігаємо дані ролей для combobox
        self._roles_list = get_all_roles()  # [(id, name), ...]
        self._roles_combo_widget = None

        # Лог внизу
        log_frame = tk.LabelFrame(content_frame, text="Статус", font=("Arial", 10), padx=10, pady=5, bg=_BG, fg=_FG)
        log_frame.pack(fill=tk.X, padx=15, pady=(0, 5))

        self.log_text = scrolledtext.ScrolledText(
            log_frame, height=4, font=("Consolas", 9), wrap=tk.WORD,
            bg=_BG_INPUT, fg=_FG, relief=tk.FLAT, insertbackground=_FG
        )
        self.log_text.pack(fill=tk.X)

        # Статус-бар
        self.status_label = tk.Label(
            content_frame, text="Подвійний клік по колонці 'Роль' для зміни",
            relief=tk.SUNKEN, anchor=tk.W, font=("Arial", 9), bg=_STATUS_BG, fg=_STATUS_FG
        )
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

        # Завантажуємо дані
        self._refresh_roles_treeview()

    def _refresh_roles_treeview(self):
        """Оновлює Treeview з даними персоналу."""
        if not hasattr(self, 'roles_tree'):
            return
        for item in self.roles_tree.get_children():
            self.roles_tree.delete(item)

        personnel = get_all_personnel()
        for i, p in enumerate(personnel, 1):
            role_display = p["role_name"] or "— не призначено —"
            self.roles_tree.insert("", tk.END, iid=p["pib"], values=(
                i, p["pib"], p["rank"], p["position"], role_display
            ))

    def _on_role_cell_click(self, event):
        """Показує overlay Combobox для зміни ролі при подвійному кліку."""
        # Визначаємо колонку
        region = self.roles_tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        column = self.roles_tree.identify_column(event.x)
        if column != "#5":  # 5-та колонка = role
            return

        item_id = self.roles_tree.identify_row(event.y)
        if not item_id:
            return

        # Видаляємо попередній combobox якщо є
        if self._roles_combo_widget:
            self._roles_combo_widget.destroy()
            self._roles_combo_widget = None

        # Координати клітинки
        bbox = self.roles_tree.bbox(item_id, column)
        if not bbox:
            return
        x, y, w, h = bbox

        # Створюємо combobox
        role_names = ["— не призначено —"] + [name for _, name in self._roles_list]
        combo = ttk.Combobox(
            self.roles_tree, values=role_names, state="readonly",
            font=("Arial", 9)
        )

        # Поточне значення
        current_values = self.roles_tree.item(item_id, "values")
        current_role = current_values[4] if len(current_values) > 4 else ""
        if current_role in role_names:
            combo.set(current_role)
        else:
            combo.set("— не призначено —")

        combo.place(x=x, y=y, width=w, height=h)
        combo.focus_set()

        pib = item_id  # iid = pib

        def on_select(ev):
            selected = combo.get()
            if selected == "— не призначено —":
                set_personnel_role(pib, None)
            else:
                role_id = None
                for rid, rname in self._roles_list:
                    if rname == selected:
                        role_id = rid
                        break
                if role_id is not None:
                    set_personnel_role(pib, role_id)
            combo.destroy()
            self._roles_combo_widget = None
            self._refresh_roles_treeview()

        def on_focus_out(ev):
            combo.destroy()
            self._roles_combo_widget = None

        combo.bind("<<ComboboxSelected>>", on_select)
        combo.bind("<FocusOut>", on_focus_out)
        self._roles_combo_widget = combo

    def _import_from_tabel_action(self):
        """Імпортує особовий склад з табеля у БД."""
        month = self.roles_month_var.get()
        if not month:
            messagebox.showwarning("Увага", "Оберіть місяць!")
            return

        self._log(f"Імпорт з аркуша '{month}'...")
        self._update_status("Імпорт...")

        def do_import():
            try:
                count = import_personnel_from_tabel(self.excel_file, month)
                self.root.after(0, lambda: self._log(f"Імпортовано {count} записів."))
                self.root.after(0, self._refresh_roles_treeview)
                self.root.after(0, lambda: self._update_status(f"Імпорт завершено: {count} записів"))
            except Exception as e:
                self.root.after(0, lambda: self._log(f"ПОМИЛКА: {e}"))
                self.root.after(0, lambda: self._update_status("Помилка імпорту"))

        threading.Thread(target=do_import, daemon=True).start()

    def _auto_assign_roles_action(self):
        """Автопризначає ролі по ключових словах у посаді."""
        self._log("Автопризначення ролей...")
        self._update_status("Автопризначення...")

        def do_assign():
            try:
                stats = auto_assign_all_roles()
                total = sum(stats.values())
                self.root.after(0, lambda: self._log(f"Автопризначено {total} ролей:"))
                for role_name, count in stats.items():
                    self.root.after(0, lambda rn=role_name, c=count: self._log(f"  {rn}: {c}"))
                if not stats:
                    self.root.after(0, lambda: self._log("  Немає нових призначень (усі вже мають ролі або не відповідають правилам)."))
                self.root.after(0, self._refresh_roles_treeview)
                self.root.after(0, lambda: self._update_status(f"Автопризначено: {total}"))
            except Exception as e:
                self.root.after(0, lambda: self._log(f"ПОМИЛКА: {e}"))
                self.root.after(0, lambda: self._update_status("Помилка"))

        threading.Thread(target=do_assign, daemon=True).start()

    # ==================== МЕТОДИ ДЛЯ WORD БР ====================

    def _preview_composition(self):
        """Показує прев'ю складу БР по ролях (тільки mark==100)."""
        date_str = self.start_date_var.get().strip()
        if not date_str:
            messagebox.showwarning("Увага", "Введіть дату БР!")
            return

        try:
            br_date = datetime.strptime(date_str, "%d.%m.%Y")
        except ValueError:
            messagebox.showerror("Помилка", "Некоректний формат дати (ДД.ММ.РРРР)")
            return

        self.log_text.delete(1.0, tk.END)
        self._log(f"Формування складу на дату БР: {date_str}")
        from datetime import timedelta as _td
        tabel_date = br_date + _td(days=1)
        self._log(f"Дата табеля (БР+1): {tabel_date.strftime('%d.%m.%Y')}")
        self._log("")

        def do_preview():
            try:
                composition = build_composition_for_date(self.excel_file, br_date)
                total = 0
                for role_name, members in composition.items():
                    count = len(members)
                    total += count
                    self.root.after(0, lambda rn=role_name, c=count: self._log(f"--- {rn} ({c}) ---"))
                    if members:
                        for m in members:
                            self.root.after(0, lambda mm=m: self._log(f"  {mm['rank']} {mm['pib']}"))
                    else:
                        self.root.after(0, lambda: self._log("  (порожньо)"))
                    self.root.after(0, lambda: self._log(""))
                self.root.after(0, lambda: self._log(f"Всього: {total} осіб з відміткою 100"))
                self.root.after(0, lambda: self._update_status(f"Склад сформовано: {total} осіб"))
            except Exception as e:
                self.root.after(0, lambda: self._log(f"ПОМИЛКА: {e}"))
                self.root.after(0, lambda: self._update_status("Помилка"))

        threading.Thread(target=do_preview, daemon=True).start()

    def _generate_word_br(self):
        """Генерує Word-документи БР для кожної дати в діапазоні."""
        start_str = self.start_date_var.get().strip()
        end_str = self.end_date_var.get().strip()
        if not start_str or not end_str:
            messagebox.showwarning("Увага", "Введіть початкову та кінцеву дату!")
            return

        try:
            start_date = datetime.strptime(start_str, "%d.%m.%Y")
            end_date = datetime.strptime(end_str, "%d.%m.%Y")
        except ValueError:
            messagebox.showerror("Помилка", "Некоректний формат дати (ДД.ММ.РРРР)")
            return

        if start_date > end_date:
            messagebox.showerror("Помилка", "Початкова дата не може бути пізнішою за кінцеву!")
            return

        if not os.path.exists(self.template_path):
            messagebox.showerror("Помилка", f"Шаблон не знайдено: {self.template_path}")
            return

        self.log_text.delete(1.0, tk.END)
        self._log(f"Генерація Word БР: {start_str} — {end_str}")

        def do_generate():
            from datetime import timedelta
            try:
                created = 0
                current = start_date
                while current <= end_date:
                    ds = current.strftime("%d.%m.%Y")
                    self.root.after(0, lambda d=ds: self._log(f"\n--- БР на {d} ---"))
                    composition = build_composition_for_date(self.excel_file, current)
                    total = sum(len(m) for m in composition.values())
                    self.root.after(0, lambda t=total: self._log(f"  Осіб з роллю: {t}"))

                    result_path = generate_br_word(
                        current, composition, self.template_path, self.output_dir,
                        br_4shb_file=self.br_4shb_file
                    )
                    self.root.after(0, lambda p=result_path: self._log(f"  Створено: {p}"))
                    created += 1
                    current += timedelta(days=1)

                self.root.after(0, lambda: self._log(f"\nВсього створено {created} файлів БР"))
                self.root.after(0, lambda: self._update_status(f"Створено {created} БР"))
                self.root.after(0, lambda: messagebox.showinfo(
                    "Готово", f"Створено {created} файлів БР"
                ))
            except Exception as e:
                self.root.after(0, lambda: self._log(f"ПОМИЛКА: {e}"))
                self.root.after(0, lambda: self._update_status("Помилка генерації"))

        threading.Thread(target=do_generate, daemon=True).start()


def main():
    """Запускає графічний інтерфейс"""
    root = tk.Tk()
    app = ReportGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

