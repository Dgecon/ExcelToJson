from tkinter import *
from tkinter import ttk
from tkinter import filedialog, messagebox
import json
import os
from datetime import datetime
import pandas as pd
import subprocess
import platform
from screeninfo import get_monitors

# Глобальные переменные
ask_window = None
current_file_path = None
text_widget = None
line_numbers_widget = None  # Виджет для номеров строк
status_label = None
editor_win = None
converter_win = None
dark_theme_enabled = False  # Глобальная переменная для темы
cyrillic_highlight_enabled = False  # Состояние подсветки кириллицы


def get_windows_theme():
    """Определяет тему Windows (светлая/темная)"""
    if platform.system() != 'Windows':
        return False  # По умолчанию светлая тема для не-Windows
    
    try:
        import winreg
        # Путь к реестру Windows для темы
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        )
        # AppsUseLightTheme: 0 = темная тема, 1 = светлая тема
        value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        winreg.CloseKey(key)
        return value == 0  # True если темная тема
    except Exception:
        return False  # По умолчанию светлая тема при ошибке

def apply_theme(window, is_dark):
    """Применяет тему к окну и всем его виджетам"""
    if is_dark:
        bg_color = "#2b2b2b"
        fg_color = "#ffffff"
        entry_bg = "#3c3c3c"
        entry_fg = "#ffffff"
        status_bg = "#404040"
        text_bg = "#1e1e1e"
        text_fg = "#ffffff"
        frame_bg = "#2b2b2b"
        button_bg = "#3c3c3c"
        button_fg = "#ffffff"
        button_active_bg = "#4a4a4a"
    else:
        bg_color = "#f9f9f9"
        fg_color = "#000000"
        entry_bg = "white"
        entry_fg = "black"
        status_bg = "white"
        text_bg = "white"
        text_fg = "black"
        frame_bg = "#f9f9f9"
        button_bg = "#e0e0e0"
        button_fg = "#000000"
        button_active_bg = "#d0d0d0"
    
    # Применяем к окну
    window.configure(bg=bg_color)
    
    # Рекурсивно применяем ко всем виджетам
    def apply_to_widgets(widget):
        widget_type = widget.winfo_class()
        
        if widget_type == "Label":
            try:
                # Проверяем, не является ли это статусной меткой (имеет relief=SUNKEN)
                relief = widget.cget("relief")
                if relief == "sunken":
                    # Это статусная метка - используем специальный цвет
                    widget.configure(bg=status_bg, fg=fg_color)
                else:
                    widget.configure(bg=bg_color, fg=fg_color)
            except:
                pass
        elif widget_type == "Frame":
            try:
                widget.configure(bg=frame_bg)
            except:
                pass
        elif widget_type == "Text":
            try:
                # Проверяем, является ли это виджетом номеров строк
                current_bg = widget.cget("bg")
                if current_bg in ["#252526", "#f0f0f0"]:
                    # Это виджет номеров строк
                    line_num_bg = "#252526" if is_dark else "#f0f0f0"
                    line_num_fg = "#858585" if is_dark else "#666666"
                    widget.configure(bg=line_num_bg, fg=line_num_fg)
                elif current_bg in ["white", "#1e1e1e"]:
                    # Это обычное текстовое поле
                    widget.configure(bg=text_bg, fg=text_fg, insertbackground=fg_color)
            except:
                pass
        elif widget_type == "Entry":
            try:
                widget.configure(bg=entry_bg, fg=entry_fg, insertbackground=fg_color)
            except:
                pass
        elif widget_type == "Button":
            try:
                widget.configure(bg=button_bg, fg=button_fg, 
                               activebackground=button_active_bg, 
                               activeforeground=button_fg,
                               relief=RAISED,
                               borderwidth=1)
            except:
                pass
        elif widget_type == "Checkbutton":
            try:
                # Для темной темы: темный фон, белый текст, темный цвет галочки
                # Для светлой темы: светлый фон, черный текст, белый цвет галочки
                checkbox_bg = frame_bg
                checkbox_fg = fg_color
                checkbox_selectcolor = "#3c3c3c" if is_dark else "white"
                widget.configure(bg=checkbox_bg, fg=checkbox_fg,
                               activebackground=checkbox_bg,
                               activeforeground=checkbox_fg,
                               selectcolor=checkbox_selectcolor)
            except:
                pass
        
        # Рекурсивно обрабатываем дочерние виджеты
        for child in widget.winfo_children():
            apply_to_widgets(child)
    
    apply_to_widgets(window)

def place_window_near_cursor(window, width, height, dx=0, dy=0, screen_margin=20):
    # Получаем координаты курсора
    x, y = window.winfo_pointerxy()

    # Находим монитор, на котором находится курсор
    target_monitor = None
    for monitor in get_monitors():
        if monitor.x <= x <= monitor.x + monitor.width and \
           monitor.y <= y <= monitor.y + monitor.height:
            target_monitor = monitor
            break

    # Если не нашли — используем первый монитор как fallback
    if not target_monitor:
        target_monitor = get_monitors()[0]

    # Вычисляем позицию окна относительно курсора
    win_x = x + dx
    win_y = y + dy

    # Ограничиваем позицию в пределах монитора с учётом отступов
    left_bound = target_monitor.x + screen_margin
    right_bound = target_monitor.x + target_monitor.width - width - screen_margin
    top_bound = target_monitor.y + screen_margin
    bottom_bound = target_monitor.y + target_monitor.height - height - screen_margin

    # Применяем ограничения
    win_x = max(left_bound, min(win_x, right_bound))
    win_y = max(top_bound, min(win_y, bottom_bound))

    # Устанавливаем геометрию окна
    window.geometry(f"{width}x{height}+{win_x}+{win_y}")

def copy_to_clipboard(text):
    """Копирует текст в буфер обмена (кроссплатформенный метод)"""
    try:
        # Пробуем использовать pyperclip если установлен
        import pyperclip
        pyperclip.copy(text)
    except ImportError:
        # Альтернативный способ для Windows
        if platform.system() == 'Windows':
            try:
                # Используем команду PowerShell для копирования
                process = subprocess.Popen(
                    ['powershell', '-command', f'Set-Clipboard -Value @\"\n{text}\n\"@'],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE
                )
                process.communicate()
            except Exception as e:
                # Если не получилось, создаем временный файл и копируем через cmd
                try:
                    import tempfile
                    with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.txt', encoding='utf-8') as f:
                        f.write(text)
                        temp_path = f.name
                    subprocess.run(['cmd', '/c', f'type "{temp_path}" | clip'], check=True)
                    os.unlink(temp_path)
                except Exception:
                    raise Exception("Не удалось скопировать в буфер обмена. Установите pyperclip: pip install pyperclip")
        else:
            # Для Linux/Mac
            try:
                subprocess.run(['xclip', '-selection', 'clipboard'], input=text.encode('utf-8'), check=True)
            except:
                raise Exception("Не удалось скопировать в буфер обмена. Установите pyperclip: pip install pyperclip")

def convert_value_by_type(value, data_type):
    """Конвертирует значение согласно типу данных"""
    if pd.isna(value) or value == '':
        return None
    
    # Приводим тип данных к строке и убираем пробелы
    data_type = str(data_type).strip().lower() if not pd.isna(data_type) else 'string'
    
    # Обработка по типу данных
    if data_type in ['number', 'int', 'integer', 'число', 'числовой']:
        try:
            # Пытаемся преобразовать в число
            if '.' in str(value):
                return float(value)
            else:
                return int(value)
        except (ValueError, TypeError):
            # Если не получилось, возвращаем как строку
            return str(value)
    elif data_type in ['bool', 'boolean', 'логический']:
        value_str = str(value).strip().lower()
        if value_str in ['true', '1', 'да', 'yes', 'истина']:
            return True
        elif value_str in ['false', '0', 'нет', 'no', 'ложь']:
            return False
        else:
            return bool(value)
    elif data_type in ['null', 'none', 'пусто']:
        return None
    else:
        # Строковый тип - экранируем специальные символы
        return str(value)

def convert_excel_to_json(excel_path, status_label=None):
    """Конвертирует Excel файл в JSON согласно ТЗ:
    - Столбец A: ключи
    - Столбец B: значения
    - Столбец C: типы данных
    - Первая строка пропускается (заголовки)
    """
    try:
        if status_label:
            status_label.config(text="Чтение файла...", fg="blue")
            status_label.update()
        
        # Читаем Excel файл без заголовков (header=None), так как первая строка - это заголовки таблицы
        # Читаем только столбцы A (0), B (1), C (2)
        df = pd.read_excel(excel_path, header=None, usecols=[0, 1, 2])
        
        # Создаем словарь для результата
        result = {}
        
        # Пропускаем первую строку (индекс 0 - это заголовки) и обрабатываем остальные
        for idx in range(1, len(df)):
            key = df.iloc[idx, 0]  # Столбец A - ключ
            value = df.iloc[idx, 1]  # Столбец B - значение
            data_type = df.iloc[idx, 2] if df.shape[1] > 2 else None  # Столбец C - тип данных
            
            # Пропускаем пустые ключи
            if pd.isna(key) or str(key).strip() == '':
                continue
            
            # Конвертируем значение согласно типу
            converted_value = convert_value_by_type(value, data_type)
            
            # Добавляем в результат
            result[str(key).strip()] = converted_value
        
        if status_label:
            status_label.config(text="Формирование JSON...", fg="blue")
            status_label.update()
        
        # Формируем JSON строку
        json_str = json.dumps(result, ensure_ascii=False, indent=2)
        
        # Определяем путь для сохранения (рядом с исходным файлом)
        folder = os.path.dirname(excel_path)
        base_name = os.path.splitext(os.path.basename(excel_path))[0]
        timestamp = datetime.now().strftime("%Y.%m.%d_%H-%M")
        save_path = os.path.join(folder, f"{base_name}_{timestamp}.json")
        
        if status_label:
            status_label.config(text="Сохранение файла...", fg="blue")
            status_label.update()
        
        # Сохраняем в JSON
        with open(save_path, 'w', encoding='utf-8') as f:
            f.write(json_str)
        
        if status_label:
            status_label.config(text="Копирование в буфер обмена...", fg="blue")
            status_label.update()
        
        # Копируем JSON в буфер обмена
        try:
            copy_to_clipboard(json_str)
        except Exception as e:
            print(f"Не удалось скопировать в буфер обмена: {e}")
        
        if status_label:
            status_label.config(text="✅ Конвертация завершена!(json строка скопирована)", fg="green")
            status_label.update()
        
        return save_path, json_str
    except Exception as e:
        error_msg = f"Не удалось конвертировать файл:\n{e}"
        if status_label:
            status_label.config(text=f"❌ Ошибка: {str(e)}", fg="red")
        messagebox.showerror("Ошибка конвертации", error_msg)
        return None, None

def select_excel_file():
    """Открывает диалог выбора Excel файла и конвертирует его"""
    global converter_win, status_label
    
    filepath = filedialog.askopenfilename(
        title="Выберите Excel файл для конвертации",
        filetypes=[
            ("Файлы Excel", "*.xl*;*.xlsx;*.xlsm;*.xls"),
            ("XLSX files", "*.xlsx"),
            ("XLSM files", "*.xlsm"),
            ("XLS files", "*.xls"),
            ("All files", "*.*")
        ]
    )
    if filepath:
        save_path, json_str = convert_excel_to_json(filepath, status_label)
        if save_path:
            messagebox.showinfo("Успех", 
                f"Файл успешно конвертирован!\n\n"
                f"Сохранено: {save_path}\n\n"
                f"JSON скопирован в буфер обмена.")

def start_xls2json_win():
    """Создает окно конвертации Excel в JSON"""
    global ask_window, converter_win, status_label, dark_theme_enabled
    
    if ask_window:
        ask_window.destroy()
        ask_window = None
    
    converter_win = Tk()
    converter_win.title("Конвертация Excel → JSON")
    converter_win.resizable(False, False)
    place_window_near_cursor(converter_win, 450, 220, 0, 0, 200)
    
    Label(converter_win, text="Конвертация Excel в JSON", 
          font=("Segoe UI", 12, "bold")).pack(pady=15)
    
    # Поле статуса (Jobizdan)
    status_bg = "#404040" if dark_theme_enabled else "white"
    status_fg = "#ffffff" if dark_theme_enabled else "gray"
    status_label = Label(converter_win, text="Готов к конвертации", 
                        relief=SUNKEN, anchor=W, bg=status_bg, fg=status_fg, 
                        font=("Segoe UI", 9))
    status_label.pack(fill=X, padx=10, pady=5)
    
    btn_frame = Frame(converter_win)
    btn_frame.pack(pady=10)
    
    Button(btn_frame, text="Выбрать файл и конвертировать", 
           command=select_excel_file, width=35).pack(pady=5)
    
    btn_frame2 = Frame(converter_win)
    btn_frame2.pack(pady=5)
    
    Button(btn_frame2, text="Справка", 
           command=show_help, width=15).pack(side=LEFT, padx=5)
    Button(btn_frame2, text="Назад", 
           command=lambda: go_back_to_main(converter_win), width=15).pack(side=LEFT, padx=5)
    
    # Применяем тему после создания всех виджетов
    apply_theme(converter_win, dark_theme_enabled)
    
    # Горячая клавиша для выбора файла
    converter_win.bind('<Control-o>', lambda e: select_excel_file())
    converter_win.bind('<Return>', lambda e: select_excel_file())
    
    converter_win.mainloop()

def show_help():
    """Открывает окно со справкой"""
    global dark_theme_enabled
    help_window = Toplevel()
    help_window.title("Справка")
    help_window.resizable(True, True)
    help_window.geometry("700x600")
    
    # Создаем текстовое поле с прокруткой
    text_frame = Frame(help_window)
    text_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)
    
    scrollbar = Scrollbar(text_frame)
    scrollbar.pack(side=RIGHT, fill=Y)
    
    text_bg = "#1e1e1e" if dark_theme_enabled else "white"
    text_fg = "#ffffff" if dark_theme_enabled else "black"
    help_text = Text(text_frame, wrap=WORD, font=("Segoe UI", 10), 
                     yscrollcommand=scrollbar.set, bg=text_bg, fg=text_fg,
                     padx=10, pady=10, insertbackground=text_fg)
    scrollbar.config(command=help_text.yview)
    help_text.pack(side=LEFT, fill=BOTH, expand=True)
    
    # Содержимое справки
    help_content = """
СПРАВКА ПО ИСПОЛЬЗОВАНИЮ УТИЛИТЫ XLSX/JSON HELPER

═══════════════════════════════════════════════════════════════

1. КОНВЕРТАЦИЯ EXCEL В JSON
───────────────────────────────────────────────────────────────

Структура Excel файла:
  • Столбец A: Ключи (названия полей JSON)
  • Столбец B: Значения (данные для полей)
  • Столбец C: Типы данных (опционально)
  • Первая строка: Заголовки (пропускается при обработке)

Поддерживаемые типы данных в столбце C:
  • number, int, integer, число, числовой - числа без кавычек
  • bool, boolean, логический - логические значения (true/false)
  • null, none, пусто - пустые значения
  • string, текст (или пусто) - строковые значения с экранированием

Пример структуры Excel:
  ┌─────────────┬──────────────┬──────────────┐
  │ Заголовок 1 │ Заголовок 2  │ Заголовок 3  │
  ├─────────────┼──────────────┼──────────────┤
  │ name        │ Иван         │ string       │
  │ age         │ 25           │ number       │
  │ active      │ true         │ bool         │
  └─────────────┴──────────────┴──────────────┘

Результат JSON:
  {
    "name": "Иван",
    "age": 25,
    "active": true
  }

Процесс конвертации:
  1. Нажмите "Выбрать файл и конвертировать"
  2. Выберите Excel файл (*.xlsx, *.xlsm, *.xls)
  3. Дождитесь завершения конвертации
  4. JSON файл будет сохранен рядом с исходным файлом
  5. JSON строка автоматически скопируется в буфер обмена

Горячие клавиши:
  • Ctrl+O - открыть диалог выбора файла
  • Enter - выбрать файл и конвертировать

═══════════════════════════════════════════════════════════════

2. РЕДАКТОР JSON С ПРОВЕРКОЙ СИНТАКСИСА
───────────────────────────────────────────────────────────────

Функции редактора:
  • Открытие JSON файлов для редактирования
  • Автоматическая проверка синтаксиса JSON
  • Подсветка строк с ошибками
  • Сохранение отредактированных файлов

Использование:
  1. Нажмите "Открыть файл"
  2. Выберите JSON файл
  3. Отредактируйте содержимое
  4. Нажмите "Проверить" для проверки синтаксиса
  5. Нажмите "Сохранить" для сохранения изменений

Горячие клавиши:
  • Ctrl+O - открыть файл
  • Ctrl+S - сохранить файл
  • F5 - проверить синтаксис JSON

Статусная строка показывает:
  • ✅ Корректный JSON - файл валиден
  • ❌ Ошибка в строке X - найдена ошибка
  • Файл пуст - файл не содержит данных

═══════════════════════════════════════════════════════════════

3. ОБРАБОТКА СПЕЦИАЛЬНЫХ СИМВОЛОВ
───────────────────────────────────────────────────────────────

Программа автоматически экранирует специальные символы в строках:
  • Кавычки (") → \\"
  • Обратный слэш (\\) → \\\\
  • Переносы строк → \\n
  • Табуляции → \\t
  • И другие управляющие символы

Пример:
  Входная строка: Привет "мир"!
  JSON результат: "Привет \\"мир\\"!"

═══════════════════════════════════════════════════════════════

4. ТРЕБОВАНИЯ К ФАЙЛАМ
───────────────────────────────────────────────────────────────

Поддерживаемые форматы Excel:
  • .xlsx (Excel 2007 и новее)
  • .xlsm (Excel с макросами)
  • .xls (Excel 97-2003)

Кодировка:
  • Все файлы обрабатываются в кодировке UTF-8
  • Поддержка кириллицы и других Unicode символов

═══════════════════════════════════════════════════════════════

5. РЕШЕНИЕ ПРОБЛЕМ
───────────────────────────────────────────────────────────────

Проблема: "Не удалось конвертировать файл"
  • Убедитесь, что файл не открыт в другой программе
  • Проверьте, что файл имеет правильный формат Excel
  • Убедитесь, что столбцы A, B, C содержат данные

Проблема: "Ошибка в строке X"
  • Проверьте синтаксис JSON в указанной строке
  • Убедитесь, что все кавычки закрыты
  • Проверьте запятые между элементами

═══════════════════════════════════════════════════════════════

Версия: 1.0
Дата обновления: 17.12.2025
Разработчик: Подпорин Н. Ю.(dgecon17@gmail.com)(n.podporin@credos.ru)
"""
    
    help_text.insert('1.0', help_content)
    help_text.config(state=DISABLED)  # Только для чтения
    
    # Кнопка закрытия
    btn_frame = Frame(help_window)
    btn_frame.pack(pady=10)
    Button(btn_frame, text="Закрыть", command=help_window.destroy, width=20).pack()
    
    # Применяем тему после создания всех виджетов
    apply_theme(help_window, dark_theme_enabled)

def go_back_to_main(current_window):
    """Возвращает к главному окну"""
    global ask_window
    if current_window:
        current_window.destroy()
    create_ask_window()

# === НОВАЯ ФУНКЦИЯ: JSON РЕДАКТОР С ПОДСВЕТКОЙ ОШИБОК ===
def validate_json(editor):
    global dark_theme_enabled
    editor.tag_remove('error', '1.0', END)
    content = editor.get('1.0', END).strip()
    if not content:
        status_fg = "#cccccc" if dark_theme_enabled else "gray"
        status_label.config(text="Файл пуст", fg=status_fg)
        return

    try:
        json.loads(content)
        status_label.config(text="✅ Корректный JSON", fg="green")
    except json.JSONDecodeError as e:
        error_line = e.lineno
        start = f"{error_line}.0"
        end = f"{error_line}.end"
        editor.tag_add('error', start, end)
        # Обновляем тег ошибки с учетом темы
        error_bg = "#8b0000" if dark_theme_enabled else "yellow"
        error_fg = "#ffcccc" if dark_theme_enabled else "red"
        editor.tag_config('error', background=error_bg, foreground=error_fg)
        status_label.config(text=f"❌ Ошибка в строке {e.lineno}: {e.msg}", fg="red")
    except Exception as ex:
        status_label.config(text=f"⚠️ Ошибка: {ex}", fg="orange")

def load_file_into_editor(filepath, editor):
    global current_file_path, cyrillic_highlight_enabled
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
        editor.delete('1.0', END)
        editor.insert('1.0', content)
        current_file_path = filepath
        update_line_numbers()  # Обновляем номера строк
        highlight_json_syntax()  # Подсвечиваем синтаксис
        # Восстанавливаем подсветку кириллицы, если она была включена
        if cyrillic_highlight_enabled:
            apply_cyrillic_highlight()
        validate_json(editor)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось загрузить файл:\n{e}")

def select_json_for_edit():
    global text_widget
    if not text_widget:
        messagebox.showerror("Ошибка", "Редактор не инициализирован")
        return
    filepath = filedialog.askopenfilename(
        title="Выберите JSON-файл",
        filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
    )
    if filepath:
        load_file_into_editor(filepath, text_widget)

def save_json():
    global current_file_path
    if not text_widget:
        messagebox.showerror("Ошибка", "Редактор не инициализирован")
        return
        
    content = text_widget.get('1.0', END).strip()
    if not content:
        messagebox.showwarning("Предупреждение", "Файл пуст.")
        return

    # Проверка корректности JSON (опционально)
    try:
        json.loads(content)
    except json.JSONDecodeError as e:
        if not messagebox.askyesno(
            "Некорректный JSON",
            f"Обнаружена ошибка:\n{e.msg} (строка {e.lineno})\n\nСохранить файл в текущем виде?"
        ):
            return

    # Если файл уже открыт, предлагаем сохранить в том же месте или выбрать новое
    if current_file_path and os.path.exists(current_file_path):
        if messagebox.askyesno("Сохранить", f"Сохранить в текущий файл?\n{current_file_path}"):
            save_path = current_file_path
        else:
            # Определяем базовое имя
            base_name = os.path.splitext(os.path.basename(current_file_path))[0]
            timestamp = datetime.now().strftime("%Y.%m.%d_%H-%M")
            suggested_name = f"{base_name}_{timestamp}.json"
            
            save_path = filedialog.asksaveasfilename(
                initialfile=suggested_name,
                defaultextension=".json",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
            )
            if not save_path:
                return
    else:
        # Определяем базовое имя
        if current_file_path:
            base_name = os.path.splitext(os.path.basename(current_file_path))[0]
        else:
            base_name = "безымянный"

        # Форматируем дату и время
        timestamp = datetime.now().strftime("%Y.%m.%d_%H-%M")
        suggested_name = f"{base_name}_{timestamp}.json"

        save_path = filedialog.asksaveasfilename(
            initialfile=suggested_name,
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not save_path:
            return

    try:
        with open(save_path, 'w', encoding='utf-8') as f:
            f.write(content)
        messagebox.showinfo("Успех", f"Файл успешно сохранён!\n\n{save_path}")
        current_file_path = save_path
        validate_json(text_widget)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{e}")

def update_line_numbers():
    """Обновляет номера строк в редакторе"""
    global text_widget, line_numbers_widget
    if not text_widget or not line_numbers_widget:
        return
    
    try:
        # Получаем количество строк через индекс последней строки
        last_line = text_widget.index('end-1c').split('.')[0]
        line_count = int(last_line)
        
        # Формируем текст с номерами строк
        line_numbers = '\n'.join(str(i) for i in range(1, line_count + 1))
        if line_count == 0:
            line_numbers = '1'
        
        # Обновляем виджет номеров строк
        line_numbers_widget.config(state=NORMAL)
        line_numbers_widget.delete('1.0', END)
        line_numbers_widget.insert('1.0', line_numbers)
        line_numbers_widget.config(state=DISABLED)
    except Exception:
        # В случае ошибки просто игнорируем
        pass

def on_text_change(event=None):
    """Обработчик изменения текста - обновляет номера строк и подсветку"""
    update_line_numbers()
    # Вызываем подсветку с небольшой задержкой, чтобы не замедлять ввод
    try:
        global editor_win, cyrillic_highlight_enabled
        if editor_win:
            editor_win.after(10, highlight_json_syntax)
            # Если подсветка кириллицы включена, обновляем её тоже
            if cyrillic_highlight_enabled:
                editor_win.after(15, highlight_cyrillic)
    except:
        pass

def sync_scroll(*args):
    """Синхронизирует прокрутку между текстовым полем и номерами строк"""
    global text_widget, line_numbers_widget
    if text_widget and line_numbers_widget:
        line_numbers_widget.yview_moveto(args[0])

def sync_line_numbers_scroll(*args):
    """Синхронизирует прокрутку номеров строк с текстовым полем"""
    global text_widget, line_numbers_widget
    if text_widget and line_numbers_widget:
        text_widget.yview_moveto(args[0])

def toggle_word_wrap():
    """Переключает перенос строк"""
    global text_widget, word_wrap_var
    if not text_widget:
        return
    
    if word_wrap_var.get():
        text_widget.config(wrap=WORD)
    else:
        text_widget.config(wrap=NONE)
    update_line_numbers()

def apply_cyrillic_highlight():
    """Применяет подсветку кириллических символов (без переключения состояния)"""
    global text_widget, dark_theme_enabled, cyrillic_highlight_enabled
    if not text_widget or not cyrillic_highlight_enabled:
        return
    
    import re
    
    # Удаляем предыдущую подсветку кириллицы
    text_widget.tag_remove("cyrillic", "1.0", END)
    
    # Получаем весь текст
    content = text_widget.get("1.0", END)
    if not content.strip():
        return
    
    # Цвета для подсветки кириллицы
    if dark_theme_enabled:
        cyrillic_bg = "#3a3a00"  # Темно-желтый фон для темной темы
        cyrillic_fg = "#ffff00"  # Ярко-желтый текст для темной темы
    else:
        cyrillic_bg = "#ffff00"  # Ярко-желтый фон для светлой темы
        cyrillic_fg = "#000000"  # Черный текст для светлой темы
    
    # Настраиваем тег для кириллицы
    text_widget.tag_configure("cyrillic", background=cyrillic_bg, foreground=cyrillic_fg)
    
    # Паттерн для поиска кириллических символов (основной блок + расширенный)
    # Кириллица: U+0400-U+04FF (основной блок) и другие связанные блоки
    cyrillic_pattern = r'[\u0400-\u04FF\u0500-\u052F\u2DE0-\u2DFF\uA640-\uA69F]'
    
    # Находим и подсвечиваем все кириллические символы
    for match in re.finditer(cyrillic_pattern, content):
        start_pos = f"1.0 + {match.start()} chars"
        end_pos = f"1.0 + {match.end()} chars"
        text_widget.tag_add("cyrillic", start_pos, end_pos)

def highlight_cyrillic():
    """Переключает подсветку кириллических символов в тексте"""
    global text_widget, dark_theme_enabled, cyrillic_highlight_enabled
    if not text_widget:
        return
    
    # Переключаем состояние подсветки
    cyrillic_highlight_enabled = not cyrillic_highlight_enabled
    
    if cyrillic_highlight_enabled:
        apply_cyrillic_highlight()
    else:
        # Удаляем подсветку
        text_widget.tag_remove("cyrillic", "1.0", END)

def highlight_json_syntax():
    """Подсвечивает синтаксис JSON как в Notepad++"""
    global text_widget, dark_theme_enabled
    if not text_widget:
        return
    
    import re
    
    # Удаляем все предыдущие теги подсветки
    text_widget.tag_remove("json_key", "1.0", END)
    text_widget.tag_remove("json_string", "1.0", END)
    text_widget.tag_remove("json_number", "1.0", END)
    text_widget.tag_remove("json_boolean", "1.0", END)
    text_widget.tag_remove("json_null", "1.0", END)
    text_widget.tag_remove("json_bracket", "1.0", END)
    text_widget.tag_remove("json_colon", "1.0", END)
    text_widget.tag_remove("json_comma", "1.0", END)
    
    # Восстанавливаем подсветку кириллицы, если она была включена
    global cyrillic_highlight_enabled
    if cyrillic_highlight_enabled:
        apply_cyrillic_highlight()
    
    # Получаем весь текст
    content = text_widget.get("1.0", END)
    if not content.strip():
        return
    
    # Цвета для темной и светлой темы
    if dark_theme_enabled:
        key_color = "#9cdcfe"      # Светло-голубой для ключей
        string_color = "#ce9178"   # Оранжево-коричневый для строк
        number_color = "#b5cea8"   # Зеленый для чисел
        boolean_color = "#569cd6"   # Синий для булевых значений
        null_color = "#569cd6"      # Синий для null
        bracket_color = "#ffd700"  # Золотой для скобок
        colon_color = "#d4d4d4"    # Светло-серый для двоеточий
        comma_color = "#d4d4d4"    # Светло-серый для запятых
    else:
        key_color = "#0451a5"      # Темно-синий для ключей
        string_color = "#a31515"   # Темно-красный для строк
        number_color = "#098658"   # Зеленый для чисел
        boolean_color = "#0000ff"  # Синий для булевых значений
        null_color = "#0000ff"     # Синий для null
        bracket_color = "#811f3f"  # Темно-красный для скобок
        colon_color = "#000000"    # Черный для двоеточий
        comma_color = "#000000"    # Черный для запятых
    
    # Настраиваем теги с цветами
    text_widget.tag_configure("json_key", foreground=key_color)
    text_widget.tag_configure("json_string", foreground=string_color)
    text_widget.tag_configure("json_number", foreground=number_color)
    text_widget.tag_configure("json_boolean", foreground=boolean_color)
    text_widget.tag_configure("json_null", foreground=null_color)
    text_widget.tag_configure("json_bracket", foreground=bracket_color, font=("Consolas", 10, "bold"))
    text_widget.tag_configure("json_colon", foreground=colon_color)
    text_widget.tag_configure("json_comma", foreground=comma_color)
    
    # Подсветка скобок и фигурных скобок
    for match in re.finditer(r'[\[\]{}]', content):
        start_pos = f"1.0 + {match.start()} chars"
        end_pos = f"1.0 + {match.end()} chars"
        text_widget.tag_add("json_bracket", start_pos, end_pos)
    
    # Подсветка двоеточий
    for match in re.finditer(r':', content):
        start_pos = f"1.0 + {match.start()} chars"
        end_pos = f"1.0 + {match.end()} chars"
        text_widget.tag_add("json_colon", start_pos, end_pos)
    
    # Подсветка запятых
    for match in re.finditer(r',', content):
        start_pos = f"1.0 + {match.start()} chars"
        end_pos = f"1.0 + {match.end()} chars"
        text_widget.tag_add("json_comma", start_pos, end_pos)
    
    # Подсветка строк (включая ключи и значения)
    # Ищем строки в кавычках, но нужно различать ключи и значения
    string_pattern = r'"(?:[^"\\]|\\.)*"'
    for match in re.finditer(string_pattern, content):
        start_pos = f"1.0 + {match.start()} chars"
        end_pos = f"1.0 + {match.end()} chars"
        # Проверяем, является ли это ключом (есть двоеточие после, возможно с пробелами)
        match_end = match.end()
        remaining = content[match_end:match_end+10].strip() if match_end < len(content) else ""
        if remaining.startswith(':'):
            text_widget.tag_add("json_key", start_pos, end_pos)
        else:
            text_widget.tag_add("json_string", start_pos, end_pos)
    
    # Подсветка чисел (целые и с плавающей точкой)
    number_pattern = r'-?\d+\.?\d*'
    for match in re.finditer(number_pattern, content):
        start_pos = f"1.0 + {match.start()} chars"
        end_pos = f"1.0 + {match.end()} chars"
        # Проверяем, что это не часть строки
        line_start = content.rfind('\n', 0, match.start())
        if line_start == -1:
            line_start = 0
        line_content = content[line_start:match.end()]
        # Если число не внутри строки
        if line_content.count('"') % 2 == 0 or (line_content.rfind('"', 0, match.start() - line_start) == -1):
            text_widget.tag_add("json_number", start_pos, end_pos)
    
    # Подсветка булевых значений
    for match in re.finditer(r'\b(true|false)\b', content, re.IGNORECASE):
        start_pos = f"1.0 + {match.start()} chars"
        end_pos = f"1.0 + {match.end()} chars"
        # Проверяем, что это не часть строки
        line_start = content.rfind('\n', 0, match.start())
        if line_start == -1:
            line_start = 0
        line_content = content[line_start:match.end()]
        if line_content.count('"') % 2 == 0:
            text_widget.tag_add("json_boolean", start_pos, end_pos)
    
    # Подсветка null
    for match in re.finditer(r'\bnull\b', content, re.IGNORECASE):
        start_pos = f"1.0 + {match.start()} chars"
        end_pos = f"1.0 + {match.end()} chars"
        # Проверяем, что это не часть строки
        line_start = content.rfind('\n', 0, match.start())
        if line_start == -1:
            line_start = 0
        line_content = content[line_start:match.end()]
        if line_content.count('"') % 2 == 0:
            text_widget.tag_add("json_null", start_pos, end_pos)

def create_json_editor_window():
    global ask_window, text_widget, line_numbers_widget, status_label, editor_win, dark_theme_enabled, word_wrap_var, cyrillic_highlight_enabled
    
    if ask_window:
        ask_window.destroy()
        ask_window = None
    
    # Сбрасываем состояние подсветки кириллицы при открытии нового редактора
    cyrillic_highlight_enabled = False

    editor_win = Tk()
    editor_win.title("JSON Редактор с проверкой")
    place_window_near_cursor(editor_win, 700, 550, 0, 0, 200)

    # Статусная строка
    status_bg = "#404040" if dark_theme_enabled else "white"
    status_fg = "#ffffff" if dark_theme_enabled else "black"
    status_label = Label(editor_win, text="Загрузите JSON-файл", relief=SUNKEN, anchor=W, 
                        bg=status_bg, fg=status_fg)
    status_label.pack(side=BOTTOM, fill=X)

    # Кнопки и чекбокс переноса строк
    btn_frame = Frame(editor_win)
    btn_frame.pack(side=TOP, fill=X, padx=10, pady=5)

    Button(btn_frame, text="Открыть файл", command=select_json_for_edit).pack(side=LEFT, padx=5)
    Button(btn_frame, text="Проверить", command=lambda: validate_json(text_widget)).pack(side=LEFT, padx=5)
    Button(btn_frame, text="Сохранить", command=save_json).pack(side=LEFT, padx=5)
    
    # Чекбокс для переноса строк
    word_wrap_var = BooleanVar(value=False)
    Checkbutton(btn_frame, text="Перенос строк", variable=word_wrap_var, 
                command=toggle_word_wrap).pack(side=LEFT, padx=5)
    
    # Кнопка для подсветки кириллицы
    Button(btn_frame, text="Подсветить кириллицу", command=highlight_cyrillic).pack(side=LEFT, padx=5)
    
    Button(btn_frame, text="Справка", command=show_help).pack(side=RIGHT, padx=5)
    Button(btn_frame, text="Назад", command=lambda: go_back_to_main(editor_win)).pack(side=RIGHT, padx=5)

    # Фрейм для текстового поля с номерами строк
    text_frame = Frame(editor_win)
    text_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)

    # Фрейм для номеров строк и текстового поля
    editor_frame = Frame(text_frame)
    editor_frame.pack(side=LEFT, fill=BOTH, expand=True)

    text_bg = "#1e1e1e" if dark_theme_enabled else "white"
    text_fg = "#ffffff" if dark_theme_enabled else "black"
    line_num_bg = "#252526" if dark_theme_enabled else "#f0f0f0"
    line_num_fg = "#858585" if dark_theme_enabled else "#666666"
    
    # Виджет для номеров строк
    line_numbers_widget = Text(editor_frame, width=5, padx=5, pady=5, 
                              font=("Consolas", 10), bg=line_num_bg, fg=line_num_fg,
                              state=DISABLED, wrap=NONE, takefocus=0)
    line_numbers_widget.pack(side=LEFT, fill=Y)
    
    # Текстовое поле с прокруткой
    text_widget = Text(editor_frame, wrap=NONE, font=("Consolas", 10), undo=True,
                      bg=text_bg, fg=text_fg, insertbackground=text_fg,
                      padx=5, pady=5)
    
    # Вертикальная прокрутка
    scroll_y = Scrollbar(text_frame, orient=VERTICAL)
    scroll_y.pack(side=RIGHT, fill=Y)
    
    # Горизонтальная прокрутка
    scroll_x = Scrollbar(text_frame, orient=HORIZONTAL)
    scroll_x.pack(side=BOTTOM, fill=X)
    
    # Настройка прокрутки
    text_widget.config(yscrollcommand=lambda *args: [scroll_y.set(*args), sync_scroll(*args)])
    text_widget.config(xscrollcommand=scroll_x.set)
    scroll_y.config(command=lambda *args: [text_widget.yview(*args), sync_line_numbers_scroll(*args)])
    scroll_x.config(command=text_widget.xview)
    
    # Синхронизация прокрутки номеров строк
    line_numbers_widget.config(yscrollcommand=lambda *args: [scroll_y.set(*args), sync_line_numbers_scroll(*args)])
    
    text_widget.pack(side=LEFT, fill=BOTH, expand=True)

    # Привязка событий для обновления номеров строк
    text_widget.bind('<KeyRelease>', on_text_change)
    text_widget.bind('<Button-1>', on_text_change)
    text_widget.bind('<Return>', on_text_change)
    text_widget.bind('<BackSpace>', on_text_change)
    text_widget.bind('<Delete>', on_text_change)
    text_widget.bind('<Button-4>', on_text_change)
    text_widget.bind('<Button-5>', on_text_change)
    
    # Инициализация номеров строк
    update_line_numbers()

    # Тег для ошибок (адаптируем под тему)
    error_bg = "#8b0000" if dark_theme_enabled else "yellow"
    error_fg = "#ffcccc" if dark_theme_enabled else "red"
    text_widget.tag_configure('error', background=error_bg, foreground=error_fg)
    
    # Применяем тему после создания всех виджетов
    apply_theme(editor_win, dark_theme_enabled)
    
    # Инициализируем подсветку синтаксиса
    highlight_json_syntax()
    
    # Горячие клавиши
    editor_win.bind('<Control-o>', lambda e: select_json_for_edit())
    editor_win.bind('<Control-s>', lambda e: save_json())
    editor_win.bind('<F5>', lambda e: validate_json(text_widget))
    
    editor_win.mainloop()

# === ОСНОВНОЕ МЕНЮ ===
def create_ask_window():
    global ask_window, editor_win, converter_win, dark_theme_enabled
    
    # Закрываем другие окна если они открыты
    if editor_win:
        try:
            editor_win.destroy()
        except:
            pass
        editor_win = None
    
    if converter_win:
        try:
            converter_win.destroy()
        except:
            pass
        converter_win = None
    
    # Определяем тему Windows при первом запуске
    if dark_theme_enabled is False and not hasattr(create_ask_window, 'theme_checked'):
        dark_theme_enabled = get_windows_theme()
        create_ask_window.theme_checked = True
    
    ask_window = Tk()
    ask_window.title("XLSX/JSON Helper")
    ask_window.resizable(False, False)
    place_window_near_cursor(ask_window, 350, 200, 0, 0, 250)

    Label(ask_window, text="Что вы хотите сделать?", 
          font=("Segoe UI", 12, "bold")).pack(pady=15)
    
    btn_frame = Frame(ask_window)
    btn_frame.pack(pady=10)
    
    Button(btn_frame, text="Конвертировать *.xlsx/*.xlsm в *.json", 
           command=start_xls2json_win, width=40).pack(anchor=CENTER, pady=5)
    Button(btn_frame, text="Исправление синтаксиса *.json", 
           command=create_json_editor_window, width=40).pack(anchor=CENTER, pady=5)
    Button(btn_frame, text="Справка", 
           command=show_help, width=40).pack(anchor=CENTER, pady=5)
    
    # Чекбокс для переключения темы
    theme_frame = Frame(ask_window)
    theme_frame.pack(pady=5)
    
    theme_var = BooleanVar(value=dark_theme_enabled)
    
    def toggle_theme():
        global dark_theme_enabled
        dark_theme_enabled = theme_var.get()
        apply_theme(ask_window, dark_theme_enabled)
    
    theme_checkbox = Checkbutton(theme_frame, text="Темная тема", 
                                 command=toggle_theme, 
                                 variable=theme_var)
    theme_checkbox.pack()
    
    # Применяем тему после создания всех виджетов
    apply_theme(ask_window, dark_theme_enabled)
    
    # Горячие клавиши для быстрого доступа
    ask_window.bind('<Control-1>', lambda e: start_xls2json_win())
    ask_window.bind('<Control-2>', lambda e: create_json_editor_window())
    ask_window.bind('<Escape>', lambda e: ask_window.destroy())

    ask_window.mainloop()

if __name__ == "__main__":
    create_ask_window()