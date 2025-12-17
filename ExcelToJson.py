from tkinter import *
from tkinter import ttk
from tkinter import filedialog, messagebox
import json
import os
from datetime import datetime
import pandas as pd
import subprocess
import platform

# Глобальные переменные
ask_window = None
current_file_path = None
text_widget = None
status_label = None
editor_win = None
converter_win = None


def place_window_near_cursor(window, width, height, dx=0, dy=0, screen_margin=20):
    x, y = window.winfo_pointerxy()
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    win_x = x + dx
    win_y = y + dy

    if win_x + width > screen_width - screen_margin:
        win_x = screen_width - width - screen_margin
    if win_y + height > screen_height - screen_margin:
        win_y = screen_height - height - screen_margin
    if win_x < screen_margin:
        win_x = screen_margin
    if win_y < screen_margin:
        win_y = screen_margin

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
    global ask_window, converter_win, status_label
    
    if ask_window:
        ask_window.destroy()
        ask_window = None
    
    converter_win = Tk()
    converter_win.title("Конвертация Excel → JSON")
    converter_win.configure(bg="#f9f9f9")
    converter_win.resizable(False, False)
    place_window_near_cursor(converter_win, 450, 220, 0, 0, 200)
    
    Label(converter_win, text="Конвертация Excel в JSON", 
          font=("Segoe UI", 12, "bold"), bg="#f9f9f9").pack(pady=15)
    
    # Поле статуса (Jobizdan)
    status_label = Label(converter_win, text="Готов к конвертации", 
                        relief=SUNKEN, anchor=W, bg="white", fg="gray", 
                        font=("Segoe UI", 9))
    status_label.pack(fill=X, padx=10, pady=5)
    
    btn_frame = Frame(converter_win, bg="#f9f9f9")
    btn_frame.pack(pady=10)
    
    ttk.Button(btn_frame, text="Выбрать файл и конвертировать", 
               command=select_excel_file, width=35).pack(pady=5)
    
    btn_frame2 = Frame(converter_win, bg="#f9f9f9")
    btn_frame2.pack(pady=5)
    
    ttk.Button(btn_frame2, text="Справка", 
               command=show_help, width=15).pack(side=LEFT, padx=5)
    ttk.Button(btn_frame2, text="Назад", 
               command=lambda: go_back_to_main(converter_win), width=15).pack(side=LEFT, padx=5)
    
    # Горячая клавиша для выбора файла
    converter_win.bind('<Control-o>', lambda e: select_excel_file())
    converter_win.bind('<Return>', lambda e: select_excel_file())
    
    converter_win.mainloop()

def show_help():
    """Открывает окно со справкой"""
    help_window = Toplevel()
    help_window.title("Справка")
    help_window.configure(bg="#f9f9f9")
    help_window.resizable(True, True)
    help_window.geometry("700x600")
    
    # Создаем текстовое поле с прокруткой
    text_frame = Frame(help_window)
    text_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)
    
    scrollbar = Scrollbar(text_frame)
    scrollbar.pack(side=RIGHT, fill=Y)
    
    help_text = Text(text_frame, wrap=WORD, font=("Segoe UI", 10), 
                     yscrollcommand=scrollbar.set, bg="white", padx=10, pady=10)
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

Проблема: "Не удалось скопировать в буфер обмена"
  • Установите pyperclip: pip install pyperclip
  • Или скопируйте JSON вручную из сохраненного файла

═══════════════════════════════════════════════════════════════

Версия: 1.0
Дата обновления: 17.12.2025
Разработчик: Подпорин Н. Ю.(dgecon17@gmail.com)(n.podporin@credos.ru)
"""
    
    help_text.insert('1.0', help_content)
    help_text.config(state=DISABLED)  # Только для чтения
    
    # Кнопка закрытия
    btn_frame = Frame(help_window, bg="#f9f9f9")
    btn_frame.pack(pady=10)
    ttk.Button(btn_frame, text="Закрыть", command=help_window.destroy, width=20).pack()

def go_back_to_main(current_window):
    """Возвращает к главному окну"""
    global ask_window
    if current_window:
        current_window.destroy()
    create_ask_window()

# === НОВАЯ ФУНКЦИЯ: JSON РЕДАКТОР С ПОДСВЕТКОЙ ОШИБОК ===
def validate_json(editor):
    editor.tag_remove('error', '1.0', END)
    content = editor.get('1.0', END).strip()
    if not content:
        status_label.config(text="Файл пуст", fg="gray")
        return

    try:
        json.loads(content)
        status_label.config(text="✅ Корректный JSON", fg="green")
    except json.JSONDecodeError as e:
        error_line = e.lineno
        start = f"{error_line}.0"
        end = f"{error_line}.end"
        editor.tag_add('error', start, end)
        editor.tag_config('error', background="yellow", foreground="red")
        status_label.config(text=f"❌ Ошибка в строке {e.lineno}: {e.msg}", fg="red")
    except Exception as ex:
        status_label.config(text=f"⚠️ Ошибка: {ex}", fg="orange")

def load_file_into_editor(filepath, editor):
    global current_file_path
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
        editor.delete('1.0', END)
        editor.insert('1.0', content)
        current_file_path = filepath
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

def create_json_editor_window():
    global ask_window, text_widget, status_label, editor_win
    
    if ask_window:
        ask_window.destroy()
        ask_window = None

    editor_win = Tk()
    editor_win.title("JSON Редактор с проверкой")
    editor_win.configure(bg="#f9f9f9")
    place_window_near_cursor(editor_win, 600, 500, 0, 0, 200)

    # Статусная строка
    status_label = Label(editor_win, text="Загрузите JSON-файл", relief=SUNKEN, anchor=W, bg="white")
    status_label.pack(side=BOTTOM, fill=X)

    # Кнопки
    btn_frame = Frame(editor_win, bg="#f9f9f9")
    btn_frame.pack(side=TOP, fill=X, padx=10, pady=5)

    ttk.Button(btn_frame, text="Открыть файл", command=select_json_for_edit).pack(side=LEFT, padx=5)
    ttk.Button(btn_frame, text="Проверить", command=lambda: validate_json(text_widget)).pack(side=LEFT, padx=5)
    ttk.Button(btn_frame, text="Сохранить", command=save_json).pack(side=LEFT, padx=5)
    ttk.Button(btn_frame, text="Справка", command=show_help).pack(side=RIGHT, padx=5)
    ttk.Button(btn_frame, text="Назад", command=lambda: go_back_to_main(editor_win)).pack(side=RIGHT, padx=5)

    # Текстовое поле с прокруткой
    text_frame = Frame(editor_win)
    text_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)

    text_widget = Text(text_frame, wrap=NONE, font=("Consolas", 10), undo=True)
    scroll_y = Scrollbar(text_frame, orient=VERTICAL, command=text_widget.yview)
    scroll_x = Scrollbar(text_frame, orient=HORIZONTAL, command=text_widget.xview)
    text_widget.config(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

    scroll_y.pack(side=RIGHT, fill=Y)
    scroll_x.pack(side=BOTTOM, fill=X)
    text_widget.pack(side=LEFT, fill=BOTH, expand=True)

    # Тег для ошибок
    text_widget.tag_configure('error', background="yellow", foreground="red")
    
    # Горячие клавиши
    editor_win.bind('<Control-o>', lambda e: select_json_for_edit())
    editor_win.bind('<Control-s>', lambda e: save_json())
    editor_win.bind('<F5>', lambda e: validate_json(text_widget))
    
    editor_win.mainloop()

# === ОСНОВНОЕ МЕНЮ ===
def create_ask_window():
    global ask_window, editor_win, converter_win
    
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
    
    ask_window = Tk()
    ask_window.title("XLSX/JSON Helper")
    ask_window.configure(bg="#f9f9f9")
    ask_window.resizable(False, False)
    place_window_near_cursor(ask_window, 350, 180, 0, 0, 250)

    Label(ask_window, text="Что вы хотите сделать?", 
          font=("Segoe UI", 12, "bold"), bg="#f9f9f9").pack(pady=15)
    
    btn_frame = Frame(ask_window, bg="#f9f9f9")
    btn_frame.pack(pady=10)
    
    ttk.Button(btn_frame, text="Конвертировать *.xlsx/*.xlsm в *.json", 
               command=start_xls2json_win, width=40).pack(anchor=CENTER, pady=5)
    ttk.Button(btn_frame, text="Исправление синтаксиса *.json", 
               command=create_json_editor_window, width=40).pack(anchor=CENTER, pady=5)
    ttk.Button(btn_frame, text="Справка", 
               command=show_help, width=40).pack(anchor=CENTER, pady=5)
    
    # Горячие клавиши для быстрого доступа
    ask_window.bind('<Control-1>', lambda e: start_xls2json_win())
    ask_window.bind('<Control-2>', lambda e: create_json_editor_window())
    ask_window.bind('<Escape>', lambda e: ask_window.destroy())

    ask_window.mainloop()

if __name__ == "__main__":
    create_ask_window()