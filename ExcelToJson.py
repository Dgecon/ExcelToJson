from tkinter import *
from tkinter import ttk
from tkinter import filedialog, messagebox
import json
import os
from datetime import datetime

# Глобальные переменные
ask_window = None
current_file_path = None
text_widget = None
status_label = None


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

def start_xls2json_win():
    messagebox.showinfo("В разработке", "Конвертация XLSX → JSON пока не реализована.")

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
    filepath = filedialog.askopenfilename(
        title="Выберите JSON-файл"
    )
    if filepath:
        load_file_into_editor(filepath, text_widget)

def save_json():
    global current_file_path
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

    # Определяем базовое имя
    if current_file_path:
        # Берём имя без расширения (например, "data.xlsx" → "data")
        base_name = os.path.splitext(os.path.basename(current_file_path))[0]
    else:
        base_name = "безымянный"

    # Форматируем дату и время
    timestamp = datetime.now().strftime("%Y.%m.%d_%H-%M")
    suggested_name = f"{base_name}_{timestamp}.json"

    # Пользователь выбирает папку
    folder = filedialog.askdirectory()
    if folder:
        save_path = filedialog.asksaveasfilename(
            initialdir=folder,
            initialfile=suggested_name,
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")]
        )

    save_path = os.path.join(folder, suggested_name)

    try:
        with open(save_path, 'w', encoding='utf-8') as f:
            f.write(content)
        messagebox.showinfo("Успех", f"Файл успешно сохранён!\n\n{save_path}")
        current_file_path = save_path
        validate_json(text_widget)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{e}")

def create_json_editor_window():
    global ask_window, text_widget, status_label
    ask_window.destroy()

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

# === ОСНОВНОЕ МЕНЮ ===
def create_ask_window():
    global ask_window
    ask_window = Tk()
    ask_window.title("XLSX/JSON Helper")
    ask_window.configure(bg="#f9f9f9")
    ask_window.resizable(False, False)
    place_window_near_cursor(ask_window, 300, 150, 0, 0, 250)

    Label(ask_window, text="Что вы хотите сделать?", font=("Segoe UI", 12, "bold"), bg="#f9f9f9").pack(pady=10)
    ttk.Button(ask_window, text="Конвертировать *.xlsx/*.xlsm в *.json", command=start_xls2json_win, width=40).pack(anchor=CENTER, pady=5)
    ttk.Button(ask_window, text="Исправление синтаксиса *.json", command=create_json_editor_window, width=30).pack(anchor=CENTER, pady=5)

    ask_window.mainloop()

if __name__ == "__main__":
    create_ask_window()