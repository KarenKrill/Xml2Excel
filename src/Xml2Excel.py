import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Any

import pandas as pd
from lxml import etree

# -------------------------------
# Глобальные переменные
# -------------------------------
df: None = None
checkbox_vars = {}
check_all_var = None
namespaces = {}
current_xml_path = None

# -------------------------------
# XML utils
# -------------------------------
def extract_namespaces(xml_path):
    ns_map = {}
    for _, elem in etree.iterparse(xml_path, events=("start-ns",)):
        prefix, uri = elem
        prefix = prefix if prefix else "default"
        ns_map[prefix] = uri
    return ns_map


def load_xml_with_namespace(xml_path, ns_uri):
    ns = {"ns": ns_uri} if ns_uri else None
    return pd.read_xml(xml_path, namespaces=ns)


# -------------------------------
# GUI callbacks
# -------------------------------
def select_xml():
    global df, namespaces, current_xml_path

    path = filedialog.askopenfilename(filetypes=[("XML files", "*.xml")])
    if not path:
        return

    try:
        namespaces = extract_namespaces(path)
        if not namespaces:
            namespaces = {"(no namespace)": None}
        
        display_values = [
            f'{k}: "{v}"' if v else k
            for k, v in namespaces.items()
        ]

        namespace_menu["values"] = display_values
        namespace_menu.current(0)

        current_xml_path = path
        load_dataframe()

    except Exception as e:
        messagebox.showerror("Ошибка", str(e))


def load_dataframe():
    global df

    try:
        selected = namespace_menu.get()

        if ':' in selected:
            prefix = selected.split(':', 1)[0]
        else:
            prefix = selected

        ns_uri = namespaces.get(prefix)

        df = load_xml_with_namespace(current_xml_path, ns_uri)
        show_fields(df.columns)

    except Exception as e:
        messagebox.showerror("Ошибка загрузки XML", str(e))


def show_fields(columns):
    for widget in fields_inner.winfo_children():
        widget.destroy()

    checkbox_vars.clear()

    for col in columns:
        var = tk.BooleanVar()
        chk = tk.Checkbutton(fields_inner, text=col, variable=var, command=update_check_all_var)
        chk.pack(anchor="w")
        checkbox_vars[col] = var

    update_check_all_var()
        
def select_all():
    for var in checkbox_vars.values():
        var.set(True)
def deselect_all():
    for var in checkbox_vars.values():
        var.set(False)
def on_check_all_var_updated():
    if check_all_var.get():
        select_all()
    else:
        deselect_all()
def update_check_all_var():
    if all(var.get() for var in checkbox_vars.values()):
        check_all_var.set(True)
    else:
        check_all_var.set(False)

def export_excel():
    if df is None:
            messagebox.showwarning("Внимание", "Сначала выберите XML-файл")
            return
            
    selected: list[Any] = [c for c, v in checkbox_vars.items() if v.get()]
    if not selected:
        messagebox.showwarning("Внимание", "Выберите хотя бы одно поле")
        return

    path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if not path:
        return
    
    try:
        df[selected].to_excel(path, index=False)
    
    except PermissionError:
        messagebox.showerror(
            "Нет доступа",
            "Невозможно сохранить файл.\n\n"
            "Возможные причины:\n"
            "• файл уже открыт в Excel\n"
            "• нет прав на запись в эту папку"
        )

    except FileNotFoundError:
        messagebox.showerror(
            "Ошибка пути",
            "Указанный путь не существует."
        )

    except Exception as e:
        messagebox.showerror(
            "Неизвестная ошибка",
            f"Произошла ошибка при экспорте:\n\n{e}"
        )

    else:
        messagebox.showinfo("Готово", "Экспорт успешно завершён")
    
def center_window_auto(window, width, height):
    window.update_idletasks()

    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)

    window.geometry(f"+{x}+{y}")

try:
    # -------------------------------
    # GUI setup
    # -------------------------------
    root = tk.Tk()
    root.title("XML → Excel")
    center_window_auto(root, 500, 600)

    # XML selection
    tk.Button(root, text="Выбрать XML", command=select_xml).pack(pady=5)

    # Namespace selector
    tk.Label(root, text="Namespace:").pack()
    namespace_menu = tk.ttk.Combobox(root, state="readonly")
    namespace_menu.pack(pady=5)
    namespace_menu.bind("<<ComboboxSelected>>", lambda e: load_dataframe())
    
    # Frame для кнопок "Выбрать всё / Снять всё"
    check_all_var = tk.BooleanVar()
    control_frame = tk.Frame(root)
    control_frame.pack(pady=10)

    chkAllButton = tk.Checkbutton(control_frame, text="Выбрать/убрать все", variable=check_all_var, command=on_check_all_var_updated)
    chkAllButton.pack(anchor="w",side="left", padx=5)
    
    uncheck_all_var = tk.BooleanVar()
    
    # Scrollable fields area
    fields_frame = tk.LabelFrame(root, text="Выберите поля для экспорта")
    fields_frame.pack(fill="both", expand=True, padx=10, pady=10)

    canvas = tk.Canvas(fields_frame)
    scrollbar = tk.Scrollbar(fields_frame, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)

    scrollbar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    fields_inner = tk.Frame(canvas, padx=15, pady=15)
    canvas.create_window((0, 0), window=fields_inner, anchor="nw")

    fields_inner.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    # Export button
    tk.Button(root, text="Экспорт в Excel", command=export_excel).pack(pady=10)

    root.mainloop()
    
except Exception as e:
    messagebox.showerror(
        "Неизвестная ошибка",
        f"Произошла ошибка при работе программы:\n\n{e}"
    )