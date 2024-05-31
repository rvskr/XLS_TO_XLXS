import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import psutil
import win32com.client as client
import subprocess
import threading

xlsx_file = None
converted_folder = None

def is_excel_running():
    return any("excel.exe" in proc.name().lower() for proc in psutil.process_iter())

def kill_excel_process():
    for proc in psutil.process_iter():
        if "excel.exe" in proc.name().lower():
            proc.kill()

def convert_xls_to_xlsx(xls_file):
    try:
        file_path, file_name = os.path.split(xls_file)
        file_name_without_extension = os.path.splitext(file_name)[0]
        
        xlsx_file = os.path.join(file_path, f"{file_name_without_extension}.xlsx")
        xlsx_file = os.path.normpath(xlsx_file)
        
        if os.path.exists(xlsx_file):
            os.remove(xlsx_file)

        excel = client.Dispatch("Excel.Application")
        wb = excel.Workbooks.Open(xls_file)
        
        wb.SaveAs(xlsx_file, FileFormat=51)
        wb.Close()
        excel.Quit()
        
        return xlsx_file
    except Exception as e:
        print(f"Ошибка при конвертации файла: {e}")
        return None

def open_folder(folder_path):
    try:
        subprocess.Popen(f'explorer "{folder_path}"')
    except Exception as e:
        print(f"Ошибка при открытии папки: {e}")

def convert_single_file(file_path, status_label, open_folder_button):
    global converted_folder
    xlsx_file = convert_xls_to_xlsx(file_path)
    if xlsx_file:
        print(f"Файл успешно сконвертирован и сохранен как: {xlsx_file}")
        converted_folder = os.path.dirname(xlsx_file)
        open_folder_button.config(state="normal")
        status_label.config(text=f"Файл успешно сконвертирован и сохранен как:\n{xlsx_file}")
    else:
        print("Ошибка конвертации файла.")
        messagebox.showerror("Ошибка", "Ошибка конвертации файла.")
        status_label.config(text="Ошибка конвертации файла.")

def convert_folder(folder_path, status_label, open_folder_button):
    global converted_folder
    converted = False
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".xls"):
            file_path = os.path.join(folder_path, file_name)
            threading.Thread(target=convert_single_file, args=(file_path, status_label, open_folder_button)).start()
            converted = True
    if not converted:
        print("Нет файлов для конвертации в выбранной папке.")
        messagebox.showinfo("Информация", "Нет файлов для конвертации в выбранной папке.")

def open_gui():
    root = tk.Tk()
    root.title("Конвертер XLS в XLSX")

    style = ttk.Style(root)
    style.theme_use("clam")  # Используем тему оформления "clam"

    # Настройка стиля для более темной темы
    root.tk_setPalette(background='#2c3e50', foreground='white', activeBackground='#34495e', activeForeground='white')

    window_width = 400
    window_height = 240

    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    x = (screen_width // 2) - (window_width // 2)
    y = (screen_height // 2) - (window_height // 2)

    root.geometry(f"{window_width}x{window_height}+{x}+{y}")

    main_frame = ttk.Frame(root, padding="20")
    main_frame.place(relx=0.5, rely=0.5, anchor="center")  # Размещаем фрейм по центру окна

    def convert_and_close():
        global xlsx_file
        xls_file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xls")])

        if not xls_file_path:
            status_label.config(text="Отменено пользователем.")
            return

        if is_excel_running():
            response = messagebox.askyesno("Предупреждение", "Excel запущен. Хотите завершить процесс Excel перед конвертацией?")
            if response:
                kill_excel_process()
            else:
                return

        threading.Thread(target=convert_single_file, args=(xls_file_path, status_label, open_folder_button)).start()

    def open_folder_click():
        global converted_folder
        if converted_folder:
            open_folder(converted_folder)
        else:
            messagebox.showinfo("Предупреждение", "Сначала сконвертируйте файлы в папке.")

    def select_folder_and_convert():
        folder_path = filedialog.askdirectory()
        if folder_path:
            threading.Thread(target=convert_folder, args=(folder_path, status_label, open_folder_button)).start()

    convert_button = ttk.Button(main_frame, text="Выбрать файл и сконвертировать", command=convert_and_close)
    convert_button.grid(row=0, column=0, pady=5, sticky="ew")

    select_folder_button = ttk.Button(main_frame, text="Выбрать папку и конвертировать все файлы", command=select_folder_and_convert)
    select_folder_button.grid(row=1, column=0, pady=5, sticky="ew")

    open_folder_button = ttk.Button(main_frame, text="Открыть папку с конвертированными файлами", command=open_folder_click, state="disabled")
    open_folder_button.grid(row=2, column=0, pady=5, sticky="ew")

    status_label = ttk.Label(main_frame, text="")
    status_label.grid(row=3, column=0, pady=5, sticky="ew")

    root.mainloop()

if __name__ == "__main__":
    open_gui()

       
