import os
import pandas as pd
import pygsheets
from openpyxl import load_workbook
from tkinter import filedialog, Tk, Label, Button, OptionMenu, StringVar, ttk
from win32com.client import Dispatch as client
from tkinter import messagebox
import threading

# Функция для выбора Google Sheets таблицы
def select_google_sheet(gc, root, var, var_tab):
    available_sheets = gc.spreadsheet_titles()
    var.set(available_sheets[0])
    OptionMenu(root, var, *available_sheets, command=lambda _: select_google_sheet_tab(gc, var_tab)).grid(row=2, column=1, pady=5)

# Функция для выбора листа в Google Sheets таблице
def select_google_sheet_tab(gc, var_tab):
    sh = gc.open(selected_sheet.get())
    available_tabs = [sheet.title for sheet in sh.worksheets()]
    var_tab.set(available_tabs[0])
    tab_menu['menu'].delete(0, 'end')
    for tab in available_tabs:
        tab_menu['menu'].add_command(label=tab, command=lambda v=var_tab, value=tab: v.set(value))

# Функция для выбора файла Excel
def select_excel_file(var):
    filename = filedialog.askopenfilename()
    var.set(filename)
    file_label.config(text=f"Выбранный файл: {filename}")

# Функция для конвертации xls файла в xlsx
def convert_xls_to_xlsx(xls_file):
    try:
        file_path, file_name = os.path.split(xls_file)
        file_name_without_extension = os.path.splitext(file_name)[0]
        xlsx_file = os.path.join(file_path, f"{file_name_without_extension}.xlsx")
        xlsx_file = os.path.normpath(xlsx_file)
        
        if os.path.exists(xlsx_file):
            os.remove(xlsx_file)

        excel = client("Excel.Application")
        wb = excel.Workbooks.Open(xls_file)
        wb.SaveAs(xlsx_file, FileFormat=51)
        wb.Close()
        excel.Quit()
        
        return xlsx_file
    except Exception as e:
        print(f"Ошибка при конвертации файла: {e}")
        return None

# Функция для загрузки данных в Google Sheets
def upload_to_google_sheets(gc, selected_sheet, selected_tab, excel_file):
    if excel_file.get().endswith('.xls'):
        converted_file = convert_xls_to_xlsx(excel_file.get())
        if converted_file:
            excel_file.set(converted_file)
        else:
            messagebox.showerror("Ошибка", "Ошибка при конвертации файла. Загрузка отменена.")
            return False
    
    sh = gc.open(selected_sheet.get())
    worksheet = sh.worksheet_by_title(selected_tab.get())

    if os.path.exists(excel_file.get()):
        wb = load_workbook(excel_file.get())
        ws = wb.active
        df = pd.DataFrame(ws.values)
        max_rows, max_cols = min(df.shape[0], 1086), min(df.shape[1], 56)
        df_selected = df.iloc[:max_rows, :max_cols]
        worksheet.clear()
        worksheet.update_values(crange='A1', values=df_selected.values.tolist())
        messagebox.showinfo("Успех", "Данные успешно загружены в Google Sheets.")
        
        if excel_file.get().endswith('.xlsx'):
            os.remove(excel_file.get())
            excel_file.set("")
        return True
    else:
        messagebox.showerror("Ошибка", "Файл Excel не найден.")
        return False

# Функция для запуска загрузки данных в отдельном потоке
def start_upload_thread():
    if excel_file.get():
        thread = threading.Thread(target=lambda: upload_to_google_sheets(gc, selected_sheet, selected_tab, excel_file))
        thread.daemon = True
        thread.start()
    else:
        messagebox.showerror("Ошибка", "Файл Excel не выбран.")

# Путь к файлу с учетными данными JSON
credentials_file = 'credentials.json'

# Аутентификация и создание клиента для доступа к API Google Sheets
gc = pygsheets.authorize(service_account_file=credentials_file)

# Создаем окно
root = Tk()
root.title("Загрузка данных в Google Sheets")
root.geometry("500x300")

# Определяем стиль для кнопок
style = ttk.Style()
style.configure('TButton', font=('Helvetica', 12))

# Переменные для хранения выбранной таблицы, листа и файла Excel
selected_sheet = StringVar()
selected_tab = StringVar()
excel_file = StringVar()

# Создаем фрейм для заголовка
header_frame = ttk.Frame(root)
header_frame.grid(row=0, column=0, columnspan=2, pady=10)

# Метка для заголовка
header_label = ttk.Label(header_frame, text="Загрузка данных в Google Sheets", font=('Helvetica', 18, 'bold'))
header_label.pack()

# Создаем фрейм для выбора файла Excel
file_frame = ttk.Frame(root)
file_frame.grid(row=1, column=0, padx=10, pady=5)

# Метка для отображения выбранного файла
file_label = ttk.Label(file_frame, text="Выбранный файл: Не выбран", wraplength=300)
file_label.grid(row=0, column=0, padx=10, pady=5)

# Кнопка для выбора файла Excel
file_button = ttk.Button(file_frame, text="Выбрать файл", command=lambda: select_excel_file(excel_file))
file_button.grid(row=0, column=1, padx=10, pady=5)

# Создаем фрейм для выбора таблицы
sheet_frame = ttk.Frame(root)
sheet_frame.grid(row=2, column=0, padx=10, pady=5)

# Метка для выбора таблицы
sheet_label = ttk.Label(sheet_frame, text="Выберите Google Sheets таблицу:")
sheet_label.grid(row=0, column=0, padx=10, pady=5)

# Выпадающий список для выбора таблицы
select_google_sheet(gc, sheet_frame, selected_sheet, selected_tab)

# Создаем фрейм для выбора листа
tab_frame = ttk.Frame(root)
tab_frame.grid(row=3, column=0, padx=10, pady=5)

# Метка для выбора листа
tab_label = ttk.Label(tab_frame, text="Выберите лист в таблице:")
tab_label.grid(row=0, column=0, padx=10, pady=5)

# Выпадающий список для выбора листа
tab_menu = OptionMenu(tab_frame, selected_tab, "")
tab_menu.grid(row=0, column=1, padx=10, pady=5)

# Кнопка для загрузки данных в Google Sheets
upload_button = ttk.Button(root, text="Загрузить данные", command=start_upload_thread).grid(row=4, column=0, columnspan=2, padx=10, pady=10)

root.mainloop()
