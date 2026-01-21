import tkinter as tk
from tkinter import messagebox, filedialog
from docx import Document
from docx2pdf import convert
import os

def replace(doc, data):
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            for key, value in data.items():
                if key in run.text:
                    run.text = run.text.replace(key, value)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        for key, value in data.items():
                            if key in run.text:
                                run.text = run.text.replace(key, value)

def load_from_txt():
    file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
    if file_path:
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                loaded_data = {}
                for line in f:
                    line = line.strip()
                    if not line or line.startswith('#'):
                        continue
                    if '=' in line:
                        key, value = line.split('=', 1)
                        loaded_data[key.strip()] = value.strip()

            entry_city.delete(0, tk.END); entry_city.insert(0, loaded_data.get("City", ""))
            entry_date.delete(0, tk.END); entry_date.insert(0, loaded_data.get("ContractDate", ""))
            entry_owner.delete(0, tk.END); entry_owner.insert(0, loaded_data.get("LandlordName", ""))
            entry_position.delete(0, tk.END); entry_position.insert(0, loaded_data.get("Position", ""))
            entry_basis.delete(0, tk.END); entry_basis.insert(0, loaded_data.get("Basis", ""))
            entry_tenant.delete(0, tk.END); entry_tenant.insert(0, loaded_data.get("TenantName", ""))
            entry_series.delete(0, tk.END); entry_series.insert(0, loaded_data.get("PassportSeries", ""))
            entry_num.delete(0, tk.END); entry_num.insert(0, loaded_data.get("PassportNumber", ""))
            entry_passport_issued.delete(0, tk.END); entry_passport_issued.insert(0, loaded_data.get("PassportIssued", ""))
            entry_code.delete(0, tk.END); entry_code.insert(0, loaded_data.get("DepartmentCode", ""))
            entry_address.delete(0, tk.END); entry_address.insert(0, loaded_data.get("PropertyAddress", ""))
            entry_year_built.delete(0, tk.END); entry_year_built.insert(0, loaded_data.get("YearBuilt", ""))
            entry_rooms.delete(0, tk.END); entry_rooms.insert(0, loaded_data.get("RoomsCount", ""))
            entry_area.delete(0, tk.END); entry_area.insert(0, loaded_data.get("TotalArea", ""))
            entry_living_area.delete(0, tk.END); entry_living_area.insert(0, loaded_data.get("LivingArea", ""))
            entry_room.delete(0, tk.END); entry_room.insert(0, loaded_data.get("PropertyType", ""))
            entry_keys.delete(0, tk.END); entry_keys.insert(0, loaded_data.get("KeysCount", ""))
            entry_price.delete(0, tk.END); entry_price.insert(0, loaded_data.get("MonthlyPayment", ""))

            messagebox.showinfo("Успех", "Данные загружены из TXT!")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить TXT: {e}")

def generate():
    data = {
        "[вписать нужное]": entry_city.get(),
        "[число, месяц, год]": entry_date.get(),
        "[Указать наименование собственника жилого помещения или управомоченного лица]": entry_owner.get(),
        "[должность, Ф. И. О. полностью]": entry_position.get(),
        "[устава, положения, доверенности]": entry_basis.get(),
        "[Ф. И. О. полностью]": entry_tenant.get(),
        "[вписать нужное]": entry_series.get(),
        "[значение]": entry_num.get(),
        "[наименование органа, выдавшего паспорт, дата выдачи]": entry_passport_issued.get(),
        "[вписать нужное]": entry_code.get(),
        "[вписать нужное]": entry_address.get(),
        "[вписать нужное]": entry_address.get(),
        "[указать год постройки]": entry_year_built.get(),
        "[значение]": entry_rooms.get(),
        "[цифрами и прописью]": entry_area.get(),
        "[цифрами и прописью]": entry_living_area.get(),
        "[квартира/жилой дом/часть квартиры/часть жилого дома]": entry_room.get(),
        "[значение]": entry_keys.get(),
        "[цифрами и прописью]": entry_price.get(),
    }

    required_fields = [entry_city, entry_date, entry_owner, entry_tenant, entry_address, entry_room, entry_price]
    for field in required_fields:
        if not field.get().strip():
            messagebox.showwarning("Ошибка", "Заполните все обязательные поля!")
            return

    if not os.path.exists("template.docx"):
        messagebox.showerror("Ошибка", "Файл template.docx не найден!\nПоместите его в папку со скриптом.")
        return

    try:
        doc = Document("template.docx")
        replace(doc, data)
        save_path = filedialog.asksaveasfilename(
            defaultextension=".docx", 
            filetypes=[("Word Document", "*.docx"), ("All files", "*.*")],
            title="Сохранить договор как",
            initialfile="Договор_аренды.docx"
        )
        if save_path:
            doc.save(save_path)
            try:
                pdf_path = save_path.replace(".docx", ".pdf")
                convert(save_path, pdf_path)
                messagebox.showinfo("Готово", f"Договор успешно создан!\n\nWord: {save_path}\nPDF: {pdf_path}")
            except Exception as pdf_error:
                messagebox.showwarning("Частичный успех", f"Word файл создан: {save_path}\nНо PDF не сконвертирован: {pdf_error}")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка при создании документа: {e}")

root = tk.Tk()
root.title("Генератор договоров аренды жилья")
root.geometry("650x900")

canvas = tk.Canvas(root)
scrollbar = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
scrollable_frame = tk.Frame(canvas)

scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
)

canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)

def create_field(parent, label_text):
    label = tk.Label(parent, text=label_text, anchor="w")
    label.pack(fill="x", pady=(10, 0))
    entry = tk.Entry(parent, width=50)
    entry.pack(pady=5)
    return entry

entry_city = create_field(scrollable_frame, "Город:")
entry_date = create_field(scrollable_frame, "Дата (чч.мм.гггг):")
entry_owner = create_field(scrollable_frame, "ФИО Наймодателя:")
entry_position = create_field(scrollable_frame, "Должность Наймодателя:")
entry_basis = create_field(scrollable_frame, "Основание действия (устав/доверенность):")
entry_tenant = create_field(scrollable_frame, "ФИО Нанимателя:")
entry_series = create_field(scrollable_frame, "Серия паспорта:")
entry_num = create_field(scrollable_frame, "Номер паспорта:")
entry_passport_issued = create_field(scrollable_frame, "Кем и когда выдан паспорт:")
entry_code = create_field(scrollable_frame, "Код подразделения:")
entry_address = create_field(scrollable_frame, "Адрес помещения:")
entry_year_built = create_field(scrollable_frame, "Год постройки:")
entry_rooms = create_field(scrollable_frame, "Количество комнат:")
entry_area = create_field(scrollable_frame, "Общая площадь (кв. м.):")
entry_living_area = create_field(scrollable_frame, "Жилая площадь (кв. м.):")
entry_room = create_field(scrollable_frame, "Тип помещения (квартира/дом/часть):")
entry_keys = create_field(scrollable_frame, "Количество ключей:")
entry_price = create_field(scrollable_frame, "Ежемесячная плата:")

btn_frame = tk.Frame(scrollable_frame)
btn_frame.pack(pady=20)

btn_load = tk.Button(btn_frame, text="Загрузить из TXT", command=load_from_txt, 
                     bg="#2196F3", fg="white", font=("Arial", 10))
btn_load.pack(side=tk.LEFT, padx=5)

btn_generate = tk.Button(btn_frame, text="Создать договор", command=generate, 
                         bg="#4CAF50", fg="white", font=("Arial", 10, "bold"))
btn_generate.pack(side=tk.LEFT, padx=5)

info_label = tk.Label(scrollable_frame, 
                      text="Убедитесь, что template.docx в папке со скриптом", 
                      fg="blue", font=("Arial", 9))
info_label.pack(pady=10)

canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

root.mainloop()