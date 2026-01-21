
А import tkinter as tk
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

            # Заполняем поля
            entry_city.delete(0, tk.END); entry_city.insert(0, loaded_data.get("City", ""))
            entry_date.delete(0, tk.END); entry_date.insert(0, loaded_data.get("ContractDate", ""))
            entry_owner.delete(0, tk.END); entry_owner.insert(0, loaded_data.get("LandlordName", ""))
            entry_tenant.delete(0, tk.END); entry_tenant.insert(0, loaded_data.get("TenantName", ""))
            entry_price.delete(0, tk.END); entry_price.insert(0, loaded_data.get("MonthlyPayment", ""))
            entry_series.delete(0, tk.END); entry_series.insert(0, loaded_data.get("PassportSeries", ""))
            entry_num.delete(0, tk.END); entry_num.insert(0, loaded_data.get("PassportNumber", ""))
            entry_address.delete(0, tk.END); entry_address.insert(0, loaded_data.get("PropertyAddress", ""))
            entry_room.delete(0, tk.END); entry_room.insert(0, loaded_data.get("PropertyType", ""))
            entry_keys.delete(0, tk.END); entry_keys.insert(0, loaded_data.get("KeysCount", ""))

            messagebox.showinfo("Успех", "Данные загружены из TXT!")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить TXT: {e}")

def generate():
    
    data = {
        "\[**вписать нужное**\]": entry_city.get(),  # Город
        "\[**число, месяц, год**\]": entry_date.get(),
        "\[**Указать наименование собственника жилого помещения или управомоченного лица**\]": entry_owner.get(),
        "\[**Ф. И. О. полностью**\]": entry_tenant.get(),
        "\[**цифрами и прописью**\]": entry_price.get(),
        "\[**вписать нужное**\]": entry_series.get(),
        "\[**значение**\]": entry_num.get(),
        "\[**вписать нужное**\]": entry_address.get(),
        "\[**квартира/жилой дом/часть квартиры/часть жилого дома**\]": entry_room.get(),
        "\[**значение**\]": entry_keys.get()
    }

    if not all(data.values()):
        messagebox.showwarning("Ошибка", "Заполните все поля!")
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
            title="Сохранить договор как"
        )
        if save_path:
            doc.save(save_path)
            
            try:
                convert(save_path, save_path.replace(".docx", ".pdf"))
                messagebox.showinfo("Готово", f"Договор успешно создан!\n\nWord: {save_path}\nPDF: {save_path.replace('.docx', '.pdf')}")
            except Exception as pdf_error:
                messagebox.showwarning("Частичный успех", f"Word файл создан: {save_path}\nНо PDF не сконвертирован: {pdf_error}")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка при создании документа: {e}")


root = tk.Tk()
root.title("Генератор договоров аренды")
root.geometry("600x800")

def create_field(label_text):
    label = tk.Label(root, text=label_text)
    label.pack(pady=(10, 0))
    entry = tk.Entry(root, width=40)
    entry.pack(pady=5)
    return entry

entry_city = create_field("Город: ")
entry_date = create_field("Дата (чч.мм.гггг): ")
entry_owner = create_field("ФИО Наймодателя: ")
entry_tenant = create_field("ФИО Нанимателя: ")
entry_price = create_field("Ежемесячная плата: ")
entry_series = create_field("Серия паспорта: ")
entry_num = create_field("Номер паспорта: ")
entry_address = create_field("Адрес: ")
entry_room = create_field("Тип помещения: ")
entry_keys = create_field("Кол-во ключей: ")

btn_load = tk.Button(root, text="Загрузить из TXT", command=load_from_txt, bg="#2196F3", fg="white", font=("Arial", 12, "bold"))
btn_load.pack(pady=10)

btn_generate = tk.Button(root, text="Создать договор", command=generate, bg="#4CAF50", fg="white", font=("Arial", 12, "bold"))
btn_generate.pack(pady=30)



root.mainloop()