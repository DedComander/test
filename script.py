П import tkinter as tk
from tkinter import messagebox, filedialog
from docx import Document

def replace(doc, data):
    for paragraph in doc.paragraphs:
        for key, value in data.items():
            paragraph.text = paragraph.text.replace(key, value)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for key, value in data.items():
                        if key in paragraph.text:
                            paragraph.text = paragraph.text.replace(key, value)

def generate():
    data = {
        "{{CITY}}": entry_city.get(),
        "{{DATE}}": entry_date.get(),
        "{{OWNER_NAME}}": entry_owner.get(),
        "{{TENANT_NAME}}": entry_tenant.get(),
        "{{PRICE}}": entry_price.get(),
        "{{PASS_SERIES}}": entry_series.get(),
        "{{PASS_NUM}}": entry_num.get(),
        "{{ADDRESS}}": entry_address.get(),
        "{{ROOM_TYPE}}": entry_room.get(),
        "{{KEYS_QUANTITY}}": entry_keys.get()
    }

    if not all(data.values()):
        messagebox.showwarning("Ошибка", "Заполните все поля!")
        return

    try:
        doc = Document("template.docx")
        replace(doc, data)
        save_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
        if save_path:
            doc.save(save_path)
            messagebox.showinfo("Готово", "Договор успешно создан!")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось открыть template.docx: {e}")

root = tk.Tk()
root.title("Генератор договоров")
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

btn_generate = tk.Button(root, text="Создать файл Word", command=generate, bg="#4CAF50", fg="white", font=("Arial", 12, "bold"))
btn_generate.pack(pady=30)
root.mainloop()