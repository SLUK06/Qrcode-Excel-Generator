import qrcode
from openpyxl import load_workbook
import os
from tkinter import *
from tkinter import filedialog, messagebox

def create_qrcode(link, text, name_archive):
    qr = qrcode.QRCode(
        version=8,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=60,
        border=1,
    )
    qr.add_data(link)
    qr.make(fit=True)

    cor_hexadecimal = "#F2EC00"
    imagem_qrcode = qr.make_image(fill_color="black", back_color=cor_hexadecimal)
    name_archive_qrcode = f"{name_archive}.png"
    imagem_qrcode.save(name_archive_qrcode)

    return name_archive_qrcode

def select_folder_excel():
    folder_excel = filedialog.askopenfilename(filetypes=[("archives Excel", "*.xlsx")])
    entry_folder_excel.delete(0, END)
    entry_folder_excel.insert(END, folder_excel)

def select_folder_destination():
    folder_destination = filedialog.askdirectory()
    entry_folder_destination.delete(0, END)
    entry_folder_destination.insert(END, folder_destination)

def generate_qrcode():
    folder_excel = entry_folder_excel.get()
    name_worksheet = entry_name_worksheet.get()
    name_column_link = entry_name_column_link.get()
    name_column_text = entry_name_column_text.get()
    folder_destination = entry_folder_destination.get()

    if not folder_excel or not name_worksheet or not name_column_link or not name_column_text or not folder_destination:
        messagebox.showwarning("warning", "Please fill in all fields.")
        return

    if not os.path.isfile(folder_excel):
        messagebox.showerror("Error", "Excel file not found.")
        return

    try:
        workbook = load_workbook(filename=folder_excel)
        worksheet = workbook[name_worksheet]
        column_links = worksheet[name_column_link]
        column_text = worksheet[name_column_text]

        qrcodes_generated = []  

        for celula_link, celula_text in zip(column_links[1:], column_text[1:]):
            link = celula_link.value
            text = celula_text.value
            name_archive_qrcode = os.path.join(folder_destination, f"QRCODE-{text}")
            name_archive_generated = create_qrcode(link, text, name_archive_qrcode)
            qrcodes_generated.append(name_archive_generated)

        if qrcodes_generated:
            messagebox.showinfo("Success", "QR codes successfully generated!")
    except Exception as e:
        messagebox.showerror("Error", str(e))
        return

window = Tk()
window.title("QR Code Generator")
window.geometry("400x500")
window.configure(bg="#D31145")

label_folder_excel = Label(window, text="Excel file:", bg="#D31145", fg="#FFFFFF", font=("KelloggsSansMedium", 15,"bold"))
label_folder_excel.pack(pady=(10,0))
entry_folder_excel = Entry(window)
entry_folder_excel.pack()
button_select_folder_excel = Button(window, text="Select Folder", font=("KelloggsSansMedium", 10,"bold"), command=select_folder_excel)
button_select_folder_excel.pack(pady=5)

label_name_worksheet = Label(window, text="Sheet Name:", bg="#D31145", fg="#FFFFFF", font=("KelloggsSansMedium", 15,"bold"))
label_name_worksheet.pack(pady=(20,0))
entry_name_worksheet = Entry(window)
entry_name_worksheet.pack()

label_name_column_link = Label(window, text="Column with the Links:", bg="#D31145", fg="#FFFFFF", font=("KelloggsSansMedium", 15,"bold"))
label_name_column_link.pack(pady=(20,0))
entry_name_column_link = Entry(window)
entry_name_column_link.pack()

label_name_column_text = Label(window, text="Text Column:", bg="#D31145", fg="#FFFFFF", font=("KelloggsSansMedium", 15,"bold"))
label_name_column_text.pack(pady=(20,0))
entry_name_column_text = Entry(window)
entry_name_column_text.pack()

label_folder_destination = Label(window, text="Destination Folder:", bg="#D31145", fg="#FFFFFF", font=("KelloggsSansMedium", 15,"bold"))
label_folder_destination.pack(pady=(20,0))
entry_folder_destination = Entry(window)
entry_folder_destination.pack()
button_select_folder_destination = Button(window, text="Select Folder", font=("KelloggsSansMedium", 10,"bold"), command=select_folder_destination)
button_select_folder_destination.pack(pady=5)

button_generate_qrcode = Button(window, text="Generate QR Codes", font=("KelloggsSansMedium", 10,"bold"), command=generate_qrcode, width=20)
button_generate_qrcode.pack(pady=(30,10))

label_signature = Label(window, text="Dev by: SLUK06", bg="#D31145", fg="#FFFFFF", font=("KelloggsSansMedium", 9,"bold"))
label_signature.pack()


window.mainloop()