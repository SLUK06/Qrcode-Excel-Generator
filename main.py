import qrcode
from openpyxl import load_workbook
import os
from tkinter import *
from tkinter import filedialog, messagebox

def criar_qrcode(link, texto, nome_arquivo):
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
    nome_arquivo_qrcode = f"{nome_arquivo}.png"
    imagem_qrcode.save(nome_arquivo_qrcode)

    return nome_arquivo_qrcode

def selecionar_pasta_excel():
    pasta_excel = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    entry_pasta_excel.delete(0, END)
    entry_pasta_excel.insert(END, pasta_excel)

def selecionar_pasta_destino():
    pasta_destino = filedialog.askdirectory()
    entry_pasta_destino.delete(0, END)
    entry_pasta_destino.insert(END, pasta_destino)

def gerar_qrcode():
    pasta_excel = entry_pasta_excel.get()
    nome_planilha = entry_nome_planilha.get()
    nome_coluna_link = entry_nome_coluna_link.get()
    nome_coluna_texto = entry_nome_coluna_texto.get()
    pasta_destino = entry_pasta_destino.get()

    if not pasta_excel or not nome_planilha or not nome_coluna_link or not nome_coluna_texto or not pasta_destino:
        messagebox.showwarning("warning", "Please fill in all fields.")
        return

    if not os.path.isfile(pasta_excel):
        messagebox.showerror("Error", "Excel file not found.")
        return

    try:
        workbook = load_workbook(filename=pasta_excel)
        planilha = workbook[nome_planilha]
        coluna_links = planilha[nome_coluna_link]
        coluna_texto = planilha[nome_coluna_texto]

        qrcodes_gerados = []  # Lista para armazenar os nomes dos QR codes gerados

        for celula_link, celula_texto in zip(coluna_links[1:], coluna_texto[1:]):
            link = celula_link.value
            texto = celula_texto.value
            nome_arquivo_qrcode = os.path.join(pasta_destino, f"QRCODE-{texto}")
            nome_arquivo_gerado = criar_qrcode(link, texto, nome_arquivo_qrcode)
            qrcodes_gerados.append(nome_arquivo_gerado)

        if qrcodes_gerados:
            messagebox.showinfo("Success", "QR codes successfully generated!")
    except Exception as e:
        messagebox.showerror("Error", str(e))
        return

janela = Tk()
janela.title("QR Code Generator")
janela.geometry("400x500")
janela.configure(bg="#D31145")

label_pasta_excel = Label(janela, text="Excel file:", bg="#D31145", fg="#FFFFFF", font=("KelloggsSansMedium", 15,"bold"))
label_pasta_excel.pack(pady=(10,0))
entry_pasta_excel = Entry(janela)
entry_pasta_excel.pack()
button_selecionar_pasta_excel = Button(janela, text="Select Folder", font=("KelloggsSansMedium", 10,"bold"), command=selecionar_pasta_excel)
button_selecionar_pasta_excel.pack(pady=5)

label_nome_planilha = Label(janela, text="Sheet Name:", bg="#D31145", fg="#FFFFFF", font=("KelloggsSansMedium", 15,"bold"))
label_nome_planilha.pack(pady=(20,0))
entry_nome_planilha = Entry(janela)
entry_nome_planilha.pack()

label_nome_coluna_link = Label(janela, text="Column with the Links:", bg="#D31145", fg="#FFFFFF", font=("KelloggsSansMedium", 15,"bold"))
label_nome_coluna_link.pack(pady=(20,0))
entry_nome_coluna_link = Entry(janela)
entry_nome_coluna_link.pack()

label_nome_coluna_texto = Label(janela, text="Text Column:", bg="#D31145", fg="#FFFFFF", font=("KelloggsSansMedium", 15,"bold"))
label_nome_coluna_texto.pack(pady=(20,0))
entry_nome_coluna_texto = Entry(janela)
entry_nome_coluna_texto.pack()

label_pasta_destino = Label(janela, text="Destination Folder:", bg="#D31145", fg="#FFFFFF", font=("KelloggsSansMedium", 15,"bold"))
label_pasta_destino.pack(pady=(20,0))
entry_pasta_destino = Entry(janela)
entry_pasta_destino.pack()
button_selecionar_pasta_destino = Button(janela, text="Select Folder", font=("KelloggsSansMedium", 10,"bold"), command=selecionar_pasta_destino)
button_selecionar_pasta_destino.pack(pady=5)

button_gerar_qrcode = Button(janela, text="Generate QR Codes", font=("KelloggsSansMedium", 10,"bold"), command=gerar_qrcode, width=20)
button_gerar_qrcode.pack(pady=(30,10))

label_assinatura = Label(janela, text="Dev by: SLUK06", bg="#D31145", fg="#FFFFFF", font=("KelloggsSansMedium", 9,"bold"))
label_assinatura.pack()

janela.mainloop()