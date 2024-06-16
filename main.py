import customtkinter
from PyPDF2 import PdfWriter
from tkinter import filedialog
from tkinter.messagebox import showwarning, showinfo
import docx2pdf
import os

#главное окно
app = customtkinter.CTk()
app.geometry("400x400")
app.minsize(400, 400)
app.title("PDF tool")

name_program = customtkinter.CTkLabel(app, text="PDF tools 🔧", fg_color="transparent", font=('Tahoma', 24))
name_program.pack(padx=20, pady=20)

files = []

def merge(): #объединение файлов pdf
    merge_window = app = customtkinter.CTk()
    merge_window.geometry("400x400")
    merge_window.minsize(400, 400)
    merge_window.title("Объединить PDF")

    name = customtkinter.CTkLabel(app, text="Выберите файлы", fg_color="transparent", font=('Tahoma', 24))
    name.pack(padx=20, pady=20)

    textbox = customtkinter.CTkTextbox(app, wrap="word", font=('Tahoma', 14))
    textbox.configure(width=400)
    textbox.pack(padx=20, pady=20)

    def select_files(): #выбор файлов
        global files
        counter = 0
        files = list(filedialog.askopenfilenames())
        textbox.delete("1.0", "end")
        if not files:
            showwarning(title="Предупреждение!", message="Вы ничего не выбрали.")
        else:
            for i in files:
                counter += 1
                textbox.insert("insert", f"{counter}) {i} \n")

    def merge_files(): #объединение файлов pdf
        if len(files) > 1:

            merger = PdfWriter()

            for pdf in files:
                extension = os.path.splitext(pdf)[1].lower()
                if extension == ".pdf":
                    merger.append(pdf)
                    merger.write("new.pdf")
                    merger.close()
                    showinfo(title="Успех!", message="Успех!")
                else:
                    showwarning(title="Предупреждение!", message="Один или несколько файлов не являются pdf")
                    merge_window.destroy()
                    break
        else:
            showwarning(title="Предупреждение!", message="Нужно выбрать несколько файлов!")
    
    select_button = customtkinter.CTkButton(merge_window, text="Выбрать PDF", command=select_files, font=('Tahoma', 14), fg_color="#A7E6FF", text_color="#050C9C", border_color="#3ABEF9")
    select_button.pack(padx=20, pady=0)

    merge_button = customtkinter.CTkButton(merge_window, text="Объединить PDF", command=merge_files, font=('Tahoma', 14), fg_color="#050C9C", text_color="#A7E6FF", border_color="#3ABEF9")
    merge_button.pack(padx=20, pady=5)

    merge_window.mainloop()

def convert(): #конвертация файлов docx в pdf
    convert_window = app = customtkinter.CTk()
    convert_window.geometry("400x400")
    convert_window.minsize(400, 400)
    convert_window.title("Конвертировать docx в pdf")

    name = customtkinter.CTkLabel(app, text="Выберите файлы", fg_color="transparent", font=('Tahoma', 24))
    name.pack(padx=20, pady=20)

    textbox = customtkinter.CTkTextbox(app, wrap="word", font=('Tahoma', 14))
    textbox.configure(width=400)
    textbox.pack(padx=20, pady=20)

    def select_files(): #выбор файлов
        global files
        counter = 0
        files = list(filedialog.askopenfilenames())
        textbox.delete("1.0", "end")
        if not files:
            showwarning(title="Предупреждение!", message="Вы ничего не выбрали.")
        else:
            for i in files:
                counter += 1
                textbox.insert("insert", f"{counter}) {i} \n")

    def convert_files(): #конвертация файлов docx в pdf
        for file in files:
            extension = os.path.splitext(file)[1].lower()
            if extension == ".docx":
                docx2pdf.convert(file)
                showinfo(title="Успех!", message="Успех!")
            else:
                showwarning(title="Предупреждение!", message="Один или несколько файлов не являются docx")
                convert_window.destroy()
                break
    
    select_button = customtkinter.CTkButton(convert_window, text="Выбрать docx", command=select_files, font=('Tahoma', 14), fg_color="#A7E6FF", text_color="#050C9C", border_color="#3ABEF9")
    select_button.pack(padx=20, pady=0)

    merge_button = customtkinter.CTkButton(convert_window, text="Конвертировать", command=convert_files, font=('Tahoma', 14), fg_color="#050C9C", text_color="#A7E6FF", border_color="#3ABEF9")
    merge_button.pack(padx=20, pady=5)

    convert_window.mainloop()

button_merge = customtkinter.CTkButton(app, text="Объединить pdf", command=merge, font=('Tahoma', 14), fg_color="#050C9C", text_color="#A7E6FF", border_color="#3ABEF9")
button_merge.pack(padx=20, pady=5)
button_convert = customtkinter.CTkButton(app, text="Конвертировать docx в pdf", command=convert, font=('Tahoma', 14), fg_color="#050C9C", text_color="#A7E6FF", border_color="#3ABEF9")
button_convert.pack(padx=20, pady=5)

app.mainloop()