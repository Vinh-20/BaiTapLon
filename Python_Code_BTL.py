import tkinter as tk
from tkinter import filedialog, messagebox
import fitz  # PyMuPDF
from pdf2docx import Converter
from docx2pdf import convert # type: ignore

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDF and Word files", "*.pdf;*.docx")])
    file_path_entry.delete(0, tk.END)
    file_path_entry.insert(0, file_path)

def convert_to_text(pdf_path, output_path):
    doc = fitz.open(pdf_path)
    with open(output_path, 'w', encoding='utf-8') as f:
        for page in doc:
            text = page.get_text()
            f.write(text)
    messagebox.showinfo("Thông báo", "Chuyển đổi PDF sang Text thành công!")

def convert_to_word(pdf_path, output_path):
    cv = Converter(pdf_path)
    cv.convert(output_path, start=0, end=None)
    cv.close()
    messagebox.showinfo("Thông báo", "Chuyển đổi PDF sang Word thành công!")

def convert_to_pdf(docx_path, output_path):
    convert(docx_path, output_path)
    messagebox.showinfo("Thông báo", "Chuyển đổi Word sang PDF thành công!")

def convert_file():
    file_path = file_path_entry.get()
    output_format = format_var.get()
    if not file_path:
        messagebox.showwarning("Cảnh báo", "Vui lòng chọn file!")
        return

    output_path = filedialog.asksaveasfilename(defaultextension=f".{output_format}",
                                               filetypes=[(f"{output_format.upper()} files", f"*.{output_format}")])

    if output_format == "txt":
        convert_to_text(file_path, output_path)
    elif output_format == "docx":
        convert_to_word(file_path, output_path)
    elif output_format == "pdf":
        # Check if the input file is a DOCX file
        if file_path.lower().endswith(".docx"):
            convert_to_pdf(file_path, output_path)
        else:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn file DOCX để chuyển đổi sang PDF!")

# Tạo cửa sổ chính
root = tk.Tk()
root.title("PDF & DOCX Converter")

# Chọn file
tk.Label(root, text="Chọn file:").grid(row=0, column=0, padx=10, pady=10)
file_path_entry = tk.Entry(root, width=50)
file_path_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Chọn file", command=select_file).grid(row=0, column=2, padx=10, pady=10)

# Chọn định dạng đầu ra
tk.Label(root, text="Chọn định dạng đầu ra:").grid(row=1, column=0, padx=10, pady=10)
format_var = tk.StringVar(value="txt")
tk.Radiobutton(root, text="Text", variable=format_var, value="txt").grid(row=1, column=1)
tk.Radiobutton(root, text="Word", variable=format_var, value="docx").grid(row=1, column=2)
tk.Radiobutton(root, text="PDF", variable=format_var, value="pdf").grid(row=1, column=3)

# Nút chuyển đổi
tk.Button(root, text="Chuyển đổi", command=convert_file).grid(row=2, column=0, columnspan=4, padx=10, pady=10)

root.mainloop()
