# -*- coding: utf-8 -*-
"""
Created on Tue Aug  6 00:49:53 2024

@author: Kevin M Thorsen
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PyPDF2 import PdfReader, PdfWriter, PageObject
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import os
from decimal import Decimal
import subprocess

# Load logos dynamically
def load_logos():
    logos_folder = os.path.join(os.getcwd(), "logos")
    if not os.path.exists(logos_folder):
        os.makedirs(logos_folder)
    return [f for f in os.listdir(logos_folder) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]

def refresh_logos():
    logos = load_logos()
    logo_selector['values'] = logos
    if logos:
        logo_selector.set(logos[0])
    else:
        logo_selector.set("No logos found")

def select_pdf_directory():
    folder_path = filedialog.askdirectory()
    pdf_directory_entry.delete(0, tk.END)
    pdf_directory_entry.insert(0, folder_path)

def open_directory(path):
    if os.path.exists(path):
        if os.name == 'nt':  # For Windows
            os.startfile(path)
        elif os.name == 'posix':  # For MacOS and Linux
            subprocess.Popen(['open' if sys.platform == 'darwin' else 'xdg-open', path])
    else:
        messagebox.showerror("Error", f"Directory does not exist: {path}")

# Create watermark on a blank PDF page
def create_watermark_pdf(watermark_image_path, watermark_pdf_path, page_width, page_height, positions):
    page_width = float(page_width) if isinstance(page_width, Decimal) else page_width
    page_height = float(page_height) if isinstance(page_height, Decimal) else page_height

    c = canvas.Canvas(watermark_pdf_path, pagesize=(page_width, page_height))
    img = ImageReader(watermark_image_path)
    img_width, img_height = map(float, img.getSize())
    c.setFillAlpha(0.3)

    for position in positions:
        if position == 'top_left':
            c.drawImage(img, 10, page_height - img_height - 10, width=img_width, height=img_height, mask='auto')
        elif position == 'top_right':
            c.drawImage(img, page_width - img_width - 10, page_height - img_height - 10, width=img_width, height=img_height, mask='auto')
        elif position == 'bottom_left':
            c.drawImage(img, 10, 10, width=img_width, height=img_height, mask='auto')
        elif position == 'bottom_right':
            c.drawImage(img, page_width - img_width - 10, 10, width=img_width, height=img_height, mask='auto')
    c.save()

# Core watermarking logic
def watermark_pdfs():
    pdf_directory = pdf_directory_entry.get()
    selected_logo = logo_selector.get()
    watermark_image_path = os.path.join(os.getcwd(), "logos", selected_logo)

    if not pdf_directory or not selected_logo:
        messagebox.showerror("Error", "Please select both PDF directory and a logo")
        return

    positions = [pos for var, pos in zip([top_left_var, top_right_var, bottom_left_var, bottom_right_var], 
                                         ['top_left', 'top_right', 'bottom_left', 'bottom_right']) if var.get()]

    if not positions:
        messagebox.showerror("Error", "Please select at least one watermark position")
        return

    output_folder = os.path.join(pdf_directory, "watermarked")
    os.makedirs(output_folder, exist_ok=True)
    pdf_files = [f for f in os.listdir(pdf_directory) if f.endswith(".pdf")]
    progress_bar["maximum"] = len(pdf_files)

    for index, filename in enumerate(pdf_files):
        try:
            input_pdf_path = os.path.join(pdf_directory, filename)
            output_pdf_path = os.path.join(output_folder, f"{os.path.splitext(filename)[0]}_watermarked.pdf")

            with open(input_pdf_path, "rb") as input_pdf_file:
                input_pdf = PdfReader(input_pdf_file)
                pdf_writer = PdfWriter()

                for page_num, page in enumerate(input_pdf.pages):
                    page_width = float(page.mediabox.width) if isinstance(page.mediabox.width, Decimal) else page.mediabox.width
                    page_height = float(page.mediabox.height) if isinstance(page.mediabox.height, Decimal) else page.mediabox.height

                    watermark_pdf_path = f"temp_watermark_{page_num}.pdf"
                    create_watermark_pdf(watermark_image_path, watermark_pdf_path, page_width, page_height, positions)

                    watermark_pdf = PdfReader(watermark_pdf_path)
                    watermark_page = watermark_pdf.pages[0]
                    page.merge_page(watermark_page)
                    pdf_writer.add_page(page)
                    os.remove(watermark_pdf_path)

                with open(output_pdf_path, "wb") as output_pdf_file:
                    pdf_writer.write(output_pdf_file)

            progress_bar["value"] = index + 1
            status_log.insert(tk.END, f"Watermarked: {filename}\n")
            root.update_idletasks()
        except Exception as e:
            status_log.insert(tk.END, f"Error with {filename}: {e}\n")

    messagebox.showinfo("Success", "All PDFs have been watermarked successfully")

# GUI Layout
root = tk.Tk()
root.title("PDF Watermarking Tool")
root.geometry("600x750")

style = ttk.Style()
style.configure('TButton', font=('Arial', 10), padding=6)
style.configure('TLabel', font=('Arial', 10))

# Logo Section
tk.Label(root, text="Select Logo for Watermarking:").pack(pady=5)
logo_selector = ttk.Combobox(root, state="readonly", width=50)
logo_selector.pack(pady=5)
refresh_logos()
tk.Button(root, text="Refresh Logos", command=refresh_logos).pack(pady=5)

# PDF Directory Selection
tk.Label(root, text="Select PDF Directory:").pack(pady=5)
pdf_directory_entry = tk.Entry(root, width=50)
pdf_directory_entry.pack(pady=5)
tk.Button(root, text="Browse", command=select_pdf_directory).pack(pady=5)

# Buttons to open directories
tk.Button(root, text="Open Selected PDF Directory", command=lambda: open_directory(pdf_directory_entry.get())).pack(pady=5)
tk.Button(root, text="Open Watermarked Directory", command=lambda: open_directory(os.path.join(pdf_directory_entry.get(), "watermarked"))).pack(pady=5)

# Watermark Positions Grid
tk.Label(root, text="Select Watermark Positions:").pack(pady=5)

position_frame = tk.Frame(root)
position_frame.pack(pady=5)

# Position Variables
top_left_var = tk.BooleanVar()
top_right_var = tk.BooleanVar()
bottom_left_var = tk.BooleanVar()
bottom_right_var = tk.BooleanVar()

# Arrange Checkbuttons in a grid
position_frame.grid_columnconfigure((0, 1), weight=1)
position_frame.grid_rowconfigure((0, 1), weight=1)

# Position Buttons
tk.Checkbutton(position_frame, text="Top Left", variable=top_left_var).grid(row=0, column=0, sticky="nw", padx=20, pady=20)
tk.Checkbutton(position_frame, text="Top Right", variable=top_right_var).grid(row=0, column=1, sticky="ne", padx=20, pady=20)
tk.Checkbutton(position_frame, text="Bottom Left", variable=bottom_left_var).grid(row=1, column=0, sticky="sw", padx=20, pady=20)
tk.Checkbutton(position_frame, text="Bottom Right", variable=bottom_right_var).grid(row=1, column=1, sticky="se", padx=20, pady=20)

# Progress Bar
progress_bar = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
progress_bar.pack(pady=10)

# Status Log
status_log = tk.Text(root, height=8, width=70)
status_log.pack(pady=10)

# Watermark Button
tk.Button(root, text="Watermark PDFs", command=watermark_pdfs).pack(pady=10)

root.mainloop()
