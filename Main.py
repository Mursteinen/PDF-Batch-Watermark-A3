# -*- coding: utf-8 -*-
"""
Created on Tue Aug  6 00:49:53 2024

@author: Kevin M Thorsen
"""

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from PyPDF2 import PdfReader, PdfWriter, PageObject
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import os

def load_logos():
    logos_folder = os.path.join(os.getcwd(), "logos")
    if not os.path.exists(logos_folder):
        os.makedirs(logos_folder)  # Create the folder if it doesn't exist
    return [f for f in os.listdir(logos_folder) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]

def refresh_logos():
    logos = load_logos()
    logo_selector['values'] = logos
    if logos:
        logo_selector.set(logos[0])  # Set the first logo as default
    else:
        logo_selector.set("No logos found")


def select_pdf_directory():
    folder_path = filedialog.askdirectory()
    pdf_directory_entry.delete(0, tk.END)
    pdf_directory_entry.insert(0, folder_path)

from decimal import Decimal

def create_watermark_pdf(watermark_image_path, watermark_pdf_path, page_width, page_height, positions):
    # Ensure dimensions are floats
    page_width = float(page_width) if isinstance(page_width, (Decimal, float)) else page_width
    page_height = float(page_height) if isinstance(page_height, (Decimal, float)) else page_height

    c = canvas.Canvas(watermark_pdf_path, pagesize=(page_width, page_height))
    img = ImageReader(watermark_image_path)
    img_width, img_height = map(float, img.getSize())  # Convert image dimensions to float
    c.setFillAlpha(0.3)  # Set transparency level

    # Calculate positions
    for position in positions:
        if position == 'top_left':
            c.drawImage(img, 0, page_height - img_height, width=img_width, height=img_height, mask='auto')
        elif position == 'top_right':
            c.drawImage(img, page_width - img_width, page_height - img_height, width=img_width, height=img_height, mask='auto')
        elif position == 'bottom_left':
            c.drawImage(img, 0, 0, width=img_width, height=img_height, mask='auto')
        elif position == 'bottom_right':
            c.drawImage(img, page_width - img_width, 0, width=img_width, height=img_height, mask='auto')

    c.save()



def watermark_pdfs():
    pdf_directory = pdf_directory_entry.get()
    selected_logo = logo_selector.get()
    watermark_image_path = os.path.join(os.getcwd(), "logos", selected_logo)

    if not pdf_directory:
        messagebox.showerror("Error", "Please select a PDF directory")
        return

    if not selected_logo:
        messagebox.showerror("Error", "Please select a logo for watermarking")
        return

    positions = []
    if top_left_var.get():
        positions.append('top_left')
    if top_right_var.get():
        positions.append('top_right')
    if bottom_left_var.get():
        positions.append('bottom_left')
    if bottom_right_var.get():
        positions.append('bottom_right')

    if not positions:
        messagebox.showerror("Error", "Please select at least one watermark position")
        return

    try:
        output_folder = os.path.join(pdf_directory, "watermarked")
        os.makedirs(output_folder, exist_ok=True)

        pdf_files = [f for f in os.listdir(pdf_directory) if f.endswith(".pdf")]
        progress_bar["maximum"] = len(pdf_files)

        for index, filename in enumerate(pdf_files):
            input_pdf_path = os.path.join(pdf_directory, filename)
            output_pdf_path = os.path.join(output_folder, f"{os.path.splitext(filename)[0]}_watermarked.pdf")

            with open(input_pdf_path, "rb") as input_pdf_file:
                input_pdf = PdfReader(input_pdf_file)
                pdf_writer = PdfWriter()

                for page_num in range(len(input_pdf.pages)):
                    page = input_pdf.pages[page_num]
                    page_width = page.mediabox.width
                    page_height = page.mediabox.height

                    # Create watermark PDF for each page size
                    watermark_pdf_path = f"temp_watermark_{page_num}.pdf"
                    create_watermark_pdf(watermark_image_path, watermark_pdf_path, page_width, page_height, positions)

                    watermark_pdf = PdfReader(watermark_pdf_path)
                    watermark_page = watermark_pdf.pages[0]

                    # Create a new page object for the watermark
                    new_page = PageObject.create_blank_page(width=page_width, height=page_height)
                    new_page.merge_page(page)
                    new_page.merge_page(watermark_page)
                    pdf_writer.add_page(new_page)

                    os.remove(watermark_pdf_path)

                with open(output_pdf_path, "wb") as output_pdf_file:
                    pdf_writer.write(output_pdf_file)

            progress_bar["value"] = index + 1
            root.update_idletasks()

        messagebox.showinfo("Success", "All PDFs have been watermarked successfully")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


# Create the main application window
root = tk.Tk()
root.title("PDF Watermarking Tool")
root.geometry("500x500")

# Add Logo Selection Section
tk.Label(root, text="Select logo for watermarking:").pack(pady=5)
logo_selector = ttk.Combobox(root, state="readonly", width=50)
logo_selector.pack(pady=5)
refresh_logos()
tk.Button(root, text="Refresh Logos", command=refresh_logos).pack(pady=5)

# Input PDF directory selection
tk.Label(root, text="Select PDF directory:").pack(pady=5)
pdf_directory_entry = tk.Entry(root, width=50)
pdf_directory_entry.pack(pady=5)
tk.Button(root, text="Browse", command=select_pdf_directory).pack(pady=5)

# Watermark position options
tk.Label(root, text="Select watermark positions:").pack(pady=5)
top_left_var = tk.BooleanVar()
top_right_var = tk.BooleanVar()
bottom_left_var = tk.BooleanVar()
bottom_right_var = tk.BooleanVar()
tk.Checkbutton(root, text="Top Left", variable=top_left_var).pack(anchor='w')
tk.Checkbutton(root, text="Top Right", variable=top_right_var).pack(anchor='w')
tk.Checkbutton(root, text="Bottom Left", variable=bottom_left_var).pack(anchor='w')
tk.Checkbutton(root, text="Bottom Right", variable=bottom_right_var).pack(anchor='w')

# Progress bar
progress_bar = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
progress_bar.pack(pady=20)

# Watermark button
tk.Button(root, text="Watermark PDFs", command=watermark_pdfs).pack(pady=20)

# Run the application
root.mainloop()
