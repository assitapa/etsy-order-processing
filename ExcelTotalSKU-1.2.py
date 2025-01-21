import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter.messagebox import showinfo
from PIL import Image
import os


def load_and_compress_excel():
    # Open file dialog to select an Excel file
    file_path = filedialog.askopenfilename(title="Select an Excel file",
                                           filetypes=[("Excel files", "*.xlsx *.xls")])
    if not file_path:
        return
    
    # Load the workbook
    workbook = pd.ExcelFile(file_path)

    # Process the 'Credit Card Skins' sheet
    credit_card_skins = workbook.parse('Credit Card Skins')
    credit_card_skins.drop('Order Number', axis=1, inplace=True)
    grouped_data = credit_card_skins.groupby(['SKU', 'Chip Size', 'Finish'], as_index=False).sum()
    sorted_data = grouped_data.sort_values(by=['Finish', 'Chip Size', 'SKU'], ascending=[True, False, True])

    # Load the 'Stickers' sheet and drop the 'Order Number' column
    stickers = workbook.parse('Stickers')
    stickers.drop('Order Number', axis=1, inplace=True)

    # Group by 'SKU' and 'Size', summing the 'Quantity'
    grouped_stickers = stickers.groupby(['SKU', 'Size'], as_index=False).agg({'Quantity': 'sum'})

    # Sort the grouped data by 'SKU' and then by 'Size'
    sorted_stickers = grouped_stickers.sort_values(by=['SKU', 'Size'])

    # Open file dialog to select save location
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx")])
    if not save_path:
        return
    
    # Create a writer object from pandas and save both sheets
    with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
        sorted_data.to_excel(writer, index=False, sheet_name='Credit Card Skins')
        sorted_stickers.to_excel(writer, index=False, sheet_name='Stickers')
        
    messagebox.showinfo("Success", "Data has been processed and saved successfully.")

def generate_sheets(finish_type):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        data = pd.read_excel(file_path)
        selected_data = data[data['Finish'] == finish_type]
        
        save_folder = filedialog.askdirectory()
        if not save_folder:
            showinfo("Info", "No save directory selected. Operation cancelled.")
            return

        images = []
        total_images = 0
        sheet_counter = 1
        
        for _, row in selected_data.iterrows():
            sku = row['SKU'].split(': ')[1]
            quantity = row['Quantity']
            chip_size = row['Chip Size'].capitalize()
            folder_path_chip = chip_size.replace(' chip', '')
            image_sku = sku.replace('SKN-', f"SKN-{folder_path_chip.upper()}-")

            folder_path = f"/Volumes/NetNVME/Documents/Dropship Stores/LocoSkins/Processing/Print Files/{folder_path_chip} Chip"
            file_name = f"{image_sku}.png"
            image_path = f"{folder_path}/{file_name}"

            try:
                base_image = Image.open(image_path).convert("RGBA")
                for _ in range(quantity):
                    images.append(base_image)
                    total_images += 1
                    if total_images == 8:  # Changed from 6 to 8
                        create_sheet(images, save_folder, sheet_counter, finish_type)
                        images = []
                        total_images = 0
                        sheet_counter += 1
            except Exception as e:
                showinfo("Error", f"Failed to open image for {sku}: {str(e)}")

        if images:
            create_sheet(images, save_folder, sheet_counter, finish_type)

def create_sheet(images, save_folder, sheet_number, prefix):
    sheet_width = 2 * images[0].width + 44  # Still 2 columns
    sheet_height = 4 * images[0].height + 3 * 44  # Now 4 rows, adjusted spacing
    sheet = Image.new("RGBA", (sheet_width, sheet_height), (0, 0, 0, 0))

    for index, img in enumerate(images[:8]):  # Changed from 6 to 8
        x = (index % 2) * (img.width + 44)
        y = (index // 2) * (img.height + 44)
        sheet.paste(img, (x, y), img)

    output_path = f"{save_folder}/{prefix}{sheet_number:03}.png"
    sheet.save(output_path, dpi=(300,300))
    showinfo("Success", f"{prefix} sheet created at {output_path} with 300 DPI.")



root = tk.Tk()
root.title("Excel Sheet Generator")

load_button = tk.Button(root, text="Load and Compress Excel File", command=load_and_compress_excel)
load_button.pack(pady=20)

generate_glossy_button = tk.Button(root, text="Generate Glossy Sheets", command=lambda: generate_sheets('Glossy'))
generate_glossy_button.pack(pady=20)

generate_holographic_button = tk.Button(root, text="Generate Holographic Sheets", command=lambda: generate_sheets('Holographic'))
generate_holographic_button.pack(pady=20)


root.mainloop()
