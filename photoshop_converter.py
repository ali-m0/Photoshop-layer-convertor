import os
import comtypes.client
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

class DialogModes:
    NO = 3

class PsSaveOptionsConstants:
    psDoNotSaveChanges = 2  # Constant for not saving changes when closing documents

def open_photoshop(template_path):
    try:
        ps_app = comtypes.client.GetActiveObject("Photoshop.Application")
        print("Connected to existing Photoshop instance.")
    except (OSError, comtypes.COMError):
        ps_app = comtypes.client.CreateObject("Photoshop.Application")
        print("Created new Photoshop instance.")
    
    ps_app.Visible = False  # Photoshop won't be visible
    ps_app.DisplayDialogs = DialogModes.NO  # Suppress all dialog boxes
    
    ps_doc = ps_app.Open(template_path)
    
    return ps_app, ps_doc

def find_layer_by_name(layers, name):
    for layer in layers:
        try:
            print(f"Checking layer: {layer.Name}")
            if layer.Name == name:
                return layer
            if hasattr(layer, 'Layers') and layer.Layers.Count > 0:  # If the layer is a group, search within it
                found_layer = find_layer_by_name(layer.Layers, name)
                if found_layer:
                    return found_layer
        except Exception as e:
            print(f"Error accessing layer '{layer.Name}': {e}")
    return None

def update_layers_and_save(ps_app, template_path, data, output_folder, column_mapping, filename_fields):
    # Re-open the document for each row to ensure it starts fresh
    ps_doc = ps_app.Open(template_path)

    for layer_name, file_column in column_mapping.items():
        try:
            layer = find_layer_by_name(ps_doc.Layers, layer_name)
            if not layer:
                print(f"Layer '{layer_name}' not found in the document. Skipping.")
                continue

            if hasattr(layer, 'TextItem'):
                layer.TextItem.Contents = str(data[file_column])
            else:
                print(f"Layer '{layer_name}' is not a text layer. Skipping.")
                
        except Exception as e:
            print(f"Error processing layer '{layer_name}' with value '{data[file_column]}': {e}")

    # Generate filename
    filename = "_".join(str(data[field]) for field in filename_fields)
    psd_output_path = os.path.join(output_folder, filename + ".psd")

    # Save PSD
    try:
        ps_doc.SaveAs(psd_output_path)  # Simplified the SaveAs call for PSD
        print(f"Saved PSD: {psd_output_path}")
    except Exception as e:
        print(f"Failed to save PSD: {e}")

    # Save as JPG
    jpg_output_folder = os.path.join(output_folder, "jpg")
    if not os.path.exists(jpg_output_folder):
        os.makedirs(jpg_output_folder)
    jpg_output_path = os.path.join(jpg_output_folder, filename + ".jpg")

    # Create and apply JPEG save options
    jpeg_options = comtypes.client.CreateObject("Photoshop.JPEGSaveOptions")
    jpeg_options.Quality = 12

    # Save as JPEG
    try:
        ps_doc.SaveAs(jpg_output_path, jpeg_options)
        print(f"Saved JPG: {jpg_output_path}")
    except Exception as e:
        print(f"Failed to save JPG: {e}")

    # Close the document after saving
    try:
        ps_doc.Close(PsSaveOptionsConstants.psDoNotSaveChanges)
    except Exception as e:
        print(f"Error closing the document: {e}")

def process_file_and_generate_images(ps_app, template_path, file_path, output_folder, column_mapping, filename_fields):
    # Determine if the file is Excel or CSV and read accordingly
    if file_path.endswith('.csv'):
        data_frame = pd.read_csv(file_path)
    else:
        data_frame = pd.read_excel(file_path)
    
    for index, row in data_frame.iterrows():
        print(f"Processing row {index + 1}")
        update_layers_and_save(ps_app, template_path, row, output_folder, column_mapping, filename_fields)

def browse_psd_file(psd_file_entry):
    file_path = filedialog.askopenfilename(filetypes=[("PSD files", "*.psd")])
    if file_path:
        psd_file_entry.config(state=tk.NORMAL)
        psd_file_entry.delete(0, tk.END)
        psd_file_entry.insert(0, file_path)
        psd_file_entry.config(state=tk.DISABLED)

def browse_file(file_entry):
    file_path = filedialog.askopenfilename(filetypes=[("Excel or CSV files", "*.xlsx *.xls *.csv")])
    if file_path:
        file_entry.config(state=tk.NORMAL)
        file_entry.delete(0, tk.END)
        file_entry.insert(0, file_path)
        file_entry.config(state=tk.DISABLED)

def browse_output_folder(output_folder_entry):
    folder_path = filedialog.askdirectory()
    if folder_path:
        output_folder_entry.config(state=tk.NORMAL)
        output_folder_entry.delete(0, tk.END)
        output_folder_entry.insert(0, folder_path)
        output_folder_entry.config(state=tk.DISABLED)

def confirm_mapping(root, file_entry, psd_file_entry, output_folder_entry):
    column_mapping = {}
    for layer_var, column_var in zip(layer_var_list, column_var_list):
        layer_name = layer_var.get().strip()
        column_name = column_var.get().strip()
        if layer_name and column_name:
            column_mapping[layer_name] = column_name

    if not column_mapping:
        messagebox.showwarning("Mapping Error", "No valid layer-column mappings provided.")
        return

    select_filename_fields(root, file_entry, psd_file_entry, output_folder_entry, column_mapping)

def map_columns_to_layers(root, file_entry, psd_file_entry, output_folder_entry):
    try:
        # Determine if the file is Excel or CSV and read columns accordingly
        file_path = file_entry.get()
        if file_path.endswith('.csv'):
            columns = pd.read_csv(file_path).columns.tolist()
        else:
            columns = pd.read_excel(file_path).columns.tolist()
    except Exception as e:
        messagebox.showerror("File Error", f"Failed to read file: {e}")
        return

    mapping_window = tk.Toplevel(root)
    mapping_window.title("Map Columns to Photoshop Layers")

    global layer_var_list, column_var_list
    layer_var_list = []
    column_var_list = []

    # Mandatory field mapping
    layer_label = tk.Label(mapping_window, text="Mandatory Layer:")
    layer_label.grid(row=0, column=0, padx=10, pady=5, sticky=tk.E)
    
    layer_var = tk.StringVar()
    layer_entry = tk.Entry(mapping_window, textvariable=layer_var, width=30)
    layer_entry.grid(row=0, column=1, padx=10, pady=5)
    layer_var_list.append(layer_var)

    column_label = tk.Label(mapping_window, text="File Column:")
    column_label.grid(row=0, column=2, padx=10, pady=5, sticky=tk.E)
    
    column_var = tk.StringVar()
    column_combobox = ttk.Combobox(mapping_window, textvariable=column_var, values=columns, width=30)
    column_combobox.grid(row=0, column=3, padx=10, pady=5)
    column_var_list.append(column_var)

    # Optional fields
    for i in range(1, 5):  # Allow up to 4 optional mappings
        layer_label = tk.Label(mapping_window, text=f"Optional Layer {i}:")
        layer_label.grid(row=i, column=0, padx=10, pady=5, sticky=tk.E)
        
        layer_var = tk.StringVar()
        layer_entry = tk.Entry(mapping_window, textvariable=layer_var, width=30)
        layer_entry.grid(row=i, column=1, padx=10, pady=5)
        layer_var_list.append(layer_var)

        column_label = tk.Label(mapping_window, text="File Column:")
        column_label.grid(row=i, column=2, padx=10, pady=5, sticky=tk.E)
        
        column_var = tk.StringVar()
        column_combobox = ttk.Combobox(mapping_window, textvariable=column_var, values=columns, width=30)
        column_combobox.grid(row=i, column=3, padx=10, pady=5)
        column_var_list.append(column_var)

    # Update this line to correctly pass psd_file_entry
    confirm_button = tk.Button(mapping_window, text="Confirm", command=lambda: confirm_mapping(root, file_entry, psd_file_entry, output_folder_entry))
    confirm_button.grid(row=5, column=1, pady=10)

def select_filename_fields(root, file_entry, psd_file_entry, output_folder_entry, column_mapping):
    try:
        # Determine if the file is Excel or CSV and read columns accordingly
        file_path = file_entry.get()
        if file_path.endswith('.csv'):
            columns = pd.read_csv(file_path).columns.tolist()
        else:
            columns = pd.read_excel(file_path).columns.tolist()
    except Exception as e:
        messagebox.showerror("File Error", f"Failed to read file: {e}")
        return

    filename_window = tk.Toplevel(root)
    filename_window.title("Select Filename Fields")

    filename_var_list = []
    for i, column in enumerate(columns):
        var = tk.IntVar()
        chk = tk.Checkbutton(filename_window, text=column, variable=var)
        chk.grid(row=i, column=0, padx=10, pady=5, sticky=tk.W)
        filename_var_list.append((var, column))

    def confirm_filename_fields_inner():
        filename_fields = [column for var, column in filename_var_list if var.get() == 1]
        if not filename_fields:
            messagebox.showwarning("Filename Error", "No fields selected for filenames.")
            return
        process_files(column_mapping, filename_fields, psd_file_entry, file_entry, output_folder_entry)
        filename_window.destroy()

    confirm_button = tk.Button(filename_window, text="Confirm", command=confirm_filename_fields_inner)
    confirm_button.grid(row=len(columns) + 1, column=0, pady=10)

def process_files(column_mapping, filename_fields, psd_file_entry, file_entry, output_folder_entry):
    psd_path = psd_file_entry.get()
    file_path = file_entry.get()
    output_folder = output_folder_entry.get()

    if not psd_path or not file_path or not output_folder:
        messagebox.showwarning("Input Error", "Please fill in all fields.")
        return

    try:
        ps_app, _ = open_photoshop(psd_path)
    except Exception as e:
        messagebox.showerror("Photoshop Error", f"Failed to open Photoshop: {e}")
        return

    try:
        process_file_and_generate_images(ps_app, psd_path, file_path, output_folder, column_mapping, filename_fields)
    except Exception as e:
        messagebox.showerror("Processing Error", f"An error occurred during processing: {e}")
        return

    messagebox.showinfo("Success", "Files processed successfully.")

def create_photoshop_tab(tab):
    # Create widgets for the Photoshop tab
    tk.Label(tab, text="Photoshop PSD File:").grid(row=0, column=0, padx=10, pady=5, sticky=tk.E)
    psd_file_entry = ttk.Entry(tab, width=50, state=tk.DISABLED)
    psd_file_entry.grid(row=0, column=1, padx=10, pady=5)
    tk.Button(tab, text="Browse...", command=lambda: browse_psd_file(psd_file_entry)).grid(row=0, column=2, padx=10, pady=5)

    tk.Label(tab, text="Excel/CSV File:").grid(row=1, column=0, padx=10, pady=5, sticky=tk.E)
    file_entry = ttk.Entry(tab, width=50, state=tk.DISABLED)
    file_entry.grid(row=1, column=1, padx=10, pady=5)
    tk.Button(tab, text="Browse...", command=lambda: browse_file(file_entry)).grid(row=1, column=2, padx=10, pady=5)

    tk.Label(tab, text="Output Folder:").grid(row=2, column=0, padx=10, pady=5, sticky=tk.E)
    output_folder_entry = ttk.Entry(tab, width=50, state=tk.DISABLED)
    output_folder_entry.grid(row=2, column=1, padx=10, pady=5)
    tk.Button(tab, text="Browse...", command=lambda: browse_output_folder(output_folder_entry)).grid(row=2, column=2, padx=10, pady=5)

    tk.Button(tab, text="Map Columns to Layers", command=lambda: map_columns_to_layers(tab, file_entry, psd_file_entry, output_folder_entry)).grid(row=3, column=1, pady=10)