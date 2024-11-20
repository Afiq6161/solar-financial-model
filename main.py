import tkinter as tk
from tkinter import filedialog, messagebox
import threading
from tkinter.ttk import Progressbar
import SolarFinancialModelFunctions as sfm  # Placeholder for your function definitions
import shutil
import os

# Function to export example input file
def export_example_input(save_path):
    try:
        # Assuming the example input file is in the same directory as this script
        example_input_file = "solar_financial_model_inputs.xlsx"
        if os.path.exists(example_input_file):
            shutil.copy(example_input_file, save_path)
            messagebox.showinfo("Success", "Example input file has been saved successfully.")
        else:
            messagebox.showerror("Error", "Example input file not found.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Function to run the solar financial model
def run_solar_financial_model(input_file, save_path, progress_var):
    try:
        progress_var.set(10)
        # Load input file and scenarios
        inputs, scenarios = sfm.load_input_file(input_file)
        progress_var.set(30)

        # Perform solar financial modeling
        results = sfm.run_model(inputs, scenarios, progress_var)
        progress_var.set(70)

        # Save the results
        sfm.save_results(results, save_path)
        progress_var.set(100)

        messagebox.showinfo("Success", "The solar financial model calculations are complete and have been saved successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Function to browse for a file
def browse_file(entry):
    filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    entry.delete(0, tk.END)
    entry.insert(0, filename)

# Function to save a file
def save_file(entry):
    filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    entry.delete(0, tk.END)
    entry.insert(0, filename)

# Function to export example input file
def export_example_input_file():
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    if save_path:
        export_example_input(save_path)

# Function to start the modeling process
def start_process(input_entry, save_entry, progress_var):
    input_file = input_entry.get()
    save_path = save_entry.get()
    
    if not input_file or not save_path:
        messagebox.showwarning("Input Required", "Please provide both the input file and save location.")
        return
    
    progress_var.set(0)
    thread = threading.Thread(target=run_solar_financial_model, args=(input_file, save_path, progress_var))
    thread.start()

# Main function to create UI
def main():
    # Create main window
    root = tk.Tk()
    root.title("Solar Financial Model UI")
    root.geometry("500x400")

    # Input file label and entry
    tk.Label(root, text="Select Input File:").pack(pady=(10, 0))
    input_entry = tk.Entry(root, width=50)
    input_entry.pack(pady=5)
    tk.Button(root, text="Browse", command=lambda: browse_file(input_entry)).pack()

    # Save file label and entry
    tk.Label(root, text="Select Save Location:").pack(pady=(10, 0))
    save_entry = tk.Entry(root, width=50)
    save_entry.pack(pady=5)
    tk.Button(root, text="Browse", command=lambda: save_file(save_entry)).pack()

    # Progress bar
    tk.Label(root, text="Progress:").pack(pady=(20, 0))
    progress_var = tk.DoubleVar()
    progress_bar = Progressbar(root, variable=progress_var, maximum=100)
    progress_bar.pack(pady=5, padx=20, fill=tk.X)

    # Start button
    tk.Button(root, text="Run Model", command=lambda: start_process(input_entry, save_entry, progress_var)).pack(pady=(20, 10))

    # Export example input button
    tk.Button(root, text="Export Example Input File", command=export_example_input_file).pack(pady=(10, 10))

    # Start the main event loop
    root.mainloop()

if __name__ == "__main__":
    main()
