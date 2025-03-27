from tkinter import filedialog, messagebox, ttk
import tkinter as tk
from datetime import datetime
import os
# from tkinter import Toplevel, Label, DoubleVar
# from tkinter.ttk import Progressbar
from tkinter import *
import customtkinter
import time  # Simulate progress for demonstration

class BillingCheckToolGUI:
    def __init__(self, master):
        # Set the app icon
        icon_path = os.path.join(os.path.dirname(__file__), 'assets', 'Amecor_Logo-01-small.ico')
        if os.path.exists(icon_path):
            try:
                master.iconbitmap(icon_path)
            except tk.TclError:
                messagebox.showwarning("Icon Error", f"Unable to set icon. File is invalid: {icon_path}")
        else:
            messagebox.showwarning("Icon Missing", f"Icon file not found: {icon_path}")
        master.title("Sabre Billing Check Tool")

        # Set window size
        master.geometry("620x600")

        # Add the logo to the GUI
        logo_path = os.path.join(os.path.dirname(__file__), 'assets', 'Amecor_Logo-01-small.png')
        if os.path.exists(logo_path):
            try:
                logo = tk.PhotoImage(file=logo_path)
                tk.Label(master, image=logo).grid(row=0, column=0, columnspan=3, pady=10)
                self.logo = logo  # Keep a reference to avoid garbage collection
            except tk.TclError:
                print(f"Warning: Unable to load logo. File not found or invalid: {logo_path}")
        else:
            tk.Label(master, text="Logo not available").grid(row=0, column=0, columnspan=3, pady=10)

        # Add name/title "SABRE" below the logo
        tk.Label(master, text="SABRE", font=("Copperplate Gothic Bold", 18), fg="#C00000").grid(row=1, column=0, columnspan=3, pady=10)

        # Move all processing functions down
        tk.Label(master, text="Select Prior Month CSV:", fg="#2457FC").grid(row=2, column=0, padx=10, pady=5)
        self.entry_prior = tk.Entry(master, width=50)
        self.entry_prior.grid(row=2, column=1, padx=10, pady=5)
        tk.Button(master, text="Browse", command=self.browse_prior_csv).grid(row=2, column=2, padx=10, pady=5)

        tk.Label(master, text="Select Current Month CSV:", fg="#2457FC").grid(row=3, column=0, padx=10, pady=5)
        self.entry_current = tk.Entry(master, width=50)
        self.entry_current.grid(row=3, column=1, padx=10, pady=5)
        tk.Button(master, text="Browse", command=self.browse_current_csv).grid(row=3, column=2, padx=10, pady=5)

        tk.Label(master, text="Select Sales CSV:", fg="#2457FC").grid(row=4, column=0, padx=10, pady=5)
        self.entry_sales = tk.Entry(master, width=50)
        self.entry_sales.grid(row=4, column=1, padx=10, pady=5)
        tk.Button(master, text="Browse", command=self.browse_sales_csv).grid(row=4, column=2, padx=10, pady=5)

        tk.Label(master, text="Select Delayed Billing Client List CSV:", fg="#2457FC").grid(row=5, column=0, padx=10, pady=5)
        self.entry_delayed = tk.Entry(master, width=50)
        self.entry_delayed.grid(row=5, column=1, padx=10, pady=5)
        tk.Button(master, text="Browse", command=self.browse_delayed_csv).grid(row=5, column=2, padx=10, pady=5)

        tk.Label(master, text="Select Client Device Report CSV:", fg="#2457FC").grid(row=6, column=0, padx=10, pady=5)
        self.entry_client_device = tk.Entry(master, width=50)
        self.entry_client_device.grid(row=6, column=1, padx=10, pady=5)
        tk.Button(master, text="Browse", command=self.browse_client_device_csv).grid(row=6, column=2, padx=10, pady=5)

        tk.Label(master, text="Select Month:", fg="#2457FC").grid(row=7, column=0, padx=10, pady=5)
        self.month_var = tk.StringVar(value=datetime.now().strftime('%m'))
        self.month_dropdown = ttk.Combobox(master, textvariable=self.month_var, values=[f"{i:02d}" for i in range(1, 13)])
        self.month_dropdown.grid(row=7, column=1, padx=10, pady=5, sticky="w")

        # Add tooltip for "Select Month" dropdown
        self.month_dropdown.bind("<Enter>", lambda e: self.show_tooltip("Select the relevant BILLING month"))
        self.month_dropdown.bind("<Leave>", lambda e: self.hide_tooltip())

        tk.Label(master, text="Select Year:", fg="#2457FC").grid(row=8, column=0, padx=10, pady=5)
        self.year_var = tk.StringVar(value=datetime.now().strftime('%Y'))
        self.year_dropdown = ttk.Combobox(master, textvariable=self.year_var, values=[str(i) for i in range(2000, datetime.now().year + 1)])
        self.year_dropdown.grid(row=8, column=1, padx=10, pady=5, sticky="w")

        tk.Label(master, text="Save Processed File As:", fg="#2457FC").grid(row=9, column=0, padx=10, pady=5)
        self.entry_save = tk.Entry(master, width=50)
        self.entry_save.grid(row=9, column=1, padx=10, pady=5)
        tk.Button(master, text="Browse", command=self.browse_save_path).grid(row=9, column=2, padx=10, pady=5)

        tk.Button(master, text="Process", command=self.process, width=20, bg="#2457FC", fg="#FFFFFF").grid(row=10, column=0, columnspan=3, padx=10, pady=20)

        # Add progress bar below the "Process" button
        self.progress_bar = customtkinter.CTkProgressBar(master, orientation="horizontal",
                                                          width=330, 
                                                          height=17, 
                                                          corner_radius=5, 
                                                          border_width=1, 
                                                          border_color="#F2F2F2",
                                                          fg_color="#DEDEDE",
                                                          progress_color="#DEDEDE")  # Initially set to #DEDEDE
        self.progress_bar.grid(row=11, column=0, columnspan=3, pady=10, padx=10)
        self.progress_bar.set(0)  # Initialize progress bar to 0%

        # Remove the progress label as CTkProgressBar does not support embedded labels

    def browse_prior_csv(self):
        filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if filename:
            self.entry_prior.delete(0, tk.END)
            self.entry_prior.insert(0, filename)

    def browse_current_csv(self):
        filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if filename:
            self.entry_current.delete(0, tk.END)
            self.entry_current.insert(0, filename)

    def browse_sales_csv(self):
        filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if filename:
            self.entry_sales.delete(0, tk.END)
            self.entry_sales.insert(0, filename)

    def browse_delayed_csv(self):
        filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if filename:
            self.entry_delayed.delete(0, tk.END)
            self.entry_delayed.insert(0, filename)

    def browse_client_device_csv(self):
        filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if filename:
            self.entry_client_device.delete(0, tk.END)
            self.entry_client_device.insert(0, filename)

    def browse_save_path(self):
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if filename:
            self.entry_save.delete(0, tk.END)
            self.entry_save.insert(0, filename)

    def process(self):
        # Change progress_color to #0336DB when processing starts
        self.progress_bar.configure(progress_color="#0336DB")
        # Simulate progress for demonstration purposes
        for i in range(101):
            self.progress_bar.set(i / 100)  # Update progress bar value
            self.progress_bar.update_idletasks()
            time.sleep(0.05)  # Simulate processing delay

        # Save the processed file
        save_path = self.entry_save.get()
        if save_path:
            try:
                # Replace placeholder with actual logic for saving the processed data
                import pandas as pd
                from processor import process_files

                # Collect input file paths and parameters
                prior_csv = self.entry_prior.get()
                current_csv = self.entry_current.get()
                sales_csv = self.entry_sales.get()
                delayed_csv = self.entry_delayed.get()
                client_device_csv = self.entry_client_device.get()
                selected_month = self.month_var.get()
                selected_year = self.year_var.get()

                # Call the processing function
                process_files(
                    prior_csv, current_csv, sales_csv, delayed_csv, client_device_csv,
                    selected_month, selected_year, save_path
                )

                messagebox.showinfo("Processing Complete", f"The file has been successfully processed and saved to:\n{save_path}")
            except Exception as e:
                messagebox.showerror("Save Error", f"An error occurred while saving the file:\n{e}")
        else:
            messagebox.showwarning("Save Path Missing", "Please specify a valid save path.")

    def show_tooltip(self, message):
        self.tooltip = tk.Toplevel()
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.geometry(f"+{self.month_dropdown.winfo_rootx() + 20}+{self.month_dropdown.winfo_rooty() + 20}")
        label = tk.Label(self.tooltip, text=message, background="white", fg="#C00000", relief="solid", borderwidth=1, font=("Arial", 10))
        label.pack()

    def hide_tooltip(self):
        if hasattr(self, 'tooltip'):
            self.tooltip.destroy()
            del self.tooltip

if __name__ == "__main__":
    root = tk.Tk()
    app = BillingCheckToolGUI(root)
    root.mainloop()