import os
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import win32com.client as win32
from datetime import datetime
import json



class ExcelSearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Search Application")
        self.root.state("zoomed")
        self.root.geometry("1000x600")  # Default window size
        self.folder_path = r"D:\NX_BACKWORK\Database_File\SMT_BOM"
        self.lcr_file_path = r"D:\NX_BACKWORK\Database_File\SMT_LCR\LCR-Correction Record.xlsx"
        self.email_config_file_path = "email_config.json"
        self.file_list = []
        self.data = {}  # To hold the entered data for sending via email
        self.is_data_saved = False  # To track if the data is saved already
        
        # Load email settings from email_config.json
        self.load_email_config()

        # Configure grid layout for responsiveness
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Main Frame
        self.main_frame = ttk.Frame(self.root, padding=10)
        self.main_frame.grid(sticky="nsew")
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(1, weight=1)
        self.main_frame.rowconfigure(4, weight=1)

        # Header Label
        self.header_label = ttk.Label(self.main_frame, text="Excel File Search", font=("Arial", 16))
        self.header_label.grid(row=0, column=0, pady=10)

        # File Selection Area
        self.file_frame = ttk.Frame(self.main_frame, borderwidth=1, relief="solid")
        self.file_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(5, 10))
        self.file_frame.columnconfigure(0, weight=1)
        self.file_frame.rowconfigure(0, weight=1)

        # File Listbox with Scrollbars
        self.file_listbox = tk.Listbox(self.file_frame, selectmode="multiple", font=("Arial", 12))
        self.file_v_scroll = ttk.Scrollbar(self.file_frame, orient="vertical", command=self.file_listbox.yview)
        self.file_h_scroll = ttk.Scrollbar(self.file_frame, orient="horizontal", command=self.file_listbox.xview)
        self.file_listbox.config(yscrollcommand=self.file_v_scroll.set, xscrollcommand=self.file_h_scroll.set)

        self.file_listbox.grid(row=0, column=0, sticky="nsew")
        self.file_v_scroll.grid(row=0, column=1, sticky="ns")
        self.file_h_scroll.grid(row=1, column=0, sticky="ew")

        # Search Entry
        self.search_frame = ttk.Frame(self.main_frame)
        self.search_frame.grid(row=2, column=0, pady=10, sticky="ew")
        self.search_frame.columnconfigure(1, weight=1)

        ttk.Label(self.search_frame, text="Enter Value to Search:").grid(row=0, column=0, padx=5)
        self.search_entry = ttk.Entry(self.search_frame, font=("Arial", 12))
        self.search_entry.grid(row=0, column=1, padx=5, sticky="ew")

        # Buttons
        self.button_frame = ttk.Frame(self.main_frame)
        self.button_frame.grid(row=3, column=0, pady=10)
        self.find_button = ttk.Button(self.button_frame, text="Find", command=self.search)
        self.find_button.grid(row=0, column=0, padx=5)
        self.clear_button = ttk.Button(self.button_frame, text="Clear", command=self.clear_results)
        self.clear_button.grid(row=0, column=1, padx=5)
        
        # Settings Button (now in the same row)
        #ttk.Button(self.main_frame, text="Settings", command=self.open_settings).grid(row=0, column=1, padx=5, pady=10)
        ttk.Button(self.button_frame, text="Settings", command=self.open_settings).grid(row=0, column=2, padx=5)

        # Results Table with Scrollbars
        self.results_frame = ttk.Frame(self.main_frame, borderwidth=1, relief="solid")
        self.results_frame.grid(row=4, column=0, sticky="nsew", padx=10, pady=10)
        self.results_frame.columnconfigure(0, weight=1)
        self.results_frame.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            self.results_frame,
            columns=("Material", "Long Description", "File"),
            show="headings",
        )
        self.tree.heading("Material", text="Material", anchor="w")
        self.tree.heading("Long Description", text="Long Description", anchor="w")
        self.tree.heading("File", text="File", anchor="w")

        # Set bold font for table headers
        bold_font = ("Arial", 12, "bold")
        self.tree.tag_configure("header", font=bold_font)

        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree_v_scroll = ttk.Scrollbar(self.results_frame, orient="vertical", command=self.tree.yview)
        self.tree_h_scroll = ttk.Scrollbar(self.results_frame, orient="horizontal", command=self.tree.xview)
        self.tree.config(yscrollcommand=self.tree_v_scroll.set, xscrollcommand=self.tree_h_scroll.set)

        self.tree_v_scroll.grid(row=0, column=1, sticky="ns")
        self.tree_h_scroll.grid(row=1, column=0, sticky="ew")

        self.tree.bind("<Double-1>", self.on_double_click)

        # Load Files
        self.load_files_from_folder()

        # Bind resizing event
        self.root.bind("<Configure>", self.on_resize)
    
    def load_email_config(self):
        """Load email settings from the email_config.json file."""
        if os.path.exists(self.email_config_file_path):
            with open(self.email_config_file_path, "r") as f:
                self.email_config = json.load(f)
        else:
            # Default email config if the file doesn't exist
            self.email_config = {
                "recipients": ["recipient1@example.com", "recipient2@example.com"],
                "subject": "LCR Correction Data for Material",
                "body": """
                    Dear Concerned,

                    Please find the following LCR correction details:

                    Material: {material}
                    Description: {description}
                    File: {file}
                    Line: {line}
                    Machine & Side: {machine_side}
                    Standard Value: {standard_value}
                    Measured Value: {measured_value}
                    AVL: {avl}
                    Error: {error}
                    Remarks: {remarks}
                    Standard Tol%: {standard_tol}
                    Correction Tol%: {correction_tol}

                    Best regards,
                    Your Team
                """
            }
            with open(self.email_config_file_path, "w") as f:
                json.dump(self.email_config, f, indent=4)

    def open_settings(self):
        """Open the settings dialog to edit the email configuration."""
        settings_window = tk.Toplevel(self.root)
        settings_window.title("Edit Email Settings")
        settings_window.state("zoomed")
        settings_window.geometry("1000x600")
        
        ttk.Label(settings_window, text="Recipients (comma separated):").pack(pady=5)
        recipients_entry = ttk.Entry(settings_window, width=100)
        recipients = self.email_config.get("recipients", ["recipient1@example.com", "recipient2@example.com"])
        recipients_entry.insert(0, ", ".join(recipients))
        recipients_entry.pack(pady=5)

        ttk.Label(settings_window, text="CC (comma separated):").pack(pady=5)
        cc_entry = ttk.Entry(settings_window, width=100)
        cc = self.email_config.get("cc", ["cc1@example.com", "cc2@example.com"])
        cc_entry.insert(0, ", ".join(cc))
        cc_entry.pack(pady=5)

        ttk.Label(settings_window, text="Subject:").pack(pady=5)
        subject_entry = ttk.Entry(settings_window, width=100)
        subject_entry.insert(0, self.email_config.get("subject", "LCR Correction Data for Material"))
        subject_entry.pack(pady=5)

        ttk.Label(settings_window, text="Body Template:").pack(pady=5)
        body_text = tk.Text(settings_window, height=30, width=100)
        body_text.insert("1.0", self.email_config.get("body", ""))
        body_text.pack(pady=5)

        def save_settings():
            """Save the updated email configuration."""
            self.email_config["recipients"] = recipients_entry.get().split(",")
            self.email_config["cc"] = cc_entry.get().split(",")
            self.email_config["subject"] = subject_entry.get()
            self.email_config["body"] = body_text.get("1.0", "end-1c")

            with open(self.email_config_file_path, "w") as f:
                json.dump(self.email_config, f, indent=4)
            messagebox.showinfo("Settings Saved", "Email settings have been updated.")

        save_button = ttk.Button(settings_window, text="Save", command=save_settings)
        save_button.pack(pady=10)

    def load_files_from_folder(self):
        """Load Excel files from the default folder."""
        if not os.path.exists(self.folder_path):
            messagebox.showerror("Error", f"Folder not found: {self.folder_path}")
            return

        all_files = os.listdir(self.folder_path)
        self.file_list = [f for f in all_files if f.lower().endswith(".xlsx")]

        if not self.file_list:
            messagebox.showwarning("No Files", f"No Excel files found in the folder: {self.folder_path}")
            return

        self.file_listbox.delete(0, "end")
        for file in self.file_list:
            self.file_listbox.insert("end", file)

    def search(self):
        """Search for the value in selected Excel files."""
        selected_indices = self.file_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("Warning", "Please select at least one file to search.")
            return

        search_value = self.search_entry.get().strip()
        if not search_value:
            messagebox.showwarning("Warning", "Please enter a value to search.")
            return

        self.clear_results()
        found = False

        for index in selected_indices:
            file_name = self.file_list[index]
            file_path = os.path.join(self.folder_path, file_name)

            try:
                df = pd.read_excel(file_path)
                if "Material" in df.columns or "Internal P/N" in df.columns:
                    material_col = "Material" if "Material" in df.columns else "Internal P/N"
                    desc_col = "Long. Description" if "Long. Description" in df.columns else "Description"

                    filtered_data = df[df[material_col].astype(str) == search_value]
                    for _, row in filtered_data.iterrows():
                        self.tree.insert("", "end", values=(row[material_col], row[desc_col], file_name))
                        found = True
            except Exception as e:
                messagebox.showerror("Error", f"Error reading file {file_name}: {str(e)}")

        if not found:
            messagebox.showinfo("No Results", "No matching data found.")

    def on_double_click(self, event):
            """Handle double-click on result row."""
            selected_item = self.tree.selection()
            if not selected_item:
                return

            # Get the values of the selected row
            row_values = self.tree.item(selected_item[0], "values")
            material, long_description, file_name = row_values

            # Open a popup window for data entry
            popup = tk.Toplevel(self.root)
            popup.title("Enter Additional Data")
            popup.state("zoomed")
            popup.geometry("800x600")  # Adjust size for new fields

            ttk.Label(popup, text=f"Material: {material}", font=("Arial", 12, "bold")).pack(pady=5)
            ttk.Label(popup, text=f"Description: {long_description}", font=("Arial", 12)).pack(pady=5)

            # Input fields for additional data
            fields = [
                "Line",
                "Machine & Side",
                "Standard Value",
                "Measured Value",
                "AVL",
                "Error",
                "Remarks",
                "Standard Tol%",
                "Correction Tol%",
            ]
            entries = {}
            for field in fields:
                frame = ttk.Frame(popup)
                frame.pack(pady=5, padx=10, fill="x")
                ttk.Label(frame, text=field, font=("Arial", 12)).pack(side="left", padx=5)
                entry = ttk.Entry(frame, font=("Arial", 12))
                entry.pack(side="right", padx=5, fill="x", expand=True)
                entries[field] = entry

            # Save button
            def save_data():
                """Save the entered data to the LCR-Correction Record file."""
                data = {field: entries[field].get().strip() for field in fields}
                data["Material"] = material
                data["Description"] = long_description
                data["File"] = file_name
                
                # Add a timestamp for when the data was saved and mailed
                data["Timestamp"] = datetime.now().strftime("%Y-%m-%d %I:%M %p")
                data["Status"] = "Saved"
                
                # Store the data in a global variable so that send_mail can access it
                self.data = data

                # Save to Excel file
                if not os.path.exists(self.lcr_file_path):
                    # Create the file if it doesn't exist
                    df = pd.DataFrame(columns=["Material", "Description", "File"] + fields + ["Timestamp", "Status"])
                    df.to_excel(self.lcr_file_path, index=False)
                    
                df = pd.read_excel(self.lcr_file_path)

                # Ensure the columns in the data dictionary match the DataFrame columns
                columns = df.columns.tolist()
                for col in columns:
                    if col not in data:
                        data[col] = None  # If any column is missing, add None

                # Convert the data dictionary to a DataFrame
                new_row = pd.DataFrame([data])
                df = pd.concat([df, new_row], ignore_index=True)  # Append new data to the existing DataFrame
                df.to_excel(self.lcr_file_path, index=False)

                messagebox.showinfo("Success", "Data saved successfully!")

            # Send Mail button
            def send_email():
                try:
                    # Initialize Outlook application
                    outlook = win32.Dispatch('Outlook.Application')
                    mail = outlook.CreateItem(0)  # 0 represents a mail item

                    # Configure email fields
                    subject = "LCR Correction Data for Material"
                    body = f"""
                    Dear Concerned,

                    Please find the following LCR correction details:

                    Material: {self.data['Material']}
                    Description: {self.data['Description']}
                    File: {self.data['File']}
                    Line: {self.data['Line']}
                    Machine & Side: {self.data['Machine & Side']}
                    Standard Value: {self.data['Standard Value']}
                    Measured Value: {self.data['Measured Value']}
                    AVL: {self.data['AVL']}
                    Error: {self.data['Error']}
                    Remarks: {self.data['Remarks']}
                    Standard Tol%: {self.data['Standard Tol%']}
                    Correction Tol%: {self.data['Correction Tol%']}

                    Best regards,
                    Your Team
                    """
                    
                    # Add recipients from your settings (can be loaded from JSON)
                    recipients = self.email_config.get("recipients", [])
                    cc_list = self.email_config.get("cc", [])  # Retrieve CC from settings

                    # Add recipients
                    for recipient in recipients:
                        mail.Recipients.Add(recipient)

                    # Add CC recipients
                    for cc in cc_list:
                        mail.CC = cc  # CC is a built-in attribute in Outlook mail items

                    # Set the subject and body
                    mail.Subject = subject
                    mail.Body = body
                    
                    # Send the email
                    mail.Send()

                    # After sending the email, update the Excel with status and timestamp
                    self.data["Timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    self.data["Status"] = "Sent"
                    df = pd.read_excel(self.lcr_file_path)

                    # Ensure the columns in the data dictionary match the DataFrame columns
                    columns = df.columns.tolist()
                    for col in columns:
                        if col not in self.data:
                            self.data[col] = None  # If any column is missing, add None

                    # Convert the data dictionary to a DataFrame
                    new_row = pd.DataFrame([self.data])
                    df = pd.concat([df, new_row], ignore_index=True)
                    df.to_excel(self.lcr_file_path, index=False)

                    messagebox.showinfo("Mail Sent", "The data has been emailed to the concerned persons.")
                    popup.destroy()

                except Exception as e:
                    messagebox.showerror("Error", f"Failed to send mail: {str(e)}")

            ttk.Button(popup, text="Save", command=save_data).pack(pady=10)
            ttk.Button(popup, text="Send Mail", command=send_email).pack(pady=10)

    def clear_results(self):
        """Clear the results table."""
        for item in self.tree.get_children():
            self.tree.delete(item)

    def on_resize(self, event):
        """Handle window resize event to adjust layout."""
        self.main_frame.update_idletasks()

def main():
    root = tk.Tk()
    app = ExcelSearchApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()

# A LiE IS A STORY THAT MAKE UP THAT TRUTH IS WHAT HAPPEND
#pyinstaller --onefile --icon=your_icon.ico LCRlog.py
