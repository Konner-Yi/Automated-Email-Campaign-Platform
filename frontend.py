import tkinter as tk
from tkinter import filedialog, messagebox
from backend import process_mail_merge

class MailMergeAppFrontend:

    #Creates GUI
    def __init__(self, root):
        self.root = root
        self.root.title("Mail Merge Tool")

        self.sheetPath = ""
        self.templatePath = ""

        # GUI Layout

        #Asks user to provide Excel Sheet
        tk.Label(root, text="Select Your Excel Sheet", font=("Arial", 12)).pack(pady=(10, 0))
        tk.Button(root, text="Browse Excel File", command=self.select_sheet).pack(pady=5)
        self.excel_label = tk.Label(root, text="", font=("Arial", 10), fg="gray")
        self.excel_label.pack()

        #Asks user to provide Email Template
        tk.Label(root, text="Select Your Email Template", font=("Arial", 12)).pack(pady=(15, 0))
        tk.Button(root, text="Browse Template File", command=self.select_template).pack(pady=5)
        self.template_label = tk.Label(root, text="", font=("Arial", 10), fg="gray")
        self.template_label.pack()

        #Asks user for email address
        tk.Label(root, text="Your Gmail Address:", font=("Arial", 12)).pack(pady=(20, 0))
        self.email_entry = tk.Entry(root, width=40)
        self.email_entry.pack(pady=5)

        #Asks user for app password
        tk.Label(root, text="App Password:", font=("Arial", 12)).pack(pady=(10, 0))
        self.password_entry = tk.Entry(root, show="*", width=40)
        self.password_entry.pack(pady=5)

        #Button to send emails
        tk.Button(root, text="Send Emails", command=self.send_emails, bg="#4CAF50", fg="white", padx=10, pady=5).pack(pady=20)

    #Prompts user to select a sheet
    def select_sheet(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if filepath:
            self.excel_label.config(text=filepath.split("/")[-1])

    #Prompts user to select a template
    def select_template(self):
        filepath = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
        if filepath:
            self.template_label.config(text=filepath.split("/")[-1])

    #Sends the emails by getting the password and email
    def send_emails(self):

        #Asks user if they are ready to send emails
        answer = messagebox.askyesno("Confirm", "Do you want to send emails to all listed recipients?")
        if (answer):

            email = self.email_entry.get().strip()
            password = self.password_entry.get().strip()

            #Checks if path of template are empty
            if not self.sheetPath or not self.templatePath:
                messagebox.showerror("Error", "Please select both an Excel sheet and a template file.")
                return
            
            #Checks if email or password are empty
            if not email or not password:
                messagebox.showerror("Error", "Please enter your Gmail and app password.")
                return
            
            #Calls on the mailmerge in backend
            try:
                process_mail_merge(self.sheetPath, self.templatePath, email, password)
                messagebox.showinfo("Success", "All emails sent successfully!")
            except Exception as e:
                messagebox.showerror("Failed", f"An error occurred: {str(e)}")

        else:
            return

    
