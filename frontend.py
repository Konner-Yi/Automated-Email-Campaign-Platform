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
        tk.Label(root, text="Select Your Excel Sheet", font=("Arial", 12)).pack(pady=(10, 0))
        tk.Button(root, text="Browse Excel File", command=self.select_sheet).pack(pady=5)

        tk.Label(root, text="Select Your Email Template", font=("Arial", 12)).pack(pady=(15, 0))
        tk.Button(root, text="Browse Template File", command=self.select_template).pack(pady=5)

        tk.Label(root, text="Your Gmail Address:", font=("Arial", 12)).pack(pady=(20, 0))
        self.email_entry = tk.Entry(root, width=40)
        self.email_entry.pack(pady=5)

        tk.Label(root, text="App Password:", font=("Arial", 12)).pack(pady=(10, 0))
        self.password_entry = tk.Entry(root, show="*", width=40)
        self.password_entry.pack(pady=5)

        tk.Button(root, text="Send Emails", command=self.send_emails, bg="#4CAF50", fg="white", padx=10, pady=5).pack(pady=20)

    #Prompts user to select a sheet
    def select_sheet(self):
        self.sheetPath = filedialog.askopenfilename(title="Select Excel File")

    #Prompts user to select a template
    def select_template(self):
        self.templatePath = filedialog.askopenfilename(title="Select Template File")

    #Sends the emails by getting the password and email
    def send_emails(self):
        email = self.email_entry.get().strip()
        password = self.password_entry.get().strip()

        #Checks if path of tempalte are empty
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

    

# Launch the UI
if __name__ == "__main__":
    root = tk.Tk()
    app = MailMergeAppFrontend(root)
    root.mainloop()