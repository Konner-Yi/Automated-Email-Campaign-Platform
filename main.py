import tkinter as tk
from frontend import MailMergeAppFrontend

# Launch the UI
if __name__ == "__main__":
    root = tk.Tk()
    app = MailMergeAppFrontend(root)
    root.mainloop()