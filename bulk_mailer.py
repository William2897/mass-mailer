import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from win32com.client import Dispatch
import pandas as pd

import os
import codecs
from bs4 import BeautifulSoup

class BulkMailerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Bulk Mailer")
        self.recipients = []
        self.csv_columns = []
        self.attachment_path = ""
        self.create_widgets()

    def create_widgets(self):
        frame = ttk.Frame(self.root, padding="10")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        ttk.Button(frame, text="Load Existing Text", command=self.load_existing_text).grid(row=0, column=0, sticky=tk.W)
        ttk.Button(frame, text="Save This Text", command=self.save_this_text).grid(row=0, column=1, sticky=tk.W)

        ttk.Label(frame, text="Email Content").grid(row=1, column=0, columnspan=2, sticky=tk.W)
        self.email_content = tk.Text(frame, width=60, height=20)
        self.email_content.grid(row=2, column=0, columnspan=2)

        # Subject Input
        ttk.Label(frame, text="Subject:").grid(row=3, column=0, sticky=tk.E)
        self.subject_entry = ttk.Entry(frame, width=50)
        self.subject_entry.grid(row=3, column=1, sticky=tk.W)

        self.variable_selector = ttk.Combobox(frame, values=self.csv_columns)
        self.variable_selector.grid(row=4, column=0, sticky=tk.W)
        ttk.Button(frame, text="Insert", command=self.insert_variable).grid(row=4, column=1, sticky=tk.W)

        ttk.Button(frame, text="Load Recipients", command=self.load_recipients).grid(row=5, column=0, sticky=tk.W)
        ttk.Button(frame, text="Text Preview", command=self.preview_email).grid(row=5, column=1, sticky=tk.W)
        ttk.Button(frame, text="Add Attachment", command=self.add_attachment).grid(row=6, column=0, sticky=tk.W)
        ttk.Button(frame, text="Send Emails", command=self.send_emails).grid(row=7, column=0, columnspan=2, sticky=tk.E)

    def load_existing_text(self):
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if file_path:
            with open(file_path, 'r') as file:
                self.email_content.delete("1.0", tk.END)
                self.email_content.insert(tk.END, file.read())

    def save_this_text(self):
        file_path = filedialog.asksaveasfilename(filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if file_path:
            with open(file_path, 'w') as file:
                file.write(self.email_content.get("1.0", tk.END))
            messagebox.showinfo("Bulk Mailer", "Template saved successfully!")

    def load_recipients(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if file_path:
            try:
                self.recipients = pd.read_csv(file_path).to_dict('records')
                self.csv_columns = list(self.recipients[0].keys()) if self.recipients else []
                self.variable_selector['values'] = self.csv_columns
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load CSV: {e}")

    def insert_variable(self):
        variable = self.variable_selector.get()
        if variable:
            self.email_content.insert(tk.INSERT, f"{{{variable}}}")

    def preview_email(self):
        if self.recipients:
            first_recipient = self.recipients[0]
            preview_text = self.email_content.get("1.0", tk.END)
            preview_subject = self.subject_entry.get()
            for key in self.csv_columns:
                preview_text = preview_text.replace(f"{{{key}}}", str(first_recipient.get(key, "")))
                preview_subject = preview_subject.replace(f"{{{key}}}", str(first_recipient.get(key, "")))
            self.show_preview_window(preview_text, preview_subject)
        else:
            messagebox.showwarning("Bulk Mailer", "No recipients loaded to preview the email.")

    def show_preview_window(self, text, subject):
        preview_window = tk.Toplevel(self.root)
        preview_window.title("Email Preview")

        ttk.Label(preview_window, text="Subject:").grid(row=0, column=0, sticky=tk.W)
        subject_label = ttk.Label(preview_window, text=subject)
        subject_label.grid(row=0, column=1, sticky=tk.W)

        preview_text = tk.Text(preview_window, width=60, height=20)
        preview_text.insert(tk.END, text)
        preview_text.grid(row=1, column=0, columnspan=2)
        ttk.Button(preview_window, text="Close", command=preview_window.destroy).grid(row=2, column=1, sticky=tk.E)

    def add_attachment(self):
        self.attachment_path = filedialog.askopenfilename()

    def send_emails(self):
        if not self.subject_entry.get():
            messagebox.showwarning("Missing Subject", "Please enter or generate an email subject.")
            return

        outlook = Dispatch('outlook.application')
        word = Dispatch('Word.Application')
        word.Visible = False
        email_item = outlook.CreateItem(0)
        email_item.GetInspector
        signature = email_item.HTMLBody

        # Find the user's default signature
        signature_dir = os.path.join(os.environ['USERPROFILE'], 'AppData\\Roaming\\Microsoft\\Signatures')
        signature_files = [f for f in os.listdir(signature_dir) if f.endswith('.htm')]
        
        if not signature_files:
            messagebox.showwarning("No Signature Found", "No Outlook signature found in the default location.")
            return
        
        signature_html_path = os.path.join(signature_dir, signature_files[0])  # Use the first found signature

        # Read the signature HTML
        with codecs.open(signature_html_path, 'r', 'utf-8', errors='ignore') as file:
            signature_html = file.read()

        # Parse the signature HTML to replace image paths
        soup = BeautifulSoup(signature_html, 'html.parser')
        for img in soup.find_all('img'):
            img_path = img['src']
            if not img_path.startswith('http'):
                img_path = os.path.join(signature_dir, img_path.replace('/', '\\'))
                if os.path.isfile(img_path):
                    attachment = email_item.Attachments.Add(img_path)
                    img['src'] = f"cid:{attachment.ContentID}"
                    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", attachment.ContentID)

        signature = str(soup)

        for recipient in self.recipients:
            try:
                mail = outlook.CreateItem(0)

                # Dynamic Subject
                mail.Subject = self.subject_entry.get()
                for key in self.csv_columns:
                    mail.Subject = mail.Subject.replace(f"{{{key}}}", str(recipient.get(key, "") or ""))

                mail_body = self.email_content.get("1.0", tk.END)
                for key in self.csv_columns:
                    mail_body = mail_body.replace(f"{{{key}}}", str(recipient.get(key, "") or ""))

                # Append the signature
                mail.BodyFormat = 2  # 2 corresponds to HTML format
                mail.HTMLBody = mail_body + signature

                # Set the recipient fields, ensuring they are strings and trimming whitespaces
                to_email = str(recipient.get('Email', '') or '').strip()
                cc = str(recipient.get('CC', '') or '').strip()
                bcc = str(recipient.get('BCC', '') or '').strip()

                mail.To = to_email
                if cc and cc.lower() != 'nan':  # Check if cc is not empty or 'nan'
                    mail.CC = cc
                if bcc and bcc.lower() != 'nan':  # Check if bcc is not empty or 'nan'
                    mail.BCC = bcc

                if self.attachment_path:
                    mail.Attachments.Add(self.attachment_path)
                mail.Send()
            except Exception as e:
                print(f"Failed to send email to {to_email}: {e}")
                if cc:
                    print(f"Failed to send CC to {cc}: {e}")
                if bcc:
                    print(f"Failed to send BCC to {bcc}: {e}")
                continue  # Skip to the next recipient if there is an error
        messagebox.showinfo("Bulk Mailer", "Emails sent successfully!")



root = tk.Tk()
app = BulkMailerApp(root)
root.mainloop()