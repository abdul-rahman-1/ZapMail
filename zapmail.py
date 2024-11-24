import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, Checkbutton, IntVar
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import logging
import re
from PIL import Image, ImageTk 
import webbrowser 

class EmailSenderApp:
    def __init__(self, master):
        self.master = master
        master.title("Email Sender")
        master.geometry("700x850")
        master.configure(bg="#1f1b2e")

        # Setup logging
        logging.basicConfig(filename="email_sender.log", level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

        self.excel_file = ""
        self.recipients = []
        self.columns = []
        self.check_vars = [] 

        # Add logo
        self.logo_image = Image.open("logo.png") 
        self.logo_image = self.logo_image.resize((100, 100), Image.Resampling.LANCZOS) 
        self.logo = ImageTk.PhotoImage(self.logo_image)

        # GUI Elements
        self.header_label = tk.Label(master, text="Email Sender", font=("Georgia", 18, "bold"), fg="#ffd700", bg="#1f1b2e")
        self.header_label.pack(pady=20)

        # Display logo at the top
        self.logo_label = tk.Label(master, image=self.logo, bg="#1f1b2e")
        self.logo_label.pack(pady=10)

        self.smtp_frame = tk.Frame(master, bg="#1f1b2e")
        self.smtp_frame.pack(pady=10)

        tk.Label(self.smtp_frame, text="SMTP Server:", bg="#1f1b2e", fg="#ffffff").grid(row=0, column=0, padx=5, pady=5)
        self.smtp_entry = tk.Entry(self.smtp_frame, width=35, font=("Arial", 11), relief="solid", bg="#2c2a40", fg="#ffffff")
        self.smtp_entry.insert(0, "smtp.hostinger.com")
        self.smtp_entry.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(self.smtp_frame, text="Port Number:", bg="#1f1b2e", fg="#ffffff").grid(row=1, column=0, padx=5, pady=5)
        self.port_entry = tk.Entry(self.smtp_frame, width=35, font=("Arial", 11), relief="solid", bg="#2c2a40", fg="#ffffff")
        self.port_entry.insert(0, "465")
        self.port_entry.grid(row=1, column=1, padx=5, pady=5)

        tk.Label(self.smtp_frame, text="Email:", bg="#1f1b2e", fg="#ffffff").grid(row=2, column=0, padx=5, pady=5)
        self.email_entry = tk.Entry(self.smtp_frame, width=35, font=("Arial", 11), relief="solid", bg="#2c2a40", fg="#ffffff")
        self.email_entry.insert(0, "yourname@example.com")
        self.email_entry.grid(row=2, column=1, padx=5, pady=5)

        tk.Label(self.smtp_frame, text="Password:", bg="#1f1b2e", fg="#ffffff").grid(row=3, column=0, padx=5, pady=5)
        self.password_entry = tk.Entry(self.smtp_frame, show="*", width=35, font=("Arial", 11), relief="solid", bg="#2c2a40", fg="#ffffff")
        self.password_entry.grid(row=3, column=1, padx=5, pady=5)

        self.excel_file_button = tk.Button(
            master, text="Browse Excel File", command=self.browse_excel_file, font=("Arial", 10), bg="#ffd700", fg="#1f1b2e", relief="flat"
        )
        self.excel_file_button.pack(pady=10)
        self.excel_file_label = tk.Label(master, text="", font=("Arial", 10), bg="#1f1b2e", fg="#ffffff")
        self.excel_file_label.pack()

        self.message_label = tk.Label(master, text="Custom Message:", font=("Georgia", 12), bg="#1f1b2e", fg="#ffd700")
        self.message_label.pack(pady=10)
        self.message_text = tk.Text(master, height=15, width=60, font=("Georgia", 12), bg="#2c2a40", fg="#ffffff", wrap="word")
        self.message_text.pack(pady=5, padx=20)

        self.message_text.tag_configure("highlight", foreground="#ffd700")

        self.message_text.bind("<KeyRelease>", self.highlight_variables)
        self.message_text.bind("<Control-v>", self.handle_paste)

        self.send_button = tk.Button(master, text="Send Emails", command=self.send_emails, bg="#4CAF50", fg="white", font=("Helvetica", 12), relief="flat")
        self.send_button.pack(pady=20)

        self.footer_frame = tk.Frame(master, bg="#1f1b2e")
        self.footer_frame.pack(side="bottom", fill="x", pady=10)

        self.name_label = tk.Label(self.footer_frame, text="Abdul Rahman | ", bg="#1f1b2e", fg="#ffd700", font=("Arial", 10))
        self.name_label.pack(side="left", padx=10)

        self.github_link_label = tk.Label(self.footer_frame, text="GitHub: https://github.com/abdul-rahman-1", bg="#1f1b2e", fg="#ffd700", font=("Arial", 10, "underline"))
        self.github_link_label.pack(side="left", padx=1)
        self.github_link_label.bind("<Button-1>", self.open_github_link) 
    def open_github_link(self, event):
        """Open the GitHub link in the default web browser."""
        webbrowser.open("https://github.com/abdul-rahman-1")

    def browse_excel_file(self):
        self.excel_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        self.excel_file_label.config(text=f"Selected file: {self.excel_file}")
        if self.excel_file:
            self.load_recipients_and_columns()

    def load_recipients_and_columns(self):
        """Load recipients and show the popup with recipient statuses."""
        try:
            data = pd.read_excel(self.excel_file)
            self.recipients = data.to_dict("records")
            self.columns = data.columns.tolist()
            self.update_message_template()
            self.show_recipient_popup()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel file.\nError: {e}")

    def update_message_template(self):
        """Update the message template with column placeholders."""
        column_placeholders = ", ".join([f"{{{col}}}" for col in self.columns])
        default_message = f"""Dear {{Name}},\n\nThis email includes the following dynamic variables:\n{column_placeholders}\n\nWarm regards,\nAbdul Rahman\nGitHub: https://github.com/abdul-rahman-1\nWarlock Cyber Security"""
        self.message_text.delete("1.0", tk.END) 
        self.message_text.insert(tk.END, default_message)  
        self.highlight_variables()

    def show_recipient_popup(self):
        """Show a popup with recipient statuses."""
        popup = Toplevel(self.master)
        popup.title("Recipient Statuses")
        popup.geometry(f"300x{50 + len(self.recipients) * 30}")
        popup.configure(bg="#1f1b2e")

        tk.Label(popup, text="Recipient Statuses:", bg="#1f1b2e", fg="#ffffff", font=("Arial", 12)).pack(pady=10)

        self.check_vars = []
        for recipient in self.recipients:
            var = IntVar(value=0) 
            cb = Checkbutton(popup, text=recipient.get("Name", "Unknown"), variable=var, bg="#1f1b2e", fg="#ffffff", state="disabled", selectcolor="#2c2a40")
            cb.pack(anchor="w", padx=10)
            self.check_vars.append((var, recipient))

    def highlight_variables(self, event=None):
        """Highlight only valid Excel column variables inside curly braces."""
        content = self.message_text.get("1.0", tk.END)
        self.message_text.tag_remove("highlight", "1.0", tk.END)

        # Highlight valid variables
        for match in re.finditer(r"{(.*?)}", content):
            variable = match.group(1)
            if variable in self.columns:
                start_idx = f"1.0+{match.start()}c"
                end_idx = f"1.0+{match.end()}c"
                self.message_text.tag_add("highlight", start_idx, end_idx)

    def handle_paste(self, event):
        """Handle pasting content and apply highlighting after the paste."""
        self.master.after(100, self.highlight_variables) 

    def send_emails(self):
        """Send emails and update recipient statuses."""
        if not self.excel_file:
            messagebox.showerror("Error", "Please select an Excel file.")
            return

        smtp_server = self.smtp_entry.get()
        smtp_port = int(self.port_entry.get())
        sender_email = self.email_entry.get()
        sender_password = self.password_entry.get()

        for var, recipient in self.check_vars:
            try:
                recipient_email = recipient.get("Email")
                email_body = self.message_text.get("1.0", tk.END).format(**recipient)
                message = MIMEMultipart()
                message["From"] = sender_email
                message["To"] = recipient_email
                message["Subject"] = "Custom Email with Dynamic Fields"
                message.attach(MIMEText(email_body, "plain"))

                with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
                    server.login(sender_email, sender_password)
                    server.send_message(message)

                logging.info(f"Email sent successfully to {recipient_email}")
                var.set(1) 
            except Exception as e:
                logging.error(f"Failed to send email to {recipient_email}: {e}")
                messagebox.showerror("Error", f"Failed to send email to {recipient_email}.\nError: {e}")

        messagebox.showinfo("Success", "All emails sent successfully!")


if __name__ == "__main__":
    root = tk.Tk()
    email_sender_app = EmailSenderApp(root)
    root.mainloop()
