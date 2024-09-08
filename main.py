import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import smtplib
import ssl
import pandas as pd
import random
import string
import zipfile
import os
from datetime import datetime, timedelta
from email.message import EmailMessage
import threading
import time

# Initialize tracking list
tracking_status = []

# Function to send an email
def send_email():
    smtp_server = smtp_server_var.get()
    port = int(port_var.get())
    sender_email = email_var.get()
    password = password_var.get()
    subject_template = subject_var.get()
    body_template = body_text.get("1.0", tk.END)
    attachment = attachment_var.get()

    try:
        context = ssl.create_default_context()
        with smtplib.SMTP(smtp_server, port) as server:
            server.starttls(context=context)
            server.login(sender_email, password)
            
            if 'recipients' in globals() and recipients:
                for recipient in recipients:
                    try:
                        # Replace placeholders with actual values
                        subject = subject_template.format(name=recipient["Name"], company=recipient["Company"])
                        body = body_template.format(name=recipient["Name"], company=recipient["Company"])

                        # Debug: Print subject and body to verify correctness
                        print(f"Sending email to {recipient['Email']}")
                        print(f"Subject: {subject}")
                        print(f"Body: {body}")

                        # Create and send the email with attachment
                        message = EmailMessage()
                        message["Subject"] = subject
                        message["From"] = sender_email
                        message["To"] = recipient["Email"]
                        message.set_content(body)

                        if attachment:
                            with open(attachment, "rb") as f:
                                file_data = f.read()
                                file_name = os.path.basename(attachment)
                                message.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_name)

                        server.send_message(message)
                        
                        # Update tracking status
                        tracking_status.append(f"Email to {recipient['Email']} sent successfully.")
                    except Exception as e:
                        # Update tracking status with failure
                        tracking_status.append(f"Failed to send email to {recipient['Email']}: {e}")

                # Display tracking status
                update_status_display()
                messagebox.showinfo("Success", "Emails processed. Check the status below.")
            else:
                messagebox.showerror("Error", "No recipients available. Please import the recipient list.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to send email: {e}")

# Function to update the status display
def update_status_display():
    status_display.config(state=tk.NORMAL)
    status_display.delete(1.0, tk.END)
    for status in tracking_status:
        status_display.insert(tk.END, status + "\n")
    status_display.config(state=tk.DISABLED)

# Function to attach a file
def attach_file():
    file_path = filedialog.askopenfilename()
    if file_path:
        attachment_var.set(file_path)
        if file_path.endswith('.zip'):
            # Automatically unzip and attach all files if zip is selected
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                zip_ref.extractall("attachments")
            attachment_var.set("attachments")

# Function to import recipients from a CSV or Excel file
def import_recipients():
    global recipients
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx *.xls")])
    if file_path:
        try:
            if file_path.endswith(".csv"):
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)

            # Check if the necessary columns are present
            required_columns = ["Email", "Name", "Company"]
            if not all(col in df.columns for col in required_columns):
                messagebox.showerror("Error", f"The file must contain the following columns: {', '.join(required_columns)}.")
                return

            # Convert DataFrame to a list of dictionaries
            recipients = df.to_dict(orient='records')
            recipient_list_var.set(', '.join([r['Email'] for r in recipients]))
            messagebox.showinfo("Success", f"Loaded {len(recipients)} recipients")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import recipients: {e}")

# Function to preview the email
def preview_email():
    preview_window = tk.Toplevel(root)
    preview_window.title("Email Preview")
    
    ttk.Label(preview_window, text="Subject:").grid(row=0, column=0, sticky=tk.W)
    ttk.Label(preview_window, text=subject_var.get()).grid(row=0, column=1, sticky=tk.W)
    
    ttk.Label(preview_window, text="Body:").grid(row=1, column=0, sticky=tk.W)
    preview_body = tk.Text(preview_window, width=50, height=10)
    preview_body.grid(row=1, column=1, pady=5)
    preview_body.insert(tk.END, body_text.get("1.0", tk.END))
    preview_body.config(state=tk.DISABLED)

# Function to generate a simple CAPTCHA
def generate_captcha():
    letters = string.ascii_uppercase + string.digits
    captcha_text = ''.join(random.choice(letters) for _ in range(6))
    return captcha_text

# Function to validate CAPTCHA
def validate_captcha():
    if captcha_var.get() == captcha_text_var.get():
        messagebox.showinfo("CAPTCHA", "CAPTCHA verified successfully!")
        send_email_button.config(state=tk.NORMAL)
    else:
        messagebox.showerror("CAPTCHA", "Incorrect CAPTCHA. Please try again.")
        captcha_text_var.set(generate_captcha())
        send_email_button.config(state=tk.DISABLED)

# Function to schedule an email
def schedule_email():
    scheduled_time = schedule_time_var.get()
    try:
        # Parse the scheduled time
        scheduled_hour, scheduled_minute = map(int, scheduled_time.split(":"))
        now = datetime.now()
        
        # Calculate the delay until the scheduled time
        scheduled_datetime = now.replace(hour=scheduled_hour, minute=scheduled_minute, second=0, microsecond=0)
        if scheduled_datetime < now:
            scheduled_datetime += timedelta(days=1)  # Schedule for the next day if the time has already passed today
        
        delay = (scheduled_datetime - now).total_seconds()
        
        # Schedule the email by starting a new thread that waits for the delay and then sends the email
        threading.Timer(delay, send_email).start()
        
        messagebox.showinfo("Scheduled", f"Email scheduled to be sent at {scheduled_time}.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to schedule email: {e}")

# Create the main window
root = tk.Tk()
root.title("Email Automation Program")
root.geometry("600x800")

# Create a frame for the form
frame = ttk.Frame(root, padding="10")
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Email Subject
ttk.Label(frame, text="Subject:").grid(row=0, column=0, sticky=tk.W)
subject_var = tk.StringVar()
subject_entry = ttk.Entry(frame, width=50, textvariable=subject_var)
subject_entry.grid(row=0, column=1, pady=5)

# Email Body
ttk.Label(frame, text="Email Body:").grid(row=1, column=0, sticky=tk.W)
body_text = tk.Text(frame, width=50, height=10)
body_text.grid(row=1, column=1, pady=5)

# Attachment
ttk.Label(frame, text="Attachment:").grid(row=2, column=0, sticky=tk.W)
attachment_var = tk.StringVar()
attachment_entry = ttk.Entry(frame, width=40, textvariable=attachment_var)
attachment_entry.grid(row=2, column=1, pady=5, sticky=tk.W)
ttk.Button(frame, text="Browse", command=attach_file).grid(row=2, column=2, pady=5)

# Recipient List
ttk.Label(frame, text="Recipient List:").grid(row=3, column=0, sticky=tk.W)
recipient_list_var = tk.StringVar(value="No file loaded")
recipient_entry = ttk.Entry(frame, width=40, textvariable=recipient_list_var)
recipient_entry.grid(row=3, column=1, pady=5, sticky=tk.W)
ttk.Button(frame, text="Import CSV/Excel", command=import_recipients).grid(row=3, column=2, pady=5)

# SMTP Configuration Labels and Entries
ttk.Label(frame, text="SMTP Server:").grid(row=4, column=0, sticky=tk.W)
smtp_server_var = tk.StringVar(value="smtp.gmail.com")
smtp_server_entry = ttk.Entry(frame, width=50, textvariable=smtp_server_var)
smtp_server_entry.grid(row=4, column=1, pady=5)

ttk.Label(frame, text="Port:").grid(row=5, column=0, sticky=tk.W)
port_var = tk.StringVar(value="587")
port_entry = ttk.Entry(frame, width=50, textvariable=port_var)
port_entry.grid(row=5, column=1, pady=5)

ttk.Label(frame, text="Email Address:").grid(row=6, column=0, sticky=tk.W)
email_var = tk.StringVar()
email_entry = ttk.Entry(frame, width=50, textvariable=email_var)
email_entry.grid(row=6, column=1, pady=5)

ttk.Label(frame, text="Password:").grid(row=7, column=0, sticky=tk.W)
password_var = tk.StringVar()
password_entry = ttk.Entry(frame, width=50, textvariable=password_var, show="*")
password_entry.grid(row=7, column=1, pady=5)

# CAPTCHA
ttk.Label(frame, text="CAPTCHA:").grid(row=8, column=0, sticky=tk.W)
captcha_text_var = tk.StringVar(value=generate_captcha())
captcha_label = ttk.Label(frame, textvariable=captcha_text_var, font=("Arial", 12, "bold"))
captcha_label.grid(row=8, column=1, pady=5, sticky=tk.W)

captcha_var = tk.StringVar()
captcha_entry = ttk.Entry(frame, width=10, textvariable=captcha_var)
captcha_entry.grid(row=8, column=1, pady=5, sticky=tk.E)
ttk.Button(frame, text="Verify", command=validate_captcha).grid(row=8, column=2, pady=5)

# Schedule Email
ttk.Label(frame, text="Schedule Time (HH:MM):").grid(row=9, column=0, sticky=tk.W)
schedule_time_var = tk.StringVar()
schedule_time_entry = ttk.Entry(frame, width=50, textvariable=schedule_time_var)
schedule_time_entry.grid(row=9, column=1, pady=5)
ttk.Button(frame, text="Schedule Email", command=schedule_email).grid(row=9, column=2, pady=5)

# Buttons for Preview, Send, and Status Display
ttk.Button(frame, text="Preview Email", command=preview_email).grid(row=10, column=0, pady=10)
send_email_button = ttk.Button(frame, text="Send Email", command=send_email)
send_email_button.grid(row=10, column=1, pady=10)
send_email_button.config(state=tk.DISABLED)

# Status Display
ttk.Label(frame, text="Status:").grid(row=11, column=0, sticky=tk.W)
status_display = tk.Text(frame, width=60, height=15, state=tk.DISABLED)
status_display.grid(row=11, column=1, pady=10)

root.mainloop()