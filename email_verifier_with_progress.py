
import tkinter as tk
from tkinter import filedialog, messagebox
import re
import dns.resolver
import smtplib
import csv
import os
import openpyxl
import threading

def is_valid_syntax(email):
    pattern = r"^[\w\.-]+@[\w\.-]+\.\w+$"
    return re.match(pattern, email) is not None

def has_mx_record(domain):
    try:
        records = dns.resolver.resolve(domain, 'MX')
        return len(records) > 0
    except:
        return False

def smtp_check(email):
    domain = email.split('@')[1]
    try:
        mx_records = dns.resolver.resolve(domain, 'MX')
        mx_record = str(mx_records[0].exchange)
        server = smtplib.SMTP(timeout=10)
        server.connect(mx_record)
        server.helo('example.com')
        server.mail('test@example.com')
        code, message = server.rcpt(email)
        server.quit()
        return code == 250 or code == 251
    except:
        return False

def read_emails(file_path):
    emails = []
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext == '.csv':
            with open(file_path, newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                for row in reader:
                    if row:
                        emails.append(row[0].strip())
        elif ext == '.txt':
            with open(file_path, 'r', encoding='utf-8') as f:
                for line in f:
                    email = line.strip()
                    if email:
                        emails.append(email)
        elif ext == '.xlsx':
            wb = openpyxl.load_workbook(file_path)
            ws = wb.active
            for row in ws.iter_rows(min_row=1, max_col=1, values_only=True):
                email = row[0]
                if email:
                    emails.append(str(email).strip())
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read file: {e}")
    return emails

def verify_emails(file_path, status_label):
    results = []
    emails = read_emails(file_path)
    total = len(emails)
    for idx, email in enumerate(emails, start=1):
        status_label.config(text=f"Verifying {idx} of {total}...")
        root.update_idletasks()
        status = "Invalid Format"
        if is_valid_syntax(email):
            domain = email.split('@')[1]
            if has_mx_record(domain):
                if smtp_check(email):
                    status = "Valid"
                else:
                    status = "SMTP Failed"
            else:
                status = "No MX Record"
        results.append((email, status))
    return results

def save_results(results, out_file):
    with open(out_file, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(["Email", "Status"])
        for row in results:
            writer.writerow(row)

def start_verification_thread():
    threading.Thread(target=start_verification).start()

def start_verification():
    file_path = filedialog.askopenfilename(filetypes=[("All Supported", "*.csv *.xlsx *.txt"),
                                                       ("CSV files", "*.csv"),
                                                       ("Excel files", "*.xlsx"),
                                                       ("Text files", "*.txt")])
    if not file_path:
        return
    status_label.config(text="Starting verification...")
    root.update_idletasks()
    results = verify_emails(file_path, status_label)
    out_file = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if out_file:
        save_results(results, out_file)
        messagebox.showinfo("Done", f"Verification complete. Results saved to: {out_file}")
        status_label.config(text="Verification complete.")
    else:
        status_label.config(text="Verification cancelled.")

root = tk.Tk()
root.title("Email Verifier")
root.geometry("350x180")

label = tk.Label(root, text="Email Verifier with SMTP Check", font=("Arial", 12))
label.pack(pady=10)

btn = tk.Button(root, text="Start Verification", command=start_verification_thread)
btn.pack(pady=5)

status_label = tk.Label(root, text="", fg="blue")
status_label.pack(pady=10)

root.mainloop()
