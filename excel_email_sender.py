import win32com.client
from tkinter import Tk, filedialog, simpledialog, messagebox
import os
import csv
from datetime import datetime

root = Tk()
root.withdraw()

try:
    
    report_path = filedialog.askopenfilename(title="Select Excel Report", filetypes=[("Excel Files", "*.xlsx *.xls")])

    if not report_path:
        raise Exception("No report selected")


    recipient = simpledialog.askstring("Recipient", "Enter recipient email \n(Seperate mutliple emails with;)")

    if not recipient:
        raise Exception("recipient email is required")

    subject = simpledialog.askstring("Subject", "Enter Email Subject")

    body = simpledialog.askstring("Message", "Enter Email Message")


    outlook = win32com.client.Dispatch("outlook.Application")

    mail = outlook.CreateItem(0)

    mail.To = recipient

    mail.Subject = subject

    mail.Body = body

    mail.Attachments.Add(report_path)

    mail.Send()

    log_file = "email_log.csv" 

    file_exists = os.path.exists(log_file)

    with open(log_file, "a", newline="", encoding="utf-8") as file:

        writer = csv.writer(file)

        if not file_exists:
            writer.writerow(["Date", "Recipient", "Report", "Status"])

        writer.writerow([
            datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
            recipient, 
            os.path.basename(report_path),
            "sent"
        ])


        messagebox.showinfo("Success", "Report sent successfully.")

except Exception as e:
    messagebox.showerror("Error", str(e))

    
