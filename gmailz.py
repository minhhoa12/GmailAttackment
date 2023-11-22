import imaplib
import email
import os
import tkinter as tk
from tkinter import messagebox

def download_attachments():
    # Load configuration from the GUI input fields
    username = username_entry.get()
    password = password_entry.get()
    sender_emails_input = sender_emails_entry.get()
    sender_emails = [email.strip() for email in sender_emails_input.split(",")]
    path_to_your_directory = directory_path_entry.get().strip()

    # Rest of your existing code...
    # Tao file luu tru
    downloaded_files_file = "downloaded_files.txt"

    # Tao tap hop de luu tru cac ten file da tai xuong
    downloaded_files = set()

    # Nap tap hop tu file luu tru (neu co)
    if os.path.exists(downloaded_files_file):
        with open(downloaded_files_file, "r", encoding="utf-8") as file:
            downloaded_files = set(file.read().splitlines())

    # Ket noi den Gmail
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(username, password)

    # Chon thu muc inbox
    mail.select("inbox")

    # Tao danh sach cac dieu kien tim kiem cho tung nguoi gui
    search_queries = [f'(FROM "{sender}")' for sender in sender_emails]

    # Ket hop ket qua tu cac dieu kien tim kiem
    all_messages = set()
    for search_query in search_queries:
        status, messages = mail.search(None, search_query)
        all_messages.update(messages[0].split())

    # Duyet qua tung email va lay file dinh kem
    for message_id in all_messages:
        # Lay noi dung email
        res, msg = mail.fetch(message_id, "(RFC822)")
        email_msg = email.message_from_bytes(msg[0][1])
        
        # Kiem tra email co file dinh kem khong
        if email_msg.get_content_maintype() != "multipart":
            continue
        
        # Duyet qua tung phan tu trong email
        for part in email_msg.walk():
            if part.get_content_maintype() == "multipart":
                continue
            if part.get("Content-Disposition") is None:
                continue
            
            # Lay ten file dinh kem
            filename = part.get_filename()
            
            # Kiem tra co ten file khong
            if filename:
                # Kiem tra xem ten file da duoc tai truoc do chua
                if filename in downloaded_files:
                    # print("File {} da duoc tai truoc do, bo qua.".format(filename))
                    text_widget.insert(tk.END, "File {} da duoc tai truoc do, bo qua.".format(filename))
                    continue

                # Xac dinh duong dan de luu tru file dinh kem
                save_path = os.path.join(path_to_your_directory, filename)
                
                # Tao thu muc
                os.makedirs(os.path.dirname(save_path), exist_ok=True)
                
                # Ghi file dinh kem vao o cung
                with open(save_path, "wb") as f:
                    f.write(part.get_payload(decode=True))
                    # print("Da tai file dinh kem tu {}: {}".format(email_msg['From'], filename))
                    text_widget.insert(tk.END, "Da tai file dinh kem tu {}: {}".format(email_msg['From'], filename))
                # Them ten file vao tap hop cac ten file da tai xuong
                downloaded_files.add(filename)

    # Dong ket noi
    mail.logout()

    # Luu trang thai da tai xuong vao file
    with open(downloaded_files_file, "w", encoding="utf-8") as file:
        file.write("\n".join(downloaded_files))
    # Update the GUI text area with download status or messages
    text_widget.insert(tk.END, "\nDownloading attachments...")

    # Rest of your existing code...

    # Display a message box when the download is complete
    messagebox.showinfo("Download Complete", "Attachments downloaded successfully.")
# Function to clear the text widget
def clear_text_widget():
    text_widget.delete(1.0, tk.END)

def start_download():
    # Disable the download button while processing to prevent multiple clicks
    download_button.config(state=tk.DISABLED)

    # Clear the text widget
    clear_text_widget()
    
    # Call the function to download attachments
    download_attachments()
    
    # Re-enable the download button after processing completes
    download_button.config(state=tk.NORMAL)

# Create a GUI window
root = tk.Tk()
root.title("Email Attachment Downloader")

# Create and place GUI elements: labels, entry fields, buttons, etc.
tk.Label(root, text="Username:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
username_entry = tk.Entry(root)
username_entry.grid(row=0, column=1, padx=5, pady=5)

tk.Label(root, text="Password:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
password_entry = tk.Entry(root, show="*")
password_entry.grid(row=1, column=1, padx=5, pady=5)

tk.Label(root, text="Email người gửi:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
sender_emails_entry = tk.Entry(root)
sender_emails_entry.grid(row=2, column=1, padx=5, pady=5)

tk.Label(root, text="Thư mục chứa file:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
directory_path_entry = tk.Entry(root)
directory_path_entry.grid(row=3, column=1, padx=5, pady=5)

# Create a button to trigger the download process
download_button = tk.Button(root, text="Download Attachments", command=start_download)
download_button.grid(row=4, columnspan=2, pady=10)

# Create a text area to display status or messages
text_widget = tk.Text(root, height=15, width=50)
text_widget.grid(row=5, columnspan=2, padx=5, pady=5)

# Start the GUI event loop
root.mainloop()
