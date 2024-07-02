import os
import re
import csv
from email import policy
from email.parser import BytesParser
from bs4 import BeautifulSoup
from tkinter import Tk, Button, Label, filedialog, messagebox, DoubleVar
from tkinter import ttk
from threading import Thread
import extract_msg

def extract_specific_url(full_url):
    pattern = re.compile(r'https%3A%2F%2Fwww\.amazon\.co\.uk%2Fg%2F([A-Z0-9]+)')
    match = pattern.search(full_url)
    if match:
        extracted_id = match.group(1)
        return f'https://www.amazon.co.uk/g/{extracted_id}'
    return None

def extract_links_from_eml(file_path):
    with open(file_path, 'rb') as f:
        msg = BytesParser(policy=policy.default).parse(f)

    unique_links = set()
    for part in msg.iter_parts():
        if part.get_content_type() == 'text/html':
            charset = part.get_content_charset()
            html_content = part.get_payload(decode=True).decode(charset)

            soup = BeautifulSoup(html_content, 'lxml')
            for a in soup.find_all('a', href=True):
                href = a['href']
                extracted_url = extract_specific_url(href)
                if extracted_url:
                    unique_links.add(extracted_url)
    return list(unique_links)

def extract_links_from_msg(file_path):
    msg = extract_msg.Message(file_path)
    html_body = msg.htmlBody
    unique_links = set()

    if html_body:
        soup = BeautifulSoup(html_body, 'lxml')
        for a in soup.find_all('a', href=True):
            href = a['href']
            extracted_url = extract_specific_url(href)
            if extracted_url:
                unique_links.add(extracted_url)
    return list(unique_links)

def process_email_files(directory, progress_var):
    all_links = set()
    files = []
    for root, _, filenames in os.walk(directory):
        for file in filenames:
            if file.endswith('.eml') or file.endswith('.msg'):
                files.append(os.path.join(root, file))
    
    total_files = len(files)
    for i, file_path in enumerate(files):
        if file_path.endswith('.eml'):
            links = extract_links_from_eml(file_path)
        elif file_path.endswith('.msg'):
            links = extract_links_from_msg(file_path)
        else:
            continue
        all_links.update(links)
        progress_var.set((i + 1) / total_files * 100)
    return list(all_links)

def save_to_csv(directory, links):
    csv_file_path = os.path.join(directory, 'extracted_links.csv')
    if os.path.exists(csv_file_path):
        overwrite = messagebox.askyesno("Overwrite Confirmation", f"{csv_file_path} already exists. Do you want to overwrite it?", icon='warning', default='no')
        if not overwrite:
            messagebox.showinfo("Cancelled", "The operation was cancelled. No file was created.")
            return None

    with open(csv_file_path, 'w', newline='') as csvfile:
        csvwriter = csv.writer(csvfile)
        for link in links:
            csvwriter.writerow([link])
    return csv_file_path

def on_select_folder():
    global directory, links, progress_var
    directory = filedialog.askdirectory(title="Select Directory with .eml or .msg Files")
    if directory:
        select_button.config(state="disabled")
        progress_var.set(0)
        progress_bar.start()
        Thread(target=process_folder, args=(directory,)).start()
    else:
        messagebox.showinfo("No Directory Selected", "No directory selected. Please try again.")

def process_folder(directory):
    global links
    links = process_email_files(directory, progress_var)
    progress_bar.stop()
    if links:
        result = messagebox.askyesno("Links Found", f"Found {len(links)} Amazon voucher links. Do you want to save these links to a CSV file?")
        if result:
            save_to_csv(directory, links)
            messagebox.showinfo("Success", "Links have been saved to a CSV file.")
    else:
        messagebox.showinfo("No Links Found", "No Amazon voucher links found in the selected directory.")
    select_button.config(state="normal")

def main():
    global root, select_button, progress_bar, progress_var, directory, links
    directory = None
    links = []

    root = Tk()
    root.title("Amazon Voucher Extractor")

    label = Label(root, text="Click the button below to select the folder containing your .eml or .msg files:")
    label.pack(pady=10)

    select_button = Button(root, text="Select Folder", command=on_select_folder)
    select_button.pack(pady=5)

    progress_var = DoubleVar()
    progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate", variable=progress_var)
    progress_bar.pack(pady=10)

    root.mainloop()

if __name__ == '__main__':
    main()
