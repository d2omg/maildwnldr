import tkinter as tk
import base64
import json
from imapclient import IMAPClient
from tkinter import messagebox
from tkinter import ttk
import email
import os
import threading

# Program version
VERSION = "1.9.2"

# Global IMAPClient object
client = None

# Window dimensions
WINDOW_WIDTH = 400
WINDOW_HEIGHT = 300

# Progress window dimensions
PROGRESS_WINDOW_WIDTH = 400
PROGRESS_WINDOW_HEIGHT = 200


def save_config():
    if config_loaded:
        return

    answer = messagebox.askyesno("Save Configuration", "Do you want to save the configuration?")
    if not answer:
        return

    try:
        config = {
            "imap_host": host_entry.get(),
            "imap_user": username_entry.get(),
            "imap_pass": base64.b64encode(password_entry.get().encode()).decode()
        }
        with open("config.json", "w") as file:
            json.dump(config, file)
        print("Configuration successfully saved!")
    except Exception as e:
        print("Error while saving the configuration:", str(e))


def load_config():
    try:
        with open("config.json", "r") as file:
            config = json.load(file)
            host_entry.delete(0, tk.END)
            host_entry.insert(tk.END, config["imap_host"])
            username_entry.delete(0, tk.END)
            username_entry.insert(tk.END, config["imap_user"])
            password_entry.delete(0, tk.END)
            password_entry.insert(tk.END, base64.b64decode(config["imap_pass"]).decode())
            print("Configuration successfully loaded!")
            return True
    except FileNotFoundError:
        print("Configuration file not found.")
    except Exception as e:
        print("Error while loading the configuration:", str(e))
    return False


def connect_to_imap():
    global client  # Declare the global client object

    try:
        imap_host = host_entry.get()
        imap_user = username_entry.get()
        imap_pass = password_entry.get()

        client = IMAPClient(imap_host)
        client.login(imap_user, imap_pass)
        print("Successfully connected to the mail server!")

        # Hide the main window
        window.withdraw()

        # Create a new window for folder selection
        folder_window = tk.Toplevel(window)
        folder_window.title("Select Folders")
        folder_window.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")

        # Get a list of available folders
        folders = [folder[2] for folder in client.list_folders()]
        folder_listbox = tk.Listbox(folder_window, selectmode=tk.MULTIPLE, font=("Arial", 16))
        for folder_name in folders:
            folder_listbox.insert(tk.END, folder_name)

        folder_listbox.pack(fill=tk.BOTH, expand=True)

        def download_selected_folders():
            selected_folders = [folder_listbox.get(index) for index in folder_listbox.curselection()]
            print("Selected folders:", selected_folders)

            # Download emails in EML format
            progress_window = tk.Toplevel(window)
            progress_window.title("Downloading Emails")
            progress_window.geometry(f"{PROGRESS_WINDOW_WIDTH}x{PROGRESS_WINDOW_HEIGHT}")

            progress_label = tk.Label(progress_window, text="", font=("Arial", 16))
            progress_label.pack()

            progress_bar = ttk.Progressbar(progress_window, mode='determinate')
            progress_bar.pack(pady=10)

            def cancel_download():
                progress_window.destroy()

            def download_emails(cancelled):
                total_files = 0
                current_file = 0

                for folder in selected_folders:
                    if cancelled[0]:
                        break

                    # Create a folder to save the emails
                    folder_path = os.path.join(os.getcwd(), folder)
                    os.makedirs(folder_path, exist_ok=True)

                    # Get the list of emails in the folder
                    client.select_folder(folder)
                    messages = client.search()
                    total_files += len(messages)

                    for message_id in messages:
                        if cancelled[0]:
                            break

                        # Get the email in EML format
                        raw_message = client.fetch([message_id], ["RFC822"])[message_id][b"RFC822"]
                        email_message = email.message_from_bytes(raw_message)

                        # Get data to form the file name
                        try:
                            from_who = email_message["From"].split("<")[1].strip().replace('>', '')  # Extract email address without ID
                        except:
                            from_who = 'empty'
                        datee = email_message["Date"].split()[1:4]
                        formatted_date = " ".join(datee)

                        # Form the file name
                        file_name = f"{current_file + 1} {from_who} {formatted_date}.eml"
                        file_path = os.path.join(folder_path, file_name)
                        try:
                            with open(file_path, "wb") as file:
                                file.write(raw_message)
                        except:
                            with open(f"{current_file + 1} {from_who}", "wb") as file:
                                file.write(raw_message)

                        current_file += 1

                        progress_label[
                            "text"] = f"Downloading Folder: {folder}\nDownloaded Files: {current_file}/{total_files}\nRemaining: {total_files - current_file}"
                        progress_bar["value"] = current_file / total_files * 100
                        progress_window.update()

                if cancelled[0]:
                    messagebox.showinfo("Information", "Download canceled")
                else:
                    messagebox.showinfo("Information", f"Download completed\n"
                                                       f"Downloaded Files: {total_files}\n"
                                                       f"Downloaded Folders: {len(selected_folders)}")

                progress_window.destroy()
                folder_window.destroy()
                window.deiconify()

            cancelled = [False]
            threading.Thread(target=download_emails, args=(cancelled,)).start()

            def cancel_download():
                cancelled[0] = True

            cancel_button = tk.Button(progress_window, text="Cancel", command=cancel_download, font=("Arial", 16))
            cancel_button.pack()

        save_button = tk.Button(folder_window, text="Download", command=download_selected_folders, font=("Arial", 16))
        save_button.pack()

        if not config_loaded:
            answer = messagebox.askyesno("Save Configuration", "Do you want to save the configuration?")
            if answer:
                save_config()

    except Exception as e:
        print("Error while connecting to the mail server:", str(e))
        messagebox.showerror("Error", "Failed to connect to the mail server.")
        client = None  # Reset the global client object on error


# Create the main window
window = tk.Tk()
window.title("Mail Server Settings")
window.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")

# Create and place labels and entry fields
host_label = tk.Label(window, text="Server Address:", font=("Arial", 16))
host_label.pack()
host_entry = tk.Entry(window, font=("Arial", 16))
host_entry.pack()

username_label = tk.Label(window, text="Username:", font=("Arial", 16))
username_label.pack()
username_entry = tk.Entry(window, font=("Arial", 16))
username_entry.pack()

password_label = tk.Label(window, text="Password:", font=("Arial", 16))
password_label.pack()
password_entry = tk.Entry(window, show="*", font=("Arial", 16))
password_entry.pack()

# Create the IMAP connection button
connect_button = tk.Button(window, text="Connect", command=connect_to_imap, font=("Arial", 16))
connect_button.pack()

# Load configuration if the file exists
config_loaded = load_config()

# Create the menu
menu_bar = tk.Menu(window)


def about_program():
    messagebox.showinfo("About", f"Author: D12 Inc.\nVersion: {VERSION}")


help_menu = tk.Menu(menu_bar, tearoff=0)
help_menu.add_command(label="About", command=about_program)
menu_bar.add_cascade(label="Help", menu=help_menu)

# Set the menu in the window
window.config(menu=menu_bar)

# Start the main event loop
try:
    window.mainloop()
    if client:
        client.logout()
except Exception as e:
    print("An error occurred:", str(e))
    if client:
        client.logout()
