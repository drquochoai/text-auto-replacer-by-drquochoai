#!filepath: src/tkGUI.py
# Updated imports: Removed unused imports.
import requests
import logging
import tkinter as tk
from tkinter import messagebox
import os
import sys
import webbrowser

from src.processXlsx import *
from src.updater import check_for_new_update  # Import check_for_new_update function

def save_settings():
    """
    Save settings to the .ini file if the URL has changed.
    """
    global previous_sheet_url, saved_language

    # Get current values from the widgets
    current_sheet_url = sheet_url_text.get("1.0", "end-1c").strip()

    # Check if the URL has changed
    if current_sheet_url != previous_sheet_url:
        # Update the previous URL value
        previous_sheet_url = current_sheet_url

        # Save the URL to the .ini file
        config["Settings"] = {
            "sheet_url": current_sheet_url,
            "language": saved_language,
        }
        with open("settings.ini", "w") as configfile:
            config.write(configfile)
        logging.info("Settings saved successfully!")

def load_settings_and_data():
    """
    Loads settings from the .ini file and reloads the replacement data.
    """
    global SHEET_URL, BEFORE_REPLACEMENT, AFTER_REPLACEMENT, current_replacement_data, LINK_EDIT_FILE

    # Load settings from the .ini file
    config.read("settings.ini")
    SHEET_URL = config.get("Settings", "sheet_url", fallback=DEFAULT_SHEET_URL)

    # Reload the replacement data from the XLSX URL
    replacement_data = load_replacement_data(SHEET_URL)
    if replacement_data:
        current_replacement_data = replacement_data  # Update the current replacement data
        logging.info("Replacement data reloaded successfully.")

        # Update the Tkinter entry fields with the dynamic values
        before_replacement_entry.config(state="normal")  # Temporarily enable to update
        before_replacement_entry.delete(0, tk.END)  # Clear the entry field
        before_replacement_entry.insert(0, BEFORE_REPLACEMENT)  # Insert the dynamic value
        before_replacement_entry.config(state="readonly")  # Disable editing again

        after_replacement_entry.config(state="normal")  # Temporarily enable to update
        after_replacement_entry.delete(0, tk.END)  # Clear the entry field
        after_replacement_entry.insert(0, AFTER_REPLACEMENT)  # Insert the dynamic value
        after_replacement_entry.config(state="readonly")  # Disable editing again

        update_link_edit_file_field()  # Update the "Link Edit File" field
        return replacement_data
    else:
        logging.error("Failed to reload replacement data.")
        return None

def toggle_pause():
    """
    Toggle the pause state.
    """
    global is_paused
    is_paused = not is_paused  # Toggle the pause state
    if is_paused:
        pause_button.config(
            text=language_config["Buttons"]["resume"],
            bg="red",
            activebackground="darkred",
        )
        logging.info("Program paused.")
    else:
        pause_button.config(
            text=language_config["Buttons"]["pause"],
            bg="green",
            activebackground="darkgreen",
        )
        logging.info("Program resumed.")

def exit_to_reload_data():
    """
    # Reload the XLSX file from the internet and update the replacement data.
    # """
    # global SHEET_URL, current_replacement_data
    # replacement_data = load_replacement_data(current_value)
    # if replacement_data:
    #     current_replacement_data = replacement_data  # Update the current replacement data
    #     logging.info("XLSX file reloaded from the internet.")
    #     messagebox.showinfo("Reload", "XLSX file reloaded successfully!")
    # else:
    #     logging.error("Failed to reload XLSX file from the internet.")
    #     messagebox.showerror("Reload", "Failed to reload XLSX file from the internet.")

    # After showing the reload message, exit and restart the program
    # Close the application
    root.quit()
    root.destroy()
    sys.exit()

def update_gui_language():
    """
    Update the GUI labels, buttons, and messages based on the selected language.
    """
    # Update Labels
    sheet_url_label.config(text=language_config["Labels"]["sheet_url"])
    before_replacement_label.config(text=language_config["Labels"]["before_replacement"])
    after_replacement_label.config(text=language_config["Labels"]["after_replacement"])

    # Update Buttons
    save_button.config(text=language_config["Buttons"]["save_settings"])
    pause_button.config(text=language_config["Buttons"]["pause"])
    reload_button.config(text=language_config["Buttons"]["reload_csv"])

def change_language(language_code):
    """
    Change the language of the program and save the selected language to the settings file.
    """
    global language_config, saved_language
    saved_language = language_code
    load_language(language_code)
    update_gui_language()

    # Save the selected language to the settings file
    config["Settings"] = {
        "language": language_code,
        "sheet_url": sheet_url_text.get("1.0", "end-1c"),
    }
    with open("settings.ini", "w") as configfile:
        config.write(configfile)
    logging.info(f"Language changed to {language_code} and saved to settings file.")

# Button to open the Google Sheet URL
def open_google_sheet():
    url = link_edit_file_text.get("1.0", "end-1c").strip()
    if url:
        webbrowser.open(url)
    else:
        messagebox.showwarning("No URL", "The 'Link Edit File' URL is not set.")

def update_link_edit_file_field():
    link_edit_file_text.config(state="normal")  # Temporarily enable to update
    link_edit_file_text.delete("1.0", tk.END)  # Clear the entry field
    link_edit_file_text.insert("1.0", LINK_EDIT_FILE)  # Insert the dynamic value
    link_edit_file_text.config(state="disabled")  # Disable editing again

# Create the GUI
root = tk.Tk()
root.title("Text Replacer by drquochoai")
root.geometry("800x400")  # Set a fixed size for the window
root.configure(bg="#f5f5f5")  # Light background color

# Custom font
custom_font = ("Roboto", 10)

# Labels
sheet_url_label = tk.Label(
    root, text="Sheet URL:", font=custom_font, bg="#f5f5f5", fg="#333"
)
sheet_url_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")

before_replacement_label = tk.Label(
    root, text="Before Replacement:", font=custom_font, bg="#f5f5f5", fg="#333"
)
before_replacement_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")

after_replacement_label = tk.Label(
    root, text="After Replacement:", font=custom_font, bg="#f5f5f5", fg="#333"
)
after_replacement_label.grid(row=2, column=0, padx=10, pady=10, sticky="w")

# Entry fields
# Use a Text widget for the sheet URL
sheet_url_text = tk.Text(
    root, width=50, height=3, font=custom_font, bg="#fff", fg="#333", relief="flat", bd=2
)
sheet_url_text.insert("1.0", SHEET_URL)  # Insert the default URL
sheet_url_text.grid(row=0, column=1, padx=10, pady=10)

# Disable editing for BEFORE_REPLACEMENT and AFTER_REPLACEMENT fields
before_replacement_entry = tk.Entry(
    root,
    width=50,
    font=custom_font,
    bg="#fff",
    fg="#333",
    relief="flat",
    bd=2,
    state="readonly",
)
before_replacement_entry.insert(0, BEFORE_REPLACEMENT)
before_replacement_entry.grid(row=1, column=1, padx=10, pady=10)

after_replacement_entry = tk.Entry(
    root,
    width=50,
    font=custom_font,
    bg="#fff",
    fg="#333",
    relief="flat",
    bd=2,
    state="readonly",
)
after_replacement_entry.insert(0, AFTER_REPLACEMENT)
after_replacement_entry.grid(row=2, column=1, padx=10, pady=10)

# Buttons
save_button = tk.Button(
    root,
    text="Save Settings",
    font=custom_font,
    bg="#0078d7",
    fg="#fff",
    relief="flat",
    activebackground="#005a9e",
    activeforeground="#fff",
    command=save_settings,
)
save_button.grid(row=3, column=0, padx=10, pady=10, sticky="ew")

pause_button = tk.Button(
    root,
    text="Pause",
    font=custom_font,
    bg="green",
    fg="#fff",
    relief="flat",
    activebackground="#005a9e",
    activeforeground="#fff",
    command=toggle_pause,
)
pause_button.grid(row=3, column=1, padx=10, pady=10, sticky="ew")

reload_button = tk.Button(
    root,
    text="Exit to Reload data",
    font=custom_font,
    bg="#0078d7",
    fg="#fff",
    relief="flat",
    activebackground="#005a9e",
    activeforeground="#fff",
    command=lambda: exit_to_reload_data(),
)
reload_button.grid(row=4, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

# Language dropdown
saved_language = config.get("Settings", "language", fallback="vi")
# Load all .ini files in the language folder
# Check if the languages folder exists, if not, create it and download the XLSX file
if not os.path.exists(LANGUAGES_FOLDER):
    os.makedirs(LANGUAGES_FOLDER)
    download_and_process_xlsx_for_languages(LINK_EDIT_FILE, LANGUAGES_FOLDER)

load_language(saved_language)
language_files = [f for f in os.listdir(LANGUAGES_FOLDER) if f.endswith(".ini")]

language_codes = [os.path.splitext(f)[0] for f in language_files]  # Extract language codes

# Update the language dropdown with all available languages
language_var = tk.StringVar(root)
language_var.set(saved_language)  # Default language
language_menu = tk.OptionMenu(root, language_var, *language_codes, command=change_language)
language_menu.config(
    font=custom_font,
    bg="#0078d7",
    fg="#fff",
    relief="flat",
    activebackground="#005a9e",
    activeforeground="#fff",
)
language_menu.grid(row=4, column=2, padx=10, pady=10, sticky="ew")

# New text field for the "Link Edit File"
link_edit_file_label = tk.Label(
    root, text="Link Edit File:", font=custom_font, bg="#f5f5f5", fg="#333"
)
link_edit_file_label.grid(row=5, column=0, padx=10, pady=10, sticky="w")

link_edit_file_text = tk.Text(
    root,
    width=50,
    height=1,
    font=custom_font,
    bg="darkred",
    fg="white",
    relief="flat",
    bd=2,
    state="disabled",
)
link_edit_file_text.grid(row=5, column=1, padx=10, pady=10)

open_sheet_button = tk.Button(
    root,
    text="Open Google Sheet",
    font=custom_font,
    bg="darkred",
    fg="white",
    relief="flat",
    activebackground="#005a9e",
    activeforeground="#fff",
    command=open_google_sheet,
)
open_sheet_button.grid(row=5, column=2, padx=10, pady=10, sticky="ew")

# Update label to show update status
update_label = tk.Label(root, text="", font=custom_font, bg="#f5f5f5", fg="#333")
update_label.grid(row=6, column=0, columnspan=2, padx=10, pady=10, sticky="w")

# Download button (initially hidden)
download_button = tk.Button(
    root,
    text="Download Update",
    font=custom_font,
    bg="#0078d7",
    fg="#fff",
    relief="flat",
    activebackground="#005a9e",
    activeforeground="#fff",
)
download_button.grid(row=6, column=2, padx=10, pady=10)
# download_button.pack_forget()
download_button.config(state=tk.DISABLED)  # Disable the button initially
# Call the function to update the GUI with the default language
change_language(saved_language)
update_gui_language()

# Bind save_settings() to the "<FocusOut>" event for the entry fields
sheet_url_text.bind("<FocusOut>", lambda event: save_settings())
# Status bar
status_bar = tk.Label(
    root,
    text="Ready",
    bd=1,
    relief=tk.SUNKEN,
    anchor=tk.W,
    font=custom_font,
    bg="#f5f5f5",
    fg="#333",
)
status_bar.grid(row=7, column=0, columnspan=3, sticky="ew", padx=10, pady=10)

# Redirect logging to the status bar
class StatusBarHandler(logging.Handler):
    def __init__(self, status_bar):
        logging.Handler.__init__(self)
        self.status_bar = status_bar

    def emit(self, record):
        log_entry = self.format(record)
        self.status_bar.after(
            0, self.status_bar.config, {"text": log_entry}
        )  # Schedule an update to GUI

# Redirect logging to the status bar
status_bar_handler = StatusBarHandler(status_bar)
status_bar_handler.setFormatter(
    logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
)
logging.getLogger().addHandler(status_bar_handler)

def load_replacement_data(xlsx_url):
    """
    Load replacement data from an XLSX file hosted online or from a local backup.
    """
    try:
        # Try to download the XLSX from the internet
        replacement_data = load_xlsx_from_url(xlsx_url)
        if replacement_data:
            # Save the downloaded XLSX to a local backup file
            save_xlsx_to_file(replacement_data, BACKUP_XLSX_PATH)
            logging.info(
                "Replacement data loaded from the internet and saved to local backup."
            )
        return replacement_data

    except requests.exceptions.RequestException as e:
        logging.error(f"Error downloading data: {e}")
        logging.info("Attempting to load replacement data from local backup...")

        # If the download fails, try loading from the local backup file
        if os.path.exists(BACKUP_XLSX_PATH):
            return load_xlsx_from_file(BACKUP_XLSX_PATH)
        else:
            logging.error("No local backup file found. Exiting program.")
            return {}

# Call check_for_new_update when the GUI is initialized
check_for_new_update(update_label, download_button)