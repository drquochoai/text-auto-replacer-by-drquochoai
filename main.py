#!filepath: main.py
# changed main.py: No functional changes, just added file path for clarity.
# Updated imports: Removed unused imports.
import atexit
from src.starterConfig import *
from src.languages import *
from src.processXlsx import *
# from src.keyboardEvent import *
from src.keyboardEvent import start_keyboard_listener, keyboard, stop_keyboard_hook
from src.tkGUI import *
import win32clipboard
import time
from threading import Thread

# Define constants manually
CF_UNICODETEXT = 13
CF_TEXT = 1

def start_program():
    """
    Restart the program by reloading all settings and replacement data.
    """
    replacement_data = load_settings_and_data()
    if replacement_data:
        start_keyboard_listener(replacement_data)
        clipboard_thread = Thread(
            target=monitor_clipboard, args=(replacement_data,), daemon=True
        )
        clipboard_thread.start()
        logging.info("Clipboard monitoring started.")
        try:
            download_and_process_xlsx_for_languages(
                LINK_LANGUAGE_FILE, LANGUAGES_FOLDER
            )
            load_language(saved_language)
            update_gui_language()
        except Exception as e:
            logging.error(f"Error downloading or processing language XLSX file: {e}")
         # Update GUI fields after loading data
        update_replacement_fields()
        update_link_edit_file_field()
    else:
        logging.error("Failed to start program. Exiting program.")

def is_clipboard_text():
    try:
        win32clipboard.OpenClipboard()
        # Check for CF_UNICODETEXT first for Unicode text
        if win32clipboard.IsClipboardFormatAvailable(CF_UNICODETEXT):
            data = win32clipboard.GetClipboardData(CF_UNICODETEXT)
            return True, data
        # Fallback to CF_TEXT for ANSI text
        elif win32clipboard.IsClipboardFormatAvailable(CF_TEXT):
            data = win32clipboard.GetClipboardData(CF_TEXT)
            return True, data
        else:
            return False, None
    except Exception as e:
        print(f"An error occurred: {e}")
        return False, None
    finally:
        win32clipboard.CloseClipboard()

def monitor_clipboard(replacement_data):
    while True:
        try:
            # must check is_clipboard_text, because if not, it will raise TypeError and clipboard will destroy in windows. copy files cannot work
            is_text, content = is_clipboard_text()
            if is_text is False:
                # print(is_text, content)
                time.sleep(TIME_INTERVAL_CLIPBOARD_CHECK)
                continue
            keys_of_replacement_data = list(replacement_data.keys())
            if content is not None:
                content = content.strip()
            if content in keys_of_replacement_data:
                print(is_text, content)
                new_content = replacement_data[content]
                logging.info(f"Clipboard content: {new_content}")
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardText(
                    new_content, win32clipboard.CF_UNICODETEXT
                )
                win32clipboard.CloseClipboard()
                keyboard.press_and_release("ctrl+v")
        except TypeError:
            # Handle the case where clipboard content is not text
            pass
        except win32clipboard.error as e: # Catch specific clipboard errors
            logging.error(f"Clipboard error: {e}")
        except Exception as e: # Keep a general exception for unexpected issues, but log it
            logging.exception("An unexpected error occurred during clipboard monitoring:") # Use logging.exception for full traceback

        time.sleep(TIME_INTERVAL_CLIPBOARD_CHECK)


if __name__ == "__main__":
    start_program()
    atexit.register(stop_keyboard_hook)
    # Run the GUI
    root.mainloop()