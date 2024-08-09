import sys
import tkinter
from tkinter import Tk, filedialog, messagebox, ttk

from deepl_pptx_translator import config_handler, files, text

PROCESSING_FOLDER = False
PROGRESS_LABEL = None
PROGRESS_BAR = None

TARGET_CHOSEN_LANG = None

LANG_SELECT_LIST = [
    "DE",
    "EN-GB",
    "EN-US",
    "FR",
    "ES",
    "IT",
]

# pylint: disable=W0603


def show_infobox_at_end():

    # Create a hidden root window
    root = Tk()
    root.attributes("-alpha", 0)  # Make the window transparent
    root.geometry("0x0")  # Set the window size to zero
    root.withdraw()  # Hide the main window

    # Your translation code here

    # Display the message box and make it always on top
    message = f"Fertig!\n{files.TRANSLATED_FILES_COUNT} Datei(en) wurde(n) übersetzt."
    messagebox.showinfo("Übersetzung abgeschlossen", message)

    # Make the message box always on top of other windows
    root.grab_set()

    # Start the Tkinter main loop
    # root.mainloop()
    root.destroy()
    root.mainloop()


def show_noconfig_error():

    # Create a hidden root window
    root = Tk()
    root.attributes("-alpha", 0)  # Make the window transparent
    root.geometry("0x0")  # Set the window size to zero
    root.withdraw()  # Hide the main window

    # Display the message box and make it always on top
    message = f"Ohne {config_handler.CONFIG_FILE_PATH} kann das Programm nicht genutzt werden."
    messagebox.showerror("Keine config.ini gefunden!", message)

    # Make the message box always on top of other windows
    root.grab_set()

    root.destroy()
    sys.exit()


def main_gui():
    # Set the initial count to 0#
    global PROGRESS_BAR
    global PROGRESS_LABEL

    def select_path():
        selected_path = filedialog.askopenfilename(
            initialdir=config_handler.output_path,
            title="Datei auswählen",
            filetypes=[("PPTX and DOCX files", "*.pptx *.docx")],
            defaultextension=".pptx",
        )
        if selected_path:
            entry_var.set(selected_path)
            entry.pack(pady=10)
            ok_button["state"] = "normal"
            letter_count_label.config(
                text=f"Anzahl der Buchstaben: {files.count_characters_in_file(selected_path)}"
            )
            _detected_file_language = text.detect_language_in_filepath(selected_path)
            print(f"Detected file language: {_detected_file_language}")
            file_count_label.config(
                text=f"Anzahl Dateien: {files.count_docx_pptx_files_in_path(selected_path)}"
            )
            language_detected_label.config(
                text=f"Erkannte Sprache: {_detected_file_language}"
            )

    def select_folder():
        folder_path = filedialog.askdirectory(title="Ordner auswählen")
        if folder_path:
            entry_var.set(folder_path)
            entry.pack(pady=10)
            ok_button["state"] = "normal"
            letter_count_label.config(
                text=f"Anzahl Buchstaben: {files.count_characters_in_folder(folder_path)}"
            )
            _detected_file_language = text.detect_language_in_filepath(folder_path)
            print(f"Detected file language: {_detected_file_language}")
            file_count_label.config(
                text=f"Anzahl Dateien: {files.count_docx_pptx_files_in_path(folder_path)}"
            )
            language_detected_label.config(
                text=f"Erkannte Sprache: {_detected_file_language}"
            )

    def on_ok():
        nonlocal user_pressed_ok
        if entry_var.get():  # Check if a folder, PPTX, or DOCX file is selected
            user_pressed_ok = True
            if entry_var.get().endswith(
                (".pptx", ".docx")
            ):  # Check if a file is selected
                process_selected_file(entry_var.get())
            else:  # Assume a folder is selected
                process_selected_folder(entry_var.get())
            root.destroy()

    def process_selected_file(file_path):
        # Call the function for processing a selected file
        print(f"Processing file: {file_path}")

        if files.check_file_type(file_path) == "pptx":
            files.translate_presentation(file_path)

        if files.check_file_type(file_path) == "docx":
            files.translate_document(file_path)

        files.open_output_folder_in_explorer()
        show_infobox_at_end()

    def process_selected_folder(folder_path):
        # Call the function for processing a selected folder
        print(f"Processing folder: {folder_path}")

        files.process_presentation_and_document_files(folder_path)
        files.open_output_folder_in_explorer()
        show_infobox_at_end()

    def toggle_subdirs():
        files.INCLUDE_SUBDIRS = not files.INCLUDE_SUBDIRS
        print(f"files.INCLUDE_SUBDIRS={files.INCLUDE_SUBDIRS}")

    user_pressed_ok = False

    root = Tk()
    root.title("ISM-DeepL-TransLite")
    root.iconbitmap("icon.ico")

    label = ttk.Label(
        root, text="Dateien (*.pptx oder *.docx) zur Übersetzung auswählen"
    )

    entry_var = tkinter.StringVar()

    button_frame = ttk.Frame(root)

    browse_file_button = tkinter.Button(
        button_frame, text="Datei auswählen...", command=select_path
    )

    browse_folder_button = tkinter.Button(
        button_frame, text="Ordner auswählen...", command=select_folder
    )

    use_subdirs_checkbox = tkinter.Checkbutton(
        root, text="Unterordner auch durchsuchen", command=toggle_subdirs
    )
    use_subdirs_checkbox.select()

    letter_count_label = ttk.Label(root, text="Kein Ordner oder Datei ausgewählt.")
    file_count_label = ttk.Label(root, text="")
    language_detected_label = ttk.Label(root, text="")

    global TARGET_CHOSEN_LANG

    target_lang_frame = ttk.Frame(root)

    target_lang_select_label = ttk.Label(target_lang_frame, text="Ziel Sprache: ")
    TARGET_CHOSEN_LANG = tkinter.StringVar(root)
    TARGET_CHOSEN_LANG.set(LANG_SELECT_LIST[1])

    target_lang_select_menu = tkinter.OptionMenu(
        target_lang_frame, TARGET_CHOSEN_LANG, *LANG_SELECT_LIST
    )

    progress_frame = tkinter.Frame(root)

    PROGRESS_LABEL = ttk.Label(progress_frame, text="")

    global PROGRESS_BAR
    PROGRESS_BAR = ttk.Progressbar(progress_frame, length=200, mode="determinate")

    entry = tkinter.Entry(root, textvariable=entry_var, state="readonly", width=40)

    ok_button = tkinter.Button(root, text="Übersetzen", command=on_ok, state="disabled")

    # VERTICAL ALIGNMENT
    label.pack(side="top", pady=5)

    button_frame.pack(side="top")
    browse_file_button.pack(side="left", pady=3)
    browse_folder_button.pack(side="left", pady=3)

    target_lang_frame.pack(side="top")
    target_lang_select_label.pack(side="left", pady=3)
    target_lang_select_menu.pack(side="left", pady=3)

    use_subdirs_checkbox.pack(side="top", pady=3)
    file_count_label.pack(side="top")
    letter_count_label.pack(side="top")
    language_detected_label.pack(side="top")

    progress_frame.pack(side="top", pady=5)
    PROGRESS_LABEL.pack(side="top", pady=5)
    PROGRESS_BAR.pack(side="top", pady=5)
    entry.pack(side="top", pady=5)
    ok_button.pack(side="top", pady=5)

    # RESIZABLE
    root.resizable(False, False)

    root.protocol(
        "WM_DELETE_WINDOW", sys.exit
    )  # Stop the script if the window is closed

    root.mainloop()
