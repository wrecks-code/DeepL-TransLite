import sys
import tkinter
from tkinter import Tk, filedialog, messagebox, ttk

from deepl_pptx_translator import configHandler, fileHandler

processing_folder = False
progress_label = None
progress_bar = None


def show_infobox_at_end():

    # Create a hidden root window
    root = Tk()
    root.attributes("-alpha", 0)  # Make the window transparent
    root.geometry("0x0")  # Set the window size to zero
    root.withdraw()  # Hide the main window

    # Your translation code here

    # Display the message box and make it always on top
    message = f"Fertig!\n{fileHandler.translated_files_count} PowerPoint Datei(en) wurde(n) übersetzt."
    messagebox.showinfo("Übersetzung abgeschlossen", message)

    # Make the message box always on top of other windows
    root.grab_set()

    # Start the Tkinter main loop
    # root.mainloop()
    root.destroy()


def show_noconfig_error():
    import deepl_pptx_translator.configHandler as configHandler

    # Create a hidden root window
    root = Tk()
    root.attributes("-alpha", 0)  # Make the window transparent
    root.geometry("0x0")  # Set the window size to zero
    root.withdraw()  # Hide the main window

    # Display the message box and make it always on top
    message = (
        f"Ohne {configHandler.config_file_path} kann das Programm nicht genutzt werden."
    )
    messagebox.showerror("Keine config.ini gefunden!", message)

    # Make the message box always on top of other windows
    root.grab_set()

    root.destroy()
    sys.exit()


def mainGUI():
    from deepl_pptx_translator import fileHandler

    global translated_files_count  # Initialize the global variable
    # Set the initial count to 0#
    global progress_bar
    global progress_label

    def select_path():
        selected_path = filedialog.askopenfilename(
            initialdir=configHandler.output_path,
            title="PowerPoint auswählen",
            filetypes=[("PPTX files", "*.pptx"), ("All files", "*.*")],
            defaultextension=".pptx",
        )
        if selected_path:
            entry_var.set(selected_path)
            entry.pack(pady=10)
            ok_button["state"] = "normal"
            additional_label.config(
                text=f"Anzahl der Buchstaben: {fileHandler.count_characters_in_presentation(selected_path)}"
            )

    def select_folder():
        folder_path = filedialog.askdirectory(title="Ordner mit PowerPoints auswählen")
        if folder_path:
            entry_var.set(folder_path)
            entry.pack(pady=10)
            ok_button["state"] = "normal"
            additional_label.config(
                text=f"Anzahl der Buchstaben: {fileHandler.count_characters_in_folder(folder_path)}"
            )

    def on_ok():
        nonlocal user_pressed_ok
        if entry_var.get():  # Check if a folder or PPTX file is selected
            user_pressed_ok = True
            if entry_var.get().endswith(".pptx"):  # Check if a file is selected
                process_selected_file(entry_var.get())
            else:  # Assume a folder is selected
                process_selected_folder(entry_var.get())
            root.destroy()

    def process_selected_file(file_path):
        import deepl_pptx_translator.apiHandler as apiHandler

        # Call the function for processing a selected file
        print(f"Processing file: {file_path}")
        processing_folder = False
        fileHandler.translate_presentation(file_path)
        fileHandler.open_output_folder_in_explorer()
        show_infobox_at_end()

    def process_selected_folder(folder_path):
        import deepl_pptx_translator.apiHandler as apiHandler

        # Call the function for processing a selected folder
        print(f"Processing folder: {folder_path}")
        processing_folder = True
        fileHandler.process_pptx_files(folder_path)
        fileHandler.open_output_folder_in_explorer()
        show_infobox_at_end()

    def toggle_subdirs():
        global include_subdirectories
        include_subdirectories = not include_subdirectories
        print(f"include_subdirectories={include_subdirectories}")

    user_pressed_ok = False

    root = Tk()
    root.title("DeepL PPTX GUI")

    label = ttk.Label(root, text="Präsentationen (*.pptx) zur Übersetzung auswählen")
    label.pack(pady=10)

    entry_var = tkinter.StringVar()

    browse_file_button = tkinter.Button(
        root, text="Einzelne PowerPoint Datei", command=select_path
    )
    browse_file_button.pack(pady=5, side="top")  # Place the button on top

    browse_folder_button = tkinter.Button(
        root, text="Ordner mit PowerPoint Dateien", command=select_folder
    )
    browse_folder_button.pack(pady=5, side="top")  # Place the button on top

    use_subdirs_checkbox = tkinter.Checkbutton(
        root, text="Unterordner auch durchsuchen", command=toggle_subdirs
    )
    use_subdirs_checkbox.select()
    use_subdirs_checkbox.pack(pady=5)

    additional_label = ttk.Label(root, text="Kein Ordner oder Datei ausgewählt.")
    additional_label.pack(pady=10)

    progress_frame = tkinter.Frame(root)
    progress_frame.pack(pady=10)

    progress_label = ttk.Label(progress_frame, text="")
    progress_label.pack(side="top")

    global progress_bar
    progress_bar = ttk.Progressbar(progress_frame, length=200, mode="determinate")
    progress_bar.pack(side="left")

    entry = tkinter.Entry(root, textvariable=entry_var, state="readonly", width=40)

    ok_button = tkinter.Button(root, text="Übersetzen", command=on_ok, state="disabled")
    ok_button.pack(pady=10)

    # VERTICAL ALIGNMENT
    label.pack(side="top", pady=5)
    browse_file_button.pack(side="top", pady=3)
    browse_folder_button.pack(side="top", pady=3)
    use_subdirs_checkbox.pack(side="top", pady=3)
    additional_label.pack(side="top", pady=5)
    progress_frame.pack(side="top", pady=5)
    entry.pack(side="top", pady=5)
    ok_button.pack(side="top", pady=5)

    # RESIZABLE
    root.resizable(False, False)

    root.protocol(
        "WM_DELETE_WINDOW", sys.exit
    )  # Stop the script if the window is closed

    root.mainloop()
