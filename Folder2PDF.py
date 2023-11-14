import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from docx2pdf import convert
import threading
import queue
import sys
import re
import time
import psutil

class StderrRedirector:
    def __init__(self, queue):
        self.queue = queue

    def write(self, string):
        # Extract relevant information from the string
        match = re.search(r"(\d+/\d+)", string)
        if match:
            self.queue.put("Conversion in progress: " + match.group(1))  # Document count like "1/45"
        else:
            # Check for estimated time and other relevant info
            time_match = re.search(r"(\d+:\d+:\d+|\d+:\d+)", string)
            if time_match:
                self.queue.put("Estimated time: " + time_match.group(1))  # Estimated time

    def flush(self):
        pass  # Required for file-like object

def set_custom_style(root):
    style = ttk.Style(root)
    style.theme_use('clam')  # Using a theme for a better look

    # Customize the label style
    style.configure('TLabel', background='white', font=('Helvetica', 12))

    # Customize the Toplevel window
    style.configure('TFrame', background='white')

    # Customize the Button
    style.configure('TButton', font=('Helvetica', 10), borderwidth=1)
    style.map('TButton', background=[('active', '!disabled', '#c0c0c0'), ('pressed', '#a0a0a0')])


def is_winword_running():
    """Check if WINWORD.exe is currently running."""
    for process in psutil.process_iter(['name']):
        if process.info['name'] == 'WINWORD.EXE':
            return True
    return False

def terminate_winword():
    """Terminate all instances of WINWORD.exe."""
    for process in psutil.process_iter(['name']):
        if process.info['name'] == 'WINWORD.EXE':
            process.terminate()  # Terminate the process

def convert_documents(folder_path, message_queue, conversion_done_event):
    sys.stderr = StderrRedirector(message_queue)

    if folder_path:
        convert(folder_path)
    conversion_done_event.set()  # Signal that conversion is done

def update_status(info_label, message_queue, conversion_done_event):
    try:
        message = message_queue.get_nowait()
        info_label.config(text=message)
    except queue.Empty:
        if conversion_done_event.is_set():
            perform_termination_countdown(info_label)
            return

    root.after(100, update_status, info_label, message_queue, conversion_done_event)

def perform_termination_countdown(info_label):
    for i in range(3, 0, -1):
        info_label.config(text=f"Conversion Complete\nTerminating in {i}...")
        root.update()
        time.sleep(1)
    root.quit()  # End tkinter main loop

def select_folder_and_convert(root, message_queue):
    if is_winword_running():
        response = messagebox.askyesno("Terminate Word", "Microsoft Word is currently running. Would you like to close it?")
        if response:
            terminate_winword()
            messagebox.showinfo("Info", "Microsoft Word has been closed.")
        else:
            root.destroy()
            return

    folder_path = filedialog.askdirectory()

    if folder_path:
        global conversion_window
        conversion_window = tk.Toplevel(root)
        conversion_window.title("Conversion Status")
        conversion_window.geometry("300x100")
        conversion_window.resizable(False, False)

        # Use ttk.Frame for improved styling
        frame = ttk.Frame(conversion_window, style='TFrame')
        frame.pack(expand=True, fill=tk.BOTH)

        global info_label
        info_label = ttk.Label(frame, text="Starting conversion...", style='TLabel')
        info_label.pack(expand=True)

        conversion_done_event = threading.Event()
        threading.Thread(target=convert_documents, args=(folder_path, message_queue, conversion_done_event)).start()
        update_status(info_label, message_queue, conversion_done_event)
    else:
        root.destroy()

# Main window setup
root = tk.Tk()
set_custom_style(root)  # Apply custom styles
root.withdraw()  # Hide main window
message_queue = queue.Queue()

select_folder_and_convert(root, message_queue)
root.mainloop()