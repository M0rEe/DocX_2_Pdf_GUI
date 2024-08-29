import os
import threading
from tkinter import Tk, Button, Label, filedialog, messagebox
from tkinter.ttk import Progressbar
from docx2pdf import convert

def convert_all_docx_in_folder(folder_path, progress_bar):
    try:
        # Disable the button while converting
        convert_button.config(state="disabled")
        
        # Get the list of .docx files in the folder
        docx_files = [f for f in os.listdir(folder_path) if f.endswith(".docx")]
        total_files = len(docx_files)

        if total_files == 0:
            messagebox.showinfo("No Files", "No .docx files found in the selected folder.")
            convert_button.config(state="normal")
            return
        
        # Initialize the progress bar
        progress_bar["maximum"] = total_files
        progress_bar["value"] = 0

        for i, docx_file in enumerate(docx_files):
            docx_file_path = os.path.join(folder_path, docx_file)
            convert(docx_file_path)
            progress_bar["value"] += 1
            root.update_idletasks()  # Update the progress bar in the UI

        messagebox.showinfo("Success", "All .docx files have been converted to .pdf successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
    finally:
        # Re-enable the button after completion
        convert_button.config(state="normal")

def start_conversion():
    folder_path = filedialog.askdirectory()
    if folder_path:  # If a folder is selected
        # Start the conversion in a separate thread to prevent GUI freezing
        threading.Thread(target=convert_all_docx_in_folder, args=(folder_path, progress_bar)).start()

# Initialize the Tkinter window
root = Tk()
root.title("DOCX to PDF Converter")
root.geometry("400x200")  # Adjust the window size as needed

# Create a label for the window
label = Label(root, text="Select a folder containing DOCX files:", font=("Arial", 12))
label.pack(pady=20)

# Create a progress bar
progress_bar = Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress_bar.pack(pady=10)

# Create a button to open the file dialog and start the conversion
convert_button = Button(root, text="Select Folder and Convert", font=("Arial", 12), command=start_conversion)
convert_button.pack(pady=20)

# Start the Tkinter event loop
root.mainloop()
