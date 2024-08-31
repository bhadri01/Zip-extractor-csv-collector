import zipfile
import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk

def copy_to_extracted_folder(source_folder, extraction_base_path):
    """
    Copies all folders and .zip files from the source folder to the extraction base path.

    Parameters:
    source_folder (str): The original folder selected by the user.
    extraction_base_path (str): The base path where files and folders will be copied.
    """
    # Ensure the extraction folder exists
    if not os.path.exists(extraction_base_path):
        os.makedirs(extraction_base_path)

    for item in os.listdir(source_folder):
        full_path = os.path.join(source_folder, item)
        destination_path = os.path.join(extraction_base_path, item)

        try:
            if os.path.isdir(full_path):
                # Copy folder to the extracted directory
                shutil.copytree(full_path, destination_path, dirs_exist_ok=True)
                log_message(f"Copied folder '{full_path}' to '{destination_path}'.")
                update_progress()

            elif os.path.isfile(full_path) and full_path.endswith('.zip'):
                # Copy ZIP file to the extracted directory
                shutil.copy2(full_path, destination_path)
                log_message(f"Copied ZIP file '{full_path}' to '{destination_path}'.")
                update_progress()

        except Exception as e:
            log_message(f"Error while copying '{full_path}' to '{destination_path}': {e}")

def find_and_extract_zip_files(extraction_base_path):
    """
    Iteratively finds all .zip files in the specified folder, extracts them,
    within the extraction base path.

    Parameters:
    extraction_base_path (str): The base path where extracted folders are located.
    """
    # Use a stack to manage folders to process
    stack = [extraction_base_path]

    while stack:
        current_path = stack.pop()

        for item in os.listdir(current_path):
            # Construct the full path of the file or folder
            full_path = os.path.join(current_path, item)

            try:
                if os.path.isfile(full_path) and full_path.endswith('.zip'):
                    # If the item is a ZIP file, extract it in place
                    extract_to_folder = os.path.join(current_path, os.path.splitext(item)[0])

                    # Check if the extraction folder exists, if not create it
                    if not os.path.exists(extract_to_folder):
                        os.makedirs(extract_to_folder)

                    # Extract the ZIP file
                    with zipfile.ZipFile(full_path, 'r') as zip_ref:
                        zip_ref.extractall(extract_to_folder)
                        log_message(f"Successfully extracted '{full_path}' to '{extract_to_folder}'.")
                        update_progress()

                    # Add the newly extracted folder to the stack for further processing
                    stack.append(extract_to_folder)

                elif os.path.isdir(full_path):
                    # If it's a folder, add it to the stack to process its contents
                    stack.append(full_path)

            except FileNotFoundError as e:
                log_message(f"Error: File not found '{e.filename}'. Skipping...")

            except zipfile.BadZipFile as e:
                log_message(f"Error: Bad ZIP file '{full_path}'. Skipping...")

            except Exception as e:
                log_message(f"Unexpected error '{e}' occurred with '{full_path}'. Skipping...")

def organize_files_by_extension(extraction_base_path, merged_path):
    """
    Finds all .csv and .pdf files in the extracted folders and subfolders,
    and copies them to respective folders in the merged directory.

    Parameters:
    extraction_base_path (str): The base path where extracted folders are located.
    merged_path (str): The path where all merged CSV and PDF files will be saved.
    """
    csv_folder = os.path.join(merged_path, 'All_CSVs')
    pdf_folder = os.path.join(merged_path, 'All_PDFs')

    # Create the destination folders if they do not exist
    if not os.path.exists(csv_folder):
        os.makedirs(csv_folder)
    if not os.path.exists(pdf_folder):
        os.makedirs(pdf_folder)

    # Walk through all folders and subfolders to find .csv and .pdf files
    for root, _, files in os.walk(extraction_base_path):
        for file_name in files:
            file_path = os.path.join(root, file_name)

            try:
                if file_name.endswith('.csv'):
                    # Copy .csv files to the All_CSVs folder
                    shutil.copy(file_path, csv_folder)
                    log_message(f"Copied CSV file '{file_path}' to '{csv_folder}'.")
                    update_progress()

                elif file_name.endswith('.pdf'):
                    # Copy .pdf files to the All_PDFs folder
                    shutil.copy(file_path, pdf_folder)
                    log_message(f"Copied PDF file '{file_path}' to '{pdf_folder}'.")
                    update_progress()

            except FileNotFoundError as e:
                log_message(f"Error: File not found '{e.filename}'. Skipping...")

            except shutil.Error as e:
                log_message(f"Error while copying '{file_path}': {e}. Skipping...")

            except Exception as e:
                log_message(f"Unexpected error '{e}' occurred with '{file_path}'. Skipping...")

def log_message(message):
    """
    Logs a message to the text widget in the GUI.

    Parameters:
    message (str): The message to log.
    """
    log_text.config(state=tk.NORMAL)
    log_text.insert(tk.END, message + '\n')
    log_text.config(state=tk.DISABLED)
    log_text.see(tk.END)

def update_progress():
    """
    Updates the progress bar to reflect the progress of the tasks.
    """
    progress_bar['value'] += 1
    root.update_idletasks()

def calculate_total_tasks(extraction_base_path):
    """
    Calculates the total number of tasks for copying and extracting files.

    Parameters:
    extraction_base_path (str): The base path where extracted folders are located.

    Returns:
    int: Total number of tasks.
    """
    total_tasks = 0
    for root, _, files in os.walk(extraction_base_path):
        for file_name in files:
            if file_name.endswith('.zip') or file_name.endswith('.csv') or file_name.endswith('.pdf'):
                total_tasks += 1
    return total_tasks

def select_folder_and_process():
    """
    Opens a file dialog for the user to select a folder and then runs the extraction
    and organization processes within that folder.
    """
    # Open a file dialog for folder selection
    folder_path = filedialog.askdirectory(title="Select a Folder")
    if not folder_path:
        messagebox.showwarning("No Folder Selected", "Please select a folder to proceed.")
        return

    # Define paths for extracted and merged folders within the selected folder
    extraction_base_path = os.path.join(folder_path, 'Extracted')
    merged_path = os.path.join(folder_path, 'Merged')

    # Create the merged folder if it does not exist
    if not os.path.exists(merged_path):
        os.makedirs(merged_path)

    # Copy original files and folders to the extraction base path
    copy_to_extracted_folder(folder_path, extraction_base_path)

    # Calculate total tasks for the progress bar
    total_tasks = calculate_total_tasks(extraction_base_path)
    progress_bar['maximum'] = total_tasks

    # Run the extraction process within the extracted folder
    find_and_extract_zip_files(extraction_base_path)

    # Organize CSV and PDF files into the merged folder
    organize_files_by_extension(extraction_base_path, merged_path)

    # Display a success message
    messagebox.showinfo("Process Complete", "Files have been extracted and organized successfully.")
    log_message("Process completed successfully!")

# GUI setup using tkinter
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Zip File Extractor and File Organizer")
    root.geometry("700x500")  # Increase window size

    # Configure styles for better UI/UX
    style = ttk.Style()
    style.configure('TButton', font=('Arial', 12, 'bold'), foreground='white', background='#4CAF50')
    style.configure('TLabel', font=('Arial', 14, 'bold'), foreground='#333')

    # Frame for the button
    frame = tk.Frame(root, pady=10, padx=10)
    frame.pack(pady=20)

    # Create a button to select folder and start processing
    select_button = ttk.Button(frame, text="Select Folder and Start Processing", command=select_folder_and_process)
    select_button.grid(row=0, column=0, padx=10, pady=10)

    # Text widget for logging messages
    log_text = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=80, height=20, state=tk.DISABLED, font=('Arial', 10))
    log_text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

    # Progress bar
    progress_bar = ttk.Progressbar(root, orient='horizontal', mode='determinate', length=500)
    progress_bar.pack(pady=10)

    # Run the tkinter main loop
    root.mainloop()
