import os
import fnmatch
from docx import Document
from PyPDF2 import PdfReader
import webbrowser
from tkinter import Tk, ttk, Scrollbar, Canvas, VERTICAL, Button, Entry, Label, filedialog
import threading

def search_files(directory, words):
    results = []
    for root, _, files in os.walk(directory):
        for filename in files:
            filepath = os.path.join(root, filename)
            if any(fnmatch.fnmatch(filename.lower(), pattern) for pattern in ['*.csv', '*.docx', '*.txt', '*.pdf', '*.xlsx']):
                try:
                    if filename.lower().endswith('.csv') or filename.lower().endswith('.txt'):
                        with open(filepath, 'r', encoding='utf-8') as file:
                            content = file.read().lower()
                            if not words or all(word.lower() in content for word in words):
                                results.append((filename, filepath))
                    elif filename.lower().endswith('.docx'):
                        try:
                            doc = Document(filepath)
                            content = ' '.join([paragraph.text.lower() for paragraph in doc.paragraphs])
                            if not words or all(word.lower() in content for word in words):
                                results.append((filename, filepath))
                        except Exception as e:
                            print(f"Error reading '{filepath}': {e}")
                    elif filename.lower().endswith('.pdf'):
                        try:
                            with open(filepath, 'rb') as file:
                                pdf_reader = PdfReader(file)
                                content = ' '.join([page.extract_text().lower() for page in pdf_reader.pages])
                                if not words or all(word.lower() in content for word in words):
                                    results.append((filename, filepath))
                        except Exception as e:
                            print(f"Error reading '{filepath}': {e}")
                    elif filename.lower().endswith('.xlsx'):
                        try:
                            pass
                        except Exception as e:
                            print(f"Error reading '{filepath}': {e}")
                except UnicodeDecodeError:
                    continue
                except PermissionError:
                    print(f"Permission denied: '{filepath}'")
                    continue
    return results

def open_file(filepath):
    webbrowser.open(filepath)

def browse_directory():
    directory = filedialog.askdirectory()
    directory_entry.delete(0, 'end')
    directory_entry.insert(0, directory)

def start_search():
    loading_label.config(text="Loading...")
    directory = directory_entry.get()
    words_input = words_entry.get().strip()
    words_to_find = [word.strip() for word in words_input.split(',')] if words_input else []

    thread = threading.Thread(target=run_search, args=(directory, words_to_find))
    thread.start()

def run_search(directory, words_to_find):
    results = search_files(directory, words_to_find)

    root.after(0, display_results, results)
    loading_label.config(text="")

def display_results(results):
    for widget in results_frame.winfo_children():
        widget.grid_forget()

    ttk.Label(results_frame, text="File name", width=30).grid(row=0, column=0, padx=10, pady=5)
    ttk.Label(results_frame, text="Where?", width=50).grid(row=0, column=1, padx=10, pady=5)
    ttk.Label(results_frame, text="Open", width=10).grid(row=0, column=2, padx=10, pady=5)

    for i, result in enumerate(results, start=1):
        ttk.Label(results_frame, text=result[0], width=30).grid(row=i, column=0, padx=10, pady=5)
        ttk.Label(results_frame, text=result[1], width=50).grid(row=i, column=1, padx=10, pady=5)
        button = Button(results_frame, text="Open", command=lambda path=result[1]: open_file(path))
        button.grid(row=i, column=2, padx=10, pady=5)

    canvas.configure(scrollregion=canvas.bbox('all'))

def on_closing():
    root.quit()
    root.destroy()
    os._exit(0) 

root = Tk()
root.title("FINDER!")

root.grid_rowconfigure(3, weight=1)
root.grid_columnconfigure(1, weight=1)

loading_label = Label(root, text="")
loading_label.grid(row=0, column=1, padx=10, pady=5)

Label(root, text="Directory:").grid(row=0, column=0, padx=10, pady=5, sticky='w')
directory_entry = Entry(root, width=50)
directory_entry.grid(row=0, column=1, padx=10, pady=5, sticky='ew')
browse_button = Button(root, text="Browse", command=browse_directory)
browse_button.grid(row=0, column=2, padx=10, pady=5)

Label(root, text="Words to find:").grid(row=1, column=0, padx=10, pady=5, sticky='w')
words_entry = Entry(root, width=50)
words_entry.grid(row=1, column=1, padx=10, pady=5, sticky='ew')

Label(root, text="").grid(row=2, column=0)
start_button = Button(root, text="Start", command=start_search)
start_button.grid(row=3, column=1, pady=10, sticky='ew')
Label(root, text="").grid(row=2, column=2)

canvas = Canvas(root)
scroll_y = Scrollbar(root, orient=VERTICAL, command=canvas.yview)
results_frame = ttk.Frame(canvas)
canvas.create_window((0, 0), window=results_frame, anchor='nw')

def on_frame_configure(event):
    canvas.configure(scrollregion=canvas.bbox('all'))

results_frame.bind("<Configure>", on_frame_configure)

canvas.configure(yscrollcommand=scroll_y.set)
canvas.grid(row=4, column=0, columnspan=3, padx=10, pady=10, sticky='nsew')
scroll_y.grid(row=4, column=3, padx=10, pady=10, sticky='ns')

root.protocol("WM_DELETE_WINDOW", on_closing)

root.mainloop()
