import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import os
import comtypes.client

wdFormatPDF = 17

input_dir = None
output_dir = None

root = tk.Tk()
root.title("RTF to PDF Converter")

input_frame = tk.Frame(root)
input_frame.pack(side=tk.TOP)

input_label = tk.Label(input_frame, text="Input Folder:")
input_label.pack(side=tk.LEFT)

input_text = tk.Text(input_frame, height=1, width=50)
input_text.pack(side=tk.LEFT)

input_button = tk.Button(input_frame, text="Browse...", command=lambda: select_input_dir())
input_button.pack(side=tk.LEFT)

output_frame = tk.Frame(root)
output_frame.pack(side=tk.TOP)

output_label = tk.Label(output_frame, text="Output Folder:")
output_label.pack(side=tk.LEFT)

output_text = tk.Text(output_frame, height=1, width=50)
output_text.pack(side=tk.LEFT)

output_button = tk.Button(output_frame, text="Browse...", command=lambda: select_output_dir())
output_button.pack(side=tk.LEFT)

convert_button = tk.Button(root, text="Convert Files to PDF", command=lambda: convert_files())
convert_button.pack(side=tk.TOP)

converting_window = None

def select_input_dir():
    global input_dir
    input_dir = filedialog.askdirectory(title='Select Input Folder').replace('/', '\\')
    input_text.delete(1.0, tk.END)
    input_text.insert(tk.END, input_dir)

def select_output_dir():
    global output_dir
    output_dir = filedialog.askdirectory(title='Select Output Folder').replace('/', '\\')
    output_text.delete(1.0, tk.END)
    output_text.insert(tk.END, output_dir)

def convert_files():
    global converting_window
    if not input_dir or not output_dir:
        messagebox.showerror("Error", "Please select an input and output folder")
        return
    converting_window = tk.Toplevel(root)
    converting_window.title("Converting...")
    converting_window.geometry("200x50")
    converting_label = tk.Label(converting_window, text="Converting...")
    converting_label.pack()
    root.update()
    try:
        # Open Word Instance
        word = comtypes.client.CreateObject('Word.Application')
        for subdir, dirs, files in os.walk(input_dir):
            for file in files:
                if file.endswith('.rtf'):
                    in_file = os.path.join(subdir, file)
                    output_file = file.split('.')[0]
                    out_file = os.path.join(output_dir, output_file+'.pdf')

                    doc = word.Documents.Open(in_file)
                    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
                    doc.Close()
        # Close Word Instance
        word.Quit()
    except Exception as e:
        messagebox.showerror("Error", str(e))
    finally:
        converting_window.destroy()
        messagebox.showinfo('Finished', 'File conversion finished!\nClick OK to continue')

root.mainloop()