from tkinter import Tk, Label, Text, Button, Frame, CENTER
import re
from datetime import datetime

def extract_info(text):
    date_match = re.search(r'Delivered Date:\s*(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})', text)
    date = datetime.strptime(date_match.group(1), '%Y-%m-%d %H:%M:%S')
    date = f"{date.month}/{date.day}/{date.year}"
    
    name = re.search(r'Location:\s*(.*?)\s*Request Date:', text, re.DOTALL).group(1).strip()
    
    # extract place from text parameter
    place_match = re.search(r'\(([^)\s][^)]*[^)\s])\)', text)
    place = place_match.group(1) if place_match else ''
    if '-' in place:
        place = place.split('-')[-1]
    
    
    incident_no = re.search(r'Incident No:\s*(.*?)\s*Part ID', text, re.DOTALL).group(1).strip()

    parts_info = re.findall(r'(\w+-\w+|\w+-\w+-\w+|\w+-\w+-\w+-\w+|\w+) \((.*?)\)\s*(\d+)', text)
    parts = [{'Date': date, 'Part ID': part_id, 'Part Name': part_name.strip(), 'Quantity': qty, 'Name': name, 'Place': place, 'Incident No.': incident_no} for part_id, part_name, qty in parts_info]

    return parts

def format_table(parts):
    table = ''
    for part in parts:
        table += '\t'.join(str(part[header]) for header in part.keys()) + '\n'

    return table

def paste_text():
    textbox.delete("1.0", "end")
    textbox.insert("end", root.clipboard_get())

def content_to_grid():
    text = textbox.get("1.0", "end-1c")
    
    try:    
        parts = extract_info(text)
        table = format_table(parts)
        
        textbox.delete("1.0", "end")
        textbox.insert("end", table)

    except:
        textbox.delete("1.0", "end")
        textbox.insert("end", "Invalid content. Please make sure that the content is in the correct format.")
    
def copy_content():
    text = textbox.get("1.0", "end-1c")
    root.clipboard_clear()
    root.clipboard_append(text)
    
    textbox.delete("1.0", "end")
    textbox.insert("end", "Text has succesfully copied!")

root = Tk()
root.title("CE Print to Excel Grid")
root.iconbitmap('icon.ico')

textbox = Text(root, bg='lightgray', bd=2, width=45, height=15, font=("Consolas", 10))
frame = Frame(root)
paste_btn = Button(frame, text="Paste", padx=10, pady=5, command=paste_text)
submit_btn = Button(frame, text="Submit", padx=10, pady=5, command=content_to_grid)
copy_btn = Button(frame, text="Copy", padx=10, pady=5, command=copy_content)

textbox.focus_set()

textbox.pack()
frame.pack()

paste_btn.grid(row=0, column=0)
submit_btn.grid(row=0, column=1)
copy_btn.grid(row=0, column=2)


root.mainloop()