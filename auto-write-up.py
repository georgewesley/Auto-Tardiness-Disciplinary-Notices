from __future__ import print_function

import os
import sys
import tkinter
from tkinter import*
from tkinter import filedialog, messagebox
from PIL import ImageTk, Image
import PyPDF2
import re
from mailmerge import MailMerge  # pip install docx-mailmerge2 (new version)
import shutil


def create_name_dict(full_text: str):
    remaining_text = clean_text(full_text)
    name_list = []
    name_dict = {}
    while True:
        current = separate_by_name(remaining_text)
        if current[1] == "":  # this is if there is no content in remaining text
            break
        name_list.append(current)
        remaining_text = remaining_text[current[2]:]  # current[2] represents the start of the next employee
        # print("Remaining Text: " + remaining_text)
    for name in name_list:
        for late in find_late(name[1]):
            if name[0] in name_dict.keys():
                name_dict[convert_first_name_last_name(name[0])].append(late)
            else:
                name_dict[convert_first_name_last_name(name[0])] = [late]
    return name_dict


def separate_by_name(remaining_text: str):
    comma_index = remaining_text.find(",")

    substring_lastname = remaining_text[:comma_index]
    start_name = substring_lastname.rfind("\n")  # finds last index of space which we can use to get name of person

    substring_firstname = remaining_text[comma_index:]
    end_name = substring_firstname.find("\n") + len(substring_lastname)

    employee_data = remaining_text[end_name:]
    employee_data = employee_data[:employee_data.find(",")]

    next_employee_start = employee_data.rfind("\n")
    employee_data = employee_data[:next_employee_start]

    next_employee_start += end_name  # makes the index relative to full_text instead of employee_data

    name = remaining_text[start_name:end_name]
    # print("Start Name: " + str(start_name) + " End Name: " + str(end_name))
    return [name.strip(), employee_data.strip(), next_employee_start]


def find_closest_date(text: str):  # finds closest date to an index going backwards
    return re.search("[0-9][0-9]/[0-9][0-9]/[0-9][0-9][0-9][0-9]", text)[0]


def get_full_text(pages: list):
    full_text = ""
    for page in pages:
        full_text += page.extract_text()
    return full_text


def find_late(work_text: str):
    late_list = []
    current_text = work_text
    second_parentheses = 0  # start as 0 to make first while loop work
    while True:
        current_text = current_text[second_parentheses:]
        first_parentheses = current_text.find("(")
        second_parentheses = current_text.find(")")
        if first_parentheses != -1:
            first_parentheses += 1
        else:
            break
        if current_text[first_parentheses-2] == 'M':
            minutes_late = convert_to_minutes(current_text[first_parentheses:second_parentheses])
            third_parentheses = second_parentheses+2
            if current_text[third_parentheses] == '(' and \
                    current_text[first_parentheses:second_parentheses] == \
                    current_text[third_parentheses+1:third_parentheses+5]:
                # this will fail if there is a double-digit time, but that would only happen if someone
                # was scheduled 20 hours or more, so I think it is safe to assume that will not happen
                # We know that the data will be written as two of the same numbers (shift time divided by two) right
                # next to each other. If they are the same we know someone missed a shift.
                late_list.append(['did not attend their shift', find_closest_date(current_text[:second_parentheses])])
            elif int(minutes_late) > root.time:
                late_list.append(['was ' + minutes_late + ' minutes late to their shift',
                                  find_closest_date(current_text[:second_parentheses])])
        second_parentheses += 1  # so we do not include next one
    return late_list


def clean_text(full_text: str):
    # maybe remove this except for the re.sub
    full_text = full_text.replace('Employee', "")
    full_text = full_text.replace('Name', "")
    full_text = full_text.replace('Date', "")
    full_text = full_text.replace('Actual', "")
    full_text = full_text.replace('Time', "")
    full_text = full_text.replace('Scheduled', "")  # if anyone's name has this in it there would be a problem,
    # capital letters should prevent that
    full_text = re.sub("[0-9][0-9]/[0-9][0-9]/[0-9][0-9][0-9][0-9] [0-9][0-9]:[0-9][0-9]:[0-9][0-9] [A-Z]M", "", full_text)
    return full_text


def convert_to_minutes(number_string: str):
    number_string = number_string.split(':')
    num_minutes = 0
    num_minutes += 60 * int(number_string[0])
    num_minutes += int(number_string[1])
    return str(num_minutes)


def convert_first_name_last_name(last_name_first_name: str):
    substring = last_name_first_name.split(",")
    print(substring[0])
    print(substring[1])
    first_name_last_name = substring[1].strip() + " " + substring[0]
    return first_name_last_name


def open_file():
    try:
        root.file = filedialog.askopenfilename(defaultextension=".pdf", initialdir="Desktop", title="Select A File",
                                               filetypes=[('pdf file', '*.pdf')])
    except:
        messagebox.showerror("Error", "Please select a valid pdf file")


def create_writeups():
    try:
        pdfFileObj = open(root.file, 'rb')
        pdfReader = PyPDF2.PdfReader(pdfFileObj)
        page_list = pdfReader.pages
        template = resource_path("files/template.docx")
        name_dict = create_name_dict(get_full_text(page_list))
        if name_dict == {}:
            messagebox.showerror("Error", "Please select a valid pdf file")
        else:
            for name in name_dict.keys():
                for write_up in name_dict[name]:
                    writeup_document = MailMerge(template)
                    writeup_document.merge(
                        Name=name,
                        Location=root.location,
                        Late=write_up[0],
                        IM=root.IM,
                        Date=write_up[1])
                    document_name = name.replace(" ", "_") + "_" + write_up[1].replace("/", "_") + '.docx'
                    writeup_document.write(remove_file_name_from_root_file()+document_name)

        pdfFileObj.close()
    except AttributeError:
        messagebox.showerror("Error", "Please select a valid pdf file")


def resource_path(relative_path):
    # credit to https://stackoverflow.com/questions/51060894/adding-a-data-file-in-pyinstaller-using-the-onefile-option/51061279
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def remove_file_name_from_root_file():
    last_index = str(root.file).rfind('/')
    return root.file[:last_index+1]


def write_IM(entry):
    root.IM = IM_name.get()


def write_location(entry):
    root.location = location.get()


def write_time(entry):
    root.time = time_to_be_late.get()

root = Tk()
root.time = 5  # default time is 5 minutes
root.title("Chick-fil-A Roster Report")
root.iconbitmap(resource_path('files/chick.ico'))
icon = tkinter.Image("photo", file=resource_path("files/chick.png"))
root.tk.call('wm','iconphoto', root._w, icon)
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
app_width = screen_width/2
app_height = screen_height/2
root.geometry(str(int(app_width)) + 'x' + str(int(app_height)))

img = (Image.open(resource_path("files/chick3.png")))
resized_image = img.resize((int(app_width), int(app_height)), Image.ANTIALIAS)
background_image = ImageTk.PhotoImage(resized_image)
background_label = Label(root, image=background_image)
background_label.place(x=0, y=0, relwidth=1, relheight=1)

IM_name = Entry(root, width=int(app_width/20))
IM_name.insert(0, "Interim Manager Name")
IM_name.place(relx=.25, rely=.1, anchor=CENTER)
IM_name.bind("<Return>", write_IM)

location = Entry(root, width=int(app_width/20))
location.insert(0, "Location Name")
location.place(relx=.75, rely=.1, anchor=CENTER)
location.bind("<Return>", write_location)

time_to_be_late = Entry(root, width=int(app_width/20))
time_to_be_late.insert(0, "Minutes to be late? Enter just number, default is 5")
time_to_be_late.place(relx=.5, rely=.2, anchor=CENTER)
time_to_be_late.bind("<Return>", write_time)

pdf_button = Button(root, text="Select PDF File", command=open_file, fg='blue').place(relx=.25, rely=.6, anchor=CENTER)
write_up_button = Button(root, text="Generate Write Ups", command=create_writeups).place(relx=.75, rely=.6, anchor=CENTER)
button_exit = Button(root, text="Exit Program", command=root.quit, fg='red').place(relx=.5, rely=.8, anchor=CENTER)
root.mainloop()

# pyinstaller --add-data='files/chick.png:files' --add-data='files/chick3.png:files' --add-data='files/template.docx:files' --onefile auto-write-up.py --windowed --icon=files/chick3.png
#  That is the command line in terminal to convert this to an exe


