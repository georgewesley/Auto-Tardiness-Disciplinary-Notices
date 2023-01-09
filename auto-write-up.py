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
    for name in name_list:
        for late in find_late(name[1]):
            if convert_first_name_last_name(name[0]) in name_dict.keys():
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
    return [name.strip(), employee_data.strip(), next_employee_start]


def find_closest_date(text: str):  # finds closest date to an index going backwards
    return re.findall("[0-9][0-9]/[0-9][0-9]/[0-9][0-9][0-9][0-9]", text)[-1]


def get_full_text(pages: list):
    full_text = ""
    for page in pages:
        full_text += page.extract_text()
    return full_text


def find_late(work_text: str):
    late_text = " minutes late to their shift"
    was_text = "was "
    no_show_text = "did not attend their shift"
    if root.spanish.get() == 1:
        late_text = " minutos tarde a su turno"
        no_show_text = "no asistió a su turno"
        was_text = "estuvo "
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
            actual_minutes_late = str(convert_shift_to_minutes(current_text[first_parentheses-18:first_parentheses-11],
                                                           current_text[first_parentheses-9:first_parentheses-2]))
            # above line is valid because we always know that if there is an M before a ( that the pattern will hold
            minutes_late = convert_to_minutes(current_text[first_parentheses:second_parentheses])
            third_parentheses = second_parentheses+2
            if current_text[third_parentheses] == '(' and \
                    current_text[first_parentheses:second_parentheses] == \
                    current_text[third_parentheses+1:third_parentheses+5] and convert_to_minutes(current_text[first_parentheses:second_parentheses]) == actual_minutes_late:
                # this will fail if there is a double-digit time, but that would only happen if someone
                # was scheduled 20 hours or more, so I think it is safe to assume that will not happen
                # We know that the data will be written as two of the same numbers (shift time divided by two) right
                # next to each other. If they are the same we know someone missed a shift.
                late_list.append([no_show_text, find_closest_date(current_text[:second_parentheses])])
            elif int(minutes_late) > int(root.time):
                late_list.append([was_text + minutes_late + late_text,
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


def convert_to_minutes(number_string: str, pm_hours=0):
    number_string = number_string.split(':')
    num_minutes = 0
    num_minutes += 60 * (int(number_string[0]) + pm_hours)
    num_minutes += int(number_string[1])
    return str(num_minutes)


def convert_shift_to_minutes(part_one_shift: str, part_two_shift: str):
    if part_one_shift[-1] == "A":
        part_one_shift = convert_to_minutes(part_one_shift[:-2])
    else:
        part_one_shift = convert_to_minutes(part_one_shift[:-2], 12)
    if part_two_shift[-1] == "A":
        part_two_shift = convert_to_minutes(part_two_shift[:-2])
    else:
        part_two_shift = convert_to_minutes(part_two_shift[:-2], 12)

    return int(abs(int(part_one_shift) - int(part_two_shift))/2)  # we divide by two because the shift missed is split into 2
    # half is in clock in, half is in clock out


def convert_first_name_last_name(last_name_first_name: str):
    substring = last_name_first_name.split(",")
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
        name_dict = create_name_dict(get_full_text(page_list))
        if name_dict == {}:
            messagebox.showerror("Error", "Please select a valid pdf file")
        else:
            for name in name_dict.keys():
                for write_up in name_dict[name]:
                    writeup_document = MailMerge(root.template)
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


def change_to_spanish():
    if root.spanish.get() == 1:
        root.template = resource_path("files/template_spanish.docx")
    else:
        root.template = resource_path("files/template.docx")


root = Tk()
root.spanish = tkinter.IntVar()
root.template = resource_path("files/template.docx")
root.time = 5  # default time is 5 minutes
root.title("Chick-fil-A Roster Report")
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

IM_trace = StringVar()
IM_trace.trace("w", lambda name, index, mode, IM_trace=IM_trace: write_IM(IM_trace))
IM_name = Entry(root, width=int(app_width/20), textvariable=IM_trace)
IM_name.insert(0, "Enter Interim Manager Name...")
IM_name.place(relx=.25, rely=.1, anchor=CENTER)

location_trace = StringVar()
location_trace.trace("w", lambda name, index, mode, location_trace=location_trace: write_location(location_trace))
location = Entry(root, width=int(app_width/20), textvariable=location_trace)
location.insert(0, "Enter Location Name...")
location.place(relx=.75, rely=.1, anchor=CENTER)

time_late_trace = StringVar()
time_late_trace.trace("w", lambda name, index, mode, time_late_trace=time_late_trace: write_time(time_late_trace))
time_to_be_late = Entry(root, width=int(app_width/20), textvariable=time_late_trace)
time_to_be_late.insert(0, "Enter time to be late in minutes (only a number)...")
time_to_be_late.place(relx=.5, rely=.2, anchor=CENTER)

c1 = Checkbutton(root, text='Write Up in Spanish', variable=root.spanish, onvalue=1, offvalue=0, command=change_to_spanish).place(relx=.5, rely=.8, anchor=CENTER)

pdf_button = Button(root, text="Select PDF File", command=open_file, fg='blue').place(relx=.25, rely=.6, anchor=CENTER)
write_up_button = Button(root, text="Generate Write Ups", command=create_writeups).place(relx=.75, rely=.6, anchor=CENTER)
button_exit = Button(root, text="Exit Program", command=root.quit, fg='red').place(relx=.5, rely=.9, anchor=CENTER)
root.mainloop()

# pyinstaller --add-data='files/chick.png:files' --add-data='files/template_spanish.docx:files' --add-data='files/chick3.png:files' --add-data='files/template.docx:files' --onefile auto-write-up.py --windowed --icon=files/chick3.png
#  That is the command line in terminal to convert this to an exe



