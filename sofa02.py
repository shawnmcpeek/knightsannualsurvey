import tkinter
from tkinter import ttk
import openpyxl
import os



def enter_data():
    category = category_combobox.get()
    name = progname_combobox.get()
    date = program_date_entry.get()
    donations = program_donation_entry.get()
    hours = program_hours_entry.get()
    

    filepath="data.xlsx"

    if not os.path.exists(filepath):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        heading = ["Program Category", "Program Name", "Date of Program", "Charitable Disbursements", "Hours of Service"]
        sheet.append(heading)
        workbook.save(filepath)
    workbook = openpyxl.load_workbook(filepath)
    sheet = workbook.active
    sheet.append([category, name, date, donations, hours])
    workbook.save(filepath)


window = tkinter.Tk()
window.title("Survey of Fraternal Activity Form")

frame = tkinter.Frame(window)
frame.pack()

program_info_frame = tkinter.LabelFrame(frame, text="Program Information")
program_info_frame.grid(row=0, padx=20, pady=10)

#3rd Box - date
program_date_label = tkinter.Label(program_info_frame, text="Date of Program")
program_date_label.grid(row=0, column=2)
program_date_entry = tkinter.Entry(program_info_frame)
program_date_entry.grid(row=1, column=2)

#4th Box - donations
program_donation_label = tkinter.Label(program_info_frame, text="Charitable Donations $")
program_donation_label.grid(row=0, column=3)
program_donation_entry = tkinter.Entry(program_info_frame)
program_donation_entry.grid(row=1, column=3)

#5th Box - hours
program_hours_label = tkinter.Label(program_info_frame, text="Hours of Service")
program_hours_label.grid(row=0, column=4)
program_hours_entry = tkinter.Entry(program_info_frame)
program_hours_entry.grid(row=1, column=4)

#1st Box - Category
category_label = tkinter.Label(program_info_frame, text="Program Category")
category_combobox = ttk.Combobox(program_info_frame, values=["Life", "Family", "Community", "Faith"])
category_label.grid(row=0, column=0)
category_combobox.grid(row=1, column=0)

#2nd Box - Program Name
progname_label = tkinter.Label(program_info_frame, text="Program Name")
progname_combobox = ttk.Combobox(program_info_frame, values=["Faith - RSVP", "Faith - Church Facilities", "Faith - Catholic Schools/Seminaries", "Faith - Religious/Vocations Education", "Faith - Prayer & Study Programs", "Faith - Sacramental Gifts", "Faith - Miscellaneous Faith Activities","Family - Food for Families", "Family - Family Formation Programs", "Family - Keep Christ in Christmas", "Family - Family Week", "Family - Family Prayer Night", "Family - Miscellaneous Family Activities","Community - Coats for Kids", "Community - Global Wheelchair Mission", "Community - Habitat for Humanity", "Community - Disaster Preparedness/Relief", "Community - Physically Disabled/Intellectual Disabilities", "Community - Elderly/Widow(er) Care", "Community - Hospitals/Health Organizations", "Community - Columbian Squires", "Community - Scouting/Youth Groups", "Community - Athletics", "Community - Youth Welfare/Service", "Community - Scholarships/Education", "Community - Veteran Military/VAVS", "Community - Miscellaneous Community/Youth Activities","Life - Special Olympics", "Life - Marches for Life", "Life - Ultrasound Initiative", "Life - Pregnancy Center Support/ASAP", "Life - Christian Refugee Relief", "Life - Memorials to Unborn Children", "Life - Miscellaneous Life Activities"])
progname_label.grid(row=0, column=1)
progname_combobox.grid(row=1, column=1)

for widget in program_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

#Accept submission button
button = tkinter.Button(frame, text="Press to Save", command = enter_data)
button.grid(row=3, column=0, sticky = "news", padx=20, pady=10)




window.mainloop()