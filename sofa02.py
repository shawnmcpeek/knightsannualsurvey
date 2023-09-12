import tkinter
from tkinter import ttk
import openpyxl
import os
import pandas as pd
import PyPDF2
import subprocess


def analyze_data():
    # Read the CSV file into a DataFrame
    df = pd.read_excel("data.xlsx")

    # Sort the DataFrame based on 'Program Category' and 'Program Name'
    sorted_df = df.sort_values(by=["Program Category", "Program Name"])

    # Group by 'Program Name' and calculate the sum of 'Hours of Service' and 'Charitable Disbursements'
    grouped = sorted_df.groupby(["Program Name"]).agg(
        {"Hours of Service": "sum", "Charitable Disbursements": "sum"}
    )

    # Prepare the content for the PDF
    content = f"Annual Survey Report\n\n"

    # Append the analysis results to the content
    for program_name, data in grouped.iterrows():
        content += f"{program_name}:\n"
        content += f"Hours of Service: {data['Hours of Service']}\n"
        content += f"Charitable Disbursements: {data['Charitable Disbursements']}\n\n"

    # Generate the PDF with the content
    generate_pdf(content)

    # Open the PDF with the default PDF viewer
    pdf_file_path = "annual_report.pdf"
    try:
        subprocess.Popen(["xdg-open", pdf_file_path])  # Linux
    except OSError:
        try:
            subprocess.Popen(["open", pdf_file_path])  # macOS
        except OSError:
            try:
                subprocess.Popen(
                    ["start", "AcroRd32.exe", pdf_file_path], shell=True
                )  # Windows
            except OSError as e:
                print(f"Unable to open PDF: {e}")


def enter_data():
    category = category_combobox.get()
    name = progname_combobox.get()
    date = program_date_entry.get()
    donations = program_donation_entry.get()
    hours = program_hours_entry.get()

    filepath = "data.xlsx"

    if not os.path.exists(filepath):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        heading = [
            "Program Category",
            "Program Name",
            "Date of Program",
            "Charitable Disbursements",
            "Hours of Service",
        ]
        sheet.append(heading)
        workbook.save(filepath)
    workbook = openpyxl.load_workbook(filepath)
    sheet = workbook.active
    sheet.append([category, name, date, donations, hours])
    workbook.save(filepath)


def generate_pdf(content):
    # Create a new PDF document
    pdf = PyPDF2.PdfWriter()

    # Create a new page and add the content to it
    page = PyPDF2.PageObject.create_blank_page(
        width=200, height=200
    )  # Adjust width and height as needed
    font = PyPDF2.pdf.TTFont("Arial", "Arial.ttf")
    font_size = 12

    # Create a text object
    text = PyPDF2.pdf.TextObject()
    text.setFont(font, font_size)
    text.setTextMatrix(1, 0, 0, 1, 50, 100)  # Adjust the position (x=50, y=100) as needed
    text.textLines = [content]

    # Draw the text on the page
    page.addText(text)

    # Add the page to the PDF document
    pdf.addPage(page)

    # Save the PDF with the content
    with open("annual_report.pdf", "wb") as pdf_file:
        pdf.write(pdf_file)


window = tkinter.Tk()
window.title("Survey of Fraternal Activity Form")

frame = tkinter.Frame(window)
frame.pack()

program_info_frame = tkinter.LabelFrame(frame, text="Program Information")
program_info_frame.grid(row=0, padx=20, pady=10)

# 3rd Box - date
program_date_label = tkinter.Label(program_info_frame, text="Date of Program")
program_date_label.grid(row=0, column=2)
program_date_entry = tkinter.Entry(program_info_frame)
program_date_entry.grid(row=1, column=2)

# 4th Box - donations
program_donation_label = tkinter.Label(
    program_info_frame, text="Charitable Donations $"
)
program_donation_label.grid(row=0, column=3)
program_donation_entry = tkinter.Entry(program_info_frame)
program_donation_entry.grid(row=1, column=3)

# 5th Box - hours
program_hours_label = tkinter.Label(program_info_frame, text="Hours of Service")
program_hours_label.grid(row=0, column=4)
program_hours_entry = tkinter.Entry(program_info_frame)
program_hours_entry.grid(row=1, column=4)

# 1st Box - Category
category_label = tkinter.Label(program_info_frame, text="Program Category")
category_combobox = ttk.Combobox(
    program_info_frame, values=["Life", "Family", "Community", "Faith"]
)
category_label.grid(row=0, column=0)
category_combobox.grid(row=1, column=0)

# 2nd Box - Program Name
progname_label = tkinter.Label(program_info_frame, text="Program Name")
progname_combobox = ttk.Combobox(
    program_info_frame,
    values=[
        "Faith - RSVP",
        "Faith - Church Facilities",
        "Faith - Catholic Schools/Seminaries",
        "Faith - Religious/Vocations Education",
        "Faith - Prayer & Study Programs",
        "Faith - Sacramental Gifts",
        "Faith - Miscellaneous Faith Activities",
        "Family - Food for Families",
        "Family - Family Formation Programs",
        "Family - Keep Christ in Christmas",
        "Family - Family Week",
        "Family - Family Prayer Night",
        "Family - Miscellaneous Family Activities",
        "Community - Coats for Kids",
        "Community - Global Wheelchair Mission",
        "Community - Habitat for Humanity",
        "Community - Disaster Preparedness/Relief",
        "Community - Physically Disabled/Intellectual Disabilities",
        "Community - Elderly/Widow(er) Care",
        "Community - Hospitals/Health Organizations",
        "Community - Columbian Squires",
        "Community - Scouting/Youth Groups",
        "Community - Athletics",
        "Community - Youth Welfare/Service",
        "Community - Scholarships/Education",
        "Community - Veteran Military/VAVS",
        "Community - Miscellaneous Community/Youth Activities",
        "Life - Special Olympics",
        "Life - Marches for Life",
        "Life - Ultrasound Initiative",
        "Life - Pregnancy Center Support/ASAP",
        "Life - Christian Refugee Relief",
        "Life - Memorials to Unborn Children",
        "Life - Miscellaneous Life Activities",
    ],
)
progname_label.grid(row=0, column=1)
progname_combobox.grid(row=1, column=1)

for widget in program_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

# Accept submission button
button = tkinter.Button(frame, text="Press to Save", command=enter_data)
button.grid(row=3, column=0, sticky="s", padx=20, pady=10)

# Generate Annual Report button
annual_button = tkinter.Button(
    frame, text="Press to Generate Annual Survey", command=analyze_data
)
annual_button.grid(row=4, column=0, sticky="s", padx=20, pady=10)

window.mainloop()
