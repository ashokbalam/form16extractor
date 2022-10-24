import tabula
import csv
import os
import xlsxwriter
from datetime import datetime
from tkinter import *
from tkinter import filedialog
from pathlib import Path


def change_status(status, color):
    root.status_text["text"] = status
    root.status_text["fg"] = color
    root.update()


def start_processing(path):
    files_location = Path(path)
    files = files_location.glob('**/*.pdf')
    output_file_name = 'Report/Extract' + datetime.now().strftime("%d-%m-%Y-%H-%M") + '.xlsx'

    workbook = xlsxwriter.Workbook(output_file_name)
    worksheet = workbook.add_worksheet("My sheet")
    headers = ["File Name",
               "Employee PAN",
               "Q1 Amount paid/credited",
               "Q1 Amount of tax deducted",
               "Q1 Amount of tax deposited/remitted",
               "Q2 Amount paid/credited",
               "Q2 Amount of tax deducted",
               "Q2 Amount of tax deposited/remitted",
               "Q3 Amount paid/credited",
               "Q3 Amount of tax deducted",
               "Q3 Amount of tax deposited/remitted",
               "Q4 Amount paid/credited",
               "Q4 Amount of tax deducted",
               "Q4 Amount of tax deposited/remitted",
               "Total amount paid/credited",
               "Total amount of tax deducted",
               "Total amount of tax deposited/remitted",
               "[1(d)] Gross salary income",
               "[9] Gross total income",
               "[12] Total taxable income",
               "[19] Net tax payable"]
    row = 0
    col = 0
    for title in headers:
        worksheet.write(row, col, title)
        col += 1

    for file in files:
        row += 1
        # print(file)
        worksheet.write(row, 0, str(file))
        pdf_name = str(file).split('\\')[-1]
        # print(pdf_name)
        change_status("Processing " + pdf_name, 'grey')

        # convert pdf to csv
        tabula.convert_into(file, "temp.csv", output_format="csv", pages='all')

        # convert csv to text
        with open('temp.txt', "w") as my_output_file:
            with open("temp.csv", "r") as my_input_file:
                [my_output_file.write(" ".join(row) + '\n') for row in csv.reader(my_input_file)]
            my_output_file.close()

        with open('temp.txt', 'r') as my_input_file:
            words = my_input_file.read().splitlines()
            my_input_file.close()

        emp_pan = ""
        for x in words:
            if emp_pan == "" and len(x) == 32 and len(x.split(" ")) == 3:
                emp_pan = x.split(" ")[-1]
                # print("Employee PAN: " + emp_pan)
                worksheet.write(row, 1, emp_pan)

            if x.startswith('Q1'):
                quarter_split = x.split(" ")
                # print("Q1 Amount paid/credited: " + quarter_split[2])
                worksheet.write(row, 2, str(quarter_split[2]))
                # print("Q1 Amount of tax deducted: " + quarter_split[3])
                worksheet.write(row, 3, str(quarter_split[3]))
                # print("Q1 Amount of tax deposited/remitted: " + quarter_split[4])
                worksheet.write(row, 4, str(quarter_split[4]))

            if x.startswith('Q2'):
                quarter_split = x.split(" ")
                # print("Q2 Amount paid/credited: " + quarter_split[2])
                worksheet.write(row, 5, str(quarter_split[2]))
                # print("Q2 Amount of tax deducted: " + quarter_split[3])
                worksheet.write(row, 6, str(quarter_split[3]))
                # print("Q2 Amount of tax deposited/remitted: " + quarter_split[4])
                worksheet.write(row, 7, str(quarter_split[4]))

            if x.startswith('Q3'):
                quarter_split = x.split(" ")
                # print("Q3 Amount paid/credited: " + quarter_split[2])
                worksheet.write(row, 8, str(quarter_split[2]))
                # print("Q3 Amount of tax deducted: " + quarter_split[3])
                worksheet.write(row, 9, str(quarter_split[3]))
                # print("Q3 Amount of tax deposited/remitted: " + quarter_split[4])
                worksheet.write(row, 10, str(quarter_split[4]))

            if x.startswith('Q4'):
                quarter_split = x.split(" ")
                # print("Q4 Amount paid/credited: " + quarter_split[2])
                worksheet.write(row, 11, str(quarter_split[2]))
                # print("Q4 Amount of tax deducted: " + quarter_split[3])
                worksheet.write(row, 12, str(quarter_split[3]))
                # print("Q4 Amount of tax deposited/remitted: " + quarter_split[4])
                worksheet.write(row, 13, str(quarter_split[4]))

            if x.startswith('Total (Rs.)') and len(x.split(" ")) == 5:
                split = x.split(" ")
                # print("Total amount paid/credited: " + split[2])
                worksheet.write(row, 14, str(split[2]))
                # print("Total amount of tax deducted: " + split[3])
                worksheet.write(row, 15, str(split[3]))
                # print("Total amount of tax deposited/remitted: " + split[4])
                worksheet.write(row, 16, str(split[4]))

            if x.startswith('(d)') and 'Total' in x and '80C' not in x:
                split = x.split(" ")
                # print("[1(d)] Gross salary income: " + str(split[-1]))
                worksheet.write(row, 17, str(split[-1]))

            if x.startswith('9') and 'Gross total income' in x:
                split = x.split(" ")
                # print("[9] Gross total income: " + str(split[-1]))
                worksheet.write(row, 18, str(split[-1]))

            if x.startswith('12') and 'Total taxable income' in x:
                split = x.split(" ")
                # print("[12] Total taxable income: " + str(split[-3]))
                worksheet.write(row, 19, str(split[-3]))

            if x.startswith('19') and 'Net tax payable' in x:
                split = x.split(" ")
                # # print("[19] Net tax payable: " + str(split[-3]))
                worksheet.write(row, 20, str(split[-3]))
        os.remove("temp.csv")
        os.remove("temp.txt")
    workbook.close()
    change_status("Extraction completed: " + output_file_name, "#007f5f")
    root.browse_button["state"] = "normal"
    root.update()


def browse_button():
    folder_name = filedialog.askdirectory()
    change_status('', 'white')
    if folder_name == '':
        change_status("Please select a folder!", "#d90429")
    if any(File.endswith(".pdf") for File in os.listdir(folder_name)):
        root.browse_button["state"] = "disabled"
        root.update()
        start_processing(folder_name)
    else:
        change_status("Selected folder does not have any PDF file.", "#d90429")


root = Tk()
root.geometry("800x200")
root.resizable(0, 0)
root.iconbitmap('extract-folder.ico')
root.title("Form16 Data Extractor")

root.selection_text = Label(text='Browse and select the folder which has all Form16 PDF\'s', font=('Roboto', 15))
root.selection_text.pack(pady=20)

root.browse_button = Button(text="Select Folder", command=browse_button, bg='#212529', fg='white', font=('Roboto', 13),
                            borderwidth="0", pady=5, padx=5)
root.browse_button.pack(pady=20)

root.status_text = Label(text="", fg='#d90429', font=('Roboto', 13))
root.status_text.pack(pady=10)

root.mainloop()
