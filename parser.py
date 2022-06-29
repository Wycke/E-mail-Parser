import email
from traceback import format_exception_only
import pandas as pd
import PySimpleGUI as sg



from email import message_from_file
import os, glob

path = "./Email Parser/msgs"

def caption(origin):
    Date = ""
    if "date" in origin: Date = origin["date"].strip()
    From = ""
    if "from" in origin: From = origin["from"].strip()
    To = ""
    if "to" in origin: To = origin["to"].strip()
    Subject = ""
    if "subject"in origin: Subject = origin["subject"].strip()
    return(From, To, Subject, Date)


def extract(msgfile, key):
    m = message_from_file(msgfile)
    From, To, Subject, Date = caption(m)
    print(From, " ", To, "Subject: ", Subject, " ", Date)
    return(From,To,Subject,Date)


templist = os.listdir("./msgs")
size = len(templist)
for i in range(size):
    templist[i] = "./msgs/"+templist[i]
print(templist)
writer = pd.ExcelWriter('emails.xlsx', engine='xlsxwriter')

from_col = []
to_col = []
subject_col = []
date_col = []


#Layout

layout = [[sg.Text("Filler")], [sg.Button("CONVERT")]]

#Window
window = sg.Window("Demo",layout)

#Events
while True:
    event, values = window.read()
    if event == "CONVERT" or event == sg.WIN_CLOSED:
        for i in range(size):
            f = open(templist[i])
            print(f.name)
            From, To, Subject, Date = extract(f, f.name)
            from_col.append(From)
            to_col.append(To)
            subject_col.append(Subject)
            date_col.append(Date)
        break




window.close()

data = {
    "From": from_col,
    "To": to_col,
    "Subject": subject_col,
    "Date": date_col
}

df = pd.DataFrame(data)

df.to_excel(writer, sheet_name='sheet1', index=False)

writer.save()
f.close()
