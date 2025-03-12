# aging_report.py

import FreeSimpleGUI as sg
import datetime
from datetime import date
import win32com.client as win32
import pandas as pd
import xlwings as xw
from openpyxl import load_workbook


layout = [[sg.Combo(sorted(sg.user_settings_get_entry('-filenames-', [])), default_value=sg.user_settings_get_entry('-last filename-', ''), size=(50, 1), key='-FILENAME-'), sg.FileBrowse(), sg.B('Clear History')],
          [sg.Button('Ok', bind_return_key=True),  sg.Button('Cancel')]]

window = sg.Window('AR Aging Report Email-er', layout, margins=(75,30))

while True:
    event, values = window.read()

    if event in (sg.WIN_CLOSED, 'Cancel'):
        break
    if event == 'Ok':
        # If OK, then need to add the filename to the list of files and also set as the last used filename
        sg.user_settings_set_entry('-filenames-', list(set(sg.user_settings_get_entry('-filenames-', []) + [values['-FILENAME-'], ])))
        sg.user_settings_set_entry('-last filename-', values['-FILENAME-'])
        app = xw.App(visible=False)  # Open Excel in the background
        wb = xw.Book(values['-FILENAME-'])
        wb.app.calculation = 'automatic'  # Ensure Excel calculates formulas
        wb.save()  # Save the updated values
        wb.close()
        app.quit()
        data = pd.read_excel(values['-FILENAME-'], engine="openpyxl")
        break
    elif event == 'Clear History':
        sg.user_settings_set_entry('-filenames-', [])
        sg.user_settings_set_entry('-last filename-', '')
        window['-FILENAME-'].update(values=[], value='')


new_headers = data.iloc[2].values
data = data.iloc[3:,:]
data.columns = new_headers
data = data.sort_values(by = "60+")

# calculate a week before today, filter out anything less than a week ago
last_week = date.today() - datetime.timedelta(weeks=1)

# First convert to datetime, then extract just the date part
data.loc[:, 30] = pd.to_datetime(data.iloc[:, 29].str.split(' ').str[0], errors='coerce').dt.date

# Now both are date objects (not Timestamps), so this comparison works
late = data[data.iloc[:, 30] <= last_week].copy()

late = late.sort_values(by="60+", ascending=False)
# Group by "manager" and store in a dictionary
manager_tables = {}

# Connect to Outlook
outlook = win32.Dispatch("Outlook.Application")
for manager, group in late.groupby("Credit Manager"):
    # Format table as HTML
    table_html = group.to_html(index=False, columns=["ID#", "Bill-To Customer", "60+", "Balance", "Last Pmt", "Last Pmt Amt"])
    
    # Create an email item
    mail = outlook.CreateItem(0)
    mail.Subject = f"Aging Report - Overdue Accounts for {manager}"
    
    # Email body
    mail.HTMLBody = f"""
    <html>
    <body>
    <p>Hello {manager},</p>
    <p>Below is the list of accounts that have an outstanding balance and have not been contacted in over a week.</p>
    {table_html}
    <p>Best regards,<br>Your Team</p>
    </body>
    </html>
    """
    
    #mail.BodyFormat = 2  # HTML format
    mail.Display()  # Open as a draft (change to mail.Send() to send immediately)


window.close()
