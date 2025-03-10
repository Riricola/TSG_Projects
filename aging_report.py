import numpy as np
import pandas as pd
import openpyxl
import os
import datetime
from datetime import date
import win32com.client as win32
import pandas as pd


# read in the excel file
data = pd.read_excel(r"C:\Users\Maria.Rodriguez\VSCode Projects\agingReport\data\PL AR Aging.xlsx",engine="openpyxl")
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

# Get unique names from the filtered dataset
unique_names = late["Credit Manager"].unique()
        # figure how to email these ppl

late = late.sort_values(by="60+")

# If you want a dataframe with only unique records based on the name column
unique_late_records = late.drop_duplicates(subset=["Credit Manager"])

# add a T/F or 1/0 column. then group by that subset, and extract the ppl

# sort 60+ column from largest to smallest

# filter out everything >0

# create a list of all the credit managers included, send them an email (on a daily,, weekly basis?)


late.to_excel(r"C:\Users\Maria.Rodriguez\VSCode Projects\agingReport\data\test.xlsx")


# Load the Excel file
df = pd.read_excel(r"C:\Users\Maria.Rodriguez\VSCode Projects\Book1.xlsx")  # Update file path as needed
df = df.drop_duplicates(subset=["Credit Manager"])

# Get email addresses from the first column
email_list = df.iloc[:, 0].dropna().tolist()
email_recipients = ";".join(email_list)

# Connect to Outlook
outlook = win32.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)

# Set email details
mail.Subject = "Your Subject Here"
mail.Body = "Your email message here."
mail.BCC = email_recipients  # Sends email in BCC

# Send email (comment out the next line if you just want to open the draft)
# mail.Send()

# Open email as a draft instead of sending
mail.Display()

print("Email draft created successfully!")
