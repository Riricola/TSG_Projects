# aging_report.py

import FreeSimpleGUI as sg
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

def format_data(input_file_path, save_path = None):
        
    # Read the first two lines of the file
    with open(input_file_path, 'r', encoding='latin-1') as file:
        for _ in range(4):  # Skip the first 4 lines
            file.readline()
        header_line = file.readline().strip()  # 5th line (headers)
        dashes_line = file.readline().strip()  # 6th line (dashes)

    # Calculate column widths based on the dashes line
    col_widths = [14, 34, 10, 14, 22, 14, 17, 14]
    # Read the fixed-width file into a DataFrame using the calculated widths
    data = pd.read_fwf(input_file_path, widths=col_widths, skiprows=4)

    # Remove empty spaces between rows
    df_cleaned = data.dropna(subset=["Invoice #"])

    
    column_index = 7  # 8th column (0-based index)
    # Identify rows with negative numbers in the 6th column

    df_cleaned['Open Amt'] = df_cleaned['Open Amt'].str.replace(',', '', regex=True)  # Remove commas
    df_cleaned['Orig Amt'] = df_cleaned['Orig Amt'].str.replace(',', '', regex=True)  # Remove commas

    df_cleaned['Open Amt'] = pd.to_numeric(df_cleaned['Open Amt'], errors='coerce')  # Converts invalid strings to NaN
    df_cleaned['Orig Amt'] = pd.to_numeric(df_cleaned['Orig Amt'], errors='coerce') 

    # Convert only negative numbers in the 6th column to "(50)" format
    df_cleaned.iloc[:, column_index] = df_cleaned.iloc[:, column_index].apply(
        lambda x: f"({abs(x)})" if x < 0 else x
    )
    df_cleaned.iloc[:, 5] = df_cleaned.iloc[:, 5].apply(
        lambda x: f"({abs(x)})" if x < 0 else x
    )

    # Filter rows that contain parentheses (which were previously negative)
    #? negatives_df = df_cleaned.iloc[5:, :].copy()  # Select rows from index 5 onwards
    #? negatives_df = negatives_df[negatives_df.iloc[:, column_index].astype(str).str.startswith("(")]

    # Remove those rows from the original DataFrame
    #? df_cleaned = df_cleaned[~df_cleaned.iloc[:, column_index].astype(str).str.startswith("(")]

    # Modified save logic to use the provided save path
    if save_path:
        # Create output filename based on the input file name
        import os
        input_filename = os.path.basename(input_file_path)
        output_filename = os.path.splitext(input_filename)[0] + "_formatted.xlsx"
        output_path = os.path.join(save_path, output_filename)
        
        # Write DataFrame to the new Excel file
        with pd.ExcelWriter(output_path, engine="openpyxl", mode="w") as writer:
            df_cleaned.to_excel(writer, sheet_name="Edited_Test0.1", index=False)
    else:
        # Original behavior - save to the same file
        with pd.ExcelWriter(input_file_path, engine="openpyxl", mode="w") as writer:
            df_cleaned.to_excel(writer, sheet_name="Edited_Test0.1", index=False)


layout = [[sg.Combo(sorted(sg.user_settings_get_entry('-filenames-', [])), default_value=sg.user_settings_get_entry('-last filename-', ''), size=(50, 1), key='-FILENAME-'), sg.FileBrowse(), sg.B('Clear History')],
          [sg.Text("Output Folder:"), sg.InputText(sg.user_settings_get_entry('-last savepath-', ''), key="-SAVEPATH-", size=(43,1)), sg.FolderBrowse()],
          [sg.Button('Ok', bind_return_key=True),  sg.Button('Cancel')]]

window = sg.Window('COD Report Formatter', layout, margins=(75,30))

while True:
    event, values = window.read()

    if event in (sg.WIN_CLOSED, 'Cancel'):
        break
    if event == 'Ok':
        # If OK, then need to add the filename to the list of files and also set as the last used filename
        sg.user_settings_set_entry('-filenames-', list(set(sg.user_settings_get_entry('-filenames-', []) + [values['-FILENAME-'], ])))
        sg.user_settings_set_entry('-last filename-', values['-FILENAME-'])
         # Save the output folder path to settings
        sg.user_settings_set_entry('-last savepath-', values['-SAVEPATH-'])
        
        # Process the data with the save path
        format_data(values['-FILENAME-'], values['-SAVEPATH-'])

         # Show a success message
        if values['-SAVEPATH-']:
            sg.popup(f"File processed and saved to {values['-SAVEPATH-']}")
        else:
            sg.popup("File processed and saved")
        
        break
    elif event == 'Clear History':
        sg.user_settings_set_entry('-filenames-', [])
        sg.user_settings_set_entry('-last filename-', '')
        window['-FILENAME-'].update(values=[], value='')

window.close()
