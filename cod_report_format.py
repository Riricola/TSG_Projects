# TO run, put in console/command prompt: py .\cod_report_format.py 'C:\ProgramData\Epicor\SolarEclipse\Eclipse\unpaid.txt' 'C:\Users\Maria.Rodriguez\VSCode Projects\crackOUT.xlsx'
# SWITCH THE FILE PATHS TO 'input_txt_file_path.txt' 'desired_output_path.xlsx
import pandas as pd
import argparse

#! Insert code that connects to eterm, goes to accounting -> AR -> reports -> unbilled
#! branch -> TSGPLUMB -> payment terms: CCONFILE, COD, Cash

def main():
    parser = argparse.ArgumentParser(description="Process a text file and output an Excel sheet.")
    parser.add_argument('input_file', type=str, help='Path to the input text file')
    parser.add_argument('output_file', type=str, help='Desired name of the output Excel file')
    
    args = parser.parse_args()
    
 # Now you can use args.input_file and args.output_file in your script
    #input_file_path = r"C:\ProgramData\Epicor\SolarEclipse\Eclipse\unpaid.txt"
    #output_file_name = r"C:\Users\Maria.Rodriguez\VSCode Projects\crackOUT.xlsx"


    # Now you can use args.input_file and args.output_file in your script
    input_file_path = args.input_file
    output_file_name = args.output_file
    
    
    # Perform your formatting here
    formatted_data = format_data(input_file_path, output_file_name)  # Replace with your actual formatting function
    
    print("done")

def format_data(input_file_path, output_path):
        
    # Read the first two lines of the file
    with open(input_file_path, 'r') as file:
        for _ in range(4):  # Skip the first 4 lines
            file.readline()
        header_line = file.readline().strip()  # 5th line (headers)
        dashes_line = file.readline().strip()  # 6th line (dashes)

    # Calculate column widths based on the dashes line
    col_widths = [14, 34, 10, 14, 22, 14, 17, 14]
    # Read the fixed-width file into a DataFrame using the calculated widths
    data = pd.read_fwf(input_file_path, widths=col_widths, skiprows=4)

    df_cleaned = data.dropna(subset=["Invoice #"])

    # for each element in the column, if it is a negative number, move it to another sheet in the excel sheet
    sheet_name = "Sheet1"
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

    # Write both DataFrames back to the same Excel file
    with pd.ExcelWriter(output_path, engine="openpyxl", mode="w") as writer:
        df_cleaned.to_excel(writer, sheet_name=sheet_name, index=False)
        #? negatives_df.to_excel(writer, sheet_name="Credits", index=False)


if __name__ == "__main__":
    main()