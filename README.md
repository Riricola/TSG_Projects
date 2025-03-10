# TSG_Projects
All of the work I did while interning at TSG (Feb-March 2025)


# daily_discount.py

This script pulls PO's that are received from the day before, filtering for orders with discount codes. The script extracts the following information into excel format: ID#, branch, receive date, amount, payTo vendor, shipFrom vendor, terms code, and the proposed "due date" (not sure how they calculate that date). The resulting excel has an extra "disc Amt" column that calculates the total price of the order after the discount is applied,and sorts indescending order from largest value to smallest value. The excel sheet is titled "daily_discount.xlsx" and if it has been made already, the code adds a new sheet to it.

TODO: 
    I get an odd "pydantic.error_wrappers.ValidationError" when translating certain vendors, so some remain as their venderId in the excel.

# cod_report_format.py

This script
