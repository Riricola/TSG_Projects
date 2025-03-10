import datetime

import pydantic
from ecl_api import EclipseApi
import pandas as pd
from datetime import date
import openpyxl
import os


TEST_URL = "http://192.168.128.6:5000/"
TEST_CREDENTIALS = {"username": "API2", "password": "TRUSTNO1"}
yesterday = date.today() - datetime.timedelta(days=1)

def main():
    client = EclipseApi(username=TEST_CREDENTIALS["username"],
                        password=TEST_CREDENTIALS["password"],
                        url=TEST_URL)
    client.connect()
    #breakpoint()
    test_data = client.terms_search(keyword="", page_size=1000)

    terms_codes = test_data.search_results
    #breakpoint()

    discount_terms = {}
    for terms_code in terms_codes:
        if terms_code.vendorFlag:
            for period in terms_code.periods:
                try:
                    discount_float = float(period.discountPercentage)
                except (ValueError, TypeError):
                    discount_float = 0.0
                if discount_float != 0:
                    # Add the terms_code as the key and discountPercentage as the value
                    discount_terms[terms_code.id] = discount_float
                    break

    # List of discount terms (extract keys from the dictionary)
    terms_list = list(discount_terms.keys())

    #breakpoint()
    po_results = client._client.session.get("http://192.168.128.6:5000/PurchaseOrders",
                                            params={"ReceiveDate": "2025-02-10", "TermsCode": terms_list,
                                                    "OrderStatus": "Received", "pageSize": 1000,
                                                    }).json()['results']
    '''
    po_results = client._client.session.get("http://192.168.128.6:5000/PurchaseOrders",
                                            params={"ReceiveDate": yesterday, "TermsCode": terms_list,
                                                    "OrderStatus": "Received", "pageSize": 1000,
                                                    }).json()['results']
    '''

    df = pd.DataFrame({
        'ID': [item["eclipseOid"] for item in po_results],
        'Branch': [item["generations"][0]["priceBranch"] for item in po_results],
        'Date': [item["generations"][0]["receiveDate"] for item in po_results],
        'Amount': [item["generations"][0]["subtotalAmount"]["value"] for item in po_results],
        'Pay_to': [item["generations"][0]["payToId"] for item in po_results],
        'Ship_from': [item["generations"][0]["shipFromName"] for item in po_results],
        'Terms': [item["generations"][0]["termsCode"] for item in po_results],
        'Due': [item["generations"][0]["dueDate"] for item in po_results],
    })

    # Add a new column "discount%" to the DataFrame
    df["Discount%"] = df["Terms"].map(discount_terms)

    # Fill NaN values with 0 (or any default value) for terms not found in the dictionary
    df["Discount%"] = df["Discount%"].fillna(0.0)

    # translate vendor ID's to their proper name
    vendors = []
    for pay_id in df["Pay_to"]:
        try:
            sample = client.vendor_retrieve(vendor_id=pay_id).name
            vendors.append(sample)
        except pydantic.error_wrappers.ValidationError as e:
            # Log the error if needed
            #print(f"Validation error for vendor {pay_id}: {e}")
            # Skip this vendor or provide a default value
            sample = "Unknown"  # Or some default value
            vendors.append(pay_id)

    # save the list of vendor names to the dataframe
    df["Pay_to"] = vendors

    # Calculate the discounted price
    df["Disc Amt"] = df["Amount"] - (df["Amount"] * (df["Discount%"] / 100))

    # Sort received PO's by cash value
    df_sort = df.sort_values(by="Amount", ascending=False)

    # Excel file name
    file_name = "daily_discount.xlsx"
    try:
        if os.path.exists(file_name):
            # Load the existing workbook
            with pd.ExcelWriter(file_name, engine="openpyxl", mode="a", if_sheet_exists="new") as writer:
                # Create a new sheet (default name will be "SheetX")
                df_sort.to_excel(writer, sheet_name="Sheet" + yesterday.strftime("%m%d"))
        else:
            # Create a new Excel file if it doesn't exist
            df_sort.to_excel(file_name, index=False)

        print("Excel file updated successfully.")

    except Exception as e:
        print(f"An error occurred: {e}")

    # !: calculate aging through the excel sheet (receive date - current date)
    #TODO:
    # 1. send it out on a nightly basis (thru email?)
    # TODO: Add a try catch for excel sheet saving: if the excel exists, delete then resave

    breakpoint()

    client.disconnect()


if __name__ == '__main__':
    main()