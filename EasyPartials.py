#!/usr/bin/python3
# Author: Anthony Garrett
#
# Small script that will read in inventory location data from one spreadsheet
# and transfer that data to another spreadsheet
#
import os
from openpyxl import Workbook, load_workbook

ORIGINAL_INPUT = "EXPORT.XLSX"
DATA_INPUT = "products.xlsx"


def main():

    # Loading the full pallet quantities from the product spreadsheet
    data_wb = load_workbook(DATA_INPUT)
    data_sheet = data_wb['Sheet1']

    data_products = data_sheet["A"]
    data_full_quantities = data_sheet["B"]
    data_products_dict = {}

    for i in range(len(data_products)):
        data_products_dict[data_products[i].value] = int(
            data_full_quantities[i].value)

    # Loading required data from totals spreadsheet

    wb = load_workbook(ORIGINAL_INPUT)
    sheet = wb['Sheet1']

    # Deleting the first row which contains headers for the columns
    sheet.delete_rows(1, 1)

    location_codes = sheet["A"]
    products = sheet["B"]
    actual_quantity = sheet["D"]
    handling_unit = sheet["G"]

    partials_list = []

    # Loop through actual_quantity list to compare actual quantity to total pallet quantity
    # of the specified product from data_products_dict
    for i in range(len(actual_quantity)):
        if actual_quantity[i].value < data_products_dict[int(products[i].value)]:
            partials_list.append(
                [location_codes[i].value, products[i].value, handling_unit[i].value, actual_quantity[i].value])

    out_book = Workbook()
    out_sheet = out_book.active

    aisle_number = location_codes[0].value.split("-")[0]
    OUTPUT_FILENAME = "Aisle-" + aisle_number + "-partials" + ".xlsx"

    partial = ['15-101-A', '11007639', '376130426105723', 84]

    # Setting Up headers for the spreadsheet
    out_sheet["A" + str(1)] = "Storage Bin"
    out_sheet["B" + str(1)] = "Product"
    out_sheet["C" + str(1)] = "Handling Unit"
    out_sheet["D" + str(1)] = "Quantity"

    for count, partial in enumerate(partials_list, start=2):
        out_sheet["A" + str(count)] = partial[0]
        out_sheet["B" + str(count)] = partial[1]
        out_sheet["C" + str(count)] = partial[2]
        out_sheet["D" + str(count)] = partial[3]

    out_book.save(OUTPUT_FILENAME)
    # os.remove(ORIGINAL_INPUT)


if __name__ == "__main__":
    main()
