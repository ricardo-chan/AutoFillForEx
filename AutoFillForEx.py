import sys
import requests
import xml.etree.ElementTree as ET
from openpyxl import load_workbook

def form_url(day: int, month: int) -> str :
    url = "http://www.floatrates.com/historical-exchange-rates.html?currency_date=2022-%02d-%02d&base_currency_code=EUR&format_type=xml" \
            % (month, day)

    return url

def load_finance(path: str):
    workbook = load_workbook(filename=path)

    return workbook

def fill_spreadsheet(workbook, sheet_choice: str, cell: str, amount: float):
    sheet = workbook[sheet_choice]
    sheet[cell] = amount

def save_finance(path: str, workbook):
    workbook.save(filename=path)

def form_cell_string(day: int, month: int) -> str :
    month_list = ["B", "E", "H", "K", "N", "Q", "T", "W", "Z", "AC", "AF", "AI"]

    cell = month_list[month - 1] + str(day + 4)

    return cell

def main():
    # Getting the url for specific days
    day = int(sys.argv[1])
    month = int(sys.argv[2])
    url = form_url(day, month)

    # Finance Spreadsheet path
    file = "<insert/your/path/here>"

    # Forming cell to fill out depending on date requested
    cell = form_cell_string(day, month)

    # Making a GET request
    r = requests.get(url)

    # check status code for response received
    # success code - 200
    if r.ok:
        root_node = ET.fromstring(r.text)

        # Loading finance spreadsheet
        workbook = load_finance(file)

        for tag in root_node.findall('item'):
            # Get the text of that tag
            target = tag.find('targetCurrency')

            # Filling out value on appropriate sheet
            if target.text == "HUF":
                exchange = tag.find('exchangeRate').text
                print("Writing HUF exchange of %s into sheet 'HUFEUR' cell %s" % (exchange, cell))
                fill_spreadsheet(workbook, "HUFEUR", cell, float(exchange))
            elif target.text == "MXN":
                exchange = tag.find('exchangeRate').text
                print("Writing MXN exchange of %s into sheet 'MXNEUR' cell %s" % (exchange, cell))
                fill_spreadsheet(workbook, "MXNEUR", cell, float(exchange))

        # Writing to spreadsheet
        save_finance(file, workbook)

        # Write to XML file
        # with open("output.xml", 'wb') as file:
        #     file.write(r.content)
    else:
        print("Server response failed!")


if __name__ == "__main__":
    main()
