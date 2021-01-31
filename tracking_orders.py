import requests
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
import time
import datetime
from datetime import date



def trackingIdDecryption(trackingNumber):
    trackingIdEncryption = {'A': 'S', 'B': 'T', 'C': 'U', 'D': 'V', 'E': 'W', 'F': 'X', 'G': 'Y', 'H': 'Z', 'I': '%5B',
                            'J': '%5C', 'K': '%5D', 'L': '%5E', 'M': '_', 'N': '%60', 'O': 'a', 'P': 'b', 'Q': 'c',
                            'R': 'd', 'S': 'e', 'T': 'f', 'U': 'g', 'V': 'h', 'W': 'i', 'X': 'j', 'Y': 'k', 'Z': 'l',
                            '0': 'B', '1': 'C', '2': 'D', '3': 'E', '4': 'F', '5': 'G', '6': 'H', '7': 'I',
                            '8': 'J', '9': 'K'}
    return "".join([trackingIdEncryption[letter] for letter in trackingNumber])

def readExcelFile(startRow, endRow):
    # get handle on existing file
    wb = load_workbook(filename=os.environ['EXCEL_SHEET'])
    # get current month and year worksheet
    month_year = date.today().strftime("%b-%Y").split("-")
    # worksheet = "{month}. {year}".format(month=month_year[0], year=month_year[1])
    worksheet = 'Dec. 2020'
    # print(worksheet)
    ws = wb[worksheet]
    font = Font(name='Times New Roman', size=12)
    alignment = Alignment(horizontal='center')
    # get columns
    # December 2020
    shipping_status_col = 'O'
    time_col = 'P'
    tracking_number_col = 'M'
    # January 2021 and on
    # shipping_status_col = 'N'
    # time_col = 'O'
    # tracking_number_col = 'P'
    # loop through range values
    for i in range(startRow, endRow+1):
        cell = tracking_number_col + str(i)
        if isinstance(ws[cell].value, str):
            # track only orders that have not been delivered yet
            if ws[tracking_number_col + str(i)].value != 'Delivered':
                print('Tracking Number: ' + ws[cell].value)
                status = getShippingStatus(ws[cell].value)
                if status is not None: ws[shipping_status_col + str(i)] = status
                ws[time_col + str(i)] = datetime.datetime.now()
            else: continue
        else:
            ws[shipping_status_col + str(i)] = 'No Tracking Number'
            ws[shipping_status_col + str(i)] = 'No Tracking Number'
        ws[shipping_status_col + str(i)].font = font
        ws[shipping_status_col + str(i)].alignment = alignment
        ws[time_col + str(i)].font = font
        ws[time_col + str(i)].alignment = alignment
        print(ws[shipping_status_col + str(i)].value)
        print(ws[time_col + str(i)].value)

    wb.save(filename=os.environ['EXCEL_SHEET'])


def getShippingStatus(trackingNum):
    request_url = os.environ['TRACKING_API']
    url = os.environ['TRACKING_SITE'] + trackingNum
    trackingId = trackingIdDecryption(trackingNum)
    data = {
        'trackingId': trackingId,
        'carrier': 'Auto-Detect',
        'language': 'en',
        'country': 'Russian Federation',
        'platform': 'web-desktop',
        'wd': 'false',
        'c': 'false',
        'p': 3,
        'l': 2,
        'se': '1792x1024,MacIntel,Gecko,Mozilla,Netscape,Google Inc.,4g,Intel Inc.,Intel(R) UHD Graphics 630,undefined,103,3584,2048'}

    try:
        resp = requests.post(request_url, data=data)
        print(resp)
        print(resp.json())
        if 'error' in resp.json():
            print('Error: ' + resp.json()['error'].title())
            return 'Error: ' + resp.json()['error'].title()
        elif 'sub_status' in resp.json():
            print('Sub-status: ' + resp.json()['sub_status'].title())
            return resp.json()['sub_status'].title()
        else:
            print('Status: ' + resp.json()['status'].title())
            return resp.json()['status'].title()
    # check to see if request does not give back a valid json
    except ValueError:
        print('Too many requests, the request did not return any JSON')
        return None
    # except ConnectionError:
    #     print(ConnectionError)
    #     return None


if __name__ == '__main__':
    try:
        # startRow = int(input('Please input the first row you would like to track: '))
        # endRow = int(input('Please input the last row you would like to track up to: '))
        startRow = 706
        endRow = 750
        start = time.time()
        readExcelFile(startRow, endRow)
        end = time.time()
        # total time taken
        print(f"Runtime of the program is {end - start}")
    except ValueError as ve:
        print('Error Message: ' + str(ve))
        print('Please input valid integers for the start and end rows.')


