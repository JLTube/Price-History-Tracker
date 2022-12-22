import time
from datetime import datetime, timedelta
import requests
from bs4 import BeautifulSoup
import gspread
from oauth2client.service_account import ServiceAccountCredentials

def AddtoGoogleSheet(Product_Link, Product_Name, Product_Cost, Delivey_Cost):
    
    # To get Current Date and time
    date_time = datetime.now()
    todate = date_time.strftime('%d-%b-%Y')
    timenow = date_time.strftime('%I:%M %p')

    # To get Current Date and time after 6 hours
    date_time_after = date_time+timedelta(hours=6)
    dateafter = date_time_after.strftime('%d-%b-%Y')
    timeafter = date_time_after.strftime('%I:%M %p')

    scopes = [ 
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    credentials = ServiceAccountCredentials.from_json_keyfile_name('amazon-price-tracker-370413-14110a865f22.json', scopes= scopes)
 
    file = gspread.authorize(credentials)
    workbook = file.open('Amazon Price Tracker')
    sheet = workbook.sheet1
    sheet.insert_row([], index= 2)
    sheet.update_acell('A2', 'Amazon IN')
    time.sleep(1)
    sheet.update_acell('B2', todate)
    time.sleep(1)
    sheet.update_acell('C2', timenow)
    time.sleep(1)
    sheet.update_acell('D2', f'=HYPERLINK("{Product_Link}","{Product_Name}")')
    time.sleep(1)
    sheet.update_acell('E2', Product_Cost)
    time.sleep(1)
    sheet.update_acell('F2', Delivey_Cost)
    time.sleep(1)
    sheet.update_acell('G2', '=SUM(E2+F2)')
    time.sleep(1)
    total_cost = sheet.acell('G2').value
    time.sleep(1)
    sheet.update_acell('G2', total_cost)
    time.sleep(1)
    sheet.update_acell('H2', timeafter)

    return (f'Recent update - {Product_Name}, {todate}, {timenow} || Next update - {dateafter}, {timeafter}')

def Amazon_IN(URL):

    
    headers = {"User-Agent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36 Edg/107.0.1418.62'}

    page = requests.get(URL, headers= headers)

    soup = BeautifulSoup(page.content, 'html.parser')

    InStock = soup.find(id="deliveryBlockMessage")
    if InStock is not None:

        AmazonIN_Product_Price = (soup.find("span", "a-offscreen").get_text()).strip()
        AmazonIN_Delivery_Charge = (soup.find(id="deliveryBlockMessage").get_text()).strip()

        AmazonIN_Product_Price = AmazonIN_Product_Price[1:]
        AmazonIN_Delivery_Charge = AmazonIN_Delivery_Charge[0:4]
        if AmazonIN_Delivery_Charge == 'FREE':
            AmazonIN_Delivery_Charge = 0

        return AmazonIN_Product_Price, AmazonIN_Delivery_Charge
    else:
        AmazonIN_Product_Price = "NA"
        AmazonIN_Delivery_Charge = "NA"
        return AmazonIN_Product_Price, AmazonIN_Delivery_Charge

def Get_Product_Code_Gsheet():
    Product_Code_Dic = {}

    ProductToScan_sheet = workbook.get_worksheet(3) 

    ProductToScan_sheet_values = ProductToScan_sheet.get_all_values() # Gets all data from the specified Googlesheet.

    Product_Code_Dic = {rows[1]: rows[2] for rows in ProductToScan_sheet_values[1:]} # From the raw data we are removing first column and adding reset to a empty dict.

    return Product_Code_Dic

def main():

    global workbook

    # Adding scopes to access Google Sheets
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]

    # Passing Credentials of Googsheet
    credentials = ServiceAccountCredentials.from_json_keyfile_name('amazon-price-tracker-370413-14110a865f22.json', scopes= scopes)

    # Accessing Googlesheet
    file = gspread.authorize(credentials)
    workbook = file.open('Amazon Price Tracker')

    AmazonIN_URL = "https://www.amazon.in/dp/"
    AmazonIN_product_code = Get_Product_Code_Gsheet()

    # AmazonIN_product_code = {
    #     'B09ZSNRX3B': 'Garmin Vivosmart 5 Large, Black',
    #     'B08N1C1GKJ': 'Garmin Venu SQ Grey',
    #     'B07HYX9P88': 'Garmin Instinct',
    #     'B08GLHDVJF': 'Garmin Instinct Solar',
    #     'B09TFNYB7M': 'Garmin Instinct 2',
    #     'B09TFNBRNF': 'Garmin Instinct 2 Solar'
    # }

    for code in AmazonIN_product_code:
        AmazonIN_Product_Price, AmazonIN_Delivery_Charge = Amazon_IN(AmazonIN_URL+code)
        time.sleep(10)
        if AmazonIN_Product_Price != 'NA':
            final_out = AddtoGoogleSheet(AmazonIN_URL+code, AmazonIN_product_code[code], AmazonIN_Product_Price, AmazonIN_Delivery_Charge)
            print(final_out)
            time.sleep(5)
            
    print('\n')
    time.sleep(21540)
    main()

if __name__ == '__main__':
    main()