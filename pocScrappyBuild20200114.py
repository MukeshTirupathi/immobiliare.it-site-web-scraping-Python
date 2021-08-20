from bs4 import BeautifulSoup
import bs4
import requests
import sys
from datetime import datetime
import xlsxwriter
import re


HEADERS = ({'User-Agent':'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/44.0.2403.157 Safari/537.36', 'Accept-Language': 'en-US, en;q=0.5'})
URL = input("Enter First Page URL : ").strip()
site = 'IMMOBILIARE.IT'

def getPropertyName(propertyDetail):
    return propertyDetail.findChildren("a")[0].contents[0].strip()


def removeFormatting(price):
    try:
        #price = price[:-4] if (price[-4:] == '.000' and len(price[:-4]) > 5) else price
        return re.sub('[^0-9]','', price)
        
    except:
        return ''


def stringToFloat(str):
    if len(str) > 0:
        return int(removeFormatting(str))
    else:
        return 0


def calculatePricePerSqMt(price, area):
    if (len(removeFormatting(price)) > 0) and (len(removeFormatting(area)) > 0):
           return str("{:.2f}".format(stringToFloat(price)/stringToFloat(area)))
    else:
           return ''


def getPrice(propertyDetail):
    try:
        #lif__item lif__pricing
        price = propertyDetail.findChildren("li", attrs={"class":'lnd-list__item in-feat__item in-feat__item--main'})[0]
        price = price.findChildren("div")
        if len(price) == 0:
            raise Exception()
        for p in price:
            if '€' in p.contents[0] and len(p.contents[0].strip()) > 0:
                return(p.contents[0].strip())
            else:
                raise Exception()
    except:
        try:
            price = propertyDetail.findChildren("li", attrs={"class":'lif__item lif__pricing'})[0]
            while isinstance(price, bs4.element.Tag):
                price = price.contents[-1]
            return price.strip()
        except:
            print('exception occured : getPrice')
            return ''


def getArea(propertyDetails):
    try:
        for detail in propertyDetails.findChildren("li", attrs={"class":'in-feat__data'}):
            if(len(detail.findChildren("sup")) > 0):
                return (detail.findChildren("span", attrs={"class":'text-bold'}))[0].contents[0]
    except:
        print('exception occured : getArea')
        return ''


def getSite():
    return 'IMMOBILIARE.IT'


def getCurrentMonth():
    today = datetime.today()
    switcher = { 
        1: "January",
        2: "February",
        3: "March",
        4: "April",
        5: "May",
        6: "June",
        7: "July",
        8: "August",
        9: "September",
        10: "October",
        11: "November",
        12: "December" 
    }
    return switcher.get(today.month, "") 


def hasNextPage(currentPage):
    if(len(currentPage.findChildren("ul", attrs={"class":'pull-right pagination'})) == 0):
        return False
    if(len(currentPage.findChildren("ul", attrs={"class":'pull-right pagination'})[0].findChildren("li",attrs={"class":'disabled'})) == 2):
        return False
    return True


def getNextPageURL(currentPage):
    try:
        if (len(currentPage.findChildren("ul", attrs={"class":'pull-right pagination'})[0].findChildren("a", attrs={"title":'Next page'})) > 0):
            result = currentPage.findChildren("ul", attrs={"class":'pull-right pagination'})[0].findChildren("a", attrs={"title":'Next page'})[0]['href']
        else:
            result = currentPage.findChildren("ul", attrs={"class":'pull-right pagination'})[0].findChildren("a", attrs={"title":'Pagina successiva'})[0]['href']
        return result
    except:
        print('exception occured : getNextPageURL')


def getAgency(propertyDetails):
    try:
        return propertyDetails.findChildren("div", attrs={"class":'nd-figure__image '
                                                                  'nd-ratio '
                                                                  'in-realEstateListCard__referent--image'
                                                          })[0].findChildren("img")[0]['alt']
    except:
        print('exception occured : getAgency')
        return ''


def generateSheet(rows):
    sheetName = datetime.now().strftime("%Y-%m-%d %H-%M") + '.xlsx'
    print(sheetName)
    workbook = xlsxwriter.Workbook(sheetName)
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})
    
    worksheet.write(0, 0, 'Name', bold)
    worksheet.write(0, 1, 'Price', bold)
    worksheet.write(0, 2, 'Area', bold)
    worksheet.write(0, 3, 'Price per sq. meter', bold)
    worksheet.write(0, 4, 'Site', bold)
    worksheet.write(0, 5, 'Month', bold)
    worksheet.write(0, 6, 'Agency', bold)
    
    row_count = 0
    for row in rows:
        worksheet.write(row_count+1, 0, row["property_name"])
        worksheet.write(row_count+1, 1, row["price"])
        worksheet.write(row_count+1, 2, row["area"])
        worksheet.write(row_count+1, 3, row["price_per_sq_meter"])
        worksheet.write(row_count+1, 4, row["site"])
        worksheet.write(row_count+1, 5, row["month"])
        worksheet.write(row_count+1, 6, row["agency"])
        row_count += 1

    workbook.close()


rows = []
pageCount = 0
while True:
    pageCount += 1
    webpage = requests.get(URL, headers=HEADERS)
    currentPage = BeautifulSoup(webpage.content, "lxml")
    ul = currentPage.find("ul", attrs={"id":'listing-container'})
    propertyDetails1 = ul.findChildren("li", recursive = False, attrs={"class":'listing-item listing-item--tiny js-row-detail'})
    propertyDetails2 = ul.findChildren("li", recursive = False, attrs={"class":'listing-item js-row-detail'})
    propertyDetails3 = ul.findChildren("li", recursive = False, attrs={"class":'listing-item listing-item--small js-row-detail'})
    propertyDetails4 = ul.findChildren("li", recursive = False, attrs={"class":'listing-item listing-item--medium js-row-detail'})
    propertyDetails5 = ul.findChildren("li", recursive = False, attrs={"class":'listing-item listing-item--wide js-row-detail'})
    for propertyDetail in propertyDetails1:
        rows.append({"property_name":getPropertyName(propertyDetail), "price":getPrice(propertyDetail), "area":getArea(propertyDetail), "price_per_sq_meter":calculatePricePerSqMt(getPrice(propertyDetail),getArea(propertyDetail)), "site":getSite(), "month":getCurrentMonth(), "agency":getAgency(propertyDetail)})
    for propertyDetail in propertyDetails2:
        rows.append({"property_name":getPropertyName(propertyDetail), "price":getPrice(propertyDetail), "area":getArea(propertyDetail), "price_per_sq_meter":calculatePricePerSqMt(getPrice(propertyDetail),getArea(propertyDetail)), "site":getSite(), "month":getCurrentMonth(), "agency":getAgency(propertyDetail)})
    for propertyDetail in propertyDetails3:
        rows.append({"property_name":getPropertyName(propertyDetail), "price":getPrice(propertyDetail), "area":getArea(propertyDetail), "price_per_sq_meter":calculatePricePerSqMt(getPrice(propertyDetail),getArea(propertyDetail)), "site":getSite(), "month":getCurrentMonth(), "agency":getAgency(propertyDetail)})
    for propertyDetail in propertyDetails4:
        rows.append({"property_name":getPropertyName(propertyDetail), "price":getPrice(propertyDetail), "area":getArea(propertyDetail), "price_per_sq_meter":calculatePricePerSqMt(getPrice(propertyDetail),getArea(propertyDetail)), "site":getSite(), "month":getCurrentMonth(), "agency":getAgency(propertyDetail)})
    for propertyDetail in propertyDetails5:
        rows.append({"property_name":getPropertyName(propertyDetail), "price":getPrice(propertyDetail), "area":getArea(propertyDetail), "price_per_sq_meter":calculatePricePerSqMt(getPrice(propertyDetail),getArea(propertyDetail)), "site":getSite(), "month":getCurrentMonth(), "agency":getAgency(propertyDetail)})
    
    print('Page '+ str(pageCount) + ' read successfully \n')

    if(not hasNextPage(currentPage)):
        print("Reached last page")
        print("Generating Spreadsheet")
        generateSheet(rows)
        break
    else:
        URL = getNextPageURL(currentPage)
