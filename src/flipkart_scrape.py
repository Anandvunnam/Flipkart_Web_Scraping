import urllib.request,urllib.error
from bs4 import BeautifulSoup as beauty
import ssl
import re
import openpyxl
from cell_settings import _init_cell_setting as set



# For websites that are having https: protocol
ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

while True:
    try:
        file_name = input('Enter Excel File Name\Path: ')
    except:
        print("Please enter valid file name(.xlsx)! ")
    try:
        work_book = openpyxl.load_workbook(file_name,read_only=False)
        break
    except:
        print("Please close "+file_name+"")

sheet = work_book['Sheet1']
sheet.protection.sheet = False
for cells in sheet.rows:
    for cell in cells:
        cell.font = set.font
        cell.alignment = set.alignment



# *****Dynamic URL for Flipkart searching*****

# flipkart_url = "https://www.flipkart.com/search?q={}&otracker=search&otracker1=search&marketplace=FLIPKART&as-show=on&as=off"
# search = input("Search For : ").strip()
# encrypted_string = urllib.parse.quote(search)
# flipkart_url = encrypted_string.join(flipkart_url.split("{}"))


flipkart_url  = "https://www.flipkart.com/search?q=mobile&otracker=search&otracker1=search&marketplace=FLIPKART&as-show=on&as=off"

try:
    html = urllib.request.urlopen(flipkart_url,context=ctx).read()
except:
    print("Please check your URL!")

soup = beauty(html,'html.parser')

product_container = soup.select("._4ddWXP")

product_name_selector = soup.select("._4ddWXP .s1Q9rs")
product_rating_selector = soup.select("._3LWZlK")
product_rating_review_selector = soup.select(".gUuXy- ._2_R_DZ")
product_current_price_selector = soup.select("._25b18c ._30jeq3")
product_original_price_selector = soup.select("._25b18c ._3I9_wc")

product_container_list = list()


for prod_i in range(len(product_container)):
    product_name = product_name_selector[prod_i].text
    product_current_price= product_current_price_selector[prod_i].text
    try:
        product_original_price = product_original_price_selector[prod_i].text
        product_rating = product_rating_selector[prod_i].text
        product_rating_review = re.search("[^()]", str(product_rating_review_selector[prod_i].text)).group()
    except:
        product_original_price = product_rating = product_rating_review = "No Data"

    product_container_list.append([product_name,product_current_price,product_original_price,
                                   product_rating,product_rating_review])


row = 2
for product_card in product_container_list:
    for j in range(len(product_card)):
        sheet.cell(row,j+1,value=product_card[j])
    row += 1


print("Saving to "+file_name)
work_book.save(file_name)
print("Data saved successfully!")


