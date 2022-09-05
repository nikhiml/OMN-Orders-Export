from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import requests
import pandas

import smtplib
import mimetypes
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.message import Message
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.text import MIMEText

emailfrom = "******"
emailto = "*******"
fileToSend = "****.csv"
username = "****"
my_password = "*****"

chrome_driver_path = '*********************'
driver = webdriver.Chrome(executable_path=chrome_driver_path)
url = '*************'
uname = '*****'
pwd = '******'

sheety_endpoint = '***************'

# For getting the current order no 

sheet_response = requests.get(url=sheety_endpoint)
print(sheet_response.json())
sheet_order_no_wip = sheet_response.json()['orders'][-1]['sNo']
if '-' in sheet_order_no_wip:
    sheet_order_no = int(sheet_order_no_wip.split('-')[0])
else:
    sheet_order_no = int(sheet_order_no_wip)

current_order_no = sheet_order_no + 1

# Logging into orders system to check for all the new orders

driver.get(url)

time.sleep(3)

uname_field = driver.find_element(by=By.XPATH, value='/html/body/div/div/form/div/div[2]/input')
uname_field.send_keys(uname)

pwd_field = driver.find_element(by=By.XPATH, value='/html/body/div/div/form/div/div[3]/input[2]')
pwd_field.send_keys(pwd)

submit = driver.find_element(by=By.XPATH, value='/html/body/div/div/form/div/div[5]/input')
submit.click()

time.sleep(3)

new_order_flag = 'Y'

order_no_list = []

# Looping through the latest orders in the system to get the list of new orders

all_rows = driver.find_elements(by=By.CSS_SELECTOR, value='.hor-scroll tbody tr')
for row in all_rows:
    row_elements = row.find_elements(by=By.CSS_SELECTOR, value='td')
    if int(row_elements[1].text) >= int(current_order_no):
        order_no_list.append(int(row_elements[1].text))

max_order_no = max(order_no_list)

# Looping through the list of new orders to extract order level information

while current_order_no <= max_order_no:
    all_rows = driver.find_elements(by=By.CSS_SELECTOR, value='.hor-scroll tbody tr')

    for row in all_rows:
        row_elements = row.find_elements(by=By.CSS_SELECTOR, value='td')
        if int(row_elements[1].text) == current_order_no:
            row_status = row_elements[-2].text
            order_no = row_elements[1].text
            
# Only selecting the 'Processing' and 'Pending' orders, not the Cancelled ones

            if row_status == 'Processing' or row_status == 'Pending':
                row.click()
                time.sleep(3)
                order_no_temp = order_no
                address = driver.find_element(by=By.CSS_SELECTOR, value='.box-right address').text
                prices = driver.find_elements(by=By.CSS_SELECTOR, value='.order-totals tbody .price')
                channel_partner = 'OMN'
                order_price = driver.find_element(by=By.CSS_SELECTOR, value='.order-totals tfoot .price').text
                if (len(prices) == 3 or len(prices) == 4) and prices[2].text.split('₹')[0] == '-':
                    print('inside for partner order')
                    order_price = prices[0].text
                    shipment_type = 'NORMAL'
                    tbody_labels = driver.find_elements(by=By.CSS_SELECTOR, value='.order-totals tbody .label')
                    for label in tbody_labels:
                        if 'discount' in label.text.lower():
                            if 'nestery' in label.text.lower():
                                channel_partner = 'Nestery'
                            elif 'mbb' in label.text.lower():
                                channel_partner = 'MBB'
                            elif 'clouds' in label.text.lower():
                                channel_partner = 'Happy Clouds'
                            elif 'hamleys' in label.text.lower():
                                channel_partner = 'Hamleys'
                            elif 'natty' in label.text.lower():
                                channel_partner = 'Natty'
                            else:
                                channel_partner = 'OMN'

                elif len(prices) == 3 or len(prices) == 4:
                    shipping_price = prices[2].text.split('₹')[1]
                    if float(shipping_price) == 295.00:
                        shipment_type = 'URGENT / COD'
                    else:
                        shipment_type = 'COD'

                else:
                    shipping_price = prices[1].text.split('₹')[1]
                    if float(shipping_price) == 295.00:
                        shipment_type = 'URGENT'
                    else:
                        shipment_type = 'NORMAL'

                total_order_items_list = []

                no_of_order_items_even = driver.find_elements(by=By.CSS_SELECTOR, value='.hor-scroll .even .item-text')
                for order_item in no_of_order_items_even:
                    total_order_items_list.append(order_item)

                no_of_order_items_odd = driver.find_elements(by=By.CSS_SELECTOR, value='.hor-scroll .odd .item-text')
                for order_item in no_of_order_items_odd:
                    total_order_items_list.append(order_item)

                print(f"no of order items in the order: {len(total_order_items_list)}")

                order_item_counter = 0

                for order_item in total_order_items_list:
                    order_item_counter += 1
                    if len(total_order_items_list) > 1:
                        order_no = order_no_temp
                        order_no = order_no + '-' + f"{order_item_counter}"
                    pages = 28
                    orient = 'LANDSCAPE'
                    sku = order_item.find_element(by=By.CSS_SELECTOR, value='.item-text div').text.split(':')[1]
                    items_options = order_item.find_elements(by=By.TAG_NAME, value='dd')

                    items_list = []
                    for item_option in items_options:
                        items_list.append(item_option.text)
                    print(items_list)

                    print(sku)

                    # DIARY #

                    if 'journal' in sku or 'creativity' in sku or 'curious' in sku:
                        print('inside diary')
                        child_name = items_list[0]
                        child_gender = items_list[1].upper()
                        book_binding = 'HC STRIP BIND'
                        book_size = 'A5'
                        pages = 100
                        orient = 'Portrait'

                    # PHOTO & NAME PUZZLES & HEIGHT CHARTS #

                    elif 'puzzle' in sku.lower() or 'height' in sku.lower():
                        print('inside puzzle')
                        if 'name-puzzle' in sku:
                            child_name = items_list[0]
                            child_gender = 'NA'

                        elif 'photo' in sku:
                            child_name = items_list[1]
                            child_gender = items_list[2].upper()

                        else:
                            child_name = items_list[0]
                            child_gender = 'NA'
                            sku = sku + '-' + items_list[1]

                        book_size = 'NA'
                        pages = 'NA'
                        orient = 'NA'
                        book_binding = 'NA'
                        inside_paper = 'NA'
                        cover_paper = 'NA'

        #
        #           # BOOKS #
        #
                    elif 'first' in sku:
                        print('inside 1st Bday')
                        child_name = items_list[0]
                        child_gender = items_list[1].upper()
                        n = len(items_list)

                        if n == 6:
                            photo_option = 'YES'
                        else:
                            photo_option = 'No'

                        if 'A5' in items_list[n - 1]:
                            book_size = '7.5X6'
                            book_size_text = 'A5'
                        else:
                            book_size = '10X8.5'
                            book_size_text = 'A4'

                        if 'Board' in items_list[n - 1]:
                            book_binding = 'BACK 2 BACK HC'
                        elif 'Hard' in items_list[n - 1]:
                            book_binding = 'HC STRIP BIND'
                        else:
                            book_binding = 'SC STRIP BIND'

                    elif 'unicorn' in sku or 'sports' in sku or 'travel' in sku or 'super' in sku or 'unique' in sku or \
                            'vehicle' in sku or 'wish' in sku or 'world' in sku or 'birthday' in sku or 'diwali' in sku:
                        print('other books')
                        child_name = items_list[0]
                        child_gender = items_list[1].upper()
                        n = len(items_list)

                        if n == 5:
                            photo_option = 'YES'
                        else:
                            photo_option = 'No'

                        if 'A5' in items_list[n-1] or '7.5' in items_list[n-1]:
                            book_size = '7.5X6'
                            book_size_text = 'A5'
                        else:
                            book_size = '10X8.5'
                            book_size_text = 'A4'
                        if 'Board' in items_list[n-1]:
                            book_binding = 'BACK 2 BACK HC'
                        elif 'Hard' in items_list[n-1]:
                            book_binding = 'HC STRIP BIND'
                        else:
                            book_binding = 'SC STRIP BIND'
                        #
                    elif 'ABC' in sku or 'habits' in sku or 'shapes' in sku or 'counting' in sku or \
                                        'colours' in sku or 'fruits' in sku:
                        print('board books')
                        child_name = items_list[0]
                        child_gender = items_list[1].upper()
                        book_binding = 'BACK 2 BACK HC'
                        book_size = '210X148'
                        n = len(items_list)

                        if n == 4:
                            photo_option = 'YES'
                        else:
                            photo_option = 'No'

                    elif 'ohmyname' in sku:
                        print('OMN')
                        if 'Yes' in items_list[1] and 'Yes' in items_list[2]:
                            child_name = items_list[3]
                        elif 'Yes' in items_list[1]:
                            child_name = items_list[2]
                        else:
                            child_name = items_list[1]
                        child_gender = sku.split('-')[1].upper()

                        if 'Mini' in items_list[0]:
                            book_size = '7.5X6'
                            book_size_text = 'A5'
                        else:
                            book_size = '10.5X8'
                            book_size_text = 'A4'

                        if 'Hard' in items_list[0]:
                            book_binding = 'HC STRIP BIND'
                        else:
                            book_binding = 'SC STRIP BIND'

                    elif 'daddy' in sku:
                        print('daddy')
                        if items_list[0] == '1':
                            sku = 'Daddy and Me'
                            child_name = items_list[1]
                            child_gender = items_list[2].upper()
                            if 'Paper' in items_list[16]:
                                book_binding = 'SC STRIP BINDING'
                            elif 'Hard' in items_list[16]:
                                book_binding = 'HC STRIP BINDING'
                            else:
                                book_binding = 'BACK 2 BACK HC'
                            if 'A5' in items_list[16]:
                                book_size = '7.5X6'
                                book_size_text = 'A5'
                            else:
                                book_size = '10X8.5'
                                book_size_text = 'A4'
                    #
                    #         # daddy_dict = {
                    #         #     'ORDER_NO': order_no,
                    #         #     'FORMAT': items_list[16],
                    #         #     'GENDER': items_list[2],
                    #         #     'NOC': items_list[1],
                    #         #     'DADDY NAME': items_list[3],
                    #         #     'DADDY': items_list[4],
                    #         #     'CPN': items_list[5],
                    #         #     'PRO': items_list[6],
                    #         #     'VICHLE': items_list[7],
                    #         #     'COLOR': items_list[8],
                    #         #     'FP': items_list[9],
                    #         #     'FT': items_list[10],
                    #         #     'HOBBY': items_list[11],
                    #         #     'CLOTH': items_list[12],
                    #         #     'H WORK': items_list[13],
                    #         #     'FOOD': items_list[14],
                    #         #     'MESSAGE': items_list[15]
                    #         # }
                    #         #
                    #         # df = pandas.DataFrame(daddy_dict, index=[0])
                    #         # df.to_csv('*********.csv', mode='a')
                    #         #
                    #         # msg = MIMEMultipart()
                    #         # msg["From"] = emailfrom
                    #         # msg["To"] = emailto
                    #         # msg["Subject"] = "Daddy CSV File"
                    #         # msg.preamble = "Daddy CSV File"
                    #         #
                    #         #     ctype, encoding = mimetypes.guess_type(fileToSend)
                    #         #     if ctype is None or encoding is not None:
                    #         #         ctype = "application/octet-stream"
                    #         #
                    #         #     maintype, subtype = ctype.split("/", 1)
                    #         #
                    #         #     if maintype == "text":
                    #         #         fp = open(fileToSend)
                    #         #         # Note: we should handle calculating the charset
                    #         #         attachment = MIMEText(fp.read(), _subtype=subtype)
                    #         #         fp.close()
                    #         #     elif maintype == "image":
                    #         #         print('inside image option')
                    #         #         fp = open(fileToSend, "rb")
                    #         #         attachment = MIMEImage(fp.read(), _subtype=subtype)
                    #         #         fp.close()
                    #         #     elif maintype == "audio":
                    #         #         fp = open(fileToSend, "rb")
                    #         #         attachment = MIMEAudio(fp.read(), _subtype=subtype)
                    #         #         fp.close()
                    #         #     else:
                    #         #         fp = open(fileToSend, "rb")
                    #         #         attachment = MIMEBase(maintype, subtype)
                    #         #         attachment.set_payload(fp.read())
                    #         #         fp.close()
                    #         #         encoders.encode_base64(attachment)
                    #         #     attachment.add_header("Content-Disposition", "attachment", filename=fileToSend)
                    #         #     msg.attach(attachment)
                    #         #
                    #         #     server = smtplib.SMTP("smtp.mail.yahoo.com")
                    #         #     server.starttls()
                    #         #     server.login(user=emailfrom, password=my_password)
                    #         #     server.sendmail(emailfrom, emailto, msg.as_string())
                    #         #     server.quit()
                    #
                        else:
                            sku = 'DADDY AND US'
                    #
                    elif 'grandma' in sku:
                        print('grandma')
                        child_name = items_list[0]
                        child_gender = items_list[1].upper()

                        n = len(items_list)

                        if n == 14:
                            photo_option = 'YES'
                        else:
                            photo_option = 'No'

                        if 'A5' in items_list[n - 1]:
                            book_size = '7.5X6'
                            book_size_text = 'A5'
                        else:
                            book_size = '10X8.5'
                            book_size_text = 'A4'

                        if 'Board' in items_list[n - 1]:
                            book_binding = 'BACK 2 BACK HC'
                        elif 'Hard' in items_list[n - 1]:
                            book_binding = 'HC STRIP BIND'
                        else:
                            book_binding = 'SC STRIP BIND'
                    #
                    elif 'grandpa' in sku:
                        print('grandpa')
                        child_name = items_list[0]
                        child_gender = items_list[1].upper()

                        n = len(items_list)

                        if n == 15:
                            photo_option = 'YES'
                        else:
                            photo_option = 'No'

                        if 'A5' in items_list[n - 1]:
                            book_size = '7.5X6'
                            book_size_text = 'A5'
                        else:
                            book_size = '10X8.5'
                            book_size_text = 'A4'

                        if 'Board' in items_list[n - 1]:
                            book_binding = 'BACK 2 BACK HC'
                        elif 'Hard' in items_list[n - 1]:
                            book_binding = 'HC STRIP BIND'
                        else:
                            book_binding = 'SC STRIP BIND'

                    elif 'mummy' in sku:
                        print('mummy')
                        child_name = items_list[1]
                        child_gender = items_list[2].upper()

                        n = len(items_list)

                        if n == 17:
                            photo_option = 'YES'
                        else:
                            photo_option = 'No'

                        if 'A5' in items_list[n - 1]:
                            book_size = '7.5X6'
                            book_size_text = 'A5'
                        else:
                            book_size = '10X8.5'
                            book_size_text = 'A4'

                        if 'Board' in items_list[n - 1]:
                            book_binding = 'BACK 2 BACK HC'
                        elif 'Hard' in items_list[n - 1]:
                            book_binding = 'HC STRIP BIND'
                        else:
                            book_binding = 'SC STRIP BIND'

                    elif 'siblings' in sku or 'twins' in sku:
                        print('siblings, twins')
                        if 'siblings' in sku:
                            child_name = f'{items_list[0]}, {items_list[2]}'

                            if items_list[1].lower() == 'female' or items_list[1].lower() == 'girl':
                                child1_gender = 'ES'
                            else:
                                child1_gender = 'EB'

                            if items_list[3].lower() == 'female' or items_list[3].lower() == 'girl':
                                child2_gender = 'YS'
                            else:
                                child2_gender = 'YB'

                            child_gender = f'{child1_gender} - {child2_gender}'
                        else:
                            child_name = items_list[0]
                            child_gender = items_list[1]

                        n = len(items_list)

                        if n == 9:
                            photo_option = 'YES'
                        else:
                            photo_option = 'No'

                        if 'A5' in items_list[n - 1]:
                            book_size = '7.5X6'
                            book_size_text = 'A5'
                        else:
                            book_size = '10X8.5'
                            book_size_text = 'A4'

                        if 'Board' in items_list[n - 1]:
                            book_binding = 'BACK 2 BACK HC'
                        elif 'Hard' in items_list[n - 1]:
                            book_binding = 'HC STRIP BIND'
                        else:
                            book_binding = 'SC STRIP BIND'
                    #
                    if book_binding == 'BACK 2 BACK HC':
                        inside_paper = '300 GSM + 300 GSM'
                    else:
                        inside_paper = '170 GSM MATT'

                    if book_binding == 'SC STRIP BIND':
                        cover_paper = 'C1S 250 GSM'
                    else:
                        cover_paper = '170 GSM MATT'

                    path = f"E:\BACKUP-WORK\WORK\Oh My Name\TEMP\{sku}\{child_gender}\{book_size}"
                    
# Prepaing the order dictionary to update in the master Orders CSV 

                    sheety_params = {
                            'order': {
                                'sNo': order_no,
                                'jobName': child_name,
                                'gender': child_gender,
                                'bookName': sku,
                                'insideLamination': 'NO',
                                'coverLamination': 'GLOSS',
                                'size': book_size,
                                'insidePages': pages,
                                'p /L': orient,
                                'binding': book_binding,
                                'paperForInside': inside_paper,
                                'paperForCover': cover_paper,
                                'templatePath': path,
                                'orderNo': order_no,
                                'patner': channel_partner,
                                'shipment': shipment_type,
                                'price': order_price,
                                'address': address
                                }
                    }

                    sheet_response = requests.post(url=sheety_endpoint, json=sheety_params)

                sales = driver.find_element(by=By.XPATH, value='/html/body/div[2]/div[1]/div[3]/ul/li[2]/a/span')
                sales.click()
                time.sleep(2)
                main_page = driver.find_element(by=By.XPATH, value='/html/body/div[2]/div[1]/div[3]/ul/li[2]/ul/li[1]/a/span')
                main_page.click()
                time.sleep(5)

            else:
                sheety_params = {
                    'order': {
                        'sNo': order_no,
                        'jobName': 'CANCELLED',
                        'gender': 'CANCELLED',
                        }
                }

                sheet_response = requests.post(url=sheety_endpoint, json=sheety_params)

                    #

    # print('***************** NEXT ORDER *****************')

            current_order_no += 1

            break

#

driver.quit()
