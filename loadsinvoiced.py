import getpass
import os
import sys
import time
import pandas as pd
import tms_login as tms
from datetime import date, datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.keys import Keys


# check if on citrix ('nt') or pi
if os.name == 'nt':
    #set to chrome defualt download folder - BOA CITRIX DESKTOP DEFAULT SETTINGS
    DOWNLOAD_FOLDER = f'C:\\Users\\{getpass.getuser().title()}\\Downloads\\'
else:
    DOWNLOAD_FOLDER = '/home/pi/Downloads/' 

# list of files before downloading
before = os.listdir(DOWNLOAD_FOLDER)

today = date.today()

# set up start date and end date for filter
# if no args passed automatically use today
if len(sys.argv) < 2:
    s_date = today
    e_date = s_date
    datestamp = today
elif len(sys.argv) == 3:
    # takes args start and end dates as MMDD ex. 0420 = APR 20
    s_date = date(today.year, int(sys.argv[1][:2]), int(sys.argv[1][2:]))
    e_date = date(today.year, int(sys.argv[2][:2]), int(sys.argv[2][2:]))
    datestamp = e_date
else:
    print('Error: Need 2 arguments: start and end dates in MMDD format.')
    sys.exit()

start = s_date.strftime("%m/%d/%Y 00:00:00")
end = e_date.strftime("%m/%d/%Y 23:59:59")

url = 'https://boa.3plsystemscloud.com/'
browser = tms.login(url, False)

# enter report code into report_code variabe
# "zLoadsInvoiced" report
report_code = '7BD7D97CBE8E'
report_url = f'{url}App_BW/staff/Reports/ReportViewer.aspx?code={report_code}'

browser.get(report_url)
startbox = browser.find_element_by_xpath("//td[1]/input[@class='filter between'][1]")
endbox = browser.find_element_by_xpath("//td[1]/input[@class='filter between'][2]")

startbox.clear()
startbox.send_keys(start)
endbox.clear()
endbox.send_keys(end)

# save & view report, then download
save_button = browser.find_element_by_id('ctl00_ContentBody_butSaveView')
save_button.click()
browser.implicitly_wait(3)
download = browser.find_element_by_id('ctl00_ContentBody_butExportToExcel')
download.click()
time.sleep(3)

browser.close()

print('Retrieved invoice report.')

# list of files in Downloads folder after downloading to extract filename
after = os.listdir(DOWNLOAD_FOLDER)
change = set(after) - set(before)

if len(change) == 1:
    file_name = change.pop()
    print(f'{file_name} downloaded.')
elif len(change) == 0:
    print('No file downloaded.')
else:
    print('More than one file downloaded.')

#file name contains .xls extention but is actually html format
filepath = f'{DOWNLOAD_FOLDER}{file_name}'
data = pd.read_html(filepath)
df = data[0]

df.drop(axis=1, columns=['R/ Lumper', 'C/ Lumper', 'R/ Processing Fee', 'C/ Processing Fee'], inplace=True)
formatting = {'AR Balance': '${:,.2f}'}
df.style.format(formatting)
df.fillna('', inplace=True)
# df.drop(df.tail(1).index, inplace=True)
print(df)

print('Exporting DataFrame to HTML...')
df.to_html('invoiced-table.html', index=False)

head = open('./templates/head.html', 'r').read()
body = open('invoiced-table.html', 'r').read()
tail = '</body>\n</html>'

html_list = [head, body, tail]

export = open('invoiced.html', 'wt')

full = '\n'.join(html_list)

export.write(full)
print('Export HTML file saved!')

print('Exporting TXT file with load numbers...')
load_list_int = list(df['Load #'])[:-1]
load_list_str = [str(load_num) for load_num in load_list_int]

load_nos_txt = open('load_nos.txt', 'w')
output = "['" + "', '".join(load_list_str) + "']"
load_nos_txt.write(output)


print('Export TXT file saved!')

export.close()
load_nos_txt.close()
