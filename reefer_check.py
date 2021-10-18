import getpass
import json
import os
import pandas as pd
import sys
import time
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import tms_login as tms
from data_extract import open_sheet

# check if on citrx ('nt') or pi
if os.name == 'nt':
    # set to Chrome default download folder - BOA CITRIX DESKTOP DEFAULT SETTINGS
    DOWNLOAD_FOLDER = f"C:\\Users\\{getpass.getuser().title()}\\Downloads\\"
else:
    # constant for downloads folder - Pi default for Chromium
    DOWNLOAD_FOLDER = '/home/pi/Downloads/'

# list of files before downloading
before = os.listdir(DOWNLOAD_FOLDER)

# get all POs from Reefer List and Delivery Schedule
reefer_list = open_sheet('Reefer Check', 'rl_query').get_all_values()
del_sched = open_sheet('Reefer Check', 'ds_export').get_all_values()

df_rl = pd.DataFrame(reefer_list[2:], columns=reefer_list[1])
df_ds = pd.DataFrame(del_sched[1:], columns=del_sched[0])

# identify which POs are on RL but not on DS
rl_on_ds_check = df_rl['PO #'].isin(df_ds['PO'])

missing = df_rl[~rl_on_ds_check]
missing_loadnos = list(missing['Load #'])

url = 'https://boa.3plsystemscloud.com/'
browser = tms.login(url, False)
print('Logged into TMS.')

# enter report code into report_code variable
# "RF Missed Check" report
report_code = '7937B6793E2E'
report_url = f'{url}App_BW/staff/Reports/ReportViewer.aspx?code={report_code}'
browser.get(report_url)

loadsbox = browser.find_element_by_xpath("//tbody/tr[@id='table-wherevalue']/td[1]/input[@class='filter']")
loadsbox.clear()

for x in missing_loadnos[:-1]:
    loadsbox.send_keys(f'\'{x}\',')
loadsbox.send_keys(f'\'{missing_loadnos[-1]}\'')

# save & view report, then download
save_button = browser.find_element_by_id('ctl00_ContentBody_butSaveView')
save_button.click()
browser.implicitly_wait(3)
download = browser.find_element_by_id('ctl00_ContentBody_butExportToExcel')
download.click()
time.sleep(5)

browser.close()
print('Retrieved shipment report.')

# list of files in Downloads folder after downloading to extract filename
after = os.listdir(DOWNLOAD_FOLDER)
change = set(after) - set(before)

while len(change) == 0:
    print('No file downloaded.')
    after = os.listdir(DOWNLOAD_FOLDER)
    change = set(after) - set(before)

if len(change) == 1:
    file_name = change.pop()
    print(f'{file_name} downloaded.')
else:
    print('More than one file downloaded. Please check only one file gets downloaded.')
    sys.exit()

# output file extension is .xls but is actually.html format
filepath = f'{DOWNLOAD_FOLDER}{file_name}'
data = pd.read_html(filepath)
df_tms = data[0]

df_tms = df_tms[[
    'PO Reference', 'Load #', 'Carrier Name', 'S/ Status', 'Customer Name',  
    'Shipper', 'S/ State', 'Consignee', 'Pallets', 'Weight',
    'C/ City', 'C/ State', 'Pickup Date', 'Estimated Delivery Date'
]]

# REMOVE ROWS WITH CONSOL CARRIER ASSIGNED
df_tms = df_tms.drop(df_tms[df_tms['Carrier Name'] == 'CONSOLIDATION PROGRAM'].index)

# SHOW ONLY LOADS IN BOOKED OR QUOTED STATUS
df_tms_booked = df_tms[df_tms['S/ Status'] == 'Booked']
df_tms_booked.fillna(value='', inplace=True)
df_tms_quoted = df_tms[df_tms['S/ Status'] == 'Quoted']
df_tms_quoted.fillna(value='', inplace=True)
df_tms_export = pd.concat([df_tms_quoted, df_tms_booked], ignore_index=True)

print('Exporting DataFrame to HTML...')
df_tms_export.to_html('rfcheck-table.html', index=False)

head = open('./templates/head.html', 'r').read()
body = open('rfcheck-table.html', 'r').read()
tail = '</body>\n</html>'

html_list = [head, body, tail]

export = open('rf-check.html', 'wt')

full = '\n'.join(html_list)

export.write(full)
export.close()
print('Export HTML file saved!')
