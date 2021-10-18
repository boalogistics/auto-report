import json
import sys
import time
from datetime import date, datetime, timedelta
import tms_login as tms

if len(sys.argv) < 2:
    print('Error: Expected employee name as argument')
    sys.exit()

today = date.today()

# json file for client owner settings
with open('cfg/owners.json', 'r') as f:
    owner_lookup = json.load(f)

owner_arg = sys.argv[1]
owner_data = owner_lookup[owner_arg]

client_ids = owner_data['clients']
email = owner_data['email']
sv = owner_data['sv']

RECEIPIENT = email
SUBJECTLINE = f'{today.strftime("%a %b %d")} Client Report - {owner_arg}'
MAILBODY = f'Client report for {today.strftime("%b %d")} attached.'

url = 'https://boa.3plsystemscloud.com/'
browser = tms.login(url)
PREFIX = 'ctl00_BodyContent_'

# enter report code into report_code variable
# "__Client Report" report
report_code = '1FF214214161'
report_url = f'{url}App_BW/staff/Reports/ReportViewer.aspx?code={report_code}'
report_email_url = f'{url}App_BW/staff/Reports/SendReport.aspx?code={report_code}'

browser.get(report_url)
print('TMS opened.')
time.sleep(5)

# check if today is Monday to determine if new date range needs to be entered
weekday = datetime.weekday(today)
if weekday == 0:
    s_date = today
    e_date = today + timedelta(7)

    start = s_date.strftime('%m/%d/%Y 00:00:00')
    end = e_date.strftime('%m/%d/%Y 23:59:59')

    startbox = browser.find_element_by_xpath("//table/tbody/tr[3]/td[1]/input[1]")
    endbox = browser.find_element_by_xpath("//table/tbody/tr[3]/td[1]/input[2]")

    startbox.clear()
    startbox.send_keys(start)
    endbox.clear()
    endbox.send_keys(end)

client_id_box = browser.find_element_by_xpath("//td[2]/input[@class='filter']")
client_id_box.clear()
for client_id in client_ids[:-1]:
    client_id_box.send_keys(f'\'{client_id}\',')
client_id_box.send_keys(f'\'{client_ids[-1]}\'')

save_btn = browser.find_element_by_id("ctl00_ContentBody_butSaveView")
save_btn.click()
time.sleep(5)

# input fields for email submission form
browser.get(report_email_url)
email = browser.find_element_by_id(f'{PREFIX}txbEmail')
subject = browser.find_element_by_id(f'{PREFIX}txbSubject')
body = browser.find_element_by_id(f'{PREFIX}txbBody')
send_btn = browser.find_element_by_id(f'{PREFIX}butSend')

email.send_keys(RECEIPIENT)
subject.send_keys(SUBJECTLINE)
body.send_keys(MAILBODY)
send_btn.click()
print(f'Email sent to Owner {owner_arg}.')
time.sleep(10)

# send to self for consolidated report
email = browser.find_element_by_id(f'{PREFIX}txbEmail')
email.clear()
email.send_keys('daigo@boalogistics.com')
send_btn = browser.find_element_by_id(f'{PREFIX}butSend')
send_btn.click()
print(f'Email sent to self.')
time.sleep(10)

# if client owner has a supervisor
if sv != 'n/a':
    email = browser.find_element_by_id(f'{PREFIX}txbEmail')
    subject = browser.find_element_by_id(f'{PREFIX}txbSubject')
    body = browser.find_element_by_id(f'{PREFIX}txbBody')
    send_btn = browser.find_element_by_id(f'{PREFIX}butSend')
    print(f'{sv}@boalogistics.com')
    email.clear()
    email.send_keys(f'{sv}@boalogistics.com')
    subject.send_keys(SUBJECTLINE)
    body.send_keys(MAILBODY)
    send_btn.click()
    print(f'Email sent to SV {sv}.')
    time.sleep(10)

browser.quit()
print('Browser closed.')
