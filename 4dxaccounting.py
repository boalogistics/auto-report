import tms_login as tms
import time

url = 'https://boa.3plsystemscloud.com/'
browser = tms.login(url)
PREFIX = 'ctl00_BodyContent_'

RECEIPIENT = 'daigo@boalogistics.com'
SUBJECTLINE = '4DX Accounting TMS'
MAILBODY = 'Attached.'

# enter report code into report_code variable
# "zzActActivity" report
report_code = '41A405436636'
report_url = f'{url}App_BW/staff/Reports/SendReport.aspx?code={report_code}'
browser.get(report_url)
print('TMS opened.')

email = browser.find_element_by_id(f'{PREFIX}txbEmail')
subject = browser.find_element_by_id(f'{PREFIX}txbSubject')
body = browser.find_element_by_id(f'{PREFIX}txbBody')
send_btn = browser.find_element_by_id(f'{PREFIX}butSend')

email.send_keys(RECEIPIENT)
subject.send_keys(SUBJECTLINE)
body.send_keys(MAILBODY)
send_btn.click()
print('Email sent.')
time.sleep(10)
browser.quit()
print('Browser closed.')

# TODO implement selenium element wait
# ctl00_BodyContent_lblMessage
# <span id="ctl00_BodyContent_lblMessage" style="color:Blue;">Send Successfully</span>