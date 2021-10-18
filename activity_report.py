import getpass
import logging
import logging.config
import os 
import sys
import time
import pandas as pd
import tms_login as tms
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By 
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select, WebDriverWait


def export_to_file(df, filename):
    export_filename = filename        
    filesave = f'{DOWNLOAD_FOLDER}{export_filename}.xlsx'
    sheetname = export_filename
    writer = pd.ExcelWriter(filesave, engine="xlsxwriter")
    df.to_excel(writer, sheet_name=sheetname)
    writer.save()
    print(f'Saved file {filesave}!')
    return filesave


if len(sys.argv) == 4:
    dept_select = sys.argv[1].upper()
    s_date_str = sys.argv[2]
    e_date_str = sys.argv[3]
elif 1 < len(sys.argv) < 4:
    print(f'Insufficient arguments. Entered: {len(sys.argv) - 1} Expected: 3')
    sys.exit()
elif len(sys.argv) > 4:
    print(f'Too many arguments. Entered: {len(sys.argv) - 1} Expected: 3')
    sys.exit()
else:
    # receive user input to set report date range
    dept_select = input('Choose department: (A)ccounting, (O)perations: ').upper()

    s_date_str = input('Enter start date as MMDDYY: ')
    while len(s_date_str) != 6:
        s_date_str = input('Invalid date. Enter start date as MMDDYY: ')

    e_date_str = input('Enter end date as MMDDYY: ')
    while len(e_date_str) != 6:
        e_date_str = input('Invalid date. Enter end date as MMDDYY: ')

# check if on citrx ('nt') or pi
if os.name == 'nt':
    # set to Chrome default download folder - BOA CITRIX DESKTOP DEFAULT SETTINGS
    DOWNLOAD_FOLDER = f"C:\\Users\\{getpass.getuser().title()}\\Downloads\\"
else:
    # constant for downloads folder - Pi default for Chromium
    DOWNLOAD_FOLDER = '/home/pi/Downloads/'

# initialize logger
logging.config.fileConfig(fname='logs/cfg/activity.conf')
logger = logging.getLogger('')

url = 'https://boa.3plsystemscloud.com/'
browser = tms.login(url, False)
browser.get(f'{url}App_BW/staff/manager/reportEmployeeSummary.aspx')

act_dict = { 
    "Mario Alanis": "3347",
    "Josie De La Torre": "1328",
    "Wilson Escalante": "3332",
    "Agnes Vanisi": "3670",
    "Alicia Napoles": "3739",
    "AP Temp": "3748",
    "Carla Moon": "3745"
}

ops_dict = {
    "Edgar Aguilar": "1260",
    "Dominique Burbank": "512",
	"Leslie Gallardo": "3392",
    "Yennie Lo": "2934",
    "Nancy Portillo": "3393",
    "Ruben Rivera": "3224",
    "Julie Salamanca": "2741",
    "Omar Salamanca": "986",
    "Dusty Sikes": "2865",
    "Hervey Vargas": "3231"
}

act_action_dict = {
    "AR Invoices Created": "invoice",
    "AP Invoices Received": "apstandard"
}

ops_action_dict = {
    "Orders Entered": "enter",
    "Carrier Selected": "carrierselect",
    "Dispatched": "dispatch",
    "Pickup Confirmed": "pickup",
    "Delivery Confirmed": "deliver"
}

## list of employee names
## use staffkey.json and pass employee as argument to loop instantiate class Employee 

# with open('staffkey.json', 'r') as f:
#     emp_lookup = json.load(f)
# emp_list =['Mario Alanis', 'Josie De La Torre', 'Wilson Escalante', 'Judy Li', 'Jessamyn Pentecost','Evelyn Reynoso']
# for employee in emp_list:
#     employee = Employee()


if dept_select == 'O':
    staff_dict = ops_dict
    action_dict = ops_action_dict
    export_prefix = 'ops-actions-'
elif dept_select == 'A':
    staff_dict = act_dict
    action_dict = act_action_dict
    export_prefix = 'acct-actions-'
else:
    logging.error('Must choose either (O)perations or (A)ccounting.')
    sys.exit()

empty_row = ['']
export_df = pd.DataFrame(empty_row)

for key in staff_dict:
    s_date_obj = datetime.strptime(s_date_str, '%m%d%y')
    e_date_obj = datetime.strptime(e_date_str, '%m%d%y') + timedelta(1)
    
    emp_name, emp_id = key, staff_dict[key]
    emp_select = Select(browser.find_element_by_id('ctl00_BodyContent_empList'))
    emp_select.select_by_value(emp_id)

    try:
        # enter each date in range individually to grab all actions for staff
        for eachdate in range((e_date_obj - s_date_obj).days):
            WebDriverWait(browser, timeout=30).until(EC.visibility_of_element_located((By.ID, 'ctl00_BodyContent_DateSelStart_tbxDatePicker')))
            startbox = browser.find_element_by_id('ctl00_BodyContent_DateSelStart_tbxDatePicker')
            endbox = browser.find_element_by_id('ctl00_BodyContent_DateSelEnd_tbxDatePicker')
            
            form_date = datetime.strftime(s_date_obj, '%m/%d/%Y')

            startbox.clear()
            startbox.send_keys(form_date)
            endbox.clear()
            endbox.send_keys(form_date)
            time.sleep(2)
            webdriver.ActionChains(browser).send_keys(Keys.ESCAPE).perform()
            time.sleep(2)
            webdriver.ActionChains(browser).send_keys(Keys.ESCAPE).perform()
            WebDriverWait(browser, timeout=30).until(EC.element_to_be_clickable((By.ID, "ctl00_BodyContent_Button1")))


            for key in action_dict:
                try:
                    before = os.listdir(DOWNLOAD_FOLDER)

                    action_name, action_id = key, action_dict[key]
                    action_select = Select(browser.find_element_by_id('ctl00_BodyContent_showDetail'))
                    action_select.select_by_value(action_id)

                    update_btn = browser.find_element_by_id('ctl00_BodyContent_Button1').click()

                    WebDriverWait(browser, timeout=30).until(EC.element_to_be_clickable((By.XPATH, "//div[@id='ctl00_BodyContent_ExportExcelDiv']/a")))
                    WebDriverWait(browser, timeout=30).until(EC.presence_of_element_located((By.XPATH, "//tbody/tr[@class='GridBFooter']/td[1]")))
                    total_loads = browser.find_element_by_xpath("//tbody/tr[@class='GridBFooter']/td[1]").text
                    
                    one_day = datetime.strftime(s_date_obj, '%m%d')
                    logging.info(f'{emp_name}: Total {total_loads} {action_name} on {one_day}!')
                    
                    download = browser.find_element_by_xpath("//div[@id='ctl00_BodyContent_ExportExcelDiv']/a").click()
                    time.sleep(3)

                    after = os.listdir(DOWNLOAD_FOLDER)
                    change = set(after) - set(before)

                    if len(change) == 1:
                        file_name = change.pop()
                        logging.info(f'{file_name} downloaded.')
                        
                        # sets filepath to downloaded file and create DataFrame from file 
                        # *output file extension is .xls but is actually.html format
                        filepath = f'{DOWNLOAD_FOLDER}{file_name}'
#                        os.rename(filepath, filepath+'.xls')


                        data = pd.read_html(filepath)
                        df = data[0]

                        col_emp = [emp_name for i in range(df.shape[0])]
                        col_date = [form_date for i in range(df.shape[0]) ]
                        col_action = [action_name for i in range(df.shape[0])]

                        df['Employee'] = col_emp
                        df['Date'] = col_date
                        df['Action'] = col_action

                        df_to_concat = [export_df, df]
                        export_df = pd.concat(df_to_concat, ignore_index=True)

                        new_source = f'{DOWNLOAD_FOLDER}{emp_name} - {action_name}{one_day}.html'
                        os.rename(filepath, new_source)
                        logging.info(f'Renamed {filepath} to {new_source}!')
                    elif len(change) == 0:
                        print ("No file downloaded")
                    else:
                        logging.info("More than one file downloaded.")
                except ElementClickInterceptedException:
                    pass
                except Exception as e:
                    logging.info('Exception thrown ' + repr(e))        
            s_date_obj += timedelta(1)
    except Exception as e:
        logging.info('Exception thrown: ' + repr(e))
        error_file = s_date_str + '_error'
        export_to_file(export_df, error_file)
        
export_filename = f'{export_prefix}{s_date_str}_{e_date_str}'        
export_filepath = export_to_file(export_df, export_filename)

# output variable filename to constant file, use for emailing purposes
og_stdout = sys.stdout

with open('activity_filepath.txt', 'w') as f:
    sys.stdout = f
    print(export_filepath)
    sys.stdout = og_stdout

browser.quit()

# GridBFooter <-- DOM class to wait for loading

# explore options on executing javascript directly
# javascript:openExcelTable('ctl00_BodyContent_DetailGrid','Employee Activity Detail Report');
