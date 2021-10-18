import csv, json, os, time
import pandas as pd
import tms_login as tms
from datetime import date
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By 
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select, WebDriverWait


# receive user input to set report date range
s_date = input('Enter start date as MMDDYY: ')
e_date = input('Enter end date as MMDDYY: ')
start = s_date[0:2] + '/' + s_date[2:4] + '/20' + s_date[4:6]
end = e_date[0:2] + '/' + e_date[2:4] + '/20' + e_date[4:6]

url = 'https://boa.3plsystemscloud.com/'
browser = tms.login(url)
browser.get('http://boa.3plsystemscloud.com/App_BW/staff/manager/reportEmployeeSummary.aspx')

WebDriverWait(browser, timeout=30).until(EC.visibility_of_element_located((By.ID, 'ctl00_BodyContent_DateSelStart_tbxDatePicker')))
startbox = browser.find_element_by_id('ctl00_BodyContent_DateSelStart_tbxDatePicker')
endbox = browser.find_element_by_id('ctl00_BodyContent_DateSelEnd_tbxDatePicker')
startbox.clear()
startbox.send_keys(start)
endbox.clear()
endbox.send_keys(end)
time.sleep(3)
webdriver.ActionChains(browser).send_keys(Keys.ESCAPE).perform()
time.sleep(3)

staff_dict = { 
	"Mario Alanis": "3347",
	"Josie De La Torre": "1328",
	"Wilson escalante": "3332",
    "Judy Li": "3122",
    "Jessamyn Pentecost": "2392",
    "Evelyn Reynoso": "3356"
}

action_dict = {
  	"Invoice Created": "invoice",
	"AP Invoices Received": "apstandard"
}

for key in staff_dict:
    emp_name = key
    emp_id = staff_dict[key]
    emp_select = Select(browser.find_element_by_id('ctl00_BodyContent_empList'))
    emp_select.select_by_value(emp_id)

    for key in action_dict:
        DOWNLOAD = 'C:\\Users\\daigo\\Downloads'

        before = os.listdir(DOWNLOAD)
        action_name = key
        action_id = action_dict[key]
        action_select = Select(browser.find_element_by_id('ctl00_BodyContent_showDetail'))
        action_select.select_by_value(action_id)

        update_btn = browser.find_element_by_id('ctl00_BodyContent_Button1')
        update_btn.click()

        WebDriverWait(browser, timeout=30).until(EC.element_to_be_clickable((By.XPATH, "//div[@id='ctl00_BodyContent_ExportExcelDiv']/a")))
        WebDriverWait(browser, timeout=30).until(EC.presence_of_element_located((By.XPATH, "//tbody/tr[@class='GridBFooter']/td[1]")))
        total_loads = browser.find_element_by_xpath("//tbody/tr[@class='GridBFooter']/td[1]").text
        
        print(emp_name + ": Total " + total_loads + " " + action_name + "!")
        
        download = browser.find_element_by_xpath("//div[@id='ctl00_BodyContent_ExportExcelDiv']/a")
        download.click()
        time.sleep(3)

        after = os.listdir(DOWNLOAD)
        change = set(after) - set(before)

        if len(change) == 1:
            file_name = change.pop()
            print(file_name + " downloaded.")
            
            # sets filepath to downloaded file and create DataFrame from file 
            # *output file extension is .xls but is actually.html format
            filepath = DOWNLOAD + "\\" + file_name
            data = pd.read_html(filepath)
            df = data[0]
    
            filesave = DOWNLOAD + "\\" + emp_name + " - " + action_name + ".xlsx"
            sheetname = action_id
    
            writer = pd.ExcelWriter(filesave, engine="xlsxwriter")
            df.to_excel(writer, sheet_name=sheetname)
            writer.save()
            print("Saved file " + filesave + "!")
            
            new_source = DOWNLOAD + "\\" + emp_name + " - " + action_name + ".html"
            os.rename(filepath, new_source)
            print("Renamed " + filepath + " to " + new_source + "!")
        elif len(change) == 0:
            print ("No file downloaded")
        else:
            print("More than one file downloaded.")


# GridBFooter <-- class to wait for loading

# explore options on executing javascript directly
# javascript:openExcelTable('ctl00_BodyContent_DetailGrid','Employee Activity Detail Report');
#


#        print(load_id)
#        print(carrier_info)
#        with open('lcr-carrier-list.csv', mode='a+') as carrier_list:
#            carrier_writer = csv.writer(carrier_list, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
#            carrier_writer.writerow([load_id, carrier_info])
#
#
#
#

#
##compares list of files in Downloads folder after downloading file to extract filename
#after = os.listdir(r"C:\Users\BOA\Downloads")
#change = set(after) - set(before)
#
#if len(change) == 1:
#    file_name = change.pop()
#    print(file_name + " downloaded.")
#else:
#    print ("More than one file or no file downloaded")
#    
## sets filepath to downloaded file and create DataFrame from file 
## *output file extension is .xls but is actually.html format
#
#filepath = r"C:\Users\BOA\Downloads" + "\\" + file_name
#data = pd.read_html(filepath)
#df = data[0]
#
## grabs list of load numbers and load count, dropping the Totals row
#load_list_numbers = list(df['Load #'])[:-1]
#load_list = [str(x) for x in load_list_numbers]
#load_count = len(df.index) -1
#
#print(load_list)
#print(str(load_count) + ' loads entered today.')
#
#
## gets today's date and adds it to new file name then saves as .xlsx file
#
#output_path = r"C:\Users\boa.sokchu\Downloads\Local_Bobtail" + "\\" + "REVISED" + file_name + ".xlsx"
#
#writer = pd.ExcelWriter(output_path, engine="xlsxwriter")
#df.to_excel(writer, sheet_name='DAILYREPORT')
#writer.save()
#print("Saved file to " + output_path + "!")
#
#
#with open('lcr-carrier-list.csv', mode='a+') as carrier_list:
#    carrier_writer = csv.writer(carrier_list, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
#    carrier_writer.writerow([load_id, carrier_info])
#



        
