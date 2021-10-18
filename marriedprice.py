from selenium import webdriver
from datetime import date, datetime, timedelta
import time
import os
import pandas as pd
import numpy as np
import sys

#from outlook_send import send_email

# constant to establish download folder path, only need to change this to change location
DOWNLOAD_FOLDER = '/home/pi/Downloads/'

# list of files before downloading
before = os.listdir(DOWNLOAD_FOLDER)

# activate chrome driver
browser = webdriver.Chrome()
browser.maximize_window()
browser.get("https://boa.3plsystemscloud.com/")

# page elements to login
boa_user = browser.find_element_by_id("txb-username")
boa_pw = browser.find_element_by_id("txb-password")
login_button = browser.find_element_by_id("ctl00_ContentBody_butLogin")

# login credentials
boa_user.send_keys("daigo@boalogistics.com")
boa_pw.send_keys("ship12345")
login_button.click()

# enter report code into report_code variable
# "BaseCost" report
report_code = "72A746731D3D"
url = "https://boa.3plsystemscloud.com/App_BW/staff/Reports/ReportViewer.aspx?code="+report_code
browser.get(url)

# sets up start date and end date for filter
#today = date.today()
#s_date = today
#e_date = today
#start = s_date.strftime("%m/%d/%Y 00:00:00")
#end = e_date.strftime("%m/%d/%Y 23:59:59")
s_date = sys.argv[1][:2] + '/' + sys.argv[1][2:]
e_date = sys.argv[2][:2] + '/' + sys.argv[2][2:]

#Manual date
# TODO HOW TO DO THIS IN BATCHES
start =f"{s_date}/2020 00:00:00"
end = f"{e_date}/2020 23:59:59"

customer = "Married"

# set up variables for parameter fields
startbox = browser.find_element_by_xpath("//td[8]/input[@class='filter between'][1]")
endbox = browser.find_element_by_xpath("//td[8]/input[@class='filter between'][2]")
customerbox = browser.find_element_by_xpath("//td[4]/input[@class='filter']")

# inserts new parameters
startbox.clear()
startbox.send_keys(start)
endbox.clear()
endbox.send_keys(end)
customerbox.clear()
customerbox.send_keys(customer)

#break

# save & view report, then download
save_button = browser.find_element_by_id("ctl00_ContentBody_butSaveView").click()
browser.implicitly_wait(3)
download = browser.find_element_by_id("ctl00_ContentBody_butExportToExcel").click()

#need to wait a few seconds before continuing to allow for file to finish downloading.

time.sleep(3)


#compares list of files in Downloads folder after downloading file to extract filename
after = os.listdir(DOWNLOAD_FOLDER)
change = set(after) - set(before)

if len(change) == 1:
    file_name = change.pop()
    print(file_name + " downloaded.")
else:
    print ("More than one file or no file downloaded")
   
# sets filepath to downloaded file and create DataFrame from file, grabs Load # column
filepath = f'{DOWNLOAD_FOLDER}{file_name}'
data = pd.read_html(filepath)
df = data[0]
df.fillna('',inplace=True)
#print(df)
load_list_full = df['Load #']

# removes last row of column and convert to string
last = len(load_list_full) - 1
load_list_int = load_list_full[0:last]
load_list = map(str, load_list_int)
#print(load_list)

#Create Shipment Notes File
Category = 'Base Cost'
fname = f'BaseCost_{sys.argv[1]}_{sys.argv[2]}.csv'
f = open(fname,'w',encoding="utf-8")
f.write(Category + ',Load,Notes')

#Grabs date and Time info from each load
for x in load_list:
    load_id = x
    print(load_id)
   
    #Shipment Notes Page
    shipment_url = 'http://boa.3plsystemscloud.com/App_BW/staff/shipment/shipmentNotes.aspx?showpop=0&loadid='+load_id
    browser.get(shipment_url)
   
    #Get number of table rows
    rows = browser.find_elements_by_xpath("//table/tbody/tr")
    row_length = len(rows)

    #Calculate the row number
    row_num_int = row_length - 15
    row_num = str(row_num_int).zfill(2)
    #print(row_num)

    basecost = 'N/A'

    while row_num_int > 0:
       
       
        #Find first shipment note entry
        row_num_str = str(row_num_int).zfill(2)
        table = browser.find_element_by_id("ctl00_BodyContent_RepeaterNotes_ctl" + row_num_str + "_TableRowItemTop")
        table2 = browser.find_element_by_id("ctl00_BodyContent_RepeaterNotes_ctl" + row_num_str + "_TableRowItemBottom")
        Table = table.text
        Table2 = table2.text
        #print(Table)
        #print(Table2)
        #print('\n')
        row_num_int -= 2
        #Specific Person
        person = Category
        action1 = 'Carrier cost changed from $0.00 to'
        action2 = 'The Do not update Customer invoices checkbox was turned on'
        filter1 = '$0.'
        if action1 in Table2 and action2 in Table2:
            split_to = Table2.split("to ",1)
            splitto = split_to[1]
            splitcost = splitto.split('\n',1)
            basecoststr = splitcost[0]
            basecost = basecoststr[0:len(basecoststr)-1]
            #print(basecost)
           
            if filter1 not in basecost:
                print(basecost)
                f.write('\n,"{}","{}"'.format(load_id, basecost))
               
               
    if basecost == 'N/A':
        print(basecost)
        f.write('\n,"{}","{}"'.format(load_id, basecost))
       

                                    
browser.quit()
f.close()
print('file saved!')


file = "BaseCost.csv"

#today = date.today()
#now = datetime.now()
#today_str = str(today)
#now_str = now.strftime("%I:%M %p")


#Pass arguments below in following order: To Address, Subject, Email Body, Path to file to attach

#send_email('sokchu@boalogistics.com, data@boalogistics.com,daigo@boalogistics.com,vince@boalogistics.com ',
#        'BaseCost Test ' + start[0:5] + ' - ' + end[0:10],
#        'Hello Team,\n\nAttached is the Base Cost file for all of 2020.\n\n\nThank you,\n\nSokchu Hwang',file)




