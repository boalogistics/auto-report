import csv, os
import pandas as pd

#FOLDER = "C:\\Users\\daigo\\Desktop\\acctjanfeb\\"

FOLDER =  '/home/pi/Downloads/merge/'

filelist = os.listdir(FOLDER)

empty_row = ['']
export_df = pd.DataFrame(empty_row)

for file_name in filelist:
    # sets filepath to downloaded file and create DataFrame from file 
    # *output file extension is .xls but is actually.html format
    filepath = FOLDER + file_name
    
    data = pd.read_html(filepath)
    df = data[0]
    df['Filename'] = file_name

    df_to_concat = [export_df, df]
    export_df = pd.concat(df_to_concat, ignore_index=True)



export_filename = 'all'        
filesave = FOLDER + export_filename + ".xlsx"
sheetname = export_filename
writer = pd.ExcelWriter(filesave, engine="xlsxwriter")
export_df.to_excel(writer, sheet_name=sheetname)
writer.save()
print("Saved file " + filesave + "!")
