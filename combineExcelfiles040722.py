# puts multiple excel files as sheets in one workbook
import pandas as pd
import glob, os, datetime, time

start_time = time.time()
# Today's date
todayDate = ' ' + datetime.date.today().strftime("%d%m%Y")
print('Today is:' + todayDate)

# Files are in the desktop (Windows)
dest_dir = os.path.normpath(os.path.expanduser("~/Desktop"))

# 2 Excel files to merge as 2 different pages. You can add more with + 
excel_files = glob.glob(dest_dir +'/*' + todayDate + ' PYTHONtoSendTo_FTP.xlsx') + glob.glob(dest_dir + '/STOCK FILE' + todayDate + '*')
print(f'Files to merge: {excel_files} ')

# Destination file
destination = dest_dir + '\Stock Project' + todayDate + '.xlsx'
print(f'Destination: {destination} ')

writer = pd.ExcelWriter(destination,engine='xlsxwriter')

# Name of the sheets when saving
resultSheets = ["Microsoft","Oracle"]

x = 0
for excel_file in excel_files:
    sheet = resultSheets[x]
    print(f'Sheet {x}: {sheet} ')
    df1 = pd.read_excel(excel_file)
    df1.fillna(value='N/A', inplace=True)
    df1.to_excel(writer, sheet_name=sheet, index=False)
    x = x + 1

writer.save()
print(df1)

print(f'Done! Completed in {round(time.time()-start_time,2)} seconds.')