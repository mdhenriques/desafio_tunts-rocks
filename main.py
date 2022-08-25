import xlsxwriter #xlsxwriter to help me create and format the .xlsx file
import requests #requests module to get the api data
import re #I used regex module to format the 'Area'

import json #importing json so I can convert and manipulate json

url = "https://restcountries.com/v2/all?fields=name,capital,area,currencies"

response = requests.request("GET", url) 


data = json.loads(response.text)
del data[21:250]
#here i'm converting the json file so i can manipulate it easier and also deleting all the other countries from the list. Deleting wasn't really necessary, but it would make the process of looping easier
  

workbook = xlsxwriter.Workbook('CountriesList.xlsx')
worksheet = workbook.add_worksheet()
worksheet.set_column(0, 0, 20)
worksheet.set_column(1, 2, 15)
worksheet.set_column(3, 3, 10)
# creating the .xlsx file and the worksheet, using xlsxwriter, and setting the columns width

title = workbook.add_format({'bold': True, 'font_color': '#4F4F4F'})
title.set_font_size(16)
title.set_align('center')
#formatting the title. Setting it to BOLD, changing font color, font size and aligning it

headers = workbook.add_format({'bold': True, 'font_color': '#808080'})
headers.set_font_size(12)
#formatting the headers

worksheet.merge_range("A1:D1", "Countries List", title)
worksheet.write('A2', 'Name', headers)
worksheet.write('B2', 'Capital', headers)
worksheet.write('C2', 'Area', headers)
worksheet.write('D2', 'Currencies', headers)
#writing the sheet's headers and title

rowName = 2
colName = 0

rowCapital = 2
colCapital = 1

rowArea = 2
colArea = 2
  
rowCurrencies = 2
colCurrencies = 3
#I created these variables so I manage to control the row and collumn for each data type that I will insert on the sheet
  
for i in range(len(data)):
  worksheet.write(rowName, colName, data[i]['name'])
  rowName += 1
    #writes the names of all the countries in data list. Since I deleted the uneeded countries(after requesting the data), it will write only the ones i need
  
for i in range(len(data)):
  #writes capital names
    if 'capital' in data[i]:
        #checks if there is a capital on the country, if there isn't, writes a '-'
        worksheet.write(rowCapital, colCapital, data[i]['capital'])
    else:
        
        worksheet.write(rowCapital, colCapital, '-')
        
    rowCapital += 1
  
for i in range(len(data)):
  #writing and formatting area
    worksheet.write(rowArea, colArea, re.sub(r'(?<!^)(?=(\d{3})+$)', r'.', str(int(data[i]['area']))) + ',00')
  #formating the area value so it has '.' and ','
    rowArea += 1

for i in range(len(data)):
  if 'currencies' in data[i]: 
  #checks if the 'currencies' list existis in certain index
    if len(data[i]['currencies']) > 1: 
    #checks if there are more than 1 item inside 'currencies'
      for j in range(len(data[i]['currencies'])):
        if j < len(data[i]['currencies']):#cheks if the loop isn't on the last item yet
          worksheet.write(rowCurrencies, colCurrencies, data[i]['currencies'][j - 1]['code'] + ',' + data[i]['currencies'][j]['code'])
    else:
      worksheet.write(rowCurrencies, colCurrencies, data[i]['currencies'][0]['code'])
      rowCurrencies += 1
  else: 
    #if there isn't a 'currencies' list, a '-' will be written instead
    worksheet.write(rowCurrencies, colCurrencies, '-')
    rowCurrencies += 1

workbook.close()