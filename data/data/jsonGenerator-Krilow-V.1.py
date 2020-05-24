import xlrd
import simplejson as json
from collections import OrderedDict

# open workbook and select the first worksheet
workBook = xlrd.open_workbook('800-171-source.xlsx')
sheet1 = workBook.sheet_by_index(0)

# list for the dict
finalData = []

# iterate over each row in the worksheet, fetching the values into a dict
for row in range(1, sheet1.nrows):
    data = OrderedDict()
    values = sheet1.row_values(row)
    data['id'] = values[0] # grab values from columns 0-2
    data['title'] = values[1]
    data['text'] = values[2]
    finalData.append(data) # append each row's data to final data dict

# serialize the list of dicts to json format
finalJson = json.dumps(finalData)

# write output to .json file
with open('resultJSON.json', 'w') as f:
    f.write(finalJson)
