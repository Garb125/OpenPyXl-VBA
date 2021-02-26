import openpyxl
from openpyxl import Workbook
import time
from time import localtime, strftime


month = strftime("%m",localtime())
day = strftime("%d", localtime())
year = strftime("%Y", localtime())
datestamp = f"{month}_{day}_{year}"

testWB = openpyxl.load_workbook(filename = "VBATest3.xlsm",keep_vba = True)
sheets = testWB.worksheets
newSheet = testWB.copy_worksheet(sheets[0])
newSheet.title = datestamp
testWB.active = newSheet
testWB.worksheets
#testWB.create_sheet(datestamp,0)

##add the difference from the set to last row of bank names
## change cell background or text or both to red 

testWB.save("VBATest3.xlsm")