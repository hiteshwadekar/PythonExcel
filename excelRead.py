import xlrd
import sys
from xlrd import open_workbook
from datetime import time

from xlwt import Workbook
from xlutils.copy import copy

wb = open_workbook('//home//gr-dev//Downloads//150529_AirS_CIG_COLL.xlsx')
sheet = wb.sheet_by_index(0)

rb = copy(wb)
r_sheet = rb.get_sheet(0)


print sheet.cell_value(0,2)
print sheet.cell_value(0,10)

print sheet.cell_value(1,2)
print sheet.cell_value(1,10)


#print sheet.ncols

curentHourValue=0
curentMinuteValue=0
curentSecValue=0
currentPMValue=0
counterForAvg=0
lastSecValue=0

print sheet.cell_value(9902,2)
cell = sheet.cell(9901,2)
print "cell type", cell.ctype
print xlrd.XL_CELL_DATE


for row in range(1,sheet.nrows):
    exceltime =  sheet.cell_value(row,2)

    print exceltime
    print row
    time_tuple = xlrd.xldate_as_tuple(exceltime,wb.datemode)
    print time_tuple
    time_value = time(*time_tuple[3:])

    if row == 1:
        curentHourValue = time_value.hour
        curentMinuteValue = time_value.minute
        curentSecValue = time_value.second

    if time_value.hour == curentHourValue:
        if time_value.minute == curentMinuteValue:
            counterForAvg +=1
            currentPMValue = currentPMValue + sheet.cell_value(row,10)
        else:
            currentPMValue = currentPMValue / counterForAvg
            # call to update the excelsheet (current to till counter)
            # update the value for new
            # reset the current values to new one
            print " Currernt PM value is -> ", currentPMValue , "for Avg ->  ", counterForAvg,"in the date range of ->  ", curentHourValue, curentMinuteValue, " with minute range of  ", curentSecValue, lastSecValue
            counterForAvg = 0
            currentPMValue = 0
            curentHourValue = time_value.hour
            curentMinuteValue = time_value.minute
            curentSecValue = time_value.second
            currentPMValue = currentPMValue + sheet.cell_value(row,10)
            counterForAvg += 1
    else:
        currentPMValue = currentPMValue / counterForAvg
        #call to update the excel sheet
        # update the value for new
        print " Currernt PM value is -> ", currentPMValue ,"for Avg ->  ", counterForAvg, "in the date range of ->  ", curentHourValue, curentMinuteValue, " with minute range of  ", curentSecValue, lastSecValue
        counterForAvg = 0
        currentPMValue = 0
        curentHourValue = time_value.hour
        curentMinuteValue = time_value.minute
        curentSecValue = time_value.second
        currentPMValue = currentPMValue + sheet.cell_value(row,10)
        counterForAvg += 1
    lastSecValue = time_value.second
    print row