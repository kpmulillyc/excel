import openpyxl,xlrd
import copy
import ctypes

filename = 'VOUCHER-PRINTING.xls'
workbook = xlrd.open_workbook(filename)
shit = workbook.sheet_by_index(0)
def countRecords():
    nr = shit.nrows
    rowxx = 0
    cel = shit.cell_value(rowxx, 2)
    tmppcel = shit.cell_value((rowxx+1),2)
    counter = 0
    for i in range(nr-2):
        if cel == tmppcel:
            rowxx += 1
            cel = shit.cell_value(rowxx,2)
            tmppcel = shit.cell_value((rowxx+1),2)
        else:
            counter+=1
            rowxx += 1
            cel = shit.cell_value(rowxx, 2)
            tmppcel = shit.cell_value((rowxx+1), 2)
    return counter
def addShit():
    temp = countRecords()
    if temp%2 is not 0:
        return int(temp/2 + 0.5)-1
    else:
        return int(temp/2)-1


def createTemplate():
    wb = openpyxl.load_workbook('template.xlsx')
    ws= wb.get_sheet_by_name('V0')
    fuck = copy.copy(ws)
    for i in range(addShit()):
        wb._add_sheet(fuck,i)
    wb.save('Voucher.xlsx')

createTemplate()
#manipulate data start from bleow
wb2 = openpyxl.load_workbook('Voucher.xlsx')
error=0
ws2Index = 6
newWs2Index = 28
sheetIndex = 0
entry = 0
counter = 0
rowIndex = 0
checkCell = shit.cell_value(rowIndex,2)
tmpcel = shit.cell_value(rowIndex+1,2)
ws2 = wb2.worksheets[sheetIndex]
while entry < countRecords():
    if counter == 0:
        ws2 = wb2.worksheets[0]
        companyName = shit.cell_value(0,6)
        ws2['B4'] = shit.cell_value(rowIndex, 0)
        ws2['D2'] = shit.cell_value(rowIndex, 2)
        ws2['B2'] = companyName
        summ = float(0)
        summ2 = float(0)
        while checkCell == tmpcel:
            ws2['A' + str(ws2Index)] = shit.cell_value(rowIndex, 1)
            ws2['B' + str(ws2Index)] = shit.cell_value(rowIndex, 3)
            ws2['D' + str(ws2Index)] = shit.cell_value(rowIndex, 4)
            ws2['E' + str(ws2Index)] = shit.cell_value(rowIndex, 5)
            ws2['A' + str(ws2Index+1)] = shit.cell_value(rowIndex+1, 1)
            ws2['B' + str(ws2Index+1)] = shit.cell_value(rowIndex+1, 3)
            ws2['D' + str(ws2Index+1)] = shit.cell_value(rowIndex+1, 4)
            ws2['E' + str(ws2Index+1)] = shit.cell_value(rowIndex+1, 5)
            if shit.row(rowIndex)[4].ctype == 2:
                summ += shit.row(rowIndex)[4].value
            if shit.row(rowIndex)[5].ctype == 2:
                summ2 += shit.row(rowIndex)[5].value
            ws2Index += 1
            rowIndex += 1
            checkCell = shit.cell_value(rowIndex,2)
            tmpcel = shit.cell_value(rowIndex+1,2)
            if checkCell != tmpcel:
                if shit.row(rowIndex)[4].ctype == 2:
                    summ += shit.row(rowIndex)[4].value
                if shit.row(rowIndex)[5].ctype == 2:
                    summ2 += shit.row(rowIndex)[5].value
                if summ != summ2:
                    error+=1
                ws2['D18'] = summ
                ws2['E18'] = summ2
                rowIndex += 1
                ws2Index = 6
                counter += 1
                entry += 1
                checkCell = shit.cell_value(rowIndex, 2)
                tmpcel = shit.cell_value(rowIndex + 1, 2)
                break
    elif counter==1:
        ws2['B26'] = shit.cell_value(rowIndex, 0)
        ws2['D24'] = shit.cell_value(rowIndex, 2)
        ws2['B24'] = companyName
        summ = float(0)
        summ2 = float(0)
        while checkCell == tmpcel:
            ws2['A' + str(newWs2Index)] = shit.cell_value(rowIndex, 1)
            ws2['B' + str(newWs2Index)] = shit.cell_value(rowIndex, 3)
            ws2['D' + str(newWs2Index)] = shit.cell_value(rowIndex, 4)
            ws2['E' + str(newWs2Index)] = shit.cell_value(rowIndex, 5)
            ws2['A' + str(newWs2Index + 1)] = shit.cell_value(rowIndex + 1, 1)
            ws2['B' + str(newWs2Index + 1)] = shit.cell_value(rowIndex + 1, 3)
            ws2['D' + str(newWs2Index + 1)] = shit.cell_value(rowIndex + 1, 4)
            ws2['E' + str(newWs2Index + 1)] = shit.cell_value(rowIndex + 1, 5)
            if shit.row(rowIndex)[4].ctype == 2:
                summ += shit.row(rowIndex)[4].value
            if shit.row(rowIndex)[5].ctype == 2:
                summ2 += shit.row(rowIndex)[5].value
            newWs2Index += 1
            rowIndex += 1
            checkCell = shit.cell_value(rowIndex, 2)
            tmpcel = shit.cell_value(rowIndex + 1, 2)
            if checkCell != tmpcel:
                if shit.row(rowIndex)[4].ctype == 2:
                    summ += shit.row(rowIndex)[4].value
                if shit.row(rowIndex)[5].ctype == 2:
                    summ2 += shit.row(rowIndex)[5].value
                if summ != summ2:
                    error += 1
                ws2['D40'] = summ
                ws2['E40'] = summ2
                rowIndex += 1
                newWs2Index = 28
                counter += 1
                sheetIndex += 1
                entry += 1
                checkCell = shit.cell_value(rowIndex, 2)
                tmpcel = shit.cell_value(rowIndex + 1, 2)
                break
    elif counter%2 is 0:
        ws2 = wb2.worksheets[sheetIndex]
        ws2['B4'] = shit.cell_value(rowIndex, 0)
        ws2['D2'] = shit.cell_value(rowIndex, 2)
        ws2['B2'] = companyName
        summ = float(0)
        summ2 = float(0)
        while checkCell == tmpcel:
            ws2['A' + str(ws2Index)] = shit.cell_value(rowIndex, 1)
            ws2['B' + str(ws2Index)] = shit.cell_value(rowIndex, 3)
            ws2['D' + str(ws2Index)] = shit.cell_value(rowIndex, 4)
            ws2['E' + str(ws2Index)] = shit.cell_value(rowIndex, 5)
            ws2['A' + str(ws2Index + 1)] = shit.cell_value(rowIndex + 1, 1)
            ws2['B' + str(ws2Index + 1)] = shit.cell_value(rowIndex + 1, 3)
            ws2['D' + str(ws2Index + 1)] = shit.cell_value(rowIndex + 1, 4)
            ws2['E' + str(ws2Index + 1)] = shit.cell_value(rowIndex + 1, 5)
            if shit.row(rowIndex)[4].ctype == 2:
                summ += shit.row(rowIndex)[4].value
            if shit.row(rowIndex)[5].ctype == 2:
                summ2 += shit.row(rowIndex)[5].value
            ws2Index += 1
            rowIndex += 1
            checkCell = shit.cell_value(rowIndex, 2)
            tmpcel = shit.cell_value(rowIndex + 1, 2)
            if checkCell != tmpcel:
                if shit.row(rowIndex)[4].ctype == 2:
                    summ += shit.row(rowIndex)[4].value
                if shit.row(rowIndex)[5].ctype == 2:
                    summ2 += shit.row(rowIndex)[5].value
                if summ != summ2:
                    error += 1
                ws2['D18'] = summ
                ws2['E18'] = summ2
                rowIndex += 1
                ws2Index = 6
                counter += 1
                entry += 1
                checkCell = shit.cell_value(rowIndex, 2)
                tmpcel = shit.cell_value(rowIndex + 1, 2)
                break
    else:
        ws2['B26'] = shit.cell_value(rowIndex, 0)
        ws2['D24'] = shit.cell_value(rowIndex, 2)
        ws2['B24'] = companyName
        summ = float(0)
        summ2 = float(0)
        while checkCell == tmpcel:
            ws2['A' + str(newWs2Index)] = shit.cell_value(rowIndex, 1)
            ws2['B' + str(newWs2Index)] = shit.cell_value(rowIndex, 3)
            ws2['D' + str(newWs2Index)] = shit.cell_value(rowIndex, 4)
            ws2['E' + str(newWs2Index)] = shit.cell_value(rowIndex, 5)
            ws2['A' + str(newWs2Index + 1)] = shit.cell_value(rowIndex + 1, 1)
            ws2['B' + str(newWs2Index + 1)] = shit.cell_value(rowIndex + 1, 3)
            ws2['D' + str(newWs2Index + 1)] = shit.cell_value(rowIndex + 1, 4)
            ws2['E' + str(newWs2Index + 1)] = shit.cell_value(rowIndex + 1, 5)
            if shit.row(rowIndex)[4].ctype == 2:
                summ += shit.row(rowIndex)[4].value
            if shit.row(rowIndex)[5].ctype == 2:
                summ2 += shit.row(rowIndex)[5].value
            newWs2Index += 1
            rowIndex += 1
            checkCell = shit.cell_value(rowIndex, 2)
            tmpcel = shit.cell_value(rowIndex + 1, 2)
            if checkCell != tmpcel:
                if shit.row(rowIndex)[4].ctype == 2:
                    summ += shit.row(rowIndex)[4].value
                if shit.row(rowIndex)[5].ctype == 2:
                    summ2 += shit.row(rowIndex)[5].value
                if summ != summ2:
                    error += 1
                ws2['D40'] = summ
                ws2['E40'] = summ2
                rowIndex += 1
                newWs2Index = 28
                counter += 1
                sheetIndex += 1
                entry += 1
                checkCell = shit.cell_value(rowIndex, 2)
                tmpcel = shit.cell_value(rowIndex + 1, 2)
                break

wb2.save('Voucher.xlsx')
if error > 0:
    ctypes.windll.user32.MessageBoxW(0, str(error)+' errors found', "Finished", 1)
else:
    ctypes.windll.user32.MessageBoxW(0, "No errors", "Finished", 1)

