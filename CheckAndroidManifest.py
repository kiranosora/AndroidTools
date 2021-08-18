import os
import xlwt
import xlrd
from ModifyExcel import *

sensitivePermissions=["android.permission.ACCESS_COARSE_LOCATION","android.permission.READ_CALENDAR","android.permission.ACCESS_FINE_LOCATION"]
sharedUserId='sharedUserId'
appPlatform="AppPlatform"

def modify(xlsxFileName, permissionList):
    wb = openpyxl.load_workbook(xlsxFileName)
    sheet = wb.get_sheet_by_name('Sheet1')
    idx=0
    rowIdx=1
    fill = openpyxl.styles.PatternFill("solid", fgColor="1874CD")
    for row in sheet.rows:
        if idx == len(permissionList):
            break
        rowValue=[cell.value for cell in row]
        print("rowValue: %s %s"%(rowValue, type(rowValue)))
        if permissionList[idx] in rowValue:
            flag = True
            for colIdx in range( rowValue.index(permissionList[idx]), len(rowValue)):
                if rowValue[colIdx] != None and '是' in rowValue[colIdx]:
                    flag = False
            if flag :
                sheet.cell(rowIdx, rowValue.index(permissionList[idx])+1).fill = fill
            idx = idx + 1
        rowIdx = rowIdx + 1
    resFileName=xlsxFileName[0:xlsxFileName.rfind('.')]+'_checked.xlsx'
    wb.save(resFileName)

def checkSensivitePermission(xmlFileName, xlsxFileName):
    xmlFile=None
    try:
        xmlFile=open(xmlFileName)
    except:
        print("opening file: %s failed"%xmlFileName)
        return
    hasPermission=[]
    lines = xmlFile.readlines()
    xmlFile.close()
    for permission in sensitivePermissions:
        for line in lines:
            if 'READ_CALENDAR' in permission:
                print('line: %s %s'%(line, permission in line))
            if permission in line:
                hasPermission.append(permission)
    modify(xlsxFileName, hasPermission)



if __name__=='__main__':
    checkSensivitePermission('AndroidManifest.xml', 'Book1.xlsx')
