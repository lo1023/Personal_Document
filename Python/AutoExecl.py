import os
import openpyxl as xl

#__file__ 현재 수행중인 파이썬의 경로를 반환한다.
# 파이썬은 {}가 없는대신 들여쓰기 사용한다.
# if 나 for 쓰고 난 후에 들여쓰기 한칸 무조건

currentPath = os.path.dirname(__file__)
reportPath = os.path.join(currentPath,'data')


reports = [] #가변 배열 쩌..쩐당

for file in os.listdir(reportPath):
    if file.endswith(".xlsx") and "20200104" in file:
        filePath = os.path.join(reportPath,file)

        wb = xl.load_workbook(filePath)
        sheet = wb.active

        #3행부터 1열 - 5열
        row = 3
        col = 1

        #10번 반복
        while 1:
            name = sheet.cell(row = row, column = col).value
            if name is None:
                break
            content = sheet.cell(row = row, column = col+1).value
            reports.append({"name" : name,"content" : content})
            row +=1

print(reports)