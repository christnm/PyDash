import os
import openpyxl

# Counting number of files
basepath = '\\\\thmahc\\shared\\Human Resources\\Benefits\\ESOP Temp\\1_2021 Applications\\Scanned'
fpasspath = '\\\\thmahc\\shared\\Human Resources\\Benefits\\ESOP Temp\\1_2021 Applications\\1st Pass'
spasspath = '\\\\thmahc\\shared\\Human Resources\\Benefits\\ESOP Temp\\1_2021 Applications\\2nd Pass'
duplipath = '\\\\thmahc\\shared\\Human Resources\\Benefits\\ESOP Temp\\1_2021 Applications\\Duplicate'
incompath = '\\\\thmahc\\shared\\Human Resources\\Benefits\\ESOP Temp\\1_2021 Applications\\Incomplete'
#scanned folder
os.chdir(basepath)
sc = 0
for file in os.listdir():
    if file.endswith(".pdf"):
        sc += 1
#1st pass folder
os.chdir(fpasspath)
fp = 0
for file in os.listdir():
    if file.endswith(".pdf"):
        fp += 1
fp = fp + sc
# 2nd pass
os.chdir(spasspath)
sp = 0
for file in os.listdir():
    sp += 1
# Duplicate
os.chdir(duplipath)
du = 0
for file in os.listdir():
    if file.endswith(".pdf"):
        du += 1
# Incomplete
os.chdir(incompath)
inc = 0
for file in os.listdir():
    if file.endswith(".pdf"):
        inc += 1
total = fp + sp + du + inc


excelpath = "C:\\Users\\cmorales\\OneDrive - American Health Partners, Inc\\AppsFeed.xlsx"

wbk = openpyxl.load_workbook(excelpath)
sheet = wbk['Sheet1']
sheet['A2'] = fp
sheet['B2'] = sp
sheet['C2'] = du
sheet['D2'] = inc
sheet['E2'] = total

wbk.save(excelpath)
wbk.close()
