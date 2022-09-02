import csv 
import xlrd
import numpy as np
from openpyxl import load_workbook
import time
from datetime import datetime as dt
import pandas as pd

x23 = []
cost = []
S23 = []

for i in range(5):
    with open('C:\pythonword/time'+str(i)+'.csv',newline='') as csvfile:
        rows = csv.DictReader(csvfile)
        for row in rows:
            cost.append(row['cost'])
    data = xlrd.open_workbook('C:\pythonword/work'+str(i)+'.xlsx')
    for n in range(len(data.sheet_names())):
        table = data.sheets()[n]
        #print(table.row_values(1)[:4])
        x23.append(table.row_values(1)[:4])

        
cost_high = cost.index(max(cost))
cost_low = cost.index(min(cost))

x_high = x23[cost_high]
x_low = x23[cost_low]
x_high[0] = int( x_high[0])
x_high[1] = int( x_high[1])
x_low[0] = int( x_low[0])
x_low[1] = int( x_low[1])

add = 0
add1 = 0
add2 = 0
add3 = 0
for j in x23:
    add +=j[0]
    add1 +=j[1]
    add2 +=j[2]
    add3 +=j[3]

S23.append(int(add))
S23.append(int(add1))
S23.append(add2)
S23.append(add3)


x_base0 = (np.array(S23)-np.array(x_high))/4
x_base = x_base0.tolist()
x_base[0] = int( x_base[0])
x_base[1] = int( x_base[1])
for k in range(5):   
    rand = np.random.random()
    alfa = 0.9+rand*0.2
    x23[k]=((x_base0+(x_base0-x_high)*alfa))
    x23[k][0] = int(x23[k][0])
    x23[k][1] = int(x23[k][1])
      
    for l in x23:
        wb=load_workbook('C:\pythonword/work'+str(k)+'.xlsx')
        sheet=wb["Sheet"]
        sheet['A2'] = l[0]
        sheet['B2'] = l[1]
        sheet['C2'] = l[2]
        sheet['D2'] = l[3]
        wb.save('C:\pythonword/work'+str(k)+'.xlsx')

    df = pd.read_csv('C:\pythonword/time'+str(k)+'.csv')
    time2 = str(dt.now().strftime('%Y%m%d%H%M%S'))+'\t'
    df.at[0,'time'] = time2
    df.at[0,'cost'] = np.random.random()
    print(str(dt.now().strftime('%Y%m%d%H%M%S')))
    time.sleep(1)
    df.to_csv('C:\pythonword/time'+str(k)+'.csv', index=False)



print()
print('x_base：',x_base)
print()
print('S23：',S23)
print()
print('Index_high：',cost_high) 
print('Index_low：',cost_low)
print()       
print('cost：',cost)
print()
print('x_high：',x_high)    
print('x_low：',x_low)  
print()
for l in x23:
   print([l])
