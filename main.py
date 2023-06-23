import scipy
import numpy as np
import pandas as pd
from scipy.optimize import linear_sum_assignment
from time import process_time

import openpyxl

workbook1 = openpyxl.load_workbook('C:\\Users\\guner\\AppData\\Local\\Programs\\Python\\Python38\\write.xlsx')
sheet1= workbook1.active
workbook1.save('C:\\Users\\guner\\AppData\\Local\\Programs\\Python\\Python38\\write.xlsx')
sheet1['A1'] = result

df = pd.read_excel('dss.xlsm',sheet_name='Sheet4', usecols='H:AYA', skiprows=83)
cost_matrix = df.values

t1_start = process_time()
# apply Hungarian algorithm to get optimal assignments
row_ind, col_ind = linear_sum_assignment(cost_matrix)

t1_stop = process_time()

results=[]

total_distance=0
# print optimal assignments
product=[]
for i, j in zip(row_ind, col_ind):
    if i < 1312:
        total_distance += cost_matrix[i, j]
        result = f" {j+1} "
        results.append(result)
        print(result)
        product.append(i+1)

print("\n\n Total Distance:" +str(total_distance))

row_num = 0

for index, item in enumerate(results):
    worksheet.write_string(index, 0, item)
    row_num += 1
workbook1.close()
print("Elapsed time during the whole program in seconds:", t1_stop-t1_start)

#f"Product {i+1} is assigned to shelf {j+1} with a distance of {cost_matrix[i,j]}"