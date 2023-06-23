import scipy
import numpy as np
import pandas as pd
from scipy.optimize import linear_sum_assignment
from time import process_time
import xlsxwriter

workbook1 = xlsxwriter.Workbook('C:\\Users\\guner\\AppData\\Local\\Programs\\Python\\Python310\\write.xlsx')
worksheet = workbook1.add_worksheet()

df = pd.read_excel('C:\\Users\\guner\\AppData\\Local\\Programs\\Python\\Python310\\dss2.xlsm',sheet_name='model2', usecols='GH:HE', skiprows=1, nrows=24)
cost_matrix = df.values
#print(cost_matrix)
t1_start = process_time()
# apply Hungarian algorithm to get optimal assignments
row_ind, col_ind = linear_sum_assignment(cost_matrix)

t1_stop = process_time()

total_distance=0
# print optimal assignme=IF(OR($GC$8=0;GC28=0);999;=IF(OR($GC$8=0;GC28=0);999;=IF(OR($GC$8=0;GC28=0);999;nts
product=[]
results=[]
for i, j in zip(row_ind, col_ind):
        total_distance += cost_matrix[i, j]
        result = f"{j+1} "\
                 f"{cost_matrix[i,j]}"
        results.append(result)
        print(result)
        product.append(i + 1)
print("\n\n Total Time:" +str(total_distance))

row_num=0

for index, item in enumerate(results):
    worksheet.write_string(index, 0, item)
    row_num += 1

workbook1.close()
print("Elapsed time during the whole program in seconds:", t1_stop-t1_start)