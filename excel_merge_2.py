import pandas as pd
import numpy as np
import string
import os.path
#import sys

#temp = sys.argv[1]

f_pos = 'D:/CSV/ex/'
f_type = '.xlsx'
f_list_get = pd.read_excel(f_pos + 'KB계좌확인' + f_type, header=None)
f_list = f_list_get[0].values.tolist()
#f_list = f_list_get[4].values.tolist()
#del f_list[0]
#del f_list[-2:]

acc_name=[]
for a in f_list:
    acc_name.append(a)

tmp = pd.read_excel(f_pos + f_list[0] + f_type)
e_tmp1 = tmp.iloc[1::2,:]
e_tmp1.reset_index(drop=True , inplace=True)
e_tmp2 = tmp.iloc[0::2,:]
e_tmp2 = e_tmp2.rename(columns=e_tmp2.iloc[0])
e_tmp2 = e_tmp2.drop(e_tmp2.columns[0], axis=1)
e_tmp2 = e_tmp2.drop(e_tmp2.index[0])
e_tmp2.reset_index(drop=True , inplace=True)
result2 = pd.concat([e_tmp1,e_tmp2], axis=1)

result2['ACC_NUM'] = acc_name[0]

for i in range(1, len(f_list)):
    tmp = pd.read_excel(f_pos + f_list[i] + f_type)
    e_tmp1 = tmp.iloc[1::2,:]
    e_tmp1.reset_index(drop=True , inplace=True)
    e_tmp2 = tmp.iloc[0::2,:]
    e_tmp2 = e_tmp2.rename(columns=e_tmp2.iloc[0])
    e_tmp2 = e_tmp2.drop(e_tmp2.columns[0], axis=1)
    e_tmp2 = e_tmp2.drop(e_tmp2.index[0])
    e_tmp2.reset_index(drop=True , inplace=True)
    result1 = pd.concat([e_tmp1,e_tmp2], axis=1)
    
    result1['ACC_NUM'] = acc_name[i]
    
    result2 = pd.concat([result2,result1], ignore_index=True)

result2.to_excel(f_pos + 'merge_excel' + f_type)
