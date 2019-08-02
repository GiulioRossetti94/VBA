# -*- coding: utf-8 -*-
"""
Created on Wed Feb 27 12:53:08 2019

@author: Giulio Rossetti
"""

import os
import pandas as pd
import numpy as np
import file_update

file = 'fee.txt'

data = pd.read_csv(file,sep='\t',header=0,names=['FEE','FEE2'],index_col=None)
data.drop(['FEE2'],axis=1,inplace=True)
data['Date'] = data.index
last_date = int(data.iloc[-1,1])
year = str(last_date)[:4]
month = str(last_date)[4:6]
day = str(last_date)[-2:]

root = r'Y:\Mobiliare\08 Finint Economia Reale Italia\03_Back Office\03_Controlli NAV'
for path, subdirs, files in os.walk(root):
    for name in files:
        name = os.path.join(path, name)
        try:
            if 'PATRIMONIALE_FERI_' in name and 'depo' not in name.lower() and 'old' not in name.lower():
                date_to_check = str((name[-15:-5].replace('.','')))
                date_to_check = int(date_to_check[-4:]+date_to_check[2:4]+date_to_check[:2])

                if date_to_check>last_date:
                    name_i = name
                    
                    read_file = pd.read_excel(name_i,sheet_name='BIL VER patrim',index_col=0,header=0)
                    try:
                        a_fee = read_file.loc['2A'][2]
                        d_fee = read_file.loc['2A'][1]
                        
                        if np.isnan( a_fee):
                             a_fee =0
                        if np.isnan( d_fee):
                             d_fee = 0
                        
                        fee = a_fee-d_fee
                        
                        if np.isnan(fee):
                            fee = 0
                        to_append = [fee,date_to_check]
                        if not(date_to_check in data.index):
                            data = data.append(pd.DataFrame(to_append,index=['FEE','Date'],columns=[date_to_check]).T)

                    except:
                        fee = 0

                    
        except:
            pass
        
data = data.reindex(columns=['Date','FEE'])
np.savetxt("fee.txt",data,header = '\t\t'.join(data.columns.tolist()),fmt='%9.2f',delimiter='\t')

file_update.main()
