# -*- coding: utf-8 -*-
"""
Created on Wed Feb 27 16:18:43 2019

@author: GIULIO ROSSETTi
"""

import os
import pandas as pd
import numpy as np
import datetime as dt



def main():
    file = 'file.txt'
    file_ta = 'total_assets.txt'
    file_sub = 'subs.txt'
    file_ref = 'refunds.txt'

    data = pd.read_csv(file,sep='\t',header=0,index_col=None)
    cols = [x for x in data.columns.tolist() if not ('Unnamed' in x)]
    data = data.dropna(how='all',axis=1)
    data.columns = cols
    
    data_ta = pd.read_csv(file_ta,sep='\t',header=0,index_col=None)
    cols_ta = [x for x in data_ta.columns.tolist() if not ('Unnamed' in x)]
    data_ta = data_ta.dropna(how='all',axis=1)
    data_ta.columns = cols_ta    
    
    data_sub = pd.read_csv(file_sub,sep='\t',header=0,index_col=None)
    cols_sub = [x for x in data_sub.columns.tolist() if not ('Unnamed' in x)]
    data_sub = data_sub.dropna(how='all',axis=1)
    data_sub.columns = cols_sub  

    data_ref = pd.read_csv(file_ref,sep='\t',header=0,index_col=None)
    cols_ref = [x for x in data_ref.columns.tolist() if not ('Unnamed' in x)]
    data_ref = data_ref.dropna(how='all',axis=1)
    data_ref.columns = cols_ref

    
    last_date = int(data.iloc[-1,0])
#    year = str(last_date)[:4]
#    month = str(last_date)[4:6]
#    day = str(last_date)[-2:]
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
    
                        read_file = pd.read_excel(name_i,sheet_name='PATRIMONIALE',index_col=0,header=4,nrows = 27)
                        total_assets = read_file.loc[['TOTALE ' in s for s in read_file.index]].iloc[0,0]
                        cash = read_file.loc[['F. POSIZIONE' in s for s in read_file.index]].iloc[0,0]
#                        listed = read_file.loc[['A. STRUMENTI' in s for s in read_file.index]].iloc[0,0]
#                        not_listed =read_file.loc[['B. STRUMENTI' in s for s in read_file.index]].iloc[0,0]
                        debt_sec_list = read_file.loc[['A1' in s for s in read_file.index]].iloc[0,0]
                        debt_sec_nl = read_file.loc[['B1' in s for s in read_file.index]].iloc[0,0]
                        eqt = read_file.loc[['A2' in s for s in read_file.index]].iloc[0,0]
                        govi = read_file.loc[['A1.1' in s for s in read_file.index]].iloc[0,0]
                        etf = read_file.loc[['A3.' in s for s in read_file.index]].iloc[0,0]
                        other = read_file.loc[['G.' in s for s in read_file.index]].iloc[0,0]
                        date =name_i[-15:].strip('.xlsx').strip('_')
                        date = dt.datetime.strptime(date,'%d.%m.%Y')
    
                        b = [date_to_check,cash,debt_sec_list,debt_sec_nl,eqt,govi,etf,other]
                        ta = [date_to_check,total_assets]
                        
                        read_sub_ref= read_file = pd.read_excel(name_i,sheet_name='BIL VER patrim',index_col=0,header=0)
                        try:
                            a_wdraw = read_sub_ref.loc['4C'][2]
                            d_wdraw = read_sub_ref.loc['4C'][1]
                            
                            if np.isnan(a_wdraw):
                                 a_wdraw = 0
                            if np.isnan( d_wdraw):
                                 d_wdraw = 0
                                 
                            wdraw = a_wdraw - d_wdraw
                            if np.isnan(wdraw):
                                wdraw = 0
                                
                            wdraw_app = [date_to_check,wdraw]
                        except:
                            wdraw = 0 
                        try:
                            a_sub = read_sub_ref.loc['5A'][2]
                            d_sub = read_sub_ref.loc['5A'][1]
                            
                            if np.isnan( a_sub):
                                 a_sub = 0
                            if np.isnan( d_sub):
                                 d_sub = 0
                              
                            sub = a_sub-d_sub
                            if np.isnan(sub):
                                sub = 0
                                
                            sub_app = [date_to_check,sub]
                        except:
                            sub = 0
                        
    
                        if not(date_to_check in data.index):
                            data = data.append(pd.DataFrame(b,index=cols,columns=[data.index[-1]+1]).T)
                            data_ta = data_ta.append(pd.DataFrame(ta,index=cols_ta,columns=[data_ta.index[-1]+1]).T)
                            data_sub = data_sub.append(pd.DataFrame(sub_app,index=cols_sub,columns=[data_sub.index[-1]+1]).T)
                            data_ref = data_ref.append(pd.DataFrame(wdraw_app,index=cols_ref,columns=[data_ref.index[-1]+1]).T)
                        
            except:
                pass
    
    np.savetxt("file.txt",data,header = '\t\t'.join(data.columns.tolist()),fmt='%9.2f',delimiter='\t')
    np.savetxt("total_assets.txt",data_ta,header = '\t\t'.join(data_ta.columns.tolist()),fmt='%9.2f',delimiter='\t')
    np.savetxt("refunds.txt",data_ref,header = '\t\t'.join(data_ref.columns.tolist()),fmt='%9.2f',delimiter='\t')
    np.savetxt("subs.txt",data_sub,header = '\t\t'.join(data_sub.columns.tolist()),fmt='%9.2f',delimiter='\t')
    
if __name__ == '__main__':
    main()
    