# -*- coding: utf-8 -*-

if __name__ != "__main__":
    print("""The program is defaulted to the file General Research Data folder path.
    Only enter the target_ticker. please update the file as necessary as time progresses""")
 
tgt_ticker='AAPL'
def Fin_Massager(tgt_ticker):
    import pandas as pd
    import glob
    import numpy as np
    if tgt_ticker is None:
        tgt_ticker = input("Please Enter a Stock Ticker: ")
        print("If the data returned does not meet expectations try again.")
    only_path = 'C:/Users/OBar/Documents/General Research/Data/'
   
    #C:/Users/OBar/Documents/General Research/Data/2015q1/pre.txt
    num_path = glob.glob(only_path + 'num\\*.txt')
    sub_path = glob.glob(only_path + 'sub\\*.txt')
    map_path = 'C:/Users/OBar/Documents/General Research/Data/cik_ticker.txt'
    ########
    #The tag file isn't necessary and can't be used with the normal import because it's read as if there are more columns than headers
    #######tag_path = 'C:/Users/OBar/Documents/General Research/Data/2015q1/tag.txt'
    map_tbl = pd.read_table(map_path, delimiter='|', header=0)
    num_complete = pd.DataFrame()
    sub_complete = pd.DataFrame()
  
    for idx,i in enumerate(sub_path):
        sub_tbl_idx = pd.read_table(i, delimiter='\t', header=0)
        sub_complete = sub_complete.append(sub_tbl_idx, ignore_index=True)
    
    #filter by Ticker
    tgt_tick = tgt_ticker
    
    map_filter = map_tbl.CIK[map_tbl.Ticker == tgt_tick]
    
    print(map_filter)
    
    #grab full year data adsh code, this will grab the year prior
    filt_sub = sub_complete[(sub_complete.cik == map_filter.iloc[0]) & (sub_complete.form == '10-K')]
    print('filt sub is:')
    print(filt_sub)
    
    try:
        the_adsh_gold = list(filt_sub.adsh)
    
    #Filter to the target companies financials.
        for idx,i in enumerate(num_path):
            num_tbl_idx = pd.read_table(i, delimiter='\t', header=0)
            num_tbl_filter = num_tbl_idx[num_tbl_idx['adsh'].isin(the_adsh_gold)]
            num_complete = num_complete.append(num_tbl_filter, ignore_index=True)
            
    except:
        print('no good')
    _z_Base_Query_Tbl = pd.merge(filt_sub[['name','cik','adsh']], num_complete, how='inner', on='adsh')

    
    """
    Now that we have our filtered financials tables we will transfer the results over to another program
    """
    return _z_Base_Query_Tbl
    
    list(filt_sub)