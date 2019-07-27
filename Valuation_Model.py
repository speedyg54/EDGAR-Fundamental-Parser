# -*- coding: utf-8 -*-
"""
Created on Wed Dec 19 08:09:56 2018

@author: OBar
"""

import sys
import pandas as pd
import numpy as np

#importing the items relevant to output
from __future__ import print_function
from mailmerge import MailMerge
import datetime

sys.path.append("C:\\Users\\OBar\\Documents\\General Research\\Scripts")
print(sys.path)
import Data_Massager_Imp as DMI
tgt_ticker = 'AAPL'


New_Tbl = DMI.Fin_Massager(tgt_ticker)
New_Tbl2 = New_Tbl[['ddate','value','tag']][ (New_Tbl.qtrs == np.int64(4)) | (New_Tbl.qtrs == np.int64(0))].set_index('ddate',drop=False)

New_Tbl2.to_excel('C:\\Users\\OBar\\Documents\\General Research\\Output\\test %s.xlsx' % tgt_ticker)

#Get the Annual Revenue - the parent company filter may not work for all
Revs = New_Tbl2.value[(New_Tbl2.tag == 'Revenues') | (New_Tbl2.tag == 'SalesRevenueNet')].drop_duplicates().sort_index(0)
print(Revs)
Revs2 = pd.DataFrame(data=Revs)
Revs3 = Revs2.T
F_Revs = Revs3.rename({'value' : 'Revenues' },axis='index')

print(F_Revs)

#get the operating income
OPINC = New_Tbl2.value[(New_Tbl2.tag == 'OperatingIncomeLoss')].drop_duplicates()

print(OPINC)
OPINC2 = pd.DataFrame(data=OPINC)
OPINC3 = OPINC2.T
F_OPINC = OPINC3.rename({'value' : 'Operating Income'},axis='index')

#Get the NetIncome

NETINC = New_Tbl2.value[(New_Tbl2.tag == 'NetIncomeLoss')].drop_duplicates()

print(NETINC)
NETINC2 = pd.DataFrame(data=NETINC)
NETINC3 = NETINC2.T
F_NETINC = NETINC3.rename({'value' : 'Net Income'},axis='index')

#Get the ROA

Inctaxpd = New_Tbl2.value[(New_Tbl2.tag == 'IncomeTaxesPaidNet') ].drop_duplicates()
IntExp = New_Tbl2.value[(New_Tbl2.tag == 'InterestExpense') ].drop_duplicates()
LTDcurr = New_Tbl2.value[(New_Tbl2.tag == 'LongTermDebtCurrent')].drop_duplicates()
LTDNC = New_Tbl2.value[(New_Tbl2.tag == 'LongTermDebtNoncurrent')].drop_duplicates()
ShE = New_Tbl2.value[(New_Tbl2.tag == 'StockholdersEquity') ].drop_duplicates().sort_index(0)

#Some manual massaging
def first_last_sas(dfs, operation):
    print("second parameter is 'first' or 'last'")
    dfs = dfs[~dfs.index.duplicated(keep=operation)]
    return dfs


Inctaxpd = first_last_sas(Inctaxpd, 'last')
IntExp = first_last_sas(IntExp, 'last')
LTDcurr = first_last_sas(LTDcurr, 'last')
LTDNC = first_last_sas(LTDNC, 'last')
ShE = first_last_sas(ShE, 'last')

print(NETINC)
print(Inctaxpd)
print(IntExp)

print(LTDcurr)
print(LTDNC)
print(ShE)

stp1 = NETINC.add(Inctaxpd, fill_value=0)
stp2 = stp1.add(IntExp, fill_value=0)

ostp1 = LTDcurr.add(LTDNC,ShE,fill_value=0)
ostp2 = ostp1.add(ShE, fill_value=0)

ROA = stp2 / ostp2

print(ROA)
ROA2 = pd.DataFrame(data=ROA)
ROA3 = ROA2.T
F_ROA = ROA3.rename({'value' : 'ROA'},axis='index')

#Get the ROE
ROE = NETINC/ShE
ROE2 = pd.DataFrame(data=ROE)
ROE3 = ROE2.T
F_ROE = ROE3.rename({'value' : 'ROE'},axis='index')


#Get the EPS
EPS = New_Tbl2.value[(New_Tbl2.tag == 'EarningsPerShareBasic') ].drop_duplicates()
EPS2 = pd.DataFrame(data=EPS)
EPS3 = EPS2.T
F_EPS = EPS3.rename({'value' : 'EPS'},axis='index')

#Get the EBITDA%
EBITDAperc = stp2 / Revs
EBITDAperc2 = pd.DataFrame(data=EBITDAperc)
EBITDAperc3 = EBITDAperc2.T
F_EBITDAperc = EBITDAperc3.rename({'value' : 'EBITDA%'},axis='index')

#GEt the OP%
OPperc = OPINC / Revs
OPperc2 = pd.DataFrame(data=OPperc)
OPperc3 = OPperc2.T
F_OPperc = OPperc3.rename({'value' : 'OP%'},axis='index')

#Get the NI%
NIperc = NETINC / Revs
NIperc2 = pd.DataFrame(data=NIperc)
NIperc3 = NIperc2.T
F_NIperc = NIperc3.rename({'value' : 'NI%'},axis='index')

#Get the Dividend per share
DivdpShare = New_Tbl2.value[(New_Tbl2.tag == 'CommonStockDividendsPerShareDeclared') ].drop_duplicates()
DivdpShare2 = pd.DataFrame(data=DivdpShare)
DivdpShare3 = DivdpShare2.T
F_DivdpShare = DivdpShare3.rename({'value' : 'Divd/Share'},axis='index')

#Get the Current Ratio
Curr_Ratio = New_Tbl2.value[(New_Tbl2.tag == 'AssetsCurrent')].drop_duplicates() / New_Tbl2.value[(New_Tbl2.tag == 'LiabilitiesCurrent') ].drop_duplicates()
Curr_Ratio2 = pd.DataFrame(data=Curr_Ratio)
Curr_Ratio3 = Curr_Ratio2.T
F_Curr_Ratio = Curr_Ratio3.rename({'value' : 'Curr Ratio'},axis='index')

#Get the FCFF
ncwrkcp1 = New_Tbl2.value[(New_Tbl2.tag == 'AssetsCurrent') ].drop_duplicates() - New_Tbl2.value[(New_Tbl2.tag == 'LiabilitiesCurrent') ].drop_duplicates()
#Note the last value, if a large negative is because the shift doesn't have a value in the most recent data.
noncash_wrkcap = ncwrkcp1.subtract(ncwrkcp1.shift(), fill_value=0)

cap1 = New_Tbl2.value[(New_Tbl2.tag == 'PropertyPlantAndEquipmentGross') ].drop_duplicates()
#Note the last value, if a large negative is because the shift doesn't have a value in the most recent data.
capex =  cap1.subtract(cap1.shift(), fill_value=0)

D_and_A = New_Tbl2.value[(New_Tbl2.tag == 'DepreciationAndAmortization')].drop_duplicates()

FCFF = stp2*(NETINC/stp1) + D_and_A - capex - noncash_wrkcap
FCFF2 = pd.DataFrame(data=FCFF)
FCFF3 = FCFF2.T
F_FCFF = FCFF3.rename({'value' : 'FCFF'},axis='index')
"""
Now that we have those we need to perform the DCF by finding the WACC
"""



#output dat ass
last_frames = [F_Revs,F_OPINC,F_OPperc,F_EBITDAperc,F_NETINC,F_NIperc,F_ROA,F_ROE,F_EPS,F_DivdpShare,F_Curr_Ratio, F_FCFF]
Final_Ass = pd.concat(last_frames)

HeaderList = list(Final_Ass.columns.values)
HeadListFinal = []
for i in HeaderList:
    HeadListFinal.append(str(i)[0:4])

col_rename_dict = {i:j for i,j in zip(HeaderList,HeadListFinal)}

Final_Ass.rename(columns=col_rename_dict, inplace=True)

#Filter to the most recent years
now_year = datetime.datetime.now().year
tgt_years = []
i = 1
while i < 5:
    temp = now_year - i
    tgt_years.append(str(temp))
    i += 1
    print(tgt_years)
tgt_years.reverse()
Final_Ass = Final_Ass.filter(tgt_years)


#begin pushing to Word doc
template_1 = "C:\\Users\\OBar\\Documents\\General Research\\Output\\Test_one.docx"
doc_1 = MailMerge(template_1)
print("Fields included in {}: {}".format(template_1,
                                         doc_1.get_merge_fields()))

doc_1.merge(
        Company = ', (Ticker: %s)' % tgt_ticker
        ,Analyst_Email = "Kevin Hebreo \n Hebreo@uconn.edu"
        )

doc_1.write('C:\\Users\\OBar\\Documents\\General Research\\Output\\Test_one_Res.docx')
Final_Ass.to_excel('C:\\Users\\OBar\\Documents\\General Research\\Output\\test %s.xlsx' % tgt_ticker)
# add a table to the end and create a reference variable
# extra row is so we can add the header row


