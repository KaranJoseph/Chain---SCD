# -*- coding: utf-8 -*-
"""
Created on Fri Sep 27 11:31:17 2019

@author: kjoseph
"""

import pandas as pd
#import numpy as np
import glob
#import xlwt
#from xlwt import Workbook
#from openpyxl import load_workbook


files=glob.glob('D:\\Test\\*.xlsx')## Change Input Location Here.   Dont change wildcard operator *.xlsx


sheets=['DemandRegion','DemandRequirement','Facility','FacilityInPeriod','InterfacilityLinkInPeriod','ProcessComponent',
            'ProductAtFacilityInPeriod','ServiceLinkInPeriod','TransportationMode','TransportationModeInPeriod']
file_names=[i[i.find('Scenario'):i.find('.xlsx')] for i in files ]


d= ['D:\\Test\\Answers' for i in file_names]## Give Ouput Location Here


target=list(map(lambda x,y:x+'\\'+y,d,file_names))
target=[i+'.xlsx' for i in target]

l=len(sheets)
count=0
for j in files:
    #wb=Workbook()
    with pd.ExcelWriter(target[count],engine='xlsxwriter') as writer:
        for i in sheets:
            data=pd.read_excel(j,sheet_name=i)
            data.Scenario=data.loc[:,'Scenario'].apply(lambda x:file_names[count])
            data.to_excel(writer,sheet_name=i,index=False)
    count=count+1
   # break

