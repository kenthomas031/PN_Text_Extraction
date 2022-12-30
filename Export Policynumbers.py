# -*- coding: utf-8 -*-
"""
Created on Fri Feb 21 07:00:20 2020

@author: u0a009
"""

import pandas as pd
import re
import os

os.chdir("""\\\\chnas07\\ced\\CED-Priv\\Analytics\\Customer Research\\Analysts\\Kenneth Thomas\\Ad Hoc\\2020 Ad Hoc\\Alex Kuch Strip Policy Nums""")

myDoc = pd.read_excel('Marvin Policy Numbers 11022020.xlsx')

count = 0

PNDict = {}
for i in range(0,len(myDoc)):
    PNDict.update({i:'None'})

for item in myDoc['Title']:
    flag = 0
    for temp in item.split():
#        number = filter(lambda x: x.isdigit(),temp)
        number = re.sub('\D','',temp)
#        print(number)
        if len(number) == 10:
            flag += 1
            if flag == 1:
                policyNum = number
                PNDict.update({count:policyNum})
            elif flag > 1:
                print('Warning: ',count,flag)
                policyNum2 = number
                if policyNum != policyNum2:
                    PNDict.update({count:[policyNum, policyNum2]})
                    print('Two PNs Found')
    count += 1
    
newFrame = pd.DataFrame.from_dict(PNDict).T
newFrame = newFrame.rename(columns={0:'PN1',1:'PN2'})

newFrame['Does PN1 = PN2'] = newFrame['PN1'] == newFrame['PN2']
newFrame['PN3'] = newFrame['PN2'][newFrame['Does PN1 = PN2'] == False]
newFrame.drop(columns = ['PN2','Does PN1 = PN2'],inplace=True)
newFrame = newFrame.rename(columns={'PN3':'PN2'})

myDoc['PolicyNumber1'] = newFrame['PN1']
myDoc['PolicyNumber2'] = newFrame['PN2']
myDoc.drop(columns=['PolicyNumber2'],inplace=True)

myDoc.to_excel('Marvin Policy Numbers 11022020 Complete.xlsx')


    