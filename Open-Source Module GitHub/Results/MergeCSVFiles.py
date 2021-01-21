# -*- coding: utf-8 -*-
"""
Created on Thu Nov 26 13:28:31 2020

@author: giorgluc
"""

import os, glob
import pandas as pd
import itertools

os.chdir('C:/Users/lucag/Documents/Scuola/Master/Master_HS_2020/Master Thesis/OpenDSS/Test results')

NbrOfTotalPCC=11    #Total number of PCC in the grid
NbrOfPCC=10      # Number of PCC where a charger is installed during simulations

ListOfCombos=list()
index_combo=0
lst = list(itertools.product([0, 1], repeat=NbrOfTotalPCC))
for i in range(len(lst)):
     if sum(lst[i])==NbrOfPCC:
         ListOfCombos.insert(index_combo,lst[i])
         index_combo=index_combo+1
NbrOfCombination=len(ListOfCombos)
print(len(ListOfCombos))



pathVoltageBus='C:/Users/lucag/Documents/Scuola/Master/Master_HS_2020/Master Thesis/OpenDSS/Test results/BusVoltage'
pathLineLoading='C:/Users/lucag/Documents/Scuola/Master/Master_HS_2020/Master Thesis/OpenDSS/Test results/LineLoading'
pathTrafoLoading='C:/Users/lucag/Documents/Scuola/Master/Master_HS_2020/Master Thesis/OpenDSS/Test results/TrafoLoading'

for i in range(NbrOfCombination):
    all_VoltageBus = glob.glob(os.path.join(pathVoltageBus, "*_Combo"+str(i)+".csv"))
    all_LineLoading = glob.glob(os.path.join(pathLineLoading, "*_Combo"+str(i)+".csv"))
    all_TrafoLoading=glob.glob(os.path.join(pathTrafoLoading, "*_Combo"+str(i)+".csv"))
    
    dfVoltage=(pd.read_csv(f, sep=',') for f in all_VoltageBus)
    dfVoltageMerged   = pd.concat(dfVoltage, ignore_index=True)
    dfVoltageMerged.to_csv('VoltageMergedCombination'+str(i)+'.csv')
    
    dfLoading=(pd.read_csv(f, sep=',') for f in all_LineLoading)
    dfLoadingMerged   = pd.concat(dfLoading, ignore_index=True)
    dfLoadingMerged.to_csv('LineLoadingMergedCombination'+str(i)+'.csv')

    dfTLoading=(pd.read_csv(f, sep=',') for f in all_TrafoLoading)
    dfTLoadingMerged   = pd.concat(dfTLoading, ignore_index=True)
    dfTLoadingMerged.to_csv('TrafoLoadingMergedCombination'+str(i)+'.csv')
