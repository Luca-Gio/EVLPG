"""
Created on Wed Dec 18 16:54:48 2019

@author: lucag
"""
"""
EVLPG - MASTER THESIS
Approach : Without Limitations
Timeframe : Must be choosen
Details: - Complex Grid
         - Load profils calculated with a 15 minutes time step
         - Allows to simulate a chosen number of cases
         - The number of repetitions per case can be chosen
"""
#***********************************************************
# IMPORT LIBRARIES
#***********************************************************
import os,fnmatch
import win32com.client
import pandas as pd
import matplotlib.pyplot as plt 
import numpy as np
import random as rd
import math
from PDFs_Functions import *
from Grid_Modelling import pathGrid,pathDSS,timestep,timeperiod,dfLines,dfLoads,dfGrid, pathElement, dfTrafo, export_result, percentage , repetitions, EV_distribution, pathEV,MethodOfCalculation,NO_EV_Grid,NbrOfDistr, pathRes, ScailingFactorDistance,PercentPower
from Results_Conversion_Functions import *
import itertools

#***********************************************************
# CONNECT OPENDSS TO PYTHON
#***********************************************************
# Attribute OpenDSS command to python variables
dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")
dssText = dssObj.Text
dssCircuit = dssObj.ActiveCircuit          


#***********************************************************
# IMPORT EXCEL FILES OF PDFs
#***********************************************************
if NO_EV_Grid==0:

    #Define pattern for folders
    pathPDF ='./PDFs/'
    for i in range(len(pathGrid)):
        char=pathGrid[len(pathGrid)-1-i]
    
        if char == '/':
            index=i
            pathInfos = pathGrid[0:len(pathGrid)-index]
            break
    
    # Import Excel files
    ArrivalTimeDistributionWD = pd.read_csv(pathPDF + "D_weekdays.csv",delimiter = ";")      #Distribution of arrival time during weekdays
    ArrivalTimeDistributionWE = pd.read_csv(pathPDF +"D_weekends.csv",delimiter = ";")       #Distribution of arrival time during weekends
    EnergyDistribution = pd.read_csv(pathPDF +"Energy.csv",delimiter = ";")                  #Distribution of energy demand
    ConnectionTimeDistribution = pd.read_csv(pathPDF +"Connection_time.csv",delimiter = ";") #Distribution of connection time
    CarsDistribution = pd.read_csv(pathPDF+"EV_Model_CH.csv",delimiter=";")                  #Tipical EV car in CH with values of battery energy and Max. power of charge
    #EV_car=CarsDistribution._get_values
    
    
    #***********************************************************
    # VARIABLES DEFINITION
    #***********************************************************
    Repetition = int(repetitions) # How many times we would like to repeat the simulation
    days=int((timeperiod*timestep)/(60*24))
    DaySteps=int(timeperiod/days)
    #TotalNumberOfCustomers=sum(dfLoads['number of customers'])
    ListOfCombos=list()
    index_combo=0
    
   #***********************************************************
    # DEFINITION OF NUMBER OF DISTRIBUTIONS AND APPLICATION OF ALGORITHM FOR SIMPLIFICATION (IF SELECTED AS OPTION)
    #***********************************************************
    if EV_distribution == 0:
        #NbrOfCS=int((int(percentage)/100)*TotalNumberOfCustomers)
        NbrOfCS=int(round((int(percentage)/100)*len(dfLoads)))
        if MethodOfCalculation == 1:
            lst = list(itertools.product([0, 1], repeat=len(dfLoads)))
            for i in range(len(lst)):
                if sum(lst[i])==NbrOfCS:
                    ListOfCombos.insert(index_combo,lst[i])
                    index_combo=index_combo+1
            NbrOfEVdistribution=len(ListOfCombos)
            
        elif MethodOfCalculation == 0:
            NbrOfEVdistribution = NbrOfDistr
        
        elif MethodOfCalculation == 2:
            dssText.Command = "compile '"+pathDSS+"'"        
            #Solve only 1 calculation in order to be able to exttract distances and bus names from OpenDSS
            dssText.Command ="solve"
            Distances = dssCircuit.AllBusDistances
            Names = dssCircuit.AllBusNames
            BusInfos=[[0 for i in range(2)] for i in range (len(Distances))]
            BusColumns=['Bus Name', 'Distance']
            for i in range(len(Distances)):
                BusInfos[i][0]=Names[i]
                BusInfos[i][1]=Distances[i]
            BusInfos=pd.DataFrame(data=BusInfos,columns=BusColumns)
            #BusInfos.to_csv(pathRes+"/Distances.csv", sep=',',index=False)
            TotBusNameList=list()
            for i in range(len(dfLoads)):
                TotBusNameList.insert(i,dfLoads['Bus'][i].lower())
            LineToBusNameList=list()
            for i in range(len(dfLines)):
                LineToBusNameList.insert(i,dfLines['To'][i].lower())
            
            dfTotBusNameList=pd.DataFrame(data=TotBusNameList)
            dfLineToBusNameList=pd.DataFrame(data=LineToBusNameList)
            Criteria=pd.DataFrame(data=[[0.0 for i in range(6)]for i in range(len(dfLoads))], columns=['Load Name', 'Bus', 'Factor','Factor_Power','number of customers','EV charger'])
            ix=0

            for i in range(len(TotBusNameList)):

                if TotBusNameList[i] in dfTrafo['LV bus'][0].lower():
                    Criteria['Load Name'][ix]=dfLoads['Name'][i]
                    Criteria['Bus'][ix]=dfLoads['Bus'][i]
                    Criteria['Factor'][ix]=5000
                    Criteria['number of customers'][ix]= dfLoads['number of customers'][i]
                    Criteria['EV charger'][ix]= dfLoads['EV charger'][i]
                    Criteria['Factor_Power'][ix]=1
                    ix=ix+1

                else:
                    IndexCriteria=np.where(BusInfos['Bus Name']==TotBusNameList[i])[0][0]
                    IndexPNom=np.where(dfLineToBusNameList==TotBusNameList[i])[0][0]
                    Criteria['Load Name'][ix]=dfLoads['Name'][i]
                    Criteria['Bus'][ix]=dfLoads['Bus'][i]
                    Criteria['Factor'][ix]=0.99*(dfLoads['annual energy consumption'][i]/1000)+(0.01*BusInfos['Distance'][IndexCriteria]*ScailingFactorDistance)
                    Criteria['number of customers'][ix]= dfLoads['number of customers'][i]
                    Criteria['EV charger'][ix]= dfLoads['EV charger'][i]
                    Criteria['Factor_Power'][ix]=(dfLoads['Pmax_PCC'][i]*1000)/(dfLines['Fuse limit'][IndexPNom]*400*math.sqrt(3))

                    ix=ix+1                        
                    #print(FactorPower)
                    #Criteria['Factor_Power'][ix]=FactorPower               
                    
            Criteria=Criteria.sort_values(by=['Factor'],ascending=False).reset_index(drop=True)
            
            for i in range(len(Criteria)):
                if Criteria['Factor_Power'][i] >=PercentPower:
                    Criteria =Criteria.drop(i)
            Criteria=Criteria.reset_index(drop=True)
            
            Algo=NbrOfCS+2
            if Algo > len(Criteria):
                Algo=len(Criteria)
            
            CS_Algorithm=Criteria[:][0:Algo]
            if Algo > NbrOfCS:
                lst = list(itertools.product([0, 1], repeat=len(CS_Algorithm)))
                for i in range(len(lst)):
                    if sum(lst[i])==NbrOfCS:
                        ListOfCombos.insert(index_combo,lst[i])
                        index_combo=index_combo+1
                NbrOfEVdistribution=len(ListOfCombos)
            else:
               NbrOfEVdistribution=1 
                    
    elif EV_distribution == 1:
        dfDistr=pd.read_excel(pathEV)
        NbrOfCS=int(round((int(percentage)/100)*len(dfDistr)))
    
        if MethodOfCalculation == 1:
            lst = list(itertools.product([0, 1], repeat=len(dfDistr)))
            for i in range(len(lst)):
                if sum(lst[i])==NbrOfCS:
                    ListOfCombos.insert(index_combo,lst[i])
                    index_combo=index_combo+1
            NbrOfEVdistribution=len(ListOfCombos)
        elif MethodOfCalculation == 0:
            NbrOfEVdistribution = NbrOfDistr
    
    elif EV_distribution == 2:
        dfDistr=pd.read_excel(pathEV)
        vect_del=list()
        for i in range(len(dfDistr)):
            #First it is necessary to delete the possibility to install chargers where are already installed based on external file
            index_del=np.where(dfLoads['Bus']==dfDistr['Bus'][i])[0][0]
            vect_del.insert(i,index_del)
        dfMix=dfLoads.drop(vect_del, axis=0)
        dfMix=dfMix.reset_index(drop=True)
        
        additionalCS=int(round((int(percentage)/100)*len(dfMix)))
        NbrOfCS=len(dfDistr)+additionalCS
        if MethodOfCalculation == 1:
            lst = list(itertools.product([0, 1], repeat=len(dfMix)))
            for i in range(len(lst)):
                if sum(lst[i])==additionalCS:
                    ListOfCombos.insert(index_combo,lst[i])
                    index_combo=index_combo+1
            NbrOfEVdistribution=len(ListOfCombos)
            
        elif MethodOfCalculation == 0:
            NbrOfEVdistribution = NbrOfDistr
            
            
        elif MethodOfCalculation == 2:
            dssText.Command = "compile '"+pathDSS+"'"        
            #Solve only 1 calculation in order to be able to exttract distances and bus names from OpenDSS
            dssText.Command ="solve"
            Distances = dssCircuit.AllBusDistances
            Names = dssCircuit.AllBusNames
            BusInfos=[[0 for i in range(2)] for i in range (len(Distances))]
            BusColumns=['Bus Name', 'Distance']
            for i in range(len(Distances)):
                BusInfos[i][0]=Names[i]
                BusInfos[i][1]=Distances[i]
            BusInfos=pd.DataFrame(data=BusInfos,columns=BusColumns)
            #BusInfos.to_csv(pathRes+"/Distances.csv", sep=',',index=False)
            TotBusNameList=list()
            for i in range(len(dfMix)):
                TotBusNameList.insert(i,dfMix['Bus'][i].lower())
            LineToBusNameList=list()
            for i in range(len(dfLines)):
                LineToBusNameList.insert(i,dfLines['To'][i].lower())
            
            dfTotBusNameList=pd.DataFrame(data=TotBusNameList)
            dfLineToBusNameList=pd.DataFrame(data=LineToBusNameList)
            Criteria=pd.DataFrame(data=[[0.0 for i in range(6)]for i in range(len(dfMix))], columns=['Load Name', 'Bus', 'Factor','Factor_Power','number of customers','EV charger'])
            ix=0

            for i in range(len(TotBusNameList)):

                if TotBusNameList[i] in dfTrafo['LV bus'][0].lower():
                    Criteria['Load Name'][ix]=dfMix['Name'][i]
                    Criteria['Bus'][ix]=dfMix['Bus'][i]
                    Criteria['Factor'][ix]=5000
                    Criteria['number of customers'][ix]= dfMix['number of customers'][i]
                    Criteria['EV charger'][ix]= dfMix['EV charger'][i]
                    Criteria['Factor_Power'][ix]=1
                    ix=ix+1

                else:
                    IndexCriteria=np.where(BusInfos['Bus Name']==TotBusNameList[i])[0][0]
                    IndexPNom=np.where(dfLineToBusNameList==TotBusNameList[i])[0][0]
                    Criteria['Load Name'][ix]=dfMix['Name'][i]
                    Criteria['Bus'][ix]=dfMix['Bus'][i]
                    Criteria['Factor'][ix]=0.99*(dfMix['annual energy consumption'][i]/1000)+(0.01*BusInfos['Distance'][IndexCriteria]*ScailingFactorDistance)
                    Criteria['number of customers'][ix]= dfMix['number of customers'][i]
                    Criteria['EV charger'][ix]= dfMix['EV charger'][i]
                    Criteria['Factor_Power'][ix]=(dfMix['Pmax_PCC'][i]*1000)/(dfLines['Fuse limit'][IndexPNom]*400*math.sqrt(3))

                    ix=ix+1                        
                    #print(FactorPower)
                    #Criteria['Factor_Power'][ix]=FactorPower               
                    
            Criteria=Criteria.sort_values(by=['Factor'],ascending=False).reset_index(drop=True)
            
            for i in range(len(Criteria)):
                if Criteria['Factor_Power'][i] >=PercentPower:
                    Criteria =Criteria.drop(i)
            Criteria=Criteria.reset_index(drop=True)
            
            Algo=additionalCS
            if Algo > len(Criteria):
                Algo=len(Criteria)
            
            CS_Algorithm=Criteria[:][0:Algo]
            if Algo > additionalCS:
                lst = list(itertools.product([0, 1], repeat=len(CS_Algorithm)))
                for i in range(len(lst)):
                    if sum(lst[i])==NbrOfCS:
                        ListOfCombos.insert(index_combo,lst[i])
                        index_combo=index_combo+1
                NbrOfEVdistribution=len(ListOfCombos)
            else:
               NbrOfEVdistribution=1
    #NbrOfCS = 2 #Nbr of EV charger choosen for the simulation
     # 
    #Variable which defines if it is Weekday or Weekend
    WDorWE = 0
    Timeperiod=int(timeperiod)
    Timestep=int(timestep)
    #Definition of power level based on the typology of charger. Assumption: three phases connection, not considered monophase charge
    PowerLevelsPriv = [11000,22000,22000]         #List of possible power of a privat EV charger
    PowerLevelsWork = [11000,22000,11000]         #List of possible power of a workplace EV charger
    PowerLevelsPub = [11000,22000]         #List of possible power of a public EV charger
    
    # Define the distribution of Arrival time splitted in private, public and workplace as well as weekdays and weekends
    VarArrivalTime = ArrivalTimeDistributionWD['Arrival time'] # From 00:00 until 23:45 (1 day)
    ProbArrTimeWDPriv = ArrivalTimeDistributionWD['Private']   # Probability Distribution of arrival time for private charger during weekdays
    ProbArrTimeWEPriv = ArrivalTimeDistributionWE['Private']   # Probability Distribution of arrival time for private charger during weekend
    
    ProbArrTimeWDPub = ArrivalTimeDistributionWD['Public'] # Distribution of arrival time for public charger during weekdays
    ProbArrTimeWEPub = ArrivalTimeDistributionWE['Public'] # Distribution of arrival time for public charger during weekends
    
    ProbArrTimeWDWork = ArrivalTimeDistributionWD['Workplace'] # Distribution of arrival time for workplace charger during weekdays
    ProbArrTimeWEWork = ArrivalTimeDistributionWE['Workplace'] # Distribution of arrival time for workplace charger during weekends
    
    # Distribution of car model
    ProbEV = CarsDistribution['probability']
    
    #Define the distribution of connection time
    VarConnectionTime=ConnectionTimeDistribution['Connection duration']
    ProbConnTimePriv = ConnectionTimeDistribution['Private']  # connection time Probability distribution in private
    ProbConnTimePub = ConnectionTimeDistribution['Public']   # connection time Probability distribution in public
    ProbConnTimeWork = ConnectionTimeDistribution['Workplace']# connection time Probability distribution in workplace
    
    #Define the distribution of energy demand
    VarEnergyDemand=EnergyDistribution['Energy demand']
    ProbEnergyPriv = EnergyDistribution['Private'] # energy distribution in private
    ProbEnergyPub = EnergyDistribution['Public'] # energy distribution in public
    ProbEnergyWork = EnergyDistribution['Workplace'] # energy distribution in workplace
    
    columns = ['Combo Nbr:','Charger Names']
    LoadsToTXT=[[0 for i in range(2)] for i in range(NbrOfEVdistribution)]
    LoadsToTXT=pd.DataFrame(data=LoadsToTXT,columns=columns)
    #******************************************************************************
    # DEFINE WHERE LOADS WILL BE LOCATED AND WHICH IS THE TOPOLOGY OF CHARGER
    #******************************************************************************
    
    for EVDistributionNbr in range(NbrOfEVdistribution):
        ListNumberOfCustomers=[ 0 for i in range(NbrOfCS)]
        NamesOfLoads = [ 0 for i in range(NbrOfCS)]
        TypeOfCS = [0 for i in range(NbrOfCS)]
        PowerLevels = [0 for i in range(NbrOfCS)]
        
    
        if EV_distribution == 0:
            LoadNumber = rd.sample(range(len(dfLoads)),NbrOfCS)
            if MethodOfCalculation == 0:
                for EV_index in range (NbrOfCS):
                 
                    NbrOfCustomer = dfLoads['number of customers'][LoadNumber[EV_index]]
                    TypeOfCharger= dfLoads['EV charger'][LoadNumber[EV_index]]
                    #if the number of customers is higher than 1 in a specific node, it is possible to
                    #choose how many chargers install
                    if NbrOfCustomer > 1:
                        Customers = np.random.choice(NbrOfCustomer)+1
                    else:
                        Customers = NbrOfCustomer
                    #print(Customers)
                    #Definition of a vector which contains the topology of chager (Priv, Pub or workplace)
                    customertypology = list()
                    VectorOfPower = list()
    
                    for cust in range(Customers):
                            
                        if TypeOfCharger == 'Private':
                            customertypology.insert(cust,'Private')
                            VectorOfPower.insert(cust, np.random.choice(PowerLevelsPriv))
                                                 
                        elif TypeOfCharger == 'Public':
                            customertypology.insert(cust,'Public')
                            VectorOfPower.insert(cust, np.random.choice(PowerLevelsPub))
            
                        elif TypeOfCharger == 'Workplace':
                            customertypology.insert(cust,'Workplace')
                            VectorOfPower.insert(cust, np.random.choice(PowerLevelsWork))
            
                    ListNumberOfCustomers[EV_index] = Customers
                    NamesOfLoads[EV_index] =dfLoads['Name'][LoadNumber[EV_index]]
                    TypeOfCS[EV_index] = customertypology
                    PowerLevels[EV_index]=VectorOfPower
                    #print(ListNumberOfCustomers)
                
            elif MethodOfCalculation == 1:
                #for EV_index in range (NbrOfCS):
                index_combos=0
                for i in range(len(ListOfCombos[EVDistributionNbr])):
                    if ListOfCombos[EVDistributionNbr][i]==1:
                        NbrOfCustomer = dfLoads['number of customers'][i]
                        TypeOfCharger= dfLoads['EV charger'][i]
                        NamesOfLoads[index_combos] =dfLoads['Name'][i]
                        if NbrOfCustomer > 1:
                            Customers = np.random.choice(NbrOfCustomer)+1
                        else:
                            Customers = NbrOfCustomer
                        #print(Customers)
                        #Definition of a vector which contains the topology of chager (Priv, Pub or workplace)
                        customertypology = list()
                        VectorOfPower = list()
        
                        for cust in range(Customers):
                                
                            if TypeOfCharger == 'Private':
                                customertypology.insert(cust,'Private')
                                VectorOfPower.insert(cust, np.random.choice(PowerLevelsPriv))
                                                     
                            elif TypeOfCharger == 'Public':
                                customertypology.insert(cust,'Public')
                                VectorOfPower.insert(cust, np.random.choice(PowerLevelsPub))
                
                            elif TypeOfCharger == 'Workplace':
                                customertypology.insert(cust,'Workplace')
                                VectorOfPower.insert(cust, np.random.choice(PowerLevelsWork))                                     
                        ListNumberOfCustomers[index_combos] = Customers
                        TypeOfCS[index_combos] = customertypology
                        PowerLevels[index_combos]=VectorOfPower  
                        index_combos=index_combos+1
                #print(NamesOfLoads)
            
            elif MethodOfCalculation == 2:
                index_combos=0
                if Algo > NbrOfCS:
                    for i in range(len(ListOfCombos[EVDistributionNbr])):
                        if ListOfCombos[EVDistributionNbr][i]==1:
                            NbrOfCustomer = int(CS_Algorithm['number of customers'][i])
                            TypeOfCharger= CS_Algorithm['EV charger'][i]
                            NamesOfLoads[index_combos] =CS_Algorithm['Load Name'][i]
                            if NbrOfCustomer > 1:
                                Customers = np.random.choice(NbrOfCustomer)+1
                            else:
                                Customers = NbrOfCustomer
                            #print(Customers)
                            #Definition of a vector which contains the topology of chager (Priv, Pub or workplace)
                            customertypology = list()
                            VectorOfPower = list()
            
                            for cust in range(Customers):
                                    
                                if TypeOfCharger == 'Private':
                                    customertypology.insert(cust,'Private')
                                    VectorOfPower.insert(cust, np.random.choice(PowerLevelsPriv))
                                                         
                                elif TypeOfCharger == 'Public':
                                    customertypology.insert(cust,'Public')
                                    VectorOfPower.insert(cust, np.random.choice(PowerLevelsPub))
                    
                                elif TypeOfCharger == 'Workplace':
                                    customertypology.insert(cust,'Workplace')
                                    VectorOfPower.insert(cust, np.random.choice(PowerLevelsWork))                                     
                            ListNumberOfCustomers[index_combos] = Customers
                            TypeOfCS[index_combos] = customertypology
                            PowerLevels[index_combos]=VectorOfPower  
                            index_combos=index_combos+1
                else:
                    for i in range(len(CS_Algorithm)):
                        NbrOfCustomer = int(CS_Algorithm['number of customers'][i])
                        TypeOfCharger= CS_Algorithm['EV charger'][i]
                        NamesOfLoads[index_combos] =CS_Algorithm['Load Name'][i]
                        if NbrOfCustomer > 1:
                            Customers = np.random.choice(NbrOfCustomer)+1
                        else:
                            Customers = NbrOfCustomer
                        #print(Customers)
                        #Definition of a vector which contains the topology of chager (Priv, Pub or workplace)
                        customertypology = list()
                        VectorOfPower = list()
        
                        for cust in range(Customers):
                                
                            if TypeOfCharger == 'Private':
                                customertypology.insert(cust,'Private')
                                VectorOfPower.insert(cust, np.random.choice(PowerLevelsPriv))
                                                     
                            elif TypeOfCharger == 'Public':
                                customertypology.insert(cust,'Public')
                                VectorOfPower.insert(cust, np.random.choice(PowerLevelsPub))
                
                            elif TypeOfCharger == 'Workplace':
                                customertypology.insert(cust,'Workplace')
                                VectorOfPower.insert(cust, np.random.choice(PowerLevelsWork))                                     
                        ListNumberOfCustomers[index_combos] = Customers
                        TypeOfCS[index_combos] = customertypology
                        PowerLevels[index_combos]=VectorOfPower  
                        index_combos=index_combos+1
                #print(NamesOfLoads)              
        
        elif EV_distribution ==1:
            if MethodOfCalculation == 0:
                LoadNumber = rd.sample(range(len(dfDistr)),NbrOfCS)
                for EV_index in range(NbrOfCS):
                    customertypology = list()
                    VectorOfPower = list()
                     
                    ListNumberOfCustomers[EV_index] = dfDistr['Customers'][LoadNumber[EV_index]]
                    for i in range(ListNumberOfCustomers[EV_index]):
                        customertypology.insert(i,dfDistr['Type'][LoadNumber[EV_index]])
                        VectorOfPower.insert(i,dfDistr['Power [W]'][LoadNumber[EV_index]])
                        
                    NamesOfLoads[EV_index] =dfDistr['Name'][LoadNumber[EV_index]]
                    TypeOfCS[EV_index] = customertypology
                    PowerLevels[EV_index]=VectorOfPower
            else:
                index_combos=0
                customertypology = list()
                VectorOfPower = list()
                for x in range(len(ListOfCombos[EVDistributionNbr])):
                    if ListOfCombos[EVDistributionNbr][x]==1:                     
                        ListNumberOfCustomers[index_combos] = dfDistr['Customers'][x]
                        for i in range(ListNumberOfCustomers[index_combos]):
                            customertypology.insert(i,dfDistr['Type'][x])
                            VectorOfPower.insert(i,dfDistr['Power [W]'][x])
                            
                        NamesOfLoads[index_combos] =dfDistr['Name'][x]
                        TypeOfCS[index_combos] = customertypology
                        PowerLevels[index_combos]=VectorOfPower
                        index_combos=index_combos+1
                #print(NamesOfLoads)
        
        elif EV_distribution ==2:
            
            if MethodOfCalculation == 0:
                LoadNumber = rd.sample(range(len(dfMix)),additionalCS)
                for EV_index in range (len(dfDistr),NbrOfCS):
                 
                    NbrOfCustomer = dfMix['number of customers'][LoadNumber[EV_index-len(dfDistr)]]
                    TypeOfCharger= dfMix['EV charger'][LoadNumber[EV_index-len(dfDistr)]]
                    #if the number of customers is higher than 1 in a specific node, it is possible to
                    #choose how many chargers install
                    if NbrOfCustomer > 1:
                        Customers = np.random.choice(NbrOfCustomer)+1
                    else:
                        Customers = NbrOfCustomer
                    #print(Customers)
                    #Definition of a vector which contains the topology of chager (Priv, Pub or workplace)
                    customertypology = list()
                    VectorOfPower = list()
    
                    for cust in range(Customers):
                            
                        if TypeOfCharger == 'Private':
                            customertypology.insert(cust,'Private')
                            VectorOfPower.insert(cust, np.random.choice(PowerLevelsPriv))
                                                 
                        elif TypeOfCharger == 'Public':
                            customertypology.insert(cust,'Public')
                            VectorOfPower.insert(cust, np.random.choice(PowerLevelsPub))
            
                        elif TypeOfCharger == 'Workplace':
                            customertypology.insert(cust,'Workplace')
                            VectorOfPower.insert(cust, np.random.choice(PowerLevelsWork))
            
                    ListNumberOfCustomers[EV_index] = Customers
                    NamesOfLoads[EV_index] =dfMix['Name'][LoadNumber[EV_index-len(dfDistr)]]
                    TypeOfCS[EV_index] = customertypology
                    PowerLevels[EV_index]=VectorOfPower
                          
            elif MethodOfCalculation ==1:
                #for EV_index in range (NbrOfCS):
                index_combos=len(dfDistr)
                for i in range(len(ListOfCombos[EVDistributionNbr])):
                    if ListOfCombos[EVDistributionNbr][i]==1:
                        NbrOfCustomer = dfMix['number of customers'][i]
                        TypeOfCharger= dfMix['EV charger'][i]
                        NamesOfLoads[index_combos] =dfMix['Name'][i]
                        if NbrOfCustomer > 1:
                            Customers = np.random.choice(NbrOfCustomer)+1
                        else:
                            Customers = NbrOfCustomer
                        #print(Customers)
                        #Definition of a vector which contains the topology of chager (Priv, Pub or workplace)
                        customertypology = list()
                        VectorOfPower = list()
        
                        for cust in range(Customers):
                                
                            if TypeOfCharger == 'Private':
                                customertypology.insert(cust,'Private')
                                VectorOfPower.insert(cust, np.random.choice(PowerLevelsPriv))
                                                     
                            elif TypeOfCharger == 'Public':
                                customertypology.insert(cust,'Public')
                                VectorOfPower.insert(cust, np.random.choice(PowerLevelsPub))
                
                            elif TypeOfCharger == 'Workplace':
                                customertypology.insert(cust,'Workplace')
                                VectorOfPower.insert(cust, np.random.choice(PowerLevelsWork))                                     
                        ListNumberOfCustomers[index_combos] = Customers
                        TypeOfCS[index_combos] = customertypology
                        PowerLevels[index_combos]=VectorOfPower  
                        index_combos=index_combos+1
             
            elif MethodOfCalculation == 2:
                index_combos=len(dfDistr)
                if Algo > NbrOfCS:
                    for i in range(len(ListOfCombos[EVDistributionNbr])):
                        if ListOfCombos[EVDistributionNbr][i]==1:
                            NbrOfCustomer = int(CS_Algorithm['number of customers'][i])
                            TypeOfCharger= CS_Algorithm['EV charger'][i]
                            NamesOfLoads[index_combos] =CS_Algorithm['Load Name'][i]
                            if NbrOfCustomer > 1:
                                Customers = np.random.choice(NbrOfCustomer)+1
                            else:
                                Customers = NbrOfCustomer
                            #print(Customers)
                            #Definition of a vector which contains the topology of chager (Priv, Pub or workplace)
                            customertypology = list()
                            VectorOfPower = list()
            
                            for cust in range(Customers):
                                    
                                if TypeOfCharger == 'Private':
                                    customertypology.insert(cust,'Private')
                                    VectorOfPower.insert(cust, np.random.choice(PowerLevelsPriv))
                                                         
                                elif TypeOfCharger == 'Public':
                                    customertypology.insert(cust,'Public')
                                    VectorOfPower.insert(cust, np.random.choice(PowerLevelsPub))
                    
                                elif TypeOfCharger == 'Workplace':
                                    customertypology.insert(cust,'Workplace')
                                    VectorOfPower.insert(cust, np.random.choice(PowerLevelsWork))                                     
                            ListNumberOfCustomers[index_combos] = Customers
                            TypeOfCS[index_combos] = customertypology
                            PowerLevels[index_combos]=VectorOfPower  
                            index_combos=index_combos+1
                else:
                    for i in range(len(CS_Algorithm)):
                        NbrOfCustomer = int(CS_Algorithm['number of customers'][i])
                        TypeOfCharger= CS_Algorithm['EV charger'][i]
                        NamesOfLoads[index_combos] =CS_Algorithm['Load Name'][i]
                        if NbrOfCustomer > 1:
                            Customers = np.random.choice(NbrOfCustomer)+1
                        else:
                            Customers = NbrOfCustomer
                        #print(Customers)
                        #Definition of a vector which contains the topology of chager (Priv, Pub or workplace)
                        customertypology = list()
                        VectorOfPower = list()
        
                        for cust in range(Customers):
                                
                            if TypeOfCharger == 'Private':
                                customertypology.insert(cust,'Private')
                                VectorOfPower.insert(cust, np.random.choice(PowerLevelsPriv))
                                                     
                            elif TypeOfCharger == 'Public':
                                customertypology.insert(cust,'Public')
                                VectorOfPower.insert(cust, np.random.choice(PowerLevelsPub))
                
                            elif TypeOfCharger == 'Workplace':
                                customertypology.insert(cust,'Workplace')
                                VectorOfPower.insert(cust, np.random.choice(PowerLevelsWork))                                     
                        ListNumberOfCustomers[index_combos] = Customers
                        TypeOfCS[index_combos] = customertypology
                        PowerLevels[index_combos]=VectorOfPower  
                        index_combos=index_combos+1           
            
            for EV_index in range(len(dfDistr)):
                customertypology = list()
                VectorOfPower = list()
                 
                ListNumberOfCustomers[EV_index] = dfDistr['Customers'][EV_index]
                for i in range(ListNumberOfCustomers[EV_index]):
                    customertypology.insert(i,dfDistr['Type'][EV_index])
                    VectorOfPower.insert(i,dfDistr['Power [W]'][EV_index])
                    
                NamesOfLoads[EV_index] =dfDistr['Name'][EV_index]
                TypeOfCS[EV_index] = customertypology
                PowerLevels[EV_index]=VectorOfPower 
           
        LoadsToTXT['Combo Nbr:'][EVDistributionNbr]=EVDistributionNbr
        LoadsToTXT['Charger Names'][EVDistributionNbr]=NamesOfLoads
        
    #******************************************************************************
    # START OF SIMULATION
    #******************************************************************************
        for index_repeat_sim in range (Repetition):
            SumLoadProfiles=[[0 for i in range(Timeperiod)] for i in range(NbrOfCS)]
            WDorWE=0
            for index_day in range(days):  
                WDorWE = WDorWE + 1
            #for day in range(1):                
    
    #******************************************************************************
    # SELECTION OF CAR AND ARRIVAL TIME, CONNECTION TIME and ENERGY DEMAND BASED ON THE TOPOLOGY OF CAR
    #******************************************************************************
                for NbrCharger in range(NbrOfCS):
                    testSLP=[[0 for i in range(Timeperiod)] for i in range(ListNumberOfCustomers[NbrCharger])]
                    for customer in range(ListNumberOfCustomers[NbrCharger]):
                        LoadProfiles=[0 for i in range(DaySteps)]                   
                        EV = CarsDistribution.iloc[ElectricVehicle(len(CarsDistribution),ProbEV)][:]                   
                        CS = TypeOfCS[NbrCharger][customer]
                    
                        #Assign the power and check that is not higher than the maximal power admissible in the car
                        ChargePower = PowerLevels[NbrCharger][customer]
                        MaxChargePower = EV['Charge power [kW] AC']                  #[W]
                        if ChargePower > MaxChargePower:
                                ChargePower = MaxChargePower
                         
                        if CS == 'Private':  #Private
                                if WDorWE % 2 == 1: #if WDorWE <= 5                       
                                    #Random values of arrival time during weekdays
                                    ArrivalTime = ArrTime(VarArrivalTime,ProbArrTimeWDPriv)
                                else:                                
                                    #Random values of arrival time during weekends
                                    ArrivalTime = ArrTime(VarArrivalTime,ProbArrTimeWEPriv)                               
                                #Random values of connection time 
                                ConnectionTime = Connection_Time(VarConnectionTime,ProbConnTimePriv)
                                #Random value of energy
                                EnergyDemand = Energy_Demand(VarEnergyDemand,ProbEnergyPriv)*1000 #[Wh]
        
                        elif CS == 'Public': #Public                           
                                if WDorWE % 2 == 1: #if WDorWE <= 5                               
                                    #Random values of arrival time during weekdays
                                    ArrivalTime = ArrTime(VarArrivalTime,ProbArrTimeWDPub)
                                else:                               
                                    #Random values of arrival time during weekends
                                    ArrivalTime = ArrTime(VarArrivalTime,ProbArrTimeWEPub)
                                    
                                #Random values of connection time 
                                ConnectionTime = Connection_Time(VarConnectionTime,ProbConnTimePub)
                                #Random value of energy
                                EnergyDemand = Energy_Demand(VarEnergyDemand,ProbEnergyPub)*1000 #[Wh]
                                
                        elif CS == 'Workplace': #Workplace
                                if WDorWE % 2 == 1: #if WDorWE <= 5                              
                                    #Random values of arrival time during weekdays
                                    ArrivalTime = ArrTime(VarArrivalTime,ProbArrTimeWDWork)
                                else:                               
                                    #Random values of arrival time during weekends
                                    ArrivalTime = ArrTime(VarArrivalTime,ProbArrTimeWEWork)
                                    
                                #Random values of connection time 
                                ConnectionTime = Connection_Time(VarConnectionTime,ProbConnTimeWork)
                                #Random value of energy
                                EnergyDemand = Energy_Demand(VarEnergyDemand,ProbEnergyWork)*1000 #[Wh]
        
        #******************************************************************************
        # BATTERY DEFINITION
        #******************************************************************************
                        BatteryEnergy = EV['energy battery [kWh]']*1000  #[Wh]
                        
                        #Check that chosen Energy isn't higher than Battery energy of the car
                        if EnergyDemand > BatteryEnergy:
                                EnergyDemand = BatteryEnergy
                               
                        #Calculation of maximal time to charge full battery (from 0 to 100%)
                        TotalTimeCC = ((0.7*BatteryEnergy)/ChargePower)*3600      #[seconds]
                        TotalTimeCV = ((0.3*BatteryEnergy*2)/ChargePower)*3600  #[seconds]
                        TotalChargeTime = round(TotalTimeCC+TotalTimeCV,0)      #[seconds]
        
        #******************************************************************************
        # CONVERTION OF ARRIVAL TIME AND CONNECTION TIME IN SECONDS AND MINUTES
        #******************************************************************************
                         #Convert arrival time (WD) to decimal in order to sum the arrival time and the total time of charge
                        (h, m) = ArrivalTime.split(':')
                        ArrivalTimeSeconds=int(h) * 3600 + int(m) * 60
                        ArrivalTimeMinutes = ArrivalTimeSeconds/60
        
                            #Convert connection time to decimal in order to sum the connection time and the total time of charge
                        ConnectionTimeSeconds = ConnectionTime*3600
                        ConnectionTimeMinutes = ConnectionTime*60
                        
        #******************************************************************************
        # CALCULATION OF LOAD PROFILE (70% IN CC MODE and 30% IN CV MODE)
        #******************************************************************************
                        if EnergyDemand > 0.3*BatteryEnergy:                         
                                RatioCCMode = (EnergyDemand-0.3*BatteryEnergy)/(0.7*BatteryEnergy) #Percent of energy in constant current charge [%] -30 because we consider that 30% is in constant voltage charge
                                TimeCC = RatioCCMode*TotalTimeCC # in seconds
                                TimeOfCharge = TimeCC+TotalTimeCV #In seconds
                                t=0
                                
                                #It is necessary to chek if time of charge is no longer than the connection time
                                if TimeOfCharge > ConnectionTimeSeconds:
                                    TimeOfCharge = ConnectionTimeSeconds
                                    
                                EndOfCharge = ArrivalTimeSeconds + TimeOfCharge #In seconds
                                #Check that the charge will not depasse mdnight
                                if EndOfCharge > 86400: #In a day there are 24*60*60 = 86400seconds
                                    TimeOfCharge = 86400 - ArrivalTimeSeconds #The charge is limited at midnight
        
                                EndOfCharge = ArrivalTimeSeconds + TimeOfCharge #In seconds
                                
                                for time in range(int(ArrivalTimeMinutes),int(EndOfCharge/(60)),Timestep): #The charge start at the moment of the arrival
                                    
                                    if time <= int((ArrivalTimeSeconds+TimeCC)/60):
                                        LoadProfiles[int(time/Timestep)]=ChargePower/1000
                                       
                                    else:                    
                                        LoadProfiles[int(time/Timestep)]=((-ChargePower/(TotalTimeCV/60))*t+ChargePower)/1000
                                        t=t+Timestep
                        
        
                        elif EnergyDemand < 0.3*BatteryEnergy: #If the energy demand is lower than 30% of Battery Energy, we are in the constant voltage part
                                
                                E = 0.3*BatteryEnergy-EnergyDemand #Energy used to calculate the start time of charge
        
                                #Calculation of square
                                E_Wsec = E*3600
                                term_4ac =8*TotalTimeCV*(E_Wsec/ChargePower)
                                term_bsquare =4*(TotalTimeCV**2)   
                                square = math.sqrt(term_bsquare-term_4ac)
                                
                                TimeX=((TotalTimeCV)-(square/2)) #In seconds
                                t=TimeX/60 #In minutes
                                TimeOfCharge = (TotalTimeCV-TimeX) #In seconds
        
                                #It is necessary to chek if time of charge is no longer than the connection time
                                if TimeOfCharge > ConnectionTimeSeconds:
                                    TimeOfCharge = ConnectionTimeSeconds
                                
                                #Check that the charge will not depasse mdnight
                                EndOfCharge = ArrivalTimeSeconds + TimeOfCharge #In seconds
                                if EndOfCharge > 86400: #In a day there are 24*60*60 = 86400seconds
                                    TimeOfCharge = 86400 - ArrivalTimeSeconds #The charge is limited at midnight
        
                                EndOfCharge = ArrivalTimeSeconds + TimeOfCharge #In seconds
                                
                                for time in range(int(ArrivalTimeMinutes),int(EndOfCharge/60),Timestep):                
                                    LoadProfiles[int(time/Timestep)]=((-ChargePower/(TotalTimeCV/60))*t+ChargePower)/1000
                                    t=t+Timestep
                        for t in range(DaySteps):
                            testSLP[customer][t+index_day*DaySteps]=LoadProfiles[t]
                        for t in range(DaySteps):
                            SumLoadProfiles[NbrCharger][t+index_day*DaySteps]=LoadProfiles[t]+SumLoadProfiles[NbrCharger][t+index_day*DaySteps]
    
                    #plt.plot(pd.DataFrame(data=testSLP).transpose())
                    #plt.show()

                
                if WDorWE >= 7:
                    WDorWE = 0
                
            LPs = pd.DataFrame(data=SumLoadProfiles).transpose()
            plt.plot(LPs)
            plt.show()
            LPs.columns=NamesOfLoads
            LPs.to_csv(pathInfos+"LoadProfiles.csv", sep=',',index=False)
    
            #***********************************************************
            # WRITE LOAD SHAPES OF EVs IN TXT FILE
            #***********************************************************
            #Delete old txt files
                    
            npts=str(timeperiod)
            minterval=str(timestep)
            loadshapeEV=[[0  for i in range(2)] for i in range(NbrOfCS)]
            for index_loadshapeEV in range(NbrOfCS):
                loadshapeEV[index_loadshapeEV][0] = "New Loadshape."+NamesOfLoads[index_loadshapeEV]+"_EV npts="+npts+" minterval="+minterval+' UseActual=True'
                loadshapeEV[index_loadshapeEV][1]="~ mult="+str(tuple(LPs[NamesOfLoads[index_loadshapeEV]]))
    
            
            # Write the .dss file used then in OpenDSS to define lines
            loadshape_file = open(pathElement+"LoadshapeEV.txt","w")
            for i in range(len(loadshapeEV)):
                loadshape_file.write(loadshapeEV[i][0]+"\n"+loadshapeEV[i][1]+"\n")
            loadshape_file.close()
                
            #***********************************************************
            # DEFINE LOADS
            #***********************************************************
            loadsEV=[0 for i in range(NbrOfCS)]
            monitorsEV=[0 for i in range(NbrOfCS)]
            # Assign to the vector loads the information of each load
            
            if EV_distribution == 0:
                
                for index_loadEV in range(NbrOfCS):
                    indexBus=np.where(dfLoads['Name']==NamesOfLoads[index_loadEV])[0][0]
                    loadsEV[index_loadEV] = "New Load."+NamesOfLoads[index_loadEV]+"_EV phases=3 bus1="+str(dfLoads["Bus"][indexBus])+" conn=wye kv=0.4 daily="+NamesOfLoads[index_loadEV]+"_EV"
            
            elif EV_distribution == 1:
                
                for index_loadEV in range(NbrOfCS):
                    indexBus=np.where(dfDistr['Name']==NamesOfLoads[index_loadEV])[0][0]
                    loadsEV[index_loadEV] = "New Load."+NamesOfLoads[index_loadEV]+"_EV phases=3 bus1="+str(dfDistr["Bus"][indexBus])+" conn=wye kv=0.4 daily="+NamesOfLoads[index_loadEV]+"_EV"
            
            elif EV_distribution == 2:
                
                for index_loadEV in range(len(dfDistr)):
                    loadsEV[index_loadEV] = "New Load."+NamesOfLoads[index_loadEV]+"_EV phases=3 bus1="+str(dfDistr["Bus"][index_loadEV])+" conn=wye kv=0.4 daily="+NamesOfLoads[index_loadEV]+"_EV"
                
                for index_loadEV in range(len(dfDistr),NbrOfCS):
                    indexBus=np.where(dfMix['Name']==NamesOfLoads[index_loadEV])[0][0]
                    loadsEV[index_loadEV] = "New Load."+NamesOfLoads[index_loadEV]+"_EV phases=3 bus1="+str(dfMix["Bus"][indexBus])+" conn=wye kv=0.4 daily="+NamesOfLoads[index_loadEV]+"_EV"
            
            # Write the .dss file used then in OpenDSS to define loads
            loads_file = open(pathElement+"LoadsEV.txt","w")
            for i in range(NbrOfCS):
                loads_file.write(loadsEV[i]+"\n")
            loads_file.close()
            
            for index_monitor in range(NbrOfCS):
                monitorsEV[index_monitor] = "New monitor.EV_"+NamesOfLoads[index_monitor]+" element=Load."+NamesOfLoads[index_monitor]+"_EV terminal=1 mode=1 ppolar=no"
            # Write the .dss file used then in OpenDSS to define lines
            monitors_file = open(pathElement+"MonitorEV.txt","w")
            for i in range(NbrOfCS):
                monitors_file.write(str(monitorsEV[i])+"\n")
            monitors_file.close()  

            #***********************************************************
            # RUN SIMULATION IN OPENDSS
            #***********************************************************      
            #Compile OpenDSS script
            dssText.Command = "compile '"+pathDSS+"'"        
            
            dssText.Command ="redirect LoadshapeEV.txt"
            dssText.Command ="redirect LoadsEV.txt"
            dssText.Command ="redirect MonitorEV.txt"
            #Define timeseries calculation parameters
            dssText.Command = "set mode=daily"
            dssText.Command ="set stepsize="+str(int(timestep)*60)
            dssText.Command ="set number="+str(timeperiod)
            
            #Solve timeseries calculation
            dssText.Command ="solve"
                   
            #Export results
            #dssText.Command="Plot monitor object=Line_Line_1 channels=(9)"
            
            dssText.Command="Export monitors all"
            
            #***********************************************************
            # RESULTS ANALYSIS
            #***********************************************************
            #Check the number of .csv files have been export from OpenDSS
            if export_result != 1:
                listOfFiles = os.listdir(pathElement)
                pattern = "*.csv"
                ListOfCSV =list()
                index_csv=0
                for entry in listOfFiles:
                    if fnmatch.fnmatch(entry, pattern):
                            ListOfCSV.insert(index_csv, entry)
                            index_csv = index_csv+1
                 
                LineName=list()
                BusLineFromName=list()
                VLoadName=list()
                PLoadName=list()
                TrafoNames=list()
                
                PLoad=list()
                VLoad=list()
                LLine=list()
                VBus=list()
                ILine=list()
                ITrafo=list()
                index_line=0
                index_volt=0
                index_power=0
                index_VLoad=0
                index_trafo=0
                
                dfLoadingLine=dfLines.sort_values(by=['Name']).reset_index(drop=True)
                
                for i in range(len(dfLoads)):
                    dfLoads['Name'][i]=dfLoads['Name'][i]+'_1' 
                dfVoltageLoad=dfLoads.sort_values(by=['Name']).reset_index(drop=True)
                
                #Call each .csv file in the folder and exctract current or voltage based on the element connected to the monitor (if line or load)
                for i in range(len(ListOfCSV)):
                
                    file = ListOfCSV[i]
                    
                    if 'line' in file:
                        # Current and loading
                        current=current_line_results(fileName=file)
                        loadingLine=(current/dfLoadingLine['Fuse limit'][index_line])*100
                        #print(file+' / '+str(dfLoadingLine['Fuse limit'][index_line]))
                        ILine.insert(index_line,current)
                        LLine.insert(index_line,loadingLine)
                        LineName.insert(index_line,dfLoadingLine['Name'][index_line])
                        index_line=index_line+1
                        
                        # Voltage at the starting point of each line
                        voltage=voltage_line_results(fileName=file)
                        busVoltage=voltage/(400/math.sqrt(3))
                        VBus.insert(index_volt,busVoltage)
                        BusLineFromName.insert(index_volt,dfLoadingLine['From'][index_volt])
                        index_volt=index_volt+1
                    
                    elif 'vload' in file:
                        #Voltage at the buses where loads are connected
                        voltage=voltage_load_results(fileName=file)
                        LoadBus=voltage/(400/math.sqrt(3))
                        VLoad.insert(index_power,LoadBus)
                        VLoadName.insert(index_VLoad,dfVoltageLoad['Bus'][index_VLoad])
                        index_VLoad=index_VLoad+1
                    
                    elif 'pload' in file:
                        power=power_load_results(fileName=file)
                        PLoad.insert(index_power,power)
                        PLoadName.insert(index_power,dfVoltageLoad['Name'][index_power].replace('_1',''))
                        index_power=index_power+1
                    
                    elif 'trafo' in file:
                        currentHV=current_trafo_results(fileName=file)
                        I_nom = dfTrafo['Snom kVA'][index_trafo]/(dfTrafo['HV kV'][index_trafo]*math.sqrt(3)) #Approximation --> only active power: cos(phi)=1
                        Loading_Trafo=(currentHV/I_nom)*100
                        ITrafo.insert(index_trafo,Loading_Trafo)
                        TrafoNames.insert(index_trafo,dfTrafo['Name'][index_trafo])
                        index_trafo=index_trafo+1
                    
                    if export_result==0:
                        os.remove(file)
                
                #Converting data from a vector to a dataframe. In this way is easy to export the results as .csv file
                LineLoading=pd.DataFrame(data=LLine)
                LineCurrent=pd.DataFrame(data=ILine)
                BusVoltage=pd.DataFrame(data=VBus)
                LoadVoltage=pd.DataFrame(data=VLoad)
                LoadPower=pd.DataFrame(data=PLoad)
                LoadingTrafo = pd.DataFrame(data=ITrafo)
                
                LoadVoltage=LoadVoltage.transpose()
                BusVoltage=BusVoltage.transpose()
                LineLoading=LineLoading.transpose() 
                LineCurrent=LineCurrent.transpose()
                LoadPower=LoadPower.transpose()
                LoadingTrafo=LoadingTrafo.transpose()
                #Change dataframe columns names with the name of lines or load based
                LineCurrent.columns=LineName
                LineLoading.columns=LineName
                BusVoltage.columns=BusLineFromName 
                LoadVoltage.columns=VLoadName
                LoadPower.columns=PLoadName
                LoadingTrafo.columns=TrafoNames
                
                
                
                LineLoading.to_csv(pathRes+"/LineLoading/Lines_Loading_Rep"+str(index_repeat_sim)+"_Combo"+str(EVDistributionNbr)+".csv", sep=',',index=False)
                #LineCurrent.to_csv(pathRes+"/Lines_Current_Rep"+str(index_repeat_sim)+"_Combo"+str(EVDistributionNbr)+".csv", sep=',',index=False)
                #BusVoltage.to_csv(pathRes+"/Voltage_Lines_Rep"+str(index_repeat_sim)+"_Combo"+str(EVDistributionNbr)+".csv", sep=',',index=False)
                LoadVoltage.to_csv(pathRes+"/BusVoltage/Loads_Voltages_Rep"+str(index_repeat_sim)+"_Combo"+str(EVDistributionNbr)+".csv", sep=',',index=False)
                #LoadPower.to_csv(pathRes+"/Loads_PowerProfile_Rep"+str(index_repeat_sim)+"_Combo"+str(EVDistributionNbr)+".csv", sep=',',index=False)
                LoadingTrafo.to_csv(pathRes+"/TrafoLoading/Loading_Transformers_Rep"+str(index_repeat_sim)+"_Combo"+str(EVDistributionNbr)+".csv", sep=',',index=False)
                
                
            for i in range(len(dfLoads)):
                dfLoads['Name'][i]=dfLoads['Name'][i][:len(dfLoads['Name'][i])-2]
    
    LoadsToTXT.to_csv(pathRes+'/'+str(len(ListOfCombos))+'_Combinations.csv', sep=',',index=False)

else:
    #***********************************************************
    # RUN SIMULATION IN OPENDSS
    #***********************************************************      
    #Compile OpenDSS script
  
    dssText.Command = "compile '"+pathDSS+"'"
      
    #Define timeseries calculation parameters
    dssText.Command = "set mode=daily"
    dssText.Command ="set stepsize="+str(int(timestep)*60)
    dssText.Command ="set number="+str(timeperiod)
    
    #Solve timeseries calculation
    dssText.Command ="solve"
           
    #Export results
    dssText.Command="Plot monitor object=Line_Line_22_ channels=(9)"
    dssText.Command="Export monitors all"
    #dssText.Command="Export monitors all"
        #***********************************************************
    # RESULTS ANALYSIS
    #***********************************************************
    #Check the number of .csv files have been export from OpenDSS
    if export_result != 1:
        listOfFiles = os.listdir(pathElement)
        pattern = "*.csv"
        ListOfCSV =list()
        index_csv=0
        for entry in listOfFiles:
            if fnmatch.fnmatch(entry, pattern):
                    ListOfCSV.insert(index_csv, entry)
                    index_csv = index_csv+1
         
        LineName=list()
        BusLineFromName=list()
        VLoadName=list()
        PLoadName=list()
        TrafoNames=list()
        
        PLoad=list()
        VLoad=list()
        LLine=list()
        VBus=list()
        ILine=list()
        ITrafo=list()
        index_line=0
        index_volt=0
        index_power=0
        index_VLoad=0
        index_trafo=0
        
        dfLoadingLine=dfLines.sort_values(by=['Name']).reset_index(drop=True)
        
        for i in range(len(dfLoads)):
            dfLoads['Name'][i]=dfLoads['Name'][i]+'_1' 
        dfVoltageLoad=dfLoads.sort_values(by=['Name']).reset_index(drop=True)
        
        #Call each .csv file in the folder and exctract current or voltage based on the element connected to the monitor (if line or load)
        for i in range(len(ListOfCSV)):
        
            file = ListOfCSV[i]
            
            if 'line' in file:
                # Current and loading
                current=current_line_results(fileName=file)
                loadingLine=(current/dfLoadingLine['Fuse limit'][index_line])*100
                #print(file+' / '+str(dfLoadingLine['Fuse limit'][index_line]))
                ILine.insert(index_line,current)
                LLine.insert(index_line,loadingLine)
                LineName.insert(index_line,dfLoadingLine['Name'][index_line])
                index_line=index_line+1
                
                # Voltage at the starting point of each line
                voltage=voltage_line_results(fileName=file)
                busVoltage=voltage/(400/math.sqrt(3))
                VBus.insert(index_volt,busVoltage)
                BusLineFromName.insert(index_volt,dfLoadingLine['From'][index_volt])
                index_volt=index_volt+1
            
            elif 'vload' in file:
                #Voltage at the buses where loads are connected
                voltage=voltage_load_results(fileName=file)
                LoadBus=voltage/(400/math.sqrt(3))
                VLoad.insert(index_power,LoadBus)
                VLoadName.insert(index_VLoad,dfVoltageLoad['Bus'][index_VLoad])
                index_VLoad=index_VLoad+1
            
            elif 'pload' in file:
                power=power_load_results(fileName=file)
                PLoad.insert(index_power,power)
                PLoadName.insert(index_power,dfVoltageLoad['Name'][index_power].replace('_1',''))
                index_power=index_power+1
            
            elif 'trafo' in file:
                currentHV=current_trafo_results(fileName=file)
                I_nom = dfTrafo['Snom kVA'][index_trafo]/(dfTrafo['HV kV'][index_trafo]*math.sqrt(3)) #Approximation --> only active power: cos(phi)=1
                Loading_Trafo=(currentHV/I_nom)*100
                ITrafo.insert(index_trafo,Loading_Trafo)
                TrafoNames.insert(index_trafo,dfTrafo['Name'][index_trafo])
                index_trafo=index_trafo+1
            
            if export_result==0:
                os.remove(file)
        
        #Converting data from a vector to a dataframe. In this way is easy to export the results as .csv file
        LineLoading=pd.DataFrame(data=LLine)
        LineCurrent=pd.DataFrame(data=ILine)
        BusVoltage=pd.DataFrame(data=VBus)
        LoadVoltage=pd.DataFrame(data=VLoad)
        LoadPower=pd.DataFrame(data=PLoad)
        LoadingTrafo = pd.DataFrame(data=ITrafo)
        
        LoadVoltage=LoadVoltage.transpose()
        BusVoltage=BusVoltage.transpose()
        LineLoading=LineLoading.transpose() 
        LineCurrent=LineCurrent.transpose()
        LoadPower=LoadPower.transpose()
        LoadingTrafo=LoadingTrafo.transpose()
        #Change dataframe columns names with the name of lines or load based
        LineCurrent.columns=LineName
        LineLoading.columns=LineName
        BusVoltage.columns=BusLineFromName 
        LoadVoltage.columns=VLoadName
        LoadPower.columns=PLoadName
        LoadingTrafo.columns=TrafoNames
        
        
        
        LineLoading.to_csv(pathRes+"/Lines_Loading_.csv", sep=',',index=False)
        LineCurrent.to_csv(pathRes+"/Lines_Current.csv", sep=',',index=False)
        BusVoltage.to_csv(pathRes+"/Voltage_Lines.csv", sep=',',index=False)
        LoadVoltage.to_csv(pathRes+"/Loads_Voltages.csv", sep=',',index=False)
        LoadPower.to_csv(pathRes+"/Loads_PowerProfile.csv", sep=',',index=False)
        LoadingTrafo.to_csv(pathRes+"/Loading_Transformers.csv", sep=',',index=False)
        



