# -*- coding: utf-8 -*-
"""
Created on Wed Nov  4 11:31:11 2020

@author: lucag
"""
# Function which allows to extract the results regarding lines (current and voltage). It returns the mean of current (per line)
# and voltage in Volt at the starting bus of the line
import pandas as pd

def current_line_results(fileName):
    res=pd.read_csv(fileName,delimiter=',')
    if len(res.columns) == 1 :
        res=pd.read_csv(fileName,delimiter=';')
    I1=res[' I1']
    I2=res[' I2']
    I3=res[' I3']

    I=(I1+I2+I3)/3         
    return I

def voltage_line_results(fileName):
    res=pd.read_csv(fileName,delimiter=',')
    if len(res.columns) == 1 :
        res=pd.read_csv(fileName,delimiter=';')
    
    U1=res[' V1']
    U2=res[' V2']
    U3=res[' V3']

    U=(U1+U2+U3)/3
         
    return U


# Function which allows to extract the results regarding loads (power and voltage). It returns the mean of voltage on the busbar
# and power profile of each load
def voltage_load_results(fileName):
    res=pd.read_csv(fileName,delimiter=',')
    if len(res.columns) == 1 :
        res=pd.read_csv(fileName,delimiter=';')
    U1=res[' V1']
    U2=res[' V2']
    U3=res[' V3']

    U=(U1+U2+U3)/3
    return U

def power_load_results(fileName):
    res=pd.read_csv(fileName,delimiter=',')
    if len(res.columns) == 1 :
        res=pd.read_csv(fileName,delimiter=';')
    P1=res[' P1 (kW)']
    P2=res[' P2 (kW)']
    P3=res[' P3 (kW)']

    P=(P1+P2+P3)
    return P
# Function which extract the value of current (on HV side) for transformer
def current_trafo_results(fileName):
    res=pd.read_csv(fileName,delimiter=',')
    if len(res.columns) == 1 :
        res=pd.read_csv(fileName,delimiter=';')
    I1=res[' I1']
    I2=res[' I2']
    I3=res[' I3']

    I=(I1+I2+I3)/3         
    return I


