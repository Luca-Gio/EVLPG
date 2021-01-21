# -*- coding: utf-8 -*-
"""
Created on Fri Nov  6 13:09:38 2020

@author: lucag
"""
import numpy as np

#This script defines functions used to chose arrival time, connection time, energy demand and type of car
# based on probability density functions

def ArrTime(var,prob):
    res=np.random.choice(var,p=prob)
    return res

def Connection_Time(var,prob):
    res=np.random.choice(var,p=prob)
    return res

def Energy_Demand(var,prob):
    res=np.random.choice(var,p=prob)
    return res

def ElectricVehicle(var,prob):
    res=np.random.choice(var,p=prob)
    return res