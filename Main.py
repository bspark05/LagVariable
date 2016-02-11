'''
Created on Feb 10, 2016

@author: Bumsub
'''
import FileIO.Excel as excel
from math import hypot
import collections
import math

if __name__ == '__main__':
    ## Read excel file##
    # filename 1 - previous year including logged UnitPrice
    # filename 2 - current year
    
    filename1 = "SaleApartment201303"
    filepath1 = filename1+".xlsx"
    sheetname1 = filename1
    
    filename2 = "SaleApartment201403"
    filepath2 = filename2+".xlsx"
    sheetname2 = filename2
    
    preVar = excel.excelRead(filepath1, sheetname1)
    curVar = excel.excelRead(filepath2, sheetname2)
        
        # structure of the variable reading excel sheet
        # 1st row - field name list - preVar[0] (list format)
        # 2nd row ~ - values
    
    ## Setting neighborhood condition
    bandwidth = 2000.0 # setting bandwidth (maximum distance)
    mindist = -1.0 # setting minimum distance to avoid identical apartment complex itself
    knn = 10 # the maximum number of neighbors
    
    lagVar = [] # initialize lag variable list (final column)
        
    for curValue in curVar[1:]: # 2nd row ~
        curX = float(curValue[0].value) # X coordinate of current year data
        curY = float(curValue[1].value) # Y coordinate of current year data
        curOID = int(curValue[2].value) # OBJECT ID of current year data
        
        distDic = dict() # define a dictionary (key:OBJECT ID, value:distance) 
        
        nCount = 1 # initializing the number of neighborhoods
        
        for preValue in preVar[1:]: # 2nd row ~
            preX = float(preValue[0].value) # X coordinate of previous year data
            preY = float(preValue[1].value) # Y coordinate of previous year data
            preOID = int(preValue[2].value) # OBJECT ID of previous year data
            preUnitPrice = float(preValue[22].value) # Unit Price of previous year data
             
            distance = math.sqrt((curX - preX)**2+(curY - preY)**2) # distance between an observation in current data and each points in previous one 
            
            # add observation satisfying the neighborhood condition 
            if distance < bandwidth :   # 1. bandwidth condition
                if distance > mindist:  # 2. minimum distance condition
                    if nCount <= knn:    # 3. number of maximum neighborhoods
                        distDic[preOID] = [distance, preUnitPrice]
                        nCount+=1
            
        distDicsorted = collections.OrderedDict(sorted(distDic.items(), key=lambda t:t[1][0]))
        
#         print distDicsorted.values()[0][0]
        
        ## calculating Wij & Wij*UnitPricej for current year observations
        sumWij = 0  # initialize sigma Wij
        sumWijP = 0 # initialize sigma Wij*UnitP
        for dist, UnitP in distDicsorted.values():
            wij = (1-(dist/bandwidth)**2)**2
            wijP = wij*UnitP
            sumWij += wij
            sumWijP += wijP
        lag = sumWijP / sumWij
        lagVar.append(lag)
     
    excel.excelWriteOnExistingFile3(filepath2, sheetname2, 'X', lagVar)
    print(lagVar)
        
            