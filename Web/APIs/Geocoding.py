#-*- coding: utf-8 -*-

#from geopy.geocoders import Nominatim
import geocoder

def geocodeList(addressExcelList):
    # -*- coding: cp949 -*-
    
    #geolocator = Nominatim()
    searchingAddress = ''
    resultList = []
    count = 0
    for rows in addressExcelList:
        searchingAddress=''
        
        searchingAddress+=rows[0].value
        searchingAddress+=' '
        searchingAddress+=str(rows[1].value)
        searchingAddress+='-'
        searchingAddress+=str(rows[2].value)
        #searchingAddress+=' '
        #searchingAddress+=str(rows[7].value)    
        
        #print(searchingAddress)
        location = geocoder.google(searchingAddress)
        print(searchingAddress)
        print(location.latlng)
        if location.latlng == [] :
            resultList.append(['',''])
        else:
            resultList.append(location.latlng)
            
        print(count)
        count+=1
    return resultList

