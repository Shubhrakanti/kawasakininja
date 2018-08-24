import requests
import pandas as pd
import numpy as np
import os
from pathlib import Path


GOOGLE_MAPS_API_URL = 'https://maps.googleapis.com/maps/api/geocode/json'

API_KEY = "AIzaSyBAKnd5BDKb52_GUgxkxj89W9xWyUxr-Pc"

def get_lat_long(address):
    for x in range(20):
        try:

            params = {
                'address': address,
                'sensor': 'false',
                'region': 'united states',
                'key': API_KEY
            }

            # Do the request and get the response data
            req = requests.get(GOOGLE_MAPS_API_URL, params=params)
            res = req.json()

            # Use the first result
            result = res['results'][0]

            #Get the lat and long from the request
            lat = result['geometry']['location']['lat']
            long = result['geometry']['location']['lng']

            print(address)
            return lat,long

        except:
            pass
    return "Error"
print(type(os.getcwd()))
print(os.getcwd())
newpath = 'Geocoded Data'
if not os.path.exists(newpath):
    os.makedirs(newpath)

def makedata(mtheat):
    my_file = Path(newpath+'/'+mtheat+' Geocoded Data.xlsx')
    print(my_file)
    if not my_file.is_file():
        df = pd.read_excel(mtheat+".xlsx")

        lats = []
        longs = []

        for x in df["Address"].values:
            print(get_lat_long(x))
            lat,long= get_lat_long(x)
            lats.append(lat)
            longs.append(long)

        lats = np.array(lats)
        longs = np.array(longs)

        df["lats"] = lats
        df["longs"] = longs

        writer = pd.ExcelWriter(mtheat+'Geocoded Data.xlsx')

        df.to_excel(writer,'Data')

        writer.save()
        os.rename(os.getcwd()+ "/" +mtheat+ ' Geocoded Data.xlsx', newpath+'/' + mtheat + ' Geocoded Data.xlsx')
    else:
        dfold = pd.read_excel(my_file)
        dfnew = pd.read_excel(mtheat+".xlsx")

        dfnewent = pd.concat([dfold, dfnew])
        dfnewent = dfnewent.drop_duplicates(subset = ['Address'], keep=False)
        print(dfnewent)

        lats = []
        longs = []

        for x in dfnewent["Address"].values:

            lat,long= get_lat_long(x)
            lats.append(lat)
            longs.append(long)

        lats = np.array(lats)
        longs = np.array(longs)

        dfnewent["lats"] = lats
        dfnewent["longs"] = longs
        df= pd.concat([dfold, dfnewent], ignore_index=True)
        os.remove(my_file)

        writer = pd.ExcelWriter(mtheat+' Geocoded Data.xlsx')

        df.to_excel(writer,'Data')
        writer.save()
        os.rename(os.getcwd()+ "/" +mtheat+ ' Geocoded Data.xlsx', newpath+'/' + mtheat + ' Geocoded Data.xlsx')


        # df_all = df.merge(dfold.drop_duplicates(),
        #            how='left', indicator=True)
        # newentry =





makedata("RGC")

# df = pd.read_excel("RGC.xlsx")
#
# lats = []
# longs = []
#
# for x in df["Address"].values:
#
#     lat,long= get_lat_long(x)
#     lats.append(lat)
#     longs.append(long)
#
# lats = np.array(lats)
# longs = np.array(longs)
#
# df["lats"] = lats
# df["longs"] = longs
#
# writer = pd.ExcelWriter('RGC Data.xlsx')
#
# df.to_excel(writer,'Data')
#
# writer.save()
