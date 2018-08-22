import requests
import pandas as pd
import numpy as np

GOOGLE_MAPS_API_URL = 'http://maps.googleapis.com/maps/api/geocode/json'

def get_lat_long(address):
    for x in range(0,10):
        try:

            params = {
                'address': address,
                'sensor': 'false',
                'region': 'united states'
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



df = pd.read_excel("RGC.xlsx")

lats = []
longs = []

for x in df["Address"].values:

    lat,long= get_lat_long(x)
    lats.append(lat)
    longs.append(long)

lats = np.array(lats)
longs = np.array(longs)

df["lats"] = lats
df["longs"] = longs

writer = pd.ExcelWriter('RGC Data.xlsx')

df.to_excel(writer,'Data')

writer.save()
