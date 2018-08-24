import pandas as pd
import plotly
from pathlib import Path

AMC_df = pd.read_excel("Geocoded Data/AMC Geocoded Data.xlsx")
CNK_df = pd.read_excel("Geocoded Data/CNK Geocoded Data.xlsx")
RGC_df = pd.read_excel("Geocoded Data/RGC Geocoded Data.xlsx")

data = dict(
        type = 'scattergeo',
        locationmode = 'USA-states',
        mode = 'markers'
        )

data_CNK = data.copy()
data_CNK['lon'] = CNK_df['longs']
data_CNK['lat'] = CNK_df['lats']
data_CNK['marker'] = dict(color = 'green')
data_CNK['name'] = 'CNK'

data_AMC = data.copy()
data_AMC['lon'] = AMC_df['longs']
data_AMC['lat'] = AMC_df['lats']
data_AMC['marker'] = dict(color = 'red')
data_AMC['name'] = 'AMC'


data_RGC = data.copy()
data_RGC['lon'] = RGC_df['longs']
data_RGC['lat'] = RGC_df['lats']
data_RGC['marker'] = dict(color = 'blue')
data_RGC['name'] = 'RGC'


layout = dict(
        geo = dict(
            scope = 'usa',
            projection = dict(type='albers usa'),
        ),
    )

fig = dict(data=[data_AMC, data_CNK, data_RGC], layout=layout)
plot_url = plotly.plotly.plot(fig)

f = open("plot_url.txt","w+")
f.write(plot_url)
f.close()
