import pandas as pd
from math import sin, cos, sqrt, atan2, radians

#read excel sheets with all the data
AMC_all_df = pd.read_excel("AMC Data.xlsx")
CNK_all_df = pd.read_excel("CNK Data.xlsx")
RGC_all_df = pd.read_excel("RGC Data.xlsx")

#seperate the premium loctations for each brand
AMC_premium_df = AMC_all_df[AMC_all_df["Seating"] == 1.0]
CNK_premium_df = CNK_all_df[CNK_all_df["Loungers"] == 1.0]
RGC_premium_df = RGC_all_df[RGC_all_df["Loungers"] == 1.0]

#seperate the non premium locations for each brand
AMC_reg_df = AMC_all_df[AMC_all_df["Seating"] != 1.0]
CNK_reg_df = CNK_all_df[CNK_all_df["Loungers"] != 1.0]
RGC_reg_df = RGC_all_df[RGC_all_df["Loungers"] != 1.0]


#formula to get dinstance beetween two geo coordiantes
def get_distance(lat1,lon1,lat2,lon2):
    R = 3959

    lat1 = radians(lat1)
    lon1 = radians(lon1)
    lat2 = radians(lat2)
    lon2 = radians(lon2)

    dlon = lon2 - lon1
    dlat = lat2 - lat1

    a = sin(dlat / 2)**2 + cos(lat1) * cos(lat2) * sin(dlon / 2)**2
    c = 2 * atan2(sqrt(a), sqrt(1 - a))

    distance = R * c
    return distance


#Functions calculate the percent of premium locations that have at least "num_loc"
#number of premium location from any company in a "rad" mile radius

def AMC_percent_in_rad_premium(num_loc,rad):

    num_locations_with_intersects = []
    for amc_location in AMC_premium_df.iterrows():
        current_intersects = 0
        check_cnk = True
        for rgc_location in RGC_premium_df.iterrows():
            if get_distance(rgc_location[1]["lats"],rgc_location[1]["longs"],amc_location[1]["lats"],amc_location[1]["longs"]) <= rad:
                current_intersects += 1
            if current_intersects == num_loc:
                num_locations_with_intersects.append(amc_location[1]["Address"])
                check_cnk = False
                break
        if check_cnk:
            for cnk_location in CNK_premium_df.iterrows():
                if get_distance(cnk_location[1]["lats"],cnk_location[1]["longs"],amc_location[1]["lats"],amc_location[1]["longs"]) <= rad:
                    current_intersects += 1
                if current_intersects == num_loc:
                    num_locations_with_intersects.append(amc_location[1]["Address"])
                    break
    return len(num_locations_with_intersects)/AMC_premium_df["Address"].count()


def RGC_percent_in_rad_premium(num_loc,rad):

    num_locations_with_intersects = []
    for rgc_location in RGC_premium_df.iterrows():
        current_intersects = 0
        check_cnk = True
        for amc_location in AMC_premium_df.iterrows():
            if get_distance(rgc_location[1]["lats"],rgc_location[1]["longs"],amc_location[1]["lats"],amc_location[1]["longs"]) <= rad:
                current_intersects += 1
            if current_intersects == num_loc:
                num_locations_with_intersects.append(rgc_location[1]["Address"])
                check_cnk = False
                break
        if check_cnk:
            for cnk_location in CNK_premium_df.iterrows():
                if get_distance(cnk_location[1]["lats"],cnk_location[1]["longs"],rgc_location[1]["lats"],rgc_location[1]["longs"]) <= rad:
                    current_intersects += 1
                if current_intersects == num_loc:
                    num_locations_with_intersects.append(rgc_location[1]["Address"])
                    break
    return len(num_locations_with_intersects)/AMC_premium_df["Address"].count()


def RGC_percent_in_rad_premium(num_loc,rad):

    num_locations_with_intersects = []
    for rgc_location in RGC_premium_df.iterrows():
        current_intersects = 0
        check_cnk = True
        for amc_location in AMC_premium_df.iterrows():
            if get_distance(rgc_location[1]["lats"],rgc_location[1]["longs"],amc_location[1]["lats"],amc_location[1]["longs"]) <= rad:
                current_intersects += 1
            if current_intersects == num_loc:
                num_locations_with_intersects.append(rgc_location[1]["Address"])
                check_cnk = False
                break
        if check_cnk:
            for cnk_location in CNK_premium_df.iterrows():
                if get_distance(cnk_location[1]["lats"],cnk_location[1]["longs"],rgc_location[1]["lats"],rgc_location[1]["longs"]) <= rad:
                    current_intersects += 1
                if current_intersects == num_loc:
                    num_locations_with_intersects.append(rgc_location[1]["Address"])
                    break
    return len(num_locations_with_intersects)/AMC_premium_df["Address"].count()

print(RGC_percent_in_rad_premium(1,15))
