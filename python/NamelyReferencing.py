import pandas as pd
import numpy as np
from collections import Counter
#If you want to run, pip install openpyxl, numpy & pandas libraries

verizon_df = pd.read_excel(r"C:\Users\jgupta\OneDrive - Collegium Pharma\Verizon Device Report.xlsx")
namely_df = pd.read_excel(r"C:\Users\jgupta\OneDrive - Collegium Pharma\Namely Report - Roster Report (25).xlsx")

#Converting fName, lName, into a full name column or nick name column
namely_df["Full name"] = namely_df["First name"] + " " + namely_df["Last name"]
namely_df["Full name"] = namely_df["Full name"].apply(lambda name: name.lower())
namely_df["Nick name"] = np.where(namely_df["Preferred name"].isna() | namely_df["Preferred name"] == "", "", namely_df["Preferred name"] + " " + namely_df["Last name"])
namely_df["Nick name"] = namely_df["Nick name"].fillna("").apply(lambda name: name.lower())

namely_active = namely_df[namely_df["Status"] == "Active"].copy() #filter based on active users

user_list = {"unassigned": [], "nonexistent in namely": [], "not active": []} 

#finding unrecognized users and bucketing into correct description
for index,row in verizon_df.iterrows():
    user = row["User name"]
    if user.lower() not in namely_active["Full name"].values and user.lower() not in namely_active["Nick name"].values:
        if "user" in user.lower():
            user_list["unassigned"].append((user, row["Wireless number"]))
        elif user.lower() not in namely_df["Full name"].values and user.lower() not in namely_df["Nick name"].values:
            user_list["nonexistent in namely"].append((user, row["Wireless number"]))
        else:
            user_list["not active"].append((user, row["Wireless number"]))

#finding potential misspelled names and putting them in their own category
misspelled_list = []
for user in user_list["nonexistent in namely"]:
    email = ""
    fullname_list = user[0].lower().split(" ")
    if len(fullname_list) > 2:
        email = fullname_list[0][0:1] + fullname_list[1][0:1] + fullname_list[len(fullname_list) - 1]
    else:
        email = fullname_list[0][0:1] + fullname_list[len(fullname_list) - 1]
    
    email += "@collegiumpharma.com"
    if email in namely_df["Company email"].values:
        misspelled_list.append(user)

user_list["nonexistent in namely"] = list(set(user_list["nonexistent in namely"]) - set(misspelled_list))
user_list["potentially misspelled"] = misspelled_list

#final output    
for desc in user_list:
    print(f"{desc}:")
    user_list[desc].sort()
    for user in user_list[desc]:
        print(user)
    print()
