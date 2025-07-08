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

#getting all users not in namely
user_dict = {"Unverified Name": [], "Verizon Phone Number": []}
for index, row in verizon_df.iterrows():
    user = row["User name"]
    if user.lower() not in namely_active["Full name"].values and user.lower() not in namely_active["Nick name"].values:
        user_dict["Unverified Name"].append(user)
        user_dict["Verizon Phone Number"].append(row["Wireless number"])

#convert dict to df
unverified_df = pd.DataFrame(user_dict)

#import intune sheet into a dataframe
intune_df = pd.read_excel(r"C:\Users\jgupta\OneDrive - Collegium Pharma\Intune iOS Export - Match Phone Numbers with VZW.xlsx")

#remove uneccesary columns/rows in intune df
intune_df = intune_df[["Primary user display name", "Phone number", "IMEI"]]
num1 = len(intune_df)
intune_df.dropna(inplace=True)
num2 = len(intune_df)
intune_df["Phone number"] = intune_df["Phone number"].astype(str)
print(f"# Of Rows Dropped In Intune DF: {num1-num2}")

#convert data to match intune phone number format
unverified_df["Formatted Verizon Phone Number"] = unverified_df["Verizon Phone Number"].apply(lambda num: "1" + num.replace("-",""))
unverified_df["Intune Name"] = None
unverified_df["Intune IMEI"] = None

#gathering Intune IMEI and name info 
for index, row in unverified_df.iterrows():
    number = row["Formatted Verizon Phone Number"]
    if number in intune_df["Phone number"].values:
       values = intune_df.loc[intune_df["Phone number"] == number, ["Primary user display name", "IMEI"]].values
       unverified_df.at[index, "Intune Name"] = values[0][0]
       unverified_df.at[index, "Intune IMEI"] = values[0][1]

#write final results to excel
unverified_df.to_excel(r"C:\Users\jgupta\OneDrive - Collegium Pharma\Verizon Updates.xlsx", index=False)
print("Written To Excel Successfully!")

user_dict = {"unassigned": [], "nonexistent in namely": [], "not active": []} 
null_rows = unverified_df[unverified_df["Intune Name"].isnull()]

#finding unrecognized users and bucketing into correct description
for index,row in null_rows.iterrows():
    user = row["Unverified Name"]
    if user.lower() not in namely_active["Full name"].values and user.lower() not in namely_active["Nick name"].values:
        if "user" in user.lower():
            user_dict["unassigned"].append((user, row["Verizon Phone Number"]))
        elif user.lower() not in namely_df["Full name"].values and user.lower() not in namely_df["Nick name"].values:
            user_dict["nonexistent in namely"].append((user, row["Verizon Phone Number"]))
        else:
            user_dict["not active"].append((user, row["Verizon Phone Number"]))

#finding potential misspelled names and putting them in their own category
misspelled_list = []
for user in user_dict["nonexistent in namely"]:
    email = ""
    fullname_list = user[0].lower().split(" ")
    if len(fullname_list) > 2:
        email = fullname_list[0][0:1] + fullname_list[1][0:1] + fullname_list[len(fullname_list) - 1]
    else:
        email = fullname_list[0][0:1] + fullname_list[len(fullname_list) - 1]
    
    email += "@collegiumpharma.com"
    if email in namely_df["Company email"].values:
        misspelled_list.append(user)

user_dict["nonexistent in namely"] = list(set(user_dict["nonexistent in namely"]) - set(misspelled_list))
user_dict["potentially misspelled"] = misspelled_list

#final output    
print("Unverified Users In Namely\n\n")
for desc in user_dict:
    print(f"{desc}:")
    user_dict[desc].sort()
    for user in user_dict[desc]:
        print(user)
    print()

