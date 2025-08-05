import pandas as pd
import numpy as np
from collections import Counter
#If you want to run, pip install openpyxl, numpy & pandas libraries
#replace any paths to the documents with the path on your local machine

verizon_df = pd.read_excel(r"C:\Users\jgupta\OneDrive - Collegium Pharma\Verizon Device Report.xlsx")
namely_df = pd.read_excel(r"C:\Users\jgupta\OneDrive - Collegium Pharma\Namely Report - Roster Report (25).xlsx")

#Converting fName, lName, into a full name column or nick name column
namely_df["Full name"] = namely_df["First name"] + " " + namely_df["Last name"]
namely_df["Full name"] = namely_df["Full name"].apply(lambda name: name.lower())
namely_df["Nick name"] = np.where(namely_df["Preferred name"].isna() | namely_df["Preferred name"] == "", "", namely_df["Preferred name"] + " " + namely_df["Last name"])
namely_df["Nick name"] = namely_df["Nick name"].fillna("").apply(lambda name: name.lower())

namely_active = namely_df[namely_df["Status"] == "Active"].copy() #filter based on active users

#getting all unverified users
user_dict = {"Unverified Name (Verizon)": [], "Verizon Phone Number": [], "Reason For Update": []}
for index, row in verizon_df.iterrows():
    user = row["User name"]
    if user.lower() not in namely_active["Full name"].values and user.lower() not in namely_active["Nick name"].values:
        user_dict["Unverified Name (Verizon)"].append(user)
        user_dict["Verizon Phone Number"].append(row["Wireless number"])  
        #finding unrecognized users and bucketing into correct description
        if "user" in user.lower():
            user_dict["Reason For Update"].append("unassigned")
        elif user.lower() not in namely_df["Full name"].values and user.lower() not in namely_df["Nick name"].values:
            user_dict["Reason For Update"].append("nonexistent in namely")
        else:
            user_dict["Reason For Update"].append("not active")

print(f'# Of Unverified Names: {len(user_dict["Unverified Name (Verizon)"])}')
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
#print(f"# Of Rows Dropped In Intune DF: {num1-num2}")

#convert data to match intune phone number format
unverified_df["Unformatted Verizon Phone Number"] = unverified_df["Verizon Phone Number"].apply(lambda num: "1" + num.replace("-",""))
unverified_df["Verified Name (Intune)"] = None
unverified_df["Intune IMEI"] = None

#gathering Intune IMEI and name info 
for index, row in unverified_df.iterrows():
    number = row["Unformatted Verizon Phone Number"]
    if number in intune_df["Phone number"].values:
       values = intune_df.loc[intune_df["Phone number"] == number, ["Primary user display name", "IMEI"]].values
       unverified_df.at[index, "Verified Name (Intune)"] = values[0][0]
       unverified_df.at[index, "Intune IMEI"] = values[0][1]

#format final results
desired_order = ["Unverified Name (Verizon)", "Verified Name (Intune)", "Verizon Phone Number", "Unformatted Verizon Phone Number", "Intune IMEI", "Reason For Update"]
unverified_df = unverified_df.reindex(columns=desired_order)

#write final results to excel
unverified_df.to_excel(r"C:\Users\jgupta\OneDrive - Collegium Pharma\Verizon Updates.xlsx", index=False)
print("Written To Excel Successfully!\n\n")
user_dict = {"unassigned": [], "nonexistent in namely": [], "not active": []} 
null_rows = unverified_df[unverified_df["Verified Name (Intune)"].isnull()]

#finding unrecognized users and bucketing into correct description
for index,row in null_rows.iterrows():
    user = row["Unverified Name (Verizon)"]
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
print("Unverified Users In Namely\n")
for desc in user_dict:
    print(f"{desc}:")
    user_dict[desc].sort()
    for user in user_dict[desc]:
        print(user)
    print()

