import pandas as pd
import numpy as np
#If you want to run, pip install openpyxl & pandas libraries

verizon_df = pd.read_excel(r"C:\Users\jgupta\OneDrive - Collegium Pharma\Verizon Device Report.xlsx")
namely_df = pd.read_excel(r"C:\Users\jgupta\OneDrive - Collegium Pharma\Namely Report - Roster Report (25).xlsx")
##print(namely_df)
##print(verizon_df)
##print(namely_df.columns)
##print(verizon_df.columns)
namely_active = namely_df[namely_df["Status"] == "Active"].copy() #filter based on active users
##print(namely_active)
namely_active["Full name"] = namely_active["First name"] + " " + namely_active["Last name"]
##print(namely_active[namely_active["Full name"] == "Shane St John"])
namely_active["Full name"] = namely_active["Full name"].apply(lambda name: name.lower().replace(".", ""))
##print(namely_active["Full name"])
namely_active["Nick name"] = np.where(namely_active["Preferred name"].isna() | namely_active["Preferred name"] == "", "", namely_active["Preferred name"] + " " + namely_active["Last name"])
namely_active["Nick name"] = namely_active["Nick name"].fillna("").apply(lambda name: name.lower().replace(".", ""))

print(namely_active["Full name"])
#print(namely_active[["First name", "Last name", "Full name", "Nick name"]])

user_list = []
for user in verizon_df["User name"]:
    print(user.lower())
    if user.lower() not in namely_active["Full name"].values and user.lower() not in namely_active["Nick name"].values:
        user_list.append(user)

print(user_list)
#print(len(user_list))
#print("tammy sanford" in namely_active["Full name"].values)
