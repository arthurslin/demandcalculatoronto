import pandas as pd
import math
import glob
import os
from collections import defaultdict
from datetime import date

demanddir = 'demandchart'
invaliddir = 'invalid'
inventorydir = 'inventory'
oordir = 'oor'
veloctydir = 'velocity'
itemmasterdir = 'itemmaster'
directories = [demanddir, invaliddir, inventorydir, oordir, veloctydir, itemmasterdir]


def load_data():
    to_return = []
    for i in directories:
        path = glob.glob(os.path.join(i, "*xlsx"))
        if not path:
            raise FileNotFoundError(i, "File not found")
        xl = pd.ExcelFile(path[0])
        df = pd.read_excel(path[0], sheet_name=xl.sheet_names[0])
        to_return.append(df)
    return to_return


def create_report(data):
    # Extracting data from input tuple
    demand, invalid, inventory, oor, vel, itemmaster = data
    
    # Processing invalid items
    invalid_items = invalid["Invalid Locations"].tolist()
    inventory["SubInv"] = inventory["SubInv"].fillna("")
    inventory = inventory[~inventory["SubInv"].str.contains('|'.join(invalid_items))]
    
    # Summarizing inventory quantities
    q_dict = inventory.groupby("Item Number", sort=False)["Item Qty"].sum(numeric_only=False).to_dict()
    
    # Preparing demand dictionary
    demand["Name"] = demand["Name"].astype(str)
    demand_dict = demand.set_index("Name")["Median Demand"].to_dict()
    
    # Preparing velocity dictionary
    vel["PART_NUMBER"] = vel["PART_NUMBER"].astype(str)
    vel_dict = vel.set_index("PART_NUMBER")["Event Class"].to_dict()

    # Preparing item master dictionary
    itemmaster["Item"] = itemmaster["Item"].astype(str)
    itemmaster_dict = itemmaster.set_index("Item")["Cumulative Total LT"].to_dict()
    
    # Summarizing out-of-reach quantities
    df2 = oor.groupby("Item Code", sort=False)['Quantity Due'].sum(numeric_only=False)
    df2d = df2.to_dict()
    oor["Quantity Due"] = oor["Item Code"].apply(lambda x: df2d.get(x))
    oor["Item Code"] = oor["Item Code"].astype(str)
    
    # Converting dictionaries to defaultdicts for safe access
    demand_dict = defaultdict(int, demand_dict)
    q_dict = defaultdict(int, q_dict)
    vel_dict = defaultdict(lambda: "Zero Demand", vel_dict)
    itemmaster_dict = defaultdict(int,itemmaster_dict)
    
    # Processing each row in oor DataFrame
    for i in range(len(oor)):
        pn, quantity, price = oor.at[i, "Item Code"], oor.at[i, "Quantity Due"], oor.at[i, "PO Price"]
        
        if pd.isna(pn):
            continue
        
        # Assigning Part Cost category
        oor.at[i, "Part Cost"] = "High $" if price > 5000 else ("Low $" if price < 1000 else "Mid $")
        
        # Calculating expected MOS and updating relevant columns
        expected_mos = q_dict[pn] / demand_dict[pn] if demand_dict[pn] != 0 else 0
        after_push = q_dict[pn] - quantity if not pd.isna(quantity) else q_dict[pn]
        
        oor.at[i, 'Expected MOS'] = round(expected_mos, 2)
        oor.at[i, "Median Demand"] = demand_dict[pn]
        oor.at[i, "Valid Inventory"] = q_dict[pn]
        oor.at[i, 'Inventory Qty - Qty Due'] = after_push
        new_date = (pd.to_datetime(date.today()) + pd.DateOffset(days=math.floor(expected_mos*30))).date()
        oor.at[i, "Month to Zero Inventory"] = new_date
        oor.at[i, "Cumulative Total LT"] = itemmaster_dict[pn]
        # Updating Part Velocity and Part Class
        oor.at[i, "Part Velocity"] = vel_dict[pn]
        oor.at[i, "Part Class"] = oor.at[i, "Part Cost"].replace(" Part", "") + " " + oor.at[i, "Part Velocity"]

    oor.drop(oor[oor['Org Code'] == 'NUS'].index, inplace=True)
    oor = oor.drop_duplicates(subset="Item Code")

    oor.to_excel("mosreport.xlsx", index=False)

print("Starting Generation.")
create_report(load_data())
print("Generation Complete")
