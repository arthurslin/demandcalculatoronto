import pandas as pd
import math
import glob
import os
from datetime import date

demanddir = 'demandchart'
invaliddir = 'invalid'
inventorydir = 'inventory'
oordir = 'oor'
veloctydir = 'velocity'
directories = [demanddir, invaliddir, inventorydir, oordir, veloctydir]


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
    demand, invalid, inventory, oor, vel = data[0], data[1], data[2], data[3], data[4]

    invalid_items = invalid["Invalid Locations"].tolist()
    inventory["SubInv"] = inventory["SubInv"].fillna("")
    inventory = inventory[~inventory["SubInv"].str.contains(
        '|'.join(invalid_items))]

    q_dict = inventory.groupby("Item Number", sort=False)[
        "Item Qty"].sum(numeric_only=False).to_dict()

    # q_dict = inventory.set_index("Item Number")["Item Qty"].to_dict()
    demand_dict = demand.set_index("Name")["Median Demand"].to_dict()
    vel["PART_NUMBER"] = vel["PART_NUMBER"].astype(str)
    vel_dict = vel.set_index("PART_NUMBER")["Event Class"].to_dict()
    print(vel_dict)

    df2 = oor.groupby("Item Code", sort=False)[
        'Quantity Due'].sum(numeric_only=False)

    df2d = df2.to_dict()
    oor["Quantity Due"] = oor["Item Code"].apply(lambda x: df2d.get(x))

    for i in range(len(oor)):
        pn, quantity, price = str(oor.at[i, "Item Code"]), oor.at[i, "Quantity Due"], oor.at[i, "PO Price"]
        if pd.isna(pn) or pd.isna(quantity):
            continue
        if pn in q_dict and pn in demand_dict:
            # Subinventory quantity / Monthy Demand mean
            expected_mos = q_dict[pn]/demand_dict[pn]
            after_push = q_dict[pn] - quantity
            oor.at[i, 'Expected MOS'] = round(expected_mos, 2)
            oor.at[i, 'Inventory Qty - Qty Due'] = after_push
            oor.at[i, "Median Demand"] = demand_dict[pn]
            oor.at[i, "Valid Inventory"] = q_dict[pn]
            thisdate = date.today()
            new_date = (pd.to_datetime(thisdate) +
                        pd.DateOffset(days=math.floor(expected_mos*30))).date()
            oor.at[i, "Month to Zero Inventory"] = new_date
            action = "Cancel" if expected_mos >= 6 else "Push Out" if 3 <= expected_mos < 6 else "No Action"
            oor.at[i, "MOS Action"] = action
        oor.at[i, "Part Cost"] = "High $" if price > 5000 else (
            "Low $" if price < 1000 else "Mid $")
        if pn in vel_dict:
            oor.at[i, "Part Velocity"] = vel_dict[pn]
            oor.at[i, "Part Class"] = (oor.at[i, "Part Cost"]).replace(
                " Part", "") + " " + oor.at[i, "Part Velocity"]
        else:
            oor.at[i,"Part Velocity"] = "Zero Demand"
            oor.at[i, "Part Class"] = (oor.at[i, "Part Cost"]).replace(
                " Part", "") + " " + oor.at[i, "Part Velocity"]

    oor.drop(oor[oor['Org Code'] == 'NUS'].index, inplace=True)
    oor = oor.drop_duplicates(subset="Item Code")

    oor.to_excel("mosreport.xlsx", index=False)


create_report(load_data())
