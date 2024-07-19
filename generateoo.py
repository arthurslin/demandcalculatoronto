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
    print(inventory)
    q_dict = inventory.groupby("Item Number",sort=False)["Item Qty"].sum(numeric_only=False).to_dict()

    # q_dict = inventory.set_index("Item Number")["Item Qty"].to_dict()
    demand_dict = demand.set_index("Name")["Median Demand"].to_dict()
    vel_dict = vel.set_index("PART_NUMBER")["Event Class"].to_dict()
    print(vel_dict)

    df2 = oor.groupby("Item Code", sort=False)[
        'Quantity Due'].sum(numeric_only=False)

    df2d = df2.to_dict()
    oor["Quantity Due"] = oor["Item Code"].apply(lambda x: df2d.get(x))

    for i in range(len(oor)):
        pn, quantity,price = oor.iloc[i, 7], oor.iloc[i, 12], oor.iloc[i,17]
        if pd.isna(pn) or pd.isna(quantity):
            continue
        if pn in q_dict and pn in demand_dict:
            # Subinventory quantity / Monthy Demand mean
            expected_mos = q_dict[pn]/demand_dict[pn]
            after_push = q_dict[pn] - quantity
            oor.at[i, 'Expected MOS'] = expected_mos.round(0.1)
            oor.at[i, 'Inventory Qty - Qty Due'] = after_push
            oor.at[i, "Median Demand"] = demand_dict[pn]
            oor.at[i, "Valid Inventory"] = q_dict[pn]
            thisdate = date.today()
            new_date = (pd.to_datetime(thisdate)+pd.DateOffset(days=math.floor(expected_mos*30))).date()
            oor.at[i, "Month to Zero Inventory"] = new_date
        if pn in vel_dict:
            oor.at[i,"Part Velocty"] = vel_dict[pn]
        if price > 5000:
            pass
        elif price < 1000:
            pass
        else: # 1000 < price < 5000
            pass

    oor = oor.drop_duplicates(subset="Item Code")

    oor.to_excel("mosreport.xlsx", index=False)


create_report(load_data())
