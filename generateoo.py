import pandas as pd
import glob
import os

demanddir = "demandchart"
invaliddir = 'invalid'
inventorydir = 'inventory'
oordir = 'oor'
directories = [demanddir, invaliddir, inventorydir, oordir]


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
    demand, invalid, inventory, oor = data[0], data[1], data[2], data[3]

    invalid_items = invalid["Invalid Locations"].tolist()

    inventory["SubInv"] = inventory["SubInv"].fillna("")
    inventory = inventory[~inventory["SubInv"].str.contains(
        '|'.join(invalid_items))]

    q_dict = inventory.set_index("Item Number")["Item Qty"].to_dict()
    demand_dict = demand.set_index("Name")["Median Demand"].to_dict()

    df2 = oor.groupby("Item Code", sort=False)[
        'Quantity Due'].sum(numeric_only=False)

    df2d = df2.to_dict()
    oor["Quantity Due"] = oor["Item Code"].apply(lambda x: df2d.get(x))

    for i in range(len(oor)):
        pn, quantity = oor.iloc[i, 7], oor.iloc[i, 12]
        if pd.isna(pn) or pd.isna(quantity):
            continue
        if pn in q_dict and pn in demand_dict:
            # Subinventory quantity / Monthy Demand mean
            expected_mos = q_dict[pn]/demand_dict[pn]
            after_push = q_dict[pn] - quantity
            oor.at[i, 'Expected MOS'] = expected_mos
            oor.at[i, 'Inventory Q - Q Due'] = after_push
            oor.at[i, "Median Demand"] = demand_dict[pn]
            oor.at[i, "Valid Inventory"] = q_dict[pn]

    oor = oor.drop_duplicates(subset="Item Code")

    oor.to_excel("mosreport.xlsx", index=False)


create_report(load_data())
