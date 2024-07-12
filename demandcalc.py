import pandas as pd
import warnings
import glob
from collections import defaultdict
import os

class PartDesc:
    def __init__(self, name, month, year, quantity):
        self.name = name
        self.month = month
        self.year = year
        self.quantity = quantity

def get_monthlydem():
    warnings.simplefilter(action='ignore', category=UserWarning)
    reportdir = "UseReport"
    classpath = glob.glob(os.path.join(reportdir, "*.xlsx"))
    if not classpath:
        raise FileNotFoundError("Classification document not found")
    
    df = pd.read_excel(classpath[0], sheet_name="Sheet1")

    item_table = []
    
    for i in range(len(df)):
        date = df.iloc[i,3].to_pydatetime()  # Assuming the date column is at index 3
        month = date.month
        year = date.year
        pn = str(df.iloc[i,0])  # Assuming the part name is at index 0
        quantity = df.iloc[i,2]  # Assuming the quantity is at index 2
        item_table.append(PartDesc(pn, month, year, quantity))
    
    data = defaultdict(lambda: defaultdict(int))

    for obj in item_table:
        key = (obj.name, obj.month, obj.year)
        data[key]['Total_Quantity'] += obj.quantity

    flattened_data = [{'Name': k[0], 'Month': k[1], 'Year': k[2], 'Median Demand': v['Total_Quantity']} for k, v in data.items()]

    df = pd.DataFrame(flattened_data)

    # Calculate the mean demand for each item across all months and years
    mean_demand_df = df.groupby('Name')['Median Demand'].median().reset_index()

    # Save the mean demand DataFrame to an Excel file
    mean_demand_df.to_excel('expected_demand.xlsx', index=False)
    
get_monthlydem()
