import pandas as pd
import math
from fastapi import FastAPI, HTTPException, Query
from typing import List

# Initialize the FastAPI application
app = FastAPI()

# Read and prepare data from Excel
try:
    dataframe = pd.read_excel('capbudg.xls')  # Load the Excel file
    data = dataframe.to_dict(orient='index')
    raw_data = pd.DataFrame(data)  # Reconstruct DataFrame from dict
except FileNotFoundError:
    raise Exception("Excel file 'capbudg.xls' not found.")
except Exception as e:
    raise Exception(f"Error reading Excel file: {e}")

# Extracting tables from the raw_data DataFrame using specific row and column slices
try:
    Initial_investment = raw_data.iloc[[0, 2], 1:9]
    Cashflow_details = raw_data.iloc[[4, 6], 1:6]
    Discount_rate = raw_data.iloc[[8, 10], 1:9]
    Working_capital = raw_data.iloc[[0, 2], 10:14]
    Growth_rates = raw_data.iloc[0:11, [15, 17, 18, 19]]
    Initial_Investment2 = raw_data.iloc[0:1, 22:30]
    Salvage_value = raw_data.iloc[0:11, 31:34]
    year = raw_data.iloc[0:11, 20:33]
    Operating_Cashflows = raw_data.iloc[0:11, 35:50]
    Investment_measures = raw_data.iloc[1:2, 51:55]
    Book_Value = raw_data.iloc[0:11, 57:61]
except Exception as e:
    raise Exception(f"Error processing data slices from Excel file: {e}")

# Storing tables in a list for easy access
tables = [
    Initial_investment, Cashflow_details, Discount_rate, Working_capital,
    Growth_rates, Initial_Investment2, Salvage_value, Operating_Cashflows,
    Investment_measures, Book_Value
]

# Assigning readable names to tables
table_names=[]
for i in range(0,len(tables)-1):
    table_names.append(tables[i].iloc[0, 0])
table_names.append(tables[9].iloc[4, 0])



# Dictionary of available table names
available_tables = {"tables": table_names}

# Endpoint to list all available table names
@app.get("/list_tables")
def list_tables():
    return available_tables

# Endpoint to get details (row names) of a specific table
@app.get("/get_table_details")
def get_table_details(table_name: str = Query(..., description="Name of the table")):
    if table_name not in table_names:
        raise HTTPException(status_code=404, detail=f"Table '{table_name}' not found.")

    try:
        idx = table_names.index(table_name)
        df = tables[idx]
        row_names = [df.iloc[0, j] for j in range(1, len(df.columns))]
        return {"table_name": table_name, "row_names": row_names}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error retrieving table details: {str(e)}")

# Endpoint to compute sum of numeric values in a row (identified by row_name) of a specific table
@app.get("/row_sum")
def row_sum(
    table_name: str = Query(..., description="Name of the table"),
    row_name: str = Query(..., description="Name of the row")
):
    if table_name not in table_names:
        raise HTTPException(status_code=404, detail=f"Table '{table_name}' not found.")

    try:
        idx = table_names.index(table_name)
        df = tables[idx]

        # Identify column index where row_name is present in first row
        col_index = None
        for j in range(1, len(df.columns)):
            if str(df.iloc[0, j]).strip().lower() == row_name.strip().lower():
                col_index = j
                break

        if col_index is None:
            raise HTTPException(status_code=404, detail=f"Row '{row_name}' not found in table '{table_name}'.")

        # Sum all numeric values in the identified column (excluding first row)
        total = 0
        for k in range(1, len(df)):
            value = df.iloc[k, col_index]
            if isinstance(value, (int, float)) and not pd.isna(value):
                total += value

        return {"table_name": table_name, "row_name": row_name, "sum": total}

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error calculating sum: {str(e)}")
