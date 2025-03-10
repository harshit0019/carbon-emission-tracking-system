import json
import random
from datetime import datetime

# Define the units and months.
units = ["C-49", "B-37", "C-91", "2B-4"]
months = [
    ("January", "01"),
    ("February", "02"),
    ("March", "03"),
    ("April", "04"),
    ("May", "05"),
    ("June", "06"),
    ("July", "07"),
    ("August", "08"),
    ("September", "09"),
    ("October", "10"),
    ("November", "11"),
    ("December", "12")
]

# Define Scope 1 - Fuel types with ranges for amounts.
fuel_types = [
    {"name": "Diesel", "factor": 2.54603, "amount_range": (80, 120), "doc_ext": ".pdf"},
    {"name": "Petrol", "factor": 2.296,   "amount_range": (80, 120), "doc_ext": ".pdf"},
    {"name": "PNG",    "factor": 2.02266, "amount_range": (100, 150), "doc_ext": ".pdf"},
    {"name": "LPG",    "factor": 1.55537, "amount_range": (120, 180), "doc_ext": ".pdf"}
]

# Define Scope 2 - Electricity with a range.
electricity = {"name": "Electricity", "factor": 0.6727, "amount_range": (40, 80), "doc_ext": ".jpeg"}

# Function to compute total given factor and amount.
def compute_total(factor, amount):
    total = factor * amount
    return f"{total:.2f}"

records = []
record_id = 0
email = "manager@gmail.com"
years = ["2025", "2026"]

# Generate records for each year, unit, and month.
for year in years:
    for unit in units:
        for month_name, month_num in months:
            entry_date = f"{year}-{month_num}-07"  # Fixed day: 7th.
            
            # Create Fuel records with randomized amounts.
            for fuel in fuel_types:
                amount = random.randint(*fuel["amount_range"])
                total = compute_total(fuel["factor"], amount)
                doc = f"CarbonData\\{unit}\\{year}\\{month_num}_{month_name}\\{unit}_07_{month_num}_{year}_{fuel['name']}_Fuel{fuel['doc_ext']}"
                record = [email, entry_date, month_name, unit, "Fuel", fuel["name"], f"{fuel['factor']}", str(amount), total, doc, record_id]
                record_id += 1
                records.append(record)
            
            # Create an Electricity record with randomized amount.
            amount = random.randint(*electricity["amount_range"])
            total = compute_total(electricity["factor"], amount)
            doc = f"CarbonData\\{unit}\\{year}\\{month_num}_{month_name}\\{unit}_07_{month_num}_{year}_{electricity['name']}_Electricity{electricity['doc_ext']}"
            record = [email, entry_date, month_name, unit, "Electricity", electricity["name"], f"{electricity['factor']}", str(amount), total, doc, record_id]
            record_id += 1
            records.append(record)

# Save records to a JSON file.
with open("emission_records.json", "w") as f:
    json.dump(records, f, indent=4)

print("Fake data generated and saved to 'emission_records.json'.")
