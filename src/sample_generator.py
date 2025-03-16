import pandas as pd
import numpy as np
from io import BytesIO
import traceback
import streamlit as st

def create_sample_files():
    """Create sample Excel files with known differences for testing"""
    try:
        # Create the first sample file (base file)
        output1 = BytesIO()

        # Use a context manager to ensure proper resource handling
        with pd.ExcelWriter(output1, engine='openpyxl') as writer1:
            # Customers sheet
            customers1 = pd.DataFrame({
                "CustomerID": list(range(1, 11)),
                "Name": [f"Customer {i}" for i in range(1, 11)],
                "Email": [f"customer{i}@example.com" for i in range(1, 11)],
                "Phone": [f"555-{i:04d}" for i in range(1, 11)],
                "Address": [f"Address {i}" for i in range(1, 11)]
            })
            customers1.to_excel(writer1, sheet_name="Customers", index=False)

            # Products sheet
            products1 = pd.DataFrame({
                "ProductID": list(range(1, 21)),
                "Name": [f"Product {i}" for i in range(1, 21)],
                "Category": ["Category A"] * 7 + ["Category B"] * 7 + ["Category C"] * 6,
                "Price": [10.99, 20.99, 15.50, 25.00, 30.99, 12.50, 22.75,
                         18.99, 24.50, 35.00, 9.99, 19.99, 29.99, 14.50,
                         22.50, 17.99, 27.50, 32.99, 11.50, 21.75],
                "Stock": [100, 50, 75, 25, 60, 90, 40, 80, 30, 70,
                         110, 55, 65, 35, 85, 45, 95, 15, 105, 20]
            })
            products1.to_excel(writer1, sheet_name="Products", index=False)

            # Employees sheet
            employees1 = pd.DataFrame({
                "EmployeeID": list(range(1, 6)),
                "Name": ["John Doe", "Jane Smith", "Bob Johnson", "Alice Brown", "Charlie Davis"],
                "Department": ["Sales", "Marketing", "IT", "HR", "Finance"],
                "Salary": [50000, 60000, 70000, 55000, 65000],
                "HireDate": ["2020-01-15", "2019-05-20", "2021-03-10", "2018-11-05", "2020-07-22"]
            })
            employees1.to_excel(writer1, sheet_name="Employees", index=False)

        # Create the second sample file (comparison file with differences)
        output2 = BytesIO()

        with pd.ExcelWriter(output2, engine='openpyxl') as writer2:
            # Create a deep copy to avoid modifying the original data
            customers2 = customers1.copy(deep=True)

            # Change a few values
            customers2.loc[2, "Email"] = "changed.email@example.com"  # Change email for customer 3
            customers2.loc[5, "Phone"] = "555-9999"  # Change phone for customer 6

            # Add a new customer
            new_customer = pd.DataFrame({
                "CustomerID": [11],
                "Name": ["New Customer"],
                "Email": ["new@example.com"],
                "Phone": ["555-1111"],
                "Address": ["New Address"]
            })
            customers2 = pd.concat([customers2, new_customer], ignore_index=True)

            # Remove a customer
            customers2 = customers2[customers2["CustomerID"] != 4]  # Remove customer 4
            customers2.to_excel(writer2, sheet_name="Customers", index=False)

            # Products sheet with differences
            products2 = products1.copy(deep=True)

            # Remove some rows
            products2 = products2[products2["ProductID"] < 18]  # Remove last 3 products

            # Change column order
            products2 = products2[["ProductID", "Category", "Name", "Stock", "Price"]]  # Reordered columns

            # Change some values - check indices first
            if 5 < len(products2):
                products2.loc[5, "Price"] = 13.99  # Change price for product 6
            if 10 < len(products2):
                products2.loc[10, "Stock"] = 200  # Change stock for product 11

            products2.to_excel(writer2, sheet_name="Products", index=False)

            # No Employees sheet in the second file (missing sheet)

            # Add a new sheet that doesn't exist in the first file
            orders = pd.DataFrame({
                "OrderID": list(range(1, 6)),
                "CustomerID": [3, 5, 2, 7, 9],
                "OrderDate": ["2023-01-10", "2023-02-15", "2023-03-20", "2023-04-25", "2023-05-30"],
                "TotalAmount": [125.50, 230.75, 75.99, 310.25, 150.00]
            })
            orders.to_excel(writer2, sheet_name="Orders", index=False)

        # Return the files as bytes
        output1.seek(0)
        output2.seek(0)
        return output1.getvalue(), output2.getvalue()

    except Exception as e:
        # More detailed error reporting
        error_msg = f"Error generating sample files: {str(e)}\n{traceback.format_exc()}"
        print(error_msg)
        st.error(error_msg)
        # Return empty files to avoid crashing
        empty1 = BytesIO()
        empty2 = BytesIO()
        with pd.ExcelWriter(empty1, engine='openpyxl') as writer:
            pd.DataFrame({"Error": ["Sample generation failed"]}).to_excel(writer, index=False)
        with pd.ExcelWriter(empty2, engine='openpyxl') as writer:
            pd.DataFrame({"Error": ["Sample generation failed"]}).to_excel(writer, index=False)
        empty1.seek(0)
        empty2.seek(0)
        return empty1.getvalue(), empty2.getvalue()