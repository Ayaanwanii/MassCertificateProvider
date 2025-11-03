import pandas as pd
import sqlite3

# Connect to SQLite database
conn = sqlite3.connect('certificates.db')

# Get all table names from the database
cursor = conn.cursor()
cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
tables = cursor.fetchall()

# Create an Excel writer object
with pd.ExcelWriter('certificates_export.xlsx', engine='openpyxl') as writer:
    # Export each table to a separate sheet in the Excel file
    for table in tables:
        table_name = table[0]
        # Read the SQL table into a pandas DataFrame
        df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
        # Write the DataFrame to an Excel sheet
        df.to_excel(writer, sheet_name=table_name, index=False)

# Close the database connection
conn.close()

print("Data has been exported to 'certificates_export.xlsx'")