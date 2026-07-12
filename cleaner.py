import os
import pandas as pd

# Ask the user for the Excel file name
file_name = input("Enter Excel file name (example: data.xlsx): ")

try:
    # Create output folder if it doesn't exist
    os.makedirs("output", exist_ok=True)

    # Read the Excel file
    df = pd.read_excel(file_name, engine="openpyxl")

    print("\nOriginal data:")
    print(df.head())

    # Count rows before cleaning
    rows_before = len(df)

    # Remove duplicate rows
    df = df.drop_duplicates()

    # Remove empty rows
    df = df.dropna(how="all")

    # Remove extra spaces from text columns
    df = df.apply(
        lambda col: col.str.strip() if col.dtype == "object" else col
    )

    # Count rows after cleaning
    rows_after = len(df)
    duplicates_removed = rows_before - rows_after

    # Save cleaned data inside the output folder
    output_file = "output/cleaned_data.xlsx"
    df.to_excel(output_file, index=False)

    print("\nCleaning completed successfully!")
    print(f"Cleaned file saved as: {output_file}")

    print("\nCleaning Report")
    print("-" * 30)
    print(f"Rows before cleaning : {rows_before}")
    print(f"Rows after cleaning  : {rows_after}")
    print(f"Duplicates removed   : {duplicates_removed}")

except Exception as e:
    print("Error:", e)