import os
import glob
import pandas as pd
import re
from datetime import datetime
import sys

# Determine the base directory path, considering whether the script is run as an executable
def get_base_directory():
    if getattr(sys, 'frozen', False):
        # Running as executable (frozen)
        return os.path.dirname(sys.executable)
    else:
        # Running as script
        return os.path.abspath(os.path.dirname(__file__))

def combine_xlsx_files_and_cleanup(base_directory):
    # Find the XLSX files
    files = glob.glob(os.path.join(base_directory, '*.xlsx'))

    # Create the dataframe by concatenating XLSX files
    all_data = pd.concat([pd.read_excel(file, header=1) for file in files])

    # Remove the original XLSX files
    for file in files:
        os.remove(file)

    # Perform data cleanup and sorting
    all_filtered = all_data.loc[all_data['Shipping Information'].str.contains('others|seller|non-shopee', flags=re.I, regex=True, na=False)].sort_values(by=["Seller ID"], ascending=True)

    return all_filtered


def create_postcode_column(all_filtered):
    # Rename the "Unnamed: 17" column to "Address"
    all_filtered.rename(columns={'Unnamed: 17': 'Address'}, inplace=True)

    # Create a new column "Postcode" based on the last 5 characters of "Address"
    all_filtered['Postcode'] = all_filtered['Address'].str[-5:]



def save_to_csv(all_filtered, base_directory):
    # Define the columns you want to keep in the desired sequence
    selected_columns = [
        'Seller ID', 'Order Number', 'SKU', 'Quantity', 'Customer Name', 'Phone', 'Postcode', 'Address'
    ]

    # Select only the desired columns and reorder them
    filtered_data = all_filtered[selected_columns]

    # Get the current date and time in the desired format (YYMMDD_HourMinute)
    current_datetime = datetime.now().strftime('%y%m%d_%H%M')

    # Specify the folder name (Results) and the CSV file name with the current date and time
    folder_name = 'Results'
    os.makedirs(os.path.join(base_directory, folder_name), exist_ok=True)
    csv_file_name = os.path.join(base_directory, folder_name, f'{current_datetime}.csv')

    # Save the filtered data to the CSV file in the "Results" folder
    filtered_data.to_csv(csv_file_name, index=False,header=False, encoding="utf-8")

def combine_sku_and_quantity(all_filtered):
    sku_values = all_filtered['SKU'].split(',')
    quantity_values = all_filtered['Quantity'].split(',')

    # Create an empty dictionary to store the SKU and Quantity pairs
    sku_quantity_dict = {}

    # Iterate through the indices and combine SKU and Quantity
    for index in range(len(sku_values)):
        sku = sku_values[index].strip()  # Remove leading/trailing spaces
        quantity = int(quantity_values[index].strip())  # Convert quantity to an integer
        sku_quantity_dict[sku] = quantity  

    all_filtered["SKU"] = ",".join([f"{key}-{value}" for key, value in sku_quantity_dict.items()])
    all_filtered["Quantity"] = ",".join(map(str, sku_quantity_dict.values()))


def main():
    base_directory = get_base_directory()
    all_filtered = combine_xlsx_files_and_cleanup(base_directory)
    create_postcode_column(all_filtered)
    combine_sku_and_quantity(all_filtered)  # Call the function to combine SKU and Quantity
    save_to_csv(all_filtered, base_directory)

if __name__ == "__main__":
    main()