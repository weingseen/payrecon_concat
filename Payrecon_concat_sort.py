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
    # Create empty lists to store the combined SKU and Quantity strings
    combined_sku_values = []
    combined_quantity_values = []
    combined_phone_number_values = []

    # Iterate through each row in the DataFrame
    for index, row in all_filtered.iterrows():
        sku_values = row['SKU'].split(',')
        quantity_values = str(row['Quantity']).split(',')

        # Create an empty dictionary to store the SKU and Quantity pairs for the current row
        sku_quantity_dict = {}

        # Iterate through the indices and combine SKU and Quantity for the current row
        for i in range(len(sku_values)):
            sku = sku_values[i].strip()  # Remove leading/trailing spaces
            quantity = int(quantity_values[i].strip())  # Convert quantity to an integer
            sku_quantity_dict[sku] = sku_quantity_dict.get(sku, 0) + quantity

        # Join SKU and Quantity for the current row and append to the lists
        combined_sku = ",".join([f"{key}" for key, value in sku_quantity_dict.items()])
        combined_quantity = ",".join(map(str, sku_quantity_dict.values()))
        combined_sku_values.append(combined_sku)
        combined_quantity_values.append(combined_quantity)

        # PHONE NUMBER LOGIC HERE
        original_phone_number = str(row['Phone']).strip()

        # Exclude phone numbers starting with "601"
        if not original_phone_number.startswith("601"):
            first_three_digits = original_phone_number[:3]
            after_three_digits = original_phone_number[3:]
            if first_three_digits == "600":
                combined_phone_number_values.append("60" + after_three_digits)
            else:
                combined_phone_number_values.append("60" + original_phone_number)
        else:
            combined_phone_number_values.append(original_phone_number)

    # Update the DataFrame with the combined SKU and Quantity values
    all_filtered['SKU'] = combined_sku_values
    all_filtered['Quantity'] = combined_quantity_values
    all_filtered['Phone'] = combined_phone_number_values

def main():
    base_directory = get_base_directory()
    all_filtered = combine_xlsx_files_and_cleanup(base_directory)
    create_postcode_column(all_filtered)
    combine_sku_and_quantity(all_filtered)  # Call the function to combine SKU and Quantity
    save_to_csv(all_filtered, base_directory)

if __name__ == "__main__":
    main()