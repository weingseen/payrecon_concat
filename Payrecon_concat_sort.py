import os
import glob
import pandas as pd
import re
from datetime import datetime
import sys

def get_base_directory():
    return os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.abspath(os.path.dirname(__file__))

def combine_xlsx_files_and_cleanup(base_directory):
    files = glob.glob(os.path.join(base_directory, '*.xlsx'))
    all_data = pd.concat([pd.read_excel(file, header=1) for file in files], ignore_index=True)
    for file in files:
        os.remove(file)

    return all_data[all_data['Shipping Information'].str.contains('others|seller|non-shopee', flags=re.I, regex=True, na=False)].sort_values(by=["Seller ID"], ascending=True)

def create_postcode_column(all_filtered):
    all_filtered.rename(columns={'Unnamed: 17': 'Address'}, inplace=True)
    all_filtered['Postcode'] = all_filtered['Address'].str[-5:]

def save_to_csv(all_filtered, base_directory):
    selected_columns = ['Seller ID', 'Order Number', 'SKU', 'Quantity', 'Customer Name', 'Phone', 'Postcode', 'Address']
    filtered_data = all_filtered[selected_columns]
    current_datetime = datetime.now().strftime('%y%m%d_%H%M')
    folder_name = 'Results'
    os.makedirs(os.path.join(base_directory, folder_name), exist_ok=True)
    csv_file_name = os.path.join(base_directory, folder_name, f'{current_datetime}.csv')
    filtered_data.to_csv(csv_file_name, index=False, header=False, encoding="utf-8")

def combine_sku_and_quantity(all_filtered):
    all_filtered['SKU'] = ''
    all_filtered['Quantity'] = ''
    all_filtered['Phone'] = all_filtered['Phone'].astype(str).str.strip().apply(lambda x: "60" + x[3:] if not x.startswith("601") and x.startswith("600") else "60" + x if not x.startswith("601") else x)

def main():
    base_directory = get_base_directory()
    all_filtered = combine_xlsx_files_and_cleanup(base_directory)
    create_postcode_column(all_filtered)
    combine_sku_and_quantity(all_filtered)
    save_to_csv(all_filtered, base_directory)

if __name__ == "__main__":
    main()
