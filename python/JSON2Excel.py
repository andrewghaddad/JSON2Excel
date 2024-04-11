import json
import datetime
from openpyxl import Workbook

# Configuration parameters
input_file_name = 'input.json'  # Name of the input JSON file
date = datetime.datetime.now() # Timestamp of output Excel file
output_file_path = './python/files/' # File path where output Excel file will be written to
output_file_name = output_file_path + 'output ' + date.strftime("%Y-%m-%d") + '.xlsx'  # Name of the output Excel file
sheet_name = 'News For ' +  date.strftime("%Y-%m-%d");  # Name of the sheet in the Excel file

# Function to read JSON data from file
def read_json(file_name):
    with open(file_name, 'r') as f:
        return json.load(f)

# Function to write JSON data to Excel
def write_excel(data):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet_name
    data = data['articles']

    # Write headers
    headers = list(data[0].keys())
    worksheet.append(headers)

    # Write data
    for item in data:
        row = [str(item[header]) for header in headers]
        worksheet.append(row)

    # Save the workbook
    workbook.save(output_file_name)
    print(f'Excel file "{output_file_name}" generated successfully.')

# Main function
def main():
    try:
        # Read JSON data from file
        json_data = read_json(input_file_name)
        # Write JSON data to Excel
        write_excel(json_data)
    except Exception as e:
        print('Error:', e)

# Execute the main function
# if __name__ == "__main__":
main()
