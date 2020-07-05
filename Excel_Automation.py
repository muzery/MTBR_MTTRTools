from openpyxl import load_workbook
import re
import os
from openpyxl import worksheet
import matplotlib.pyplot as plt
from datetime import datetime
import Excel_Extraction


def excel_extraction_function(dictonary_file_path1):
    return Excel_Extraction.main(dictonary_file_path1)


def extract_dictionary_value(dictionary_station_name, key_value):
    if isinstance(dictionary_station_name, dict):
        return dictionary_station_name[key_value]


def extract_dictionary_key(dictionary_station_name, BE_to_check):
    if isinstance(dictionary_station_name, dict):
        key_values = dictionary_station_name.keys()
        for key_value in key_values:
            if BE_to_check == key_value:
                return extract_dictionary_value
            else:
                print("Nothing Match. Please check template file")
                return None
    else:
        print("Nothing Match. Please check template file")


def excel_automate(compile_file, list_stations, start_date, end_date, title_cell, worksheet, BU, ):
    # template :
    # Title (C13) + 0 title cell
    # Next Station Part number C (col) + 1
    # Start Date (B14) + 2
    # End Date:  (B15) + 3
    # # of station B(17) + 5
    # count of MI B(19) + 7
    # Total of # of days B(20) +8
    Found = False
    write = False
    Number_Of_Stations = 0
    store_all = []
    two_D_store = []
    # this is title
    items = re.split(r"([A-Z]{1,2})(\d+)", title_cell)  # <-----------problem happen at here
    ASCII_code = ord(items[1])
    counter = 1
    with open(compile_file, 'r') as fp:
        # it is read line by line
        chunk_line = fp.readline()  # jump one line due to the title
        while True:
            chunk_line = fp.readline()
            # split the link by using tab
            each_element = chunk_line.split('\t')
            if not chunk_line:
                break
            if each_element[0] == BU:
                # get the element out and store into array
                store_all.append(each_element[1])  # store part number
                store_all.append(each_element[2])  # store Total_Days
                store_all.append(each_element[3])  # store Number_Of_Stations
                store_all.append(each_element[4])  # store Number_Of_MI_Occurance
                two_D_store.append(store_all)  # store all
                store_all = []

            else:
                pass
        # if write:
        #     two_D_store.append(store_all)
        #     write =False
    # Checking all to get location key value
    dictionary_key_value = {}
    for stations in list_stations:
        for stn in stations:
            # loop thru
            for i in range(1, 10, 1):
                ASCII_code = ASCII_code + 1
                character = chr(ASCII_code)
                row_col = character + str(int(items[2]) + 1)
                # get value
                cell_value = worksheet[row_col]
                if cell_value.value == stn:
                    dictionary_key_value[stn] = row_col
                    # reset the ascii code
                    ASCII_code = ord(items[1])
                    break
    print(dictionary_key_value)
    # get all keys
    all_location = dictionary_key_value.keys()

    for location in all_location:
        # get value
        home_position = dictionary_key_value[location]
        # doing checking on the spreadsheet
        station_part_number_in_cell = worksheet[home_position]
        if location == station_part_number_in_cell.value:
            items = re.split(r"([A-Z]{1,2})(\d+)", home_position)
            ASCII_code = ord(items[1])
            ASCII_code = ASCII_code + 0
            character = chr(ASCII_code)

            # Start Date
            rows = int(items[2]) + 1
            # Appending both row and column
            row_col = character + str(rows)
            worksheet[row_col] = start_date
            # End Date
            rows = int(items[2]) + 2
            # Appending both row and column
            row_col = character + str(rows)
            worksheet[row_col] = end_date
            # of station
            # get station name
            rows = int(items[2]) + 0
            row_col = character + str(rows)  # need to change to dynamic
            # Get the element from the cell value for excel
            station_part_number_in_cell = worksheet[row_col]
            # Can apply re if needed
            rows = int(items[2]) + 4
            for station in stations:

                if station == station_part_number_in_cell.value:
                    print(stations[station])
                    print(len(stations[station]))
                    Number_Of_Stations = len(stations[station])
                    Found = True
                if Found:
                    row_col = character + str(rows)
                    worksheet[row_col].value = Number_Of_Stations
                    Number_Of_Stations = 0
                    Found = False
                # look thru templateFile
            for one_D_store in two_D_store:
                if station_part_number_in_cell.value == one_D_store[0]:
                    # count of MI
                    rows = int(items[2]) + 6
                    row_col = character + str(rows)
                    # for templateIO
                    worksheet[row_col].value = int(one_D_store[3])
                    # Total of # of days
                    rows = int(items[2]) + 7
                    row_col = character + str(rows)
                    worksheet[row_col].value = int(one_D_store[1])
                    # for templateIO
    return worksheet


def save_as_file(template_file_io_to_write, workbook, worksheet, start_month, end_month):
    # month_year[1]  - month
    # month_year[2]   - year
    print(template_file_io_to_write)
    template_file_io_to_write = template_file_io_to_write.replace('Month', start_month[1] + '-' + end_month[1])
    template_file_io_to_write = template_file_io_to_write.replace('YYYY', start_month[2])
    template_file_io_to_write = template_file_io_to_write.replace('Calculation_template', 'Calculation')
    workbook.save(template_file_io_to_write)
    workbook.close()


def open_template_file(template_file_io_to_write, dictionaryFile, compile_file, start_date, end_date):
    # 1. open the template file.
    # 2. looking for the B12 cell , it should be having string if no let user choose
    # 3. anything inside 10 row
    # 4. looping for another 10 row or back to 1
    # excel_extraction_function(dictionaryFile)
    starting_point = []
    # loading dictionary
    dictionary_station_name = excel_extraction_function(dictionaryFile)
    workbook = load_workbook(filename=template_file_io_to_write)
    worksheet = workbook["Sheet1"]
    print(worksheet)
    # looking for title cell # Assume all title in B column
    all_station_name = dictionary_station_name.keys()
    # looking for the station_name
    # run thru all B column until continue next row is empty
    for station in all_station_name:
        print(station)
        # run thru the station list and compare first column
        for row in range(12, 300, 1):
            row_col = "B" + str(row)
            cell_value = worksheet[row_col]
            if cell_value.value == station:
                # Save inside the position array
                starting_point.append(row_col)
                break
    # ------------------New Product run at here--------------------#
    if len(starting_point) != 0:
        for home_position in starting_point:
            BU = worksheet[home_position]
            print(BU.value)
            if not BU.value:  # string empty
                # user input
                pass
            else:
                # format is pretty constant
                pass
            # get from dictionary
            list_station_from_dictionary = extract_dictionary_value(dictionary_station_name, BU.value)
            if len(list_station_from_dictionary) == None:
                print("No list file, it will not continue ")
                return
            else:
                # start_Date = input('Enter Start Date:')
                # end_Date = input('Enter End Date')
                start_Date = start_date  # Will change to user input
                end_Date = end_date
                month_year_start = re.split(r'\d+\-(\w+)\-([0-2]+)', start_Date)
                month_year_end = re.split(r'\d+\-(\w+)\-([0-2]+)', end_Date)
                worksheet = excel_automate(compile_file, list_station_from_dictionary, start_Date, end_Date,
                                           home_position, worksheet, BU.value)

        # month_year[1]  - month
        # month_year[2]   - year
        save_as_file(template_file_io_to_write, workbook, worksheet, month_year_start, month_year_end)

    else:
        print("Nothing Found!")
    # format is pretty constant
    # starting worksheet


# 1. Enter Starting date and end date:
# 2. Get The number of station from dictionary and  put into the template
# B is title next 10 rows column is unknown

if __name__ == "__main__":
    template_file_io_to_write = r"C:\Users\willlee\Desktop\Manufacturing Issue\2020\MI_MTBF_MTTR\MTBF_MTTR Month YYYY Calculation_template.xlsx"
    dictionaryFile = r"C:\Users\willlee\Desktop\Manufacturing Issue\2020\Test Station Part Number.xlsx"
    compile_file = r"C:\Users\willlee\.PyCharmCE2018.2\config\scratches\file_use_in_template.txt"  # save same source file
    open_template_file(template_file_io_to_write, dictionaryFile, compile_file, '1-MAY-2020', '31-MAY-2020')
