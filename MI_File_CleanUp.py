import pandas as pd
import os
import Data_calculation
import Excel_Extraction
import numpy as np

full_path = os.path.abspath(os.getcwd())
dictonary_file_path = file_name = r"C:\Users\willlee\Desktop\Manufacturing Issue\2020\Test Station Part Number.xlsx"
file_filter_to_save = full_path + "\\file_use_in_template.txt"  # this is temporary file code generate itself


def create_new_file(file_save, file_name_and_ext_change):
    filepath, file_name = os.path.split(file_save)
    return filepath + "\\" + file_name_and_ext_change


def excel_extraction_function(dictonary_file_path1):
    return Excel_Extraction.main(dictonary_file_path1)


# loading data_analysis
def data_analysis():
    dictionary_cal = Data_calculation.rules_operation()
    return dictionary_cal


def data_frame_merging_filter(usr, input_file, output_file, dictonary_file_path):
    # usr = input('Please select you want to do the correction only (1) or  grouping after doing correction(2) or using predict analysis(3):')
    # Open MI file
    # 1. doing clean up on data frame
    # 2. change to columns (more flexible) -- should need to filter dataframe as some are not used

    if usr == 1:
        df = pd.read_table(input_file)

    elif usr == 2:
        df = pd.read_excel(output_file)
        # read for the column that Unnamed and drop it
        for column in df.columns:
            if "Unnamed" in column or "unnamed" in column:  # will not use the column name using unnamed
                df = df.drop(columns=column, axis=1)

    # check df and see already update
    print(df)
    # looking for some important needed columns
    x = df.columns == 'Days Open'
    y = np.where(x == True)
    # Check content for the array
    # get the length -- Expected one only else whole sequence out

    if len(y) > 1 or len(y) == 0:
        return None, 3

    # Upper case for the test station
    df['Test Stations Used'] = df['Test Stations Used'].str.upper()
    df['Test Stations Used'] = df['Test Stations Used'].str.replace(" ", "")
    # Open the dictonary file and get the sheet
    xl = pd.ExcelFile(dictonary_file_path)
    print(xl.sheet_names)
    xl.close()
    # 2. Open Excel File Read with the Correct Sheet
    df_dictonary = pd.read_excel(io=dictonary_file_path, sheet_name=xl.sheet_names[2])
    # Create empty data frame column
    # https://stackoverflow.com/questions/39050539/how-to-add-multiple-columns-to-pandas-dataframe-in-one-assignment
    # applicable for user choose number 1 and 3
    if usr == 1:
        df = df.join(pd.DataFrame(
            {
                'Station Name': "Undefined",
                'Location': 'Undefined',
                'TE PN': 'Undefined',
                'Nearest Correction Value': '-',
                'Original Value': '-'
            }, index=df.index
        ))
    print(df)
    # reserve for machine learning use
    # prediction_dict = data_analysis()
    cannot_find = False

    for row in range(0, len(df)):
        valueOut = df['Test Stations Used'][row]
        df_filter = df_dictonary.loc[df_dictonary['Station Name'] == valueOut]
        try:

            df.loc[row, 'Station Name'] = df_filter['Station Name'].values[0]
            df.loc[row, 'Location'] = df_filter['Location'].values[0]
            df.loc[row, 'TE PN'] = df_filter['TE PN'].values[0]

        except:
            cannot_find = True
            # pull to other file need to add  auto checking
            print(row)
            nearest_station_name = Excel_Extraction.correction_station_name(valueOut, dictonary_file_path)
            # assign station name
            str_text = ""
            for station_naming in nearest_station_name:
                str_text = str_text + station_naming + ";"
            # write into the row
            if usr == 1:
                df.loc[row, 'Nearest Correction Value'] = str_text
                df.loc[row, 'Original Value'] = valueOut
                # df = predict_analysis(prediction_dict,df,row,valueOut )

            cannot_find = False
            pass
            print(df)
            if cannot_find:
                print("Some Part number cannot find from dictionary. Please Manual Check it ")
                print(row)

    df.to_excel(output_file)
    # df.to_excel('C:\\Users\\willlee\\Desktop\\DataSet\\May_Data_2020_2.xlsx')
    # put color at here if needed
    return [df, usr]


def predict_analysis(dictonary_data, df, row, value_out):
    values = {}
    print("This prediction based on the highest probability value and it might not correct")
    print("Press Y to continue")
    # get all keys value and compare key with the value out
    all_keys = dictonary_data.keys()
    for key in list(all_keys):
        ori_value, corr_value = key.split('_')
        # doing comparison
        if ori_value == value_out:
            # check the probability
            probability_value = dictonary_data[ori_value + "_" + corr_value]
            values[ori_value + "_" + corr_value] = probability_value
    print(value_out)
    print(row)
    print(values)
    # check whether the dataset having the correct value
    if len(values) == 0:
        pass
    else:
        # after filtering and get the maximun value for key
        max_keys = max(values)
        # split_max_values
        ori_value, corr_value = max_keys.split('_')
        # replace the corr_value to the test station
        df.loc[row, 'Test Stations Used'] = corr_value
    # #get_value_list
    # val_list = list(values)
    # index_value = val_list.index(max_value)
    # list(values.keys())[index_value]
    return df


def grouping(data_frame, table):
    # doing groupby
    file_store = {}
    print(table)
    number_of_stations = 0
    number_of_MI = 0
    total_days = 0
    df_grouping = data_frame.groupby(['Location', 'TE PN'])
    print(df_grouping)
    for table_key in df_grouping.groups.keys():
        # reset number_of_stations
        number_of_stations = 0
        # look for number station
        part_number = table_key[1]
        for each_list in table:
            try:
                number_of_stations = each_list[part_number]
            except:
                print("Cannot look for it")
            else:
                break
        values = df_grouping.groups[table_key]
        for value in values:
            days_open = data_frame.loc[value, 'Days Open']
            total_days = total_days + days_open
            number_of_MI = number_of_MI + 1
        file_store[table_key] = [total_days, number_of_stations, number_of_MI]
        total_days = 0
        number_of_MI = 0
    return file_store


def save_file(file_store, file_name):
    title = "Location" + "\t" + "TE PN" + "\t" + "Total_Days" + "\t" + "Number_Of_Stations" + "\t" + "Number_Of_MI_Occurance" + "\n"
    if isinstance(file_store, dict):
        # create an empty file
        fp = open(file_name, "w")
        fp.write(title)
        fp.close()
        for each_element in file_store:
            print(each_element)
            [Total_Days, Number_Of_Stations, Number_Of_MI_Occurance] = file_store[each_element]
            with open(file_name, 'a+') as fp:
                total_string = str(each_element[0]) + "\t" + str(each_element[1]) + "\t" + str(Total_Days) + "\t" + str(
                    Number_Of_Stations) + "\t" + str(Number_Of_MI_Occurance)
                fp.write(total_string + "\n")

    else:
        return


# https://stackoverflow.com/questions/41428539/data-frame-to-file-txt-python
def file_rename(file_to_open, file_write_in):
    # column_to_swap = list
    title = ''
    data = pd.read_excel(file_to_open)
    # get the column name as title
    for column in data.columns:
        title = title + column + "\t"

    np.savetxt(file_write_in, data.values, delimiter="\t", header=title, fmt="%s")


def main(dictonary_file_path=None, usr1=1, input_file=None, save_file_path=None):
    # read for the file extension and change to the txt file format
    file_path, file_name = os.path.splitext(input_file)
    print(file_name)
    if '.xlsx' in file_name:
        # save the file in the .txt
        file_path = file_path + ".txt"
        # create path
        try:
            file_rename(input_file, file_path)
            input_file = file_path
        except PermissionError as err:
            print("Error {0}".format(err))
            return

    try:
        input_file = open(input_file, "r+")  # or "a+", whatever you need


    except IOError:
        print("Could not open file! Please close Excel manually!")
        # close file
        return
    else:
        pass

    table_list = []
    dict_part_number = {}
    table_station = excel_extraction_function(dictonary_file_path)
    print(table_station)
    for table_key in table_station.keys():
        values = table_station[table_key]
        for value in values:
            print(len(value))
            if isinstance(value, dict):
                for val in value.keys():
                    print(len(value[val]))
                    dict_part_number[val] = len(value[val])
                    table_list.append(dict_part_number)
                    dict_part_number = {}
    print(table_list)

    [df, usr] = data_frame_merging_filter(usr1, input_file, save_file_path, dictonary_file_path)

    if usr == 2:
        file_store = grouping(data_frame=df, table=table_list)
        print(file_store)
        file_filter_to_save = create_new_file(save_file_path, 'templatefile.txt')
        save_file(file_store, file_filter_to_save)
    if usr == 3:
        # raise error
        raise Exception("Cannot look for important columns. Please check your file.")
    # Save into the file


if __name__ == "__main__":
    main(dictonary_file_path=r"C:\Users\willlee\Desktop\Manufacturing Issue\2020\Test Station Part Number.xlsx"
         , usr1=1, input_file=r"C:\Users\willlee\Desktop\Manufacturing Issue\2020\MI_MTBF_MTTR\Other\MI_Jan-Feb.xlsx",
         save_file_path=r"C:\Users\willlee\Desktop\DataSet\1H_2020_3.xlsx")
