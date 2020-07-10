import pandas as pd
import os
import Data_calculation
import Excel_Extraction

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


def data_frame_merging_filter(usr, input_file, output_file):
    # usr = input('Please select you want to do the correction only (1) or  grouping after doing correction(2) or using predict analysis(3):')
    # Open MI file
    if usr == 1:
        # ----------------select which one you want to run---------------------------
        df = pd.read_csv(input_file, names=["Issue Number", "Planner Code", "Product Description", "Managment Stripe",
                                            "P7_TEST_ITA", \
                                            "P7_TEST_FIXTURE", "P7_TEST_KIT", "P7_POT_SERIAL_NUMBER",
                                            "P7_ROOT_CAUSE_TYPE_IDP7_ROOT_CAUSE_TYPE_ID", \
                                            "P7_ROOT_CAUSE_SUBTYPE_ID", "P7_ROOT_CAUSE_DESCRIPTION", "Severity",
                                            "Status", "Type", "Category", "Assignee Name", \
                                            "Reviewer Name", "Days Open", "Test Station(s)", "Date Reported",
                                            "Last Update Date", \
                                            "Orangization Code", "Problem Summary","complete?","Days Open_2"], sep="\t", skiprows=1,
                         encoding="utf-8", )

        # df = pd.read_csv('C:\\Users\\willlee\\Desktop\\DataSet\\May_Data_2020_2.csv',names =["Issue Number","Planner Code","Product Description", \
        #              "Managment Stripe","P7_TEST_ITA","P7_TEST_FIXTURE","P7_TEST_KIT","P7_POT_SERIAL_NUMBER"\
        #              "P7_ROOT_CAUSE_TYPE_ID", "P7_ROOT_CAUSE_SUBTYPE_ID","P7_ROOT_CAUSE_DESCRIPTION","Severity","Status","Type",	\
        #              "Category","Assignee Name","Reviewer Name","Days Open","Test Station(s)","Date Reported",\
        #              "Last Update Date"	,"Orangization Code","Problem Summary"],sep="\t", skiprows=1,encoding = "utf-8",)
        df = df
    elif usr == 2:
        df = pd.read_excel(io=output_file, sheet_name='Sheet1')
        # df = pd.read_excel(io=r"C:\\Users\\willlee\\Desktop\\DataSet\\May_Data_2020_2.xlsx", sheet_name='Sheet1')

    print(len(df))
    #    df = df.reset_index()

    # df.rename(columns={'index': 'Issue Number', 'Issue Number': 'Planner Code', 'Planner Code': 'Product Description',
    #                     'Product Description': 'Managment Stripe', 'Managment Stripe': 'P7_TEST_ITA', \
    #                     'P7_TEST_ITA': 'P7_TEST_FIXTURE', 'P7_TEST_FIXTURE': 'P7_TEST_KIT',
    #                    'P7_TEST_KIT': 'P7_POT_SERIAL_NUMBER',
    #                     'P7_POT_SERIAL_NUMBERP7_ROOT_CAUSE_TYPE_ID': 'P7_ROOT_CAUSE_TYPE_ID', \
    #                    }, inplace=True)
    #    df = df.reset_index()
    #    df.drop(['Orangization Code'],inplace= True,axis = 1) # for csv file not for ABCD-1
    # no longer using
    # df.rename(columns={'level_0': 'Issue Number','level_1':'Severity','Issue_Number':'Status','Severity':'Type',
    #                'Status':'Category','Type':'Assignee Name','Category':'Reviewer Name','Assignee Name':'Days Open',
    #                'Reviewer Name':'Test Station(s)','Days Open':'Date Reported','Test Station(s)':'Last Update Date',
    #                'Date Reported':'Orangization Code','Last Update Date':'Problem Summary'}, inplace= True)

    df['Test Station(s)'] = df['Test Station(s)'].str.upper()
    df['Test Station(s)'] = df['Test Station(s)'].str.replace(" ", "")
    # Open the dictonary file
    xl = pd.ExcelFile(dictonary_file_path)
    print(xl.sheet_names)
    xl.close()
    # 2. Open Excel File Read with the Correct Sheet
    df_dictonary = pd.read_excel(io=dictonary_file_path, sheet_name=xl.sheet_names[2])
    print("AAAA")
    # Create empty data frame column

    # applicable for user choose number 1 and 3
    if usr == 1:
        df['Station Name'] = "Undefined"
        df['Location'] = "Undefined"
        df['TE PN'] = "Undefined"
        df['Nearest Correction Value'] = "-"
        df['Original Value'] = "-"

    # prediction_dict = data_analysis()
    cannot_find = False

    for row in range(0, len(df)):
        valueOut = df['Test Station(s)'][row]
        df_filter = df_dictonary.loc[df_dictonary['Station Name'] == valueOut]
        try:

            df.loc[row, 'Station Name'] = df_filter['Station Name'].values[0]
            df.loc[row, 'Location'] = df_filter['Location'].values[0]
            df.loc[row, 'TE PN'] = df_filter['TE PN'].values[0]

        except:

            cannot_find = True
            # pull to other file need to add  auto checking
            print(row)
            nearest_station_name = Excel_Extraction.correction_station_name(valueOut)
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
        print("Some Part number cannot find from dictonary. Please Manual Check it ")
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
        df.loc[row, 'Test Station(s)'] = corr_value
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


def main(dictonary_file_path=None, usr1=1, input_file=None, save_file_path=None):
    if input_file != None or save_file_path != None:
        try:
            input_file = open(input_file, "r+")  # or "a+", whatever you need
            # save_file_path = open(save_file_path, "r+")  # or "a+", whatever you need
        except IOError:
            print("Could not open file! Please close Excel manually!")
            # close file
            return
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

    [df, usr] = data_frame_merging_filter(usr1, input_file, save_file_path)

    if usr == 2:
        file_store = grouping(data_frame=df, table=table_list)
        print(file_store)
        file_filter_to_save = create_new_file(save_file_path, 'templatefile.txt')
        save_file(file_store, file_filter_to_save)
    # Save into the file


if __name__ == "__main__":
    main()
