# 1. open the folder and read the filename get the Data from the excel file
# 2. Save the data in the excel file as a DataSet
# 3. Cacluate
import os
import re
import pandas as pd

path = "C:\\Users\\willlee\\Desktop\\DataSet\\"
path_to_excel = "Dummy_text_set.txt"
files = []


# r=root, d=directories, f = files
def main_2():
    i = 0
    for r, d, f in os.walk(path):
        for file in f:
            if '.xlsx' in file:
                # file split
                title = file.split('.')[0]
                met = re.match(r"[A-Za-z]+_[A-Za-z]+_\d{4}", title)
                if not met is None:
                    files.append(os.path.join(r, file))

    for f in files:
        print(f)
        # 1.  file open to and group by Nearest Correction Value
        df = pd.read_excel(f)
        print(df.head())
        df_grouping = df.groupby(by='Nearest Correction Value')
        print(df_grouping)
        # df loctation
        df = df.loc[df['Nearest Correction Value'] != '-']
        # save in the excel
        df_a = df[['Test Station(s)', 'Nearest Correction Value', 'Original Value']]
        # Appending all value in one dataframe
        if i == 0:
            df_main = df_a
            i = 1
        else:
            df_main = df_main.append(df_a)

    return df_main


def probability_checking(df):
    dict_abc = {}
    # Calcuclate probability
    # convert to set
    # 1. get the unique value from the original station
    # 2. run it one by one and check with original value together with the test station
    # take unique value
    list_of_unique_value = list(set(df['Original Value']))
    # 1. looks for location
    # define location and from location take out the correct value and check whether all correct value are the same by \
    # by grouping  if yes
    for i in list_of_unique_value:
        df_b = df.loc[df['Original Value'] == i]
        # grouping by the correct station name
        df_b_grouping = df_b.groupby(by='Test Station(s)')
        print(df_b_grouping)
        print(len(df_b_grouping))
        # once group check whether have more than one group if more than one group take the group doing AND to check both
        if len(df_b_grouping) == 1:
            # if one group only, simply take the length
            correct_station_name = list(df_b_grouping.groups.keys())[0]
            total_amount_with_the_same_name = len(df_b)
            str = i + "_" + correct_station_name
            dict_abc[str] = total_amount_with_the_same_name
        elif len(df_b_grouping) > 1:
            # if more than one group, Mix it and applying AND group A*B = A*C
            # here is not correct
            for data_user_key_in in list(df_b_grouping.groups.keys()):  # this statement not correct
                # df.loc[(df['Original Value'] == 'CHAS4012102PEN') & (df['Test Station(s)'] == 'CHAS4012102PEN')]
                correct_station_name = data_user_key_in
                # Checking
                df_c = df.loc[(df['Original Value'] == i) & (df['Test Station(s)'] == data_user_key_in)]
                # Checking the length:
                total_amount_with_the_same_name = len(df_c)
                str = i + "_" + correct_station_name
                dict_abc[str] = total_amount_with_the_same_name
        else:
            print("Nothing match")
        # put in dictonary
    print()
    return dict_abc


def dict_key_separation(dict_abc):
    # 1.get key
    dict_corr_value_probability = {}
    dict_ori_value_total = {}
    ori_value_list = []
    all_keys = dict_abc.keys()
    # convert dictionary to list
    for key in list(all_keys):
        ori_value, corr_value = key.split('_')
        ori_value_list.append(ori_value)
    unique_value_list = list(set(ori_value_list))
    # look for the list and find the unique value

    for unique_value in unique_value_list:
        total = 0
        for key in list(all_keys):
            ori_value, corr_value = key.split('_')
            if ori_value == unique_value:
                total = total + dict_abc[key]
        # store in the dictionary
        dict_ori_value_total[unique_value] = total
    # Calculation probability
    for unique_value in unique_value_list:
        for key in list(all_keys):
            ori_value, corr_value = key.split('_')
            if ori_value == unique_value:
                # get the value
                total = dict_ori_value_total[unique_value]
                value_corr_value = dict_abc[ori_value + '_' + corr_value]
                dict_corr_value_probability[ori_value + '_' + corr_value] = value_corr_value / total
    return dict_corr_value_probability


def rules_operation():
    df = main_2()
    dict_abc = probability_checking(df)
    dict_corr_value_probability = dict_key_separation(dict_abc)
    return dict_corr_value_probability


if __name__ == "__main__":
    rules_operation()
