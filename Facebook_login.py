from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
import pickle
from time import sleep, time
from bs4 import BeautifulSoup as bs4
from bs4.element import NavigableString, Tag
from datetime import datetime, date
from tkinter import Listbox, END
import os
import re
import Excel_Extraction

dictonary_file_path = file_name = r"C:\Users\willlee\Desktop\Manufacturing Issue\2020\Test Station Part Number.xlsx"
full_path = os.path.abspath(os.getcwd())
# file_save = full_path + "\May_Data_2020_2.txt"
file_save = 'Global String'
Title = "Issue Number" + "\t" + "Planner Code" + "\t" + "Product Description" + "\t" + "Managment Stripe" + "\t" + \
        "P7_TEST_ITA" + "\t" + "P7_TEST_FIXTURE" + "\t" + "P7_TEST_KIT" + "\t" + "P7_POT_SERIAL_NUMBER" + "\t " + "P7_ROOT_CAUSE_TYPE_ID" + \
        "P7_ROOT_CAUSE_TYPE_ID" + "\t" + "P7_ROOT_CAUSE_SUBTYPE_ID" + "\t" + "P7_ROOT_CAUSE_DESCRIPTION" + "\t" \
        + "Severity" + "\t" + "Status" + "\t" + "Type" + "\t" \
        + "Category" + "\t" + "Assignee Name" + "\t" + "Reviewer Name" + "\t" + "Days Open" + "\t" \
        + "Test Station(s)" + "\t" + "Date Reported" + "\t" + "Last Update Date" + "\t" + "Orangization Code" + "\t" \
        + "Problem Summary"


def write_log_message(ListBox, text):
    ListBox.insert(END, text)
    # ListBox.select_set(END)
    ListBox.yview(END)


def create_new_file(file_save, ext_change):
    filename, ext = os.path.splitext(file_save)
    filepath, file_name = os.path.split(filename)
    return filepath + "\\" + file_name + "." + ext_change


def rename_file_ext(file_save, ext_change):
    pre, ext = os.path.splitext(file_save)
    return os.rename(file_save, pre + ext_change)


def excel_extraction_function(dictonary_file_path):
    return Excel_Extraction.main(dictonary_file_path)


def text_splitter(text, regex_expression, take_which_group=1):
    matches = re.finditer(regex_expression, text, re.MULTILINE)

    for matchNum, match in enumerate(matches, start=1):

        print("Match {matchNum} was found at {start}-{end}: {match}".format(matchNum=matchNum, start=match.start(),
                                                                            end=match.end(), match=match.group()))

        for groupNum in range(0, len(match.groups())):
            groupNum = groupNum + 1

            print("Group {groupNum} found at {start}-{end}: {group}".format(groupNum=groupNum,
                                                                            start=match.start(groupNum),
                                                                            end=match.end(groupNum),
                                                                            group=match.group(groupNum)))
    #   if take_which_group == 1:
    #       return match.group(1)
    #   elif take_which_group == 2:
    #       return match.group(2)
    #   else:
    #       print ("No Defined Group. Assume Group 1")
    #       return match.group(1)
    try:
        match.group(take_which_group)
    except NameError:
        match = None
    if match == None:
        return "No Defined"
    else:
        return match.group(take_which_group)


def table_lookup(driver, this_page):
    # Should be DateNow
    year_to_check = datetime.now()
    # Or User defined.
    #  year_to_check = datetime(2019,1,1,0,0,0)
    still_have_next_page = False
    sleep(5)
    page_source = driver.page_source
    soup = bs4(page_source, 'lxml')
    sleep(5)
    # get new page data
    base_addr = text_splitter(driver.current_url, r"(https:\/\/\w+\.\w+.com\/\w+\/)", 1)
    # Get the page number
    pag_num = soup.find('span', class_='a-IRR-pagination-label')
    #    this_page = pag_num.text
    if this_page == "ABC":
        this_page = pag_num.text
    else:  # Compare
        if this_page == pag_num.text:
            print("No new page. Test will stop")
            return False, this_page
        else:
            this_page = pag_num.text
    # Data Process
    # ------------------------------------------------------------------------------------------------------------
    table = soup.find(lambda tag: tag.name == 'table' and tag.has_attr('id') and tag['id'] == '4125346614174390009')
    rows = table.findAll('tr')
    for row in rows:
        column = 0
        write_in = False
        print(row)
        t_row = []
        for td in row.findAll('td'):
            # Put into list
            print(td.text)
            t_row.append(td.text)
            write_in = True
            if column == 0:
                text_out = text_splitter(str(td), r"<a\s+(?:[^>]*?\s+)?href=([\"\'])(.*?)\1", 2)
                # Open as a new Tab
                output = data_open_in_new_tab(driver, base_addr, text_out)
                for single_output in output: t_row.append(single_output)
                column = column + 1  # To disable for other column to open
        if write_in:
            # Date and time to checking
            #            if text_to_date_conversion(t_row, year_to_check):
            # data_presentation_textfile_pickle(t_row,0)
            data_presentation_textfile(t_row, 0)
    # ------------------------------------------------------------------------------------------------------------
    # finding whether the arrow being found -- for last page use only
    try:
        if driver.find_element_by_xpath(
                "/html/body/form/div[2]/div[2]/div[2]/div/div/div[2]/div/div/div/div/div[2]/div[2]/div[6]/div[2]/ul/li[3]/button"):
            driver.find_element_by_xpath(
                "/html/body/form/div[2]/div[2]/div[2]/div/div/div[2]/div/div/div/div/div[2]/div[2]/div[6]/div[2]/ul/li[3]/button").click()
            still_have_next_page = True
    except:
        print("Not Next Button being Found! Program End Now")

        return False, this_page
    else:
        return still_have_next_page, this_page


def data_open_in_new_tab(driver, base_address, heflink):  # <------Try tomorrow
    Full_addr = base_address + heflink
    # Open Window
    driver.execute_script("window.open('');")
    # Switch to the new window and open URL B
    driver.switch_to.window(driver.window_handles[1])
    driver.get(Full_addr)
    #    Test
    #    ITA\(s\)\s + (?:(.*))\s + Test
    # …Do something here ---  Data Process
    # https://stackoverflow.com/questions/43376074/optional-matching-in-regex
    # format = dictonary_key: ['readwhich content, text or context','tag_name', 'id name to find','expected id to look for',r'regex',group interest]
    dictonary_for_extraction = {
        1: ['text', 'h2', 'P7_PRODUCT_INFORMATION_REGION_heading', 'Product', r'Planner Code:\s+(\w+)', 1],
        2: ['text', 'h2', 'P7_PRODUCT_INFORMATION_REGION_heading', 'Product',
            r'Product Description:\s+(?:(.*?)(?: - |$))?(?:(.*?)(?:, |$))?', 1],
        3: ['text', 'h2', 'P7_PRODUCT_INFORMATION_REGION_heading', 'Product',
            r'Mgmt Stripe:\s+(?:(.*?)(?: - |$))?(?:(.*?)(?:, |$))?', 1],
        4: ['content', 'select', 'P7_TEST_ITA', 'selected',
            r'selected', 1],
        5: ['content', 'select', 'P7_TEST_FIXTURE', 'selected',
            r'selected', 1],
        6: ['content', 'input', 'P7_TEST_KIT', 'value',
            r'nil', 1],
        7: ['content', 'input', 'P7_POT_SERIAL_NUMBER', 'value',
            r'nil', 1],
        8: ['content', 'select', 'P7_ROOT_CAUSE_TYPE_ID', 'selected',
            r'selected', 1],
        9: ['content', 'select', 'P7_ROOT_CAUSE_SUBTYPE_ID', 'selected',
            r'selected', 1],
        10: ['content', 'input', 'P7_ROOT_CAUSE_DESCRIPTION', 'value',
             r'nil', 1],

    }
    #    dictonary_for_extraction  = {1:['P7_PRODUCT_INFORMATION_REGION_heading','Product',r'Planner Code:\s+(\w+)',1],
    #                                 2: ['P7_PRODUCT_INFORMATION_REGION_heading', 'Product', r'Product Description:\s+(?:(.*?)(?: - |$))?(?:(.*?)(?:, |$))?', 1],
    #                                 3: ['P7_PRODUCT_INFORMATION_REGION_heading', 'Product', r'Mgmt Stripe:\s+(?:(.*?)(?: - |$))?(?:(.*?)(?:, |$))?', 1]}
    # 4:['P7_TEST_STATION_INFORMATION_REGION_heading','Test Stations',r'Test ITA\(s\)\s+(\w+)\s××',1]}

    output_list = webpage_extract(driver, dictonary_for_extraction)

    print("Current Page Title is : %s" % driver.title)  # Close the tab with URL B
    driver.close()  # Switch back to the first tab with URL A
    driver.switch_to.window(driver.window_handles[0])
    print("Current Page Title is : %s" % driver.title)
    return output_list


def text_to_date_conversion(text_list, year_to_check):
    str_date = text_list[9]
    # Processing text file
    split_date = str_date.split(' ')
    date = split_date[0].split('-')
    d, m, y = date[0].split('-')
    text_date = m + " " + d + " " + y
    recorded_date = datetime.strptime(text, '%b %d %Y')
    # If more than 365 date will not check
    if int(str(year_to_check - recorded_date).split(',')[0].split(' ')[0]) > 365:
        return False
    else:
        return True


def data_presentation_textfile_pickle(text_list, i):
    text = ""
    found = False
    if os.path.exists(file_save):  # True
        with open(file_save, 'r') as fp:
            chunk = fp.readline()  # string # skip for title
            while True:
                chunk = fp.readline()  # string
                # Separate
                if chunk:
                    words = chunk.split("\t")
                    if text_list[0] == words[0]:
                        found = True
                        i = i + 1
                        # put 50 for continue searching
                        #                        if i > 50:
                        break
                    else:
                        found = False
                # no next line find found, stop reading
                else:
                    break

        if not found:
            with open(file_save, 'a') as fp:
                text = list_to_string(text_list)
                fp.write(text + "\n")


    else:
        # file not exists
        with open(file_save, 'w') as fp:
            # create
            fp.write(Title + "\n")
            text = list_to_string(text_list)
            fp.write(text + "\n")


def data_presentation_textfile(text_list, i):
    # Save in file
    text = ""
    found = False
    if os.path.exists(file_save):  # True
        with open(file_save, 'r') as fp:
            chunk = fp.readline()  # string # skip for title
            while True:
                chunk = fp.readline()  # string
                # Separate
                if chunk:
                    words = chunk.split("\t")
                    if text_list[0] == words[0]:
                        found = True
                        i = i + 1
                        # put 50 for continue searching
                        #                        if i > 50:
                        break
                    else:
                        found = False
                # no next line find found, stop reading
                else:
                    break

        if not found:
            with open(file_save, 'a') as fp:
                text = list_to_string(text_list)
                fp.write(text + "\n")


    else:
        # file not exists
        with open(file_save, 'w') as fp:
            # create
            fp.write(Title + "\n")
            text = list_to_string(text_list)
            fp.write(text + "\n")


def list_to_string(text_list):
    text = ""
    for i in range(len(text_list)):
        if i < len(text_list) - 1:
            text = text + text_list[i] + "\t"
        elif i == len(text_list) - 1:
            text = text + text_list[i]
        else:
            pass
    return text


def webpage_extract(driver, *args):  # ["id to check", "compare value", ......] or put as dictonary
    page_source = driver.page_source
    sleep(2)
    soup = bs4(page_source, 'html')
    returnValue = []
    # Checking the length of the args
    for i in args:
        for j in i.values():  # j[0] = id_to_Compare j[1]= text_to_compare ,j[2] = regex, j[3] = index_to_get
            if j[0] == 'text':
                reviews_selectors = soup.find_all('div', class_='t-Body-main')
                for reviews_selector in reviews_selectors:
                    review_div = reviews_selector.find_all('div', class_='t-Region')
                    # Assume take out the Test_station_Information_region... can change accordingly
                    # Since array[0] having all information, we just access first one only
                    #                   for single_rev_div in review_div[0]:
                    for single_content in review_div[0].contents:
                        # look for whether it is bs4.ewlement.Tag (Hardcoded)
                        if isinstance(single_content, Tag):
                            # find for the banner that interested
                            try:
                                # single_content.find('h2',id = j[0]).text == j[1]:
                                if single_content.find(j[1], id=j[2]).text == j[3]:
                                    # directly goto contents[3]
                                    string_to_process = single_content.contents[3].text
                                    string_interest = text_splitter(string_to_process, j[4], j[5])
                                    if string_interest:
                                        returnValue.append(string_interest)
                                        string_interest = ""
                                        break  # When found break
                                    else:
                                        string_interest = "No Defined"
                                        returnValue.append(string_interest)
                            except:
                                pass
                            else:
                                pass
            elif j[0] == 'content':
                i = -1
                string_interestA = []
                some_text = ''
                reviews_selectors = soup.find_all(j[1], id=j[2])
                for reviews_selector in reviews_selectors:
                    # look whether it is an Nagivable String or Tag
                    if isinstance(reviews_selector, Tag):
                        if j[1] == 'input':
                            # input maybe no need to loop do input need to loop?
                            # for review_select in reviews_selector.contents:
                            # read for the attrs value look for the words selected
                            # if isinstance(review_select, Tag):
                            try:
                                string_interest = reviews_selector[j[3]]
                                string_interestA.append(string_interest)
                            except:
                                pass
                            if len(string_interestA) != 0:
                                for str_interestA in string_interestA:
                                    some_text = some_text + str_interestA + ';'
                            else:
                                some_text = string_interest

                            returnValue.append(some_text)
                            string_interest = "-"
                            # cannot break due to user might put more value

                        elif j[1] == 'select':
                            # read for contesnts
                            i = 0
                            for review_select in reviews_selector.contents:
                                if isinstance(review_select, Tag):
                                    try:
                                        if review_select[j[3]] == j[4]:
                                            string_interest = review_select.text
                                            string_interestA.append(string_interest)
                                        elif review_select[j[3]] != j[4]:
                                            pass  # not sure do what??
                                        else:
                                            pass
                                    except:
                                        pass
                            if len(string_interestA) != 0:
                                for str_interestA in string_interestA:
                                    some_text = some_text + str_interestA + ';'
                            else:
                                some_text = string_interest
                            returnValue.append(some_text)
                            string_interest = "-"

                        else:  # future purposed
                            pass

            else:
                pass

    print(returnValue)
    return returnValue


def data_presentation_other_file(text_list):
    pass


def automate(username, password, selection, start_date, end_date, source_file, ListBox):
    # table_station = excel_extraction_function(dictonary_file_path)
    global file_save  # Make the file save as a global
    start_time = time()
    if username is None:
        usr = input('Enter Email Id:')
    else:
        usr = username
    if password is None:
        pwd = input('Enter Password')
    else:
        pwd = password

    driver = webdriver.Firefox()
    #driver.get('https://apex.natinst.com/apex/f?p=NIAPEXMENU:101')
    driver.get('https://echo.natinst.com/')
    print('Opened Echo')
    #driver.find_elements_by_css_selector('div.tile:nth-child(1) > div:nth-child(1)')

    # write_log_message(ListBox,'Opened APEX')
    print('Opened APEX')
    sleep(2)
    username_box = driver.find_element_by_id('i0116')
    #username_box = driver.find_element_by_id('P101_USERNAME')
    username_box.send_keys(usr)
    # write_log_message(ListBox, 'Email Id entered')
    print('Email Id entered')
    sleep(2)

    driver.find_element_by_id('idSIButton9').click()
    sleep(2)

    # User need to enter email and password id at the message box due to no way to control

    input_part = driver.find_element_by_id('inputPart')
    input_part.send_keys('159572B-000L')

    sleep(10)

    search_box = driver.find_element_by_id('searchButton')
    search_box.click()






    password_box = driver.find_element_by_id('P101_PASSWORD')
    password_box.send_keys(pwd)
    print('Password Entered')
    # write_log_message(ListBox, 'Password Entered')
    sleep(2)

    login_box = driver.find_element_by_xpath('/html/body/form/div[1]/div/div/div/div/div/div/div/div[2]/button')
    # login_box = driver.find_element_by_id('P101_LOGIN')
    login_box.click()

    print("Done")
    # write_log_message(ListBox, 'Done')
    sleep(1)
    driver.get(driver.current_url)
    sleep(7)
    # minimize the window
    driver.minimize_window()

    page_source = driver.page_source
    soup = bs4(page_source, 'lxml')
    a = soup.find_all('li', class_='a-TreeView-node')
    driver.find_element_by_link_text("Manufacturing").click()
    # driver.find_element_by_xpath("/html/body/form/div[1]/div/div[2]/div/div/div[1]/div/div/div[2]/div[2]/div/div/div/div/div[2]/div[2]/div/ul/li/ul/li[6]/ul/li[2]/div[2]/a").click()
    driver.find_element_by_xpath(
        "/html/body/form/div[1]/div/div[2]/div/div/div[1]/div/div/div[2]/div[2]/div/div/div/div/div[2]/div[2]/div/ul/li/ul/li[6]/span").click()
    page_source = driver.page_source
    # using javascript soon
    driver.find_element_by_css_selector('#authTree_tree_12 > div:nth-child(2) > a:nth-child(1)').click()
    sleep(10)
    # remove all the filter
    driver.get(driver.current_url)
    sleep(2)
    page_source = driver.page_source
    sleep(2)
    soup = bs4(page_source, 'lxml')

    Item_To_Remove = soup.find_all('span', class_='a-IRR-controls-cell a-IRR-controls-cell--remove')
    # if remove come out if clicking
    for i in range(len(Item_To_Remove)):
        sleep(4)
        driver.find_element_by_css_selector(
            "li.a-IRR-controls-item:nth-child(1) > span:nth-child(4) > button:nth-child(1)").click()
        sleep(4)
    # driver.find_element_by_class_name("TreeView-node").click()
    driver.find_element_by_css_selector('#P14_ORGANIZATION_CODE > option:nth-child(7)').click()  # <--cover
    sleep(3)
    driver.find_element_by_css_selector('#P14_ISSUES_REGION_row_select > option:nth-child(10)').click()  # <--cover
    # Press Search
    que = driver.find_element_by_css_selector('#P14_ISSUES_REGION_search_field')  # <--cover
    que.send_keys('PEN Mfg Test Tech')
    driver.find_element_by_css_selector('#P14_ISSUES_REGION_column_search_root').click()  # <--cover
    sleep(3)
    page_source = driver.page_source
    sleep(3)
    driver.find_element_by_css_selector('#P14_ISSUES_REGION_column_search_drop_2_c5i').click()  # <--cover
    driver.find_element_by_css_selector('#P14_ISSUES_REGION_search_button').click()  # <--cover
    sleep(5)
    if selection is None:
        usrSelection = input('Enter You want to look for closed(0) or waiting for 3rd Party(1):')
    else:
        usrSelection = selection

    if usrSelection == 1:
        que = driver.find_element_by_css_selector('#P14_ISSUES_REGION_search_field')
        sleep(5)
        que.send_keys('Closed')
    elif usrSelection == 2:
        que = driver.find_element_by_css_selector('#P14_ISSUES_REGION_search_field')
        sleep(5)
        que.send_keys('Waiting for 3rd Party')

    driver.find_element_by_css_selector('#P14_ISSUES_REGION_column_search_root').click()
    sleep(3)
    page_source = driver.page_source
    sleep(3)
    driver.find_element_by_css_selector('#P14_ISSUES_REGION_column_search_drop_2_c2i').click()
    driver.find_element_by_css_selector('#P14_ISSUES_REGION_search_button').click()
    # Press for filter
    driver.find_element_by_css_selector('#P14_ISSUES_REGION_actions_button').click()
    sleep(3)
    page_source = driver.page_source
    sleep(3)
    driver.find_element_by_css_selector('#P14_ISSUES_REGION_actions_menu_2i').click()
    sleep(3)
    page_source = driver.page_source
    sleep(3)
    driver.find_element_by_css_selector('#P14_ISSUES_REGION_column_name').click()
    sleep(3)
    # Date Reported
    driver.find_element_by_xpath(
        '/html/body/div[7]/div[2]/div[2]/table[2]/tbody/tr[1]/td/table/tbody/tr[2]/td[1]/select/optgroup[1]/option[10]').click()  # <---Cannot cover
    sleep(3)
    driver.find_element_by_css_selector('#P14_ISSUES_REGION_DATE_OPT > option:nth-child(13)').click()
    sleep(3)
    que = driver.find_element_by_css_selector('#P14_ISSUES_REGION_between_from')
    que.send_keys(start_date)
    que = driver.find_element_by_css_selector('#P14_ISSUES_REGION_between_to')
    que.send_keys(end_date)
    driver.find_element_by_xpath('/html/body/div[7]/div[3]/div/button[2]').click()
    sleep(3)
    # ==============Press for filter ----------------------
    # driver.find_element_by_css_selector('#P14_ISSUES_REGION_actions_button').click()
    # sleep(3)
    # page_source = driver.page_source
    # sleep(3)
    # driver.find_element_by_css_selector('#P14_ISSUES_REGION_actions_menu_2i').click()
    # sleep(3)
    # page_source = driver.page_source
    # sleep(3)
    # driver.find_element_by_css_selector('#P14_ISSUES_REGION_column_name').click()
    # sleep(3)
    # #Last Update Date Reported
    # driver.find_element_by_xpath('/html/body/div[7]/div[2]/div[2]/table[2]/tbody/tr[1]/td/table/tbody/tr[2]/td[1]/select/optgroup[1]/option[11]').click()
    # sleep(3)
    # driver.find_element_by_css_selector('#P14_ISSUES_REGION_DATE_OPT > option:nth-child(13)').click()
    # sleep(3)
    # que = driver.find_element_by_css_selector('#P14_ISSUES_REGION_between_from')
    # que.send_keys('1-MAY-2020')
    # que = driver.find_element_by_css_selector('#P14_ISSUES_REGION_between_to')
    # que.send_keys('31-MAY-2020')
    # driver.find_element_by_xpath('/html/body/div[7]/div[3]/div/button[2]').click()
    sleep(5)

    file_save = create_new_file(source_file, 'txt')
    print(file_save)
    default_page = "ABC"
    while True:
        next_page_valid, next_page_numer = table_lookup(driver, default_page)

        if not next_page_valid:
            print(next_page_numer)
            break
        else:
            default_page = next_page_numer

    print("--- %s seconds ---" % (time() - start_time))

    # ---------------------end----------------------------------------------------------

    # ---------- Will defined in whole function ----------------------------------------

    # page_source = driver.page_source
    # sleep(10)
    # soup = bs4(page_source,'lxml')
    # # get new page data
    #
    # # data process
    # #------------------------------------------------------------------------------------------------------------
    # table = soup.find(lambda tag: tag.name =='table' and tag.has_attr('id') and tag['id']=='4125346614174390009')
    # rows = table.findall('tr')
    # for row in rows:
    #     print (row)
    #     t_row ={}
    #     for td in row.findall('td'):
    #         print (td.text)
    # #------------------------------------------------------------------------------------------------------------
    # # finding whether the arrow being found -- for last page use only
    # if  driver.find_element_by_css_selector('.a-irr-button--pagination'):
    #     driver.find_element_by_css_selector('.a-irr-button--pagination').click()
    #     have_next_page = true
    # #-----------------------------------------------------------------------------------
    # #found = soup.find('div',{'id':'authtree_tree'})
    # #driver.find_element_by_css_selector("#authtree_tree_6 > div:nth-child(3) > a:nth-child(1)").click()
    # #driver.find_element_by_id('authtree_tree')
    # #for i in found.children:
    # #    print (i)
    # #driver.find_element_by_class_name("a-treeview-toggle").click()
    # #page_source = driver.page_source
    # #driver.find_element_by_css_selector("#authtree_tree_10 > div:nth-child(2) > a:nth-child(1)").click()
    # #print (frame)
    # #manufacturing_issue = driver.find_element_by_id('apexcbmdummyselection')
    # #print (manufacturing_issue)
    # input('Press anything to quit')
    driver.close()
    driver.quit()
    print("Finished")
    rename_file_ext(file_save, '.csv')
    return "Finish"


if __name__ == "__main__":
    automate(None, None, None, '1-JAN-2020', '30-JUNE-2020', None, None)
