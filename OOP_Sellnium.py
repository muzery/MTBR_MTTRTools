# https://stackoverflow.com/questions/24001427/use-selenium-webdriver-as-a-baseclass-python
from selenium import webdriver
from bs4 import BeautifulSoup
from time import sleep, time
from bs4.element import NavigableString, Tag
import os
import functools
from webpage_extraction import WebpageExtract, DataFrameFeature

# https://qxf2.com/blog/auto-generate-xpaths-using-python/
# here is hardcoded
Title = "Issue Number" + "\t" + "Planner Code" + "\t" + "Product Description" + "\t" + "Managment Stripe" + "\t" + \
        "P7_TEST_ITA" + "\t" + "P7_TEST_FIXTURE" + "\t" + "P7_TEST_KIT" + "\t" + "P7_POT_SERIAL_NUMBER" + "\t " + "P7_ROOT_CAUSE_TYPE_ID" + \
        "P7_ROOT_CAUSE_TYPE_ID" + "\t" + "P7_ROOT_CAUSE_SUBTYPE_ID" + "\t" + "P7_ROOT_CAUSE_DESCRIPTION" + "\t" \
        + "Severity" + "\t" + "Status" + "\t" + "Type" + "\t" \
        + "Category" + "\t" + "Assignee Name" + "\t" + "Reviewer Name" + "\t" + "Days Open" + "\t" \
        + "Test Station(s)" + "\t" + "Date Reported" + "\t" + "Last Update Date" + "\t" + "Orangization Code" + "\t" + "Problem Summary"

Title_Party = "Issue Number" + "\t" + "Planner Code" + "\t" + "Product Description" + "\t" + "Managment Stripe" + "\t" + \
              "P7_TEST_ITA" + "\t" + "P7_TEST_FIXTURE" + "\t" + "P7_TEST_KIT" + "\t" + "P7_POT_SERIAL_NUMBER" + "\t " + "P7_ROOT_CAUSE_TYPE_ID" + \
              "P7_ROOT_CAUSE_TYPE_ID" + "\t" + "P7_ROOT_CAUSE_SUBTYPE_ID" + "\t" + "P7_ROOT_CAUSE_DESCRIPTION" + "\t" \
              + "Complete?" + "\t" + "Days_Open_2" + "\t" + "Severity" + "\t" + "Status" + "\t" + "Type" + "\t" \
              + "Category" + "\t" + "Assignee Name" + "\t" + "Reviewer Name" + "\t" + "Days Open" + "\t" \
              + "Test Station(s)" + "\t" + "Date Reported" + "\t" + "Last Update Date" + "\t" + "Orangization Code" + "\t" + "Problem Summary"

PATH = "C:\Program Files (x86)\chromedriver.exe"


class DriverClass(webdriver.Firefox, webdriver.Ie, webdriver.Chrome):

    def __init__(self):
        # super().__init__()
        print('In Child Class')

    def setup_page(self, browser):
        if browser.lower() == 'ie':
            return webdriver.Ie.__init__(self)
        elif browser.lower() == 'chrome':
            return webdriver.Chrome(PATH)
        else:
            return webdriver.Firefox()

    def element_by_text_get_x_path(self, text):
        element = self.driver.find_element_by_link_text(text)
        generated_xpath = self.driver.execute_script("return window.getPathTo(arguments[0]);", element)
        print(generated_xpath)
        self.driver.find_element_by_xpath(generated_xpath).click()

    def retrieve_first_page(self):
        pass

    def display_file(self, text):
        print(text)

    def website_parse_all_content(self, page_source, parser, tag_look_for='li', cls='a-TreeView-node'):
        self.soup = BeautifulSoup(page_source, parser)
        return self.soup.find_all(tag_look_for, class_=cls)

    def __recursive_loop__(self, contents, main_tree, class_name_to_look_for, attribute, i=0):
        internal_ctr = 0
        found = False
        # Check for the amount of the class name to look for
        for counter in range(i, len(class_name_to_look_for), 1):
            # retrieve class
            for cnt in range(0, len(contents), 1):
                # look for the type
                if not isinstance(contents[cnt], NavigableString):
                    internal_ctr += 1
                    if isinstance(contents[cnt].attrs[attribute], list):
                        if contents[cnt].attrs[attribute][0] == class_name_to_look_for[counter] and isinstance(
                                contents[cnt], Tag):
                            child = internal_ctr
                            main_tree = main_tree + " > " + contents[cnt].name + ":" + "nth-child" + "(" + str(
                                child) + ")"
                            found = True
                            if len(contents[cnt].contents) != 0:
                                i = i + 1  # look for search file
                                main_tree, complete = self.__recursive_loop__(contents[cnt].contents, main_tree,
                                                                              class_name_to_look_for, attribute, i)
                            else:
                                return main_tree, True
                            # if found: return main_tree,True  # if found break the loop
                    elif isinstance(contents[cnt].attrs[attribute], str):
                        if contents[cnt].attrs[attribute] == class_name_to_look_for[counter] and isinstance(
                                contents[cnt], Tag):
                            child = internal_ctr
                            main_tree = main_tree + " > " + contents[cnt].name + ":" + "nth-child" + "(" + str(
                                child) + ")"
                            found = True
                            if len(contents[cnt].contents) != 0:
                                i = i + 1  # look for search file
                                main_tree, complete = self.__recursive_loop__(contents[cnt].contents, main_tree,
                                                                              class_name_to_look_for, attribute, i)
                            else:
                                return main_tree, True
                    else:
                        pass

            if not found:
                # if search nothing for 2nd element, break, will not continue
                main_tree = main_tree + 'empty'  # this is empty
                return main_tree  # break the for loop
        return main_tree, True

    def ____recursive_loop_list__(self, contents, main_tree, class_name_to_look_for, attribute, i=0):
        internal_ctr = 0
        found = False
        # Check for the amount of the class name to look for
        for counter in range(i, len(class_name_to_look_for), 1):
            # retrieve class
            for cnt in range(0, len(contents), 1):
                # look for the type
                if not isinstance(contents[cnt], NavigableString):
                    internal_ctr += 1
                    if isinstance(contents[cnt].attrs[attribute], list):
                        if contents[cnt].attrs[attribute][0] == class_name_to_look_for[counter] and isinstance(
                                contents[cnt], Tag):
                            child = internal_ctr
                            main_tree = main_tree + " > " + contents[cnt].name + ":" + "nth-child" + "(" + str(
                                child) + ")"
                            found = True
                            if len(contents[cnt].contents) != 0:
                                i = i + 1  # look for search file
                                main_tree, complete = self.__recursive_loop__(contents[cnt].contents, main_tree,
                                                                              class_name_to_look_for, attribute, i)
                            else:
                                return main_tree, True
                            # if found: return main_tree,True  # if found break the loop
                    elif isinstance(contents[cnt].attrs[attribute], str):
                        if contents[cnt].attrs[attribute] == class_name_to_look_for[counter] and isinstance(
                                contents[cnt], Tag):
                            child = internal_ctr
                            main_tree = main_tree + " > " + contents[cnt].name + ":" + "nth-child" + "(" + str(
                                child) + ")"
                            found = True
                            if len(contents[cnt].contents) != 0:
                                i = i + 1  # look for search file
                                main_tree, complete = self.__recursive_loop__(contents[cnt].contents, main_tree,
                                                                              class_name_to_look_for, attribute, i)
                            else:
                                return main_tree, True
                    else:
                        pass

            if not found:
                # if search nothing for 2nd element, break, will not continue
                main_tree = main_tree + 'empty'  # this is empty
                return main_tree  # break the for loop
        return main_tree, True

    def soup_parsing_by_text_name(self, all_elements_tag, search_by_textname, args):
        complete = False
        main_tree = ''  # change to dictionary more directive
        if len(all_elements_tag) == 0:
            return
        for element_tag in all_elements_tag:
            if complete: break
            if element_tag.text == search_by_textname:
                # get the id
                main_tree = '#' + element_tag.attrs['id']
                # get the content
                # function_name (internal content, list to store, class to search
                if len(args) != 0:
                    main_tree, complete = self.__recursive_loop__(element_tag.contents, main_tree, args, 'class')

        return main_tree

    def soup_parsing_by_id_name(self, soup, tag_name, search_by_id, args,
                                attr):  # driver.find_element_by_css_selector('#P14_ORGANIZATION_CODE > option:nth-child(7)').click()
        complete = False
        main_tree = ''
        contents = soup.find_all(tag_name, id=search_by_id)
        if isinstance(attr, str):
            if len(contents) != 0:
                for cnt in range(0, len(contents), 1):
                    if isinstance(contents[cnt], Tag):
                        # get id
                        main_tree = '#' + contents[cnt].attrs['id']
                        if len(args) != 0:
                            main_tree, complete = self.__recursive_loop__(contents[cnt].contents, main_tree, args,
                                                                          'value')
            else:
                return main_tree, complete
        elif isinstance(attr, list):
            if len(contents) != 0:
                for cnt in range(0, len(contents), 1):
                    if isinstance(contents[cnt], Tag):
                        # get id
                        main_tree = '#' + contents[cnt].attrs['id']
                        if len(args) != 0 and len(attr) != 0:
                            # here doing retrieve

                            main_tree, complete = self.__recursive_loop_list__(contents[cnt].contents, main_tree, args,
                                                                               attr)
            else:
                return main_tree, complete
        else:
            pass

        return main_tree


def webpage_refresh(driver):
    sleep(3)
    page_source = driver.page_source
    sleep(3)
    soup = BeautifulSoup(page_source, 'lxml')
    return soup


# not a good implementation
def delete(driver, soup):
    # need to improve
    Item_To_Remove = soup.find_all('span', class_='a-IRR-controls-cell a-IRR-controls-cell--remove')
    # if remove come out if clicking
    for i in range(len(Item_To_Remove)):
        sleep(3)
        driver.find_element_by_css_selector(
            "li.a-IRR-controls-item:nth-child(1) > span:nth-child(4) > button:nth-child(1)").click()
        sleep(3)


def first_page_view_control(driver, user_name, password):
    # first _page
    driver.get('https://apex.natinst.com/apex/f?p=NIAPEXMENU:101')

    usr = user_name
    pwd = password
    # usr = input('Enter Email Id:')
    # pwd = input('Enter Password')
    print('Opened APEX')
    sleep(2)
    # driver = DriverClass('firefox')
    # driver.goto_address('https://apex.natinst.com/apex/f?p=NIAPEXMENU:101')

    username_box = driver.find_element_by_id('P101_USERNAME')
    username_box.send_keys(usr)
    print('Email Id entered')
    sleep(2)

    password_box = driver.find_element_by_id('P101_PASSWORD')
    password_box.send_keys(pwd)
    print('Password Entered')
    sleep(2)

    login_box = driver.find_element_by_css_selector('.t-Button').click()
    print("Done")
    # do I need to put some checking at here?


def tree_view_control(driver, feature):
    # this function is control tree view
    str_construct = ''
    driver.get(driver.current_url)
    soup = webpage_refresh(driver)
    # sleep(5)
    # page_source = driver.page_source
    # soup = BeautifulSoup(page_source, 'lxml')
    a = soup.find_all('li', class_='a-TreeView-node')
    str_construct = feature.soup_parsing_by_text_name(a, "Manufacturing", ['a-TreeView-toggle'])
    driver.find_element_by_css_selector(str_construct).click()
    soup = webpage_refresh(driver)
    # soup = BeautifulSoup(page_source, 'lxml')
    b = soup.find_all('li', class_='a-TreeView-node')
    str_construct = feature.soup_parsing_by_text_name(b, "Manufacturing Issues",
                                                      ['a-TreeView-content', 'a-TreeView-label'])
    driver.find_element_by_css_selector(str_construct).click()


def mfg_issue_view(driver, feature, status, start_date, end_date):
    str_construct = ''
    driver.get(driver.current_url)

    soup = webpage_refresh(driver)

    # soup = BeautifulSoup(page_source, 'lxml')
    # ----------delete------------------
    delete(driver, soup)
    # ----------------------------------
    soup = webpage_refresh(driver)
    # soup = BeautifulSoup(page_source, 'lxml')
    str_construct = feature.soup_parsing_by_id_name(soup, 'select', 'P14_ORGANIZATION_CODE', ['PEN'], 'value')
    driver.find_element_by_css_selector(str_construct).click()
    sleep(3)
    str_construct = feature.soup_parsing_by_id_name(soup, 'select', 'P14_ISSUES_REGION_row_select', ['200'], 'value')
    driver.find_element_by_css_selector(str_construct).click()
    sleep(3)
    str_construct = feature.soup_parsing_by_id_name(soup, 'input', 'P14_ISSUES_REGION_search_field', [], 'value')
    que = driver.find_element_by_css_selector(str_construct)
    que.send_keys('PEN Mfg Test Tech')
    sleep(3)
    str_construct = feature.soup_parsing_by_id_name(soup, 'button', 'P14_ISSUES_REGION_column_search_root', [], 'value')
    driver.find_element_by_css_selector(str_construct).click()
    sleep(2)
    soup = webpage_refresh(driver)
    driver.find_element_by_css_selector('#P14_ISSUES_REGION_column_search_drop_2_c5i').click()
    # unable to work on small window
    # str_construct = driver.soup_parsing_by_id_name(soup, 'button', 'P14_ISSUES_REGION_column_search_drop_2_c5i', [])   #<---cannot find
    # driver.find_element_by_css_selector(str_construct).click()
    sleep(2)
    str_construct = feature.soup_parsing_by_id_name(soup, 'button', 'P14_ISSUES_REGION_search_button', [], 'value')
    driver.find_element_by_css_selector(str_construct).click()
    sleep(5)
    str_construct = feature.soup_parsing_by_id_name(soup, 'input', 'P14_ISSUES_REGION_search_field', [], 'value')
    # doing data key in
    que = driver.find_element_by_css_selector(str_construct)
    que.send_keys(status)
    sleep(2)
    # page_source = webpage_refresh(driver)
    str_construct = feature.soup_parsing_by_id_name(soup, 'button', 'P14_ISSUES_REGION_column_search_root', [], 'value')
    driver.find_element_by_css_selector(str_construct).click()
    sleep(2)
    driver.find_element_by_css_selector('#P14_ISSUES_REGION_column_search_drop_2_c2i').click()
    # str_construct = driver.soup_parsing_by_id_name(soup, '', 'P14_ISSUES_REGION_column_search_drop_2_c2i', [])
    # driver.find_element_by_css_selector(str_construct).click()
    str_construct = feature.soup_parsing_by_id_name(soup, 'button', 'P14_ISSUES_REGION_search_button', [], 'value')
    driver.find_element_by_css_selector(str_construct).click()
    str_construct = feature.soup_parsing_by_id_name(soup, 'button', 'P14_ISSUES_REGION_actions_button', [], 'value')
    driver.find_element_by_css_selector(str_construct).click()
    soup = webpage_refresh(driver)
    str_construct = feature.soup_parsing_by_id_name(soup, 'button', 'P14_ISSUES_REGION_actions_menu_2i', [], 'value')
    driver.find_element_by_css_selector(str_construct).click()
    soup = webpage_refresh(driver)
    str_construct = feature.soup_parsing_by_id_name(soup, 'select', 'P14_ISSUES_REGION_column_name', [], '-')
    driver.find_element_by_css_selector(str_construct).click()
    driver.find_element_by_css_selector(
        'option.DATE:nth-child(10)').click()  # <--------Special case, cannot observe any relationship
    soup = webpage_refresh(driver)
    str_construct = feature.soup_parsing_by_id_name(soup, 'select', 'P14_ISSUES_REGION_DATE_OPT', ['between'], 'value')
    driver.find_element_by_css_selector(str_construct).click()
    soup = webpage_refresh(driver)
    str_construct = feature.soup_parsing_by_id_name(soup, 'input', 'P14_ISSUES_REGION_between_from', [], 'value')
    # doing data key in
    que = driver.find_element_by_css_selector(str_construct)
    que.send_keys(start_date)
    str_construct = feature.soup_parsing_by_id_name(soup, 'input', 'P14_ISSUES_REGION_between_to', [], 'value')
    # doing data key in
    que = driver.find_element_by_css_selector(str_construct)
    que.send_keys(end_date)
    sleep(2)
    driver.find_element_by_css_selector('.ui-button--hot').click()
    print("Done")


def webpage_init(file_save):
    return WebpageExtract(file_save, Title)


def create_new_file(file_save, ext_change):
    filename, ext = os.path.splitext(file_save)
    filepath, file_name = os.path.split(filename)
    return filepath + "\\" + file_name + "." + ext_change


def rename_file_ext(file_save, ext_change):
    pre, ext = os.path.splitext(file_save)
    return os.rename(file_save, pre + ext_change)


def teardown(driver):
    driver.close()
    driver.quit()
    print("Finished")


def analyzing_status(status):
    lst = status.split("_")
    if len(lst) == 2:
        return lst[0], lst[1]
    else:
        return lst[0], None


# put an decorator for time recording
def logging_time(func):
    @functools.wraps(func)
    def wrapper_time(*args):
        start_time = time()
        func(*args)
        end_time = time()
        run_time = end_time - start_time
        print("--- %s seconds ---" % (time() - start_time))

    return wrapper_time


# @logging_time
def main(user_name, password, status, start_date, end_date, source_file, browser):
    title1 = ""
    df_ctrl = DataFrameFeature()
    still_have_next_page = False
    update_status = analyzing_status(status)
    feature = DriverClass()
    driver = feature.setup_page(browser)
    # driver.setup_page(browser)
    first_page_view_control(driver, user_name, password)

    driver.implicitly_wait(5)
    sleep(5)
    tree_view_control(driver, feature)
    driver.implicitly_wait(5)
    sleep(5)
    mfg_issue_view(driver, feature, update_status[0], start_date, end_date)
    driver.implicitly_wait(5)
    sleep(5)
    file_save = create_new_file(source_file, 'txt')
    title1 = Title_Party if update_status[1] == 'Party' else Title
    extract = WebpageExtract(file_save, title1)
    print(file_save)
    driver.implicitly_wait(5)
    sleep(5)
    default_page = "ABC"
    while True:
        table_complete, next_page_number = extract.table_lookup(driver,
                                                                default_page, update_status[
                                                                    1])  # make sure whole page is complete, else test will stop
        try:
            if driver.find_element_by_css_selector('.icon-right-chevron'):
                driver.find_element_by_css_selector('.icon-right-chevron').click()
                driver.implicitly_wait(5)
                sleep(5)
                still_have_next_page = True
        except:
            print("Not Next Button being Found! Program End Now")
            break
        else:
            if table_complete:
                if next_page_number:
                    print(next_page_number)
                    default_page = next_page_number
                    print("Goto Next Page)")
                else:
                    print("Scrapping Complete!")
            else:
                print("Table not read completely.... Please Check and code will error!")
                # throw exception
                raise Exception("Table not read completely.... Please Check and code will error!")
                break

    print()
    teardown(driver)
    if update_status[1] == 'Party':
        Title1 = Title
        df_ctrl.column_swapping(file_save, ["Days_Open_2", "Days Open"], True, Title1)
    rename_file_ext(file_save, '.csv')
    return "Finish"


if __name__ == "__main__":
    main(os.environ.get('Username'), 'QuanQuan_90', 'Closed', '1-JAN-2020', '15-JAN-2020',
         r"C:\Users\willlee\Desktop\DataSet\1H_test_2020_2.txt", 'firefox')
