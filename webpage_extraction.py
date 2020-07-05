from time import sleep, time
from bs4 import BeautifulSoup as bs4
from bs4.element import NavigableString, Tag
from datetime import datetime, date
from tkinter import Listbox, END
import os
import re


class WebpageExtract:

    def __init__(self, file_saving, title):
        self.file_save = file_saving
        self.title = title

    # def excel_extraction_function(self,dictonary_file_path):
    #     return Excel_Extraction.main(dictonary_file_path)

    def text_splitter(self, text, regex_expression, take_which_group=1):

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

    def __data_open_in_new_tab__(self, driver, base_address, heflink):  # <------mark as private method
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

        output_list = self.__webpage_extract__(driver, dictonary_for_extraction)

        print("Current Page Title is : %s" % driver.title)  # Close the tab with URL B
        driver.close()  # Switch back to the first tab with URL A
        driver.switch_to.window(driver.window_handles[0])
        print("Current Page Title is : %s" % driver.title)
        return output_list

    def __webpage_extract__(self, driver, *args):  # ["id to check", "compare value", ......] or put as dictonary
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
                                        string_interest = self.text_splitter(string_to_process, j[4], j[5])
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

    def __list_to_string__(self, text_list):
        text = ""
        for i in range(len(text_list)):
            if i < len(text_list) - 1:
                text = text + text_list[i] + "\t"
            elif i == len(text_list) - 1:
                text = text + text_list[i]
            else:
                pass
        return text

    def __data_presentation_textfile__(self, text_list, i):
        # Save in file
        text = ""
        found = False
        if os.path.exists(self.file_save):  # True
            with open(self.file_save, 'r') as fp:
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
                with open(self.file_save, 'a') as fp:
                    text = self.__list_to_string__(text_list)
                    fp.write(text + "\n")


        else:
            # file not exists
            with open(self.file_save, 'w') as fp:
                # create
                fp.write(self.title + "\n")
                text = self.__list_to_string__(text_list)
                fp.write(text + "\n")

    def table_lookup(self, driver, this_page):
        # Should be DateNow
        ctr = 0
        year_to_check = datetime.now()
        # Or User defined.
        #  year_to_check = datetime(2019,1,1,0,0,0)
        still_have_next_page = False
        sleep(5)
        page_source = driver.page_source
        soup = bs4(page_source, 'lxml')
        sleep(5)
        # get new page data
        base_addr = self.text_splitter(driver.current_url, r"(https:\/\/\w+\.\w+.com\/\w+\/)", 1)
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
                    text_out = self.text_splitter(str(td), r"<a\s+(?:[^>]*?\s+)?href=([\"\'])(.*?)\1", 2)
                    # Open as a new Tab
                    output = self.__data_open_in_new_tab__(driver, base_addr, text_out)
                    for single_output in output: t_row.append(single_output)
                    column = column + 1  # To disable for other column to open
            if write_in:
                # Date and time to checking
                #            if text_to_date_conversion(t_row, year_to_check):
                # data_presentation_textfile_pickle(t_row,0)
                ctr = ctr + 1
                self.__data_presentation_textfile__(t_row, 0)
        # ------------------------------------------------------------------------------------------------------------
        # finding whether the arrow being found -- for last page use only
        # try:
        #     if driver.find_element_by_xpath(
        #             "/html/body/form/div[2]/div[2]/div[2]/div/div/div[2]/div/div/div/div/div[2]/div[2]/div[6]/div[2]/ul/li[3]/button"):
        #         driver.find_element_by_xpath(
        #             "/html/body/form/div[2]/div[2]/div[2]/div/div/div[2]/div/div/div/div/div[2]/div[2]/div[6]/div[2]/ul/li[3]/button").click()
        #         still_have_next_page = True
        # except:
        #     print("Not Next Button being Found! Program End Now")
        #
        #     return False, this_page
        # else:
        #     return still_have_next_page, this_page
        # put need to put end of the file
        if len(rows) - 1 == ctr:
            return True, this_page
        return False, this_page
