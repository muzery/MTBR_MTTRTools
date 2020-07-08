from tkinter import Tk, RIGHT, BOTH, RAISED, X, N, LEFT, Text, IntVar, StringVar, BOTTOM, W, Listbox, END, Y, TOP, Menu, \
    DISABLED, NORMAL
from tkinter.ttk import Frame, Button, Style, Label, Entry, Checkbutton, LabelFrame, Scrollbar, Radiobutton, Progressbar
import re
import os
from multiprocessing import Process, Manager, Queue
from queue import Empty
import MI_File_CleanUp
import OOP_Sellnium
import webpage_extraction
import Facebook_login
import Excel_Automation

DELAY1 = 80
DELAY2 = 20

# Queue must be global
q = Queue()


class ListCheck:

    def __int__(self):
        pass

    def date_month(self, listInput):
        # match for pattern only
        start_date_match = False
        end_date_match = False
        pattern = r"(\d{1,2})-([A-Z]{3})-(\d{4})"
        start_date_match = re.match(pattern, listInput['Start_Date'])
        end_date_match = re.match(pattern, listInput['End_Date'])
        if start_date_match is None and end_date_match is None:
            return False
        else:
            return True

    def list_checked(self, listInput, listbox):
        print("I am in List Checked")
        # 0 - UserName
        # 1 - password
        # 2 - Start_Date
        # 3 - End_Date
        # 4 - Status
        # 5 - TemplateFile
        # 6 - PartStation
        # 7 - SourceFile
        # 8 - save_directory
        # 9 - WebScrapping
        # 10 - Correction1
        # 11 - Correction2
        # 12 - Correction3
        # 13 -Browser
        file_is_good1 = False
        directory_found = False
        correction_function_template = False
        correction_function = False
        all_setting = False
        compile_function = False
        web_scrapping_setting = False
        for i in listInput:
            print(i)
        if listInput['UserName'] and listInput['Password'] and listInput['Start_Date'] and listInput['End_Date'] and \
                listInput['Status'] and listInput['WebScrapping'] and listInput['SourceFile']:
            # web scrapping able to run
            try:
                filename2, file_ext2 = os.path.splitext(listInput['SourceFile'])
            except:
                print("Setting not complete, web scrapping are unable to run")
                listbox.insert(END, 'Setting not complete, web scrapping are unable to run')

            if file_ext2.strip() == '.csv':
                file_is_good1 = True

            info = self.date_month(listInput)
            if info and file_is_good1:
                web_scrapping_setting = True
            else:
                web_scrapping_setting = False
                print("Date is not correct. Please check. Web scrapping are unable to run ")
                listbox.insert(END, 'Date is not correct. Please check. Web scrapping are unable to run ')
            # listInput [2] and [3] is Date... regression == status should be 1
        else:
            print("Setting not complete, web scrapping are unable to run")
            # listbox.insert(END, 'Setting not complete, web scrapping are unable to run')

        if True:
            # if listInput['TemplateFile'] and listInput['PartStation'] and listInput['SourceFile']and  listInput['CompileType'] \
            #        and listInput['save_directory'] and listInput['xls_file_with_correction']:
            # able to compile
            try:
                filename_template_file, file_ext_template_file = os.path.splitext(
                    listInput['TemplateFile'])  # MTTR Template file
                filename_part_station, file_ext_part_station = os.path.splitext(
                    listInput['PartStation'])  # Part Station file
                filename_source_file, file_ext_source_file = os.path.splitext(listInput['SourceFile'])  # source file
                filename_xls_with_correction, file_ext_xls_with_correction = os.path.splitext(
                    listInput['xls_file_with_correction'])  # Excel Pfile save path

                if file_ext_source_file.strip() == '.csv' and \
                        file_ext_part_station.strip() == '.xlsx' and file_ext_xls_with_correction.strip() == '.xlsx':

                    if os.path.exists(listInput['SourceFile'].strip()) and os.path.exists(
                            listInput['PartStation'].strip()) or os.path.exists(
                        listInput['xls_file_with_correction'].strip()):
                        compile_function = True

                    else:
                        print("File (csv) cannot found! Please Check")
                        listbox.insert(END, 'File (csv) cannot found! Please Check')

                # write together with template file
                if file_ext_source_file.strip() == '.csv' and file_ext_template_file.strip() == '.xlsx' and \
                        file_ext_template_file.strip() == '.xlsx' and file_ext_xls_with_correction.strip() == '.xlsx':

                    if os.path.exists(listInput['SourceFile'].strip()) and os.path.exists(
                            listInput['TemplateFile'].strip()) and \
                            os.path.exists(listInput['xls_file_with_correction'].strip()):
                        correction_function_template = True
                    else:
                        print("File (csv) cannot found! Please Check")
                        listbox.insert(END, 'File (csv) cannot found! Please Check')

                if os.path.isdir(listInput['save_directory'].strip()):
                    directory_found = True
                else:
                    # create new directory
                    os.mkdir(listInput['save_directory'].strip())
            except:
                print("Setting not complete, compile feature are unable to run")
                listbox.insert(END, 'Setting not complete, compile feature are unable to run')
        else:
            print("Setting not complete, compile feature are unable to run")
            listbox.insert(END, 'Setting not complete, compile feature are unable to run')
            listbox.yview(END)

        return web_scrapping_setting, compile_function, correction_function, correction_function_template

    def create_new_file(self, file_save):
        filepath, file_name = os.path.split(file_save)
        return filepath + "\\" + "templatefile.txt"


class Example(Frame):

    def onGetValue_1(self):

        if self.p1.is_alive():

            self.after(DELAY1, self.onGetValue)
            return

        else:

            try:
                self.mylist.insert('end', q.get(0))
                self.mylist.insert('end', "\n")
                self.pbar.stop()
                self.okButton.config(state=NORMAL)

            except Empty:
                print("queue is empty")

    def onGetValue(self):

        if self.p1.is_alive():

            self.after(DELAY1, self.onGetValue)
            return

        else:

            try:
                self.mylist.insert('end', q.get(0))
                self.mylist.insert('end', "\n")
                self.pbar.stop()
                self.okButton.config(state=NORMAL)

            except Empty:
                print("queue is empty")

    def entry_box_apperance(self, *args):
        pass

    def client_exit(self):
        exit()

    def status_selection(self, list_user_input):
        if list_user_input['Status'] == 1:
            list_user_input['Status'] = "Closed"
        elif list_user_input['Status'] == 2:
            list_user_input['Status'] = "Waiting for 3rd Party"
        else:
            list_user_input['Status'] = "Nothing"

        return list_user_input

    def browser_selection(self, list_user_input):
        if list_user_input['Browser'] == 1:
            list_user_input['Browser'] = "firefox"
        elif list_user_input['Browser'] == 2:
            list_user_input['Browser'] = "chrome"
        else:
            list_user_input['Browser'] = "ie"

        return list_user_input

    def run(self, *args):

        # self.okButton.config(state = DISABLED)
        # store everything inside list
        key = ['UserName', 'Password', 'Start_Date', 'End_Date', 'Status', 'SourceFile', 'TemplateFile',
               'PartStation', 'xls_file_with_correction', 'save_directory', 'WebScrapping', 'CompileType', 'Browser']
        # self.okButton.config(state=DISABLED)

        list_user_input = {}
        counter = 0
        if isinstance(self.mylist, Listbox):
            for i in range(0, len(args)):
                print(i)
                # Check without is empty and the iteration is fix
                list_user_input[key[i]] = args[i].get()
                self.mylist.insert(END, args[i].get())

            list_user_input = self.status_selection(list_user_input)
            list_user_input = self.browser_selection(list_user_input)

            web_scrapping_setting, compile_function, correction_function, correction_function_template = self.lc.list_checked(
                list_user_input, self.mylist)
            # Caling Facebook_login.py
            if web_scrapping_setting:
                # Separate process (long process)
                # Facebook_login.automate(list_user_input['UserName'],list_user_input['Password'],list_user_input['Status'],
                #                         list_user_input['Start_Date'],list_user_input['End_Date'], list_user_input['SourceFile'].strip(), self.mylist)
                self.p1 = Process(target=data_crawling, args=[q, list_user_input, Listbox])
                self.p1.start()
                self.pbar.start(DELAY2)
                self.after(DELAY1, self.onGetValue)

            if compile_function and list_user_input['CompileType'] == 1:
                # MI_File_CleanUp.main(list_user_input['PartStation'].strip(), list_user_input['CompileType'],
                #                      list_user_input['SourceFile'].strip(),
                #                      list_user_input['xls_file_with_correction'].strip())

                self.p1 = Process(target=func_run, args=(q, list_user_input))
                self.p1.start()
                self.pbar.start(DELAY2)
                self.after(DELAY1, self.onGetValue)

                self.mylist.insert(END,
                                   'Check the file before proceed to do the calculation. Please manual do the correction')
                self.mylist.yview(END)
                self.mylist.insert(END, 'Done')
                self.mylist.yview(END)

            if compile_function and list_user_input['CompileType'] == 2:
                self.p1 = Process(target=func_run, args=(q, list_user_input))
                self.p1.start()
                self.pbar.start(DELAY2)
                self.after(DELAY1, self.onGetValue)
                # MI_File_CleanUp.main(list_user_input['PartStation'].strip(), list_user_input['CompileType'],
                #                      list_user_input['SourceFile'].strip(),
                #                      list_user_input['xls_file_with_correction'].strip())
                self.mylist.insert(END, 'Done')
                self.mylist.yview(END)

            if correction_function_template and list_user_input['CompileType'] == 3:
                template_file = self.lc.create_new_file(list_user_input['xls_file_with_correction'].strip())
                self.p1 = Process(target=func_run_automation, args=(q, list_user_input, template_file))
                self.p1.start()
                self.pbar.start(DELAY2)
                self.after(DELAY1, self.onGetValue)
                self.mylist.insert(END, 'Done')
                self.mylist.yview(END)

        else:
            print("Code having problem. Please contact William Lee")
        # self.okButton.config(state=NORMAL)

    def __init__(self):
        super().__init__()
        lc = ListCheck()
        self.lc = lc
        self.initUI()

    def initUI(self):

        self.master.title("MTTR Automation Tool")
        self.pack(fill=BOTH, expand=False)

        # -------------menu---------------------------#

        # menu = Menu(self.master)
        # self.master.config (menu)
        # filemenu = Menu(menu)
        # menu.add_cascade(label='File', menu=filemenu)
        # filemenu.add_command(label='New')
        # filemenu.add_command(label='Open...')
        # filemenu.add_separator()
        # filemenu.add_command(label='Exit', command=self.quit)
        # helpmenu = Menu(menu)
        # menu.add_cascade(label='Help', menu=helpmenu)
        # helpmenu.add_command(label='About')

        # -------------user name ---------------------------#
        # StringVar() is the variable class
        # we create an instance of this class
        username = StringVar()

        frame1 = Frame(self)
        frame1.pack(fill=X)

        lbl1 = Label(frame1, text="User Name : ", width=30)
        lbl1.pack(side=LEFT, padx=5, pady=5)

        entry1 = Entry(frame1, textvariable=username)
        entry1.pack(fill=X, padx=5, expand=False)
        # -------------Password ---------------------------#

        password = StringVar()

        frame2 = Frame(self)
        frame2.pack(fill=X)

        lbl2 = Label(frame2, text="Password :", width=30)
        lbl2.pack(side=LEFT, padx=5, pady=5)

        entry2 = Entry(frame2, textvariable=password)
        entry2.pack(fill=X, padx=5, expand=False)
        entry2.config(show="*")

        # -----------Start Date--------------------------- #

        start_date = StringVar()

        frame17 = Frame(self)
        frame17.pack(fill=X)

        lbl8 = Label(frame17, text="Start Date (ONLY) DD-MMM-YYYY :", width=30)
        lbl8.pack(side=LEFT, padx=5, pady=5)

        entry6 = Entry(frame17, textvariable=start_date)
        entry6.pack(fill=X, padx=5, expand=False)

        # -----------End Date--------------------------- #

        end_date = StringVar()

        frame18 = Frame(self)
        frame18.pack(fill=X)

        lbl9 = Label(frame18, text="End Date (ONLY) DD-MMM-YYYY :", width=30)
        lbl9.pack(side=LEFT, padx=5, pady=5)

        entry7 = Entry(frame18, textvariable=end_date)
        entry7.pack(fill=X, padx=5, expand=False)

        # -------------Status---------------------------#
        frame11 = Frame(self)
        frame11.pack(fill=BOTH, expand=False)

        lbl10 = Label(frame11, text="Status", width=100)
        lbl10.pack(side=LEFT, anchor=N, padx=5, pady=5)

        #
        frame19 = Frame(self)
        # frame1 = tk.Frame(root, width=100, height=100, background="bisque")
        frame19.pack(fill=BOTH, expand=False)

        status = IntVar()
        rdo3 = Radiobutton(frame19, text='Closed', variable=status, value=1)
        rdo3.pack(side=LEFT, anchor=N, padx=5, pady=5)
        status.set(1)
        # chk1.pack(side=BOTTOM)

        rdo4 = Radiobutton(frame19, text='Waiting for 3rd Party', variable=status, value=2)
        rdo4.pack(side=LEFT, anchor=W, padx=5, pady=5)

        # -------------Source File---------------------------#
        source_file = StringVar()

        frame5 = Frame(self)
        frame5.pack(fill=BOTH, expand=False)

        lbl5 = Label(frame5, text="Source File(.csv)/File to Save (for webscrapping)(.csv)", width=30)
        lbl5.pack(side=LEFT, anchor=N, padx=5, pady=5)

        entry5 = Entry(frame5, textvariable=source_file)
        entry5.pack(fill=BOTH, pady=5, padx=5, expand=False)

        # -------------MTTR Template File ---------------------------#
        template_file_path = StringVar()

        frame3 = Frame(self)
        frame3.pack(fill=BOTH, expand=False)

        lbl3 = Label(frame3, text="MTTR Template File(.xlsx)", width=30)
        lbl3.pack(side=LEFT, anchor=N, padx=5, pady=5)

        entry3 = Entry(frame3, textvariable=template_file_path)
        entry3.pack(fill=BOTH, pady=5, padx=5, expand=False)

        # -------------Part Station File---------------------------#
        part_station_file = StringVar()

        frame4 = Frame(self)
        frame4.pack(fill=BOTH, expand=False)

        lbl4 = Label(frame4, text="Part Station File(.xlsx)", width=30)
        lbl4.pack(side=LEFT, anchor=N, padx=5, pady=5)

        entry4 = Entry(frame4, textvariable=part_station_file)
        entry4.pack(fill=BOTH, pady=5, padx=5, expand=False)

        # -------------Save File Directory---------------------------#
        xls_file_with_correction = StringVar()

        frame13 = Frame(self)
        frame13.pack(fill=BOTH, expand=False)

        lbl11 = Label(frame13, text="Excel File Save Path(.xlsx)", width=30)
        lbl11.pack(side=LEFT, anchor=N, padx=5, pady=5)

        entry7 = Entry(frame13, textvariable=xls_file_with_correction)
        entry7.pack(fill=BOTH, pady=5, padx=5, expand=False)

        # -------------Save File Directory---------------------------#
        save_directory = StringVar()

        frame20 = Frame(self)
        frame20.pack(fill=BOTH, expand=False)

        lbl5 = Label(frame20, text="Directory to save", width=30)
        lbl5.pack(side=LEFT, anchor=N, padx=5, pady=5)

        entry6 = Entry(frame20, textvariable=save_directory)
        entry6.pack(fill=BOTH, pady=5, padx=5, expand=False)

        # -------------Operation Selection ---------------------------#
        frame6 = Frame(self)
        frame6.pack(fill=BOTH, expand=False)

        lbl6 = Label(frame6, text="Select the operation that wish to run:", width=100)
        lbl6.pack(side=LEFT, anchor=N, padx=5, pady=5)
        #
        frame7 = Frame(self)
        # frame1 = tk.Frame(root, width=100, height=100, background="bisque")
        frame7.pack(fill=BOTH, expand=False)

        var1 = IntVar()
        chk1 = Checkbutton(frame7, text='Web Scrapping (firefox only)', variable=var1)
        chk1.pack(side=LEFT, anchor=N, padx=5, pady=5)
        # chk1.pack(side=BOTTOM)

        frame8 = Frame(self)
        # frame1 = tk.Frame(root, width=100, height=100, background="bisque")
        frame8.pack(fill=BOTH, expand=False)

        var2 = IntVar()
        rdo4 = Radiobutton(frame8, text='Correction Without Calculating', variable=var2, value=1)
        rdo4.pack(side=LEFT, anchor=N, padx=5, pady=5)
        var2.set(1)

        rdo5 = Radiobutton(frame8, text='Correction With Calculating and export to txt file', variable=var2, value=2,
                           command="")
        rdo5.pack(side=LEFT, anchor=W, padx=5, pady=5)

        frame9 = Frame(self)
        # frame1 = tk.Frame(root, width=100, height=100, background="bisque")
        frame9.pack(fill=BOTH, expand=False)

        rdo6 = Radiobutton(frame9,
                           text='Correction With Calculating and export to excel file according to template file format',
                           variable=var2, value=3, command="")
        rdo6.pack(side=BOTTOM, anchor=W, padx=5, pady=5)

        # -------------Web Browser---------------------------#
        frame11 = Frame(self)
        frame11.pack(fill=BOTH, expand=False)

        lbl7 = Label(frame11, text="Select which web Browser to use :(use for web scrapping only)", width=100)
        lbl7.pack(side=LEFT, anchor=N, padx=5, pady=5)

        #
        frame12 = Frame(self)
        # frame1 = tk.Frame(root, width=100, height=100, background="bisque")
        frame12.pack(fill=BOTH, expand=False)

        v = IntVar()
        rdo1 = Radiobutton(frame12, text='Firefox', variable=v, value=1)
        rdo1.pack(side=LEFT, anchor=N, padx=5, pady=5)
        v.set(1)
        # chk1.pack(side=BOTTOM)

        rdo2 = Radiobutton(frame12, text='Chrome', variable=v, value=2)
        rdo2.pack(side=LEFT, anchor=W, padx=5, pady=5)

        rdo3 = Radiobutton(frame12, text='IE', variable=v, value=3)
        rdo3.pack(side=LEFT, anchor=W, padx=5, pady=5)

        # -------------Log Message ---------------------------#
        frame15 = Frame(self)
        frame15.pack(fill=BOTH, expand=False)

        lbl8 = Label(frame15, text="Message Log:", width=100)
        lbl8.pack(side=LEFT, anchor=N, padx=5, pady=5)

        frame16 = Frame(self)
        # frame1 = tk.Frame(root, width=100, height=100, background="bisque")
        frame16.pack(fill=BOTH, expand=False)

        scrollbar = Scrollbar(frame16)
        scrollbar.pack(side=RIGHT, anchor=N, fill=Y, padx=5, pady=5)

        # Example
        self.mylist = Listbox(frame16, yscrollcommand=scrollbar.set)

        # for line in range(200):
        #     mylist.insert(END, "I am EMPTY, PLease don't surprise" )
        self.mylist.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.config(command=self.mylist.yview)

        # self.master.title("Buttons")
        # self.style = Style()
        # self.style.theme_use("default")
        #
        frame = Frame(self, relief=RAISED, borderwidth=1)
        frame.pack(fill=BOTH, expand=True)
        #
        self.pack(fill=BOTH, expand=True)
        #
        closeButton = Button(self, text="Close/Quit", command=self.client_exit)
        closeButton.pack(side=RIGHT, padx=5, pady=5)
        self.okButton = Button(self, text="OK to Run",
                               command=lambda: self.run(username, password, start_date, end_date, status,
                                                        source_file, template_file_path, part_station_file,
                                                        xls_file_with_correction, save_directory,
                                                        var1, var2, v))
        self.okButton.pack(side=RIGHT)

        # -------------Progress Bar ---------------------------#

        frame16 = Frame(self)
        frame16.pack(fill=BOTH, expand=False)

        self.pbar = Progressbar(frame16, length=300, mode='indeterminate')
        self.pbar.pack(side=LEFT, anchor=W, padx=5, pady=5)


# ----------------Generate function must be a top-level module function---------------------------------------

def data_crawling(q, list_user_input, listbox):
    # Separate process (long process)
    OOP_Sellnium.main(list_user_input['UserName'], list_user_input['Password'], list_user_input['Status'],
                      list_user_input['Start_Date'], list_user_input['End_Date'],
                      list_user_input['SourceFile'],list_user_input['Browser'])
    # Facebook_login.automate(list_user_input['UserName'], list_user_input['Password'], list_user_input['Status'],
    #                             list_user_input['Start_Date'], list_user_input['End_Date'],
    #                             list_user_input['SourceFile'], listbox)
    q.put('Done')


def func_run(q, list_user_input):
    MI_File_CleanUp.main(list_user_input['PartStation'].strip(), list_user_input['CompileType'],
                         list_user_input['SourceFile'].strip(),
                         list_user_input['xls_file_with_correction'].strip())

    q.put('Done')


def func_run_automation(q, list_user_input, template_file):
    MI_File_CleanUp.main(list_user_input['PartStation'].strip(), 2, list_user_input['SourceFile'].strip(),
                         list_user_input['xls_file_with_correction'].strip())

    # here should be .txt file
    Excel_Automation.open_template_file(list_user_input['TemplateFile'].strip(), list_user_input['PartStation'].strip(),
                                        template_file,
                                        list_user_input['Start_Date'].strip(), list_user_input['End_Date'].strip())
    q.put('Done')


def main():
    root = Tk()
    root.geometry("650x750")  # <-Original "650x750
    root.resizable(0, 0)
    app = Example()
    root.mainloop()


if __name__ == '__main__':
    main()
