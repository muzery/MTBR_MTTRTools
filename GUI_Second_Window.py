from tkinter import Tk, RIGHT, BOTH, RAISED, X, N, LEFT, Text, IntVar, StringVar, BOTTOM, W, Listbox, END, Y, TOP, Menu, \
    DISABLED, NORMAL, Toplevel, CENTER
from tkinter.ttk import Frame, Button, Style, Label, Entry, Checkbutton, LabelFrame, Scrollbar, Radiobutton, \
    Progressbar, \
    OptionMenu
import Graphical_Representation
import pandas as pd
import MI_File_CleanUp

file_name = r"C:\Users\willlee\Desktop\DataSet\Try_Party_Modified.xlsx"


class GUI(Frame):

    def get_data(self, file_name):
        # This part should be at Graphical_Representation
        self.file_name = file_name
        df1 = pd.read_excel(self.file_name)
        df3 = df1.groupby('Location')
        self.unique_part_number = df3['TE PN'].unique()
        print(self.unique_part_number)
        # Construct list

    def __init__(self, root, file_name):
        self.newWindow = Toplevel(root)
        self.get_data(file_name)
        self.widget_construct()

    def widget_construct(self):
        self.newWindow.title("Graph Plot")
        self.newWindow.resizable(0, 0)
        self.newWindow.geometry("300x180")
        # ----------Create frame for the window-----------------------#
        # -----------Frame 0 for Main Label---------------------------#
        frame0 = Frame(self.newWindow)
        frame0.pack(fill=X)
        lbl0 = Label(frame0, text="Plot Graph: ", width=30)  # <--- increase size
        lbl0.pack(side=TOP, padx=5, pady=5)
        # ----------Frame 1 for option Menu and Label-----------------#
        frame1 = Frame(self.newWindow)
        frame1.pack(fill=X)
        variable = StringVar(self.newWindow)
        lbl1 = Label(frame1, text="BU Location ", width=12)
        lbl1.pack(side=LEFT, padx=5, pady=5)
        w = OptionMenu(frame1, variable, *self.unique_part_number.index)
        w.pack(side=LEFT, padx=5, pady=5)
        variable.set(self.unique_part_number.index[0])
        # ----------- Radio Button -----------------------------------#
        frame2 = Frame(self.newWindow)
        frame2.pack(fill=BOTH, expand=False)

        lbl12 = Label(frame2, text="Type Plotting", width=100)
        lbl12.pack(side=LEFT, anchor=N, padx=5, pady=5)

        frame3 = Frame(self.newWindow)
        # frame1 = tk.Frame(root, width=100, height=100, background="bisque")
        frame3.pack(fill=BOTH, expand=False)

        v = StringVar()
        rdo3 = Radiobutton(frame3, text='HeatMap (All)', variable=v, value='1')
        rdo3.pack(side=LEFT, anchor=N, padx=5, pady=5)
        v.set('1')

        rdo4 = Radiobutton(frame3, text='Bar Plot - By Product', variable=v, value='2')
        rdo4.pack(side=LEFT, anchor=W, padx=5, pady=5)

        # ----------- Button -----------------------------------#
        frame4 = Frame(self.newWindow, relief=RAISED, borderwidth=1)
        frame4.pack(fill=BOTH, expand=True)

        plot_button = Button(self.newWindow, text="Plot Graph", command=lambda:graph_plotting(self.file_name, v,variable))
        plot_button.pack(side=RIGHT, padx=5, pady=5)


def graph_plotting(file_name, selection,opt_menu_variable):
    # Handle Which kind of file to handle
    # directly get the temp file
    graph_plot = Graphical_Representation.Graph_Plot()
    if selection.get() == '1':
        file_filter_to_save = MI_File_CleanUp.create_new_file(file_name, 'templatefile.txt')
        graph_plot.read_table(file_filter_to_save)
        graph_plot.pivot_table_assign(index_value='Location', values='Number_Of_MI_Occurance', column='TE PN')
        graph_plot.heatmap_plotting()
    else:
        bu_location = opt_menu_variable.get()
        graph_plot.read_excel(file_name)
        graph_plot.get_location_element(bu_location)



if __name__ == "__main__":
    root = Tk()
    gui = GUI(root, file_name)
    root.mainloop()
