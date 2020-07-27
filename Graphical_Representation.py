import seaborn as sns
import pandas as pd
import matplotlib.pyplot as plt


# file_name_2 = r"C:\Users\willlee\Desktop\DataSet\Try_Party_Modified.xlsx"
# df2 = pd.read_excel(file_name_2)
# df2
# df3 = df2.groupby('Location')
# a = df3['TE PN'].unique()
# print(a)
# df2.loc[df2['TE PN'] =='540648A-00']
# df2.loc[df2['TE PN'] =='540648A-00'].count()
# df2.loc[df2['TE PN'] =='540648A-00']['P7_ROOT_CAUSE_SUBTYPE_ID'].unique()
# a = df2.loc[df2['TE PN'] =='540648A-00']
# a
# a.loc[a['P7_ROOT_CAUSE_SUBTYPE_ID'] == 'Connection;']
# df3 = df2.groupby('Location')
# a = df3['TE PN'].unique()
# df2.loc[df2['TE PN'] =='540648A-00'].count()
# create heatmap (all product) and barplot (individual product)
# a.loc[a['P7_ROOT_CAUSE_SUBTYPE_ID'] == 'Connection;']['# Issue Number'].count()
# ax = sns.countplot(x = 'P7_ROOT_CAUSE_SUBTYPE_ID',data = a)
# g = sns.catplot(x = 'P7_ROOT_CAUSE_SUBTYPE_ID', data = b, kind ='count', col ='TE PN')
# plt.show()

class Graph_Plot:
    def __init__(self):
        pass

    def read_table(self,file_name):
        self.df = pd.read_table(file_name)

    def pivot_table_assign(self, index_value, values, column):
        self.df1 = pd.pivot_table(self.df, index=index_value, values=values, columns=column)

    def heatmap_plotting(self):
        fig, ax = plt.subplots(1)
        plot1 = sns.heatmap(self.df1, cmap='coolwarm', annot=True, fmt=".0f",cbar = False,ax = ax,annot_kws={'size':16})
        plt.xlabel('Location')
        plt.ylabel('Station Part Number')
        ax.set_ylim((0, 10))
        plt.text(5, 10.0, "MI Overall Count", fontsize=25, color='Black', fontstyle='italic')
        plt.show()

    @property
    def file_name_op(self):
        return self.__file_name

    # @file_name_op.setter
    # def file_name_op(self, file_name):
    #     self.file_name = file_name
    def read_excel(self,file_name):
        self.df2 = pd.read_excel(file_name)

    def get_location_element(self,location_search):
        a = self.df2[self.df2['Location'] == location_search]
        #look for is there empty
        if a['# Issue Number'].count() == 0:
            return
        else:
            self.__barplot_plotting__(a)
            #unique_part_number = a['TE PN'].unique()

    def __barplot_plotting__(self,a):
        g = sns.catplot(x='P7_ROOT_CAUSE_SUBTYPE_ID', data=a, kind='count', col='TE PN')
        plt.show()
