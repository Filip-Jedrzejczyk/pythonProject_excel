from pandas import DataFrame
import pandas as pd


filename1 = r'Book1.xlsx'
filename2 = r'Book2.xlsx'
filename3 = r'Book3.xlsx'
filename4 = r'Book4.xlsx'
filename5 = r'Book5.xlsx'
df1 = pd.read_excel(filename1)
df2 = pd.read_excel(filename2)
df3 = pd.read_excel(filename3)
df4 = pd.read_excel(filename4)
df5 = pd.read_excel(filename5)
list1 = list(df1['MMITNO'])
list2 = list(df2['MMITNO'])
list3 = list(df3['MMITNO'])
list4 = list(df4['MMITNO'])
list5 = list(df5['MMITNO'])

MMITNO_list = list(df1['MMITNO']) + list(df2['MMITNO']) + list(df3['MMITNO']) + list(df4['MMITNO']) + list(df5['MMITNO'])
MMTIDS_list = list(df1['MMTIDS']) + list(df2['MMTIDS']) + list(df3['MMTIDS']) + list(df4['MMTIDS']) + list(df5['MMTIDS'])
MMFUDS_list = list(df1['MMFUDS']) + list(df2['MMFUDS']) + list(df3['MMFUDS']) + list(df4['MMFUDS']) + list(df5['MMFUDS'])
MMITGR_list = list(df1['MMITGR']) + list(df2['MMITGR']) + list(df3['MMITGR']) + list(df4['MMITGR']) + list(df5['MMITGR'])
MMITCL_list = list(df1['MMITCL']) + list(df2['MMITCL']) + list(df3['MMITCL']) + list(df4['MMITCL']) + list(df5['MMITCL'])
MMITTY_list = list(df1['MMITTY']) + list(df2['MMITTY']) + list(df3['MMITTY']) + list(df4['MMITTY']) + list(df5['MMITTY'])
MMNEWE_list = list(df1['MMNEWE']) + list(df2['MMNEWE']) + list(df3['MMNEWE']) + list(df4['MMNEWE']) + list(df5['MMNEWE'])
MMGRWE_list = list(df1['MMGRWE']) + list(df2['MMGRWE']) + list(df3['MMGRWE']) + list(df4['MMGRWE']) + list(df5['MMGRWE'])
MMPRGP_list = list(df1['MMPRGP']) + list(df2['MMPRGP']) + list(df3['MMPRGP']) + list(df4['MMPRGP']) + list(df5['MMPRGP'])
MMCONO_list = list(df1['MMCONO']) + list(df2['MMCONO']) + list(df3['MMCONO']) + list(df4['MMCONO']) + list(df5['MMCONO'])


class excel_table:
    def __init__(self, MMITNO, MMTIDS, MMFUDS, MMITGR, MMITCL, MMITTY, MMNEWE, MMGRWE, MMPRGP, MMCONO):
        self.MMITNO = MMITNO
        self.MMTIDS = MMTIDS
        self.MMFUDS = MMFUDS
        self.MMITGR = MMITGR
        self.MMITCL = MMITCL
        self.MMITTY = MMITTY
        self.MMNEWE = MMNEWE
        self.MMGRWE = MMGRWE
        self.MMPRGP = MMPRGP
        self.MMCONO = MMCONO

    def __eq__(self, other):
        return self.MMITNO == other.MMITNO and self.MMTIDS == other.MMTIDS and self.MMFUDS == other.MMFUDS and self.MMITGR == other.MMITGR and self.MMITCL == other.MMITCL and self.MMITTY == other.MMITTY and self.MMNEWE == other.MMNEWE and self.MMGRWE == other.MMGRWE and self.MMPRGP == other.MMPRGP and self.MMCONO == other.MMCONO

    def __hash__(self):
        return hash((self.MMITNO, self.MMTIDS, self.MMFUDS, self.MMITGR, self.MMITCL, self.MMITTY, self.MMNEWE, self.MMGRWE, self.MMPRGP, self.MMCONO))

    def __str__(self):
        return str(vars(self))

    def show(self):
        print(str(self.MMITNO) + "   " + str(self.MMTIDS) + "   "
              + str(self.MMFUDS) + "   " + str(self.MMITGR) + "   "
              + str(self.MMITCL) + "   " + str(self.MMITTY) + "   "
              + str(self.MMNEWE) + "   " + str(self.MMGRWE) + "   "
              + str(self.MMPRGP) + "   " + str(self.MMCONO) + "\n")


first_list = []
for i in range(len(MMITNO_list)):
    first_list.append(excel_table(MMITNO_list[i], MMTIDS_list[i], MMFUDS_list[i], MMITGR_list[i], MMITCL_list[i], MMITTY_list[i], MMNEWE_list[i], MMGRWE_list[i], MMPRGP_list[i], MMCONO_list[i]))

second_list = []


def checkio(data):
    a = [i for i in data if data.count(i) > 1]
    return a


temp_list_ID = []
for i in range(len(first_list)):
    temp_list_ID.append(first_list[i].MMITNO)

temp_list_ID = checkio(temp_list_ID)
temp_list_ID = list(set(temp_list_ID))
print(temp_list_ID)

for i in range(len(first_list)):
    for j in range(len(temp_list_ID)):
        if first_list[i].MMITNO == temp_list_ID[j]:
            second_list.append(first_list[i])


second_list_temp = set(second_list)
for i in second_list_temp:
    print(i)

second_list_temp = list(set(second_list))
second_list_temp.sort(key=lambda x: x.MMITNO, reverse=True)
for i in range(len(second_list_temp)):
    second_list_temp[i].show()

