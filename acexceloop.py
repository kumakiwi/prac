import xlrd
class acexcel:
    def __init__(self, startstr, endstr,txtname):
        self.startstr = startstr
        self.endstr = endstr
        self.txtname = txtname
    def find_position(self,colume_list, startstr, endstr, txtname):
        start_position_num = colume_list.index(startstr)
        end_position_num = colume_list.index(endstr)
        #tmp_l=colume_A[start_position_num:end_position_num+1]

        for i in range(start_position_num,end_position_num+1):
            cell_temp = table.cell(i,4).value
            print(str(colume_A[i])+"="+str(cell_temp))
        print ("hello,\n")
        fo = open("foo.txt", "wb")
        result = []
        for i in range(start_position_num,end_position_num+1):
            cell_temp = table.cell(i,4).value
            result.append(str(colume_A[i])+"="+str(cell_temp))
        print(result)
        fo = open(txtname, "w")
        for j in range(end_position_num-start_position_num+1):
            fo.write(result[j])
            fo.write("\n")
        fo.close()
        '''
        for j in range(end_position_num-start_position_num):
            fo = open("foo1.txt", "a+")
            fo.write( result[i])
            print("\n")
            fo.close()
        '''
data = xlrd.open_workbook('ACParameterList_integrated.xlsx')
table = data.sheets()[0]
colume_0 = []
colume_A = table.col_values(0)
aero = acexcel("[WING_0]", "ANGLE","area.ini")
aero.find_position(colume_A, aero.startstr, aero.endstr,aero.txtname)
