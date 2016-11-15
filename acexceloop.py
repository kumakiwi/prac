import xlrd
#class define
class acexcel:
    def __init__(self, startstr, endstr,txtname):
        self.startstr = startstr
        self.endstr = endstr
        self.txtname = txtname
    def excel_to_txt(self,colume_list, startstr, endstr, txtname):
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
        print(result)# to make sure the output value is right
        fo = open(txtname, "w")
        for j in range(end_position_num-start_position_num+1):
            fo.write(result[j])
            fo.write("\n")
        fo.close()

data = xlrd.open_workbook('ACParameterList_integrated.xlsx')
table = data.sheets()[0]
colume_0 = []
colume_A = table.col_values(0)
#areo.ini
aero = acexcel("[WING_0]", "ANGLE","areo.ini")
aero.excel_to_txt(colume_A, aero.startstr, aero.endstr,aero.txtname)
#brake.ini
brake = acexcel("MAX_TORQUE", "ADJUST_STEP", "brake.ini")
brake.excel_to_txt(colume_A, brake.startstr, brake.endstr,brake.txtname)
#car.ini
car = acexcel("[INFO]", "SUSP_REPAIR_TIME_SEC", "car.ini")
car.excel_to_txt(colume_A, car.startstr, car.endstr,car.txtname)
#Drivetrain.ini
drivetrain = acexcel("[TRACTION]", "RPM_WINDOW_K", "drivetrain.ini")
drivetrain.excel_to_txt(colume_A, drivetrain.startstr, drivetrain.endstr,drivetrain.txtname)
#electronics.ini
electronics = acexcel("[ABS]", "DEAD_ZONE_COAST", "electronics.ini")
electronics.excel_to_txt(colume_A, electronics.startstr, electronics.endstr,electronics.txtname)
#engine.ini
engine = acexcel("[HEADER]", "DEFAULT_TURBO_ADJUSTMENT", "engine.ini")
engine.excel_to_txt(colume_A, engine.startstr, engine.endstr,engine.txtname)
#suspensions.ini
suspension = acexcel("[BASIC]", "DEBUG_LOG", "suspension.ini")
suspension.excel_to_txt(colume_A, suspension.startstr, suspension.endstr,suspension.txtname)
#tyres.ini
tyres = acexcel("[COMPOUND_DEFAULT]", "BLISTER_GAIN" , "tyres.ini")
tyres.excel_to_txt(colume_A, tyres.startstr, tyres.endstr,tyres.txtname)
