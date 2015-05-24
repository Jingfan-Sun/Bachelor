# -*- coding: utf-8 -*-
import os
import string

import math
# Add powerfactory.pyd path to python path.
# This is an example for 32 bit PowerFactory architecture.
import sys
sys.path.append("D:\\Program Files\\DIgSILENT\\PowerFactory 15.2\\Python\\3.4\\")
#import PowerFactory module
import powerfactory
#start PowerFactory in engine mode
app = powerfactory.GetApplication()
#run Python code below

#清屏
app.ClearOutputWindow()

# ---------------------------------------------------------------------------------------------------
def OpenFile():
    """Opens the BPA file."""
    path = 'C:\\Users\\xj\\Desktop\\13xg\\2013年稳定_华东.swi' 
    bpa_file = open(path, 'r')

    return bpa_file

# ---------------------------------------------------------------------------------------------------
def OpenBusFile():
    """Opens the BPA file."""
    path = 'C:\\Users\\xj\\Desktop\\bus.txt' 
    bus_file = open(path, 'r')

    return bus_file

# ----------------------------------------------------------------------------------------------------
def isfloat(str):
    """Checks if the string is a floating point number."""

    try:
        float(str)
        return True			#Returns true if the string is a floating point number
    except (ValueError, TypeError):
        return False			#Returns false otherwise

# ----------------------------------------------------------------------------------------------------
def isint(str):
    """Checks if the string is an integer."""

    try:
        int(str)
        return True			#Returns true if the string is an integer
    except (ValueError, TypeError):
        return False			#Returns false otherwise

# ----------------------------------------------------------------------------------------------------
def getfloatvalue(str_line,num):
    if str_line.strip() == '':
        return 0
    elif str_line.strip() == '.':
        return 0
    elif '.' in str_line:
        return float(str_line.strip())
    elif '.' not in str_line:
        #return float((str_line[0:len(str_line)-num]+'.'+str_line[len(str_line)-num:]).strip())
        return float(str_line.strip())/pow(10,num)

# ----------------------------------------------------------------------------------------------------
def GetCASECard(bpa_file, bpa_str):
    """Gets the CASECard in the BPA file."""
    
    IXO = 0          #非零，用于形成变压器零序数据卡
    IWSCC = 0  
    X2FAC = 0.65     #不对称故障时发阿电机X2/X'd之比
    XFACT = 0        #X"d/X'd之比，当考虑阻尼绕组时才需填 
    TDODPS = 0.03    #隐极机的T"do，以秒计
    TQODPS = 0.05    #隐极机的T"qo，以秒计
    TDODPH = 0.04    #凸极水轮机T"do，以秒计
    TQODPH = 0.3     #凸极水轮机T"qo，以秒计
    CFACL2 = 0.36    #负序负荷导纳标么值Y2*＝0.19+jCFACL2
    

#		 ！使用find报错
#    pos = string.find(bpa_str, "CASE")
#    if pos != -1:
#        bpa_file.seek(pos)
#        bpa_file.readline()			#Just to position on to the next line
#        line = bpa_file.readline()
    for line in bpa_str_ar:			#Loop over every line of the BPA file
        #line = line.lstrip() 
        line = line.rstrip('\n')
        if line == "": continue			#To continue if it is a blank line
        
        if line[0] == 'C' and line[1] == 'A' and line[2] == 'S' and line[3] == 'E':	#If it is an CASE card
            line = line + ' '*(80-len(line))	#To pad each line with spaces up to 34 records
            if line[3:13].strip() == area_name or line[14:24].strip() == area_name: return True		#Returns true if the area is found in an I card

        if isfloat(line[20]):
        	if float(line[20]) != 0: IXO = 1
        if isfloat(line[22]): 
        	if float(line[22]) > 0: IWSCC = 1
        if isfloat(line[44:49]): X2FAC= getfloatvalue(line[44:49],5)
        if isfloat(line[49:54]): XFACT= getfloatvalue(line[49:54],5)
        if isfloat(line[54:59]): TDODPS= getfloatvalue(line[54:59],5)
        if isfloat(line[59:64]): TQODPS= getfloatvalue(line[59:64],5)
        if isfloat(line[64:69]): TDODPH= getfloatvalue(line[64:69],5)
        if isfloat(line[69:74]): TQODPH= getfloatvalue(line[69:74],5)
        if isfloat(line[74:80]): CFACL2= getfloatvalue(line[74:80],5)
        
        #只有第一张CASE卡为计算控制卡
        break
        	
    return IXO,IWSCC,X2FAC,XFACT,TDODPS,TQODPS,TDODPH,TQODPH,CFACL2
# ----------------------------------------------------------------------------------------------------
def GetF1Card(bpa_str_ar):
    #F1卡不完善！！！！！

    DMPALL = 0
    IAMRTS = 0 #该值不为0，且CASE卡中的XFACT也不为0，采用缺省次暂态参数
    for line in bpa_str_ar:			#Loop over every line of the BPA file
        #line = line.lstrip() 
        line = line.rstrip('\n')
        if line == "": continue			#To continue if it is a blank line
        
        if line[0] == 'F' and line[1] == '1':
            line = line + ' '*(80-len(line))	#To pad each line with spaces up to 80 records

            if isfloat(line[19:22]):
                DMPALL = getfloatvalue(line[19:22],2)

            if isint(line[25]): IAMRTS = int(line[25])

    return DMPALL,IAMRTS

# ----------------------------------------------------------------------------------------------------
def GetFFCard(bpa_str_ar):
    #FF卡不完善

    DMPMLT = 1
    NOSAT = 0
    
    for line in bpa_str_ar:			#Loop over every line of the BPA file
        #line = line.lstrip() 
        line = line.rstrip('\n')
        if line == "": continue			#To continue if it is a blank line
        
        if line[0] == 'F' and line[1] == 'F':
            line = line + ' '*(80-len(line))	#To pad each line with spaces up to 80 records

            if isfloat(line[44:47]):
                DMPALL = getfloatvalue(line[44:47],3)

            if isint(line[72]): NOSAT = int(line[72])

    return DMPMLT,NOSAT
    
# ----------------------------------------------------------------------------------------------------
def GetECard(bpa_file, bpa_str_ar,bus_name_ar):
    #励磁机E卡

    prj = app.GetActiveProject()
    if prj is None:
        raise Exception("No project activated. Python Script stopped.")

    app.PrintInfo(prj)
    Network_Modell = prj.SearchObject("Network Model.IntPrjfolder")
    Network_Model=Network_Modell[0]

    Network_Dataa = Network_Model.SearchObject("Network Data.IntPrjfolder")
    Network_Data = Network_Dataa[0]
    app.PrintInfo(Network_Data)

    Nett = Network_Data.SearchObject("*.ElmNet")
    Net = Nett[0]
    app.PrintInfo(Net)

    Libraryy = prj.SearchObject("Library.IntPrjfolder")
    Library = Libraryy[0]
    app.PrintInfo(Library)

    user_defined_modell = Library.SearchObject("User Defined Models.IntPrjfolder")
    user_defined_model = user_defined_modell[0]
    app.PrintInfo(user_defined_model)

    BPA_Frame_EE = user_defined_model.SearchObject("BPA Frame(E).BlkDef")
    BPA_Frame_E = BPA_Frame_EE[0]
    app.PrintInfo(BPA_Frame_E)

    avr_EAA = user_defined_model.SearchObject("avr_EA.BlkDef")
    avr_EA = avr_EAA[0]
    app.PrintInfo(avr_EA)

    avr_EGG = user_defined_model.SearchObject("avr_EG.BlkDef")
    avr_EG = avr_EGG[0]
    app.PrintInfo(avr_EG)

    avr_EJJ = user_defined_model.SearchObject("avr_EJ.BlkDef")
    avr_EJ = avr_EJJ[0]
    app.PrintInfo(avr_EJ)

    motor_Tonee = user_defined_model.SearchObject("motor_Tone.BlkDef")
    motor_Tone = motor_Tonee[0]
    app.PrintInfo(motor_Tone)

    Compensatorr = user_defined_model.SearchObject("Compensator.BlkDef")
    Compensator = Compensatorr[0]
    app.PrintInfo(Compensator)    
    
    for line in bpa_str_ar:
        line = line.rstrip('\n')
        if line == "": continue        

        #E卡
        if line[0] == 'E':
            chinese_count = 0
            #判断中文的个数
            for i in range (3,10):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count = chinese_count + 1
            name = line[3:11-chinese_count].strip()
            base = float(line[11-chinese_count:15-chinese_count])
            #名称+电压
            Generator = name + str(base)                
            #找到节点名对应的编号
            if Generator in bus_name_ar:
                Generator_num = bus_name_ar[Generator]

                #找到对应的comp模块
                Generator_comp_name = str(Generator_num) + '_comp'
                Generator_compp = Net.SearchObject(str(Generator_comp_name))
                Generator_comp = Generator_compp[0]
    
                if not Generator_comp == None:
                    #avr.dsl
    #                app.PrintInfo(Generator_num)
                    Generator_avr_dsll = Generator_comp.SearchObject('avr_dsl.ElmDsl')
                    Generator_avr_dsl = Generator_avr_dsll[0]
                    if Generator_avr_dsl == None:
                        Generator_comp.CreateObject('ElmDsl','avr_dsl')
                        Generator_avr_dsll = Generator_comp.SearchObject('avr_dsl.ElmDsl')
                        Generator_avr_dsl = Generator_avr_dsll[0]
                    Generator_comp.Avr = Generator_avr_dsl

                    Generator_motor_dsll = Generator_comp.SearchObject('motor.ElmDsl')
                    Generator_motor_dsl = Generator_motor_dsll[0]
                    if Generator_motor_dsl == None:
                        Generator_comp.CreateObject('ElmDsl','motor')
                        Generator_motor_dsll = Generator_comp.SearchObject('motor.ElmDsl')
                        Generator_motor_dsl = Generator_motor_dsll[0]                
                    Generator_comp.Prime_Moter = Generator_motor_dsl
                    Generator_comp_motor = Generator_comp.Prime_Moter
                    Generator_comp_motor.typ_id = motor_Tone

                    Generator_Compensator_dsll = Generator_comp.SearchObject('Compensator.ElmDsl')
                    Generator_Compensator_dsl = Generator_Compensator_dsll[0]
                    if Generator_Compensator_dsl == None:
                        Generator_comp.CreateObject('ElmDsl','Compensator')
                        Generator_Compensator_dsll = Generator_comp.SearchObject('Compensator.ElmDsl')
                        Generator_Compensator_dsl = Generator_Compensator_dsll[0]                
                    Generator_comp.Compensator = Generator_Compensator_dsl
                    Generator_comp_com = Generator_comp.Compensator
                    Generator_comp_com.typ_id = Compensator
                    Generator_comp_com.RC = 0
                    Generator_comp_com.XC = 0

                    if line[1] == 'A':
                        Generator_comp_avr = Generator_comp.Avr
                        Generator_comp_avr.typ_id = avr_EA
                                                             
                        avr_type = Generator_comp.Avr
    #                app.PrintInfo(avr_type)

                        TR = 0.03
                        KF = 0.05
                        TF = 0.5
                        KA = 200
                        TA = 0
                        TA1 = 0            #????????????TA+TA1 = TA TA,TA1????
                        TE = 0.7
                        SEMAX = 0.267
                        SE75MAX = 0.074
                        KE = 0
                        VRMIN = 0
                        VRMAX = 0
                        M = 0
                        N = 0
                        EFDMAX = 3.5
                        VRMINMULT = -1
                    
                        if isfloat(line[16-chinese_count:20-chinese_count]):
                            TR = getfloatvalue(line[16-chinese_count:20-chinese_count],3)
                            
                        if isfloat(line[20-chinese_count:25-chinese_count]):
                            KA = getfloatvalue(line[20-chinese_count:25-chinese_count],2)
                                            
                        if isfloat(line[25-chinese_count:29-chinese_count]):
                            TA = getfloatvalue(line[25-chinese_count:29-chinese_count],2)
                            
                        if isfloat(line[29-chinese_count:33-chinese_count]):
                            TA1 = getfloatvalue(line[29-chinese_count:33-chinese_count],3)                            
    
                        if isfloat(line[37-chinese_count:41-chinese_count]):
                            KE = getfloatvalue(line[37-chinese_count:41-chinese_count],3)
                            
                        if isfloat(line[41-chinese_count:45-chinese_count]):
                            TE = getfloatvalue(line[41-chinese_count:45-chinese_count],3)

                        if isfloat(line[45-chinese_count:49-chinese_count]):
                            SE75MAX = getfloatvalue(line[45-chinese_count:49-chinese_count],3)

                        if isfloat(line[49-chinese_count:53-chinese_count]):
                            SEMAX = getfloatvalue(line[49-chinese_count:53-chinese_count],3)
                            
                    #？？？？
                    #EFDMIN???????????

                        if isfloat(line[53-chinese_count:58-chinese_count]):
                            EFDMIN = getfloatvalue(line[53-chinese_count:58-chinese_count],3)

                        if isfloat(line[58-chinese_count:62-chinese_count]):
                            EFDMAX = getfloatvalue(line[58-chinese_count:62-chinese_count],3)

                        N = math.log(SEMAX/SE75MAX)/(abs(EFDMAX-EFDMAX*0.75))
                        M = SEMAX/(math.exp(abs(N*EFDMAX)))
                                                
#                         app.PrintInfo(M)                

                        if isfloat(line[33-chinese_count:37-chinese_count]):                        
                            VRMINMULT = getfloatvalue(line[33-chinese_count:37-chinese_count],2)

                        VRMAX = (SEMAX + KE)*EFDMAX
                        VRMIN = VRMAX * VRMINMULT
    
                        if isfloat(line[62-chinese_count:66-chinese_count]):
                            KF = getfloatvalue(line[62-chinese_count:66-chinese_count],3)                            
                
                        if isfloat(line[66-chinese_count:70-chinese_count]):
                            TF = getfloatvalue(line[66-chinese_count:70-chinese_count],3)                            
    
                        avr_type.TR = TR
                        avr_type.KF = KF
                        avr_type.TF = TF
                        avr_type.KA = KA
                        avr_type.TA = TA
                        avr_type.TE = TE
                        avr_type.M = M
                        avr_type.N = N
                        avr_type.KE = KE                        
                        avr_type.TA1 = TA1
                        avr_type.VRMIN = VRMIN
                        avr_type.VRMAX = VRMAX
                                                                   

                    #EG
                    elif line[1] == 'G':
                        Generator_comp_avr = Generator_comp.Avr
                        Generator_comp_avr.typ_id = avr_EG
                                                             
                        avr_type = Generator_comp.Avr
    #                app.PrintInfo(avr_type)

                        TR = 0
                        KF = 0
                        TF = 0
                        KA = 0
                        TA = 0
                        TA1 = 0
                        VRMIN = -9999
                        VRMAX = 9999
                        EFDMAX = 9999
                        EFDMIN = -9999

                    
                        if isfloat(line[16-chinese_count:20-chinese_count]):
                            TR = getfloatvalue(line[16-chinese_count:20-chinese_count],3)
                            
                        if isfloat(line[20-chinese_count:25-chinese_count]):
                            KA = getfloatvalue(line[20-chinese_count:25-chinese_count],2)
                
                        if isfloat(line[25-chinese_count:29-chinese_count]):
                            TA = getfloatvalue(line[25-chinese_count:29-chinese_count],2)

                        if isfloat(line[29-chinese_count:33-chinese_count]):
                            TA1 = getfloatvalue(line[29-chinese_count:33-chinese_count],3)                           

                        if isfloat(line[62-chinese_count:66-chinese_count]):
                            KF = getfloatvalue(line[62-chinese_count:66-chinese_count],3)                            
                
                        if isfloat(line[66-chinese_count:70-chinese_count]):
                            TF = getfloatvalue(line[66-chinese_count:70-chinese_count],3)      
                            
                        if isfloat(line[53-chinese_count:58-chinese_count]):
                            EFDMIN = getfloatvalue(line[53-chinese_count:58-chinese_count],3)

                        if isfloat(line[58-chinese_count:62-chinese_count]):
                            EFDMAX = getfloatvalue(line[58-chinese_count:62-chinese_count],3)
    
                        if isfloat(line[62-chinese_count:66-chinese_count]):
                            KF = getfloatvalue(line[62-chinese_count:66-chinese_count],3)

                        if isfloat(line[66-chinese_count:70-chinese_count]):
                            TF = getfloatvalue(line[66-chinese_count:70-chinese_count],3)

                        VRMAX=EFDMAX
                        VRMIN=EFDMIN

                        avr_type.TR = TR
                        avr_type.KA = KA
                        avr_type.TA = TA
                        avr_type.TA1 = TA1
                        avr_type.KF = KF
                        avr_type.TF = TF
                        avr_type.VRMIN = VRMIN
                        avr_type.VRMAX = VRMAX
                        
                    #EJ
                    elif line[1] == 'J':
                        Generator_comp_avr = Generator_comp.Avr
                        Generator_comp_avr.typ_id = avr_EJ
                                                             
                        avr_type = Generator_comp.Avr
    #                app.PrintInfo(avr_type)

                        TR = 0
                        KF = 0
                        TF = 0
                        KA = 0
                        TA = 0
                        TA1 = 0
                        VRMIN = -9999
                        VRMAX = 9999
                        EFDMAX = 9999
                        EFDMIN = -9999

                    
                        if isfloat(line[16-chinese_count:20-chinese_count]):
                            TR = getfloatvalue(line[16-chinese_count:20-chinese_count],3)
                            
                        if isfloat(line[20-chinese_count:25-chinese_count]):
                            KA = getfloatvalue(line[20-chinese_count:25-chinese_count],2)
                
                        if isfloat(line[25-chinese_count:29-chinese_count]):
                            TA = getfloatvalue(line[25-chinese_count:29-chinese_count],2)

                        if isfloat(line[29-chinese_count:33-chinese_count]):
                            TA1 = getfloatvalue(line[29-chinese_count:33-chinese_count],3)

                        if isfloat(line[33-chinese_count:37-chinese_count]):                        
                            VRMIN = getfloatvalue(line[33-chinese_count:37-chinese_count],2)                                

                        if isfloat(line[62-chinese_count:66-chinese_count]):
                            KF = getfloatvalue(line[62-chinese_count:66-chinese_count],3)                            
                
                        if isfloat(line[66-chinese_count:70-chinese_count]):
                            TF = getfloatvalue(line[66-chinese_count:70-chinese_count],3)      
                            
                        if isfloat(line[53-chinese_count:58-chinese_count]):
                            EFDMIN = getfloatvalue(line[53-chinese_count:58-chinese_count],3)

                        if isfloat(line[58-chinese_count:62-chinese_count]):
                            EFDMAX = getfloatvalue(line[58-chinese_count:62-chinese_count],3)
    
                        if isfloat(line[62-chinese_count:66-chinese_count]):
                            KF = getfloatvalue(line[62-chinese_count:66-chinese_count],3)

                        if isfloat(line[66-chinese_count:70-chinese_count]):
                            TF = getfloatvalue(line[66-chinese_count:70-chinese_count],3)

                        VRMAX=VRMIN*(-1)

                        avr_type.Tr = TR
                        avr_type.Kf = KF
                        avr_type.Tf = TF
                        avr_type.Ta1 = TA1
                        avr_type.KA = KA
                        avr_type.TA = TA
                        avr_type.Vrmin = VRMIN
                        avr_type.EFDMIN = EFDMIN
                        avr_type.Vrmax = VRMAX                        
                        avr_type.EFDMAX = EFDMAX
                                

# ----------------------------------------------------------------------------------------------------
def GetFCard(bpa_file, bpa_str_ar,bus_name_ar):    
    #励磁机F卡
    
    prj = app.GetActiveProject()
    if prj is None:
        raise Exception("No project activated. Python Script stopped.")

    app.PrintInfo(prj)
    Network_Modell = prj.SearchObject("Network Model.IntPrjfolder")
    Network_Model=Network_Modell[0]

    Network_Dataa = Network_Model.SearchObject("Network Data.IntPrjfolder")
    Network_Data = Network_Dataa[0]
    app.PrintInfo(Network_Data)

    Nett = Network_Data.SearchObject("*.ElmNet")
    Net = Nett[0]
    app.PrintInfo(Net)

    Libraryy = prj.SearchObject("Library.IntPrjfolder")
    Library = Libraryy[0]
    app.PrintInfo(Library)

    user_defined_modell = Library.SearchObject("User Defined Models.IntPrjfolder")
    user_defined_model = user_defined_modell[0]
    app.PrintInfo(user_defined_model)

    BPA_Frame_FF = user_defined_model.SearchObject("BPA Frame(F).BlkDef")
    BPA_Frame_F = BPA_Frame_FF[0]
    app.PrintInfo(BPA_Frame_F)

    BPA_Frame_EE = user_defined_model.SearchObject("BPA Frame(E).BlkDef")
    BPA_Frame_E = BPA_Frame_EE[0]
    app.PrintInfo(BPA_Frame_E)

    Compensatorr = user_defined_model.SearchObject("Compensator.BlkDef")
    Compensator = Compensatorr[0]
    app.PrintInfo(Compensator)

    motor_Tonee = user_defined_model.SearchObject("motor_Tone.BlkDef")
    motor_Tone = motor_Tonee[0]
    app.PrintInfo(motor_Tone)

    avr_FFF = user_defined_model.SearchObject("avr_FF.BlkDef")
    avr_FF = avr_FFF[0]
    app.PrintInfo(avr_FF)

    avr_FGG = user_defined_model.SearchObject("avr_FG.BlkDef")
    avr_FG = avr_FGG[0]
    app.PrintInfo(avr_FG)

    avr_FJJ = user_defined_model.SearchObject("avr_FJ.BlkDef")
    avr_FJ = avr_FJJ[0]
    app.PrintInfo(avr_FJ)    

    avr_FKK = user_defined_model.SearchObject("avr_FK.BlkDef")
    avr_FK = avr_FKK[0]
    app.PrintInfo(avr_FK)

    avr_FM_00 = user_defined_model.SearchObject("avr_FM_0.BlkDef")
    avr_FM_0 = avr_FM_00[0]
    app.PrintInfo(avr_FM_0)

    avr_FM_11 = user_defined_model.SearchObject("avr_FM_1.BlkDef")
    avr_FM_1 = avr_FM_11[0]
    app.PrintInfo(avr_FM_1)

    avr_FN_00 = user_defined_model.SearchObject("avr_FN_0.BlkDef")
    avr_FN_0 = avr_FN_00[0]
    app.PrintInfo(avr_FN_0)

    avr_FN_11 = user_defined_model.SearchObject("avr_FN_1.BlkDef")
    avr_FN_1 = avr_FN_11[0]
    app.PrintInfo(avr_FN_1)

    avr_FO_00 = user_defined_model.SearchObject("avr_FO_0.BlkDef")
    avr_FO_0 = avr_FO_00[0]
    app.PrintInfo(avr_FO_0)

    avr_FO_11 = user_defined_model.SearchObject("avr_FO_1.BlkDef")
    avr_FO_1 = avr_FO_11[0]
    app.PrintInfo(avr_FO_1) 

    avr_FP_00 = user_defined_model.SearchObject("avr_FP_0.BlkDef")
    avr_FP_0 = avr_FP_00[0]
    app.PrintInfo(avr_FP_0)

    avr_FP_11 = user_defined_model.SearchObject("avr_FP_1.BlkDef")
    avr_FP_1 = avr_FP_11[0]
    app.PrintInfo(avr_FP_1) 

    avr_FQ_00 = user_defined_model.SearchObject("avr_FQ_0.BlkDef")
    avr_FQ_0 = avr_FQ_00[0]
    app.PrintInfo(avr_FQ_0)

    avr_FQ_11 = user_defined_model.SearchObject("avr_FQ_1.BlkDef")
    avr_FQ_1 = avr_FQ_11[0]
    app.PrintInfo(avr_FQ_1) 

    avr_FR_00 = user_defined_model.SearchObject("avr_FR_0.BlkDef")
    avr_FR_0 = avr_FR_00[0]
    app.PrintInfo(avr_FR_0)

    avr_FR_11 = user_defined_model.SearchObject("avr_FR_1.BlkDef")
    avr_FR_1 = avr_FR_11[0]
    app.PrintInfo(avr_FR_1) 

    avr_FS_00 = user_defined_model.SearchObject("avr_FS_0.BlkDef")
    avr_FS_0 = avr_FS_00[0]
    app.PrintInfo(avr_FS_0)

    avr_FS_11 = user_defined_model.SearchObject("avr_FS_1.BlkDef")
    avr_FS_1 = avr_FS_11[0]
    app.PrintInfo(avr_FS_1) 

    avr_FT_00 = user_defined_model.SearchObject("avr_FT_0.BlkDef")
    avr_FT_0 = avr_FT_00[0]
    app.PrintInfo(avr_FT_0)

    avr_FT_11 = user_defined_model.SearchObject("avr_FT_1.BlkDef")
    avr_FT_1 = avr_FT_11[0]
    app.PrintInfo(avr_FT_1) 

    avr_FU_00 = user_defined_model.SearchObject("avr_FU_0.BlkDef")
    avr_FU_0 = avr_FU_00[0]
    app.PrintInfo(avr_FU_0)

    avr_FU_11 = user_defined_model.SearchObject("avr_FU_1.BlkDef")
    avr_FU_1 = avr_FU_11[0]
    app.PrintInfo(avr_FU_1) 

    avr_FV_00 = user_defined_model.SearchObject("avr_FV_0.BlkDef")
    avr_FV_0 = avr_FV_00[0]
    app.PrintInfo(avr_FV_0)

    avr_FV_11 = user_defined_model.SearchObject("avr_FV_1.BlkDef")
    avr_FV_1 = avr_FV_11[0]
    app.PrintInfo(avr_FV_1)

    avr_FXX = user_defined_model.SearchObject("avr_FX.BlkDef")
    avr_FX = avr_FXX[0]
    app.PrintInfo(avr_FX) 

    #读取所有的发电机数据，后面作判断依据
    
    for line in bpa_str_ar:
        line = line.rstrip('\n')
        if line == "": continue        

        #F卡
        if line[0] == 'F':
            line = line + ' '*(80-len(line))
            chinese_count = 0
            #判断中文的个数
            for i in range (3,10):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count = chinese_count + 1
            name = line[3:11-chinese_count].strip()
            if isfloat(line[11-chinese_count:15-chinese_count]):base = float(line[11-chinese_count:15-chinese_count])
            #名称+电压
            Generator = name + str(base)                
            #找到节点名对应的编号
            if Generator in bus_name_ar:
                Generator_num = bus_name_ar[Generator]

                #找到对应的comp模块
                Generator_comp_name = str(Generator_num) + '_comp'
                Generator_compp = Net.SearchObject(str(Generator_comp_name))
                Generator_comp = Generator_compp[0]
    
                if not Generator_comp == None:

                    sym = Generator_comp.Sym
                        
                    sym_typ = sym.typ_id
                    MVABASE = sym_typ.sgn
                    
                    Generator_Compensator_dsll = Generator_comp.SearchObject('Compensator.ElmDsl')
                    Generator_Compensator_dsl = Generator_Compensator_dsll[0]
                    if Generator_Compensator_dsl == None:
                        Generator_comp.CreateObject('ElmDsl','Compensator')
                        Generator_Compensator_dsll = Generator_comp.SearchObject('Compensator.ElmDsl')
                        Generator_Compensator_dsl = Generator_Compensator_dsll[0]
                    Generator_comp.Compensator = Generator_Compensator_dsl
                    Generator_comp_com = Generator_comp.Compensator
                    Generator_comp_com.typ_id = Compensator
#                    Generator_comp_com.RC = 0
#                    Generator_comp_com.XC = 0

                    Generator_motor_dsll = Generator_comp.SearchObject('motor.ElmDsl')
                    Generator_motor_dsl = Generator_motor_dsll[0]
                    if Generator_motor_dsl == None:
                        Generator_comp.CreateObject('ElmDsl','motor')
                        Generator_motor_dsll = Generator_comp.SearchObject('motor.ElmDsl')
                        Generator_motor_dsl = Generator_motor_dsll[0]                
                    Generator_comp.Prime_Moter = Generator_motor_dsl
                    Generator_comp_motor = Generator_comp.Prime_Moter
                    Generator_comp_motor.typ_id = motor_Tone
                    

                                                          
                    #avr.dsl
    #                app.PrintInfo(Generator_num)
                    Generator_avr_dsll = Generator_comp.SearchObject('avr_dsl.ElmDsl')
                    Generator_avr_dsl = Generator_avr_dsll[0]
                    if Generator_avr_dsl == None:
                        Generator_comp.CreateObject('ElmDsl','avr_dsl')
                        Generator_avr_dsll = Generator_comp.SearchObject('avr_dsl.ElmDsl')
                        Generator_avr_dsl = Generator_avr_dsll[0]
                    Generator_comp.Avr= Generator_avr_dsl

                    #FF
                    if line[1] == 'F':
                        Generator_comp_avr = Generator_comp.Avr
                        Generator_comp_avr.typ_id = avr_FF

                        avr_type = Generator_comp.Avr
                        Generator_Compensator = Generator_comp.Compensator

                        RC = 0
                        XC = 0
                        
                        
                        if isfloat(line[16-chinese_count:21-chinese_count]):
                            RC = getfloatvalue(line[16-chinese_count:21-chinese_count],4)

                        if isfloat(line[21-chinese_count:26-chinese_count]):
                            XC = getfloatvalue(line[21-chinese_count:26-chinese_count],4)

                        Generator_Compensator.RC = RC
                        Generator_Compensator.XC = XC

                        TR = 0
                        TC = 0
                        TB = 0
                        KA = 0
                        TA = 0
                        TE = 0
                        KE = 0
                        VAMIN = -9999
                        VRMIN = -9999
                        VAMAX = 9999
                        VRMAX = 9999

                        if isfloat(line[26-chinese_count:31-chinese_count]):
                            TR = getfloatvalue(line[26-chinese_count:31-chinese_count],4)
                        
                        if isfloat(line[31-chinese_count:36-chinese_count]):
                            VAMAX = getfloatvalue(line[31-chinese_count:36-chinese_count],3)
                        
                        if isfloat(line[36-chinese_count:41-chinese_count]):
                            VAMIN = getfloatvalue(line[36-chinese_count:41-chinese_count],3)                        
                
                        if isfloat(line[41-chinese_count:46-chinese_count]):
                            TB = getfloatvalue(line[41-chinese_count:46-chinese_count],3)

                        if isfloat(line[46-chinese_count:51-chinese_count]):
                            TC = getfloatvalue(line[46-chinese_count:51-chinese_count],3)

                        if isfloat(line[51-chinese_count:56-chinese_count]):
                            KA = getfloatvalue(line[51-chinese_count:56-chinese_count],2)                        
                            
                        if isfloat(line[56-chinese_count:61-chinese_count]):
                            TA = getfloatvalue(line[56-chinese_count:61-chinese_count],3)

                        if isfloat(line[61-chinese_count:66-chinese_count]):
                            VRMAX = getfloatvalue(line[61-chinese_count:66-chinese_count],3)                            

                        if isfloat(line[66-chinese_count:71-chinese_count]):
                            VRMIN = getfloatvalue(line[66-chinese_count:71-chinese_count],3)

                        if isfloat(line[71-chinese_count:76-chinese_count]):
                            KE = getfloatvalue(line[71-chinese_count:76-chinese_count],3)
                            
                        if isfloat(line[76-chinese_count:80-chinese_count]):
                            TE = getfloatvalue(line[76-chinese_count:80-chinese_count],3)

                        avr_type.Tr = TR
                        avr_type.Tc = TC
                        avr_type.Tb = TB
                        avr_type.Ka = KA
                        avr_type.Ta = TA 
                        avr_type.Te = TE
                        avr_type.Ke = KE
                        avr_type.Vamin = VAMIN
                        avr_type.Vrmin = VRMIN
                        avr_type.Vamax = VAMAX
                        avr_type.Vrmax = VRMAX

                    #FG
                    if line[1] == 'G':
                        Generator_comp_avr = Generator_comp.Avr
                        Generator_comp_avr.typ_id = avr_FG

                        avr_type = Generator_comp.Avr
                        Generator_Compensator = Generator_comp.Compensator

                        RC = 0
                        XC = 0
                        
                        if isfloat(line[16-chinese_count:21-chinese_count]):
                            RC = getfloatvalue(line[16-chinese_count:21-chinese_count],4)

                        if isfloat(line[21-chinese_count:26-chinese_count]):
                            XC = getfloatvalue(line[21-chinese_count:26-chinese_count],4)


                        Generator_Compensator.RC = RC
                        Generator_Compensator.XC = XC

                        TR = 0
                        TC = 0
                        TB = 0
                        KA = 0
                        TA = 0
                        VIMIN = -9999
                        VRMIN = -9999
                        VIMAX = 9999
                        VRMAX = 9999

                        if isfloat(line[26-chinese_count:31-chinese_count]):
                            TR = getfloatvalue(line[26-chinese_count:31-chinese_count],4)
                        
                        if isfloat(line[31-chinese_count:36-chinese_count]):
                            VIMAX = getfloatvalue(line[31-chinese_count:36-chinese_count],3)
                        
                        if isfloat(line[36-chinese_count:41-chinese_count]):
                            VIMIN = getfloatvalue(line[36-chinese_count:41-chinese_count],3)                        
                
                        if isfloat(line[41-chinese_count:46-chinese_count]):
                            TB = getfloatvalue(line[41-chinese_count:46-chinese_count],3)

                        if isfloat(line[46-chinese_count:51-chinese_count]):
                            TC = getfloatvalue(line[46-chinese_count:51-chinese_count],3)

                        if isfloat(line[51-chinese_count:56-chinese_count]):
                            KA = getfloatvalue(line[51-chinese_count:56-chinese_count],2)                        
                            
                        if isfloat(line[56-chinese_count:61-chinese_count]):
                            TA = getfloatvalue(line[56-chinese_count:61-chinese_count],3)

                        if isfloat(line[61-chinese_count:66-chinese_count]):
                            VRMAX = getfloatvalue(line[61-chinese_count:66-chinese_count],3)                            

                        if isfloat(line[66-chinese_count:71-chinese_count]):
                            VRMIN = getfloatvalue(line[66-chinese_count:71-chinese_count],3)
                        
                        avr_type.Tr = TR
                        avr_type.Tb = TB
                        avr_type.Tc = TC
                        avr_type.Ka = KA
                        avr_type.Ta = TA
                        avr_type.Vimin = VIMIN
                        avr_type.Vrmin = VRMIN
                        avr_type.Vimax = VIMAX
                        avr_type.Vrmax = VRMAX
                        

                    #FJ                   
                    if line[1] == 'J':
                        Generator_comp_avr = Generator_comp.Avr
                        Generator_comp_avr.typ_id = avr_FJ

                        avr_type = Generator_comp.Avr
                        Generator_Compensator = Generator_comp.Compensator

                        RC = 0
                        XC = 0
                        
                        
                        if isfloat(line[16-chinese_count:21-chinese_count]):
                            RC = getfloatvalue(line[16-chinese_count:21-chinese_count],4)

                        if isfloat(line[21-chinese_count:26-chinese_count]):
                            XC = getfloatvalue(line[21-chinese_count:26-chinese_count],4)

                        Generator_Compensator.RC = RC
                        Generator_Compensator.XC = XC

                        TR = 0
                        TC = 0
                        TB = 0
                        KA = 0
                        TA = 0
                        VRMIN = -9999
                        VRMAX = 9999

                        if isfloat(line[26-chinese_count:31-chinese_count]):
                            TR = getfloatvalue(line[26-chinese_count:31-chinese_count],4)

                        
                        if isfloat(line[41-chinese_count:46-chinese_count]):
                            TB = getfloatvalue(line[41-chinese_count:46-chinese_count],3)

                        if isfloat(line[46-chinese_count:51-chinese_count]):
                            TC = getfloatvalue(line[46-chinese_count:51-chinese_count],3)

                        if isfloat(line[51-chinese_count:56-chinese_count]):
                            KA = getfloatvalue(line[51-chinese_count:56-chinese_count],2)                        
                            
                        if isfloat(line[56-chinese_count:61-chinese_count]):
                            TA = getfloatvalue(line[56-chinese_count:61-chinese_count],3)

                        if isfloat(line[61-chinese_count:66-chinese_count]):
                            VRMAX = getfloatvalue(line[61-chinese_count:66-chinese_count],3)                            

                        if isfloat(line[66-chinese_count:71-chinese_count]):
                            VRMIN = getfloatvalue(line[66-chinese_count:71-chinese_count],3)

                        avr_type.Tr = TR
                        avr_type.Tc = TC
                        avr_type.Tb = TB
                        avr_type.Ka = KA
                        avr_type.Ta = TA 
                        avr_type.Vrmin = VRMIN
                        avr_type.Vrmax = VRMAX

                    #FK
                    if line[1] == 'K':
                        Generator_comp_avr = Generator_comp.Avr
                        Generator_comp_avr.typ_id = avr_FK

                        avr_type = Generator_comp.Avr
                        Generator_Compensator = Generator_comp.Compensator

                        RC = 0
                        XC = 0
                        
                        if isfloat(line[16-chinese_count:21-chinese_count]):
                            RC = getfloatvalue(line[16-chinese_count:21-chinese_count],4)
                            
                        if isfloat(line[21-chinese_count:26-chinese_count]):
                            XC = getfloatvalue(line[21-chinese_count:26-chinese_count],4)

                        Generator_Compensator.RC = RC
                        Generator_Compensator.XC = XC

                        TR = 0
                        TC = 0
                        TB = 0
                        KA = 0
                        TA = 0
                        VIMIN = -9999
                        VIMAX = 9999
                        VRMIN = -9999
                        VRMAX = 9999

                        if isfloat(line[26-chinese_count:31-chinese_count]):
                            TR = getfloatvalue(line[26-chinese_count:31-chinese_count],4)
                        
                        if isfloat(line[31-chinese_count:36-chinese_count]):
                            VIMAX = getfloatvalue(line[31-chinese_count:36-chinese_count],3)
                        
                        if isfloat(line[36-chinese_count:41-chinese_count]):
                            VIMIN = getfloatvalue(line[36-chinese_count:41-chinese_count],3)                        
                
                        if isfloat(line[41-chinese_count:46-chinese_count]):
                            TB = getfloatvalue(line[41-chinese_count:46-chinese_count],3)

                        if isfloat(line[46-chinese_count:51-chinese_count]):
                            TC = getfloatvalue(line[46-chinese_count:51-chinese_count],3)

                        if isfloat(line[51-chinese_count:56-chinese_count]):
                            KA = getfloatvalue(line[51-chinese_count:56-chinese_count],2)                        
                            
                        if isfloat(line[56-chinese_count:61-chinese_count]):
                            TA = getfloatvalue(line[56-chinese_count:61-chinese_count],3)

                        if isfloat(line[61-chinese_count:66-chinese_count]):
                            VRMAX = getfloatvalue(line[61-chinese_count:66-chinese_count],3)                            

                        if isfloat(line[66-chinese_count:71-chinese_count]):
                            VRMIN = getfloatvalue(line[66-chinese_count:71-chinese_count],3)

                        avr_type.Tr = TR
                        avr_type.Tc = TC
                        avr_type.Tb = TB
                        avr_type.Ka = KA
                        avr_type.Ta = TA
                        avr_type.Vimin = VIMIN
                        avr_type.Vrmin = VRMIN
                        avr_type.Vimax = VIMAX
                        avr_type.Vrmax = VRMAX

                    #FM/FQ
                    if line[1] == 'M' or line[1] == 'Q':
                        Generator_comp_avr = Generator_comp.Avr
                        pss_id = 0
                        if isint(line[2]):
                            pss_id = int(line[2])
                        if pss_id == 0:
                            Generator_comp_avr.typ_id = avr_FM_0
                        else:
                            Generator_comp_avr.typ_id = avr_FM_1

                        avr_type = Generator_comp.Avr
                        Generator_Compensator = Generator_comp.Compensator

                        RC = 0
                        XC = 0
                        
                        if isfloat(line[16-chinese_count:20-chinese_count]):
                            RC = getfloatvalue(line[16-chinese_count:20-chinese_count],3)

                        if isfloat(line[20-chinese_count:24-chinese_count]):
                            XC = getfloatvalue(line[20-chinese_count:24-chinese_count],3)

                        Generator_Compensator.RC = RC
                        Generator_Compensator.XC = XC

                        TR = 0
                        T3 = 0
                        T4 = 0
                        KF = 0
                        TF = 0
                        KH = 0
                        K = 0
                        KV = 0
                        T1 = 0
                        T2 = 0
                        KA = 0
                        TA = 0

                        if isfloat(line[24-chinese_count:29-chinese_count]):
                            TR = getfloatvalue(line[24-chinese_count:29-chinese_count],3)

                        if isfloat(line[29-chinese_count:34-chinese_count]):
                            K = getfloatvalue(line[29-chinese_count:34-chinese_count],3)
                            
                        if isfloat(line[34-chinese_count:37-chinese_count]):
                            KV = getfloatvalue(line[34-chinese_count:37-chinese_count],0)

                        if isfloat(line[37-chinese_count:42-chinese_count]):
                            T1 = getfloatvalue(line[37-chinese_count:42-chinese_count],3)

                        if isfloat(line[42-chinese_count:47-chinese_count]):
                            T2 = getfloatvalue(line[42-chinese_count:47-chinese_count],3)

                        if isfloat(line[47-chinese_count:52-chinese_count]):
                            T3 = getfloatvalue(line[47-chinese_count:52-chinese_count],3)

                        if isfloat(line[52-chinese_count:57-chinese_count]):
                            T4 = getfloatvalue(line[52-chinese_count:57-chinese_count],3)

                        if isfloat(line[57-chinese_count:62-chinese_count]):
                            KA = getfloatvalue(line[57-chinese_count:62-chinese_count],3)

                        if isfloat(line[62-chinese_count:67-chinese_count]):
                            TA = getfloatvalue(line[62-chinese_count:67-chinese_count],3)

                        if isfloat(line[67-chinese_count:72-chinese_count]):
                            KF = getfloatvalue(line[67-chinese_count:72-chinese_count],3)

                        if isfloat(line[72-chinese_count:76-chinese_count]):
                            TF = getfloatvalue(line[72-chinese_count:76-chinese_count],3)
                                
                        if isfloat(line[76-chinese_count:80-chinese_count]):
                            KH = getfloatvalue(line[76-chinese_count:80-chinese_count],2)

                        avr_type.Tr = TR
                        avr_type.T3 = T3
                        avr_type.T4 = T4
                        avr_type.Ka = KA
                        avr_type.Ta = TA
                        avr_type.Kh = KH
                        avr_type.Kf = KF
                        avr_type.Tf = TF
                        avr_type.K = K
                        avr_type.T1 = T1
                        avr_type.Kv = KV
                        avr_type.T2 = T2



                    #FN/FR
                    if line[1] == 'N' or line[1] == 'R':
                        Generator_comp_avr = Generator_comp.Avr
                        pss_id = 0
                        if isint(line[2]):
                            pss_id = int(line[2])
                        if pss_id == 0:
                            Generator_comp_avr.typ_id = avr_FN_0
                        else:
                            Generator_comp_avr.typ_id = avr_FN_1

                        avr_type = Generator_comp.Avr
                        Generator_Compensator = Generator_comp.Compensator

                        RC = 0
                        XC = 0
                        
                        if isfloat(line[16-chinese_count:20-chinese_count]):
                            RC = getfloatvalue(line[16-chinese_count:20-chinese_count],3)

                        if isfloat(line[20-chinese_count:24-chinese_count]):
                            XC = getfloatvalue(line[20-chinese_count:24-chinese_count],3)

                        Generator_Compensator.RC = RC
                        Generator_Compensator.XC = XC

                        TR = 0
                        T3 = 0
                        T4 = 0
                        KF = 0
                        TF = 0
                        KH = 0
                        K = 0
                        KV = 0
                        T1 = 0
                        T2 = 0
                        KA = 0
                        TA = 0

                        if isfloat(line[24-chinese_count:29-chinese_count]):
                            TR = getfloatvalue(line[24-chinese_count:29-chinese_count],3)

                        if isfloat(line[29-chinese_count:34-chinese_count]):
                            K = getfloatvalue(line[29-chinese_count:34-chinese_count],3)
                            
                        if isfloat(line[34-chinese_count:37-chinese_count]):
                            KV = getfloatvalue(line[34-chinese_count:37-chinese_count],0)

                        if isfloat(line[37-chinese_count:42-chinese_count]):
                            T1 = getfloatvalue(line[37-chinese_count:42-chinese_count],3)

                        if isfloat(line[42-chinese_count:47-chinese_count]):
                            T2 = getfloatvalue(line[42-chinese_count:47-chinese_count],3)

                        if isfloat(line[47-chinese_count:52-chinese_count]):
                            T3 = getfloatvalue(line[47-chinese_count:52-chinese_count],3)

                        if isfloat(line[52-chinese_count:57-chinese_count]):
                            T4 = getfloatvalue(line[52-chinese_count:57-chinese_count],3)

                        if isfloat(line[57-chinese_count:62-chinese_count]):
                            KA = getfloatvalue(line[57-chinese_count:62-chinese_count],3)

                        if isfloat(line[62-chinese_count:67-chinese_count]):
                            TA = getfloatvalue(line[62-chinese_count:67-chinese_count],3)

                        if isfloat(line[67-chinese_count:72-chinese_count]):
                            KF = getfloatvalue(line[67-chinese_count:72-chinese_count],3)

                        if isfloat(line[72-chinese_count:76-chinese_count]):
                            TF = getfloatvalue(line[72-chinese_count:76-chinese_count],3)
                                
                        if isfloat(line[76-chinese_count:80-chinese_count]):
                            KH = getfloatvalue(line[76-chinese_count:80-chinese_count],2)

                        avr_type.Tr = TR
                        avr_type.T3 = T3
                        avr_type.T4 = T4
                        avr_type.Ka = KA
                        avr_type.Ta = TA
                        avr_type.Kh = KH
                        avr_type.Kf = KF
                        avr_type.Tf = TF
                        avr_type.K = K
                        avr_type.T1 = T1
                        avr_type.Kv = KV
                        avr_type.T2 = T2
                        

                    #FO/FS
                    if line[1] == 'O' or line[1] == 'S':
                        Generator_comp_avr = Generator_comp.Avr
                        pss_id = 0
                        if isint(line[2]):
                            pss_id = int(line[2])
                        if pss_id == 0:
                            Generator_comp_avr.typ_id = avr_FO_0
                        else:
                            Generator_comp_avr.typ_id = avr_FO_1

                        avr_type = Generator_comp.Avr
                        Generator_Compensator = Generator_comp.Compensator

                        RC = 0
                        XC = 0
                        
                        if isfloat(line[16-chinese_count:20-chinese_count]):
                            RC = getfloatvalue(line[16-chinese_count:20-chinese_count],3)

                        if isfloat(line[20-chinese_count:24-chinese_count]):
                            XC = getfloatvalue(line[20-chinese_count:24-chinese_count],3)

                        Generator_Compensator.RC = RC
                        Generator_Compensator.XC = XC
                                                

                        TR = 0
                        T3 = 0
                        T4 = 0
                        KF = 0
                        TF = 0
                        KH = 0
                        K = 0
                        KV = 0
                        T1 = 0
                        T2 = 0
                        KA = 0
                        TA = 0

                        if isfloat(line[24-chinese_count:29-chinese_count]):
                            TR = getfloatvalue(line[24-chinese_count:29-chinese_count],3)

                        if isfloat(line[29-chinese_count:34-chinese_count]):
                            K = getfloatvalue(line[29-chinese_count:34-chinese_count],3)
                            
                        if isfloat(line[34-chinese_count:37-chinese_count]):
                            KV = getfloatvalue(line[34-chinese_count:37-chinese_count],0)

                        if isfloat(line[37-chinese_count:42-chinese_count]):
                            T1 = getfloatvalue(line[37-chinese_count:42-chinese_count],3)

                        if isfloat(line[42-chinese_count:47-chinese_count]):
                            T2 = getfloatvalue(line[42-chinese_count:47-chinese_count],3)

                        if isfloat(line[47-chinese_count:52-chinese_count]):
                            T3 = getfloatvalue(line[47-chinese_count:52-chinese_count],3)

                        if isfloat(line[52-chinese_count:57-chinese_count]):
                            T4 = getfloatvalue(line[52-chinese_count:57-chinese_count],3)

                        if isfloat(line[57-chinese_count:62-chinese_count]):
                            KA = getfloatvalue(line[57-chinese_count:62-chinese_count],3)

                        if isfloat(line[62-chinese_count:67-chinese_count]):
                            TA = getfloatvalue(line[62-chinese_count:67-chinese_count],3)

                        if isfloat(line[67-chinese_count:72-chinese_count]):
                            KF = getfloatvalue(line[67-chinese_count:72-chinese_count],3)

                        if isfloat(line[72-chinese_count:76-chinese_count]):
                            TF = getfloatvalue(line[72-chinese_count:76-chinese_count],3)
                                
                        if isfloat(line[76-chinese_count:80-chinese_count]):
                            KH = getfloatvalue(line[76-chinese_count:80-chinese_count],2)

                        avr_type.Tr = TR
                        avr_type.T3 = T3
                        avr_type.T4 = T4
                        avr_type.Ka = KA
                        avr_type.Ta = TA
                        avr_type.Kh = KH
                        avr_type.Kf = KF
                        avr_type.Tf = TF
                        avr_type.K = K
                        avr_type.T1 = T1
                        avr_type.Kv = KV
                        avr_type.T2 = T2
                        
                    #FP/FT
                    if line[1] == 'P' or line[1] == 'T':
                        Generator_comp_avr = Generator_comp.Avr
                        pss_id = 0
                        if isint(line[2]):
                            pss_id = int(line[2])
                        if pss_id == 0:
                            Generator_comp_avr.typ_id = avr_FP_0
                        else:
                            Generator_comp_avr.typ_id = avr_FP_1

                        avr_type = Generator_comp.Avr
                        Generator_Compensator = Generator_comp.Compensator

                        RC = 0
                        XC = 0
                        
                        if isfloat(line[16-chinese_count:20-chinese_count]):
                            RC = getfloatvalue(line[16-chinese_count:20-chinese_count],3)

                        if isfloat(line[20-chinese_count:24-chinese_count]):
                            XC = getfloatvalue(line[20-chinese_count:24-chinese_count],3)

                        Generator_Compensator.RC = RC
                        Generator_Compensator.XC = XC

                        TR = 0
                        T3 = 0
                        T4 = 0
                        KF = 0
                        TF = 0
                        KH = 0
                        K = 0
                        KV = 0
                        T1 = 0
                        T2 = 0
                        KA = 0
                        TA = 0

                        if isfloat(line[24-chinese_count:29-chinese_count]):
                            TR = getfloatvalue(line[24-chinese_count:29-chinese_count],3)

                        if isfloat(line[29-chinese_count:34-chinese_count]):
                            K = getfloatvalue(line[29-chinese_count:34-chinese_count],3)
                            
                        if isfloat(line[34-chinese_count:37-chinese_count]):
                            KV = getfloatvalue(line[34-chinese_count:37-chinese_count],0)

                        if isfloat(line[37-chinese_count:42-chinese_count]):
                            T1 = getfloatvalue(line[37-chinese_count:42-chinese_count],3)

                        if isfloat(line[42-chinese_count:47-chinese_count]):
                            T2 = getfloatvalue(line[42-chinese_count:47-chinese_count],3)

                        if isfloat(line[47-chinese_count:52-chinese_count]):
                            T3 = getfloatvalue(line[47-chinese_count:52-chinese_count],3)

                        if isfloat(line[52-chinese_count:57-chinese_count]):
                            T4 = getfloatvalue(line[52-chinese_count:57-chinese_count],3)

                        if isfloat(line[57-chinese_count:62-chinese_count]):
                            KA = getfloatvalue(line[57-chinese_count:62-chinese_count],3)

                        if isfloat(line[62-chinese_count:67-chinese_count]):
                            TA = getfloatvalue(line[62-chinese_count:67-chinese_count],3)

                        if isfloat(line[67-chinese_count:72-chinese_count]):
                            KF = getfloatvalue(line[67-chinese_count:72-chinese_count],3)

                        if isfloat(line[72-chinese_count:76-chinese_count]):
                            TF = getfloatvalue(line[72-chinese_count:76-chinese_count],3)
                                
                        if isfloat(line[76-chinese_count:80-chinese_count]):
                            KH = getfloatvalue(line[76-chinese_count:80-chinese_count],2)

                        avr_type.Tr = TR
                        avr_type.T3 = T3
                        avr_type.T4 = T4
                        avr_type.Ka = KA
                        avr_type.Ta = TA
                        avr_type.Kh = KH
                        avr_type.Kf = KF
                        avr_type.Tf = TF
                        avr_type.K = K
                        avr_type.T1 = T1
                        avr_type.Kv = KV
                        avr_type.T2 = T2

                    #FU
                    if line[1] == 'U':
                        Generator_comp_avr = Generator_comp.Avr
                        pss_id = 0
                        if isint(line[2]):pss_id = int(line[2])
                        if pss_id == 0:
                            Generator_comp_avr.typ_id = avr_FU_0
                        else:
                            Generator_comp_avr.typ_id = avr_FU_1

                        avr_type = Generator_comp.Avr
                        Generator_Compensator = Generator_comp.Compensator

                        RC = 0
                        XC = 0                        
                        
                        if isfloat(line[16-chinese_count:20-chinese_count]):
                            RC = getfloatvalue(line[16-chinese_count:20-chinese_count],3)

                        if isfloat(line[20-chinese_count:24-chinese_count]):
                            XC = getfloatvalue(line[20-chinese_count:24-chinese_count],3)

                        Generator_Compensator.RC = RC
                        Generator_Compensator.XC = XC

                        TR = 0
                        T3 = 0
                        T4 = 0
                        KF = 0
                        TF = 0
                        K = 0
                        KV = 0
                        T1 = 0
                        T2 = 0
                        KA = 0
                        TA = 0

                        if isfloat(line[24-chinese_count:29-chinese_count]):
                            TR = getfloatvalue(line[24-chinese_count:29-chinese_count],3)

                        if isfloat(line[29-chinese_count:34-chinese_count]):
                            K = getfloatvalue(line[29-chinese_count:34-chinese_count],3)
                            
                        if isfloat(line[34-chinese_count:37-chinese_count]):
                            KV = getfloatvalue(line[34-chinese_count:37-chinese_count],0)

                        if isfloat(line[37-chinese_count:42-chinese_count]):
                            T1 = getfloatvalue(line[37-chinese_count:42-chinese_count],3)

                        if isfloat(line[42-chinese_count:47-chinese_count]):
                            T2 = getfloatvalue(line[42-chinese_count:47-chinese_count],3)

                        if isfloat(line[47-chinese_count:52-chinese_count]):
                            T3 = getfloatvalue(line[47-chinese_count:52-chinese_count],3)

                        if isfloat(line[52-chinese_count:57-chinese_count]):
                            T4 = getfloatvalue(line[52-chinese_count:57-chinese_count],3)

                        if isfloat(line[57-chinese_count:62-chinese_count]):
                            KA = getfloatvalue(line[57-chinese_count:62-chinese_count],3)

                        if isfloat(line[62-chinese_count:67-chinese_count]):
                            TA = getfloatvalue(line[62-chinese_count:67-chinese_count],3)
                            if TA == 0:
                                TA = 0.001

                        if isfloat(line[67-chinese_count:72-chinese_count]):
                            KF = getfloatvalue(line[67-chinese_count:72-chinese_count],3)

                        if isfloat(line[72-chinese_count:76-chinese_count]):
                            TF = getfloatvalue(line[72-chinese_count:76-chinese_count],3)

                        avr_type.Tr = TR
                        avr_type.T3 = T3
                        avr_type.T4 = T4
                        avr_type.Ka = KA
                        avr_type.Ta = TA
                        avr_type.Kf = KF
                        avr_type.Tf = TF
                        avr_type.K = K
                        avr_type.T1 = T1
                        avr_type.Kv = KV
                        avr_type.T2 = T2

                    #FV
                    if line[1] == 'V':
                        Generator_comp_avr = Generator_comp.Avr
                        pss_id = 0
                        if isint(line[2]):pss_id = int(line[2])
                        if pss_id == 0:
                            Generator_comp_avr.typ_id = avr_FV_0
                        else:
                            Generator_comp_avr.typ_id = avr_FV_1

                        avr_type = Generator_comp.Avr
                        Generator_Compensator = Generator_comp.Compensator

                        RC = 0
                        XC = 0                        
                        
                        if isfloat(line[16-chinese_count:20-chinese_count]):
                            RC = getfloatvalue(line[16-chinese_count:20-chinese_count],3)

                        if isfloat(line[20-chinese_count:24-chinese_count]):
                            XC = getfloatvalue(line[20-chinese_count:24-chinese_count],3)

                        Generator_Compensator.RC = RC
                        Generator_Compensator.XC = XC

                        TR = 0
                        T3 = 0
                        T4 = 0
                        KF = 0
                        TF = 0
                        K = 0
                        KV = 0
                        T1 = 0
                        T2 = 0
                        KA = 0
                        TA = 0

                        if isfloat(line[24-chinese_count:29-chinese_count]):
                            TR = getfloatvalue(line[24-chinese_count:29-chinese_count],3)

                        if isfloat(line[29-chinese_count:34-chinese_count]):
                            K = getfloatvalue(line[29-chinese_count:34-chinese_count],3)
                            
                        if isfloat(line[34-chinese_count:37-chinese_count]):
                            KV = getfloatvalue(line[34-chinese_count:37-chinese_count],0)

                        if isfloat(line[37-chinese_count:42-chinese_count]):
                            T1 = getfloatvalue(line[37-chinese_count:42-chinese_count],3)

                        if isfloat(line[42-chinese_count:47-chinese_count]):
                            T2 = getfloatvalue(line[42-chinese_count:47-chinese_count],3)

                        if isfloat(line[47-chinese_count:52-chinese_count]):
                            T3 = getfloatvalue(line[47-chinese_count:52-chinese_count],3)

                        if isfloat(line[52-chinese_count:57-chinese_count]):
                            T4 = getfloatvalue(line[52-chinese_count:57-chinese_count],3)

                        if isfloat(line[57-chinese_count:62-chinese_count]):
                            KA = getfloatvalue(line[57-chinese_count:62-chinese_count],3)

                        if isfloat(line[62-chinese_count:67-chinese_count]):
                            TA = getfloatvalue(line[62-chinese_count:67-chinese_count],3)
                            if TA == 0:
                                TA = 0.001

                        if isfloat(line[67-chinese_count:72-chinese_count]):
                            KF = getfloatvalue(line[67-chinese_count:72-chinese_count],3)

                        if isfloat(line[72-chinese_count:76-chinese_count]):
                            TF = getfloatvalue(line[72-chinese_count:76-chinese_count],3)

                        avr_type.Tr = TR
                        avr_type.T3 = T3
                        avr_type.T4 = T4
                        avr_type.Ka = KA
                        avr_type.Ta = TA
                        avr_type.Kf = KF
                        avr_type.Tf = TF
                        avr_type.K = K
                        avr_type.T1 = T1
                        avr_type.Kv = KV
                        avr_type.T2 = T2

                    #FX
                    if line[1] == 'X' and line[2] != '+': 
                        Generator_comp_avr = Generator_comp.Avr
                        Generator_comp_avr.typ_id = avr_FX

                        avr_type = Generator_comp.Avr
                        Generator_Compensator = Generator_comp.Compensator

                        RC = 0
                        XC = 0
                        
                        if isfloat(line[16-chinese_count:20-chinese_count]):
                            RC = getfloatvalue(line[16-chinese_count:20-chinese_count],3)

                        if isfloat(line[20-chinese_count:24-chinese_count]):
                            XC = getfloatvalue(line[20-chinese_count:24-chinese_count],3)

                        Generator_Compensator.RC = RC
                        Generator_Compensator.XC = XC

                        TR = 0
                        KA = 0
                        TA = 0
                        KP = 0
                        KI = 0
                        IKP = 0
                        IKI = 0
                        KT = 0
                        TT = 0
                        VRMIN = -9999
                        VFMIN = -9999
                        VRMAX = 9999
                        VFMAX = 9999
                        KIFD = 1
                        TIFD = 0
                        EFDMAX = 9999
                        EFDMIN = -9999

                        if isfloat(line[24-chinese_count:28-chinese_count]):
                            TR = getfloatvalue(line[24-chinese_count:28-chinese_count],4)

                        if isfloat(line[28-chinese_count:32-chinese_count]):
                            KA = getfloatvalue(line[28-chinese_count:32-chinese_count],3)
                            
                        if isfloat(line[32-chinese_count:36-chinese_count]):
                            TA = getfloatvalue(line[32-chinese_count:36-chinese_count],4)

                        if isfloat(line[36-chinese_count:40-chinese_count]):
                            KP = getfloatvalue(line[36-chinese_count:40-chinese_count],3)

                        if isfloat(line[40-chinese_count:44-chinese_count]):
                            KI = getfloatvalue(line[40-chinese_count:44-chinese_count],3)

                        if isfloat(line[44-chinese_count:49-chinese_count]):
                            VRMAX = getfloatvalue(line[44-chinese_count:49-chinese_count],3)

                        if isfloat(line[49-chinese_count:54-chinese_count]):
                            VRMIN = getfloatvalue(line[49-chinese_count:54-chinese_count],3)

                        if isfloat(line[54-chinese_count:58-chinese_count]):
                            IKP = getfloatvalue(line[54-chinese_count:58-chinese_count],3)

                        if isfloat(line[58-chinese_count:62-chinese_count]):
                            IKI = getfloatvalue(line[58-chinese_count:62-chinese_count],3)

                        if isfloat(line[62-chinese_count:67-chinese_count]):
                            VFMAX = getfloatvalue(line[62-chinese_count:67-chinese_count],3)

                        if isfloat(line[67-chinese_count:72-chinese_count]):
                            VFMIN = getfloatvalue(line[67-chinese_count:72-chinese_count],3)
                                
                        if isfloat(line[72-chinese_count:76-chinese_count]):
                            KT = getfloatvalue(line[72-chinese_count:76-chinese_count],3)

                        if isfloat(line[76-chinese_count:80-chinese_count]):
                            TT = getfloatvalue(line[76-chinese_count:80-chinese_count],4)
                            

                        avr_type.Tr = TR
                        avr_type.KA = KA
                        avr_type.TA = TA
                        avr_type.Kp = KP
                        avr_type.Ki = KI
                        avr_type.IKP = IKP
                        avr_type.IKI = IKI
                        avr_type.Kt = KT
                        avr_type.Tt = TT
                        avr_type.VRMIN = VRMIN
                        avr_type.VFMIN = VFMIN
                        avr_type.VRMAX = VRMAX
                        avr_type.VFMAX = VFMAX
                        avr_type.EFDMAX = EFDMAX
                        avr_type.EFDMIN = EFDMIN

                                
    #FZ卡
    for line in bpa_str_ar:
        line = line.rstrip('\n')
        if line == "": continue        

        if line[0] == 'F' and line[1] == 'Z':
            line = line + ' '*(80-len(line))
            chinese_count = 0
            #判断中文的个数
            for i in range (3,10):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count = chinese_count + 1
            name = line[3:11-chinese_count].strip()
            if isfloat(line[11-chinese_count:15-chinese_count]):base = float(line[11-chinese_count:15-chinese_count])
            #名称+电压
            Generator = name + str(base)                
            #找到节点名对应的编号
            if Generator in bus_name_ar:
                Generator_num = bus_name_ar[Generator]
                #app.PrintInfo(Generator_num)

                #找到对应的comp模块
                Generator_comp_name = str(Generator_num) + '_comp'
                Generator_compp = Net.SearchObject(str(Generator_comp_name))
                Generator_comp = Generator_compp[0]

                if not Generator_comp == None:
                    avr_type = Generator_comp.Avr
                    avr_type_id = avr_type.typ_id

                    #FF
                    if avr_type_id != None and avr_type_id.loc_name == avr_FF.loc_name:
                        KC = 0
                        KD = 0
                        EFDN = 0 #?????????作用
                        SE1 = 0
                        SE2 = 0
                        KH = 0
                        KF = 0
                        TF = 0
                        VLR = 0
                        KL = 0                        
#??????????????????????????????????????????????????????????????????? 1 or 0.75                        
                        if isfloat(line[16-chinese_count:21-chinese_count]):#1
                            SE1 = getfloatvalue(line[16-chinese_count:21-chinese_count],3)                             

                        if isfloat(line[21-chinese_count:26-chinese_count]):#0.75
                            SE2 = getfloatvalue(line[21-chinese_count:26-chinese_count],3)

                        if isfloat(line[26-chinese_count:31-chinese_count]):
                            EFDN = getfloatvalue(line[26-chinese_count:31-chinese_count],3) #标幺值
                            
                        if isfloat(line[31-chinese_count:36-chinese_count]):
                            VE1 = getfloatvalue(line[31-chinese_count:36-chinese_count],3) #标幺值

                        if isfloat(line[36-chinese_count:41-chinese_count]):
                            KF = getfloatvalue(line[36-chinese_count:41-chinese_count],3)
                            
                        if isfloat(line[41-chinese_count:46-chinese_count]):
                            TF = getfloatvalue(line[41-chinese_count:46-chinese_count],3) 
    
                        if isfloat(line[46-chinese_count:51-chinese_count]):
                            KC = getfloatvalue(line[46-chinese_count:51-chinese_count],4)
                            
                        if isfloat(line[51-chinese_count:56-chinese_count]):
                            KD = getfloatvalue(line[51-chinese_count:56-chinese_count],4) 

                        if isfloat(line[56-chinese_count:61-chinese_count]):
                            KB = getfloatvalue(line[56-chinese_count:61-chinese_count],3) 

                        if isfloat(line[61-chinese_count:66-chinese_count]):
                            KL = getfloatvalue(line[61-chinese_count:66-chinese_count],3) 
                            
                        if isfloat(line[66-chinese_count:71-chinese_count]):
                            KH = getfloatvalue(line[66-chinese_count:71-chinese_count],3) 

                        if isfloat(line[71-chinese_count:76-chinese_count]):
                            VLR = getfloatvalue(line[71-chinese_count:76-chinese_count],3)

                        #N = math.log(SE1/SE2)/(abs(VE1-VE1*0.75))
                        #M = SE1/(math.exp(abs(N*VE1)))

                        avr_type.Kb = KB
                        avr_type.Kl = KL
                        avr_type.Kh = KH
                        avr_type.Kf = KF
                        avr_type.Tf = TF
                        avr_type.Kc = KC
                        avr_type.Kd = KD
                        avr_type.E1 = VE1
                        avr_type.Se1 = SE1
                        avr_type.E2= VE1*0.75
                        avr_type.Se2 = SE2
                        avr_type.Vlr = VLR                                                

                    #FG
                    if avr_type_id != None and avr_type_id.loc_name == avr_FG.loc_name:
                        KC = 0
#?????????????????????????????????????????????????????  KF作用？？
                        KF = 0                                           

                        if isfloat(line[36-chinese_count:41-chinese_count]):
                            KF = getfloatvalue(line[36-chinese_count:41-chinese_count],3)
                            
                        if isfloat(line[46-chinese_count:51-chinese_count]):
                            KC = getfloatvalue(line[46-chinese_count:51-chinese_count],4)
                            
                        avr_type.params[5] = KC
              
                    #FJ
                    if avr_type_id != None and avr_type_id.loc_name == avr_FJ.loc_name:
                        KC = 0
                        EFDMAX = 9999
                        EFDMIN = -9999
                        KF = 0
                        TF = 0

                        
                        if isfloat(line[26-chinese_count:31-chinese_count]):
                            EFDMIN = getfloatvalue(line[26-chinese_count:31-chinese_count],3)
                            EFDMIN = EFDMIN
                            
                        if isfloat(line[31-chinese_count:36-chinese_count]):
                            EFDMAX = getfloatvalue(line[31-chinese_count:36-chinese_count],3) 
                            EFDMAX = EFDMAX

                        if isfloat(line[36-chinese_count:41-chinese_count]):
                            KF = getfloatvalue(line[36-chinese_count:41-chinese_count],3)
                            
                        if isfloat(line[41-chinese_count:46-chinese_count]):
                            TF = getfloatvalue(line[41-chinese_count:46-chinese_count],3) 
    
                        if isfloat(line[46-chinese_count:51-chinese_count]):
                            KC = getfloatvalue(line[46-chinese_count:51-chinese_count],4)
                            
                        avr_type.Kc = KC
                        avr_type.Kf = KF
                        avr_type.Tf = TF
                        avr_type.EFDmin = EFDMIN
                        avr_type.EFDmax = EFDMAX

                    #FK
                    if avr_type_id != None and avr_type_id.loc_name == avr_FK.loc_name:
                        KC = 0
                        KF = 0
                        TF = 0
                        
                        if isfloat(line[36-chinese_count:41-chinese_count]):
                            KF = getfloatvalue(line[36-chinese_count:41-chinese_count],3)
                            
                        if isfloat(line[41-chinese_count:46-chinese_count]):
                            TF = getfloatvalue(line[41-chinese_count:46-chinese_count],3) 
    
                        if isfloat(line[46-chinese_count:51-chinese_count]):
                            KC = getfloatvalue(line[46-chinese_count:51-chinese_count],4)
                            
                        avr_type.Kc = KC
                        avr_type.Kf = KF
                        avr_type.Tf = TF


#FZ可代替F+
                    #FM/FQ
                    if avr_type_id != None and (avr_type_id.loc_name == avr_FM_0.loc_name or avr_type_id.loc_name == avr_FM_1.loc_name):

                        KB = 0
                        T5 = 0
                        TE = 0
                        KL1 = 0
                        KD = 0
                        KC = 0
                        SE2 = 0
                        KE = 0
                        VL1R = 0
                        VRMIN = -9999
                        VA1MIN = -9999
                        VAMIN = -9999
                        VRMAX = 9999
                        EFDMAX = 9999
                        VA1MAX = 9999
                        VAMAX = 9999
                        M = 0
                        N = 0

                        if isfloat(line[16-chinese_count:21-chinese_count]):
                            VAMAX = getfloatvalue(line[16-chinese_count:21-chinese_count],3)

                        if isfloat(line[21-chinese_count:26-chinese_count]):
                            VAMIN = getfloatvalue(line[21-chinese_count:26-chinese_count],3)

                        if isfloat(line[26-chinese_count:30-chinese_count]):
                            KB = getfloatvalue(line[26-chinese_count:30-chinese_count],2)

                        if isfloat(line[31-chinese_count:34-chinese_count]):
                            T5 = getfloatvalue(line[30-chinese_count:34-chinese_count],2)

                        if isfloat(line[34-chinese_count:38-chinese_count]):
                            KE = getfloatvalue(line[34-chinese_count:38-chinese_count],2)

                        if isfloat(line[38-chinese_count:42-chinese_count]):
                            TE = getfloatvalue(line[38-chinese_count:42-chinese_count],2)
                            
                        if isfloat(line[42-chinese_count:47-chinese_count]):
                            SEMAX = getfloatvalue(line[42-chinese_count:47-chinese_count],4)

                        if isfloat(line[47-chinese_count:52-chinese_count]):
                            SE75MAX = getfloatvalue(line[47-chinese_count:52-chinese_count],4)

                        if isfloat(line[52-chinese_count:56-chinese_count]):
                            VRMAX = getfloatvalue(line[52-chinese_count:56-chinese_count],2)

                        if isfloat(line[56-chinese_count:60-chinese_count]):
                            VRMIN = getfloatvalue(line[56-chinese_count:60-chinese_count],2)
                            if VRMIN > 0:
                                VRMIN = VRMIN*(-1)

                        if isfloat(line[60-chinese_count:64-chinese_count]):
                            KC = getfloatvalue(line[60-chinese_count:64-chinese_count],2)

                        if isfloat(line[64-chinese_count:68-chinese_count]):
                            KD = getfloatvalue(line[64-chinese_count:68-chinese_count],2)

                        if isfloat(line[68-chinese_count:72-chinese_count]):
                            KL1 = getfloatvalue(line[68-chinese_count:72-chinese_count],2)

                        if isfloat(line[72-chinese_count:76-chinese_count]):
                            VL1R = getfloatvalue(line[72-chinese_count:76-chinese_count],2)

                        if isfloat(line[76-chinese_count:80-chinese_count]):
                            EFDMAX = getfloatvalue(line[76-chinese_count:80-chinese_count],2)

                        avr_type.Kb = KB
                        avr_type.T5 = T5
                        avr_type.Te = TE
                        avr_type.Kl = KL1
                        avr_type.Kc = KC
                        avr_type.Kd = KD
                        avr_type.E1 = EFDMAX
                        avr_type.Se1 = SEMAX
                        avr_type.E2 = EFDMAX*0.75
                        avr_type.Se2 = SE75MAX
                        avr_type.Vlr = VL1R
                        avr_type.Ke = KE
                        avr_type.Vamin = VAMIN
                        avr_type.Vrmin = abs(VRMIN)*(-1)
                        avr_type.VA1min = VA1MIN
                        avr_type.Vamax = VAMAX
                        avr_type.Vrmax = VRMAX
                        avr_type.EFDMAX = EFDMAX
                        avr_type.VA1max = VA1MAX
                        

                    #FN/FR
                    if avr_type_id != None and (avr_type_id.loc_name == avr_FN_0.loc_name or avr_type_id.loc_name == avr_FN_1.loc_name):

                        KB = 0
                        T5 = 0
                        TE = 0
                        KL1 = 0
                        KD = 0
                        KC = 0
                        SE2 = 0
                        KE = 0
                        VL1R = 0
                        VRMIN = -9999
                        VA1MIN = -9999
                        VAMIN = -9999
                        VRMAX = 9999
                        EFDMAX = 9999
                        VA1MAX = 9999
                        VAMAX = 9999
                        M = 0
                        N = 0

                        if isfloat(line[16-chinese_count:21-chinese_count]):
                            VAMAX = getfloatvalue(line[16-chinese_count:21-chinese_count],3)

                        if isfloat(line[21-chinese_count:26-chinese_count]):
                            VAMIN = getfloatvalue(line[21-chinese_count:26-chinese_count],3)

                        if isfloat(line[26-chinese_count:30-chinese_count]):
                            KB = getfloatvalue(line[26-chinese_count:30-chinese_count],2)

                        if isfloat(line[31-chinese_count:34-chinese_count]):
                            T5 = getfloatvalue(line[30-chinese_count:34-chinese_count],2)

                        if isfloat(line[34-chinese_count:38-chinese_count]):
                            KE = getfloatvalue(line[34-chinese_count:38-chinese_count],2)

                        if isfloat(line[38-chinese_count:42-chinese_count]):
                            TE = getfloatvalue(line[38-chinese_count:42-chinese_count],2)
                            
                        if isfloat(line[42-chinese_count:47-chinese_count]):
                            SEMAX = getfloatvalue(line[42-chinese_count:47-chinese_count],4)

                        if isfloat(line[47-chinese_count:52-chinese_count]):
                            SE75MAX = getfloatvalue(line[47-chinese_count:52-chinese_count],4)

                        if isfloat(line[52-chinese_count:56-chinese_count]):
                            VRMAX = getfloatvalue(line[52-chinese_count:56-chinese_count],2)

                        if isfloat(line[56-chinese_count:60-chinese_count]):
                            VRMIN = getfloatvalue(line[56-chinese_count:60-chinese_count],2)
                            if VRMIN > 0:
                                VRMIN = VRMIN*(-1)

                        if isfloat(line[60-chinese_count:64-chinese_count]):
                            KC = getfloatvalue(line[60-chinese_count:64-chinese_count],2)

                        if isfloat(line[64-chinese_count:68-chinese_count]):
                            KD = getfloatvalue(line[64-chinese_count:68-chinese_count],2)

                        if isfloat(line[68-chinese_count:72-chinese_count]):
                            KL1 = getfloatvalue(line[68-chinese_count:72-chinese_count],2)

                        if isfloat(line[72-chinese_count:76-chinese_count]):
                            VL1R = getfloatvalue(line[72-chinese_count:76-chinese_count],2)

                        if isfloat(line[76-chinese_count:80-chinese_count]):
                            EFDMAX = getfloatvalue(line[76-chinese_count:80-chinese_count],2)

                        avr_type.Kb = KB
                        avr_type.T5 = T5
                        avr_type.Te = TE
                        avr_type.Kl = KL1
                        avr_type.Kc = KC
                        avr_type.Kd = KD
                        avr_type.E1 = EFDMAX
                        avr_type.Se1 = SEMAX
                        avr_type.E2 = EFDMAX*0.75
                        avr_type.Se2 = SE75MAX
                        avr_type.Vlr = VL1R
                        avr_type.Ke = KE
                        avr_type.Vamin = VAMIN
                        avr_type.Vrmin = abs(VRMIN)*(-1)
                        avr_type.VA1min = VA1MIN
                        avr_type.Vamax = VAMAX
                        avr_type.Vrmax = VRMAX
                        avr_type.EFDMAX = EFDMAX
                        avr_type.VA1max = VA1MAX

                    #FO/FS
                    if avr_type_id != None and (avr_type_id.loc_name == avr_FO_0.loc_name or avr_type_id.loc_name == avr_FO_1.loc_name):

                        KB = 0
                        T5 = 0
                        TE = 0
                        KL1 = 0
                        KD = 0
                        KC = 0
                        SE2 = 0
                        KE = 0
                        VL1R = 0
                        VRMIN = -9999
                        VA1MIN = -9999
                        VAMIN = -9999
                        VRMAX = 9999
                        EFDMAX = 9999
                        VA1MAX = 9999
                        VAMAX = 9999
                        M = 0
                        N = 0

                        if isfloat(line[16-chinese_count:21-chinese_count]):
                            VAMAX = getfloatvalue(line[16-chinese_count:21-chinese_count],3)

                        if isfloat(line[21-chinese_count:26-chinese_count]):
                            VAMIN = getfloatvalue(line[21-chinese_count:26-chinese_count],3)

                        if isfloat(line[26-chinese_count:30-chinese_count]):
                            KB = getfloatvalue(line[26-chinese_count:30-chinese_count],2)

                        if isfloat(line[31-chinese_count:34-chinese_count]):
                            T5 = getfloatvalue(line[30-chinese_count:34-chinese_count],2)

                        if isfloat(line[34-chinese_count:38-chinese_count]):
                            KE = getfloatvalue(line[34-chinese_count:38-chinese_count],2)

                        if isfloat(line[38-chinese_count:42-chinese_count]):
                            TE = getfloatvalue(line[38-chinese_count:42-chinese_count],2)
                            
                        if isfloat(line[42-chinese_count:47-chinese_count]):
                            SEMAX = getfloatvalue(line[42-chinese_count:47-chinese_count],4)

                        if isfloat(line[47-chinese_count:52-chinese_count]):
                            SE75MAX = getfloatvalue(line[47-chinese_count:52-chinese_count],4)

                        if isfloat(line[52-chinese_count:56-chinese_count]):
                            VRMAX = getfloatvalue(line[52-chinese_count:56-chinese_count],2)

                        if isfloat(line[56-chinese_count:60-chinese_count]):
                            VRMIN = getfloatvalue(line[56-chinese_count:60-chinese_count],2)
                            if VRMIN > 0:
                                VRMIN = VRMIN*(-1)

                        if isfloat(line[60-chinese_count:64-chinese_count]):
                            KC = getfloatvalue(line[60-chinese_count:64-chinese_count],2)

                        if isfloat(line[64-chinese_count:68-chinese_count]):
                            KD = getfloatvalue(line[64-chinese_count:68-chinese_count],2)

                        if isfloat(line[68-chinese_count:72-chinese_count]):
                            KL1 = getfloatvalue(line[68-chinese_count:72-chinese_count],2)

                        if isfloat(line[72-chinese_count:76-chinese_count]):
                            VL1R = getfloatvalue(line[72-chinese_count:76-chinese_count],2)

                        if isfloat(line[76-chinese_count:80-chinese_count]):
                            EFDMAX = getfloatvalue(line[76-chinese_count:80-chinese_count],2)

                        app.PrintInfo(VRMAX)
                        app.PrintInfo(VRMIN)
                        
                        
                        avr_type.Kb = KB
                        avr_type.T5 = T5
                        avr_type.Te = TE
                        avr_type.Kl = KL1
                        avr_type.Kc = KC
                        avr_type.Kd = KD
                        avr_type.E1 = EFDMAX
                        avr_type.Se1 = SEMAX
                        avr_type.E2 = EFDMAX*0.75
                        avr_type.Se2 = SE75MAX
                        avr_type.Vlr = VL1R
                        avr_type.Ke = KE
                        avr_type.Vamin = VAMIN
                        avr_type.Vrmin = abs(VRMIN)*(-1)
                        avr_type.VA1min = VA1MIN
                        avr_type.Vamax = VAMAX
                        avr_type.Vrmax = VRMAX
                        avr_type.EFDMAX = EFDMAX
                        avr_type.VA1max = VA1MAX

                        

                    #FP/FT
                    if avr_type_id != None and (avr_type_id.loc_name == avr_FP_0.loc_name or avr_type_id.loc_name == avr_FP_1.loc_name):

                        KB = 0
                        T5 = 0
                        TE = 0
                        KL1 = 0
                        KD = 0
                        KC = 0
                        SE2 = 0
                        KE = 0
                        VL1R = 0
                        VRMIN = -9999
                        VA1MIN = -9999
                        VAMIN = -9999
                        VRMAX = 9999
                        EFDMAX = 9999
                        VA1MAX = 9999
                        VAMAX = 9999
                        M = 0
                        N = 0

                        if isfloat(line[16-chinese_count:21-chinese_count]):
                            VAMAX = getfloatvalue(line[16-chinese_count:21-chinese_count],3)

                        if isfloat(line[21-chinese_count:26-chinese_count]):
                            VAMIN = getfloatvalue(line[21-chinese_count:26-chinese_count],3)

                        if isfloat(line[26-chinese_count:30-chinese_count]):
                            KB = getfloatvalue(line[26-chinese_count:30-chinese_count],2)

                        if isfloat(line[31-chinese_count:34-chinese_count]):
                            T5 = getfloatvalue(line[30-chinese_count:34-chinese_count],2)

                        if isfloat(line[34-chinese_count:38-chinese_count]):
                            KE = getfloatvalue(line[34-chinese_count:38-chinese_count],2)

                        if isfloat(line[38-chinese_count:42-chinese_count]):
                            TE = getfloatvalue(line[38-chinese_count:42-chinese_count],2)
                            
                        if isfloat(line[42-chinese_count:47-chinese_count]):
                            SEMAX = getfloatvalue(line[42-chinese_count:47-chinese_count],4)

                        if isfloat(line[47-chinese_count:52-chinese_count]):
                            SE75MAX = getfloatvalue(line[47-chinese_count:52-chinese_count],4)

                        if isfloat(line[52-chinese_count:56-chinese_count]):
                            VRMAX = getfloatvalue(line[52-chinese_count:56-chinese_count],2)

                        if isfloat(line[56-chinese_count:60-chinese_count]):
                            VRMIN = getfloatvalue(line[56-chinese_count:60-chinese_count],2)
                            if VRMIN > 0:
                                VRMIN = VRMIN*(-1)

                        if isfloat(line[60-chinese_count:64-chinese_count]):
                            KC = getfloatvalue(line[60-chinese_count:64-chinese_count],2)

                        if isfloat(line[64-chinese_count:68-chinese_count]):
                            KD = getfloatvalue(line[64-chinese_count:68-chinese_count],2)

                        if isfloat(line[68-chinese_count:72-chinese_count]):
                            KL1 = getfloatvalue(line[68-chinese_count:72-chinese_count],2)

                        if isfloat(line[72-chinese_count:76-chinese_count]):
                            VL1R = getfloatvalue(line[72-chinese_count:76-chinese_count],2)

                        if isfloat(line[76-chinese_count:80-chinese_count]):
                            EFDMAX = getfloatvalue(line[76-chinese_count:80-chinese_count],2)

                        avr_type.Kb = KB
                        avr_type.T5 = T5
                        avr_type.Te = TE
                        avr_type.Kl = KL1
                        avr_type.Kc = KC
                        avr_type.Kd = KD
                        avr_type.E1 = EFDMAX
                        avr_type.Se1 = SEMAX
                        avr_type.E2 = EFDMAX*0.75
                        avr_type.Se2 = SE75MAX
                        avr_type.Vlr = VL1R
                        avr_type.Ke = KE
                        avr_type.Vamin = VAMIN
                        avr_type.Vrmin = abs(VRMIN)*(-1)
                        avr_type.VA1min = VA1MIN
                        avr_type.Vamax = VAMAX
                        avr_type.Vrmax = VRMAX
                        avr_type.EFDMAX = EFDMAX
                        avr_type.VA1max = VA1MAX
                    #FU
                    if avr_type_id != None and (avr_type_id.loc_name == avr_FU_0.loc_name or avr_type_id.loc_name == avr_FU_1.loc_name):

                        KB = 0
                        T5 = 0
                        KC = 0
                        VRMIN = -9999
                        VA1MIN = -9999
                        VAMIN = -9999
                        VRMAX = 9999
                        VA1MAX = 9999
                        VAMAX = 9999

                        if isfloat(line[16-chinese_count:21-chinese_count]):
                            VAMAX = getfloatvalue(line[16-chinese_count:21-chinese_count],3)

                        if isfloat(line[21-chinese_count:26-chinese_count]):
                            VAMIN = getfloatvalue(line[21-chinese_count:26-chinese_count],3)

                        if isfloat(line[52-chinese_count:56-chinese_count]):
                            VRMAX = getfloatvalue(line[52-chinese_count:56-chinese_count],2)

                        if isfloat(line[56-chinese_count:60-chinese_count]):
                            VRMIN = getfloatvalue(line[56-chinese_count:60-chinese_count],2)
                            if VRMIN > 0:
                                VRMIN = VRMIN*(-1)

                        if isfloat(line[60-chinese_count:64-chinese_count]):
                            KC = getfloatvalue(line[60-chinese_count:64-chinese_count],2)
                         

                        avr_type.Kc = KC
                        avr_type.VAmin = VAMIN
                        avr_type.VA1min = VA1MIN
                        avr_type.Vrmin = VRMIN
                        avr_type.VAmax = VAMAX
                        avr_type.VA1max = VA1MAX
                        avr_type.Vrmax = VRMAX                        
                        
                    #FV
                    if avr_type_id != None and (avr_type_id.loc_name == avr_FV_0.loc_name or avr_type_id.loc_name == avr_FV_1.loc_name):

                        KB = 0
                        T5 = 0
                        KC = 0
                        VRMIN = -9999
                        VA1MIN = -9999
                        VAMIN = -9999
                        VRMAX = 9999
                        VA1MAX = 9999
                        VAMAX = 9999

                        if isfloat(line[16-chinese_count:21-chinese_count]):
                            VAMAX = getfloatvalue(line[16-chinese_count:21-chinese_count],3)

                        if isfloat(line[21-chinese_count:26-chinese_count]):
                            VAMIN = getfloatvalue(line[21-chinese_count:26-chinese_count],3)

                        if isfloat(line[52-chinese_count:56-chinese_count]):
                            VRMAX = getfloatvalue(line[52-chinese_count:56-chinese_count],2)

                        if isfloat(line[56-chinese_count:60-chinese_count]):
                            VRMIN = getfloatvalue(line[56-chinese_count:60-chinese_count],2)
                            if VRMIN > 0:
                                VRMIN = VRMIN*(-1)

                        if isfloat(line[60-chinese_count:64-chinese_count]):
                            KC = getfloatvalue(line[60-chinese_count:64-chinese_count],2)
                         

                        avr_type.Kc = KC
                        avr_type.VAmin = VAMIN
                        avr_type.VA1min = VA1MIN
                        avr_type.Vrmin = VRMIN
                        avr_type.VAmax = VAMAX
                        avr_type.VA1max = VA1MAX
                        avr_type.Vrmax = VRMAX      

                    #FX+ F+和FX配合使用时，相当于FX+
                    if avr_type_id != None and avr_type_id.loc_name == avr_FX.loc_name:

                        KIFD = 0
                        TIFD = 0
                        EFDMAX = 9999
                        EFDMIN = -9999

                        if isfloat(line[26-chinese_count:30-chinese_count]):
                            KIFD = getfloatvalue(line[26-chinese_count:30-chinese_count],2)

                        if isfloat(line[30-chinese_count:34-chinese_count]):
                            TIFD = getfloatvalue(line[30-chinese_count:34-chinese_count],2)

                        if isfloat(line[72-chinese_count:76-chinese_count]):
                            EFDMIN = getfloatvalue(line[72-chinese_count:76-chinese_count],2)

                        if isfloat(line[76-chinese_count:80-chinese_count]):
                            EFDMAX = getfloatvalue(line[76-chinese_count:80-chinese_count],2)

                        avr_type.KIFD = KIFD
                        avr_type.TIFD = TIFD
                        avr_type.EFDMIN = EFDMIN
                        avr_type.EFDMAX = EFDMAX



    #F+卡
    for line in bpa_str_ar:
        line = line.rstrip('\n')
        if line == "": continue        

        if line[0] == 'F' and line[1] == '+':
            line = line + ' '*(80-len(line))
            chinese_count = 0
            #判断中文的个数
            for i in range (3,10):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count = chinese_count + 1
            name = line[3:11-chinese_count].strip()
            if isfloat(line[11-chinese_count:15-chinese_count]):base = float(line[11-chinese_count:15-chinese_count])
            #名称+电压
            Generator = name + str(base)                
            #找到节点名对应的编号
            if Generator in bus_name_ar:
                Generator_num = bus_name_ar[Generator]
                #app.PrintInfo(Generator_num)

                #找到对应的comp模块
                Generator_comp_name = str(Generator_num) + '_comp'
                Generator_compp = Net.SearchObject(str(Generator_comp_name))
                Generator_comp = Generator_compp[0]

                if not Generator_comp == None:
                    avr_type = Generator_comp.Avr
                    avr_type_id = avr_type.typ_id

                    #FM/FQ
                    if avr_type_id != None and (avr_type_id.loc_name == avr_FM_0.loc_name or avr_type_id.loc_name == avr_FM_1.loc_name):

                        KB = 0
                        T5 = 0
                        TE = 0
                        KL1 = 0
                        KD = 0
                        KC = 0
                        SE2 = 0
                        KE = 0
                        VL1R = 0
                        VRMIN = -9999
                        VA1MIN = -9999
                        VAMIN = -9999
                        VRMAX = 9999
                        EFDMAX = 9999
                        VA1MAX = 9999
                        VAMAX = 9999
                        M = 0
                        N = 0

                        if isfloat(line[16-chinese_count:21-chinese_count]):
                            VAMAX = getfloatvalue(line[16-chinese_count:21-chinese_count],3)

                        if isfloat(line[21-chinese_count:26-chinese_count]):
                            VAMIN = getfloatvalue(line[21-chinese_count:26-chinese_count],3)

                        if isfloat(line[26-chinese_count:30-chinese_count]):
                            KB = getfloatvalue(line[26-chinese_count:30-chinese_count],2)

                        if isfloat(line[31-chinese_count:34-chinese_count]):
                            T5 = getfloatvalue(line[30-chinese_count:34-chinese_count],2)

                        if isfloat(line[34-chinese_count:38-chinese_count]):
                            KE = getfloatvalue(line[34-chinese_count:38-chinese_count],2)

                        if isfloat(line[38-chinese_count:42-chinese_count]):
                            TE = getfloatvalue(line[38-chinese_count:42-chinese_count],2)
                            
                        if isfloat(line[42-chinese_count:47-chinese_count]):
                            SEMAX = getfloatvalue(line[42-chinese_count:47-chinese_count],4)

                        if isfloat(line[47-chinese_count:52-chinese_count]):
                            SE75MAX = getfloatvalue(line[47-chinese_count:52-chinese_count],4)

                        if isfloat(line[52-chinese_count:56-chinese_count]):
                            VRMAX = getfloatvalue(line[52-chinese_count:56-chinese_count],2)

                        if isfloat(line[56-chinese_count:60-chinese_count]):
                            VRMIN = getfloatvalue(line[56-chinese_count:60-chinese_count],2)
                            if VRMIN > 0:
                                VRMIN = VRMIN*(-1)

                        if isfloat(line[60-chinese_count:64-chinese_count]):
                            KC = getfloatvalue(line[60-chinese_count:64-chinese_count],2)

                        if isfloat(line[64-chinese_count:68-chinese_count]):
                            KD = getfloatvalue(line[64-chinese_count:68-chinese_count],2)

                        if isfloat(line[68-chinese_count:72-chinese_count]):
                            KL1 = getfloatvalue(line[68-chinese_count:72-chinese_count],2)

                        if isfloat(line[72-chinese_count:76-chinese_count]):
                            VL1R = getfloatvalue(line[72-chinese_count:76-chinese_count],2)

                        if isfloat(line[76-chinese_count:80-chinese_count]):
                            EFDMAX = getfloatvalue(line[76-chinese_count:80-chinese_count],2)

                        avr_type.Kb = KB
                        avr_type.T5 = T5
                        avr_type.Te = TE
                        avr_type.Kl = KL1
                        avr_type.Kc = KC
                        avr_type.Kd = KD
                        avr_type.E1 = EFDMAX
                        avr_type.Se1 = SEMAX
                        avr_type.E2 = EFDMAX*0.75
                        avr_type.Se2 = SE75MAX
                        avr_type.Vlr = VL1R
                        avr_type.Ke = KE
                        avr_type.Vamin = VAMIN
                        avr_type.Vrmin = abs(VRMIN)*(-1)
                        avr_type.VA1min = VA1MIN
                        avr_type.Vamax = VAMAX
                        avr_type.Vrmax = VRMAX
                        avr_type.EFDMAX = EFDMAX
                        avr_type.VA1max = VA1MAX
                        

                    #FN/FR
                    if avr_type_id != None and (avr_type_id.loc_name == avr_FN_0.loc_name or avr_type_id.loc_name == avr_FN_1.loc_name):

                        KB = 0
                        T5 = 0
                        TE = 0
                        KL1 = 0
                        KD = 0
                        KC = 0
                        SE2 = 0
                        KE = 0
                        VL1R = 0
                        VRMIN = -9999
                        VA1MIN = -9999
                        VAMIN = -9999
                        VRMAX = 9999
                        EFDMAX = 9999
                        VA1MAX = 9999
                        VAMAX = 9999
                        M = 0
                        N = 0

                        if isfloat(line[16-chinese_count:21-chinese_count]):
                            VAMAX = getfloatvalue(line[16-chinese_count:21-chinese_count],3)

                        if isfloat(line[21-chinese_count:26-chinese_count]):
                            VAMIN = getfloatvalue(line[21-chinese_count:26-chinese_count],3)

                        if isfloat(line[26-chinese_count:30-chinese_count]):
                            KB = getfloatvalue(line[26-chinese_count:30-chinese_count],2)

                        if isfloat(line[31-chinese_count:34-chinese_count]):
                            T5 = getfloatvalue(line[30-chinese_count:34-chinese_count],2)

                        if isfloat(line[34-chinese_count:38-chinese_count]):
                            KE = getfloatvalue(line[34-chinese_count:38-chinese_count],2)

                        if isfloat(line[38-chinese_count:42-chinese_count]):
                            TE = getfloatvalue(line[38-chinese_count:42-chinese_count],2)
                            
                        if isfloat(line[42-chinese_count:47-chinese_count]):
                            SEMAX = getfloatvalue(line[42-chinese_count:47-chinese_count],4)

                        if isfloat(line[47-chinese_count:52-chinese_count]):
                            SE75MAX = getfloatvalue(line[47-chinese_count:52-chinese_count],4)

                        if isfloat(line[52-chinese_count:56-chinese_count]):
                            VRMAX = getfloatvalue(line[52-chinese_count:56-chinese_count],2)

                        if isfloat(line[56-chinese_count:60-chinese_count]):
                            VRMIN = getfloatvalue(line[56-chinese_count:60-chinese_count],2)
                            if VRMIN > 0:
                                VRMIN = VRMIN*(-1)

                        if isfloat(line[60-chinese_count:64-chinese_count]):
                            KC = getfloatvalue(line[60-chinese_count:64-chinese_count],2)

                        if isfloat(line[64-chinese_count:68-chinese_count]):
                            KD = getfloatvalue(line[64-chinese_count:68-chinese_count],2)

                        if isfloat(line[68-chinese_count:72-chinese_count]):
                            KL1 = getfloatvalue(line[68-chinese_count:72-chinese_count],2)

                        if isfloat(line[72-chinese_count:76-chinese_count]):
                            VL1R = getfloatvalue(line[72-chinese_count:76-chinese_count],2)

                        if isfloat(line[76-chinese_count:80-chinese_count]):
                            EFDMAX = getfloatvalue(line[76-chinese_count:80-chinese_count],2)

                        avr_type.Kb = KB
                        avr_type.T5 = T5
                        avr_type.Te = TE
                        avr_type.Kl = KL1
                        avr_type.Kc = KC
                        avr_type.Kd = KD
                        avr_type.E1 = EFDMAX
                        avr_type.Se1 = SEMAX
                        avr_type.E2 = EFDMAX*0.75
                        avr_type.Se2 = SE75MAX
                        avr_type.Vlr = VL1R
                        avr_type.Ke = KE
                        avr_type.Vamin = VAMIN
                        avr_type.Vrmin = abs(VRMIN)*(-1)
                        avr_type.VA1min = VA1MIN
                        avr_type.Vamax = VAMAX
                        avr_type.Vrmax = VRMAX
                        avr_type.EFDMAX = EFDMAX
                        avr_type.VA1max = VA1MAX

                    #FO/FS
                    if avr_type_id != None and (avr_type_id.loc_name == avr_FO_0.loc_name or avr_type_id.loc_name == avr_FO_1.loc_name):

                        KB = 0
                        T5 = 0
                        TE = 0
                        KL1 = 0
                        KD = 0
                        KC = 0
                        SE2 = 0
                        KE = 0
                        VL1R = 0
                        VRMIN = -9999
                        VA1MIN = -9999
                        VAMIN = -9999
                        VRMAX = 9999
                        EFDMAX = 9999
                        VA1MAX = 9999
                        VAMAX = 9999
                        M = 0
                        N = 0

                        if isfloat(line[16-chinese_count:21-chinese_count]):
                            VAMAX = getfloatvalue(line[16-chinese_count:21-chinese_count],3)

                        if isfloat(line[21-chinese_count:26-chinese_count]):
                            VAMIN = getfloatvalue(line[21-chinese_count:26-chinese_count],3)

                        if isfloat(line[26-chinese_count:30-chinese_count]):
                            KB = getfloatvalue(line[26-chinese_count:30-chinese_count],2)

                        if isfloat(line[31-chinese_count:34-chinese_count]):
                            T5 = getfloatvalue(line[30-chinese_count:34-chinese_count],2)

                        if isfloat(line[34-chinese_count:38-chinese_count]):
                            KE = getfloatvalue(line[34-chinese_count:38-chinese_count],2)

                        if isfloat(line[38-chinese_count:42-chinese_count]):
                            TE = getfloatvalue(line[38-chinese_count:42-chinese_count],2)
                            
                        if isfloat(line[42-chinese_count:47-chinese_count]):
                            SEMAX = getfloatvalue(line[42-chinese_count:47-chinese_count],4)

                        if isfloat(line[47-chinese_count:52-chinese_count]):
                            SE75MAX = getfloatvalue(line[47-chinese_count:52-chinese_count],4)

                        if isfloat(line[52-chinese_count:56-chinese_count]):
                            VRMAX = getfloatvalue(line[52-chinese_count:56-chinese_count],2)

                        if isfloat(line[56-chinese_count:60-chinese_count]):
                            VRMIN = getfloatvalue(line[56-chinese_count:60-chinese_count],2)
                            if VRMIN > 0:
                                VRMIN = VRMIN*(-1)

                        if isfloat(line[60-chinese_count:64-chinese_count]):
                            KC = getfloatvalue(line[60-chinese_count:64-chinese_count],2)

                        if isfloat(line[64-chinese_count:68-chinese_count]):
                            KD = getfloatvalue(line[64-chinese_count:68-chinese_count],2)

                        if isfloat(line[68-chinese_count:72-chinese_count]):
                            KL1 = getfloatvalue(line[68-chinese_count:72-chinese_count],2)

                        if isfloat(line[72-chinese_count:76-chinese_count]):
                            VL1R = getfloatvalue(line[72-chinese_count:76-chinese_count],2)

                        if isfloat(line[76-chinese_count:80-chinese_count]):
                            EFDMAX = getfloatvalue(line[76-chinese_count:80-chinese_count],2)

                        app.PrintInfo(VRMAX)
                        app.PrintInfo(VRMIN)
                        
                        
                        avr_type.Kb = KB
                        avr_type.T5 = T5
                        avr_type.Te = TE
                        avr_type.Kl = KL1
                        avr_type.Kc = KC
                        avr_type.Kd = KD
                        avr_type.E1 = EFDMAX
                        avr_type.Se1 = SEMAX
                        avr_type.E2 = EFDMAX*0.75
                        avr_type.Se2 = SE75MAX
                        avr_type.Vlr = VL1R
                        avr_type.Ke = KE
                        avr_type.Vamin = VAMIN
                        avr_type.Vrmin = abs(VRMIN)*(-1)
                        avr_type.VA1min = VA1MIN
                        avr_type.Vamax = VAMAX
                        avr_type.Vrmax = VRMAX
                        avr_type.EFDMAX = EFDMAX
                        avr_type.VA1max = VA1MAX

                        

                    #FP/FT
                    if avr_type_id != None and (avr_type_id.loc_name == avr_FP_0.loc_name or avr_type_id.loc_name == avr_FP_1.loc_name):

                        KB = 0
                        T5 = 0
                        TE = 0
                        KL1 = 0
                        KD = 0
                        KC = 0
                        SE2 = 0
                        KE = 0
                        VL1R = 0
                        VRMIN = -9999
                        VA1MIN = -9999
                        VAMIN = -9999
                        VRMAX = 9999
                        EFDMAX = 9999
                        VA1MAX = 9999
                        VAMAX = 9999
                        M = 0
                        N = 0

                        if isfloat(line[16-chinese_count:21-chinese_count]):
                            VAMAX = getfloatvalue(line[16-chinese_count:21-chinese_count],3)

                        if isfloat(line[21-chinese_count:26-chinese_count]):
                            VAMIN = getfloatvalue(line[21-chinese_count:26-chinese_count],3)

                        if isfloat(line[26-chinese_count:30-chinese_count]):
                            KB = getfloatvalue(line[26-chinese_count:30-chinese_count],2)

                        if isfloat(line[31-chinese_count:34-chinese_count]):
                            T5 = getfloatvalue(line[30-chinese_count:34-chinese_count],2)

                        if isfloat(line[34-chinese_count:38-chinese_count]):
                            KE = getfloatvalue(line[34-chinese_count:38-chinese_count],2)

                        if isfloat(line[38-chinese_count:42-chinese_count]):
                            TE = getfloatvalue(line[38-chinese_count:42-chinese_count],2)
                            
                        if isfloat(line[42-chinese_count:47-chinese_count]):
                            SEMAX = getfloatvalue(line[42-chinese_count:47-chinese_count],4)

                        if isfloat(line[47-chinese_count:52-chinese_count]):
                            SE75MAX = getfloatvalue(line[47-chinese_count:52-chinese_count],4)

                        if isfloat(line[52-chinese_count:56-chinese_count]):
                            VRMAX = getfloatvalue(line[52-chinese_count:56-chinese_count],2)

                        if isfloat(line[56-chinese_count:60-chinese_count]):
                            VRMIN = getfloatvalue(line[56-chinese_count:60-chinese_count],2)
                            if VRMIN > 0:
                                VRMIN = VRMIN*(-1)

                        if isfloat(line[60-chinese_count:64-chinese_count]):
                            KC = getfloatvalue(line[60-chinese_count:64-chinese_count],2)

                        if isfloat(line[64-chinese_count:68-chinese_count]):
                            KD = getfloatvalue(line[64-chinese_count:68-chinese_count],2)

                        if isfloat(line[68-chinese_count:72-chinese_count]):
                            KL1 = getfloatvalue(line[68-chinese_count:72-chinese_count],2)

                        if isfloat(line[72-chinese_count:76-chinese_count]):
                            VL1R = getfloatvalue(line[72-chinese_count:76-chinese_count],2)

                        if isfloat(line[76-chinese_count:80-chinese_count]):
                            EFDMAX = getfloatvalue(line[76-chinese_count:80-chinese_count],2)

                        avr_type.Kb = KB
                        avr_type.T5 = T5
                        avr_type.Te = TE
                        avr_type.Kl = KL1
                        avr_type.Kc = KC
                        avr_type.Kd = KD
                        avr_type.E1 = EFDMAX
                        avr_type.Se1 = SEMAX
                        avr_type.E2 = EFDMAX*0.75
                        avr_type.Se2 = SE75MAX
                        avr_type.Vlr = VL1R
                        avr_type.Ke = KE
                        avr_type.Vamin = VAMIN
                        avr_type.Vrmin = abs(VRMIN)*(-1)
                        avr_type.VA1min = VA1MIN
                        avr_type.Vamax = VAMAX
                        avr_type.Vrmax = VRMAX
                        avr_type.EFDMAX = EFDMAX
                        avr_type.VA1max = VA1MAX
                    #FU
                    if avr_type_id != None and (avr_type_id.loc_name == avr_FU_0.loc_name or avr_type_id.loc_name == avr_FU_1.loc_name):

                        KB = 0
                        T5 = 0
                        KC = 0
                        VRMIN = -9999
                        VA1MIN = -9999
                        VAMIN = -9999
                        VRMAX = 9999
                        VA1MAX = 9999
                        VAMAX = 9999

                        if isfloat(line[16-chinese_count:21-chinese_count]):
                            VAMAX = getfloatvalue(line[16-chinese_count:21-chinese_count],3)

                        if isfloat(line[21-chinese_count:26-chinese_count]):
                            VAMIN = getfloatvalue(line[21-chinese_count:26-chinese_count],3)

                        if isfloat(line[52-chinese_count:56-chinese_count]):
                            VRMAX = getfloatvalue(line[52-chinese_count:56-chinese_count],2)

                        if isfloat(line[56-chinese_count:60-chinese_count]):
                            VRMIN = getfloatvalue(line[56-chinese_count:60-chinese_count],2)
                            if VRMIN > 0:
                                VRMIN = VRMIN*(-1)

                        if isfloat(line[60-chinese_count:64-chinese_count]):
                            KC = getfloatvalue(line[60-chinese_count:64-chinese_count],2)
                         

                        avr_type.Kc = KC
                        avr_type.VAmin = VAMIN
                        avr_type.VA1min = VA1MIN
                        avr_type.Vrmin = VRMIN
                        avr_type.VAmax = VAMAX
                        avr_type.VA1max = VA1MAX
                        avr_type.Vrmax = VRMAX                        
                        
                    #FV
                    if avr_type_id != None and (avr_type_id.loc_name == avr_FV_0.loc_name or avr_type_id.loc_name == avr_FV_1.loc_name):

                        KB = 0
                        T5 = 0
                        KC = 0
                        VRMIN = -9999
                        VA1MIN = -9999
                        VAMIN = -9999
                        VRMAX = 9999
                        VA1MAX = 9999
                        VAMAX = 9999

                        if isfloat(line[16-chinese_count:21-chinese_count]):
                            VAMAX = getfloatvalue(line[16-chinese_count:21-chinese_count],3)

                        if isfloat(line[21-chinese_count:26-chinese_count]):
                            VAMIN = getfloatvalue(line[21-chinese_count:26-chinese_count],3)

                        if isfloat(line[52-chinese_count:56-chinese_count]):
                            VRMAX = getfloatvalue(line[52-chinese_count:56-chinese_count],2)

                        if isfloat(line[56-chinese_count:60-chinese_count]):
                            VRMIN = getfloatvalue(line[56-chinese_count:60-chinese_count],2)
                            if VRMIN > 0:
                                VRMIN = VRMIN*(-1)

                        if isfloat(line[60-chinese_count:64-chinese_count]):
                            KC = getfloatvalue(line[60-chinese_count:64-chinese_count],2)
                         

                        avr_type.Kc = KC
                        avr_type.VAmin = VAMIN
                        avr_type.VA1min = VA1MIN
                        avr_type.Vrmin = VRMIN
                        avr_type.VAmax = VAMAX
                        avr_type.VA1max = VA1MAX
                        avr_type.Vrmax = VRMAX      

                    #FX+ F+和FX配合使用时，相当于FX+
                    if avr_type_id != None and avr_type_id.loc_name == avr_FX.loc_name:

                        KIFD = 0
                        TIFD = 0
                        EFDMAX = 9999
                        EFDMIN = -9999

                        if isfloat(line[26-chinese_count:30-chinese_count]):
                            KIFD = getfloatvalue(line[26-chinese_count:30-chinese_count],2)

                        if isfloat(line[30-chinese_count:34-chinese_count]):
                            TIFD = getfloatvalue(line[30-chinese_count:34-chinese_count],2)

                        if isfloat(line[72-chinese_count:76-chinese_count]):
                            EFDMIN = getfloatvalue(line[72-chinese_count:76-chinese_count],2)

                        if isfloat(line[76-chinese_count:80-chinese_count]):
                            EFDMAX = getfloatvalue(line[76-chinese_count:80-chinese_count],2)

                        avr_type.KIFD = KIFD
                        avr_type.TIFD = TIFD
                        avr_type.EFDMIN = EFDMIN
                        avr_type.EFDMAX = EFDMAX


    #FX+卡
    for line in bpa_str_ar:
        line = line.rstrip('\n')
        if line == "": continue        

        if line[0] == 'F' and line[1] == 'X' and line[2] == '+':
            line = line + ' '*(80-len(line))
            chinese_count = 0
            #判断中文的个数
            for i in range (3,10):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count = chinese_count + 1
            name = line[3:11-chinese_count].strip()
            if isfloat(line[11-chinese_count:15-chinese_count]):base = float(line[11-chinese_count:15-chinese_count])
            #名称+电压
            Generator = name + str(base)                
            #找到节点名对应的编号
            if Generator in bus_name_ar:
                Generator_num = bus_name_ar[Generator]
                #app.PrintInfo(Generator_num)

                #找到对应的comp模块
                Generator_comp_name = str(Generator_num) + '_comp'
                Generator_compp = Net.SearchObject(str(Generator_comp_name))
                Generator_comp = Generator_compp[0]

                if not Generator_comp == None:
                    avr_type = Generator_comp.Avr
                    avr_type_id = avr_type.typ_id
                    #FX+
                    if avr_type_id != None and avr_type_id.loc_name == avr_FX.loc_name:

                        KIFD = 0
                        TIFD = 0
                        EFDMAX = 9999
                        EFDMIN = -9999

                        if isfloat(line[26-chinese_count:30-chinese_count]):
                            KIFD = getfloatvalue(line[26-chinese_count:30-chinese_count],2)

                        if isfloat(line[30-chinese_count:34-chinese_count]):
                            TIFD = getfloatvalue(line[30-chinese_count:34-chinese_count],2)

                        if isfloat(line[72-chinese_count:76-chinese_count]):
                            EFDMIN = getfloatvalue(line[72-chinese_count:76-chinese_count],2)

                        if isfloat(line[76-chinese_count:80-chinese_count]):
                            EFDMAX = getfloatvalue(line[76-chinese_count:80-chinese_count],2)

                        avr_type.KIFD = KIFD
                        avr_type.TIFD = TIFD
                        avr_type.EFDMIN = EFDMIN
                        avr_type.EFDMAX = EFDMAX

    #F#卡
    for line in bpa_str_ar:
        line = line.rstrip('\n')
        if line == "": continue        

        #F#卡
        if line[0] == 'F' and line[0] == '#':
            line = line + ' '*(80-len(line))
            chinese_count = 0
            #判断中文的个数
            for i in range (3,10):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count = chinese_count + 1
            name = line[3:11-chinese_count].strip()
            if isfloat(line[11-chinese_count:15-chinese_count]):base = float(line[11-chinese_count:15-chinese_count])
            #名称+电压
            Generator = name + str(base)
            
            #找到节点名对应的编号
            if Generator in bus_name_ar:
                Generator_num = bus_name_ar[Generator]

                #找到对应的comp模块
                Generator_comp_name = str(Generator_num) + '_comp'
                Generator_compp = Net.SearchObject(str(Generator_comp_name))
                Generator_comp = Generator_compp[0]

                if not Generator_comp == None:
                    avr_type = Generator_comp.Avr
                    avr_type_id = avr_type.typ_id

                    if line[17-chinese_count] == 1:
                        app.PrintInfo('Need PID')
                    else:
                        if avr_type_id != None and (avr_type_id.loc_name == avr_FV_0.loc_name or avr_type_id.loc_name == avr_FV_1.loc_name\
                                                or avr_type_id.loc_name == avr_FV_1.loc_name or avr_type_id.loc_name == avr_FV_1.loc_name):
                            if isfloat(line[52-chinese_count:56-chinese_count]):
                                VA1MAX = getfloatvalue(line[58-chinese_count:63-chinese_count],0)

                            if isfloat(line[63-chinese_count:68-chinese_count]):
                                VA1MIN = getfloatvalue(line[63-chinese_count:68-chinese_count],0)

                            avr_type.VA1min = VA1MIN
                            avr_type.VA1max = VA1MAX
                        else:                            
                            if isfloat(line[52-chinese_count:56-chinese_count]):
                                VA1MAX = getfloatvalue(line[58-chinese_count:63-chinese_count],0)

                            if isfloat(line[63-chinese_count:68-chinese_count]):
                                VA1MIN = getfloatvalue(line[63-chinese_count:68-chinese_count],0)

                            avr_type.VA1min = VA1MIN
                            avr_type.VA1max = VA1MAX    


# ----------------------------------------------------------------------------------------------------
def GetSCard(bpa_file, bpa_str_ar,bus_name_ar):
    #PSS S卡
    
    prj = app.GetActiveProject()
    if prj is None:
        raise Exception("No project activated. Python Script stopped.")

    app.PrintInfo(prj)
    Network_Modell = prj.SearchObject("Network Model.IntPrjfolder")
    Network_Model=Network_Modell[0]

    Network_Dataa = Network_Model.SearchObject("Network Data.IntPrjfolder")
    Network_Data = Network_Dataa[0]
    app.PrintInfo(Network_Data)

    Nett = Network_Data.SearchObject("*.ElmNet")
    Net = Nett[0]
    app.PrintInfo(Net)

    Libraryy = prj.SearchObject("Library.IntPrjfolder")
    Library = Libraryy[0]
    app.PrintInfo(Library)

    user_defined_modell = Library.SearchObject("User Defined Models.IntPrjfolder")
    user_defined_model = user_defined_modell[0]
    app.PrintInfo(user_defined_model)

    BPA_Frame_EE = user_defined_model.SearchObject("BPA Frame(E).BlkDef")
    BPA_Frame_E = BPA_Frame_EE[0]
    app.PrintInfo(BPA_Frame_E)

    BPA_Frame_FF = user_defined_model.SearchObject("BPA Frame(F).BlkDef")
    BPA_Frame_F = BPA_Frame_FF[0]
    app.PrintInfo(BPA_Frame_F)

    pss_SPP = user_defined_model.SearchObject("pss_SP.BlkDef")
    pss_SP = pss_SPP[0]
    app.PrintInfo(pss_SP)

    pss_SGG = user_defined_model.SearchObject("pss_SG.BlkDef")
    pss_SG = pss_SGG[0]
    app.PrintInfo(pss_SG)

    pss_SSS = user_defined_model.SearchObject("pss_SS.BlkDef")
    pss_SS = pss_SSS[0]
    app.PrintInfo(pss_SS)

    pss_SH11 = user_defined_model.SearchObject("pss_SH1.BlkDef")
    pss_SH1 = pss_SH11[0]
    app.PrintInfo(pss_SH1) 

    pss_SH22 = user_defined_model.SearchObject("pss_SH2.BlkDef")
    pss_SH2 = pss_SH22[0]
    app.PrintInfo(pss_SH2)     

    pss_SI_00 = user_defined_model.SearchObject("pss_SI_0.BlkDef")
    pss_SI_0 = pss_SI_00[0]
    app.PrintInfo(pss_SI_0)     

    pss_SI_11 = user_defined_model.SearchObject("pss_SI_1.BlkDef")
    pss_SI_1 = pss_SI_11[0]
    app.PrintInfo(pss_SI_1)

    pss_SI_22 = user_defined_model.SearchObject("pss_SI_2.BlkDef")
    pss_SI_2 = pss_SI_22[0]
    app.PrintInfo(pss_SI_2)         

    pss_SAA = user_defined_model.SearchObject("pss_SA.BlkDef")
    pss_SA = pss_SAA[0]
    app.PrintInfo(pss_SA)  

    pss_SB_00 = user_defined_model.SearchObject("pss_SB_0.BlkDef")
    pss_SB_0 = pss_SB_00[0]
    app.PrintInfo(pss_SB_0)     

    pss_SB_11 = user_defined_model.SearchObject("pss_SB_1.BlkDef")
    pss_SB_1 = pss_SB_11[0]
    app.PrintInfo(pss_SB_1)

    pss_SB_22 = user_defined_model.SearchObject("pss_SB_2.BlkDef")
    pss_SB_2 = pss_SB_22[0]
    app.PrintInfo(pss_SB_2)   

    
    for line in bpa_str_ar:
        line = line.rstrip('\n')
        if line == "": continue        

        #S卡
        if line[0] == 'S':
            line = line + ' '*(80-len(line))
            chinese_count = 0
            #判断中文的个数
            for i in range (3,10):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count = chinese_count + 1
            name = line[3:11-chinese_count].strip()
            if isfloat(line[11-chinese_count:15-chinese_count]):base = float(line[11-chinese_count:15-chinese_count])
            #名称+电压
            Generator = name + str(base)                
            #找到节点名对应的编号
            if Generator in bus_name_ar:
                Generator_num = bus_name_ar[Generator]

                #找到对应的comp模块
                Generator_comp_name = str(Generator_num) + '_comp'
                Generator_compp = Net.SearchObject(str(Generator_comp_name))
                Generator_comp = Generator_compp[0]

                if not Generator_comp == None:       
                    #avr.dsl
#                    app.PrintInfo(Generator_num)
                    Generator_pss_dsll = Generator_comp.SearchObject('pss_dsl.ElmDsl')
                    Generator_pss_dsl = Generator_pss_dsll[0]
                    if Generator_pss_dsl == None:
                        Generator_comp.CreateObject('ElmDsl','pss_dsl')
                        Generator_pss_dsll = Generator_comp.SearchObject('pss_dsl.ElmDsl')
                        Generator_pss_dsl = Generator_pss_dsll[0]

#                    app.PrintInfo(Generator_pss_dsl)
                    Generator_comp.Pss= Generator_pss_dsl
                    sym = Generator_comp.Sym

                    sym_typ = sym.typ_id
                    MVABASE = sym_typ.sgn                    

                    #SP
                    if line[1] == 'P':

                        Generator_comp_pss = Generator_comp.Pss
                        Generator_comp_pss.typ_id = pss_SP
                        pss_type = Generator_comp.Pss

                        sym = Generator_comp.Sym

                        sym_typ = sym.typ_id
                        MVABASE = sym_typ.sgn   

#?????????????????????????????????????????????????????? KQS pu
                        if isfloat(line[76-chinese_count:80-chinese_count]):
                            MVABASE1 = getfloatvalue(line[76-chinese_count:80-chinese_count],0)
                        else:
                            MVABASE1 = MVABASE
                        if MVABASE1 == 0:
                            MVABASE1 = MVABASE

                        KQV = 0
                        TQV = 0
                        KQS = 0
                        TQS = 0
                        TQ = 0
                        TQ1_ = 0                            
                        TQ1 = 0
                        TQ2_ = 0
                        TQ2 = 0
                        TQ3_ = 0
                        TQ3 = 0
                        VCUTOFF = 0
                        VSMIN = -9999
                        VSMAX = 9999
                        VSLOW = 0

                        if isfloat(line[16-chinese_count:20-chinese_count]):
                            KQV = getfloatvalue(line[16-chinese_count:20-chinese_count],3)

                        if isfloat(line[20-chinese_count:23-chinese_count]):
                            TQV = getfloatvalue(line[20-chinese_count:23-chinese_count],3)                         
                        
                        if isfloat(line[23-chinese_count:27-chinese_count]):
                            KQS = getfloatvalue(line[23-chinese_count:27-chinese_count],3)

                        if isfloat(line[27-chinese_count:30-chinese_count]):
                            TQS = getfloatvalue(line[27-chinese_count:30-chinese_count],3)

                        if isfloat(line[30-chinese_count:34-chinese_count]):
                            TQ = getfloatvalue(line[30-chinese_count:34-chinese_count],2)

                        if isfloat(line[34-chinese_count:38-chinese_count]):
                            TQ1 = getfloatvalue(line[34-chinese_count:38-chinese_count],3)

                        if isfloat(line[38-chinese_count:42-chinese_count]):
                            TQ1_ = getfloatvalue(line[38-chinese_count:42-chinese_count],3)

                        if isfloat(line[42-chinese_count:46-chinese_count]):
                            TQ2 = getfloatvalue(line[42-chinese_count:46-chinese_count],3)

                        if isfloat(line[46-chinese_count:50-chinese_count]):
                            TQ2_ = getfloatvalue(line[46-chinese_count:50-chinese_count],3)

                        if isfloat(line[50-chinese_count:54-chinese_count]):
                            TQ3 = getfloatvalue(line[50-chinese_count:54-chinese_count],3)

                        if isfloat(line[54-chinese_count:58-chinese_count]):
                            TQ3_ = getfloatvalue(line[54-chinese_count:58-chinese_count],3)                                

                        if isfloat(line[58-chinese_count:62-chinese_count]):
                            VSMAX = getfloatvalue(line[58-chinese_count:62-chinese_count],3)

                        if isfloat(line[62-chinese_count:66-chinese_count]):
                            VCUTOFF = getfloatvalue(line[62-chinese_count:66-chinese_count],3)

                        if isfloat(line[66-chinese_count:68-chinese_count]):
                            VSLOW = getfloatvalue(line[66-chinese_count:68-chinese_count],2)

                        if VSLOW<=0: VSMIN = VSMAX*(-1)
                        else: VSMIN = VSLOW*(-1)

                        pss_type.KQV = KQV
                        pss_type.TQV = TQV
                        pss_type.KQS = abs(KQS*MVABASE/MVABASE1)
                        pss_type.TQS = TQS
                        pss_type.TQ = TQ
                        pss_type.TQ1_ = TQ1_                            
                        pss_type.TQ1 = TQ1
                        pss_type.TQ2_ = TQ2_
                        pss_type.TQ2 = TQ2
                        pss_type.TQ3_ = TQ3_
                        pss_type.TQ3 = TQ3
                        pss_type.Vcutoff= VCUTOFF
                        pss_type.VSMIN = VSMIN
                        pss_type.VSMAX = VSMAX

                    #SG
                    if line[1] == 'G':

                        Generator_comp_pss = Generator_comp.Pss
                        Generator_comp_pss.typ_id = pss_SG
                        pss_type = Generator_comp.Pss

                        sym = Generator_comp.Sym

                        sym_typ = sym.typ_id
                        MVABASE = sym_typ.sgn   
#?????????????????????????????????????????????????????? KQS pu
                        if isfloat(line[76-chinese_count:80-chinese_count]):
                            MVABASE1 = getfloatvalue(line[76-chinese_count:80-chinese_count],0)
                        else:
                            MVABASE1 = MVABASE
                        if MVABASE1 == 0:
                            MVABASE1 = MVABASE
                            
                        KQV = 0
                        TQV = 0
                        KQS = 0
                        TQS = 0
                        TQ = 0
                        TQ1_ = 0                            
                        TQ1 = 0
                        TQ2_ = 0
                        TQ2 = 0
                        TQ3_ = 0
                        TQ3 = 0
                        VCUTOFF = 0
                        VSMIN = -9999
                        VSMAX = 9999
                        VSSLOW = 0

                        if isfloat(line[16-chinese_count:20-chinese_count]):
                            KQV = getfloatvalue(line[16-chinese_count:20-chinese_count],3)

                        if isfloat(line[20-chinese_count:23-chinese_count]):
                            TQV = getfloatvalue(line[20-chinese_count:23-chinese_count],3)                         
                        
                        if isfloat(line[23-chinese_count:27-chinese_count]):
                            KQS = getfloatvalue(line[23-chinese_count:27-chinese_count],3)

                        if isfloat(line[27-chinese_count:30-chinese_count]):
                            TQS = getfloatvalue(line[27-chinese_count:30-chinese_count],3)

                        if isfloat(line[30-chinese_count:34-chinese_count]):
                            TQ = getfloatvalue(line[30-chinese_count:34-chinese_count],2)

                        if isfloat(line[34-chinese_count:38-chinese_count]):
                            TQ1 = getfloatvalue(line[34-chinese_count:38-chinese_count],3)

                        if isfloat(line[38-chinese_count:42-chinese_count]):
                            TQ1_ = getfloatvalue(line[38-chinese_count:42-chinese_count],3)

                        if isfloat(line[42-chinese_count:46-chinese_count]):
                            TQ2 = getfloatvalue(line[42-chinese_count:46-chinese_count],3)

                        if isfloat(line[46-chinese_count:50-chinese_count]):
                            TQ2_ = getfloatvalue(line[46-chinese_count:50-chinese_count],3)

                        if isfloat(line[50-chinese_count:54-chinese_count]):
                            TQ3 = getfloatvalue(line[50-chinese_count:54-chinese_count],3)

                        if isfloat(line[54-chinese_count:58-chinese_count]):
                            TQ3_ = getfloatvalue(line[54-chinese_count:58-chinese_count],3)                                

                        if isfloat(line[58-chinese_count:62-chinese_count]):
                            VSMAX = getfloatvalue(line[58-chinese_count:62-chinese_count],3)

                        if isfloat(line[62-chinese_count:66-chinese_count]):
                            VCUTOFF = getfloatvalue(line[62-chinese_count:66-chinese_count],3)

                        if isfloat(line[66-chinese_count:68-chinese_count]):
                            VSLOW = getfloatvalue(line[66-chinese_count:68-chinese_count],2)

                        if VSLOW<=0: VSMIN = VSMAX*(-1)
                        else: VSMIN = VSLOW*(-1)

                        pss_type.KQV = KQV
                        pss_type.TQV = TQV
                        pss_type.KQS = abs(KQS*MVABASE/MVABASE1)
                        pss_type.TQS = TQS
                        pss_type.TQ = TQ
                        pss_type.TQ1_ = TQ1_                            
                        pss_type.TQ1 = TQ1
                        pss_type.TQ2_ = TQ2_
                        pss_type.TQ2 = TQ2
                        pss_type.TQ3_ = TQ3_
                        pss_type.TQ3 = TQ3
                        pss_type.Vcutoff= VCUTOFF
                        pss_type.VSMIN = VSMIN
                        pss_type.VSMAX = VSMAX

                    #SS
                    if line[1] == 'S':

                        Generator_comp_pss = Generator_comp.Pss
                        Generator_comp_pss.typ_id = pss_SS
                        pss_type = Generator_comp.Pss

                        sym = Generator_comp.Sym

                        sym_typ = sym.typ_id
                        MVABASE = sym_typ.sgn   

                        if isfloat(line[76-chinese_count:80-chinese_count]):
                            MVABASE1 = getfloatvalue(line[76-chinese_count:80-chinese_count],0)
                        else:
                            MVABASE1 = MVABASE

                        KQV = 0
                        TQV = 0
                        KQS = 0
                        TQS = 0
                        TQ = 0
                        TQ1_ = 0                            
                        TQ1 = 0
                        TQ2_ = 0
                        TQ2 = 0
                        TQ3_ = 0
                        TQ3 = 0
                        VCUTOFF = 0
                        VSMIN = -9999
                        VSMAX = 9999
                        VSLOW = 0

                        if isfloat(line[16-chinese_count:20-chinese_count]):
                            KQV = getfloatvalue(line[16-chinese_count:20-chinese_count],3)

                        if isfloat(line[20-chinese_count:23-chinese_count]):
                            TQV = getfloatvalue(line[20-chinese_count:23-chinese_count],3)                         
                        
                        if isfloat(line[23-chinese_count:27-chinese_count]):
                            KQS = getfloatvalue(line[23-chinese_count:27-chinese_count],3)

                        if isfloat(line[27-chinese_count:30-chinese_count]):
                            TQS = getfloatvalue(line[27-chinese_count:30-chinese_count],3)

                        if isfloat(line[30-chinese_count:34-chinese_count]):
                            TQ = getfloatvalue(line[30-chinese_count:34-chinese_count],2)

                        if isfloat(line[34-chinese_count:38-chinese_count]):
                            TQ1 = getfloatvalue(line[34-chinese_count:38-chinese_count],3)

                        if isfloat(line[38-chinese_count:42-chinese_count]):
                            TQ1_ = getfloatvalue(line[38-chinese_count:42-chinese_count],3)

                        if isfloat(line[42-chinese_count:46-chinese_count]):
                            TQ2 = getfloatvalue(line[42-chinese_count:46-chinese_count],3)

                        if isfloat(line[46-chinese_count:50-chinese_count]):
                            TQ2_ = getfloatvalue(line[46-chinese_count:50-chinese_count],3)

                        if isfloat(line[50-chinese_count:54-chinese_count]):
                            TQ3 = getfloatvalue(line[50-chinese_count:54-chinese_count],3)

                        if isfloat(line[54-chinese_count:58-chinese_count]):
                            TQ3_ = getfloatvalue(line[54-chinese_count:58-chinese_count],3)                                

                        if isfloat(line[58-chinese_count:62-chinese_count]):
                            VSMAX = getfloatvalue(line[58-chinese_count:62-chinese_count],3)

                        if isfloat(line[62-chinese_count:66-chinese_count]):
                            VCUTOFF = getfloatvalue(line[62-chinese_count:66-chinese_count],3)

                        if isfloat(line[66-chinese_count:68-chinese_count]):
                            VSLOW = getfloatvalue(line[66-chinese_count:68-chinese_count],2)

                        if VSLOW<=0: VSMIN = VSMAX*(-1)
                        else: VSMIN = VSLOW*(-1)

                        pss_type.KQV = KQV
                        pss_type.TQV = TQV
                        pss_type.KQS = KQS
                        pss_type.TQS = TQS
                        pss_type.TQ = TQ
                        pss_type.TQ1_ = TQ1_                            
                        pss_type.TQ1 = TQ1
                        pss_type.TQ2_ = TQ2_
                        pss_type.TQ2 = TQ2
                        pss_type.TQ3_ = TQ3_
                        pss_type.TQ3 = TQ3
                        pss_type.Vcutoff= VCUTOFF
                        pss_type.VSMIN = VSMIN
                        pss_type.VSMAX = VSMAX

                    #SH1/SH2
#SH1没用到，而且有问题！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！
                    if line[1] == 'H' :
                        if line[2] == '1' :#SH1

                            Generator_comp_pss = Generator_comp.Pss
                            Generator_comp_pss.typ_id = pss_SH1
                            pss_type = Generator_comp.Pss

                            sym = Generator_comp.Sym

                            sym_typ = sym.typ_id
                            MVABASE = sym_typ.sgn   
                                                           
                            KP = 0
                            TD = 0
                            K0 = 0
                            K1 = 0
                            K2 = 0
                            K3 = 0
                            T2 = 0
                            T3 = 0
                            T1 = 0
                            K = 0
                            VSMIN = -9999
                            VSMAX = 9999

#???????????????????????????????????????? KP pu
                            if isfloat(line[16-chinese_count:21-chinese_count]):
                                TD = getfloatvalue(line[16-chinese_count:21-chinese_count],3)

                            if isfloat(line[21-chinese_count:26-chinese_count]):
                                 T1 = getfloatvalue(line[21-chinese_count:26-chinese_count],3)

                            if isfloat(line[26-chinese_count:31-chinese_count]):
                                T2 = getfloatvalue(line[26-chinese_count:31-chinese_count],3)

                            if isfloat(line[31-chinese_count:36-chinese_count]):
                                T3 = getfloatvalue(line[31-chinese_count:36-chinese_count],3)

                            if isfloat(line[36-chinese_count:41-chinese_count]):
                                K0 = getfloatvalue(line[36-chinese_count:41-chinese_count],4)

                            if isfloat(line[41-chinese_count:46-chinese_count]):
                                K1 = getfloatvalue(line[41-chinese_count:46-chinese_count],4)

                            if isfloat(line[46-chinese_count:51-chinese_count]):
                                K2 = getfloatvalue(line[46-chinese_count:51-chinese_count],4)

                            if isfloat(line[51-chinese_count:56-chinese_count]):
                                K3 = getfloatvalue(line[51-chinese_count:56-chinese_count],4)

                            if isfloat(line[56-chinese_count:61-chinese_count]):
                                K = getfloatvalue(line[56-chinese_count:61-chinese_count],3)

                            if isfloat(line[61-chinese_count:66-chinese_count]):
                                VSMAX = getfloatvalue(line[61-chinese_count:66-chinese_count],3)

                            if isfloat(line[66-chinese_count:71-chinese_count]):
                                VSMIN = getfloatvalue(line[66-chinese_count:71-chinese_count],3)

                            if isfloat(line[71-chinese_count:76-chinese_count]):
                                KP = getfloatvalue(line[71-chinese_count:76-chinese_count],4)

                            if isfloat(line[76-chinese_count:80-chinese_count]):
                                MVABASE1 = getfloatvalue(line[76-chinese_count:80-chinese_count],0)
                            else:
                                MVABASE1 = MVABASE
                            if MVABASE1 == 0:
                                MVABASE1 = MVABASE
                            
                            pss_type.KP = KP*MVABASE/MVABASE1
                            pss_type.TD = TD
                            pss_type.K0 = K0
                            pss_type.K1 = K1
                            pss_type.K2 = K2
                            pss_type.K3 = K3
                            pss_type.T2 = T2
                            pss_type.T3 = T3
                            pss_type.T1 = T1
                            pss_type.K = K
                            pss_type.VSMIN = VSMIN
                            pss_type.VSMAX = VSMAX

                            raise Exception("SH1")

                            #SH2
                        elif line[2] == ' ' :
                            Generator_comp_pss = Generator_comp.Pss
                            Generator_comp_pss.typ_id = pss_SH2
                            pss_type = Generator_comp.Pss

                            sym = Generator_comp.Sym

                            sym_typ = sym.typ_id
                            MVABASE = sym_typ.sgn   
                        	
                            K0 = 0
                            K1 = 0
                            K2 = 0
                            K3 = 0
                            T2 = 0
                            T3 = 0
                            T1 = 0
                            K = 0
                            VSMIN = -9999
                            VSMAX = 9999

                            if isfloat(line[21-chinese_count:26-chinese_count]):
                                T1 = getfloatvalue(line[21-chinese_count:26-chinese_count],3)

                            if isfloat(line[26-chinese_count:31-chinese_count]):
                                T2 = getfloatvalue(line[26-chinese_count:31-chinese_count],3)

                            if isfloat(line[31-chinese_count:36-chinese_count]):
                                T3 = getfloatvalue(line[31-chinese_count:36-chinese_count],3)

                            if isfloat(line[36-chinese_count:41-chinese_count]):
                                K0 = getfloatvalue(line[36-chinese_count:41-chinese_count],4)

                            if isfloat(line[41-chinese_count:46-chinese_count]):
                                K1 = getfloatvalue(line[41-chinese_count:46-chinese_count],4)

                            if isfloat(line[46-chinese_count:51-chinese_count]):
                                K2 = getfloatvalue(line[46-chinese_count:51-chinese_count],4)

                            if isfloat(line[51-chinese_count:56-chinese_count]):
                                K3 = getfloatvalue(line[51-chinese_count:56-chinese_count],4)

                            if isfloat(line[56-chinese_count:61-chinese_count]):
                                K = getfloatvalue(line[56-chinese_count:61-chinese_count],3)

                            if isfloat(line[61-chinese_count:66-chinese_count]):
                                VSMAX = getfloatvalue(line[61-chinese_count:66-chinese_count],3)

                            if isfloat(line[66-chinese_count:71-chinese_count]):
                                VSMIN = getfloatvalue(line[66-chinese_count:71-chinese_count],3)

                            pss_type.K3 = K3
                            pss_type.K0 = K0
                            pss_type.K1 = K1
                            pss_type.K2 = K2
                            pss_type.T3 = T3
                            pss_type.T1 = T1
                            pss_type.T2 = T2
                            pss_type.K = K
                            pss_type.VSMIN = VSMIN
                            pss_type.VSMAX = VSMAX

                    #SI
                    if line[1] == 'I' and line[2] != '+':

                        INP = 0
                        if isint(line[79-chinese_count]):
                            INP = int(line[79-chinese_count])
                        if INP ==  0:
                            Generator_comp_pss = Generator_comp.Pss
                            Generator_comp_pss.typ_id = pss_SI_0
                            pss_type = Generator_comp.Pss

                            sym = Generator_comp.Sym

                            sym_typ = sym.typ_id
                            MVABASE = sym_typ.sgn   

                            Trw = 0
                            T12 = 0
                            Kr = 0
                            Trp = 0
                            TW = 0
                            TW1 = 0
                            T10 = 0
                            T9 = 0
                            T5 = 0
                            T6 = 0
                            TW2 = 0
                            KS = 0
                            T7 = 0

                            if isfloat(line[16-chinese_count:20-chinese_count]):
                                Trw = getfloatvalue(line[16-chinese_count:20-chinese_count],4)
                                
                            if isfloat(line[20-chinese_count:25-chinese_count]):
                                T5 = getfloatvalue(line[20-chinese_count:25-chinese_count],3)

                            if isfloat(line[25-chinese_count:30-chinese_count]):
                                T6 = getfloatvalue(line[25-chinese_count:30-chinese_count],3)

                            if isfloat(line[30-chinese_count:35-chinese_count]):
                                T7 = getfloatvalue(line[30-chinese_count:35-chinese_count],3)

                            if isfloat(line[35-chinese_count:41-chinese_count]):
                                Kr = getfloatvalue(line[35-chinese_count:41-chinese_count],4)
    
                            if isfloat(line[41-chinese_count:45-chinese_count]):
                                Trp = getfloatvalue(line[41-chinese_count:45-chinese_count],4)

                            if isfloat(line[45-chinese_count:50-chinese_count]):
                                TW = getfloatvalue(line[45-chinese_count:50-chinese_count],3)
        
                            if isfloat(line[50-chinese_count:55-chinese_count]):
                                TW1 = getfloatvalue(line[50-chinese_count:55-chinese_count],3)

                            if isfloat(line[55-chinese_count:60-chinese_count]):
                                TW2 = getfloatvalue(line[55-chinese_count:60-chinese_count],3)

                            if isfloat(line[60-chinese_count:64-chinese_count]):
                                KS = getfloatvalue(line[60-chinese_count:64-chinese_count],2)

                            if isfloat(line[64-chinese_count:69-chinese_count]):
                                T9 = getfloatvalue(line[64-chinese_count:69-chinese_count],3)

                            if isfloat(line[69-chinese_count:74-chinese_count]):
                                T10 = getfloatvalue(line[69-chinese_count:74-chinese_count],3)
        
                            if isfloat(line[74-chinese_count:79-chinese_count]):
                                T12 = getfloatvalue(line[74-chinese_count:79-chinese_count],3)

                            pss_type.Trw = Trw
                            pss_type.T12 = T12
                            pss_type.Kr = Kr
                            pss_type.Trp = Trp
                            pss_type.Tw = TW
                            pss_type.Tw1 = TW1
                            pss_type.T10 = T10
                            pss_type.T9 = T9
                            pss_type.T5 = T5
                            pss_type.T6 = T6
                            pss_type.Tw2 = TW2
                            pss_type.Ks = KS
                            pss_type.T7 = T7

                        elif INP ==  1:
                            Generator_comp_pss = Generator_comp.Pss
                            Generator_comp_pss.typ_id = pss_SI_1
                            pss_type = Generator_comp.Pss

                            sym = Generator_comp.Sym

                            sym_typ = sym.typ_id
                            MVABASE = sym_typ.sgn   

                            Trw = 0
                            T12 = 0
                            T10 = 0
                            T9 = 0
                            T5 = 0
                            T6 = 0
                            T7 = 0

                            if isfloat(line[16-chinese_count:20-chinese_count]):
                                Trw = getfloatvalue(line[16-chinese_count:20-chinese_count],4)
                                
                            if isfloat(line[20-chinese_count:25-chinese_count]):
                                T5 = getfloatvalue(line[20-chinese_count:25-chinese_count],3)

                            if isfloat(line[25-chinese_count:30-chinese_count]):
                                T6 = getfloatvalue(line[25-chinese_count:30-chinese_count],3)

                            if isfloat(line[30-chinese_count:35-chinese_count]):
                                T7 = getfloatvalue(line[30-chinese_count:35-chinese_count],3)

                            if isfloat(line[64-chinese_count:69-chinese_count]):
                                T9 = getfloatvalue(line[64-chinese_count:69-chinese_count],3)

                            if isfloat(line[69-chinese_count:74-chinese_count]):
                                T10 = getfloatvalue(line[69-chinese_count:74-chinese_count],3)
        
                            if isfloat(line[74-chinese_count:79-chinese_count]):
                                T12 = getfloatvalue(line[74-chinese_count:79-chinese_count],3)

                            pss_type.Trw = Trw
                            pss_type.T12 = T12
                            pss_type.T10 = T10
                            pss_type.T9 = T9
                            pss_type.T5 = T5
                            pss_type.T6 = T6
                            pss_type.T7 = T7

                        elif INP == 2:
                            Generator_comp_pss = Generator_comp.Pss
                            Generator_comp_pss.typ_id = pss_SI_2
                            pss_type = Generator_comp.Pss

                            sym = Generator_comp.Sym

                            sym_typ = sym.typ_id
                            MVABASE = sym_typ.sgn       

                            T12 = 0
                            Kr = 0
                            Trp = 0
                            TW = 0
                            TW1 = 0
                            T10 = 0
                            T9 = 0
                            TW2 = 0
                            KS = 0

                            if isfloat(line[35-chinese_count:41-chinese_count]):
                                Kr = getfloatvalue(line[35-chinese_count:41-chinese_count],4)
    
                            if isfloat(line[41-chinese_count:45-chinese_count]):
                                Trp = getfloatvalue(line[41-chinese_count:45-chinese_count],4)

                            if isfloat(line[45-chinese_count:50-chinese_count]):
                                TW = getfloatvalue(line[45-chinese_count:50-chinese_count],3)
        
                            if isfloat(line[50-chinese_count:55-chinese_count]):
                                TW1 = getfloatvalue(line[50-chinese_count:55-chinese_count],3)

                            if isfloat(line[55-chinese_count:60-chinese_count]):
                                TW2 = getfloatvalue(line[55-chinese_count:60-chinese_count],3)

                            if isfloat(line[60-chinese_count:64-chinese_count]):
                                KS = getfloatvalue(line[60-chinese_count:64-chinese_count],2)

                            if isfloat(line[64-chinese_count:69-chinese_count]):
                                T9 = getfloatvalue(line[64-chinese_count:69-chinese_count],3)

                            if isfloat(line[69-chinese_count:74-chinese_count]):
                                T10 = getfloatvalue(line[69-chinese_count:74-chinese_count],3)
        
                            if isfloat(line[74-chinese_count:79-chinese_count]):
                                T12 = getfloatvalue(line[74-chinese_count:79-chinese_count],3)

                            pss_type.T12 = T12
                            pss_type.Kr = Kr
                            pss_type.Trp = Trp
                            pss_type.Tw = TW
                            pss_type.Tw1 = TW1
                            pss_type.T10 = T10
                            pss_type.T9 = T9
                            pss_type.Tw2 = TW2
                            pss_type.Ks = KS

                    #SA
                    if line[1] == 'A':

                        Generator_comp_pss = Generator_comp.Pss
                        Generator_comp_pss.typ_id = pss_SA
                        pss_type = Generator_comp.Pss

                        sym = Generator_comp.Sym

                        sym_typ = sym.typ_id
                        MVABASE = sym_typ.sgn       

                        T1 = 0
                        T2 = 0
                        T3 = 0
                        T4 = 0
                        T5 = 0
                        T6 = 0
                        K2 = 0
                        K1 = 0
                        T6 = 0
                        A = 0
                        P = 0
                        K = 0
                        VSMIN = -9999
                        VSMAX = 9999
                        
#???????????????????????????????????????? K pu
                        if isfloat(line[75-chinese_count:79-chinese_count]):
                            MVABASE1 = getfloatvalue(line[75-chinese_count:79-chinese_count],0)
                        else:
                            MVABASE1= MVABASE
                                
                        if isfloat(line[16-chinese_count:20-chinese_count]):
                            T1 = getfloatvalue(line[16-chinese_count:20-chinese_count],4)

                        if isfloat(line[20-chinese_count:24-chinese_count]):
                            T2 = getfloatvalue(line[20-chinese_count:24-chinese_count],4)

                        if isfloat(line[24-chinese_count:28-chinese_count]):
                            T3 = getfloatvalue(line[24-chinese_count:28-chinese_count],3)

                        if isfloat(line[28-chinese_count:32-chinese_count]):
                            T4 = getfloatvalue(line[28-chinese_count:32-chinese_count],3)

                        if isfloat(line[32-chinese_count:36-chinese_count]):
                            T5 = getfloatvalue(line[32-chinese_count:36-chinese_count],3)

                        if isfloat(line[36-chinese_count:40-chinese_count]):
                            T6 = getfloatvalue(line[36-chinese_count:40-chinese_count],3)

                        if isfloat(line[40-chinese_count:45-chinese_count]):
                            K1 = getfloatvalue(line[40-chinese_count:45-chinese_count],3)

                        if isfloat(line[45-chinese_count:50-chinese_count]):
                            K2 = getfloatvalue(line[45-chinese_count:50-chinese_count],3)

                        if isfloat(line[50-chinese_count:55-chinese_count]):
                            K = getfloatvalue(line[50-chinese_count:55-chinese_count],3)
                    
                        if isfloat(line[55-chinese_count:60-chinese_count]):
                            A = getfloatvalue(line[55-chinese_count:60-chinese_count],3)

                        if isfloat(line[60-chinese_count:65-chinese_count]):
                            P = getfloatvalue(line[60-chinese_count:65-chinese_count],3)

                        if isfloat(line[65-chinese_count:70-chinese_count]):
                            VSMAX = getfloatvalue(line[65-chinese_count:70-chinese_count],3)

                        if isfloat(line[70-chinese_count:75-chinese_count]):
                            VSMIN = getfloatvalue(line[70-chinese_count:75-chinese_count],3)

                        pss_type.T1 = T1
                        pss_type.T2 = T2
                        pss_type.T3 = T3
                        pss_type.T4 = T4
                        pss_type.T5 = T5
                        pss_type.K2 = K2
                        pss_type.K1 = K1
                        pss_type.T6 = T6
                        pss_type.A = A
                        pss_type.P = P
                        pss_type.K = K*MVABASE/MVABASE1
                        pss_type.VSMIN = VSMIN
                        pss_type.VSMAX = VSMAX

                        
                    #SB
                    if line[1] == 'B' and  line[2] != '+':
                        ISIG = 0
                        if isint(line[16-chinese_count]):
                            ISIG = int(line[16-chinese_count])

                        if ISIG == 0:
                            Generator_comp_pss = Generator_comp.Pss
                            Generator_comp_pss.typ_id = pss_SB_0
                            pss_type = Generator_comp.Pss

                            sym = Generator_comp.Sym

                            sym_typ = sym.typ_id
                            MVABASE = sym_typ.sgn

                            TW1 = 0
                            TW2 = 0
                            T6 = 0
                            TD = 0
                            KS3 = 0
                            KS2 = 0
                            T7 = 0
                            KPG = 0
                            TW3 = 0
                            TW4 = 0

                            if isfloat(line[67-chinese_count:72-chinese_count]):
                                MVABASE1 = getfloatvalue(line[67-chinese_count:72-chinese_count],0)
                            else:
                                MVABASE1 = MVABASE
                            if MVABASE1 == 0:
                                MVABASE1 = MVABASE

                            if isfloat(line[17-chinese_count:22-chinese_count]):
                                TD = getfloatvalue(line[17-chinese_count:22-chinese_count],0)

                            if isfloat(line[22-chinese_count:27-chinese_count]):
                                TW1 = getfloatvalue(line[22-chinese_count:27-chinese_count],0)

                            if isfloat(line[27-chinese_count:32-chinese_count]):
                                TW2 = getfloatvalue(line[27-chinese_count:32-chinese_count],0)

                            if isfloat(line[32-chinese_count:37-chinese_count]):
                                T6 = getfloatvalue(line[32-chinese_count:37-chinese_count],0)
                            
                            if isfloat(line[37-chinese_count:42-chinese_count]):
                                TW3 = getfloatvalue(line[37-chinese_count:42-chinese_count],0)

                            if isfloat(line[42-chinese_count:47-chinese_count]):
                                TW4 = getfloatvalue(line[42-chinese_count:47-chinese_count],0)

                            if isfloat(line[47-chinese_count:52-chinese_count]):
                                T7 = getfloatvalue(line[47-chinese_count:52-chinese_count],0)

                            if isfloat(line[52-chinese_count:57-chinese_count]):
                                KS2 = getfloatvalue(line[52-chinese_count:57-chinese_count],0)

                            if isfloat(line[57-chinese_count:62-chinese_count]):
                                KS3 = getfloatvalue(line[57-chinese_count:62-chinese_count],0)

                            if isfloat(line[62-chinese_count:67-chinese_count]):
                                KPG = getfloatvalue(line[62-chinese_count:67-chinese_count],0)
                                KPG = KPG*MVABASE/MVABASE1
                            else:
                                KPG = 1

                            pss_type.Tw1 = TW1
                            pss_type.TW2 = TW2
                            pss_type.T6 = T6
                            pss_type.TD = TD
                            pss_type.KS3 = KS3
                            pss_type.KS2 = KS2
                            pss_type.T7 = T7
                            pss_type.KPG = KPG
                            pss_type.TW3 = TW3
                            pss_type.TW4 = TW4

                        #SB_1
                        elif ISIG == 1:
                            Generator_comp_pss = Generator_comp.Pss
                            Generator_comp_pss.typ_id = pss_SB_1
                            pss_type = Generator_comp.Pss

                            sym = Generator_comp.Sym

                            sym_typ = sym.typ_id
                            MVABASE = sym_typ.sgn   

                            TW1 = 0
                            TW2 = 0
                            T6 = 0
                            TD = 0                                

                            if isfloat(line[17-chinese_count:22-chinese_count]):
                                TD = getfloatvalue(line[17-chinese_count:22-chinese_count],0)

                            if isfloat(line[22-chinese_count:27-chinese_count]):
                                TW1 = getfloatvalue(line[22-chinese_count:27-chinese_count],0)

                            if isfloat(line[27-chinese_count:32-chinese_count]):
                                TW2 = getfloatvalue(line[27-chinese_count:32-chinese_count],0)

                            if isfloat(line[32-chinese_count:37-chinese_count]):
                                T6 = getfloatvalue(line[32-chinese_count:37-chinese_count],0)
                            

                            pss_type.Tw1 = TW1
                            pss_type.TW2 = TW2
                            pss_type.T6 = T6
                            pss_type.TD = TD

                        elif ISIG == 2:
                            Generator_comp_pss = Generator_comp.Pss
                            Generator_comp_pss.typ_id = pss_SB_2
                            pss_type = Generator_comp.Pss

                            sym = Generator_comp.Sym

                            sym_typ = sym.typ_id
                            MVABASE = sym_typ.sgn   
                        	
                            TD = 0
                            KS3 = 0
                            KS2 = 0
                            T7 = 0
                            KPG = 0
                            TW3 = 0
                            TW4 = 0

#???????????????????????????????????????? KPG pu
                            if isfloat(line[67-chinese_count:72-chinese_count]):
                                MVABASE1 = getfloatvalue(line[67-chinese_count:72-chinese_count],0)
                            else:
                                MVABASE1 = MVABASE
                            if MVABASE1 == 0:
                                MVABASE1 = MVABASE                                

                            if isfloat(line[17-chinese_count:22-chinese_count]):
                                TD = getfloatvalue(line[17-chinese_count:22-chinese_count],0)
                            
                            if isfloat(line[37-chinese_count:42-chinese_count]):
                                TW3 = getfloatvalue(line[37-chinese_count:42-chinese_count],0)

                            if isfloat(line[42-chinese_count:47-chinese_count]):
                                TW4 = getfloatvalue(line[42-chinese_count:47-chinese_count],0)

                            if isfloat(line[47-chinese_count:52-chinese_count]):
                                T7 = getfloatvalue(line[47-chinese_count:52-chinese_count],0)

                            if isfloat(line[52-chinese_count:57-chinese_count]):
                                KS2 = getfloatvalue(line[52-chinese_count:57-chinese_count],0)

                            if isfloat(line[57-chinese_count:62-chinese_count]):
                                KS3 = getfloatvalue(line[57-chinese_count:62-chinese_count],0)

                            if isfloat(line[62-chinese_count:67-chinese_count]):
                                KPG = getfloatvalue(line[62-chinese_count:67-chinese_count],0)
                                KPG = KPG*MVABASE/MVABASE1
                            else:
                                KPG = 1

                            pss_type.TD = TD
                            pss_type.KS3 = KS3
                            pss_type.KS2 = KS2
                            pss_type.T7 = T7
                            pss_type.KPG = KPG
                            pss_type.TW3 = TW3
                            pss_type.TW4 = TW4

                            
                
    #SH+卡
    for line in bpa_str_ar:
        line = line.rstrip('\n')
        if line == "": continue        

        #SH+卡
        if line[0] == 'S' and line[1] == 'H' and line[2] == '+':
            line = line + ' '*(80-len(line))
            chinese_count = 0
            #判断中文的个数
            for i in range (3,10):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count = chinese_count + 1
            name = line[3:11-chinese_count].strip()
            if isfloat(line[11-chinese_count:15-chinese_count]):base = float(line[11-chinese_count:15-chinese_count])
            #名称+电压
            Generator = name + str(base)                
            #找到节点名对应的编号
            if Generator in bus_name_ar:
                Generator_num = bus_name_ar[Generator]

                #找到对应的comp模块
                Generator_comp_name = str(Generator_num) + '_comp'
                Generator_compp = Net.SearchObject(str(Generator_comp_name))
                Generator_comp = Generator_compp[0]

                if not Generator_comp == None:
                    pss_type = Generator_comp.Pss
                    pss_type_id = pss_type.typ_id

                    sym = Generator_comp.Sym

                    sym_typ = sym.typ_id
                    MVABASE = sym_typ.sgn   

                    #SH1
                    if pss_type_id != None and pss_type_id.loc_name == pss_SH1.loc_name:

                        T4 = 0
                        K4 = 0
                            
                        if isfloat(line[61-chinese_count:66-chinese_count]):
                            T4 = getfloatvalue(line[61-chinese_count:66-chinese_count],4)
                            
                        if isfloat(line[66-chinese_count:71-chinese_count]):
                            K4 = getfloatvalue(line[66-chinese_count:71-chinese_count],3)

                        pss_type.params[9] = T4
                        pss_type.params[6] = K4
                    #SH2
                    if pss_type_id != None and pss_type_id.loc_name == pss_SH2.loc_name:

#???????????????????????????????????????? KPM KPE pu ???
                        if isfloat(line[76-chinese_count:80-chinese_count]):
                            MVABASE1 = getfloatvalue(line[76-chinese_count:80-chinese_count],0)
                        else:
                            MVABASE1 = MVABASE
                        if MVABASE1 == 0:
                            MVABASE1 = MVABASE

                        KPM = 0
                        TPM = 0
                        KPE = 0
                        TPE = 0
                        KW = 0
                        TW = 0
                        KD1 = 0
                        TD1 = 0
                        TD2 = 0
                        K4 = 0
                        KW = 0
                        T4 = 0   
                            
                        if isfloat(line[16-chinese_count:21-chinese_count]):
                            KPM = getfloatvalue(line[16-chinese_count:21-chinese_count],3)
                            
                        if isfloat(line[21-chinese_count:26-chinese_count]):
                            TPM = getfloatvalue(line[21-chinese_count:26-chinese_count],3)

                        if isfloat(line[26-chinese_count:31-chinese_count]):
                            KPE = getfloatvalue(line[26-chinese_count:31-chinese_count],3)

                        if isfloat(line[31-chinese_count:36-chinese_count]):
                            TPE = getfloatvalue(line[31-chinese_count:36-chinese_count],3)

                        if isfloat(line[36-chinese_count:41-chinese_count]):
                            KW = getfloatvalue(line[36-chinese_count:41-chinese_count],3)

                        if isfloat(line[41-chinese_count:46-chinese_count]):
                            TW = getfloatvalue(line[41-chinese_count:46-chinese_count],3)

                        if isfloat(line[46-chinese_count:51-chinese_count]):
                            TD1 = getfloatvalue(line[46-chinese_count:51-chinese_count],3)

                        if isfloat(line[51-chinese_count:56-chinese_count]):
                            TD2 = getfloatvalue(line[51-chinese_count:56-chinese_count],3)

                        if isint(line[56-chinese_count]):
                            KD1 = int(line[56-chinese_count])
                            

                        if isfloat(line[61-chinese_count:66-chinese_count]):
                            T4 = getfloatvalue(line[61-chinese_count:66-chinese_count],3)

                        if isfloat(line[66-chinese_count:71-chinese_count]):
                            K4 = getfloatvalue(line[66-chinese_count:71-chinese_count],3)

                        pss_type.KPM = KPM*MVABASE/MVABASE1
                        pss_type.TPM = TPM
                        pss_type.KPE = KPE*MVABASE/MVABASE1
                        pss_type.TPE = TPE
                        pss_type.KW = KW
                        pss_type.TW = TW
                        pss_type.KD1 = KD1
                        pss_type.TD1 = TD1
                        pss_type.TD2 = TD2
                        pss_type.K4 = K4
                        pss_type.T4 = T4        


    #SI+卡
    for line in bpa_str_ar:
        line = line.rstrip('\n')
        if line == "": continue        

        #SI+卡
        if line[0] == 'S' and line[1] == 'I' and line[2] == '+':
            line = line + ' '*(80-len(line))
            chinese_count = 0
            #判断中文的个数
            for i in range (3,10):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count = chinese_count + 1
            name = line[3:11-chinese_count].strip()
            if isfloat(line[11-chinese_count:15-chinese_count]):base = float(line[11-chinese_count:15-chinese_count])
            #名称+电压
            Generator = name + str(base)                
            #找到节点名对应的编号
            if Generator in bus_name_ar:
                Generator_num = bus_name_ar[Generator]

                #找到对应的comp模块
                Generator_comp_name = str(Generator_num) + '_comp'
                Generator_compp = Net.SearchObject(str(Generator_comp_name))
                Generator_comp = Generator_compp[0]

                if not Generator_comp == None:

                    pss_type = Generator_comp.Pss
                    pss_type_id = pss_type.typ_id
                    sym = Generator_comp.Sym
                        
                    sym_typ = sym.typ_id
                    MVABASE = sym_typ.sgn

                    #SI_0
                    if pss_type_id != None and pss_type_id.loc_name == pss_SI_0.loc_name:

                        T14 = 0
                        T13 = 0
                        T4 = 0
                        T3 = 0
                        T2 = 0
                        T1 = 0
                        KP = 0
                        VSMAX = 9999
                        VSMIN = -9999
                        
                        if isint(line[63-chinese_count]):
                            IB = int(line[63-chinese_count])
                            if IB != 0:
                                app.PrintInfo('SI feq wrong')
#??????????????????????????????????????????? kr pu
                        if isfloat(line[76-chinese_count:80-chinese_count]):
                            MVABASE1 = getfloatvalue(line[76-chinese_count:80-chinese_count],0)
                        else:
                            MVABASE1 = MVABASE
                        if MVABASE1 == 0:
                            MVABASE1 = MVABASE

                        
                        if isfloat(line[16-chinese_count:21-chinese_count]):
                            KP = getfloatvalue(line[16-chinese_count:21-chinese_count],3)

                        if isfloat(line[21-chinese_count:26-chinese_count]):
                            T1 = getfloatvalue(line[21-chinese_count:26-chinese_count],3)

                        if isfloat(line[26-chinese_count:31-chinese_count]):
                            T2 = getfloatvalue(line[26-chinese_count:31-chinese_count],3)

                        if isfloat(line[31-chinese_count:36-chinese_count]):
                            T13 = getfloatvalue(line[31-chinese_count:36-chinese_count],3)

                        if isfloat(line[36-chinese_count:41-chinese_count]):
                            T14 = getfloatvalue(line[36-chinese_count:41-chinese_count],3)

                        if isfloat(line[41-chinese_count:46-chinese_count]):
                            T3 = getfloatvalue(line[41-chinese_count:46-chinese_count],3)

                        if isfloat(line[46-chinese_count:51-chinese_count]):
                            T4 = getfloatvalue(line[46-chinese_count:51-chinese_count],3)

                        if isfloat(line[51-chinese_count:57-chinese_count]):
                            VSMAX = getfloatvalue(line[51-chinese_count:57-chinese_count],4)

                        if isfloat(line[57-chinese_count:63-chinese_count]):
                            VSMIN = getfloatvalue(line[57-chinese_count:63-chinese_count],4)

                        pss_type.T14 = T14
                        pss_type.T13 = T13
                        pss_type.T4 = T4
                        pss_type.T3 = T3
                        pss_type.T2 = T2
                        pss_type.T1 = T1
                        pss_type.Kp = KP
                        pss_type.Vsmin = VSMIN
                        pss_type.Vsmax = VSMAX
                        
                        #修正KR
                        pss_type.Kr = pss_type.Kr*MVABASE/MVABASE1


                    #SI_1
                    if pss_type_id != None and pss_type_id.loc_name == pss_SI_1.loc_name:

                        T14 = 0
                        T13 = 0
                        T4 = 0
                        T3 = 0
                        T2 = 0
                        T1 = 0
                        KP = 0
                        VSMAX = 9999
                        VSMIN = -9999
                        
                        if isint(line[63-chinese_count]):
                            IB = int(line[63-chinese_count])
                            if IB != 0:
                                app.PrintInfo('SI feq wrong')
#??????????????????????????????????????????? kr pu
                        if isfloat(line[76-chinese_count:80-chinese_count]):
                            MVABASE1 = getfloatvalue(line[76-chinese_count:80-chinese_count],0)
                        else:
                            MVABASE1 =MVABASE
                        
                        if isfloat(line[16-chinese_count:21-chinese_count]):
                            KP = getfloatvalue(line[16-chinese_count:21-chinese_count],3)

                        if isfloat(line[21-chinese_count:26-chinese_count]):
                            T1 = getfloatvalue(line[21-chinese_count:26-chinese_count],3)

                        if isfloat(line[26-chinese_count:31-chinese_count]):
                            T2 = getfloatvalue(line[26-chinese_count:31-chinese_count],3)

                        if isfloat(line[31-chinese_count:36-chinese_count]):
                            T13 = getfloatvalue(line[31-chinese_count:36-chinese_count],3)

                        if isfloat(line[36-chinese_count:41-chinese_count]):
                            T14 = getfloatvalue(line[36-chinese_count:41-chinese_count],3)

                        if isfloat(line[41-chinese_count:46-chinese_count]):
                            T3 = getfloatvalue(line[41-chinese_count:46-chinese_count],3)

                        if isfloat(line[46-chinese_count:51-chinese_count]):
                            T4 = getfloatvalue(line[46-chinese_count:51-chinese_count],3)

                        if isfloat(line[51-chinese_count:57-chinese_count]):
                            VSMAX = getfloatvalue(line[51-chinese_count:57-chinese_count],4)

                        if isfloat(line[57-chinese_count:63-chinese_count]):
                            VSMIN = getfloatvalue(line[57-chinese_count:63-chinese_count],4)

                        pss_type.T14 = T14
                        pss_type.T13 = T13
                        pss_type.T4 = T4
                        pss_type.T3 = T3
                        pss_type.T2 = T2
                        pss_type.T1 = T1
                        pss_type.Kp = KP
                        pss_type.Vsmin = VSMIN
                        pss_type.Vsmax = VSMAX

                        #修正KR
                        pss_type.Kr = pss_type.Kr*MVABASE/MVABASE1

                    #SI_2
                    if pss_type_id != None and pss_type_id.loc_name == pss_SI_2.loc_name:

                        T14 = 0
                        T13 = 0
                        T4 = 0
                        T3 = 0
                        T2 = 0
                        T1 = 0
                        KP = 0
                        VSMAX = 9999
                        VSMIN = -9999
                        
                        if isint(line[63-chinese_count]):
                            IB = int(line[63-chinese_count])
                            if IB != 0:
                                app.PrintInfo('SI feq wrong')
#??????????????????????????????????????????? kr pu
                        if isfloat(line[76-chinese_count:80-chinese_count]):
                            MVABASE1 = getfloatvalue(line[76-chinese_count:80-chinese_count],0)
                        else:
                            MVABASE1 =MVABASE
                        
                        if isfloat(line[16-chinese_count:21-chinese_count]):
                            KP = getfloatvalue(line[16-chinese_count:21-chinese_count],3)

                        if isfloat(line[21-chinese_count:26-chinese_count]):
                            T1 = getfloatvalue(line[21-chinese_count:26-chinese_count],3)

                        if isfloat(line[26-chinese_count:31-chinese_count]):
                            T2 = getfloatvalue(line[26-chinese_count:31-chinese_count],3)

                        if isfloat(line[31-chinese_count:36-chinese_count]):
                            T13 = getfloatvalue(line[31-chinese_count:36-chinese_count],3)

                        if isfloat(line[36-chinese_count:41-chinese_count]):
                            T14 = getfloatvalue(line[36-chinese_count:41-chinese_count],3)

                        if isfloat(line[41-chinese_count:46-chinese_count]):
                            T3 = getfloatvalue(line[41-chinese_count:46-chinese_count],3)

                        if isfloat(line[46-chinese_count:51-chinese_count]):
                            T4 = getfloatvalue(line[46-chinese_count:51-chinese_count],3)

                        if isfloat(line[51-chinese_count:57-chinese_count]):
                            VSMAX = getfloatvalue(line[51-chinese_count:57-chinese_count],4)

                        if isfloat(line[57-chinese_count:63-chinese_count]):
                            VSMIN = getfloatvalue(line[57-chinese_count:63-chinese_count],4)

                        pss_type.T14 = T14
                        pss_type.T13 = T13
                        pss_type.T4 = T4
                        pss_type.T3 = T3
                        pss_type.T2 = T2
                        pss_type.T1 = T1
                        pss_type.Kp = KP
                        pss_type.Vsmin = VSMIN
                        pss_type.Vsmax = VSMAX

                        #修正KR
                        pss_type.Kr = pss_type.Kr*MVABASE/MVABASE1

    #SB+卡
    for line in bpa_str_ar:
        line = line.rstrip('\n')
        if line == "": continue        

        #SB+卡
        if line[0] == 'S' and line[1] == 'B' and line[2] == '+':
            line = line + ' '*(80-len(line))
            chinese_count = 0
            #判断中文的个数
            for i in range (3,10):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count = chinese_count + 1
            name = line[3:11-chinese_count].strip()
            if isfloat(line[11-chinese_count:15-chinese_count]):base = float(line[11-chinese_count:15-chinese_count])
            #名称+电压
            Generator = name + str(base)                
            #找到节点名对应的编号
            if Generator in bus_name_ar:
                Generator_num = bus_name_ar[Generator]

                #找到对应的comp模块
                Generator_comp_name = str(Generator_num) + '_comp'
                Generator_compp = Net.SearchObject(str(Generator_comp_name))
                Generator_comp = Generator_compp[0]

                if not Generator_comp == None:       
                    pss_type = Generator_comp.Pss
                    pss_type_id = pss_type.typ_id
                    
                    #SB_0
                    if pss_type_id != None and pss_type_id.loc_name == pss_SB_0.loc_name:

                        T8 = 0
                        T9 = 0
                        N = 0
                        M = 0
                        KS1 = 0
                        T1 = 0
                        T2 = 0
                        T3 = 0
                        T4 = 0
                        VPMIN = -9999
                        VPMAX = 9999
                        
                        if isfloat(line[16-chinese_count:21-chinese_count]):
                            T8 = getfloatvalue(line[16-chinese_count:21-chinese_count],0)

                        if isfloat(line[21-chinese_count:26-chinese_count]):
                            T9 = getfloatvalue(line[21-chinese_count:26-chinese_count],0)

                        if isint(line[26-chinese_count]):
                            M = int(line[26-chinese_count])

                        if isint(line[27-chinese_count]):
                            N = int(line[27-chinese_count])

                        if isfloat(line[28-chinese_count:33-chinese_count]):
                            KS1 = getfloatvalue(line[28-chinese_count:33-chinese_count],0)

                        if isfloat(line[33-chinese_count:38-chinese_count]):
                            T1 = getfloatvalue(line[33-chinese_count:38-chinese_count],0)

                        if isfloat(line[38-chinese_count:43-chinese_count]):
                            T2 = getfloatvalue(line[38-chinese_count:43-chinese_count],0)
                        
                        if isfloat(line[43-chinese_count:48-chinese_count]):
                            T3 = getfloatvalue(line[43-chinese_count:48-chinese_count],0)

                        if isfloat(line[48-chinese_count:53-chinese_count]):
                            T4 = getfloatvalue(line[48-chinese_count:53-chinese_count],0)

                        if isfloat(line[53-chinese_count:58-chinese_count]):
                            VPMAX = getfloatvalue(line[53-chinese_count:58-chinese_count],0)

                        if isfloat(line[58-chinese_count:63-chinese_count]):
                            VPMIN = getfloatvalue(line[58-chinese_count:63-chinese_count],0)

                        pss_type.T8 = T8
                        pss_type.T9 = T9
                        pss_type.N = N
                        pss_type.M = M
                        pss_type.KS1 = KS1
                        pss_type.T1 = T1
                        pss_type.T2 = T2
                        pss_type.T3 = T3
                        pss_type.T4 = T4
                        pss_type.VPMIN = VPMIN
                        pss_type.VPMAX = VPMAX

                    #SB_1
                    if pss_type_id != None and pss_type_id.loc_name == pss_SB_1.loc_name:

                        T8 = 0
                        T9 = 0
                        N = 0
                        M = 0
                        KS1 = 0
                        T1 = 0
                        T2 = 0
                        T3 = 0
                        T4 = 0
                        VPMIN = -9999
                        VPMAX = 9999
                        
                        if isfloat(line[16-chinese_count:21-chinese_count]):
                            T8 = getfloatvalue(line[16-chinese_count:21-chinese_count],0)

                        if isfloat(line[21-chinese_count:26-chinese_count]):
                            T9 = getfloatvalue(line[21-chinese_count:26-chinese_count],0)

                        if isint(line[26-chinese_count]):
                            M = int(line[26-chinese_count])

                        if isint(line[27-chinese_count]):
                            N = int(line[27-chinese_count])

                        if isfloat(line[28-chinese_count:33-chinese_count]):
                            KS1 = getfloatvalue(line[28-chinese_count:33-chinese_count],0)

                        if isfloat(line[33-chinese_count:38-chinese_count]):
                            T1 = getfloatvalue(line[33-chinese_count:38-chinese_count],0)

                        if isfloat(line[38-chinese_count:43-chinese_count]):
                            T2 = getfloatvalue(line[38-chinese_count:43-chinese_count],0)
                        
                        if isfloat(line[43-chinese_count:48-chinese_count]):
                            T3 = getfloatvalue(line[43-chinese_count:48-chinese_count],0)

                        if isfloat(line[48-chinese_count:53-chinese_count]):
                            T4 = getfloatvalue(line[48-chinese_count:53-chinese_count],0)

                        if isfloat(line[53-chinese_count:58-chinese_count]):
                            VPMAX = getfloatvalue(line[53-chinese_count:58-chinese_count],0)

                        if isfloat(line[58-chinese_count:63-chinese_count]):
                            VPMIN = getfloatvalue(line[58-chinese_count:63-chinese_count],0)

                        pss_type.T8 = T8
                        pss_type.T9 = T9
                        pss_type.N = N
                        pss_type.M = M
                        pss_type.KS1 = KS1
                        pss_type.T1 = T1
                        pss_type.T2 = T2
                        pss_type.T3 = T3
                        pss_type.T4 = T4
                        pss_type.VPMIN = VPMIN
                        pss_type.VPMAX = VPMAX

                        
                    #SB_2
                    if pss_type_id != None and pss_type_id.loc_name == pss_SB_2.loc_name:

                        T8 = 0
                        T9 = 0
                        N = 0
                        M = 0
                        KS1 = 0
                        T1 = 0
                        T2 = 0
                        T3 = 0
                        T4 = 0
                        VPMIN = -9999
                        VPMAX = 9999
                        
                        if isfloat(line[16-chinese_count:21-chinese_count]):
                            T8 = getfloatvalue(line[16-chinese_count:21-chinese_count],0)

                        if isfloat(line[21-chinese_count:26-chinese_count]):
                            T9 = getfloatvalue(line[21-chinese_count:26-chinese_count],0)

                        if isint(line[26-chinese_count]):
                            M = int(line[26-chinese_count])

                        if isint(line[27-chinese_count]):
                            N = int(line[27-chinese_count])

                        if isfloat(line[28-chinese_count:33-chinese_count]):
                            KS1 = getfloatvalue(line[28-chinese_count:33-chinese_count],0)

                        if isfloat(line[33-chinese_count:38-chinese_count]):
                            T1 = getfloatvalue(line[33-chinese_count:38-chinese_count],0)

                        if isfloat(line[38-chinese_count:43-chinese_count]):
                            T2 = getfloatvalue(line[38-chinese_count:43-chinese_count],0)
                        
                        if isfloat(line[43-chinese_count:48-chinese_count]):
                            T3 = getfloatvalue(line[43-chinese_count:48-chinese_count],0)

                        if isfloat(line[48-chinese_count:53-chinese_count]):
                            T4 = getfloatvalue(line[48-chinese_count:53-chinese_count],0)

                        if isfloat(line[53-chinese_count:58-chinese_count]):
                            VPMAX = getfloatvalue(line[53-chinese_count:58-chinese_count],0)

                        if isfloat(line[58-chinese_count:63-chinese_count]):
                            VPMIN = getfloatvalue(line[58-chinese_count:63-chinese_count],0)

                        pss_type.T8 = T8
                        pss_type.T9 = T9
                        pss_type.N = N
                        pss_type.M = M
                        pss_type.KS1 = KS1
                        pss_type.T1 = T1
                        pss_type.T2 = T2
                        pss_type.T3 = T3
                        pss_type.T4 = T4
                        pss_type.VPMIN = VPMIN
                        pss_type.VPMAX = VPMAX

                        
# ----------------------------------------------------------------------------------------------------
def GetGovCard(bpa_file, bpa_str_ar,bus_name_ar):
    #调速器G卡
    prj = app.GetActiveProject()
    if prj is None:
        raise Exception("No project activated. Python Script stopped.")

    app.PrintInfo(prj)
    Network_Modell = prj.SearchObject("Network Model.IntPrjfolder")
    Network_Model=Network_Modell[0]

    Network_Dataa = Network_Model.SearchObject("Network Data.IntPrjfolder")
    Network_Data = Network_Dataa[0]
    app.PrintInfo(Network_Data)

    Nett = Network_Data.SearchObject("*.ElmNet")
    Net = Nett[0]
    app.PrintInfo(Net)

    Libraryy = prj.SearchObject("Library.IntPrjfolder")
    Library = Libraryy[0]
    app.PrintInfo(Library)

    user_defined_modell = Library.SearchObject("User Defined Models.IntPrjfolder")
    user_defined_model = user_defined_modell[0]
    app.PrintInfo(user_defined_model)

    BPA_Frame_EE = user_defined_model.SearchObject("BPA Frame(E).BlkDef")
    BPA_Frame_E = BPA_Frame_EE[0]
    app.PrintInfo(BPA_Frame_E)

    BPA_Frame_FF = user_defined_model.SearchObject("BPA Frame(F).BlkDef")
    BPA_Frame_F = BPA_Frame_FF[0]
    app.PrintInfo(BPA_Frame_F)

    gov_GSS = user_defined_model.SearchObject("gov_GS.BlkDef")
    gov_GS = gov_GSS[0]
    app.PrintInfo(gov_GS)

    gov_GA_GII = user_defined_model.SearchObject("gov_GA_GI.BlkDef")
    gov_GA_GI = gov_GA_GII[0]
    app.PrintInfo(gov_GA_GI)

    gov_GA_GII_1111 = user_defined_model.SearchObject("gov_GA_GI_111.BlkDef")
    gov_GA_GI_111 = gov_GA_GII_1111[0]
    app.PrintInfo(gov_GA_GI_111)

    gov_GA_GII_1121 = user_defined_model.SearchObject("gov_GA_GI_112.BlkDef")
    gov_GA_GI_112 = gov_GA_GII_1121[0]
    app.PrintInfo(gov_GA_GI_112)

    gov_GA_GII_1211 = user_defined_model.SearchObject("gov_GA_GI_121.BlkDef")
    gov_GA_GI_121 = gov_GA_GII_1211[0]
    app.PrintInfo(gov_GA_GI_121)
    
    gov_GA_GII_1221 = user_defined_model.SearchObject("gov_GA_GI_122.BlkDef")
    gov_GA_GI_122 = gov_GA_GII_1221[0]
    app.PrintInfo(gov_GA_GI_122)

    gov_GA_GII_2111 = user_defined_model.SearchObject("gov_GA_GI_211.BlkDef")
    gov_GA_GI_211 = gov_GA_GII_2111[0]
    app.PrintInfo(gov_GA_GI_211)

    gov_GA_GII_2121 = user_defined_model.SearchObject("gov_GA_GI_212.BlkDef")
    gov_GA_GI_212 = gov_GA_GII_2121[0]
    app.PrintInfo(gov_GA_GI_212)

    gov_GA_GII_2211 = user_defined_model.SearchObject("gov_GA_GI_221.BlkDef")
    gov_GA_GI_221 = gov_GA_GII_2211[0]
    app.PrintInfo(gov_GA_GI_221)

    gov_GA_GII_2221 = user_defined_model.SearchObject("gov_GA_GI_222.BlkDef")
    gov_GA_GI_222 = gov_GA_GII_2221[0]
    app.PrintInfo(gov_GA_GI_222)


    gov_GHH = user_defined_model.SearchObject("gov_GH.BlkDef")
    gov_GH = gov_GHH[0]
    app.PrintInfo(gov_GH)
    
    for line in bpa_str_ar:
        line = line.rstrip('\n')
        if line == "": continue        
#        line = line + ' '*(80-len(line))

        #G卡
        if line[0] == 'G' and line[1] != ' ' and line[1] != '+' and line[4] != ' ' and line[5] != ' ':
            line = line + ' '*(80-len(line))
            chinese_count = 0
            #判断中文的个数
            for i in range (3,10):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count = chinese_count + 1
            name = line[3:11-chinese_count].strip()
            if isfloat(line[11-chinese_count:15-chinese_count]):base = float(line[11-chinese_count:15-chinese_count])
            #名称+电压
            Generator = name + str(base)                
            #找到节点名对应的编号
            if Generator in bus_name_ar:
                Generator_num = bus_name_ar[Generator]
                #找到对应的comp模块
                Generator_comp_name = str(Generator_num) + '_comp'
                Generator_compp = Net.SearchObject(str(Generator_comp_name))
                Generator_comp = Generator_compp[0]

                if not Generator_comp == None:       
                    #avr.dsl
    #                app.PrintInfo(Generator_num)
                    Generator_gov_dsll = Generator_comp.SearchObject('gov_dsl.ElmDsl')
                    Generator_gov_dsl = Generator_gov_dsll[0]
                    if Generator_gov_dsl == None:
                        Generator_comp.CreateObject('ElmDsl','gov_dsl')
                        Generator_gov_dsll = Generator_comp.SearchObject('gov_dsl.ElmDsl')
                        Generator_gov_dsl = Generator_gov_dsll[0]

#                    app.PrintInfo(Generator_pss_dsl)

                    Generator_comp.Gov= Generator_gov_dsl
                    sym = Generator_comp.Sym

                    sym_typ = sym.typ_id
                    MVABASE = sym_typ.sgn                    

                    #GS
                    if line[1] == 'S':
                        Generator_comp_gov = Generator_comp.Gov
                        Generator_comp_gov.typ_id = gov_GS
                        gov_type = Generator_comp.Gov

                        T2 = 0
                        T1 = 0
                        K = 0
                        T3 = 0
                        e = 0
                        PMIN = -9999
                        PDOWN = -9999
                        PMAX = 9999
                        PUP = 9999


#?????????????????????????????????功率基准值
                        
                        if isfloat(line[16-chinese_count:22-chinese_count]):
                            PMAX = getfloatvalue(line[16-chinese_count:22-chinese_count],1)
                        
                        if isfloat(line[22-chinese_count:28-chinese_count]):
                            PMIN = getfloatvalue(line[22-chinese_count:28-chinese_count],3)

                        if isfloat(line[28-chinese_count:33-chinese_count]):
                            R = getfloatvalue(line[28-chinese_count:33-chinese_count],3)

                        if isfloat(line[33-chinese_count:38-chinese_count]):
                            T1 = getfloatvalue(line[33-chinese_count:38-chinese_count],3)

                        if isfloat(line[38-chinese_count:43-chinese_count]):
                            T2 = getfloatvalue(line[38-chinese_count:43-chinese_count],3)

                        if isfloat(line[43-chinese_count:48-chinese_count]):
                            T3 = getfloatvalue(line[43-chinese_count:48-chinese_count],3)

                        if isfloat(line[48-chinese_count:54-chinese_count]):
                            VELOPEN = getfloatvalue(line[48-chinese_count:54-chinese_count],1)

                        if isfloat(line[54-chinese_count:60-chinese_count]):
                            VELCLOSE = getfloatvalue(line[54-chinese_count:60-chinese_count],1)

                        if isfloat(line[60-chinese_count:66-chinese_count]):
                            e = getfloatvalue(line[60-chinese_count:66-chinese_count],5)

                        K = PMAX/R/MVABASE
                        PUP = VELOPEN*PMAX/MVABASE
                        PDOWN = -1*VELCLOSE*PMAX/MVABASE
                        PMAX = PMAX/MVABASE
                        PMIN = PMIN/MVABASE

                        gov_type.T1 = T1
                        gov_type.T2 = T2
                        gov_type.K = K
                        gov_type.T3 = T3
                        gov_type.e = e
                        gov_type.PMIN = PMIN
                        gov_type.PDOWN = PDOWN
                        gov_type.PMAX = PMAX
                        gov_type.PUP = PUP
#??????????????????????????????????????????????????????????????????????/
                    #GA
                    elif line[1] == 'A':

                        PE = 0
                        TC = 0
                        TO = 0
                        VELclose = 9999
                        VELopen = -9999
                        PMAX = 9999
                        PMIN = -9999
                        T1LVDT = 0
                        KP = 0
                        KD = 0
                        INTG_MAX = 9999
                        INTG_MIN = -9999
                        PID_MAX = 9999
                        PID_MIN = -9999
                        PGV_DELAY = 0

                        if isfloat(line[16-chinese_count:22-chinese_count]):
                            PE = getfloatvalue(line[16-chinese_count:22-chinese_count],2)

                        if isfloat(line[22-chinese_count:26-chinese_count]):
                            TC = getfloatvalue(line[22-chinese_count:26-chinese_count],2)

                        if isfloat(line[26-chinese_count:30-chinese_count]):
                            TO = getfloatvalue(line[26-chinese_count:30-chinese_count],2)

                        if isfloat(line[30-chinese_count:34-chinese_count]):
                            VELclose = getfloatvalue(line[30-chinese_count:34-chinese_count],2)

                        if isfloat(line[34-chinese_count:38-chinese_count]):
                            VELopen = getfloatvalue(line[34-chinese_count:38-chinese_count],2)

                        if isfloat(line[38-chinese_count:42-chinese_count]):
                            PMAX = getfloatvalue(line[38-chinese_count:42-chinese_count],2)

                        if isfloat(line[42-chinese_count:46-chinese_count]):
                            PMIN = getfloatvalue(line[42-chinese_count:46-chinese_count],2)

                        if isfloat(line[46-chinese_count:50-chinese_count]):
                            T2 = getfloatvalue(line[46-chinese_count:50-chinese_count],2)

                        if isfloat(line[50-chinese_count:54-chinese_count]):
                            KP = getfloatvalue(line[50-chinese_count:54-chinese_count],2)

                        if isfloat(line[54-chinese_count:58-chinese_count]):
                            KD = getfloatvalue(line[54-chinese_count:58-chinese_count],2)

                        if isfloat(line[58-chinese_count:62-chinese_count]):
                            KI = getfloatvalue(line[58-chinese_count:62-chinese_count],2)
                            
                        if isfloat(line[62-chinese_count:66-chinese_count]):
                            INTG_MAX = getfloatvalue(line[62-chinese_count:66-chinese_count],2)
                            if INTG_MAX == 0:
                                INTG_MAX = 9999

                        if isfloat(line[66-chinese_count:70-chinese_count]):
                            INTG_MIN = getfloatvalue(line[66-chinese_count:70-chinese_count],2)
                            if INTG_MIN == 0:
                                INTG_MIN == -9999

                        if isfloat(line[70-chinese_count:74-chinese_count]):
                            PID_MAX = getfloatvalue(line[70-chinese_count:74-chinese_count],2)
                            app.PrintInfo(PID_MAX)
                            if PID_MAX == 0:
                                PID_MAX = 9999

                        if isfloat(line[74-chinese_count:78-chinese_count]):
                            PID_MIN = getfloatvalue(line[74-chinese_count:78-chinese_count],2)
                            if PID_MIN == 0:
                                PID_MIN = -9999

                        for line in bpa_str_ar:
                            line = line.rstrip('\n')
                            if line == "": continue        

                            #GA+卡
                            if line[0] == 'G' and line[1] == 'A' and line[2] == '+':
                                line = line + ' '*(80-len(line))
                                chinese_count = 0

                            #判断中文的个数
                                for i in range (3,10):
                                    if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                                        chinese_count = chinese_count + 1
                                name = line[3:11-chinese_count].strip()
                                if isfloat(line[11-chinese_count:15-chinese_count]):base = float(line[11-chinese_count:15-chinese_count])
                                #名称+电压
                                Generator = name + str(base)                
                                #找到节点名对应的编号
                                if Generator in bus_name_ar:
                                    Generator_num1 = bus_name_ar[Generator]
                                    
                                    if Generator_num1 == Generator_num:

                                        if isfloat(line[16-chinese_count:21-chinese_count]):
                                            PGV_DELAY = getfloatvalue(line[16-chinese_count:21-chinese_count],0)

                                        Generator_comp_gov = Generator_comp.Gov
                                        if Generator_comp_gov.typ_id == gov_GA_GI:
                                            gov_type.params[17] = PGV_DELAY
                                        else:
                                            app.PrintInfo('GA+')

                    
                            #GI卡
                            elif line[0] == 'G' and line[1] == 'I' and line[2] != '+':
                                line = line + ' '*(80-len(line))
                                chinese_count = 0

                            #判断中文的个数
                                for i in range (3,10):
                                    if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                                        chinese_count = chinese_count + 1
                                name = line[3:11-chinese_count].strip()
                                if isfloat(line[11-chinese_count:15-chinese_count]):base = float(line[11-chinese_count:15-chinese_count])
                                #名称+电压
                                Generator = name + str(base)                
                                #找到节点名对应的编号
                                if Generator in bus_name_ar:
                                    Generator_num1 = bus_name_ar[Generator]
                                    
                                    if Generator_num1 == Generator_num:
#                                      app.PrintInfo('222222222222222222222222222222222222222')


                                        Generator_comp_gov = Generator_comp.Gov
                                        if Generator_comp_gov.typ_id == gov_GA_GI:
                                            gov_type = Generator_comp.Gov
                                        else:
                                            Generator_comp_gov.typ_id = gov_GA_GI
                                            gov_type = Generator_comp.Gov

                                        T1 = 0
                                        e = 0
                                        K = 0
                                        fuhezidong = 0
                                        KP1 = 0
                                        KD1 = 0
                                        KI1 = 0
                                        INTG_MAX1 = 9999
                                        INTG_MIN1 = -9999
                                        PID_MAX1 = 9999
                                        PID_MIN1 = -9999
                                        fuheqiankui = 0
                                        FIRST_MAX = 9999
                                        FIRSR_MIN = -9999
        
                                        if isfloat(line[16-chinese_count:21-chinese_count]):
                                            T1 = getfloatvalue(line[16-chinese_count:21-chinese_count],3)
                                                     
                                        if isfloat(line[21-chinese_count:27-chinese_count]):
                                            e = getfloatvalue(line[21-chinese_count:27-chinese_count],4)
                                            e = e/2
 
                                        if isfloat(line[27-chinese_count:32-chinese_count]):
                                            K = getfloatvalue(line[27-chinese_count:32-chinese_count],2)

                                        if isfloat(line[32-chinese_count]):
                                            fuhezidong = int(line[32-chinese_count])
                                            if fuhezidong == 2:
                                                fuhezidong = 0#0对应切除

                                        if isfloat(line[33-chinese_count:38-chinese_count]):
                                            KP1 = getfloatvalue(line[33-chinese_count:38-chinese_count],3)
                                             
                                        if isfloat(line[38-chinese_count:43-chinese_count]):
                                            KD1 = getfloatvalue(line[38-chinese_count:43-chinese_count],3)
       
                                        if isfloat(line[43-chinese_count:48-chinese_count]):
                                            KI1 = getfloatvalue(line[43-chinese_count:48-chinese_count],3)
  
                                        if isfloat(line[48-chinese_count:53-chinese_count]):
                                            INTG_MAX1 = getfloatvalue(line[48-chinese_count:53-chinese_count],3)
                                            if INTG_MAX1 == 0:
                                                INTG_MAX1 = 9999
  
                                        if isfloat(line[53-chinese_count:58-chinese_count]):
                                            INTG_MIN1 = getfloatvalue(line[53-chinese_count:58-chinese_count],3)
                                            if INTG_MIN1 == 0:
                                                INTG_MIN1 = -9999
          
                                        if isfloat(line[58-chinese_count:63-chinese_count]):
                                            PID_MAX1 = getfloatvalue(line[58-chinese_count:63-chinese_count],3)
                                            if PID_MAX1 == 0:
                                                PID_MAX1 = 9999

                                        if isfloat(line[63-chinese_count:68-chinese_count]):
                                            PID_MIN1 = getfloatvalue(line[63-chinese_count:68-chinese_count],3)
                                            if PID_MAX1 == 0:
                                                PID_MAX1 = -9999
    
                                        if isfloat(line[68-chinese_count]):
                                            fuheqiankui = int(line[68-chinese_count])
                                            if fuheqiankui == 2:
                                                fuheqiankui = 0

                                        if isfloat(line[69-chinese_count:74-chinese_count]):
                                            FIRST_MAX = getfloatvalue(line[69-chinese_count:74-chinese_count],3)
                                              
                                        if isfloat(line[74-chinese_count:79-chinese_count]):
                                            FIRST_MIN = getfloatvalue(line[74-chinese_count:79-chinese_count],3)

                                        gov_type.e = e
                                        gov_type.T2 = T2
                                        gov_type.K = K
                                        gov_type.T1 = T1
                                        gov_type.KP1 = KP1
                                        gov_type.KD1 = KD1
                                        gov_type.KI1 = KI1
                                        gov_type.PE = PE
                                        gov_type.KP = KP
                                        gov_type.KD = KD
                                        gov_type.KI = KI
                                        gov_type.fuheqiankui = fuheqiankui
                                        gov_type.fuhezidong = fuhezidong
                                        gov_type.TO = TO
                                        gov_type.TC = TC
                                        gov_type.PGV_DELAY = PGV_DELAY
                                        gov_type.INTG_MIN1 = INTG_MIN1
                                        gov_type.PID_MIN1 = PID_MIN1
                                        gov_type.FIRST_MIN = FIRST_MIN
                                        gov_type.INTG_MIN = INTG_MIN
                                        gov_type.PID_MIN = PID_MIN
                                        gov_type.VELclose = VELclose
                                        gov_type.PMIN = PMIN
                                        gov_type.INTG_MAX1 = INTG_MAX1
                                        gov_type.PID_MAX1 = PID_MAX1
                                        gov_type.FIRST_MAX = FIRST_MAX
                                        gov_type.INTG_MAX = INTG_MAX
                                        gov_type.PID_MAX = PID_MAX
                                        gov_type.VELopen = VELopen
                                        gov_type.PMAX = PMAX

 
                            #GI+卡
                            elif line[0] == 'G' and line[1] == 'I' and line[2] == '+':
                                line = line + ' '*(80-len(line))
                                chinese_count = 0

                            #判断中文的个数
                                for i in range (3,10):
                                    if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                                        chinese_count = chinese_count + 1
                                name = line[3:11-chinese_count].strip()
                                if isfloat(line[11-chinese_count:15-chinese_count]):base = float(line[11-chinese_count:15-chinese_count])
                                #名称+电压
                                Generator = name + str(base)                
                                #找到节点名对应的编号
                                if Generator in bus_name_ar:
                                    Generator_num1 = bus_name_ar[Generator]
                                    
                                    if Generator_num1 == Generator_num:
#                                 app.PrintInfo('222222222222222222222222222222222222222')                                    

                                        Generator_comp_gov = Generator_comp.Gov
                                        if Generator_comp_gov.typ_id == gov_GA_GI:
                                            gov_type = Generator_comp.Gov
                                        else:
                                            Generator_comp_gov.typ_id = gov_GA_GI
                                            gov_type = Generator_comp.Gov

                                        app.PrintInfo(gov_type)
                                        app.PrintInfo(Generator_comp_gov.typ_id)
                                        app.PrintInfo(Generator_comp.Gov)

                                        yalikongzhi = 0
                                        KP2 = 0
                                        KD2 = 0
                                        KI2 = 0
                                        INTG_MAX2 = 9999
                                        INTG_MIN2 = -9999
                                        PID_MAX2 = 9999
                                        PID_MIN2 = -9999
                                        CON_MAX = 9999
                                        CON_MIN = -9999

                                        if isfloat(line[16-chinese_count]):
                                            yalikongzhi = int(line[16-chinese_count])
                                            if yalikongzhi == 2:
                                                yalikongzhi = 0

                                        if isfloat(line[17-chinese_count:22-chinese_count]):
                                            KP2 = getfloatvalue(line[17-chinese_count:22-chinese_count],3)
                                                     
                                        if isfloat(line[22-chinese_count:27-chinese_count]):
                                            KD2 = getfloatvalue(line[22-chinese_count:27-chinese_count],3)
 
                                        if isfloat(line[27-chinese_count:32-chinese_count]):
                                            KI2 = getfloatvalue(line[27-chinese_count:32-chinese_count],3)

                                        if isfloat(line[32-chinese_count:37-chinese_count]):
                                            INTG_MAX2 = getfloatvalue(line[32-chinese_count:37-chinese_count],3)
                                            if INTG_MAX2 == 0:
                                                INTG_MAX2 = 9999
                                                     
                                        if isfloat(line[37-chinese_count:42-chinese_count]):
                                            INTG_MIN2 = getfloatvalue(line[37-chinese_count:42-chinese_count],3)
                                            if INTG_MIN2 == 0:
                                                INTG_MIN2 = -9999
       
                                        if isfloat(line[42-chinese_count:47-chinese_count]):
                                            PID_MAX2 = getfloatvalue(line[42-chinese_count:47-chinese_count],3)
                                            if PID_MAX2 == 0:
                                                PID_MAX2 = 9999
  
                                        if isfloat(line[47-chinese_count:52-chinese_count]):
                                            PID_MIN2 = getfloatvalue(line[47-chinese_count:52-chinese_count],3)
                                            if PID_MIN2 == 0:
                                                PID_MIN2 = -9999
  
                                        if isfloat(line[52-chinese_count:57-chinese_count]):
                                            CON_MAX = getfloatvalue(line[52-chinese_count:57-chinese_count],3)
                                            if CON_MAX == 0:
                                                CON_MAX = 9999
                                      
                                        if isfloat(line[57-chinese_count:62-chinese_count]):
                                            CON_MIN = getfloatvalue(line[57-chinese_count:62-chinese_count],3)
                                            if CON_MIN == 0:
                                                CON_MIN = -9999


                                        gov_type.KP2 = KP2
                                        gov_type.KD2 = KD2
                                        gov_type.KI2 = KI2
                                        gov_type.yalikongzhi = yalikongzhi
                                        gov_type.INTG_MAX2 = INTG_MAX2
                                        gov_type.PID_MAX2 = PID_MAX2
                                        gov_type.CON_MAX = CON_MAX
                                        gov_type.INTG_MIN2 = INTG_MIN2
                                        gov_type.PID_MIN2 = PID_MIN2
                                        gov_type.CON_MIN = CON_MIN


                                        e = gov_type.e
                                        T2 = gov_type.T2
                                        K = gov_type.K
                                        T1 = gov_type.T1
                                        KP1 = gov_type.KP1
                                        KD1 = gov_type.KD1
                                        KI1 = gov_type.KI1
                                        PE = gov_type.PE
                                        KP = gov_type.KP
                                        KD = gov_type.KD
                                        KI = gov_type.KI
                                        fuheqiankui = gov_type.fuheqiankui
                                        fuhezidong = gov_type.fuhezidong
                                        TO = gov_type.TO
                                        TC = gov_type.TC
                                        PGV_DELAY = gov_type.PGV_DELAY
                                        INTG_MIN1 = gov_type.INTG_MIN1
                                        PID_MIN1 = gov_type.PID_MIN1
                                        FIRST_MIN = gov_type.FIRST_MIN 
                                        INTG_MIN = gov_type.INTG_MIN
                                        PID_MIN = gov_type.PID_MIN 
                                        VELclose = gov_type.VELclose
                                        PMIN = gov_type.PMIN
                                        INTG_MAX1 = gov_type.INTG_MAX1
                                        PID_MAX1 = gov_type.PID_MAX1 
                                        FIRST_MAX = gov_type.FIRST_MAX 
                                        INTG_MAX = gov_type.INTG_MAX 
                                        PID_MAX = gov_type.PID_MAX
                                        VELopen = gov_type.VELopen
                                        PMAX = gov_type.PMAX

                                        if fuhezidong == 1 and fuheqiankui == 1 and yalikongzhi == 1:
                                            Generator_comp_gov = Generator_comp.Gov
                                            Generator_comp_gov.typ_id = gov_GA_GI_111
                                            gov_type = Generator_comp.Gov
                                            raise Exception("GA_GI_111.")
                                                
                                        elif fuhezidong == 1 and fuheqiankui == 1 and yalikongzhi == 0:
                                            Generator_comp_gov = Generator_comp.Gov
                                            Generator_comp_gov.typ_id = gov_GA_GI_111
                                            gov_type = Generator_comp.Gov
                                            raise Exception("GA_GI_112.")

                                        elif fuhezidong == 1 and fuheqiankui == 0 and yalikongzhi == 1:
                                            Generator_comp_gov = Generator_comp.Gov
                                            Generator_comp_gov.typ_id = gov_GA_GI_111
                                            gov_type = Generator_comp.Gov
                                            raise Exception("GA_GI_121.")

                                        elif fuhezidong == 1 and fuheqiankui == 0 and yalikongzhi == 0:
                                            Generator_comp_gov = Generator_comp.Gov
                                            Generator_comp_gov.typ_id = gov_GA_GI_122
                                            gov_type = Generator_comp.Gov

                                            gov_type.e = e
                                            gov_type.PE = PE/MVABASE
                                            gov_type.PGV_DELAY = PGV_DELAY
                                            gov_type.K = K
                                            gov_type.T1 = T1
                                            gov_type.KP1 = KP1
                                            gov_type.KD1 = KD1
                                            gov_type.KI1 = KI1
                                            gov_type.KP = KP
                                            gov_type.KD = KD
                                            gov_type.KI = KI
                                            gov_type.TO = TO
                                            gov_type.TC = TC
                                            gov_type.T2 = T2
                                            gov_type.INTG_MIN1 = INTG_MIN1
                                            gov_type.PID_MIN1 = PID_MIN1
                                            gov_type.CON_MIN = CON_MIN
                                            gov_type.INTG_MIN = INTG_MIN
                                            gov_type.PID_MIN = PID_MIN
                                            gov_type.VELclose = VELclose
                                            gov_type.PMIN = PMIN
                                            gov_type.INTG_MAX1 = INTG_MAX1
                                            gov_type.PID_MAX1 = PID_MAX1
                                            gov_type.CON_MAX = CON_MAX
                                            gov_type.params[24] = INTG_MAX
                                            gov_type.params[25] = PID_MAX
                                            gov_type.params[26] = VELopen
                                            gov_type.params[27] = PMAX
                                        

                                        elif fuhezidong == 0 and fuheqiankui == 1 and yalikongzhi == 1:
                                            Generator_comp_gov = Generator_comp.Gov
                                            Generator_comp_gov.typ_id = gov_GA_GI_111
                                            gov_type = Generator_comp.Gov
                                            raise Exception("GA_GI_211.")

                                        elif fuhezidong == 0 and fuheqiankui == 1 and yalikongzhi == 0:
                                            Generator_comp_gov = Generator_comp.Gov
                                            Generator_comp_gov.typ_id = gov_GA_GI_111
                                            gov_type = Generator_comp.Gov
                                            raise Exception("GA_GI_212.")

                                        elif fuhezidong == 0 and fuheqiankui == 0 and yalikongzhi == 1:
                                            Generator_comp_gov = Generator_comp.Gov
                                            Generator_comp_gov.typ_id = gov_GA_GI_111
                                            gov_type = Generator_comp.Gov
                                            raise Exception("GA_GI_221.")
                                        
                                        elif fuhezidong == 0 and fuheqiankui == 0 and yalikongzhi == 0:
                                            Generator_comp_gov = Generator_comp.Gov
                                            Generator_comp_gov.typ_id = gov_GA_GI_111
                                            gov_type = Generator_comp.Gov
                                            raise Exception("GA_GI_222.")



                    #GJ卡
                    elif line[0] == 'G' and line[1] == 'J' and line[2] != '+':
                        line = line + ' '*(80-len(line))
                        chinese_count = 0

                    #判断中文的个数
                        for i in range (3,10):
                            if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                                chinese_count = chinese_count + 1
                        name = line[3:11-chinese_count].strip()
                        if isfloat(line[11-chinese_count:15-chinese_count]):base = float(line[11-chinese_count:15-chinese_count])
                        #名称+电压
                        Generator = name + str(base)                
                        #找到节点名对应的编号
                        if Generator in bus_name_ar:
                            Generator_num1 = bus_name_ar[Generator]
                                    
                            if Generator_num1 == Generator_num:
                                app.PrintInfo('GJ')
                                raise Exception("No project activated. Python Script stopped.")




                    #GJ+卡
                    elif line[0] == 'G' and line[1] == 'J' and line[2] == '+':
                        line = line + ' '*(80-len(line))
                        chinese_count = 0

                    #判断中文的个数
                        for i in range (3,10):
                            if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                                chinese_count = chinese_count + 1
                        name = line[3:11-chinese_count].strip()
                        if isfloat(line[11-chinese_count:15-chinese_count]):base = float(line[11-chinese_count:15-chinese_count])
                        #名称+电压
                        Generator = name + str(base)                
                        #找到节点名对应的编号
                        if Generator in bus_name_ar:
                            Generator_num1 = bus_name_ar[Generator]
                                    
                            if Generator_num1 == Generator_num:
                                app.PrintInfo('GJ+')
                                raise Exception("No project activated. Python Script stopped.")



                    #GK卡
                    elif line[0] == 'G' and line[1] == 'K':
                        line = line + ' '*(80-len(line))
                        chinese_count = 0

                    #判断中文的个数
                        for i in range (3,10):
                            if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                                chinese_count = chinese_count + 1
                        name = line[3:11-chinese_count].strip()
                        if isfloat(line[11-chinese_count:15-chinese_count]):base = float(line[11-chinese_count:15-chinese_count])
                        #名称+电压
                        Generator = name + str(base)                
                        #找到节点名对应的编号
                        if Generator in bus_name_ar:
                            Generator_num1 = bus_name_ar[Generator]
                                    
                            if Generator_num1 == Generator_num:
                                app.PrintInfo('GK')
                                raise Exception("No project activated. Python Script stopped.")


                    #GH卡
                    elif line[0] == 'G' and line[1] == 'H':
#                        app.PrintInfo('2222222222222222222222')

                        line = line + ' '*(80-len(line))
                        chinese_count = 0

                    #判断中文的个数
                        for i in range (3,10):
                            if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                                chinese_count = chinese_count + 1
                        name = line[3:11-chinese_count].strip()
                        if isfloat(line[11-chinese_count:15-chinese_count]):base = float(line[11-chinese_count:15-chinese_count])
                        #名称+电压
                        Generator = name + str(base)                
                        #找到节点名对应的编号
                        if Generator in bus_name_ar:
                            Generator_num1 = bus_name_ar[Generator]

                            Generator_comp_gov = Generator_comp.Gov
                            if Generator_comp_gov.typ_id == gov_GH:
                                gov_type = Generator_comp.Gov
                                sym = Generator_comp.Sym
                            else:
                                Generator_comp_gov.typ_id = gov_GH
                                gov_type = Generator_comp.Gov
                                sym = Generator_comp.Sym
#                            app.PrintInfo(gov_type)
                            sym_typ = sym.typ_id
                            MVABASE = sym_typ.sgn

                            PMAX = 9999
                            R = 0
                            TG = 0.3
                            TP = 0.04
                            TD = 10
                            TW = 2
                            VELCLOSE = 0.15
                            VELOPEN = 0.15
                            Dd = 0
                            e = 0
                            PMIN = 0

                            if isfloat(line[16-chinese_count:22-chinese_count]):
                                PMAX = getfloatvalue(line[16-chinese_count:22-chinese_count],1)
                                        
                            if isfloat(line[22-chinese_count:27-chinese_count]):
                                R = getfloatvalue(line[22-chinese_count:27-chinese_count],3)
                                        
                            if isfloat(line[27-chinese_count:32-chinese_count]):
                                TG = getfloatvalue(line[27-chinese_count:32-chinese_count],3)

                            if isfloat(line[32-chinese_count:37-chinese_count]):
                                TP = getfloatvalue(line[32-chinese_count:37-chinese_count],3)

                            if isfloat(line[37-chinese_count:42-chinese_count]):
                                TD = getfloatvalue(line[37-chinese_count:42-chinese_count],3)

                            if isfloat(line[42-chinese_count:47-chinese_count]):
                                TW_2 = getfloatvalue(line[42-chinese_count:47-chinese_count],3)
                                TW = TW_2 * 2

                            if isfloat(line[47-chinese_count:52-chinese_count]):
                                VELCLOSE = getfloatvalue(line[47-chinese_count:52-chinese_count],3)

                            if isfloat(line[52-chinese_count:57-chinese_count]):
                                VELOPEN = getfloatvalue(line[52-chinese_count:57-chinese_count],3)

                            if isfloat(line[57-chinese_count:62-chinese_count]):
                                Dd = getfloatvalue(line[57-chinese_count:62-chinese_count],3)

                            if isfloat(line[62-chinese_count:68-chinese_count]):
                                e = getfloatvalue(line[62-chinese_count:68-chinese_count],5)

                            PMAX_pu = PMAX/MVABASE
                            PDOWN = -abs(1*VELCLOSE)*PMAX_pu
                            PUP = VELOPEN*PMAX_pu
                                    
                            gov_type.R = R
                            gov_type.e = e
                            gov_type.Pmax_pu = PMAX_pu
                            gov_type.TG = TG
                            gov_type.TP = TP
                            gov_type.TW = TW
                            gov_type.Dd = Dd
                            gov_type.Td = TD
                            gov_type.PDOWN = PDOWN
                            gov_type.Pmin = PMIN
                            gov_type.PUP = PUP
                            gov_type.Pmaxpu = PMAX_pu

                                    



# ----------------------------------------------------------------------------------------------------
def GetTCard(bpa_file, bpa_str_ar,bus_name_ar):
    #原动机T卡
    prj = app.GetActiveProject()
    if prj is None:
        raise Exception("No project activated. Python Script stopped.")

    app.PrintInfo(prj)
    Network_Modell = prj.SearchObject("Network Model.IntPrjfolder")
    Network_Model=Network_Modell[0]

    Network_Dataa = Network_Model.SearchObject("Network Data.IntPrjfolder")
    Network_Data = Network_Dataa[0]
    app.PrintInfo(Network_Data)

    Nett = Network_Data.SearchObject("*.ElmNet")
    Net = Nett[0]
    app.PrintInfo(Net)

    Libraryy = prj.SearchObject("Library.IntPrjfolder")
    Library = Libraryy[0]
    app.PrintInfo(Library)

    user_defined_modell = Library.SearchObject("User Defined Models.IntPrjfolder")
    user_defined_model = user_defined_modell[0]
    app.PrintInfo(user_defined_model)

    BPA_Frame_EE = user_defined_model.SearchObject("BPA Frame(E).BlkDef")
    BPA_Frame_E = BPA_Frame_EE[0]
    app.PrintInfo(BPA_Frame_E)

    BPA_Frame_FF = user_defined_model.SearchObject("BPA Frame(F).BlkDef")
    BPA_Frame_F = BPA_Frame_FF[0]
    app.PrintInfo(BPA_Frame_F)

    motor_TAA = user_defined_model.SearchObject("motor_TA.BlkDef")
    motor_TA = motor_TAA[0]
    app.PrintInfo(motor_TA)
    
    motor_TBB = user_defined_model.SearchObject("motor_TB.BlkDef")
    motor_TB = motor_TBB[0]
    app.PrintInfo(motor_TB)
    
    for line in bpa_str_ar:
        line = line.rstrip('\n')
        if line == "": continue        

        #T卡
        if line[0] == 'T':
            line = line + ' '*(80-len(line))
            chinese_count = 0
            #判断中文的个数
            for i in range (3,10):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count = chinese_count + 1
            name = line[3:11-chinese_count].strip()
            if isfloat(line[11-chinese_count:15-chinese_count]):base = float(line[11-chinese_count:15-chinese_count])
            #名称+电压
            Generator = name + str(base)                
            #找到节点名对应的编号
            if Generator in bus_name_ar:
                Generator_num = bus_name_ar[Generator]

                #找到对应的comp模块
                Generator_comp_name = str(Generator_num) + '_comp'
                Generator_compp = Net.SearchObject(str(Generator_comp_name))
                Generator_comp = Generator_compp[0]

                if not Generator_comp == None:       
                    #avr.dsl
                    Generator_motor_dsll = Generator_comp.SearchObject('motor.ElmDsl')
                    Generator_motor_dsl = Generator_motor_dsll[0]
                    if Generator_motor_dsl == None:
                        Generator_comp.CreateObject('ElmDsl','Motor')
                        Generator_motor_dsll = Generator_comp.SearchObject('motor.ElmDsl')
                        Generator_motor_dsl = Generator_motor_dsll[0]                


                    Generator_comp.Prime_Moter = Generator_motor_dsl
                    sym = Generator_comp.Sym


                    sym_typ = sym.typ_id
                    MVABASE = sym_typ.sgn                    

                    #TB
                    if line[1] == 'B':
                        Generator_comp_motor = Generator_comp.Prime_Moter
                        Generator_comp_motor.typ_id = motor_TB
                        motor_type = Generator_comp.Prime_Moter

                        TCH = 0
                        TRH = 0
                        TCO = 0 
                        lambda_ = 0
                        FHP = 0
                        FIP = 0
                        FLp = 0

                        if isfloat(line[16-chinese_count:21-chinese_count]):
                            TCH = getfloatvalue(line[16-chinese_count:21-chinese_count],3)
                        
                        if isfloat(line[21-chinese_count:26-chinese_count]):
                            FHP = getfloatvalue(line[21-chinese_count:26-chinese_count],3)

                        if isfloat(line[31-chinese_count:36-chinese_count]):
                            TRH = getfloatvalue(line[31-chinese_count:36-chinese_count],3)

                        if isfloat(line[36-chinese_count:41-chinese_count]):
                            FIP = getfloatvalue(line[36-chinese_count:41-chinese_count],3)

                        if isfloat(line[46-chinese_count:51-chinese_count]):
                            TCO = getfloatvalue(line[46-chinese_count:51-chinese_count],3)

                        if isfloat(line[51-chinese_count:56-chinese_count]):
                            FLP = getfloatvalue(line[51-chinese_count:56-chinese_count],3)

                        if isfloat(line[76-chinese_count:80-chinese_count]):
                            lambda_ = getfloatvalue(line[76-chinese_count:80-chinese_count],2)

                        motor_type.TCH = TCH
                        motor_type.TRH = TRH
                        motor_type.TCO = TCO
                        motor_type.lambda_ = lambda_
                        motor_type.FHP = FHP
                        motor_type.FIP = FIP
                        motor_type.FLP = FLP

                    #TA                
                    if line[1] == 'A':
                        Generator_comp_motor = Generator_comp.Prime_Moter
                        Generator_comp_motor.typ_id = motor_TA
                        motor_type = Generator_comp.Prime_Moter

                        TCH = 0
                        if isfloat(line[16-chinese_count:21-chinese_count]):
                            TCH = getfloatvalue(line[16-chinese_count:21-chinese_count],3)

                        motor_type.TCH = TCH                        
    
# ----------------------------------------------------------------------------------------------------
def GetLNCard(bpa_file, bpa_str_ar,bus_name_ar):
    #当前工程
    prj = app.GetActiveProject()
    if prj is None:
        raise Exception("No project activated. Python Script stopped.")

    app.PrintInfo(prj)
    Network_Modell = prj.SearchObject("Network Model.IntPrjfolder")
    Network_Model=Network_Modell[0]

    Network_Dataa = Network_Model.SearchObject("Network Data.IntPrjfolder")
    Network_Data = Network_Dataa[0]
    app.PrintInfo(Network_Data)

    Nett = Network_Data.SearchObject("*.ElmNet")
    Net = Nett[0]
    app.PrintInfo(Net)

    Libraryy = prj.SearchObject("Library.IntPrjfolder")
    Library = Libraryy[0]
    app.PrintInfo(Library)

    user_defined_modell = Library.SearchObject("User Defined Models.IntPrjfolder")
    user_defined_model = user_defined_modell[0]
    app.PrintInfo(user_defined_model)

    Equipment_Type_Libraryy = Library.SearchObject("Equipment Type Library.IntPrjfolder")
    Equipment_Type_Library = Equipment_Type_Libraryy[0]
    app.PrintInfo(Equipment_Type_Library)

#    folder_Study_Casess = Library.SearchObject("Study Cases.IntPrjfolder")
#    folder_Study_Cases = folder_Study_Casess[0]
#    app.PrintInfo(folder_Study_Cases)

#    Study_Cases = folder_Study_Cases.SearchObject("Study Case.IntCase")
#    Study_Case = Study_Cases[0]
#    app.PrintInfo(Study_Case)


#    Load_Flow_Calculationn = Study_Case.SearchObject("Load Flow Calculation.ComLdf")
#    Load_Flow_Calculation = Load_Flow_Calculationn[0]
#    app.PrintInfo(Load_Flow_Calculation)

    Generator_num_str = []
    Active_power_str = []
    Reactive_power_str = []

#潮流计算
    #retrieve load-flow object
    ldf = app.GetFromStudyCase("ComLdf")
    #force balanced load flow
    ldf.iopt_net = 0
    #考虑无功限制
    app.PrintInfo(ldf.iopt_lim)
    ldf.iopt_lim = 1
    #execute load flow
    ldf.Execute()

##计算一遍后project内才会出现Load Flow Calculation的选项
#    Load_Flow_Calculation.ipot_lim = 1




    for line in bpa_str_ar:
        line = line.rstrip('\n')
        if line == "": continue

        line = line + ' '*(80-len(line))

        #LN卡
        if line[0] == 'L' and line[1] == 'N':
            chinese_count1 = 0
            #判断中文的个数
            for i in range (3,10):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count1 = chinese_count1 + 1
            name = line[3:11-chinese_count1].strip()
            base = float(line[11-chinese_count1:15-chinese_count1])
            #名称+电压
            Generator = name + str(base)
            #找到节点名对应的编号
            Generator_num = 0
            if Generator in bus_name_ar:
                Generator_num = bus_name_ar[Generator]

                Generator_name = 'sym_' + str(Generator_num) + '_1'
              
                #选中发电机
                symm = Net.SearchObject(str(Generator_name))
                sym = symm[0]
                if sym != None:

                    app.PrintInfo(sym)

                    Generator_num_str.append(Generator_num)

                    Active_power = sym.GetAttribute("m:P:bus1")
                    Reactive_power = sym.GetAttribute("m:Q:bus1")
                    app.PrintInfo(Active_power)
                    app.PrintInfo(Reactive_power)

                    Active_power_str.append(Active_power)
                    Reactive_power_str.append(Reactive_power)

                      



            #判断中文的个数
            chinese_count2 = chinese_count1
            for i in range (18-chinese_count1,30-chinese_count1):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count2 = chinese_count2 + 1
            name = line[18-chinese_count1:26-chinese_count2].strip()
            if isfloat(line[26-chinese_count2:30-chinese_count2]):base = float(line[26-chinese_count2:30-chinese_count2])
            #名称+电压
            Generator = name + str(base)
            #找到节点名对应的编号
            if Generator in bus_name_ar:
                Generator_num = bus_name_ar[Generator]

                Generator_name = 'sym_' + str(Generator_num) + '_1'
              
                #选中发电机
                symm = Net.SearchObject(str(Generator_name))
                sym = symm[0]
                if sym != None:
                    Generator_num_str.append(Generator_num)

                    Active_power = sym.GetAttribute("m:P:bus1")
                    Reactive_power = sym.GetAttribute("m:Q:bus1")
                    app.PrintInfo(Active_power)
                    app.PrintInfo(Reactive_power)

                    Active_power_str.append(Active_power)
                    Reactive_power_str.append(Reactive_power)


            #判断中文的个数
            chinese_count3 = chinese_count2
            for i in range (33-chinese_count2,45-chinese_count2):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count3 = chinese_count3 + 1
            name = line[33-chinese_count2:41-chinese_count3].strip()
            if isfloat(line[41-chinese_count3:45-chinese_count3]):base = float(line[41-chinese_count3:45-chinese_count3])
            #名称+电压
            Generator = name + str(base)
            #找到节点名对应的编号
            if Generator in bus_name_ar:
                Generator_num = bus_name_ar[Generator]
                
                Generator_name = 'sym_' + str(Generator_num) + '_1'
              
                #选中发电机
                symm = Net.SearchObject(str(Generator_name))
                sym = symm[0]
                if sym != None:
                    Generator_num_str.append(Generator_num)

                    Active_power = sym.GetAttribute("m:P:bus1")
                    Reactive_power = sym.GetAttribute("m:Q:bus1")
                    app.PrintInfo(Active_power)
                    app.PrintInfo(Reactive_power)

                    Active_power_str.append(Active_power)
                    Reactive_power_str.append(Reactive_power)

                


            #判断中文的个数
            chinese_count4 = chinese_count3
            for i in range (63-chinese_count3,75-chinese_count3):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count4 = chinese_count4 + 1
            name = line[63-chinese_count3:75-chinese_count4].strip()
            if isfloat(line[71-chinese_count4:75-chinese_count4]): base = float(line[71-chinese_count4:75-chinese_count4])
            #名称+电压
            Generator = name + str(base)
            #找到节点名对应的编号
            if Generator in bus_name_ar:
                Generator_num = bus_name_ar[Generator]
                
                Generator_name = 'sym_' + str(Generator_num) + '_1'
              
                #选中发电机
                symm = Net.SearchObject(str(Generator_name))
                sym = symm[0]
                if sym != None:
                    Generator_num_str.append(Generator_num)

                    Active_power = sym.GetAttribute("m:P:bus1")
                    Reactive_power = sym.GetAttribute("m:Q:bus1")
                    app.PrintInfo(Active_power)
                    app.PrintInfo(Reactive_power)

                    Active_power_str.append(Active_power)
                    Reactive_power_str.append(Reactive_power)



    Generator_length = len(Generator_num_str)
    for i in range (0,Generator_length):

        Generator_num = Generator_num_str[i]

        Generator_name = 'sym_' + str(Generator_num) + '_1'

        app.PrintInfo(Generator_name)
              
        #选中发电机
        symm = Net.SearchObject(str(Generator_name))
        sym = symm[0]

        app.PrintInfo(sym)

        sym.outserv = 1

        #找到对应的comp模块
        Generator_comp_name = str(Generator_num) + '_comp'
        Generator_compp = Net.SearchObject(str(Generator_comp_name))
        Generator_comp = Generator_compp[0]
        if not Generator_comp == None:
            app.PrintInfo('LN')
            app.PrintInfo(Generator_comp_name)
            Generator_comp.outserv = 1                    
        
        Lod_name ='sys_' + 'Lod_' + str(Generator_num) + '_1'
        Lod_full_name = Lod_name + '.ElmLod'

        Lodd = Net.SearchObject(Lod_full_name)
        Lod = Lodd[0]
        if Lod == None:
            Net.CreateObject('ElmLod',Lod_name)                
            Lodd = Net.SearchObject(Lod_full_name)
            Lod = Lodd[0]

        TypLod_name = 'TypLod' + Lod_name
        lod_typlodd = Net.SearchObject(TypLod_name)
        lod_typlod = lod_typlodd[0]
        if lod_typlod == None:
            Net.CreateObject('TypLod',TypLod_name)
            lod_typlodd = Net.SearchObject(TypLod_name)
            lod_typlod = lod_typlodd[0]
                    
        lod_typlod.phtech = 2
        Lod.typ_id = lod_typlod
                    
        bus_name = str(Generator_num) + ' ' + str(Generator_num) + '.ElmTerm'
        buss = Net.SearchObject(bus_name)
        bus = buss[0]

        Lod_StaCubic_fullname = Lod_name + '.StaCubic'
        Lod_StaCubicc = bus.SearchObject(Lod_StaCubic_fullname)
        Lod_StaCubic = Lod_StaCubicc[0]
        if Lod_StaCubic == None:
            bus.CreateObject('StaCubic',Lod_name)
            Lod_StaCubic_fullname = Lod_name + '.StaCubic'
            Lod_StaCubicc = bus.SearchObject(Lod_StaCubic_fullname)
            Lod_StaCubic = Lod_StaCubicc[0]
            
        Lod_StaCubic.obj_id = Lod
        Lod.bus1 = Lod_StaCubic
        Lod.plini = Active_power_str[i]*(-1)
        Lod.qlini = Reactive_power_str[i]*(-1)



# ----------------------------------------------------------------------------------------------------
def GetGenerator(bpa_file, bpa_str_ar,bus_name_ar,XFACT,TDODPS,TQODPS,TDODPH,TQODPH,IAMRTS,DMPALL,DMPMLT,NOSAT,MVABASE):

    Generator_num = 0
    MFCard_str = []
    MCard_str = []

    #当前工程
    prj = app.GetActiveProject()
    if prj is None:
        raise Exception("No project activated. Python Script stopped.")

    app.PrintInfo(prj)
    Network_Modell = prj.SearchObject("Network Model.IntPrjfolder")
    Network_Model=Network_Modell[0]

    Network_Dataa = Network_Model.SearchObject("Network Data.IntPrjfolder")
    Network_Data = Network_Dataa[0]
    app.PrintInfo(Network_Data)
    

    Nett = Network_Data.SearchObject("*.ElmNet")
    Net = Nett[0]
    app.PrintInfo(Net)

    Libraryy = prj.SearchObject("Library.IntPrjfolder")
    Library = Libraryy[0]
    app.PrintInfo(Library)

    Settingss = prj.SearchObject("Settings.SetFold")
    Settings = Settingss[0]
    app.PrintInfo(Settings)

    Input_Optionss = Settings.SearchObject("Input Options.IntOpt")
    Input_Options = Input_Optionss[0]
    app.PrintInfo(Input_Options)

    user_defined_modell = Library.SearchObject("User Defined Models.IntPrjfolder")
    user_defined_model = user_defined_modell[0]
    app.PrintInfo(user_defined_model)

    BPA_Frame_EE = user_defined_model.SearchObject("BPA Frame(E).BlkDef")
    BPA_Frame_E = BPA_Frame_EE[0]
    app.PrintInfo(BPA_Frame_E)

    BPA_Frame_FF = user_defined_model.SearchObject("BPA Frame(F).BlkDef")
    BPA_Frame_F = BPA_Frame_FF[0]
    app.PrintInfo(BPA_Frame_F)

    MVABASE0 = MVABASE

    for line in bpa_str_ar:
        line = line.rstrip('\n')
        if line == "": continue

        #MC/MF卡
        if line[0] == 'M' and (line[1] == 'F' or line[1] == 'C'):
            chinese_count = 0

            name = '_____'
            base = 0
            
            #判断中文的个数
            for i in range (3,10):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count = chinese_count + 1
            name = line[3:11-chinese_count].strip()
            base = float(line[11-chinese_count:15-chinese_count])
            #名称+电压
            Generator = name + str(base)
            Generator_num = 0
            #找到节点名对应的编号
            if Generator in bus_name_ar:
                Generator_num = bus_name_ar[Generator]
                MFCard_str.append(Generator_num)
           
#            else:
#                app.PrintInfo(Generator.encode('GBK'))

                Generator_name = 'sym_' + str(Generator_num) + '_1'

                sym_optionss = Input_Options.SearchObject("*.OptTypsym") #设置显示Td0s还是Tds，切换后才能输入Td0s
                sym_options = sym_optionss[0]
                sym_options.iopt_tag = 2
                sym_options.iopt_t = 1


                #选中发电机
                symm = Net.SearchObject(str(Generator_name))
                sym = symm[0]
                #判断MF卡是否有对应的发电机
                if not(sym == None) :
                    sym_type = sym.typ_id
                    #频率变化的影响
                    sym_type.i_speedVar = 0

                    MVABASE = MVABASE0
                    #MVABASE基准容量
                    if isfloat(line[28-chinese_count:32-chinese_count]):
                        MVABASE = float(line[28-chinese_count:32-chinese_count])
                        #无功上下限变随基准容量变化
                        QMIN = sym.cQ_min
                        QMAX = sym.cQ_max
                        sym_type.sgn = MVABASE
                        sym.cQ_min = QMIN
                        sym.cQ_max = QMAX
                    else:
                        MVABASE = sym_type.sgn
                           
                    #EMWS 发电机动能
                    EMWS_value = 0
                    if isfloat(line[16-chinese_count:22-chinese_count]):
                        EMWS_value = getfloatvalue(line[16-chinese_count:22-chinese_count],0)
                        SN_GEN = MVABASE                                  #额定容量？？？？
                        H = EMWS_value  / SN_GEN
                        
                        sym_type.tag = H*2/sym_type.cosn



                    #Ra，定子电阻（pu）
                    Ra = 0
                    if isfloat(line[32-chinese_count:36-chinese_count]):
                        Ra = getfloatvalue(line[32-chinese_count:36-chinese_count],4)
                        

                    #Xd'，直轴暂态电抗
                    xds_value = 0
                    if isfloat(line[36-chinese_count:41-chinese_count]):
                        xds_value = getfloatvalue(line[36-chinese_count:41-chinese_count],4)
                        

                    #Xq'，交流暂态电抗（pu）
                    xqs_value = 0
                    if isfloat(line[41-chinese_count:46-chinese_count]):
                        xqs_value = getfloatvalue(line[41-chinese_count:46-chinese_count],4)
                        

                    #Xd，直轴不饱和同步电抗（pu）
                    xd_value = 0
                    if isfloat(line[46-chinese_count:51-chinese_count]):
                        xd_value = getfloatvalue(line[46-chinese_count:51-chinese_count],4)
                        
                    
                    #Xq，交轴不饱和同步电抗（pu）
                    xq_value = 0
                    if isfloat(line[51-chinese_count:56-chinese_count]):
                        xq_value = getfloatvalue(line[51-chinese_count:56-chinese_count],4)
                        

                    #Tdo'，直轴暂态开路时间常数（秒）
                    tds_value = 0
                    if isfloat(line[56-chinese_count:60-chinese_count]):
                        tds_value = getfloatvalue(line[56-chinese_count:60-chinese_count],2)
                        

                    #Tqo'，交轴暂态开路时间常数（秒）
                    tqs_value = 0
                    if isfloat(line[60-chinese_count:63-chinese_count]):
                        tqs_value = getfloatvalue(line[60-chinese_count:63-chinese_count],2)                        

                    #XL，定子漏抗（pu）
                    xl_value = 0
                    if isfloat(line[63-chinese_count:68-chinese_count]):
                        xl_value = getfloatvalue(line[63-chinese_count:68-chinese_count],4)

                    #SG1.0，额定电压时的电机饱和系数（pu）
                    SG10 = 0
                    if isfloat(line[68-chinese_count:73-chinese_count]):
                        SG10 = getfloatvalue(line[68-chinese_count:73-chinese_count],4)

                    #SG1.2，1.2倍电压时的电机饱和系数（pu）
                    SG12 = 0
                    if isfloat(line[73-chinese_count:77-chinese_count]):
                        SG12 = getfloatvalue(line[73-chinese_count:77-chinese_count],3)

                    #电机阻尼转距系数
                    D = 0
                    if isfloat(line[77-chinese_count:80-chinese_count]):
                        D = getfloatvalue(line[77-chinese_count:80-chinese_count],2)
                    

                    sym_type.rstr = Ra
                    sym_type.xds = xds_value
                    sym_type.xqs = xqs_value
                    sym_type.xd = xd_value
                    sym_type.xq = xq_value
                    sym_type.tds0 = tds_value
                    sym_type.tqs0 = tqs_value
                    sym_type.xl = xl_value
#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                    if sym_type.xqs == sym_type.xq:
                        sym_type.tqs0 = 0

                    
#                    #饱和 二次
#                    if SG10 != 0 and SG12 != 0 and NOSAT == 0:
#                        sym_type.isat = 1
#                        sym_type.sg12 = SG12
#                        sym_type.sg10 = SG10

                    DMPALL,DMPMLT
                    if DMPALL != 0:
                        sym_type.dpu = DMPALL
                    else:
                        if D != 0:  
                            sym_type.dpu = DMPMLT*D
                            app.PrintInfo('DampDampDampDampDamp')
                            app.PrintInfo(sym)

                    #comp
                    Generator_comp_name = str(Generator_num) + '_comp'
                    Generator_compp = Net.SearchObject(str(Generator_comp_name))
                    Generator_comp = Generator_compp[0]
                    if Generator_comp == None:
                        Net.CreateObject('Elmcomp',Generator_comp_name)
                        Generator_compp = Net.SearchObject(str(Generator_comp_name))
                        Generator_comp = Generator_compp[0]

                
                    Generator_comp.typ_id = BPA_Frame_F

           #         Generator_comp.CreateObject('ElmDsl','sym_dsl')
           #         Generator_sym_dsll = Generator_comp.SearchObject(str(Generator_comp_name))
           #         Generator_sym_dsl = Generator_sym_dsll[0]            
    #                app.PrintInfo(sym)
                    Generator_comp.Sym= sym
           #         Generator_comp.pelm[2] = Generator_sym_dsl

    #            else:
    #                app.PrintInfo(Generator_name)


    #IAMRTS == 0 直接使用CASE卡参数
    if IAMRTS == 0:
        for line in bpa_str_ar:
            line = line.rstrip('\n')
            if line == "": continue        

            #M卡
            if line[0] == 'M' and line[1] == ' ':
                chinese_count = 0

                #判断中文的个数
                for i in range (3,10):
                    if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                        chinese_count = chinese_count + 1
                name = line[3:11-chinese_count].strip()
                base = float(line[11-chinese_count:15-chinese_count])
                #名称+电压
                Generator = name + str(base)                
                #找到节点名对应的编号
                if Generator in bus_name_ar:
                    Generator_num = bus_name_ar[Generator]
                    MCard_str.append(Generator_num)

                    Generator_name = 'sym_' + str(Generator_num) + '_1'

                    #选中发电机
                    symm = Net.SearchObject(str(Generator_name))
                    sym = symm[0]
                    #判断M卡是否有对应的发电机
                    if not(sym == None) :
                        sym_type = sym.typ_id

                        xdss_value = 0
                        xqss_value = 0
                        tdss_value = 0
                        tqss_value = 0

                        
                        #d轴次暂态电（pu）
                        if isfloat(line[37-chinese_count:42-chinese_count]):
                            xdss_value = getfloatvalue(line[37-chinese_count:42-chinese_count],4)
                        
                        #q轴次暂态电（pu）
                        if isfloat(line[42-chinese_count:47-chinese_count]):
                            xqss_value = getfloatvalue(line[42-chinese_count:47-chinese_count],4)
                        
                        #d轴次暂态时间常数（秒）
                        if isfloat(line[47-chinese_count:51-chinese_count]):
                            tdss_value = getfloatvalue(line[47-chinese_count:51-chinese_count],4)
                        
                        #q轴次暂态时间常数（秒）
                        if isfloat(line[51-chinese_count:55-chinese_count]):
                            tqss_value = getfloatvalue(line[51-chinese_count:55-chinese_count],4)
                        if sym_type.xl >= xdss_value or sym_type.xl >= xqss_value:
                            if xdss_value < xqss_value:
                                sym_type.xl = xdss_value-0.001
                            else:
                                sym_type.xl = xqss_value-0.001
                        sym_type.xdss = xdss_value
                        sym_type.xqss = xqss_value
                        sym_type.tdss0 = tdss_value
                        sym_type.tqss0 = tqss_value

#                        app.PrintInfo(sym)
#                        app.PrintInfo(xqss_value)
#                        app.PrintInfo(line[0:20].encode('GBK'))
#                        app.PrintInfo(sym_type.xqss)
                    

                    


        #没有M卡，使用CASE卡参数
        length = len(MFCard_str)
        for i in range(1,length+1):
            if not MFCard_str[i-1] in MCard_str:

                app.PrintInfo(MFCard_str[i-1])
                
                Generator_name = 'sym_' + str(MFCard_str[i-1]) + '_1'

                app.PrintInfo(Generator_name)
                #选中发电机
                symm = Net.SearchObject(str(Generator_name))
                sym = symm[0]

                app.PrintInfo(sym)

                #XFACT大于 0 CASE卡参数才有效，否则使用默认参数
                if  XFACT > 0:               
                    sym_type.xdss = sym_type.xds*XFACT
                    sym_type.xqss = sym_type.xds*XFACT
                    if sym_type.tqs > 0:
                        sym_type.tdss = TDODPS
                        sym_type.tqss = TQODPS
                    else:
                        sym_type.tdss = TDODPH
                        sym_type.tqss = TQODPH
                else:
                    sym_type.xdss = sym_type.xds*0.65
                    sym_type.xqss = sym_type.xds*0.65
                    if sym_type.tqs > 0:
                        sym_type.tdss = 0.03
                        sym_type.tqss = 0.05
                    else:
                        sym_type.tdss = 0.04
                        sym_type.tqss = 0.06
    else:
        length = len(MFCard_str)
        for i in range(1,length+1):
                Generator_name = 'sym_' + str(MFCard_str[i-1]) + '_1'
                #选中发电机
                symm = Net.SearchObject(str(Generator_name))
                sym = symm[0]

                #XFACT大于 0 CASE卡参数才有效，否则使用默认参数
                if  XFACT > 0:               
                    sym_type.xdss = sym_type.xds*XFACT
                    sym_type.xqss = sym_type.xds*XFACT
                    if sym_type.tqs > 0:
                        sym_type.tdss = TDODPS
                        sym_type.tqss = TQODPS
                    else:
                        sym_type.tdss = TDODPH
                        sym_type.tqss = TQODPH
                else:
                    sym_type.xdss = sym_type.xds*0.65
                    sym_type.xqss = sym_type.xds*0.65
                    if sym_type.tqs > 0:
                        sym_type.tdss = 0.03
                        sym_type.tqss = 0.05
                    else:
                        sym_type.tdss = 0.04
                        sym_type.tqss = 0.06
    #励磁机E卡
    GetECard(bpa_file, bpa_str_ar,bus_name_ar)

    #励磁机F卡
    GetFCard(bpa_file, bpa_str_ar,bus_name_ar)

    #PSS S卡
    GetSCard(bpa_file, bpa_str_ar,bus_name_ar)

    #调速器G卡
    GetGovCard(bpa_file, bpa_str_ar,bus_name_ar)

    #原动机T卡
    GetTCard(bpa_file, bpa_str_ar,bus_name_ar)

    #发电机转负荷
    GetLNCard(bpa_file, bpa_str_ar,bus_name_ar)

# ----------------------------------------------------------------------------------------------------
def GetXOCard(bpa_file,bpa_str_ar,bus_name_ar):

    XO_ar = {}
    XO_nbr = 0

        #当前工程
    prj = app.GetActiveProject()
    if prj is None:
        raise Exception("No project activated. Python Script stopped.")

    app.PrintInfo(prj)
    Network_Modell = prj.SearchObject("Network Model.IntPrjfolder")
    Network_Model=Network_Modell[0]

    Network_Dataa = Network_Model.SearchObject("Network Data.IntPrjfolder")
    Network_Data = Network_Dataa[0]
    app.PrintInfo(Network_Data)

    Nett = Network_Data.SearchObject("*.ElmNet")
    Net = Nett[0]
    app.PrintInfo(Net)

    for line in bpa_str_ar:
        line = line.rstrip('\n')
        if line == "": continue

        #XO卡
        if line[0] == 'X' and line[1] == 'O' :
            chinese_count1 = 0
            #判断中文的个数
            for i in range (4,12):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count1 = chinese_count1 + 1
            name1 = line[4:12-chinese_count1]
            base1 = float(line[12-chinese_count1:16-chinese_count1])
            #名称+电压
            bus1 = name1 + str(base1)
            bus1_num = 0
            #找到节点名对应的编号
            if bus1 in bus_name_ar:
                bus1_num = bus_name_ar[bus1]
            else:
                bus1_num = 0

            chinese_count2 = 0
            for i in range (18-chinese_count1,26-chinese_count1):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count2 = chinese_count2 + 1
            chinese_count = chinese_count2 + chinese_count1
            name2 = line[18-chinese_count1:26-chinese_count]
            base2 = float(line[26-chinese_count:30-chinese_count])

            #名称+电压
            bus2 = name2 + str(base2)
            bus2_num = 0
            num = 1
            #找到节点名对应的编号
            if bus2 in bus_name_ar:
                bus2_num = bus_name_ar[bus2]

            if isint(line[33-chinese_count]):
                num = int(line[33-chinese_count])

            trf_name = 'trf_' + str(bus1_num) + '_' + str(bus2_num) + '_' + str(num)

#            app.PrintInfo(trf_name)
            

            #选中变压器
            trff = Net.SearchObject(str(trf_name))
            trf = trff[0]

            if not(trf == None):
#                app.PrintInfo(trf)
                XO_nbr = XO_nbr+1
                XO_ar[trf.loc_name]=XO_nbr

                SID = int(line[31-chinese_count])
#                app.PrintInfo(SID)

                R0 = 0
                X0 = 0

                if isfloat(line[37-chinese_count:44-chinese_count]):
                    X0 = getfloatvalue(line[37-chinese_count:44-chinese_count],4)

                if isfloat(line[44-chinese_count:51-chinese_count]):
                    R0 = getfloatvalue(line[44-chinese_count:51-chinese_count],4)

                trf_type = trf.typ_id
                trf_type.x0pu = X0*100/100            #直接输入会被认作0 原因未知
                trf_type.r0pu = R0*100/100

                trf.outserv = 1  #has_key and in 不能用 用outserv作为判断依据

                if SID == 1:
                    trf_type.zx0hl_h = 1
                    trf_type.tr2cn_h = 'YN'
                    trf_type.tr2cn_l = 'D'
                elif SID == 2:
                    trf_type.zx0hl_h = 0
                    trf_type.tr2cn_h = 'YN'
                    trf_type.tr2cn_l = 'Y'                    
                elif SID == 3:
                    trf_type.zx0hl_h = 0.5
                    trf_type.tr2cn_h = 'YN'
                    trf_type.tr2cn_l = 'YN'

            trf_name = 'trf_' + str(bus2_num) + '_' + str(bus1_num) + '_' + str(num)   #顺序颠倒的情况
            #选中变压器
            trff = Net.SearchObject(str(trf_name))
            trf = trff[0]

            if not(trf == None):
#                app.PrintInfo(trf)
                XO_nbr = XO_nbr+1
                XO_ar[trf]=XO_nbr      

                SID = int(line[31-chinese_count])
#                app.PrintInfo(SID)

                R0 = 0
                X0 = 0

                if isfloat(line[37-chinese_count:44-chinese_count]):
                    X0 = getfloatvalue(line[37-chinese_count:44-chinese_count],4)

                if isfloat(line[44-chinese_count:51-chinese_count]):
                    R0 = getfloatvalue(line[44-chinese_count:51-chinese_count],4)

                trf_type = trf.typ_id
                trf_type.x0pu = X0*100/100            #直接输入会被认作0 原因未知
                trf_type.r0pu = R0*100/100

                trf.outserv = 1  #has_key and in 不能用 用outserv作为判断依据

                if SID == 1:
                    trf_type.zx0hl_h = 1
                    trf_type.tr2cn_h = 'YN'
                    trf_type.tr2cn_l = 'D'
                elif SID == 2:
                    trf_type.zx0hl_h = 0
                    trf_type.tr2cn_h = 'YN'
                    trf_type.tr2cn_l = 'Y'                    
                elif SID == 3:
                    trf_type.zx0hl_h = 0.5
                    trf_type.tr2cn_h = 'YN'
                    trf_type.tr2cn_l = 'YN'



            

    length = len(XO_ar)
#    app.PrintInfo(length)
#    for i in range (1,length:
#        app.PrintInfo(XO_ar[i-1])
                
    Trans = app.GetCalcRelevantObjects('*.ElmTr2')
    length = len(Trans)
#    app.PrintInfo(length)
    for i in range (1,length+1):      
        if  Trans[i-1].outserv == 0:
#            app.PrintInfo(Trans[i-1].typ_id.iopt_uk0)
#            app.PrintInfo(Trans[i-1])
            Trans[i-1].typ_id.x0pu = 9999
            Trans[i-1].typ_id.r0pu = 9999
            Trans[i-1].outserv = 0
        else:
            Trans[i-1].outserv = 0
            

            
#            app.PrintInfo(Trans[i-1])


                
# ----------------------------------------------------------------------------------------------------                
def GetLOCard(bpa_file,bpa_str_ar,bus_name_ar,MVABASE):
    #线路零序阻抗

    LO_ar = []
    
    prj = app.GetActiveProject()
    if prj is None:
        raise Exception("No project activated. Python Script stopped.")

    app.PrintInfo(prj)
    Network_Modell = prj.SearchObject("Network Model.IntPrjfolder")
    Network_Model=Network_Modell[0]

    Network_Dataa = Network_Model.SearchObject("Network Data.IntPrjfolder")
    Network_Data = Network_Dataa[0]
    app.PrintInfo(Network_Data)

    Nett = Network_Data.SearchObject("*.ElmNet")
    Net = Nett[0]
    app.PrintInfo(Net)

    for line in bpa_str_ar:
        line = line.rstrip('\n')
        if line == "": continue

        #LO卡
        if line[0] == 'L' and line[1] == 'O' :
            chinese_count1 = 0
            #判断中文的个数
            for i in range (4,12):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count1 = chinese_count1 + 1
            name1 = line[4:12-chinese_count1]
            base1 = float(line[12-chinese_count1:16-chinese_count1])
            #名称+电压
            bus1 = name1 + str(base1)
            bus1_num = 0
            #找到节点名对应的编号
            if bus1 in bus_name_ar:
                bus1_num = bus_name_ar[bus1]

            chinese_count2 = 0
            for i in range (18-chinese_count1,26-chinese_count1):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count2 = chinese_count2 + 1
            chinese_count = chinese_count2 + chinese_count1
            name2 = line[18-chinese_count1:26-chinese_count]
            base2 = float(line[26-chinese_count:30-chinese_count])

            #名称+电压
            bus2 = name2 + str(base2)
            bus2_num = 0
            num = 1
            #找到节点名对应的编号
            if bus2 in bus_name_ar:
                bus2_num = bus_name_ar[bus2]

            if isint(line[32-chinese_count]):
                num = int(line[32-chinese_count])
                
            lne_name = 'lne_' + str(bus1_num) + '_' + str(bus2_num) + '_' + str(num)

            #选中输电线
            lnee = Net.SearchObject(str(lne_name))
            lne = lnee[0]
            
            if not(lne == None) :

                LO_ar.append(lne_name)
                lne_type = lne.typ_id
                length =lne.dline
                base = lne_type.uline

                R0 = 0
                X0 = 0
                Ga0 = 0
                Gb0 = 0
                Ba0 = 0
                Bb0 = 0
                                

                if isfloat(line[35-chinese_count:42-chinese_count]):
                    R0 = getfloatvalue(line[35-chinese_count:42-chinese_count],4)
                 
                if isfloat(line[42-chinese_count:49-chinese_count]):
                    X0 = getfloatvalue(line[42-chinese_count:49-chinese_count],4)

                if isfloat(line[49-chinese_count:56-chinese_count]):
                    Ga0 = getfloatvalue(line[49-chinese_count:56-chinese_count],4)

                if isfloat(line[56-chinese_count:63-chinese_count]):
                    Ba0 = getfloatvalue(line[56-chinese_count:63-chinese_count],4)

                if isfloat(line[63-chinese_count:70-chinese_count]):
                    Gb0 = getfloatvalue(line[63-chinese_count:70-chinese_count],4)

                if isfloat(line[70-chinese_count:77-chinese_count]):
                    Bb0 = getfloatvalue(line[70-chinese_count:77-chinese_count],4)

                lne_type.rline0 = R0*base*base/MVABASE/length
                lne_type.xline0 = X0*base*base/MVABASE/length
                if Gb0 == 0 and Bb0 == 0:
                    lne_type.bline0 = Ba0*base*base/MVABASE/length
                    lne_type.gline0 = Ga0*base*base/MVABASE/length
    #G/B不平衡时，等效电抗器命名规则未知，等出现该情况时再做更改
                else:
                    app.PrintInfo(lne_name) 
                    raise Exception("Line is unbalanced. Python Script stopped.")

#bus1 bus2次序颠倒
            lne_name = 'lne_' + str(bus2_num) + '_' + str(bus1_num) + '_' + str(num)

            #选中输电线
            lnee = Net.SearchObject(str(lne_name))
            lne = lnee[0]
            
            if not(lne == None) :

                LO_ar.append(lne_name)
                lne_type = lne.typ_id
                length =lne.dline
                base = lne_type.uline

                R0 = 0
                X0 = 0
                Ga0 = 0
                Gb0 = 0
                Ba0 = 0
                Bb0 = 0
                                

                if isfloat(line[35-chinese_count:42-chinese_count]):
                    R0 = getfloatvalue(line[35-chinese_count:42-chinese_count],4)
                 
                if isfloat(line[42-chinese_count:49-chinese_count]):
                    X0 = getfloatvalue(line[42-chinese_count:49-chinese_count],4)

                if isfloat(line[49-chinese_count:56-chinese_count]):
                    Ga0 = getfloatvalue(line[49-chinese_count:56-chinese_count],4)

                if isfloat(line[56-chinese_count:63-chinese_count]):
                    Ba0 = getfloatvalue(line[56-chinese_count:63-chinese_count],4)

                if isfloat(line[63-chinese_count:70-chinese_count]):
                    Gb0 = getfloatvalue(line[63-chinese_count:70-chinese_count],4)

                if isfloat(line[70-chinese_count:77-chinese_count]):
                    Bb0 = getfloatvalue(line[70-chinese_count:77-chinese_count],4)

                lne_type.rline0 = R0*base*base/MVABASE/length
                lne_type.xline0 = X0*base*base/MVABASE/length
                if Gb0 == 0 and Bb0 == 0:
                    lne_type.bline0 = Ba0*base*base/MVABASE/length
                    lne_type.gline0 = Ga0*base*base/MVABASE/length
    #G/B不平衡时，等效电抗器命名规则未知，等出现该情况时再做更改
                else:
                    app.PrintInfo(lne_name) 
                    raise Exception("Line is unbalanced. Python Script stopped.")
                     

    #LO-Z卡的区域字典
    zone_ar = []
    LOZ_ar = []
#LO-Z线路缺省零序参数模型
    #?????????????????????????????????????????????????????????????????????????????????????????????? R/X未知
    
    for line in bpa_str_ar:
        line = line.rstrip('\n')
        if line == "": continue

        line = line + ' '*(80-len(line))

        #LO Z卡
        if line[0] == 'L' and line[1] == 'O' and line[2] == ' ' and line[3] == 'Z':
            zone = line[5:7]
            zone_ar.append(zone)

            zone = line[8:10]
            zone_ar.append(zone)    

            zone = line[11:13]
            zone_ar.append(zone)    
 
            zone = line[14:16]
            zone_ar.append(zone)    
 
            zone = line[17:19]
            zone_ar.append(zone)    
 
            zone = line[20:22]
            zone_ar.append(zone)    
 
            zone = line[23:25]
            zone_ar.append(zone)    
 
            zone = line[26:28]
            zone_ar.append(zone)    
 
            zone = line[29:31]
            zone_ar.append(zone)    
 
            zone = line[32:34]
            zone_ar.append(zone)    
 
            zone = line[35:37]
            zone_ar.append(zone)    
 
            zone = line[38:40]
            zone_ar.append(zone)    
 
            zone = line[41:43]
            zone_ar.append(zone)    
 
            zone = line[44:46]
            zone_ar.append(zone)    
 
            zone = line[47:49]
            zone_ar.append(zone)
            
            VMIN = -9999
            VMAX = 9999
            
            if isfloat(line[50:55]): KG0 = getfloatvalue(line[50:55],4)
            if isfloat(line[55:60]): KB0 = getfloatvalue(line[55:60],4)
            if isfloat(line[60:65]): KGC0 = getfloatvalue(line[60:65],4)
            if isfloat(line[65:70]): KBC0 = getfloatvalue(line[65:70],4)
            if isfloat(line[70:75]): VMIN = getfloatvalue(line[70:75],1)
            if isfloat(line[75:80]): VMAX = getfloatvalue(line[75:80],1)

            line = app.GetCalcRelevantObjects('*.ElmLne')
            line_count = len(line)
            for i in range (1,line_count):
                line_type = line[i-1].typ_id
                if line_type.uline >= VMIN and line_type.uline <= VMAX and (line[i-1].loc_name not in LO_ar) and line[i-1].cpZone.loc_name in zone_ar:
                    LOZ_ar.append(line[i-1].loc_name)
                    base = line[i-1].typ_id.uline
                    R = line[i-1].typ_id.rline
                    X = line[i-1].typ_id.xline
                    G = line[i-1].typ_id.gline
                    B = line[i-1].typ_id.bline
                    line[i-1].typ_id.rline0 = KG0*R/(R*R+X*X)
                    line[i-1].typ_id.xline0 = KB0*X/(R*R+X*X)

            #G/B不平衡时，等效电抗器命名规则未知，等出现该情况时再做更改
                    
                    line[i-1].typ_id.gline0 = KGC0 * line[i-1].typ_id.gline
                    line[i-1].typ_id.bline0 = KBC0 * line[i-1].typ_id.bline

#缺省情况
#    lolength = len(LO_ar)
#    app.PrintInfo(LO_ar)
    line = app.GetCalcRelevantObjects('*.ElmLne')
    line_count = len(line)
#    app.PrintInfo(line_count)
    for i in range (1,line_count+1):
#        app.PrintInfo(line[i-1].loc_name)
        if (line[i-1].loc_name not in LO_ar) and(line[i-1].loc_name not in LOZ_ar):
#            app.PrintInfo(line[i-1])
            line_type = line[i-1].typ_id
            base = line[i-1].typ_id.uline
            R = line[i-1].typ_id.rline
            X = line[i-1].typ_id.xline
            G = line[i-1].typ_id.gline
            B = line[i-1].typ_id.bline
            G0 = 0.8*R/(R*R+X*X)
            B0 = 0.313*X/(R*R+X*X)


            line[i-1].typ_id.rline0 = G0/(G0*G0+B0*B0)
            line[i-1].typ_id.xline0 = B0/(G0*G0+B0*B0)
        #G/B不平衡时，等效电抗器命名规则未知，等出现该情况时再做更改
                    
            line[i-1].typ_id.gline0 = 0.7 * line[i-1].typ_id.gline
            line[i-1].typ_id.bline0 = 0.68 * line[i-1].typ_id.bline

#XR卡
    for line in bpa_str_ar:
        line = line.rstrip('\n')
        if line == "": continue

        if line[0] == 'X' and line[1] == 'R':
            Load = 0
            chinese_count = 0
            #判断中文的个数
            for i in range (4,12):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count = chinese_count + 1
            name = line[4:12-chinese_count].strip()
            if isfloat(line[12-chinese_count:16-chinese_count]):
                base = float(line[12-chinese_count:16-chinese_count])
                #名称+电压
                xr_shnt = name + str(base)                
            #找到节点名对应的编号
                if xr_shnt in bus_name_ar:
                    shnt_num = bus_name_ar[xr_shnt]
#                    app.PrintInfo(shnt_num)
            
                    shnt_name = 'shntfix_' + str(shnt_num) + '_1'
                    shntt = Net.SearchObject(shnt_name)
                    shnt = shntt[0]

                    if shnt != None:
#                        app.PrintInfo(shnt)
                        base = shnt.ushnm

                        R0 = 0
                        X0 = 0
                    
                        if isfloat(line[21-chinese_count:28-chinese_count]): R0 = getfloatvalue(line[21-chinese_count:28-chinese_count],4)
                        if isfloat(line[28-chinese_count:35-chinese_count]): X0 = getfloatvalue(line[28-chinese_count:35-chinese_count],4)

                        if X0 != 0:
                            shnt.Bg = 1000000000/(X0*base*base/MVABASE)

#LO+卡 
    for line in bpa_str_ar:
        line = line.rstrip('\n')
        if line == "": continue

        #LO+卡
        if line[0] == 'L' and line[1] == 'O' and line[2] == '+':
            chinese_count1 = 0
            #判断中文的个数
            for i in range (4,12):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count1 = chinese_count1 + 1
            name1 = line[4:12-chinese_count1]
            base1 = float(line[12-chinese_count1:16-chinese_count1])
            #名称+电压
            bus1 = name1 + str(base1)
            bus1_num = 0
            #找到节点名对应的编号
            if bus1 in bus_name_ar:
                bus1_num = bus_name_ar[bus1]

            chinese_count2 = 0
            for i in range (18-chinese_count1,26-chinese_count1):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count2 = chinese_count2 + 1
            chinese_count = chinese_count2 + chinese_count1
            name2 = line[18-chinese_count1:26-chinese_count]
            base2 = float(line[26-chinese_count:30-chinese_count])

            #名称+电压
            bus2 = name2 + str(base2)
            bus2_num = 0
            num = 1
            #找到节点名对应的编号
            if bus2 in bus_name_ar:
                bus2_num = bus_name_ar[bus2]

            if isint(line[32-chinese_count]):
                num = int(line[32-chinese_count])

            if isfloat(line[35-chinese_count:42-chinese_count]): XL1 = float(line[35-chinese_count:42-chinese_count])
            if isfloat(line[42-chinese_count:49-chinese_count]): XL2 = float(line[42-chinese_count:49-chinese_count])

            
            shnt_name1 = 'shnt_' + str(bus1_num) + '_' + str(bus2_num) + '_' + str(num)

            #选中输电线
            shnt11 = Net.SearchObject(str(shnt_name1))
            shnt1 = shnt11[0]

            if shnt1 != None:

                app.PrintInfo(shnt1)

                base1 = shnt1.ushnm
                if XL1 != 0:
                    shnt1.B0 = 1000000000/(XL1*base1*base1/MVABASE)

            shnt_name2 = 'shnt_' + str(bus2_num) + '_' + str(bus1_num) + '_' + str(num)

            #选中输电线
            shnt22 = Net.SearchObject(str(shnt_name2))
            shnt2 = shnt22[0]

            if shnt2 != None:

                app.PrintInfo(shnt2)

                base2 = shnt2.ushnm
                if XL2 != 0:
                    shnt2.B0 = 1000000000/(XL2*base2*base2/MVABASE)                        
            
                                        
# ----------------------------------------------------------------------------------------------------                    
def GetLoadCard(bpa_file,bpa_str_ar,bus_name_ar,MVABASE):
#??????????????????????????????????????????????????????????????????????????????????? LB和L+的参数有问题
    #当前工程
    prj = app.GetActiveProject()
    if prj is None:
        raise Exception("No project activated. Python Script stopped.")

    app.PrintInfo(prj)
    Network_Modell = prj.SearchObject("Network Model.IntPrjfolder")
    Network_Model=Network_Modell[0]

    Network_Dataa = Network_Model.SearchObject("Network Data.IntPrjfolder")
    Network_Data = Network_Dataa[0]
    app.PrintInfo(Network_Data)

    Nett = Network_Data.SearchObject("*.ElmNet")
    Net = Nett[0]
    app.PrintInfo(Net)


    for line in bpa_str_ar:
        line = line.rstrip('\n')
        if line == "": continue

        if line[0] == 'L' and line[1] == 'A':
            Load = 0
            chinese_count = 0
            #判断中文的个数
            for i in range (3,10):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count = chinese_count + 1
            name = line[3:11-chinese_count].strip()
            if isfloat(line[11-chinese_count:15-chinese_count]):
                base = float(line[11-chinese_count:15-chinese_count])
                #名称+电压
                Load = name + str(base)                
            #找到节点名对应的编号
                if Load in bus_name_ar:
                    Load_num = bus_name_ar[Load]

                    lod_name = 'lod_' + str(Load_num) + '_1'
                    lodd = Net.SearchObject(lod_name)
                    lod = lodd[0]
                    TypLod_name = 'TypLod' + lod_name
                    lod_typlodd = Net.SearchObject(TypLod_name)
                    lod_typlod = lod_typlodd[0]
                    if lod_typlod == None:
                        Net.CreateObject('TypLod',TypLod_name)
                        lod_typlodd = Net.SearchObject(TypLod_name)
                        lod_typlod = lod_typlodd[0]

                    P1 = 0
                    Q1 = 0
                    P2 = 0
                    Q2 = 0
                    P3 = 0
                    Q3 = 0
                    P4 = 0
                    Q4 = 0
                    LDP = 0
                    LDQ = 0
                                                                    

                    if isfloat(line[27-chinese_count:32-chinese_count]): P1 = getfloatvalue(line[27-chinese_count:32-chinese_count],3)
                    if isfloat(line[32-chinese_count:37-chinese_count]): Q1 = getfloatvalue(line[32-chinese_count:37-chinese_count],3)
                    if isfloat(line[37-chinese_count:42-chinese_count]): P2 = getfloatvalue(line[37-chinese_count:42-chinese_count],3)
                    if isfloat(line[42-chinese_count:47-chinese_count]): Q2 = getfloatvalue(line[42-chinese_count:47-chinese_count],3)
                    if isfloat(line[47-chinese_count:52-chinese_count]): P3 = getfloatvalue(line[47-chinese_count:52-chinese_count],3)
                    if isfloat(line[52-chinese_count:57-chinese_count]): Q3 = getfloatvalue(line[52-chinese_count:57-chinese_count],3)
                    if isfloat(line[57-chinese_count:67-chinese_count]): Q4 = getfloatvalue(line[62-chinese_count:67-chinese_count],3)
                    if isfloat(line[67-chinese_count:72-chinese_count]): LDP = getfloatvalue(line[67-chinese_count:72-chinese_count],3)
                    if isfloat(line[72-chinese_count:77-chinese_count]): LDQ = getfloatvalue(line[72-chinese_count:77-chinese_count],3)

                    lod.typ_id = lod_typlod
                    lod_typlod.phtech = 2
                    lod_typlod.loddy = 100                    
                    lod_typlod.i_nln = 1 

                    if (P1+P2+P3!=1):
                        P1 = P1/(P1+P2+P3)
                        P2 = P2/(P1+P2+P3)
                        P3 = P3/(P1+P2+P3)

                    if (Q1+Q2+Q3!=1):
                        Q1 = Q1/(Q1+Q2+Q3)
                        Q2 = Q2/(Q1+Q2+Q3)
                        Q3 = Q3/(Q1+Q2+Q3)

                    lod_typlod.aP = P3
                    lod_typlod.bP = P2
                    lod_typlod.kpf = 0
                    lod_typlod.tpf = 0
                    lod_typlod.t1 = 0
                    lod_typlod.aQ = Q3
                    lod_typlod.bQ = Q2
                    lod_typlod.kqf = 0
                    lod_typlod.tqf = 0


            if Load == 0:
                if line[16] == ' ':
                    zone = str(line[15])
#                    app.PrintInfo(zone)
                else:
                    zone = str(line[15:17].rstrip('\n')) 
                lod = app.GetCalcRelevantObjects('*.ElmLod')
                length = len(lod)
                for i in range (1,length+1):
                    if lod[i-1].cpZone.loc_name == zone:
                        TypLod_name = 'TypLod' + lod[i-1].loc_name
                        lod_typlodd = Net.SearchObject(TypLod_name)
                        lod_typlod = lod_typlodd[0]
                        if lod_typlod == None:
                            Net.CreateObject('TypLod',TypLod_name)
                            lod_typlodd = Net.SearchObject(TypLod_name)
                            lod_typlod = lod_typlodd[0]

                        P1 = 0
                        Q1 = 0
                        P2 = 0
                        Q2 = 0
                        P3 = 0
                        Q3 = 0
                        P4 = 0
                        Q4 = 0
                        LDP = 0
                        LDQ = 0
                                                                    

                        if isfloat(line[27-chinese_count:32-chinese_count]): P1 = getfloatvalue(line[27-chinese_count:32-chinese_count],3)
                        if isfloat(line[32-chinese_count:37-chinese_count]): Q1 = getfloatvalue(line[32-chinese_count:37-chinese_count],3)
                        if isfloat(line[37-chinese_count:42-chinese_count]): P2 = getfloatvalue(line[37-chinese_count:42-chinese_count],3)
                        if isfloat(line[42-chinese_count:47-chinese_count]): Q2 = getfloatvalue(line[42-chinese_count:47-chinese_count],3)
                        if isfloat(line[47-chinese_count:52-chinese_count]): P3 = getfloatvalue(line[47-chinese_count:52-chinese_count],3)
                        if isfloat(line[52-chinese_count:57-chinese_count]): Q3 = getfloatvalue(line[52-chinese_count:57-chinese_count],3)
                        if isfloat(line[57-chinese_count:67-chinese_count]): Q4 = getfloatvalue(line[62-chinese_count:67-chinese_count],3)
                        if isfloat(line[67-chinese_count:72-chinese_count]): LDP = getfloatvalue(line[67-chinese_count:72-chinese_count],3)
                        if isfloat(line[72-chinese_count:77-chinese_count]): LDQ = getfloatvalue(line[72-chinese_count:77-chinese_count],3)

                        lod[i-1].typ_id = lod_typlod
                        lod_typlod.phtech = 2
                        lod_typlod.loddy = 100                   
#                        app.PrintInfo(P4)
                        lod_typlod.i_nln = 1

                        if (P1+P2+P3!=1):
                            P1 = P1/(P1+P2+P3)
                            P2 = P2/(P1+P2+P3)
                            P3 = P3/(P1+P2+P3)

                        if (Q1+Q2+Q3!=1):
                            Q1 = Q1/(Q1+Q2+Q3)
                            Q2 = Q2/(Q1+Q2+Q3)
                            Q3 = Q3/(Q1+Q2+Q3)
                            

                        lod_typlod.aP = P3
                        lod_typlod.bP = P2
                        lod_typlod.kpf = 0
                        lod_typlod.tpf = 0
                        lod_typlod.t1 = 0
                        lod_typlod.aQ = Q3
                        lod_typlod.bQ = Q2
                        lod_typlod.kqf = 0
                        lod_typlod.tqf = 0

        elif line[0] == 'L' and line[1] == 'B':
            Load = 0
            chinese_count = 0
            #判断中文的个数
            for i in range (3,10):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count = chinese_count + 1
            name = line[3:11-chinese_count].strip()
            if isfloat(line[11-chinese_count:15-chinese_count]):
                base = float(line[11-chinese_count:15-chinese_count])
                #名称+电压
                Load = name + str(base)                
            #找到节点名对应的编号
                if Load in bus_name_ar:
                    Load_num = bus_name_ar[Load]

                    lod_name = 'lod_' + str(Load_num) + '_1'
                    lodd = Net.SearchObject(lod_name)
                    lod = lodd[0]
                    TypLod_name = 'TypLod' + lod_name
                    lod_typlodd = Net.SearchObject(TypLod_name)
                    lod_typlod = lod_typlodd[0]
                    if lod_typlod == None:
                        Net.CreateObject('TypLod',TypLod_name)
                        lod_typlodd = Net.SearchObject(TypLod_name)
                        lod_typlod = lod_typlodd[0]

                    lod.typ_id = lod_typlod
                    lod_typlod.phtech = 2
                    lod_typlod.loddy = 100
                    lod_typlod.i_nln = 1

                    P1 = 0
                    Q1 = 0
                    P2 = 0
                    Q2 = 0
                    P3 = 0
                    Q3 = 0
                    LDP = 0
                    LDQ = 0
                                                                    

                    if isfloat(line[27-chinese_count:32-chinese_count]): P1 = getfloatvalue(line[27-chinese_count:32-chinese_count],3)
                    if isfloat(line[32-chinese_count:37-chinese_count]): Q1 = getfloatvalue(line[32-chinese_count:37-chinese_count],3)
                    if isfloat(line[37-chinese_count:42-chinese_count]): P2 = getfloatvalue(line[37-chinese_count:42-chinese_count],3)
                    if isfloat(line[42-chinese_count:47-chinese_count]): Q2 = getfloatvalue(line[42-chinese_count:47-chinese_count],3)
                    if isfloat(line[47-chinese_count:52-chinese_count]): P3 = getfloatvalue(line[47-chinese_count:52-chinese_count],3)
                    if isfloat(line[52-chinese_count:57-chinese_count]): Q3 = getfloatvalue(line[52-chinese_count:57-chinese_count],3)
                    if isfloat(line[57-chinese_count:67-chinese_count]): Q4 = getfloatvalue(line[62-chinese_count:67-chinese_count],3)
                    if isfloat(line[67-chinese_count:72-chinese_count]): LDP = getfloatvalue(line[67-chinese_count:72-chinese_count],3)
                    if isfloat(line[72-chinese_count:77-chinese_count]): LDQ = getfloatvalue(line[72-chinese_count:77-chinese_count],3)

                    lod_typlod.aP = P3
                    lod_typlod.bP = P2
                    lod_typlod.kpf = LDP
                    lod_typlod.tpf = 0
                    lod_typlod.t1 = 0
                    lod_typlod.aQ = Q3
                    lod_typlod.bQ = Q2
                    lod_typlod.kqf = LDQ
                    lod_typlod.tqf = 0


            if Load == 0:
                if line[16] == ' ':
                    zone = str(line[15])
                else:
                    zone = str(line[15:17].rstrip('\n'))            
                lod = app.GetCalcRelevantObjects('*.ElmLod')
                length = len(lod)
                for i in range (1,length+1):
                    if lod[i-1].cpZone.loc_name == zone:
                        TypLod_name = 'TypLod' + lod[i-1].loc_name
                        lod_typlodd = Net.SearchObject(TypLod_name)
                        lod_typlod = lod_typlodd[0]
                        if lod_typlod == None:
                            Net.CreateObject('TypLod',TypLod_name)
                            lod_typlodd = Net.SearchObject(TypLod_name)
                            lod_typlod = lod_typlodd[0]

                        P1 = 0
                        Q1 = 0
                        P2 = 0
                        Q2 = 0
                        P3 = 0
                        Q3 = 0
                        P4 = 0
                        Q4 = 0
                        LDP = 0
                        LDQ = 0
                                                                    

                        if isfloat(line[27-chinese_count:32-chinese_count]): P1 = getfloatvalue(line[27-chinese_count:32-chinese_count],3)
                        if isfloat(line[32-chinese_count:37-chinese_count]): Q1 = getfloatvalue(line[32-chinese_count:37-chinese_count],3)
                        if isfloat(line[37-chinese_count:42-chinese_count]): P2 = getfloatvalue(line[37-chinese_count:42-chinese_count],3)
                        if isfloat(line[42-chinese_count:47-chinese_count]): Q2 = getfloatvalue(line[42-chinese_count:47-chinese_count],3)
                        if isfloat(line[47-chinese_count:52-chinese_count]): P3 = getfloatvalue(line[47-chinese_count:52-chinese_count],3)
                        if isfloat(line[52-chinese_count:57-chinese_count]): Q3 = getfloatvalue(line[52-chinese_count:57-chinese_count],3)
                        if isfloat(line[57-chinese_count:67-chinese_count]): Q4 = getfloatvalue(line[62-chinese_count:67-chinese_count],3)
                        if isfloat(line[67-chinese_count:72-chinese_count]): LDP = getfloatvalue(line[67-chinese_count:72-chinese_count],3)
                        if isfloat(line[72-chinese_count:77-chinese_count]): LDQ = getfloatvalue(line[72-chinese_count:77-chinese_count],3)

                        lod[i-1].typ_id = lod_typlod
                        lod_typlod.phtech = 2
                        lod_typlod.loddy = 100
                        lod_typlod.i_nln = 1

                        lod_typlod.aP = P3
                        lod_typlod.bP = P2
                        lod_typlod.kpf = LDP
                        lod_typlod.tpf = 0
                        lod_typlod.t1 = 0
                        lod_typlod.aQ = Q3
                        lod_typlod.bQ = Q2
                        lod_typlod.kqf = LDQ
                        lod_typlod.tqf = 0
                        

                 
# ----------------------------------------------------------------------------------------------------
def revise_zone():
    #修改区域名
    
    prj = app.GetActiveProject()
    if prj is None:
        raise Exception("No project activated. Python Script stopped.")

    app.PrintInfo(prj)
    Network_Modell = prj.SearchObject("Network Model.IntPrjfolder")
    Network_Model=Network_Modell[0]

    Network_Dataa = Network_Model.SearchObject("Network Data.IntPrjfolder")
    Network_Data = Network_Dataa[0]
    app.PrintInfo(Network_Data)

    Zoness = Network_Data.SearchObject("Zones.IntZone")
    Zones = Zoness[0]
    app.PrintInfo(Zones)


    zone = app.GetCalcRelevantObjects('*.ElmZone')
    length = len(zone)
    for i in range (0,length):        
        name = zone[i].loc_name
        name_length = len(name)
        for j in range (0,name_length-1):
            if name[j] == '_':
                name_new = name[j+1:name_length]
                zone[i].loc_name = name_new
                break;
            
# ----------------------------------------------------------------------------------------------------        
def revise_bus(bpa_file,bpa_str_ar,bus_name_ar,bus_name_str):

    prj = app.GetActiveProject()
    if prj is None:
        raise Exception("No project activated. Python Script stopped.")

    app.PrintInfo(prj)
    Network_Modell = prj.SearchObject("Network Model.IntPrjfolder")
    Network_Model=Network_Modell[0]

    Network_Dataa = Network_Model.SearchObject("Network Data.IntPrjfolder")
    Network_Data = Network_Dataa[0]
    app.PrintInfo(Network_Data)

    Nett = Network_Data.SearchObject("*.ElmNet")
    Net = Nett[0]
    app.PrintInfo(Net)
    
    length = len(bus_name_str)
    for i in range (1,length+1):
        bus_name = str(i) + ' ' + str(i)
#        app.PrintInfo(comp_name)
        buss = Net.SearchObject(bus_name)
        bus = buss[0]
#        app.PrintInfo(name)
        if bus != None:
            name = bus_name_str[i-1].encode('GBK')
#            app.PrintInfo(name)
            bus.loc_name = name
                
# ----------------------------------------------------------------------------------------------------        
def revise_comp(bpa_file,bpa_str_ar,bus_name_ar,bus_name_str):

    prj = app.GetActiveProject()
    if prj is None:
        raise Exception("No project activated. Python Script stopped.")

    app.PrintInfo(prj)
    Network_Modell = prj.SearchObject("Network Model.IntPrjfolder")
    Network_Model=Network_Modell[0]

    Network_Dataa = Network_Model.SearchObject("Network Data.IntPrjfolder")
    Network_Data = Network_Dataa[0]
    app.PrintInfo(Network_Data)

    Nett = Network_Data.SearchObject("*.ElmNet")
    Net = Nett[0]
    app.PrintInfo(Net)
    
    length = len(bus_name_str)
    for i in range (1,length+1):
        comp_name = str(i) + '_' + 'comp'
#        app.PrintInfo(comp_name)
        compp = Net.SearchObject(comp_name)
        comp = compp[0]
#        app.PrintInfo(name)
        if comp != None:
#            name = bus_name_str[i-1].encode('GBK')
#            app.PrintInfo(name)
            name = 'comp' + ' ' + bus_name_str[i-1]
            name = name.encode('GBK')
            comp.loc_name = name
        
# ----------------------------------------------------------------------------------------------------        
def revise_sym(bpa_file,bpa_str_ar,bus_name_ar,bus_name_str):

    prj = app.GetActiveProject()
    if prj is None:
        raise Exception("No project activated. Python Script stopped.")

    app.PrintInfo(prj)
    Network_Modell = prj.SearchObject("Network Model.IntPrjfolder")
    Network_Model=Network_Modell[0]

    Network_Dataa = Network_Model.SearchObject("Network Data.IntPrjfolder")
    Network_Data = Network_Dataa[0]
    app.PrintInfo(Network_Data)

    Nett = Network_Data.SearchObject("*.ElmNet")
    Net = Nett[0]
    app.PrintInfo(Net)
    
    length = len(bus_name_str)
    for i in range (1,length+1):
        sym_name = 'sym' + '_' + str(i) + '_' + '1'
#        app.PrintInfo(comp_name)
        symm = Net.SearchObject(sym_name)
        sym = symm[0]
#        app.PrintInfo(name)
        if sym != None:
#            name = bus_name_str[i-1].encode('GBK')
#            app.PrintInfo(name)
            name = 'sym' + ' ' + bus_name_str[i-1]
            name = name.encode('GBK')
            sym.loc_name = name
        
# ----------------------------------------------------------------------------------------------------
DEBUG = 1
if DEBUG:
    
    bpa_file = OpenFile() # 打开指定文件
    bus_file = OpenBusFile()

    #读入节点名及电压幅值
    bus_str = bus_file.read()
    bus_file.seek(0)
    bus_str_ar = bus_file.readlines()

    bus_name_ar = {}
    bus_nbr = 0
    bus_name_str = []
    MVABASE = 100
    for line in bus_str_ar:
        #把读入的行补足
        line = line + ' '*(17-len(line))
        if line[0] == '1':
            busname = line[1:9].strip()
    #        app.PrintInfo(busname.encode('gb2312'))
            busbase = float(line[9:17])
    #        app.PrintInfo(busbase)
            bus_nbr = bus_nbr + 1
            bus_name_ar[busname+str(busbase)]=bus_nbr
            bus_name_str.append(busname+str(busbase))
        if line[0] == '2':
            MVABASE = float(line[1:8])
            
if bpa_file:					#If the file opened successfully
    #psspy.progress_output(6,"",[0,0])          # 不显示
    #psspy.progress_output(1,"",[0,0])           # 显示

    bpa_str = bpa_file.read()			#The string containing the text file, to use the find() function
    bpa_file.seek(0)		#To position back at the beginning
    bpa_str_ar = bpa_file.readlines()		#The array that is containing all the lines of the BPA file

    #读计算控制卡（CASE卡）    
    IXO,IWSCC,X2FAC,XFACT,TDODPS,TQODPS,TDODPH,TQODPH,CFACL2 = GetCASECard(bpa_file, bpa_str)

    #读F1卡
    DMPALL,IAMRTS = GetF1Card(bpa_str_ar)

    #读FF卡
    DMPMLT,NOSAT = GetFFCard(bpa_str_ar)

    #修改区域名
    revise_zone()
    
    #读发电机及控制系统的数据
    GetGenerator(bpa_file, bpa_str_ar,bus_name_ar,XFACT,TDODPS,TQODPS,TDODPH,TQODPH,IAMRTS,DMPALL,DMPMLT,NOSAT,MVABASE)
    
    #读变压器数据
    GetXOCard(bpa_file,bpa_str_ar,bus_name_ar)

    #读线路数据
    GetLOCard(bpa_file,bpa_str_ar,bus_name_ar,MVABASE)

    #读负荷数据
    GetLoadCard(bpa_file,bpa_str_ar,bus_name_ar,MVABASE)
        
    #busname改为中文
    revise_bus(bpa_file,bpa_str_ar,bus_name_ar,bus_name_str)
   
    #compname改为中文    
    revise_comp(bpa_file,bpa_str_ar,bus_name_ar,bus_name_str)

    #symname改为中文    
    revise_sym(bpa_file,bpa_str_ar,bus_name_ar,bus_name_str)

    app.PrintInfo(bus_name_str[0].encode('gbk'))
    app.PrintInfo(bus_name_str[1].encode('gbk'))
    a = bus_name_str[1].encode('gbk')
    app.PrintInfo(a)



#浙北仑 2554 VRMAX VRMIN
#沪石二_224.0 应为F+卡，实为FZ卡
#闽后石
