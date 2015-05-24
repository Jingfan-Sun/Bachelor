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
    path = 'C:\\Users\\Outrace\\Desktop\\bpa4.1_crack\\SAMPLE\\IEEE9\\IEEE90.dat' 
    bpa_file = open(path, 'r')
    app.PrintPlain("Open Successfully!")

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
def GetBCard(bpa_file, bpa_str_ar):
    #B卡
    
    bus_name = []
    bus_index = 0
    
    generator_name = []
    generator_index = 0
    
    shunt_index = 0
    
    load_index = 0
    

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
    app.PrintInfo('Net')
    app.PrintInfo(Net)

    Libraryy = prj.SearchObject("Library.IntPrjfolder")
    Library = Libraryy[0]
    app.PrintInfo(Library)

    user_defined_modell = Library.SearchObject("User Defined Models.IntPrjfolder")
    user_defined_model = user_defined_modell[0]
    app.PrintInfo(user_defined_model)

    BPA_Frame_BB = user_defined_model.SearchObject("BPA Frame(E).BlkDef")
    BPA_Frame_B = BPA_Frame_BB[0]
    app.PrintInfo(BPA_Frame_B)
    
    Zoness = Network_Data.SearchObject("Zones.IntZone")
    Zones = Zoness[0]
    app.PrintInfo(Zones)

    for line in bpa_str_ar:
        line = line.rstrip('\n')
        if line == "": continue
        
        if line[0] == 'B':
            chinese_count = 0
            #判断中文的个数
            for i in range (6,14):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count = chinese_count + 1
            name = line[6:14-chinese_count].strip()
            base = float(line[14-chinese_count:19-chinese_count])
            #名称+电压
            Variable_name = name + '_' + str(base)
            
            if line[1] == ' ' or line[1] == 'T' or line[1] == 'C' or line[1] == 'V' or line[1] == 'F' or line[1] == 'J' or line [1] == 'X':
                #The bus type code for a PQ bus
                
                bus_index = bus_index + 1
                bus = Net.CreateObject('ElmTerm', 'bus' + str(bus_index))
                bus = bus[0]
                bus_name.append(Variable_name)
                
                cubic = bus.CreateObject('StaCubic', 'Cubic')
                app.PrintPlain(cubic[0].loc_name.encode('GBK'))
                cubic = cubic[0]

                #Zone
                zone = Zones.SearchObject(str(line[18-chinese_count:21-chinese_count].strip()))
                # app.PrintInfo(zone)
                zone = zone[0]
                
                if zone == None: 
                    zone = Zones.CreateObject('ElmZone', 'Zone')
                    zone = zone[0]
                    zone.loc_name = line[18-chinese_count:21-chinese_count]
                    
                bus.cpZone = zone
                
                load = line[20-chinese_count:25-chinese_count]
                # app.PrintInfo(load+'1')
                if not(load[0:3] == '   ' or load == ''):
                    load_index = load_index + 1
                    load = Net.CreateObject('ElmLod', 'Load' + str(load_index))
                    load = load[0]
#                     app.PrintInfo(line[25-chinese_count:28-chinese_count].strip().rstrip('.'))
                    load.plini = float(line[20-chinese_count:25-chinese_count].strip().rstrip('.'))
                    load.qlini = float(line[25-chinese_count:30-chinese_count].strip().rstrip('.'))
                    
                    app.PrintInfo(load)
                    load.bus1 = cubic
                    
                shunt = line[34-chinese_count:38-chinese_count]
                # app.PrintInfo(load+'1')
                if not(shunt[0:3] == '   ' or load == ''):
                    shunt_index = shunt_index + 1
                    shunt = Net.CreateObject('ElmShnt', 'shunt' + str(shunt_index))
                    shunt = shunt[0]
                    shunt.qtotn = float(line[34-chinese_count:38-chinese_count].strip().rstrip('.'))
                    
                    app.PrintInfo(shunt)
                    shunt.bus1 = cubic
                
            elif line[1] == 'E' or line[1] == 'Q' or line[1] == 'G' or line[1] == 'K' or line[1] == 'L':
                #The bus type code for a PV bus
                
                # Generator
                generator_index = generator_index + 1
                generator = Net.CreateObject('ElmSym', 'generator' + str(generator_index))
                generator = generator[0]
                generator_name.append(Variable_name)
                
                generator.ip_ctrl = 0;  #PV
                generator.iv_mode = 1;
                generator.pgini = float(line[42-chinese_count:47-chinese_count])
                generator.q_max = float(line[47-chinese_count:52-chinese_count])
                generator.usetp = float(line[57-chinese_count:61-chinese_count])
                
            elif line[1] == 'S':
                #The bus type code for a swing bus
                # continue;
                
                # Generator
                generator_index = generator_index + 1
                generator = Net.CreateObject('ElmSym', 'generator' + str(generator_index))
                generator = generator[0]
                generator_name.append(Variable_name)
                
                generator.ip_ctrl = 1; #reference
                generator.iv_mode = 1;
                generator.pgini = float(line[42-chinese_count:47-chinese_count])
                generator.q_max = float(line[47-chinese_count:52-chinese_count])
                generator.usetp = float(line[57-chinese_count:61-chinese_count])
                
    # app.PrintInfo(generator_name[0].encode('GBK'))

# ----------------------------------------------------------------------------------------------------
DEBUG = 1
if DEBUG:
    
    bpa_file = OpenFile() # 打开指定文件

            
if bpa_file:					#If the file opened successfully

    bpa_str = bpa_file.read()			#The string containing the text file, to use the find() function
    bpa_file.seek(0)		#To position back at the beginning
    bpa_str_ar = bpa_file.readlines()		#The array that is containing all the lines of the BPA file
    
    # app.PrintPlain(bpa_str_ar[0])
    
    for line in bpa_str_ar: 
        if line[0:2] == "/M": 
            MVABASE = float(line[line.find("=")+1:line.find("\\")].lstrip())            #To continue if it is a blank line
            break
            
    GetBCard(bpa_file, bpa_str_ar)


