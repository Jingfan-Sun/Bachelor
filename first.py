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
def OpenFile(system):
    """Opens the BPA file."""
    path_win = 'E:\\Bachelor\\2013年夏低_华东.dat' 
    path_mac = 'Z:\\Dev\\Bachelor\\2013年夏低_华东.dat'
    if system == 'win': path = path_win
    else: path = path_mac
    bpa_file = open(path, 'r')
    app.PrintPlain("Open Successfully!")

    return bpa_file

# ----------------------------------------------------------------------------------------------------
def myFloat(str, defalt):
    """Enpty -> defalt; Positive Stay; Negetive -> defalt"""

    if str.strip() == '': return defalt
    elif float(str) < 0: return defalt
    else: return float(str)

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
def changeName():
    """Change the name of Models"""

    global bus_name_cn
    global bus_name

    global generator_name_cn
    global generator_name

    global transLine_name_cn
    global transLine_name

    global shunt_name
    global shunt_name_cn

    global scap_name_cn
    global scap_name

    global transformers_name_cn
    global transformers_name

    prj = app.GetActiveProject()
    if prj is None:
        raise Exception("No project activated. Python Script stopped.")

    # app.PrintInfo(prj)
    Network_Modell = prj.SearchObject("Network Model.IntPrjfolder")
    Network_Model=Network_Modell[0]

    Network_Dataa = Network_Model.SearchObject("Network Data.IntPrjfolder")
    Network_Data = Network_Dataa[0]
    # app.PrintInfo(Network_Data)

    Nett = Network_Data.SearchObject("*.ElmNet")
    Net = Nett[0]
    # app.PrintInfo('Net')
    # app.PrintInfo(Net)

    Libraryy = prj.SearchObject("Library.IntPrjfolder")
    Library = Libraryy[0]
    # app.PrintInfo(Library)

    user_defined_modell = Library.SearchObject("User Defined Models.IntPrjfolder")
    user_defined_model = user_defined_modell[0]
    # app.PrintInfo(user_defined_model)

    BPA_Frame_BB = user_defined_model.SearchObject("BPA Frame(E).BlkDef")
    BPA_Frame_B = BPA_Frame_BB[0]
    # app.PrintInfo(BPA_Frame_B)
    
    Zoness = Network_Data.SearchObject("Zones.IntZone")
    Zones = Zoness[0]
    # app.PrintInfo(Zones)

    for model in bus_name:
        bus = Net.SearchObject(model)
        bus = bus[0]
        bus.loc_name = bus_name_cn[bus_name.index(model)].encode("GBK")

    for model in generator_name:
        generator = Net.SearchObject(model)
        generator = generator[0]
        generator.loc_name = generator_name_cn[generator_name.index(model)].encode("GBK")

    for model in transLine_name:
        transLine = Net.SearchObject(model)
        transLine = transLine[0]
        transLine.loc_name = transLine_name_cn[transLine_name.index(model)].encode("GBK")

    for model in shunt_name:
        shunt = Net.SearchObject(model)
        shunt = shunt[0]
        shunt.loc_name = shunt_name_cn[shunt_name.index(model)].encode("GBK")

    for model in scap_name:
        scap = Net.SearchObject(model)
        scap = scap[0]
        scap.loc_name = scap_name_cn[scap_name.index(model)].encode("GBK")

    for model in transformers_name:
        transformers = Net.SearchObject(model)
        transformers = transformers[0]
        transformers.loc_name = transformers_name_cn[transformers_name.index(model)].encode("GBK")

# ----------------------------------------------------------------------------------------------------
def GetBCard(bpa_file, bpa_str_ar):
    #B卡 
    global MVABASE

    global bus_name_cn
    global bus_name
    global bus_index
    
    global generator_name_cn
    global generator_name
    global generator_index
    
    global shunt_index
    global shunt_name
    global shunt_name_cn
    
    global load_index

    prj = app.GetActiveProject()
    if prj is None:
        raise Exception("No project activated. Python Script stopped.")

    # app.PrintInfo(prj)
    Network_Modell = prj.SearchObject("Network Model.IntPrjfolder")
    Network_Model=Network_Modell[0]

    Network_Dataa = Network_Model.SearchObject("Network Data.IntPrjfolder")
    Network_Data = Network_Dataa[0]
    # app.PrintInfo(Network_Data)

    Nett = Network_Data.SearchObject("*.ElmNet")
    Net = Nett[0]
    # app.PrintInfo('Net')
    # app.PrintInfo(Net)

    Libraryy = prj.SearchObject("Library.IntPrjfolder")
    Library = Libraryy[0]
    # app.PrintInfo(Library)

    user_defined_modell = Library.SearchObject("User Defined Models.IntPrjfolder")
    user_defined_model = user_defined_modell[0]
    # app.PrintInfo(user_defined_model)

    BPA_Frame_BB = user_defined_model.SearchObject("BPA Frame(E).BlkDef")
    BPA_Frame_B = BPA_Frame_BB[0]
    # app.PrintInfo(BPA_Frame_B)
    
    Zoness = Network_Data.SearchObject("Zones.IntZone")
    Zones = Zoness[0]
    # app.PrintInfo(Zones)

    for line in bpa_str_ar:
        line = line.rstrip('\n')
        line = line[0:80]
        line = line + ' '*(80-len(line))
        if line == "": continue
        
        if line[0] == 'B':
            chinese_count = 0
            #判断中文的个数
            for i in range (6,14):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count = chinese_count + 1
            name = line[6:14-chinese_count].strip()
            base = float(line[14-chinese_count:18-chinese_count])
            #名称+电压
            Variable_name = name + '   ' + str(base)

            #Zone
            zone = Zones.SearchObject(str(line[18-chinese_count:20-chinese_count].strip()))
            zone = zone[0]
                
            if zone == None: 
                zone = Zones.CreateObject('ElmZone', 'Zone')
                zone = zone[0]
                zone.loc_name = line[18-chinese_count:20-chinese_count]
            
            if line[1] == ' ' or line[1] == 'T' or line[1] == 'C' or line[1] == 'V' or line[1] == 'F' or line[1] == 'J' or line [1] == 'X':
                #The bus type code for a PQ bus
                
                bus_index = bus_index + 1
                bus = Net.SearchObject('bus' + str(bus_index))
                bus = bus[0]
                if bus == None:
                	bus = Net.CreateObject('ElmTerm', 'bus' + str(bus_index))
                	bus = bus[0]
                bus_name_cn.append('Bus_' + Variable_name)
                bus_name.append('bus' + str(bus_index))

                bus.uknom = base
                bus.cpZone = zone
                
                load = line[20-chinese_count:25-chinese_count]
                # app.PrintInfo(load+'1')
                if isfloat(load):
                    load_index = load_index + 1
                    load = Net.SearchObject('load' + str(load_index))
                    load = load[0]
                    if load == None:
                    	load = Net.CreateObject('ElmLod', 'load' + str(load_index))
                    	load = load[0]
                    load.plini = float(line[20-chinese_count:25-chinese_count].strip().rstrip('.'))
                    load.qlini = float(line[25-chinese_count:30-chinese_count].strip().rstrip('.'))
                    cubic = bus.SearchObject('Cubic_load' + str(load_index))
                    cubic = cubic[0]
                    if cubic == None:
                        cubic = bus.CreateObject('StaCubic', 'Cubic_load' + str(load_index))
                        cubic = cubic[0]
                    #新建typeLoad
                    TypLod_name = 'TypeLoad' + str(load_index)
                    TypLod = Library.SearchObject(TypLod_name)
                    TypLod = TypLod[0]
                    if TypLod == None:
                        TypLod = Library.CreateObject('TypLod',TypLod_name)
                        TypLod = TypLod[0]
                    TypLod.aP = 1
                    TypLod.aQ = 1
                    load.bus1 = cubic
                    load.typ_id = TypLod
					
                shunt = line[34-chinese_count:38-chinese_count]
                # app.PrintInfo(load+'1')
                if isfloat(shunt):
                    shunt_index = shunt_index + 1
                    shunt_name_cn.append('Shunt_' + Variable_name)
                    shunt_name.append('shunt' + str(shunt_index))
                    shunt = Net.SearchObject('shunt' + str(shunt_index))
                    shunt = shunt[0]
                    if shunt == None:
                    	shunt = Net.CreateObject('ElmShnt', 'shunt' + str(shunt_index))
                    	shunt = shunt[0]
                    shunt.qtotn = float(line[34-chinese_count:38-chinese_count].strip().rstrip('.'))
                    cubic = bus.SearchObject('Cubic_shunt' + str(shunt_index))
                    cubic = cubic[0]
                    if cubic == None:
                        cubic = bus.CreateObject('StaCubic', 'Cubic_shunt' + str(shunt_index))
                        cubic = cubic[0]
                    shunt.bus1 = cubic
                    shunt.ushnm = base
                    value = float(line[34-chinese_count:38-chinese_count].strip().rstrip('.'))
                    if value > 0:
                        shunt.shtype = 2
                        shunt.qcapn = value
                    else:
                        shunt.shtype = 1
                        shunt.qrean = value

                
            elif line[1] == 'E' or line[1] == 'Q' or line[1] == 'G' or line[1] == 'K' or line[1] == 'L':
                #The bus type code for a PV bus
                
                # Generator
                generator_index = generator_index + 1
                generator = Net.SearchObject('generator' + str(generator_index))
                generator = generator[0]
                if generator == None:
                    generator = Net.CreateObject('ElmSym', 'generator' + str(generator_index))
                    generator = generator[0]
                    g_bus = Net.CreateObject('ElmTerm', 'bus_generator' + str(generator_index))
                    g_bus = g_bus[0]
                    cubic = g_bus.CreateObject('StaCubic', 'Cubic_generator' + str(generator_index))
                    cubic = cubic[0]
                    Typgen = Library.CreateObject('TypSym', 'TypeGenerator' + str(generator_index))
                    Typgen = Typgen[0]
                g_bus = Net.SearchObject('bus_generator' + str(generator_index))
                g_bus = g_bus[0]
                cubic = g_bus.SearchObject('Cubic_generator' + str(generator_index))
                cubic = cubic[0]
                Typgen = Library.SearchObject('TypeGenerator' + str(generator_index))
                Typgen = Typgen[0]
                generator_name_cn.append('Generator_' + Variable_name)
                generator_name.append('generator' + str(generator_index))
                bus_name_cn.append('Bus_' + Variable_name)
                bus_name.append('bus_generator' + str(generator_index))
                Typgen.sgn = MVABASE
                Typgen.ugn = base
                Typgen.cosn = 1

                g_bus.uknom = base
                g_bus.cpZone = zone
                
                generator.bus1 = cubic
                generator.typ_id = Typgen
                generator.ip_ctrl = 0;  #PV
                generator.iv_mode = 1;
                generator.Pmax_uc = float(line[38-chinese_count:42-chinese_count])
                generator.pgini = float(line[42-chinese_count:47-chinese_count])
                generator.q_max = float(line[47-chinese_count:52-chinese_count]) / MVABASE  #无功出力最大值
                generator.q_min = float(line[53-chinese_count:57-chinese_count]) / MVABASE  #无功出力最小值
                generator.usetp = float(line[57-chinese_count:61-chinese_count])

                #load
                if isfloat(line[20-chinese_count:25-chinese_count]):
                    load = Net.SearchObject('load_generator' + str(generator_index))
                    load = load[0]
                    if load == None:
                        load = Net.CreateObject('ElmLod', 'load_generator' + str(generator_index))
                        load = load[0]
                    load.plini = float(line[20-chinese_count:25-chinese_count].strip().rstrip('.'))
                    load.qlini = float(line[25-chinese_count:30-chinese_count].strip().rstrip('.'))
                    cubic = g_bus.SearchObject('Cubic_load_generator' + str(load_index))
                    cubic = cubic[0]
                    if cubic == None:
                        cubic = g_bus.CreateObject('StaCubic', 'Cubic_load_generator' + str(load_index))
                        cubic = cubic[0]
                    #新建typeLoad
                    TypLod_name = 'TypeLoadGenerator' + str(load_index)
                    TypLod = Library.SearchObject(TypLod_name)
                    TypLod = TypLod[0]
                    if TypLod == None:
                        TypLod = Library.CreateObject('TypLod',TypLod_name)
                        TypLod = TypLod[0]
                    TypLod.aP = 1
                    TypLod.aQ = 1
                    load.bus1 = cubic
                    load.typ_id = TypLod
                
            elif line[1] == 'S':    #SL
                generator_index = generator_index + 1
                generator = Net.SearchObject('generator' + str(generator_index))
                generator = generator[0]
                if generator == None:
                    generator = Net.CreateObject('ElmSym', 'generator' + str(generator_index))
                    generator = generator[0]
                    g_bus = Net.CreateObject('ElmTerm', 'bus_generator' + str(generator_index))
                    g_bus = g_bus[0]
                    cubic = g_bus.CreateObject('StaCubic', 'Cubic_generator' + str(generator_index))
                    cubic = cubic[0]
                    Typgen = Library.CreateObject('TypSym', 'TypeGenerator' + str(generator_index))
                    Typgen = Typgen[0]
                g_bus = Net.SearchObject('bus_generator' + str(generator_index))
                g_bus = g_bus[0]
                cubic = g_bus.SearchObject('Cubic_generator' + str(generator_index))
                cubic = cubic[0]
                Typgen = Library.SearchObject('TypeGenerator' + str(generator_index))
                Typgen = Typgen[0]
                Typgen.sgn = MVABASE
                Typgen.ugn = base
                Typgen.cosn = 1

                generator_name_cn.append('Generator_' + Variable_name)
                generator_name.append('generator' + str(generator_index))
                bus_name_cn.append('Bus_' + Variable_name)
                bus_name.append('bus_generator' + str(generator_index))

                g_bus.uknom = base
                g_bus.cpZone = zone

                generator.bus1 = cubic
                generator.typ_id = Typgen
                generator.ip_ctrl = 1; #reference
                generator.iv_mode = 1;
                generator.pgini = 0
                generator.q_max = float(line[47-chinese_count:52-chinese_count]) / MVABASE
                generator.usetp = float(line[57-chinese_count:61-chinese_count])
                
# ----------------------------------------------------------------------------------------------------
def GetLCard(bpa_file, bpa_str_ar):
    #L卡
    global MVABASE
    global MileToKm
    
    global bus_name_cn
    global bus_name
    global bus_index
    
    global generator_name_cn
    global generator_name
    global generator_index
    
    global shunt_index
    global shunt_name
    global shunt_name_cn
    
    global load_index

    global scap_name_cn
    global scap_name
    
    global transLine_index
    global transLine_name_cn
    global transLine_name

    prj = app.GetActiveProject()
    if prj is None:
        raise Exception("No project activated. Python Script stopped.")

    # app.PrintInfo(prj)
    Network_Modell = prj.SearchObject("Network Model.IntPrjfolder")
    Network_Model=Network_Modell[0]

    Network_Dataa = Network_Model.SearchObject("Network Data.IntPrjfolder")
    Network_Data = Network_Dataa[0]
    # app.PrintInfo(Network_Data)

    Nett = Network_Data.SearchObject("*.ElmNet")
    Net = Nett[0]
    # app.PrintInfo('Net')
    # app.PrintInfo(Net)

    Libraryy = prj.SearchObject("Library.IntPrjfolder")
    Library = Libraryy[0]
    # app.PrintInfo(Library)

    user_defined_modell = Library.SearchObject("User Defined Models.IntPrjfolder")
    user_defined_model = user_defined_modell[0]
    # app.PrintInfo(user_defined_model)

    BPA_Frame_BB = user_defined_model.SearchObject("BPA Frame(E).BlkDef")
    BPA_Frame_B = BPA_Frame_BB[0]
    # app.PrintInfo(BPA_Frame_B)
    
    Zoness = Network_Data.SearchObject("Zones.IntZone")
    Zones = Zoness[0]
    # app.PrintInfo(Zones)

    for line in bpa_str_ar:
        line = line.rstrip('\n')
        line = line[0:80]
        line = line + ' '*(80-len(line))

        if line == "": continue

        # app.PrintPlain(line.encode('GBK'))
        if line[0] == 'L':
            chinese_count_from = 0
            #判断中文个数
            for i in range (6,14):
            	if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
            	    chinese_count_from = chinese_count_from + 1
            name_from = line[6:14-chinese_count_from].strip()
            base_from = float(line[14-chinese_count_from:18-chinese_count_from])
            name_from = name_from + '   ' + str(base_from)
        	
            chinese_count_to = 0 
            #判断中文个数
            for i in range (19-chinese_count_from,27-chinese_count_from):
            	if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
        	    	chinese_count_to = chinese_count_to + 1
            name_to = line[19-chinese_count_from:27-chinese_count_from-chinese_count_to].strip()
            base_to = float(line[27-chinese_count_from-chinese_count_to:31-chinese_count_from-chinese_count_to])
            name_to = name_to + '   ' + str(base_to)

            #起始bus
            bus_from = bus_name[bus_name_cn.index('Bus_' + name_from)]   
            # app.PrintPlain(bus_from)
            bus_from = Net.SearchObject(bus_from)
            bus_from = bus_from[0]

            #终止bus
            bus_to = bus_name[bus_name_cn.index('Bus_' + name_to)]
            # app.PrintPlain(bus_to)
            bus_to = Net.SearchObject(bus_to)
            bus_to = bus_to[0]

            if line[1] == ' ':
                if myFloat(line[44-chinese_count_from-chinese_count_to:50-chinese_count_from-chinese_count_to].strip().rstrip('.'), 0) > 0:
                     #新建ElmLne   
                    transLine_index = transLine_index + 1
                    transLine = Net.SearchObject('transLine' + str(transLine_index))
                    transLine = transLine[0]
                    if transLine == None:
                        transLine = Net.CreateObject('ElmLne', 'transLine' + str(transLine_index))
                        transLine = transLine[0]

                    cubic = bus_from.SearchObject('Cubic_' + 'transLine' + str(transLine_index))
                    cubic = cubic[0]
                    if cubic == None:
                        cubic = bus_from.CreateObject('StaCubic', 'Cubic_' + 'transLine' + str(transLine_index))
                        cubic = cubic[0]
                    transLine.bus1 = cubic

                    cubic = bus_to.SearchObject('Cubic_' + 'transLine' + str(transLine_index))
                    cubic = cubic[0]
                    if cubic == None:
                        cubic = bus_to.CreateObject('StaCubic', 'Cubic_' + 'transLine' + str(transLine_index))
                        cubic = cubic[0]
                    transLine.bus2 = cubic

                    #保存线路名称
                    chinese_count_line = 0 
                    for i in range (66-chinese_count_from-chinese_count_to, 74-chinese_count_from-chinese_count_to):
                        if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                            chinese_count_line = chinese_count_line + 1
                    name_line = line[66-chinese_count_from-chinese_count_to: 74-chinese_count_from-chinese_count_to-chinese_count_line].strip()

                    name_transLine = name_from + '_' + name_to + '_' + line[31-chinese_count_from-chinese_count_to: 32-chinese_count_from-chinese_count_to]
                    transLine_name_cn.append(name_transLine)
                    transLine_name.append('transLine' + str(transLine_index))

                    # app.PrintInfo(name_line.encode('GBK'))

                    #新建typeLines
                    TypLne_name = 'TypeLine_' + 'transLine' + str(transLine_index)
                    TypLne = Library.SearchObject(TypLne_name)
                    TypLne = TypLne[0]
                    if TypLne == None:
                    	TypLne = Library.CreateObject('TypLne',TypLne_name)
                    	TypLne = TypLne[0]
                    TypLne.uline = base_from

                    R_pu = myFloat(line[38-chinese_count_from-chinese_count_to:44-chinese_count_from-chinese_count_to].strip().rstrip('.'), 0)
                    X_pu = myFloat(line[44-chinese_count_from-chinese_count_to:50-chinese_count_from-chinese_count_to].strip().rstrip('.'), 0)
                    B_pu = myFloat(line[56-chinese_count_from-chinese_count_to:62-chinese_count_from-chinese_count_to].strip().rstrip('.'), 0)
                    Current = myFloat(line[33-chinese_count_from-chinese_count_to:37-chinese_count_from-chinese_count_to].strip().rstrip('.'), 1)
                    transLineLength = myFloat(line[62-chinese_count_from-chinese_count_to:66-chinese_count_from-chinese_count_to].strip().rstrip('.'), 1/MileToKm)

                    transLine.dline = transLineLength * MileToKm
                    TypLne.sline = Current
                    TypLne.rline = base_from * base_from / MVABASE / (MileToKm * transLineLength) * R_pu
                    TypLne.xline = base_from * base_from / MVABASE / (MileToKm * transLineLength) * X_pu
                    TypLne.bline = 2 * pow(10, 6) * MVABASE / base_from / base_from / (MileToKm * transLineLength) * B_pu

                    transLine.typ_id = TypLne
                else:
                    #Scap
                    transLine_index = transLine_index + 1
                    scap = Net.SearchObject('scap' + str(transLine_index))
                    scap = scap[0]
                    if scap == None:
                        scap = Net.CreateObject('ElmScap', ('scap' + str(transLine_index)))
                        scap = scap[0]
                    cubic = bus_from.SearchObject('Cubic_' + 'scap' + str(transLine_index))
                    cubic = cubic[0]
                    if cubic == None:
                        cubic = bus_from.CreateObject('StaCubic', 'Cubic_' + 'scap' + str(transLine_index))
                        cubic = cubic[0]
                    scap.bus1 = cubic

                    cubic = bus_to.SearchObject('Cubic_' + 'scap' + str(transLine_index))
                    cubic = cubic[0]
                    if cubic == None:
                        cubic = bus_to.CreateObject('StaCubic', 'Cubic_' + 'scap' + str(transLine_index))
                        cubic = cubic[0]
                    scap.bus2 = cubic

                    scap.ucn = base_from

                    name_scap = 'Scap_' + name_from + '_' + name_to + '_' + line[31-chinese_count_from-chinese_count_to: 32-chinese_count_from-chinese_count_to]
                    scap_name_cn.append(name_scap)
                    scap_name.append('scap' + str(transLine_index))


            if line[1] == '+':
                if myFloat(line[33-chinese_count_from-chinese_count_to:38-chinese_count_from-chinese_count_to].strip().rstrip('.'), 0) > 0:
                    name_shunt = name_from + '_' + name_to + '_' + line[31-chinese_count_from-chinese_count_to: 32-chinese_count_from-chinese_count_to] + '_i'
                    shunt_index = shunt_index + 1
                    shunt_name_cn.append(name_shunt)
                    shunt_name.append('shunt' + str(shunt_index))

                    shunt = Net.SearchObject('shunt' + str(shunt_index))
                    shunt = shunt[0]
                    if shunt == None:
                        shunt = Net.CreateObject('ElmShnt', 'shunt' + str(shunt_index))
                        shunt = shunt[0]
                    cubic = bus_from.SearchObject('Cubic_shunt' + str(shunt_index))
                    cubic = cubic[0]
                    if cubic == None:
                        cubic = bus_from.CreateObject('StaCubic', 'Cubic_shunt' + str(shunt_index))
                        cubic = cubic[0]
                    shunt.bus1 = cubic
                    shunt.ushnm = base_from
                    shunt.shtype = 1
                    shunt.grea = 9999
                    shunt.qrean = myFloat(line[33-chinese_count_from-chinese_count_to:38-chinese_count_from-chinese_count_to].strip().rstrip('.'), 0)

                if myFloat(line[43-chinese_count_from-chinese_count_to:48-chinese_count_from-chinese_count_to].strip().rstrip('.'), 0) > 0:
                    name_shunt = name_from + '_' + name_to + '_' + line[31-chinese_count_from-chinese_count_to: 32-chinese_count_from-chinese_count_to] + '_j'
                    shunt_index = shunt_index + 1
                    shunt_name_cn.append(name_shunt)
                    shunt_name.append('shunt' + str(shunt_index))

                    shunt = Net.SearchObject('shunt' + str(shunt_index))
                    shunt = shunt[0]
                    if shunt == None:
                        shunt = Net.CreateObject('ElmShnt', 'shunt' + str(shunt_index))
                        shunt = shunt[0]
                    cubic = bus_to.SearchObject('Cubic_shunt' + str(shunt_index))
                    cubic = cubic[0]
                    if cubic == None:
                        cubic = bus_to.CreateObject('StaCubic', 'Cubic_shunt' + str(shunt_index))
                        cubic = cubic[0]
                    shunt.bus1 = cubic
                    shunt.ushnm = base_to
                    shunt.shtype = 1
                    shunt.grea = 9999
                    shunt.qrean = myFloat(line[43-chinese_count_from-chinese_count_to:48-chinese_count_from-chinese_count_to].strip().rstrip('.'), 0)

# ----------------------------------------------------------------------------------------------------
def GetTCard(bpa_file, bpa_str_ar):
    #T卡
    global MVABASE
    
    global bus_name_cn
    global bus_name
    global bus_index
    
    global generator_name_cn
    global generator_name
    global generator_index
    
    global shunt_index
    
    global load_index
    
    global transLine_index
    
    global transformers_index
    global transformers_name_cn
    global transformers_name

    prj = app.GetActiveProject()
    if prj is None:
        raise Exception("No project activated. Python Script stopped.")

    # app.PrintInfo(prj)
    Network_Modell = prj.SearchObject("Network Model.IntPrjfolder")
    Network_Model=Network_Modell[0]

    Network_Dataa = Network_Model.SearchObject("Network Data.IntPrjfolder")
    Network_Data = Network_Dataa[0]
    # app.PrintInfo(Network_Data)

    Nett = Network_Data.SearchObject("*.ElmNet")
    Net = Nett[0]
    # app.PrintInfo('Net')
    # app.PrintInfo(Net)

    Libraryy = prj.SearchObject("Library.IntPrjfolder")
    Library = Libraryy[0]
    # app.PrintInfo(Library)

    user_defined_modell = Library.SearchObject("User Defined Models.IntPrjfolder")
    user_defined_model = user_defined_modell[0]
    # app.PrintInfo(user_defined_model)

    BPA_Frame_BB = user_defined_model.SearchObject("BPA Frame(E).BlkDef")
    BPA_Frame_B = BPA_Frame_BB[0]
    # app.PrintInfo(BPA_Frame_B)
    
    Zoness = Network_Data.SearchObject("Zones.IntZone")
    Zones = Zoness[0]
    # app.PrintInfo(Zones)

    for line in bpa_str_ar:
        line = line.rstrip('\n')
        line = line[0:80]
        line = line + ' '*(80-len(line))
        if line == "": continue
        
        if line[0] == 'T':
            chinese_count_from = 0
            #判断中文个数
            for i in range (6,14):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count_from = chinese_count_from + 1
            name_from = line[6:14-chinese_count_from].strip()
            base_from = float(line[14-chinese_count_from:18-chinese_count_from])
            name_from = name_from + '   ' + str(base_from)
            
            chinese_count_to = 0 
            #判断中文个数
            for i in range (19-chinese_count_from,27-chinese_count_from):
                if line[i] >= u'\u4e00' and line[i] <= u'\u9fa5':
                    chinese_count_to = chinese_count_to + 1
            name_to = line[19-chinese_count_from:27-chinese_count_from-chinese_count_to].strip()
            base_to = float(line[27-chinese_count_from-chinese_count_to:31-chinese_count_from-chinese_count_to])
            name_to = name_to + '   ' + str(base_to)

            transformers_index = transformers_index + 1
            name_transformers =  'Transformers_' + name_from + '_' + name_to + '_' + line[31-chinese_count_from-chinese_count_to: 32-chinese_count_from-chinese_count_to]
            transformers_name_cn.append(name_transformers)
            transformers_name.append('transformers' + str(transformers_index))

            transformers = Net.SearchObject('transformers' + str(transformers_index))
            transformers = transformers[0]
            if transformers == None:
                transformers = Net.CreateObject('ElmTr2', 'transformers' + str(transformers_index)) #新建ElmTr2
                transformers = transformers[0]
                
            if base_from > base_to: 
                base_hv = base_from; base_lv = base_to
                name_hv = name_from; name_lv = name_to
            else: 
                base_hv = base_to; base_lv = base_from
                name_hv = name_to; name_lv = name_from
                
            #高压bus
            bus_hv = bus_name[bus_name_cn.index('Bus_' + name_hv)]
            app.PrintPlain(bus_hv)
            bus_hv = Net.SearchObject(bus_hv)
            bus_hv = bus_hv[0]
            cubic = bus_hv.SearchObject('Cubic_' + 'transformers' + str(transformers_index))
            cubic = cubic[0]
            if cubic == None:
                cubic = bus_hv.CreateObject('StaCubic', 'Cubic_' + 'transformers' + str(transformers_index))
                cubic = cubic[0]
            transformers.bushv = cubic
            #低压bus
            bus_lv = bus_name[bus_name_cn.index('Bus_' + name_lv)]
            app.PrintPlain(bus_lv)
            bus_lv = Net.SearchObject(bus_lv)
            bus_lv = bus_lv[0]
            cubic = bus_lv.SearchObject('Cubic_' + 'transformers' + str(transformers_index))
            cubic = cubic[0]
            if cubic == None:
                cubic = bus_lv.CreateObject('StaCubic', 'Cubic_' + 'transformers' + str(transformers_index))
                cubic = cubic[0]
            transformers.buslv = cubic
            #新建typeLines
            TypTr_name = 'TypeLine_' + 'transformers' + str(transformers_index)
            TypTr = Library.SearchObject(TypTr_name)
            TypTr = TypTr[0]
            if TypTr == None:
                TypTr = Library.CreateObject('TypTr2',TypTr_name)
                TypTr = TypTr[0]
            # if float(line[62-chinese_count_from-chinese_count_to:67-chinese_count_from-chinese_count_to].strip().rstrip('.')) > float(line[67-chinese_count_from-chinese_count_to:72-chinese_count_from-chinese_count_to].strip().rstrip('.')):
            #     TypTr.utrn_h = float(line[62-chinese_count_from-chinese_count_to:67-chinese_count_from-chinese_count_to].strip().rstrip('.'))
            #     TypTr.utrn_l = float(line[67-chinese_count_from-chinese_count_to:72-chinese_count_from-chinese_count_to].strip().rstrip('.'))
            # else:   
            #     TypTr.utrn_l = float(line[62-chinese_count_from-chinese_count_to:67-chinese_count_from-chinese_count_to].strip().rstrip('.'))
            #     TypTr.utrn_h = float(line[67-chinese_count_from-chinese_count_to:72-chinese_count_from-chinese_count_to].strip().rstrip('.'))
            TypTr.utrn_h = base_hv
            TypTr.utrn_l = base_lv
            if myFloat(line[33-chinese_count_from-chinese_count_to:37-chinese_count_from-chinese_count_to].strip().rstrip('.'), 0) > 0: 
                TypTr.strn = myFloat(line[33-chinese_count_from-chinese_count_to:37-chinese_count_from-chinese_count_to].strip().rstrip('.'), 0)
            else:
                TypTr.strn = MVABASE

            TypTr.uktr = float(line[44-chinese_count_from-chinese_count_to:50-chinese_count_from-chinese_count_to].strip().rstrip('.')) * TypTr.strn
            TypTr.pcutr = 1000 / MVABASE * myFloat(line[38-chinese_count_from-chinese_count_to:44-chinese_count_from-chinese_count_to].strip().rstrip('.'), 0) * TypTr.strn * TypTr.strn

            transformers.typ_id = TypTr
                		
# ----------------------------------------------------------------------------------------------------
DEBUG = 1
if DEBUG:

    bpa_file = OpenFile('win') # 打开指定文件

            
if bpa_file:					#If the file opened successfully
    
    global MVABASE
    global MileToKm

    MVABASE = 100
    MileToKm = 1.609344

    bpa_str = bpa_file.read()            #The string containing the text file, to use the find() function
    bpa_file.seek(0)        #To position back at the beginning
    bpa_str_ar = bpa_file.readlines()        #The array that is containing all the lines of the BPA file
    
    for line in bpa_str_ar: 
        if line[0:2] == "/M": 
            MVABASE = float(line[line.find("=")+1:line.find("\\")].lstrip())            #To continue if it is a blank line
            break

    # for i in range(0, len(bpa_str_ar)):
    #     bpa_str_ar[i] = bpa_str_ar[i][0:80]
    #     bpa_str_ar[i] = bpa_str_ar[i] + ' '*(80-len(bpa_str_ar[i]))
                    
    global bus_name_cn
    global bus_name
    global bus_index
    
    global generator_name_cn
    global generator_name
    global generator_index
    
    global shunt_index
    global shunt_name
    global shunt_name_cn
    
    global load_index
    
    global transLine_index
    global transLine_name_cn
    global transLine_name

    global scap_name_cn
    global scap_name
    
    global transformers_index
    global transformers_name_cn
    global transformers_name
    
    bus_name_cn = []
    bus_name = []
    bus_index = 0
    generator_name_cn = []
    generator_name = []
    generator_index = 0
    shunt_index = 0
    shunt_name = []
    shunt_name_cn = []
    load_index = 0
    scap_name_cn = []
    scap_name = []
    transLine_index = 0
    transLine_name_cn = []
    transLine_name = []
    transformers_index = 0
    transformers_name_cn = []
    transformers_name = []


    GetBCard(bpa_file, bpa_str_ar[0:52])

    GetLCard(bpa_file, bpa_str_ar[0:52])

    GetTCard(bpa_file, bpa_str_ar[0:52])

    changeName()
