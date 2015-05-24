# -*- coding: gbk -*-
import os
import string
from math import sqrt
PQPV_threshold = 10000.0    # B卡有有功输出时，小于该值转换为负荷，大于该值时转换为发电机
windmachine_threshold = 5   # 风机等效门限

#  B卡和A卡的读取可能还存在bug

# ----------------------------------------------------------------------------------------------------
def isfloat(str):
    """Checks if the string is a floating point number."""

    try:
        float(str)
        return True			#Returns true if the string is a floating point number
    except (ValueError, TypeError), e:
        return False			#Returns false otherwise

# ----------------------------------------------------------------------------------------------------
def isint(str):
    """Checks if the string is an integer."""

    try:
        int(str)
        return True			#Returns true if the string is an integer
    except (ValueError, TypeError), e:
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
def OpenFile():
    """Opens the BPA file specified in prompt."""

    psspy.prompt("Enter the BPA file path, followed by its name:\n")
    ierr, fnamestr = psspy.userin()		#User types in the path name

    if ierr != 0: return
#    fnamestr='C:\\Users\\eexinhh.eexinhh-THINK\\Desktop\\DW\\xb15sfn.dat'

#    fnamestr='C:\\Users\\eexinhh.eexinhh-THINK\\Desktop\\DW\\zjbpa1.dat'
    fpath, fext = os.path.splitext(fnamestr)
    if not fext: fnamestr = fpath + '.dat'	#To add the extension if left blank

    if os.path.isfile(fnamestr) == False:
        psspy.alert("The specified path or file name is invalid\n")
        return

    bpa_file = open(fnamestr, 'r')		#Opens the file in read mode

    return bpa_file

# ----------------------------------------------------------------------------------------------------
def GetMVA(bpa_file, bpa_str):
    """Gets the base MVA in the BPA file."""

    pos = string.find(bpa_str, "MVA_BASE")
    if pos != -1:
        bpa_file.seek(pos)
        bpa_file.readline()			#Just to position on to the next line
        mva_base_str = bpa_file.readline()
        mva_base_str = mva_base_str.replace(' ', '')	#To delete any spaces
        mva_base_str = mva_base_str.strip()	#To remove trailing and leading whitespaces
        mva_base_str = mva_base_str[10:]	        #To remove the "/MVABASE=" part
        basemva = float(mva_base_str[:-1])      # 去掉最后一个字符'\'

    else:
        basemva = 100.0				#The default base MVA in BPA is 100 MVA

    return basemva

# ----------------------------------------------------------------------------------------------------
def GetTitles(bpa_file, bpa_str):
    """Gets the titles from the BPA file."""

    pos = string.find(bpa_str, "CASEID")	#To find title1

    if pos != -1:
        bpa_file.seek(pos)
        #bpa_file.readline()			#Just to position on to the next line
        title_data = bpa_file.readline()
        pos1 = string.find(title_data, '=')
        pos2 = string.find(title_data, ',')
        case_id = title_data[pos1 + 1 : pos2]
        case_id = case_id.strip()		#To remove trailing and leading whitespaces

    else:
        case_id = ""				#Title1 is blank by default


    pos = string.find(bpa_str, "PROJECT")	#To find title2

    if pos != -1:
        bpa_file.seek(pos)
        title_data = bpa_file.readline()
        pos1 = string.find(title_data, '=')
        pos2 = string.find(title_data, ',')	
        if pos2 == -1:				#If there are no other parameters, the line finishes with ')'
            pos2 = string.find(title_data, ')')

        project_id = title_data[pos1 + 1 : pos2]
        project_id = project_id.strip()		#To remove trailing and leading whitespaces

    else:
        project_id = ""				#Title2 is blank by default

    return case_id, project_id

# ----------------------------------------------------------------------------------------------------
def GetPCard(bpa_str_ar): # 返回P卡数据
    """Gets the scale data from the BPA file."""

    scale_str = []

    for line in bpa_str_ar:			#Loop over every line of the BPA file
        #line = line.lstrip() 
        line = line.rstrip('\n')
        if line == "": continue			#To continue if it is a blank line

        if line[0] == 'P' and line[2] == ' ':	#If the line is a P card
            line = line + ' '*(80-len(line))	#To pad each line with spaces up to 80 records
            scale_str.append(line)

    return scale_str				#Returns an array of P cards

# ----------------------------------------------------------------------------------------------------
def GetScaleData(scale_str, owner_name, zone_name, type_code):
    """Gets the scale factors for each load and generator."""

    scale_ar = [1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0]		#Default factors are unit factors

    for line in scale_str:					#For each P card
        if line[1] == 'O' and line[3:6].strip() == owner_name:
            if zone_name == line[34:36].strip() or zone_name == line[37:39].strip() or zone_name == line[40:42].strip() or zone_name == line[43:45].strip()\
            or zone_name == line[46:48].strip() or zone_name == line[49:51].strip() or zone_name == line[52:54].strip() or zone_name == line[55:57].strip()\
            or zone_name == line[58:60].strip() or zone_name == line[61:63].strip() or zone_name == line[64:66].strip() or zone_name == line[67:69].strip()\
            or zone_name == line[70:72].strip() or zone_name == line[73:75].strip() or zone_name == line[76:78].strip() or line[34:80].strip() == "":

                if isfloat(line[9:14]): scale_ar[0] = scale_ar[0] * float(line[9:14])		#Load P factor
                if isfloat(line[15:20]): scale_ar[1] = scale_ar[1] * float(line[15:20])		#Load Q factor
                elif isfloat(line[9:14]): scale_ar[1] = scale_ar[1] * float(line[9:14])		#Load Q factor is load P factor

                if isfloat(line[21:26]): scale_ar[6] = scale_ar[6] * float(line[21:26])		#Generation P factor
                if isfloat(line[27:32]): scale_ar[7] = scale_ar[7] * float(line[27:32])		#Generation Q factor
                elif isfloat(line[21:26]): scale_ar[7] = scale_ar[7] * float(line[21:26])	#Generation Q factor is generation P factor

        elif line[1] == 'Z' and line[3:5].strip() == zone_name or line[1] == 'N' and line[3:5].strip() == zone_name and type_code == 'N':
            if owner_name == line[34:37].strip() or owner_name == line[38:41].strip() or owner_name == line[42:45].strip() or owner_name == line[46:49].strip()\
            or owner_name == line[50:53].strip() or owner_name == line[54:57].strip() or owner_name == line[58:61].strip() or owner_name == line[62:65].strip()\
            or owner_name == line[66:69].strip() or owner_name == line[70:73].strip() or owner_name == line[74:77].strip() or line[34:80].strip() == "":

                if isfloat(line[9:14]): scale_ar[0] = scale_ar[0] * float(line[9:14])		#Load P factor
                if isfloat(line[15:20]): scale_ar[1] = scale_ar[1] * float(line[15:20])		#Load Q factor
                elif isfloat(line[9:14]): scale_ar[1] = scale_ar[1] * float(line[9:14])		#Load Q factor is load P factor

                if isfloat(line[21:26]): scale_ar[6] = scale_ar[6] * float(line[21:26])		#Generation P factor
                if isfloat(line[27:32]): scale_ar[7] = scale_ar[7] * float(line[27:32])		#Generation Q factor
                elif isfloat(line[21:26]): scale_ar[7] = scale_ar[7] * float(line[21:26])	#Generation Q factor is generation P factor

        elif line[1] == 'A':
            if isfloat(line[9:14]): scale_ar[0] = scale_ar[0] * float(line[9:14])		#Load P factor
            if isfloat(line[15:20]): scale_ar[1] = scale_ar[1] * float(line[15:20])		#Load Q factor
            elif isfloat(line[9:14]): scale_ar[1] = scale_ar[1] * float(line[9:14])		#Load Q factor is load P factor

            if isfloat(line[21:26]): scale_ar[6] = scale_ar[6] * float(line[21:26])		#Generation P factor
            if isfloat(line[27:32]): scale_ar[7] = scale_ar[7] * float(line[27:32])		#Generation Q factor
            elif isfloat(line[21:26]): scale_ar[7] = scale_ar[7] * float(line[21:26])		#Generation Q factor is generation P factor

        elif line[1] == 'B' and type_code == 'X' and line[3:6].strip() == owner_name:
            if zone_name == line[34:36].strip() or zone_name == line[37:39].strip() or zone_name == line[40:42].strip() or zone_name == line[43:45].strip()\
            or zone_name == line[46:48].strip() or zone_name == line[49:51].strip() or zone_name == line[52:54].strip() or zone_name == line[55:57].strip()\
            or zone_name == line[58:60].strip() or zone_name == line[61:63].strip() or zone_name == line[64:66].strip() or zone_name == line[67:69].strip()\
            or zone_name == line[70:72].strip() or zone_name == line[73:75].strip() or zone_name == line[76:78].strip() or line[34:80].strip() == "":

                if isfloat(line[9:14]): scale_ar[2] = scale_ar[2] * float(line[9:14])		#Constant current load P factor
                if isfloat(line[15:20]): scale_ar[3] = scale_ar[3] * float(line[15:20])		#Constant current load Q factor
                elif isfloat(line[9:14]): scale_ar[3] = scale_ar[3] * float(line[9:14])		#Constant current load Q factor is load P factor

                if isfloat(line[21:26]): scale_ar[4] = scale_ar[4] * float(line[21:26])		#Constant admittance load P factor
                if isfloat(line[27:32]): scale_ar[5] = scale_ar[5] * float(line[27:32])		#Constant admittance load Q factor
                elif isfloat(line[21:26]): scale_ar[5] = scale_ar[5] * float(line[21:26])	#Constant admittance load Q factor is load P factor

        elif line[1] == 'B' and type_code == 'Y' and line[3:6].strip() == owner_name:
            if zone_name == line[34:36].strip() or zone_name == line[37:39].strip() or zone_name == line[40:42].strip() or zone_name == line[43:45].strip()\
            or zone_name == line[46:48].strip() or zone_name == line[49:51].strip() or zone_name == line[52:54].strip() or zone_name == line[55:57].strip()\
            or zone_name == line[58:60].strip() or zone_name == line[61:63].strip() or zone_name == line[64:66].strip() or zone_name == line[67:69].strip()\
            or zone_name == line[70:72].strip() or zone_name == line[73:75].strip() or zone_name == line[76:78].strip() or line[34:80].strip() == "":

                if isfloat(line[21:26]): scale_ar[4] = scale_ar[4] * float(line[21:26])		#Constant admittance load P factor
                if isfloat(line[27:32]): scale_ar[5] = scale_ar[5] * float(line[27:32])		#Constant admittance load Q factor
                elif isfloat(line[21:26]): scale_ar[5] = scale_ar[5] * float(line[21:26])	#Constant admittance load Q factor is load P factor

        elif line[1] == 'C' and type_code == 'X' and line[3:5].strip() == zone_name:
            if owner_name == line[34:37].strip() or owner_name == line[38:41].strip() or owner_name == line[42:45].strip() or owner_name == line[46:49].strip()\
            or owner_name == line[50:53].strip() or owner_name == line[54:57].strip() or owner_name == line[58:61].strip() or owner_name == line[62:65].strip()\
            or owner_name == line[66:69].strip() or owner_name == line[70:73].strip() or owner_name == line[74:77].strip() or line[34:80].strip() == "":

                if isfloat(line[9:14]): scale_ar[2] = scale_ar[2] * float(line[9:14])		#Constant current load P factor
                if isfloat(line[15:20]): scale_ar[3] = scale_ar[3] * float(line[15:20])		#Constant current load Q factor
                elif isfloat(line[9:14]): scale_ar[3] = scale_ar[3] * float(line[9:14])		#Constant current load Q factor is load P factor

                if isfloat(line[21:26]): scale_ar[4] = scale_ar[4] * float(line[21:26])		#Constant admittance load P factor
                if isfloat(line[27:32]): scale_ar[5] = scale_ar[5] * float(line[27:32])		#Constant admittance load Q factor
                elif isfloat(line[21:26]): scale_ar[5] = scale_ar[5] * float(line[21:26])	#Constant admittance load Q factor is load P factor

        elif line[1] == 'C' and type_code == 'Y' and line[3:5].strip() == zone_name:
            if owner_name == line[34:37].strip() or owner_name == line[38:41].strip() or owner_name == line[42:45].strip() or owner_name == line[46:49].strip()\
            or owner_name == line[50:53].strip() or owner_name == line[54:57].strip() or owner_name == line[58:61].strip() or owner_name == line[62:65].strip()\
            or owner_name == line[66:69].strip() or owner_name == line[70:73].strip() or owner_name == line[74:77].strip() or line[34:80].strip() == "":

                if isfloat(line[21:26]): scale_ar[4] = scale_ar[4] * float(line[21:26])		#Constant admittance load P factor
                if isfloat(line[27:32]): scale_ar[5] = scale_ar[5] * float(line[27:32])		#Constant admittance load Q factor
                elif isfloat(line[21:26]): scale_ar[5] = scale_ar[5] * float(line[21:26])	#Constant admittance load Q factor is load P factor

    return scale_ar		#Returns the scale factors
 
# ----------------------------------------------------------------------------------------------------
def CheckICard(bpa_str_ar, area_name):
    """To check if there is an I card with the specified area name."""

    for line in bpa_str_ar:			#Loop over every line of the BPA file
        #line = line.lstrip() 
        line = line.rstrip('\n')
        if line == "": continue			#To continue if it is a blank line
        
        if line[0] == 'I' and line[2] == ' ':	#If it is an I card
            line = line + ' '*(34-len(line))	#To pad each line with spaces up to 34 records
            if line[3:13].strip() == area_name or line[14:24].strip() == area_name: return True		#Returns true if the area is found in an I card

    return False			#Returns false if the area has not been found in I cards

# ----------------------------------------------------------------------------------------------------
def GetXCard(bpa_str_ar, bus_name, Vmax, Vmin, remote_bus_name, RMPCT, remote_bus_nbr, bus_nbr):
    """Gets the switched shunts attached to the specified bus."""

    for line in bpa_str_ar:		#This loop is used to get the corresponding switched shunts
        #line = line.lstrip() 
        line = line.rstrip('\n')
        if line == "": continue		#To continue if it is a blank line
        
        #--------------#
        #Default values#
        #--------------#
        intgar = [0, 0, 0, 0, 0, 0, 0, 0, 1, 0]
        realar = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 1.0]

        if isfloat(line[28:32]): remote_kv = float(line[28:32])	#The remote bus base voltage
        else: remote_kv = 0.0

        if line[0] == 'X' and bus_name == (line[6:14].strip() + str(float(line[14:18]))) and remote_bus_name == (line[20:28].strip() + str(remote_kv)):
            if line[2] == ' ':					#The modification code for a new record
                line = line + ' '*(80-len(line))		#To pad each line with spaces up to 80 records

                if isint(line[32]): intgar[0] = int(line[32])	#Number of steps for block 1
                if isint(line[38]): intgar[1] = int(line[38])	#Number of steps for block 2
                if isint(line[44]): intgar[2] = int(line[44])	#Number of steps for block 3
                if isint(line[50]): intgar[3] = int(line[50])	#Number of steps for block 4
                if isint(line[56]): intgar[4] = int(line[56])	#Number of steps for block 5
                if isint(line[62]): intgar[5] = int(line[62])	#Number of steps for block 6
                if isint(line[68]): intgar[6] = int(line[68])	#Number of steps for block 7
                if isint(line[74]): intgar[7] = int(line[74])	#Number of steps for block 8

                if isfloat(line[33:38]): realar[0] = float(line[33:38])	#Admittance increment for block 1
                if isfloat(line[39:44]): realar[1] = float(line[39:44])	#Admittance increment for block 2
                if isfloat(line[45:50]): realar[2] = float(line[45:50])	#Admittance increment for block 3
                if isfloat(line[51:56]): realar[3] = float(line[51:56])	#Admittance increment for block 4
                if isfloat(line[57:62]): realar[4] = float(line[57:62])	#Admittance increment for block 5
                if isfloat(line[63:68]): realar[5] = float(line[63:68])	#Admittance increment for block 6
                if isfloat(line[69:74]): realar[6] = float(line[69:74])	#Admittance increment for block 7
                if isfloat(line[75:80]): realar[7] = float(line[75:80])	#Admittance increment for block 8

                if isfloat(Vmax): realar[8] = float(Vmax)	#Upper voltage limit
                if isfloat(Vmin): realar[9] = float(Vmin)	#Lower voltage limit
                if isfloat(RMPCT):
                    realar[11] = float(RMPCT)	#The RMPCT
                    if realar[11] > 1.0: realar[11] = realar[11] / 100.0	#To bring back the RMPCT from percent to unit

                intgar[9] = remote_bus_nbr

                count_nbr = -1
                for nbr in realar[0:8]:		#This loop is used to get the reactors first in the list for PSS/E to work
                    count_nbr = count_nbr + 1
                    if nbr < 0.0 and count_nbr != 0:	#If the admittance is negative (reactor) and it is not already first in the list
                        realar.insert(0, nbr)
                        realar.pop(count_nbr + 1)
                        intgar.insert(0, intgar[count_nbr])
                        intgar.pop(count_nbr + 1)

                ierr = psspy.switched_shunt_data(bus_nbr, intgar, realar, '')	#The API used to load switched shunts in PSS/E

# ----------------------------------------------------------------------------------------------------
def GetACard(bpa_str_ar, zone_str, bus_str, bus_nbr, area_str, area_slack_nbr, area_nbr):
    "Gets the area data from the BPA file."""

    if zone_str != "":			#Only look for area number if the zone is not blank
        for line in bpa_str_ar:		#This loop is used to get the corresponding area data
            #line = line.lstrip() 
            line = line.rstrip('\n')
            if line == "": continue	#To continue if the line is blank
            
            if line[0] == 'A' and line[1] != 'O':
                if line[2] == ' ' or line[2] == '+':			#The modification code for a new record
                    line = line + ' '*(95-len(line))	#To pad each line with spaces up to 95 records
                    if zone_str == line[35:37].strip() or zone_str == line[38:40].strip() or zone_str == line[41:43].strip() or zone_str == line[44:46].strip()\
                    or zone_str == line[47:49].strip() or zone_str == line[50:52].strip() or zone_str == line[53:55].strip() or zone_str == line[56:58].strip()\
                    or zone_str == line[59:61].strip() or zone_str == line[62:64].strip() or zone_str == line[65:67].strip() or zone_str == line[68:70].strip()\
                    or zone_str == line[71:73].strip() or zone_str == line[74:76].strip() or zone_str == line[77:79].strip() or zone_str == line[80:82].strip()\
                    or zone_str == line[83:85].strip() or zone_str == line[86:88].strip() or zone_str == line[89:91].strip() or zone_str == line[92:94].strip():

#                        if not bus_str.has_key(line[13:21].strip() + str(float(line[21:25]))):	#If the bus is not in the list, add it
#                            bus_nbr = bus_nbr + 1
#                            bus_str[line[13:21].strip() + str(float(line[21:25]))] = bus_nbr	#A new dictionnary entry
#                            bus_flag = True							#True if the bus number has been incremented
#                        else:
#                            bus_nbr = bus_str[line[13:21].strip() + str(float(line[21:25]))]	#Get the bus number
#                            bus_flag = False							#False if the bus number hasn't been incremented

#  将上面语句修改为下面：
                        if line[2] == ' ' and bus_str.has_key(line[13:21].strip() + str(float(line[21:25]))):
                            bus_slack_nbr = bus_str[line[13:21].strip() + str(float(line[21:25]))]
                            if bus_slack_nbr not in area_slack_nbr:
                                area_slack_nbr.append(bus_slack_nbr)
                        bus_flag = False
                        
#  这样修改的目的是为了在A卡中不增加新的母线，而只是增加Area的平衡节点了。因此，要求数据满足：在A卡中的母线需要在文件中有定义。如果按上面的做法，存在一个bug：一个区域里面可能会导致存在两个平衡节点，或者不存在平衡节点，仿真发现的

                        if isfloat(line[26:34]):					#If the value of interchange is specified
                            if CheckICard(bpa_str_ar, line[3:13].strip()):		#If there is an I card, do not include interchange given
                                if not area_str.has_key((line[3:13].strip(), 0.0)):	#If the area is not in the list, add it
#                                    area_nbr = area_nbr + 1
                                    area_nbr = max(area_str.values()) + 1
                                    area_str[(line[3:13].strip(), 0.0)] = area_nbr
#                                    area_slack_nbr.append(bus_nbr)			#An array with slack bus numbers for each area
                                else:
                                    area_nbr = area_str[(line[3:13].strip(), 0.0)]	#Get the area number
                            else:							#If there is no I card, include the interchange value given
                                if not area_str.has_key((line[3:13].strip(),float(line[26:34]))):	#If the area is not in the list, add it
#                                    area_nbr = area_nbr + 1
                                    area_nbr = max(area_str.values()) + 1
                                    area_str[(line[3:13].strip(), float(line[26:34]))] = area_nbr
#                                    area_slack_nbr.append(bus_nbr)			#Add the corresponding slack bus number
                                else:
                                    area_nbr = area_str[(line[3:13].strip(),float(line[26:34]))]	#Get the area number

                        else:				#If there is no value of interchange specified
                            if not area_str.has_key((line[3:13].strip(), 0.0)):		#If the area is not in the list, add it
#                                area_nbr = area_nbr + 1
                                area_nbr = max(area_str.values()) + 1
                                area_str[(line[3:13].strip(), 0.0)] = area_nbr
#                                area_slack_nbr.append(bus_nbr)				#Add the corresponding slack bus number
                            else:
                                area_nbr = area_str[(line[3:13].strip(), 0.0)]		#Get the area number

                        area_flag = True		#True if the zone name has a corresponding area card

                        return area_flag, bus_flag, area_nbr, bus_nbr, bus_str, area_str, area_slack_nbr


    area_flag = False		#The zone name has no corresponding area card
    bus_flag = False		#The bus number hasn't changed

    return area_flag, bus_flag, area_nbr, bus_nbr, bus_str, area_str, area_slack_nbr

# ----------------------------------------------------------------------------------------------------
def GetBCard(bpa_str_ar, base_mva, scale_str, DEBUG):
    """Gets the AC bus data from the BPA file."""

    bus_data_str = []			#The array to contain all AC bus lines 存储所有节点数据(直流节点数据除外)
    bus_str = {}			#A dictionnary with bus names and the corresponding bus numbers 节点字典
    bus_owner_nbr = [0]			#An array with the owner number for each bus
    bus_zone_str = ["0"]		#An array with the zone of each bus
    load_id_ar = [0]			#An array used to determine the load CKT for each bus
    machine_id_ar = [0]			#An array used to determine the machine CKT for each bus
    owner_str = {'DEFAULT': 1}		#The default owner                                      所有者字典
    zone_str = {'DEFAULT': 1}		#The default zone
    area_str = {("DEFAULT", 0.0): 1}	#The default area is used for zones without A cards or for data without zones
    area_slack_nbr = [0]		#An array with the slack bus of each area
    bus_nbr = 0				#The bus number                                                 节点编号统计
    bus_tot = 0                         #Total bus  
    owner_nbr = 1			#The owner number                                               所有者编号统计
    zone_nbr = 1			#The zone number
    area_nbr = 1			#The area number

    for line in bpa_str_ar:		#This loop is used to fill the bus data array with each line of bus data
        #line = line.lstrip()           # BPA中，开始为空格也表示此行数据为注释
        line = line.rstrip('\n')
        if line == "": continue		#Continue if the line is blank
        if line[0] == 'B':		#If it is a bus
            if line[1] != 'D' and line[1] != 'M':	#But not a DC bus
                bus_data_str.append(line)		#Add the line to the array
#            else:
#                raise('aaaa')    # 如果出现直流报错
                

    for line in bus_data_str:
        line = line + ' '*(80-len(line))	#To pad each line with spaces up to 80 records

        #--------------#
        #Default values#
        #--------------#
        bus_intgar = [1, 0, 1, 0]
        load_intgar = [1, 1, 1, 1]
        machine_intgar = [0, 0, 0, 0, 0, 0]
        bus_realar = [0.0, 0.0, 0.0, 1.0, 0.0]
        load_realar = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
        #machine_realar = [0.0, 0.0, 9999.0, -9999.0, 9999.0, -9999.0, base_mva, 0.0, 1.0, 0.0, 0.0, 0.0, 1.0, 1.0, 1.0, 1.0, 1.0]
        machine_realar = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, base_mva, 0.0, 1.0, 0.0, 0.0, 0.0, 1.0, 1.0, 1.0, 1.0, 1.0]        
        plant_realar = [1.0, 1.0]
        remote_bus_nbr = 0
        remote_kv = 0.0
        load_id = 0                 # 负荷类型标志：1：正常负荷 2：B卡有功输出转换，该B卡负荷为0 3：B卡有功输出转换，该B卡负荷不为0

        

        if line[2] == ' ' or line[2] == '@':			#The modification code for a new record
            if line[1] == ' ' or line[1] == 'T' or line[1] == 'C' or line[1] == 'V' or line[1] == 'F' or line[1] == 'J' or line [1] == 'X':
                bus_intgar[0] = 1		#The bus type code for a PQ bus
            elif line[1] == 'E' or line[1] == 'Q' or line[1] == 'G' or line[1] == 'K' or line[1] == 'L':
                bus_intgar[0] = 2		#The bus type code for a PV bus
            elif line[1] == 'S':
                bus_intgar[0] = 3		#The bus type code for a swing bus

            if line[3:6] != "   " and not owner_str.has_key(line[3:6].strip()):	#If the owner is not in the list, add it to the list with a new number
                owner_nbr = owner_nbr + 1
                owner_str[line[3:6].strip()] = owner_nbr
                bus_intgar[3] = owner_nbr			#The owner number
            elif line[3:6] != "   " and owner_str.has_key(line[3:6].strip()):
                bus_intgar[3] =  owner_str[line[3:6].strip()]	#Get the owner number
            elif line[3:6] == "   ":
                bus_intgar[3] = 1				#The default owner number

            bus_namear = line[6:14].strip()		#The bus name 节点名
            bus_realar[2] = float(line[14:18])		#The bus base voltage in kV 母线额定电压

            if not bus_str.has_key(bus_namear + str(bus_realar[2])):	#If the bus is not in the list, add it (the bus could be in the list if assigned as a remote for another bus)
                bus_tot = bus_tot + 1
#                bus_nbr = bus_nbr + 1   
                bus_nbr = bus_tot            # 防止上次执行的是else里面的内容，故不执行上一句而采用这句
                bus_str[bus_namear + str(bus_realar[2])] = bus_nbr	#A new dictionnary entry                
            else:
                bus_nbr = bus_str[bus_namear + str(bus_realar[2])]	#Get the bus number
                psspy.progress_output(1,"",[0,0])           # 显示
                psspy.progress_output('Remote Bus or Defined duplicatedly ' + str(bus_nbr1)+' '+bus_namear +' '+ str(bus_realar[2])+'\n')
                raise('aaaa')    # 如果出现这种，先中断一下看看排除错误后再注释掉这句开始
                psspy.progress_output(6,"",[0,0])           # 关闭显示

#                                                        
#                          
            bus_owner_nbr.insert(bus_nbr + 1, bus_intgar[3])		#Insert the corresponding owner of the bus

            load_id_ar.append(0)	#To initialize the load CKT number
            machine_id_ar.append(0)	#To initialize the machine CKT number

            if line[18:20] != "  ":	#If the zone is not blank
                if not zone_str.has_key(line[18:20].strip()):	#If the zone is not in the list, add it
                    zone_nbr = zone_nbr + 1
                    zone_str[line[18:20].strip()] = zone_nbr
                    bus_intgar[2] = zone_nbr		#The zone number
                else:
                    bus_intgar[2] = zone_str[line[18:20].strip()]	#Get the zone number

            bus_zone_str.insert(bus_nbr+1, bus_intgar[2])	#Insert the corresponding zone of the bus
                
            load_id = 1
            if isfloat(line[20:25]): load_realar[0] = float(line[20:25])	#Load P in MW
            if isfloat(line[25:30]): load_realar[1] = float(line[25:30])	#Load Q in Mvar

            if isfloat(line[30:34]): bus_realar[0] = float(line[30:34])		#Fixed bus shunt active load
            if isfloat(line[34:38]): bus_realar[1] = float(line[34:38])		#Fixed bus shunt reactive load

            if isfloat(line[38:42]): machine_realar[4] = float(line[38:42])	#Maximum active power generation 
            if isfloat(line[42:47]): machine_realar[0] = float(line[42:47])	#Actual active power generation

            if line[1] != ' ' and line[1] != 'C' and line[1] != 'T' and line[1] != 'V':	#If not a PQ bus
                if isfloat(line[47:52]): machine_realar[2] = float(line[47:52])	#Machine reactive upper limit
                if isfloat(line[52:57]): machine_realar[3] = float(line[52:57])	#Machine reactive lower limit

            if line[1] == ' ' or line[1] == 'C' or line[1] == 'T' or line[1] == 'V':	#If it is a PQ bus
                if max(abs(machine_realar[0]),abs(machine_realar[1])) >= PQPV_threshold:         #判断是转换成负荷还是发电机
                    bus_intgar[0] = 2                           #在阀值以上变为PV节点
                    if isfloat(line[47:52]):
                        machine_realar[1] = float(line[47:52])	#Machine reactive power output
                        machine_realar[2] = machine_realar[1]	  #Machine reactive lower limit
                        machine_realar[3] = machine_realar[1]	  #Machine reactive upper limit
                elif (max(abs(machine_realar[0]),abs(machine_realar[1]))< PQPV_threshold) and machine_realar[0] != 0.0:
                    if isfloat(line[47:52]):
                        machine_realar[1] = float(line[47:52])	#Machine reactive power output
                    else: machine_realar[1] = 0
                    if load_realar[0] == 0: load_id = 2
                    elif load_realar[0] != 0: load_id = 3
                    #load_realar[0] = load_realar[0] * scale_ar[0]		#To scale the load P
                    #load_realar[1] = load_realar[1] * scale_ar[1]		#To scale the load Q
                    #machine_realar[0] = machine_realar[0] * scale_ar[6]		#To scale the machine P
                    #machine_realar[1] = machine_realar[1] * scale_ar[7]		#To scale the machine Q
                    #load_realar[0] = load_realar[0] - machine_realar[0]         #计算有功无功负荷
                    #machine_realar[0] = 0                                       #消除发电机判断依据
                    #machine_realar[4] = 0                                       #消除发电机判断依据
                    #load_realar[1] = load_realar[1] - machine_realar[1]
                    #machine_realar[1] = 0
                    

            elif line[1] == 'E' or line[1] == 'Q' or line[1] == 'G':		#If it is a PV bus
                if isfloat(line[57:61]):
                    float_value = getfloatvalue(line[57:61],3)
                    plant_realar[0] = float_value			#Scheduled voltage
                    bus_realar[3] = float_value				#Bus voltage magnitude
                    if line[1] == 'E':   # 不受无功约束的PV母线，所以需要把无功的上下限改掉
                        machine_realar[2] =  9998	   #Machine reactive upper limit
                        machine_realar[3] = -9998     #Machine reactive lower limit                      
                else: raise('The mag is not set for a PV bus'+str(bus_nbr)) 
            elif line[1] == 'S':						#If it is a swing bus
                if isfloat(line[57:61]):
                    float_value = getfloatvalue(line[57:61],3)
                    plant_realar[0] = float_value			        #Scheduled voltage
                    bus_realar[3] = float_value				        #Bus voltage magnitude
                    
                if isfloat(line[61:65]): bus_realar[4] = getfloatvalue(line[61:65],3)	#The voltage phase angle F4.3需修改

            if line[65:73] != "        " and line[1] in ['X', 'G']:		#If there is a remote bus
                remote_kv = float(line[73:77])					#The remote bus base voltage
                if not bus_str.has_key(line[65:73].strip() + str(remote_kv)):	#If the remote bus name is not in the list, add it
                    remote_bus_nbr = bus_nbr + 1
                    bus_str[line[65:73].strip() + str(remote_kv)] = remote_bus_nbr
                else:
                    remote_bus_nbr = bus_str[line[65:73].strip() + str(remote_kv)]	#Get the remote bus number

            if isfloat(line[77:80]):
                plant_realar[1] = getfloatvalue(line[77:80],2)	#The RMPCT
                #if plant_realar[1] > 1.0001: plant_realar[1] = plant_realar[1]/100.0	#To bring back the RMPCT from percent to unit
                

            area_flag, bus_flag, area_nbr, swing_bus_nbr, bus_str, area_str, area_slack_nbr \
                                                 = GetACard(bpa_str_ar, line[18:20].strip(), bus_str, bus_tot, area_str, area_slack_nbr, area_nbr)
                                                #     = GetACard(bpa_str_ar, line[18:20].strip(), bus_str, bus_nbr, area_str, area_slack_nbr, area_nbr)
                                                     
            if bus_flag == True:
                bus_nbr = swing_bus_nbr	#To change the bus number if it has been incremented
                bus_tot = swing_bus_nbr

            if area_flag == True:		#To assign the area number to the buses and loads
                bus_intgar[1] = area_nbr
                load_intgar[1] = area_nbr
            else:				#The default area number
                bus_intgar[1] = 1
                load_intgar[1] = 1

            if bus_nbr in area_slack_nbr:
                bus_intgar[0] = 3


            ierr = psspy.bus_data(bus_nbr, bus_intgar, bus_realar, bus_namear)	#Loads the bus data in PSS/E

            if line[1] == 'X': GetXCard(bpa_str_ar, bus_namear + str(bus_realar[2]), line[57:61], line[61:65], line[65:73].strip() + str(remote_kv), line[77:80], remote_bus_nbr, bus_nbr)

            scale_ar = GetScaleData(scale_str, line[3:6].strip(), line[18:20].strip(), 'N')	#To get the scale factors

            #if load_id == 1:
            load_realar[0] = load_realar[0] * scale_ar[0]		#To scale the load P
            load_realar[1] = load_realar[1] * scale_ar[1]		#To scale the load Q
            machine_realar[0] = machine_realar[0] * scale_ar[6]		#To scale the machine P
            machine_realar[1] = machine_realar[1] * scale_ar[7]		#To scale the machine Q
            if load_id == 2 or load_id == 3:
                load_realar[0] = load_realar[0] - machine_realar[0]         #计算有功无功负荷
                load_realar[1] = load_realar[1] - machine_realar[1]
                machine_realar[0] = 0                                       #消除发电机判断依据
                machine_realar[1] = 0
                machine_realar[2] = 0
                machine_realar[3] = 0
                machine_realar[4] = 0                                       #消除发电机判断依据

            if  bus_intgar[0] == 2 and machine_realar[2] == machine_realar[3]: #machine_realar[0] <= windmachine_threshold and 
 #               bus_intgar[0] = 1                                             #等效为风机
                machine_intgar[5] = 3
                if (machine_realar[0]**2+machine_realar[2]**2)>0: 
                   machine_realar[16] = machine_realar[0]/sqrt(machine_realar[0]**2+machine_realar[2]**2)   #风机功率因数

            if load_realar[0] != 0.0 or load_realar[1] != 0.0: 			#If there is a load
                load_intgar[2] = bus_intgar[2]		#The zone number
                load_intgar[3] = bus_intgar[3]		#The owner number
                load_id_ar[bus_nbr] = 1			#The load CKT number is now 1
                if DEBUG:
                   if load_id == 1:
                       ierr = psspy.load_data(bus_nbr, '1', load_intgar, load_realar)	#Loads the load data in PSS/E
                   elif load_id == 2:
                       ierr = psspy.load_data(bus_nbr, '2', load_intgar, load_realar)	#Loads the load data in PSS/E
                   elif load_id == 3:
                       ierr = psspy.load_data(bus_nbr, '3', load_intgar, load_realar)	#Loads the load data in PSS/E

            if (bus_intgar[0]!=1) or ( (max(abs(machine_realar[0]),abs(machine_realar[1]))>0.0001)):		#If there is a generator 可以用最大有功来判断
                machine_intgar[0] = 1			    #The machine status
                machine_id_ar[bus_nbr] = 1		#The machine CKT number is now 1
                machine_intgar[1] = bus_intgar[3]	#The owner number
                ierr = psspy.plant_data(bus_nbr, remote_bus_nbr, plant_realar)	        #First load the plant data
                ierr = psspy.machine_data_2(bus_nbr, '1', machine_intgar, machine_realar)	#After add the generator data

            if remote_bus_nbr == bus_nbr + 1:
                bus_nbr = bus_nbr + 1	  #To change the actual bus number
                bus_tot = bus_tot + 1   #增加一个计数                

    bus_nbr =  bus_tot           
    return area_str, area_nbr, area_slack_nbr, owner_str, owner_nbr, bus_owner_nbr, bus_str, bus_nbr, zone_str, zone_nbr, load_id_ar, machine_id_ar, bus_zone_str

# ----------------------------------------------------------------------------------------------------
def GetRCard(bpa_str_ar, from_bus, to_bus, bus_str):
    """Gets the regulation parameters for the two-winding transformers."""

    #--------------#
    #Default values#
    #--------------#
    #Only the returned values are useful
    intgar = [1, 0, 1, 0, 0, 0, 33, 0, 0, 0, 0, 0, 2, 1, 1]
    realari = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 0.0, 1.0, 1.0, 1.0, 1.0, 0.0, 0.0, 1.1, 0.9, 1.1, 0.9, 0.0, 0.0]

    for line in bpa_str_ar:
        #line = line.lstrip() 
        line = line.rstrip('\n')
        if line == "": continue		#Continue if the line is blank

        if line[0] == 'R' and line[1] != 'Z':	#If it is a R card but not RZ
            if line[2] == ' ' and from_bus == (line[6:14].strip() + str(float(line[14:18]))) and to_bus == (line[19:27].strip() + str(float(line[27:31]))):
                line = line + ' '*(67-len(line))	#To pad each line with spaces up to 67 records

                if line[33:41] != "        ": intgar[9] = bus_str[line[33:41].strip() + str(float(line[41:45]))]	#Controlled bus number
                if isfloat(line[45:50]): realari[17] = getfloatvalue(line[45:50],2)	#Winding one ratio/angle high limit F5.2需修改
                if isfloat(line[50:55]): realari[18] = getfloatvalue(line[50:55],2)	#Winding one ratio/angle low limit  F5.2需修改
                if isint(line[55:57]): intgar[6] = int(line[55:57])		#Number of tap positions

                if line[1] == ' ':	#If it is a R card
                    intgar[11] = 1	#Control mode = 1
                elif line[1] == 'V':	#If it is a RV card
                    intgar[11] = 1	#Control mode = 1
 
                elif line[1] == 'Q':	#If it is a RQ card
                    intgar[11] = 2	#Control mode = 2
                elif line[1] == 'N':	#If it is a RN card
                    intgar[11] = 2	#Control mode = 2
                    if isfloat(line[57:62]): realari[19] = float(line[57:62])	#Voltage or flow upper limit
                    else: realari[19] = 0.0					#Default
                    if isfloat(line[62:67]): realari[20] = float(line[62:67])	#Voltage or flow lower limit
                    else: realari[20] = 0.0					#Default

                elif line[1] == 'P':	#If it is a RP card
                    intgar[11] = 3	#Control mode = 3
                elif line[1] == 'M':	#If it is a RM card
                    intgar[11] = 3	#Control mode = 3
                    if isfloat(line[57:62]): realari[19] = float(line[57:62])	#Voltage or flow upper limit
                    else: realari[19] = 0.0					#Default
                    if isfloat(line[62:67]): realari[20] = float(line[62:67])	#Voltage or flow lower limit
                    else: realari[19] = 0.0					#Default

                return intgar[11], intgar[9], realari[17], realari[18], intgar[6], realari[19], realari[20]

    return intgar[11], intgar[9], realari[17], realari[18], intgar[6], realari[19], realari[20]

# ----------------------------------------------------------------------------------------------------
def GetSectionLine(bpa_str_ar, iarg, jarg, section_nbr, dummy_nbr, bus_nbr, bus_str, bus_owner_nbr, owner_nbr, multi_intgar, ckt_nbr, base_mva):
    """To get the multi-section line data from BPA.

    This function is recursive.
    """

    for line in bpa_str_ar:
        #line = line.lstrip() 
        line = line.rstrip('\n')
        if line == "": continue		#Continue if the line is blank

        #--------------#
        #Default values#
        #--------------#
        bus_intgar = [1, 1, 1, owner_nbr]
        bus_realar = [0.0, 0.0, 0.0, 1.0, 0.0]
        line_intgar = [1, 0, 1, 0, 0, 0]
        line_realar = [0.0, 1e-6, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 1.0, 1.0, 1.0, 1.0]
        trans_intgar = [1, 0, 1, 0, 0, 0, 33, 0, 0, 0, 0, 0, 2, 1, 1]
        trans_realari = [0.0, 0.0, base_mva, 1.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 0.0, 1.0, 1.0, 1.0, 1.0, 0.0, 0.0, 1.1, 0.9, 1.1, 0.9, 0.0, 0.0]

        if line[0] == 'L' and line[1] == ' ' and line[2] == ' ' or line[0] == 'E' and line[1] == ' ' and line[2] == ' ':
            if iarg == bus_str[line[6:14].strip() + str(float(line[14:18]))] and jarg == bus_str[line[19:27].strip() + str(float(line[27:31]))] and line[32] == str(section_nbr) and ckt_nbr == line[31]:
                line = line + ' '*(80-len(line))		#To pad each line with spaces up to 80 records
                section_nbr = section_nbr + 1			#To increment the section number
                dummy_nbr = dummy_nbr + 1			#To increment the dummy bus number
                bus_nbr = bus_nbr + 1				#To increment the bus number
                bus_namear = "DUMMY #" + str(dummy_nbr)		#The name of the new dummy bus
                bus_str[bus_namear] = bus_nbr			#Add the bus to the dictionnary
                bus_owner_nbr.insert(bus_nbr + 1, owner_nbr)	#Insert the owner number of that bus
                if isfloat(line[14:18]): bus_realar[2] = float(line[14:18])	#The bus base voltage
                multi_intgar[section_nbr - 2] = bus_nbr		#The bus number of the dummy bus
                from_bus = bus_nbr				#To remember the value of the from bus before recalling this function

                ierr = psspy.bus_data(bus_nbr, bus_intgar, bus_realar, bus_namear)		#Loads the bus data in PSS/E

                to_bus, dummy_nbr, bus_nbr, bus_str, bus_owner_nbr, owner_nbr, multi_intgar = GetSectionLine(bpa_str_ar, iarg, jarg, section_nbr, dummy_nbr, bus_nbr, bus_str, bus_owner_nbr, owner_nbr, multi_intgar, ckt_nbr, base_mva)	#Recursive function

                if line[3:6] != "   " and not owner_str.has_key(line[3:6].strip()):		#If the owner is not in the list, add it to the list with a new number
                    owner_nbr = owner_nbr + 1
                    owner_str[line[3:6].strip()] = owner_nbr		#A new dictionnary entry
                    line_intgar[2] = owner_nbr				#The new owner number
                elif line[3:6] != "   " and owner_str.has_key(line[3:6].strip()):
                    line_intgar[2] =  owner_str[line[3:6].strip()]	#Get the owner number
                elif line[3:6] == "   ":
                    line_intgar[2] = 1			#The default owner

                line_intgar[1] = from_bus		#Every section of the line is set as metered from (default)

                if isfloat(line[33:37]) and isfloat(line[14:18]): line_realar[3] = sqrt(3)*float(line[33:37])*float(line[14:18])/1000.0	#The Rate A in MVA

                if line[31] != ' ': ckt = line[31]		#The circuit identifier
                else: ckt = '1'			#If none is mentionned, the default is 1

                if isfloat(line[38:44]): line_realar[0] = getfloatvalue(line[38:44],5)			#Nominal branch resistance F6.5需修改
                if isfloat(line[44:50]) and float(line[44:50]) != 0.0: line_realar[1] = getfloatvalue(line[44:50],5)	#Nominal branch reactance F6.5需修改

                if line[0] == 'L':		#If it is an L card
                    if isfloat(line[50:56]):
                    #	GI，BI = 在 母线“I”一端上的并联支路的复数导纳，标幺值表示。对一个连接于母线上的电抗器来说BI是负值。缺省情况下，GI+jBI=0。
                    # 线路对地导纳标么值（对称支路只需填一侧值）：G/2（51-56），B/2（57-62）。
                    #在线路节点1端的对地导纳标么值：G1－(51-56)，B1－(57-62)
                        line_realar[6] = getfloatvalue(line[50:56],5)					#Real line shunt at bus IARG end  F6.5需修改
                        line_realar[8] = line_realar[6]						#Real line shunt at bus JARG end         
                    if isfloat(line[56:62]): line_realar[2] = getfloatvalue(line[56:62],5)*2		#Total line charging F6.5需修改
                    if isfloat(line[62:66]): line_realar[10] = getfloatvalue(line[62:66],1)		#Line's length in miles F4.1需修改

                else:				#If it is an E card
                    if isfloat(line[50:56]): line_realar[6] = getfloatvalue(line[50:56],5)		#Real line shunt at bus IARG end F6.5需修改
                    if isfloat(line[56:62]): line_realar[7] = getfloatvalue(line[56:62],5)		#Reactive line shunt at bus IARG end F6.5需修改
                    if isfloat(line[62:68]): line_realar[8] = getfloatvalue(line[62:68],5)		#Real line shunt at bus JARG end F6.5需修改
                    if isfloat(line[68:74]): line_realar[9] = getfloatvalue(line[68:74],5)		#Reactive line shunt at bus JARG end F6.5需修改

                #It seems the format for the date in the file 2006eh-tmp-1.dat is not the same as in the user manual
                #if line[77:80] != "   ": line_intgar[0] = 0		#If there is an out of service date, the status is offline

                ierr = psspy.branch_data(from_bus, to_bus, ckt, line_intgar, line_realar)	#The API to load branch data in PSS/E

                return from_bus, dummy_nbr, bus_nbr, bus_str, bus_owner_nbr, owner_nbr, multi_intgar

        if line[0] == 'T' and line[2] == ' ' and bus_str.has_key(line[6:14].strip() + str(float(line[14:18]))) and bus_str.has_key(line[19:27].strip() + str(float(line[27:31]))):
            if iarg == bus_str[line[6:14].strip() + str(float(line[14:18]))] and jarg == bus_str[line[19:27].strip() + str(float(line[27:31]))] and line[32] == str(section_nbr) and ckt_nbr == line[31]:
                line = line + ' '*(80-len(line))		#To pad each line with spaces up to 80 records
                section_nbr = section_nbr + 1			#To increment the section number
                dummy_nbr = dummy_nbr + 1			#To increment the dummy bus number
                bus_nbr = bus_nbr + 1				#To increment the bus number
                bus_namear = "DUMMY #" + str(dummy_nbr)		#The name of the dummy bus
                bus_str[bus_namear] = bus_nbr			#A new dictionnary entry
                bus_owner_nbr.insert(bus_nbr + 1, owner_nbr)	#Insert the owner of this bus
                if isfloat(line[14:18]): bus_realar[2] = float(line[14:18])	#The bus base voltage
                multi_intgar[section_nbr - 2] = bus_nbr		#The bus number of the dummy bus
                from_bus = bus_nbr

                ierr = psspy.bus_data(bus_nbr, bus_intgar, bus_realar, bus_namear)	#Loads the bus data in PSS/E

                to_bus, dummy_nbr, bus_nbr, bus_str, bus_owner_nbr, owner_nbr, multi_intgar = GetSectionLine(bpa_str_ar, iarg, jarg, section_nbr, dummy_nbr, bus_nbr, bus_str, bus_owner_nbr, owner_nbr, multi_intgar, ckt_nbr, base_mva)	#Recursive function

                if line[3:6] != "   " and not owner_str.has_key(line[3:6].strip()):	#If the owner is not in the list, add it to the list with a new number
                    owner_nbr = owner_nbr + 1
                    owner_str[line[3:6].strip()] = owner_nbr	#A new dictionnary entry
                    trans_intgar[2] = owner_nbr			#The new owner number
                elif line[3:6] != "   " and owner_str.has_key(line[3:6].strip()):
                    trans_intgar[2] =  owner_str[line[3:6].strip()]	#Get the owner number
                elif line[3:6] == "   ":
                    trans_intgar[2] = 1			#The default owner

                trans_intgar[1] = from_bus		#Every section of the line is set as metered from (default)

                trans_intgar[8] = from_bus		#The winding one side

                if isfloat(line[14:18]): trans_realari[4] = float(line[14:18])		#Winding one nominal voltage

                if isfloat(line[27:31]): trans_realari[7] = float(line[27:31])		#Winding two nominal voltage

                if isfloat(line[33:37]): trans_realari[8] = float(line[33:37])		#The Rate A in MVA

                if line[31] != ' ': ckt = line[31]	#The circuit identifier
                else: ckt = '1'				#If none is mentionned, the default is 1

                if isfloat(line[38:44]): trans_realari[0] = getfloatvalue(line[38:44],5)	#Nominal transformer resistance  F6.5需修改
                if isfloat(line[44:50]) and float(line[44:50]) != 0.0: trans_realari[1] = getfloatvalue(line[44:50],5)	#Nominal transformer reactance  F6.5需修改

                if isfloat(line[50:56]): trans_realari[15] = getfloatvalue(line[50:56],5)		#The magnetizing conductance F6.5需修改

                if isfloat(line[56:62]): trans_realari[16] = getfloatvalue(line[56:62],5)		#The magnetization susceptance F6.5需修改

                if line[1] == ' ':
                    if isfloat(line[62:67]): trans_realari[3] = getfloatvalue(line[62:67],2)		#The winding one ratio/voltage  F5.2需修改
                    if isfloat(line[67:72]): trans_realari[6] = getfloatvalue(line[67:72],2)		#The winding one ratio/voltage  F5.2需修改

                else:		#For the phase shifting transformer
                    if isfloat(line[62:67]): trans_realari[5] = getfloatvalue(line[62:67],2)		#The winding one phase shift angle F5.2需修改
                    trans_realari[3] = trans_realari[4]
                    trans_realari[6] = trans_realari[7]

                #It seems the format for the date in the file 2006eh-tmp-1 is not the same as in the user manual
                #if line[77:80] != "   ": trans_intgar[0] = 0		#If there is an out of service date, the status is offline

                trans_intgar[11], trans_intgar[9], trans_realari[17], trans_realari[18], trans_intgar[6], trans_realari[19], trans_realari[20] = GetRCard(bpa_str_ar, line[6:14].strip() + str(float(line[14:18])), line[19:27].strip() + str(float(line[27:31])), bus_str)

                ierr, realaro  = psspy.two_winding_data(from_bus, to_bus, ckt, trans_intgar, trans_realari, "")	#The API to load transformer data in PSS/E

                return from_bus, dummy_nbr, bus_nbr, bus_str, bus_owner_nbr, owner_nbr, multi_intgar

    return jarg, dummy_nbr, bus_nbr, bus_str, bus_owner_nbr, owner_nbr, multi_intgar

# ----------------------------------------------------------------------------------------------------
def GetLCard(bpa_str_ar, owner_str, owner_nbr, bus_str, bus_nbr, bus_owner_nbr, base_mva, DEBUG):
    """Gets the symmetrical and asymmetrical lines from BPA."""

    line_str = []	#An array to contain all AC line data
    dummy_nbr = 0	#The dummy bus number for multi-section lines
    line_counter = 0

    iarray=[]
    jarray=[]
    bckt=[]
    
    for line in bpa_str_ar:		#This loop is used to fill the branch data array with each line of branch data
        #line = line.lstrip() 
        line = line.rstrip('\n')
        if line == "": continue		#Continue if the line is blank
        if (line[0] == 'L' and line[1] == ' ') or (line[0] == 'E' and line[1] == ' '): line_str.append(line)

    for line in line_str:
        line = line + ' '*(80-len(line))	#To pad each line with spaces up to 80 records

        #--------------#
        #Default values#
        #--------------#
        intgar = [1, 0, 1, 0, 0, 0]
        realar = [0.0, 1e-6, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 1.0, 1.0, 1.0, 1.0]

        if line[2] == ' ':			#The modification code for a new record

            if line[3:6] != "   " and not owner_str.has_key(line[3:6].strip()):		#If the owner is not in the list, add it to the list with a new number
                owner_nbr = owner_nbr + 1
                owner_str[line[3:6].strip()] = owner_nbr	#A new dictionnary entry
                intgar[2] = owner_nbr				#The new owner number
            elif line[3:6] != "   " and owner_str.has_key(line[3:6].strip()):
                intgar[2] =  owner_str[line[3:6].strip()]	#Get the owner number
            elif line[3:6] == "   ":
                intgar[2] = 1					#The default owner

            if bus_str.has_key(line[6:14].strip() + str(float(line[14:18]))) and bus_str.has_key(line[19:27].strip() + str(float(line[27:31]))):
               iarg = bus_str[line[6:14].strip() + str(float(line[14:18]))]		#From bus
               jarg = bus_str[line[19:27].strip() + str(float(line[27:31]))]		#To bus
               iarray.append(iarg)
               jarray.append(jarg)
            else:
               psspy.progress_output(1,"",[0,0])           # 显示
               psspy.progress("Line termianl bus "+line[6:14]+' or '+line[19:27].strip()+" is not well defined\n\n")
               psspy.progress_output(6,"",[0,0])           # 关闭显示
               continue

            if line[18] == '1': intgar[1] = iarg	#The metered end is IARG
            elif line[18] == '2': intgar[1] = jarg	#The metered end is JARG
            elif line[18] == ' ':			#If the metered end is left blank, BPA chooses
                if bus_owner_nbr[iarg] == bus_owner_nbr[jarg]: intgar[1] = iarg		#If both ends have the same owner, IARG is the metered end
                elif bus_owner_nbr[iarg] == intgar[2]: intgar[1] = jarg			#Else the end that has a different owner from the line is the metered end
                elif bus_owner_nbr[jarg] == intgar[2]: intgar[1] = iarg
                else: intgar[1] = iarg #存在线路两端的所有者均与节点所有者不同的情况，bpa没有说明，手动设置为节点1

            if isfloat(line[33:37]) and intgar[1] == iarg and isfloat(line[14:18]): realar[3] = sqrt(3)*float(line[33:37])*float(line[14:18])/1000.0	#The Rate A in MVA
            elif isfloat(line[33:37]) and intgar[1] == jarg and isfloat(line[27:31]): realar[3] = sqrt(3)*float(line[33:37])*float(line[27:31])/1000.0	#The Rate A in MVA
            else: realar[3] =	base_mva   # 用系统基准容量作为变压器A容量

            if line[31] != ' ': ckt = line[31]		#The circuit identifier
            else: ckt = '1'				#If none is mentionned, the default is 1
            bckt.append(ckt)

            if isfloat(line[38:44]): realar[0] = getfloatvalue(line[38:44], 5)			# Nominal branch resistance #  F6.5
#            if isfloat(line[44:50]) and float(line[44:50]) != 0.0: realar[1] = getfloatvalue(line[44:50], 5)	#Nominal branch reactance
            if isfloat(line[44:50]):
            	 if abs(float(line[44:50]))>=1e-5:  realar[1] = getfloatvalue(line[44:50], 5)	# Nominal branch reactance  F6.5

            if line[0] == 'L':		#If it is an L card
                if isfloat(line[50:56]):
                    realar[6] = getfloatvalue(line[50:56], 5)				#Real line shunt at bus IARG end  F6.5
                    realar[8] = realar[6]						#Real line shunt at bus JARG end
                if isfloat(line[56:62]): realar[2] = getfloatvalue(line[56:62], 5)*2	#Total line charging    F6.5
                if isfloat(line[62:66]): 
                	 realar[10] = getfloatvalue(line[62:66],1)		              #Line's length in miles  F4.1需修改
                else: 
                	 realar[10] = 1.0		                              #默认1km

            else:			#If it is an E card
                if isfloat(line[50:56]): realar[6] = getfloatvalue(line[50:56], 5)		#Real line shunt at bus IARG end
                if isfloat(line[56:62]): realar[7] = getfloatvalue(line[56:62], 5)		#Reactive line shunt at bus IARG end
                if isfloat(line[62:68]): realar[8] = getfloatvalue(line[62:68], 5)		#Real line shunt at bus JARG end
                if isfloat(line[68:74]): realar[9] = getfloatvalue(line[68:74], 5)		#Reactive line shunt at bus JARG end

            #It seems the format for the date in the file 2006eh-tmp-1.dat is not the same as in the user manual
            #if line[77:80] != "   ": intgar[0] = 0		#If there is an out of service date, the status is offline

            if line[32] == ' ' or line[32] == '0':              
                ierr = psspy.branch_data(iarg, jarg, ckt, intgar, realar)	#The API to load branch data in PSS/E
                line_counter += 1
                
                #if line[0]=='E': 
                        
                   #psspy.progress_output(1,"",[0,0])           # 显示
                   #psspy.progress("Line termianl bus "+line[6:14]+' or '+line[19:27].strip()+" is not well defined\n\n")
                   #
                   #print intgar, realar
                   #psspy.progress_output(6,"",[0,0])           # 关闭显示
                   #raise(line+ '*'+str(iarg)+' '+str(jarg))
 
            elif line[32] == '1':	#If the branch is the section number 1
                multi_intgar = [intgar[1], 0, 0, 0, 0, 0, 0, 0, 0, 0]
                section_nbr = 2		#To search for a section number 2
                to_bus, dummy_nbr, bus_nbr, bus_str, bus_owner_nbr, owner_nbr, multi_intgar = GetSectionLine(bpa_str_ar, iarg, jarg, section_nbr, dummy_nbr, bus_nbr, bus_str, bus_owner_nbr, owner_nbr, multi_intgar, line[31], base_mva)

                intgar[1] = iarg	#The metered end is IARG
                ierr = psspy.branch_data(iarg, to_bus, ckt, intgar, realar)	#The API to load branch data in PSS/E
                if multi_intgar[1] != 0: ierr = psspy.multi_section_line_data(iarg, jarg, "&" + ckt, multi_intgar)	#The API to load the multi-section line data in PSS/E	

    return owner_str, owner_nbr, dummy_nbr, bus_nbr, bus_str, bus_owner_nbr, line_counter, iarray, jarray, bckt

# ----------------------------------------------------------------------------------------------------
def GetLPlusCard(bpa_str_ar, owner_str, owner_nbr, bus_str, bus_nbr, bus_owner_nbr, base_mva, DEBUG):

    # bpa_str_ar: bpa文件
    # owner_str: 区域字典
    # owner_nbr: 区域编号
    # bus_str: 节点字典
    # bus_nbr: 节点编号
    # bus_owner_nbr: 根据节点编号查区域编号
    # base_mva: 基准容量

    line_str = []
    line_counter = 0
    for line in bpa_str_ar:		#This loop is used to fill the branch data array with each line of branch data
        if line[0:2] == 'L+':
            line_str.append(line)

    for line in line_str:
        line = line + ' '*(80-len(line))	#To pad each line with spaces up to 80 records

        if line[2] == ' ':			#The modification code for a new record
            iarg = bus_str[line[6:14].strip() + str(float(line[14:18]))]		#From bus
            jarg = bus_str[line[19:27].strip() + str(float(line[27:31]))]		#To bus
            if line[31] != ' ': ckt = line[31]		#The circuit identifier
            else: ckt = '1'				#If none is mentionned, the default is 1

            if isfloat(line[33:38]):
                
                ierr, ishnt = psspy.brndt2(iarg, jarg, ckt, 'ISHNT')
                ig = ishnt.imag - float(line[33:38])/base_mva
                ierr = psspy.branch_data(iarg, jarg, ckt, [_i,_i,_i,_i,_i,_i], [_f,_f,_f,_f,_f,_f,_f, ig,_f,_f,_f,_f,_f,_f,_f])
                
                #ierr, shunt_q = psspy.busdt1(iarg, 'YS', 'NOM')
                #shunt_q += float(line[33:38])
                #psspy.shunt_data(iarg,r"""1""",_i,[_f,-shunt_q])

            if isfloat(line[43:48]):

                ierr, jshnt = psspy.brndt2(iarg, jarg, ckt, 'JSHNT')
                jg = jshnt.imag - float(line[43:48])/base_mva
                ierr = psspy.branch_data(iarg, jarg, ckt, [_i,_i,_i,_i,_i,_i], [_f,_f,_f,_f,_f,_f,_f,_f,_f, jg,_f,_f,_f,_f,_f])
                
                #ierr, shunt_q = psspy.busdt1(jarg, 'YS', 'NOM')
                #shunt_q += float(line[43:48])
                #psspy.shunt_data(jarg,r"""1""",_i,[_f,-shunt_q])

# ----------------------------------------------------------------------------------------------------
def GetTCard(bpa_str_ar, owner_str, owner_nbr, bus_str, bus_nbr, bus_owner_nbr, dummy_nbr, base_mva, DEBUG):
    """Gets the two-winding transformer data from BPA."""

    transformer_str = []	#An array to contain all lines with two-winding transformers connecting AC buses
    trans_counter = 0
    for line in bpa_str_ar:		#This loop is used to fill the transformer data array with each line of transformer data
        #line = line.lstrip()           # BPA中，开始为空格也表示此行数据为注释          # BPA中，开始为空格也表示此行数据为注释
        line = line.rstrip('\n')
        if line == "": continue		#Continue if the line is blank
        if line[0] == 'T' and line[1] == ' ' or line[0] == 'T' and line[1] == 'P': transformer_str.append(line)

    for line in transformer_str:
        line = line + ' '*(80-len(line))	#To pad each line with spaces up to 80 records

        #--------------#
        #Default values#
        #--------------#
        intgar = [1, 0, 1, 0, 0, 0, 33, 0, 0, 0, 0, 0, 2, 1, 1]
        realari = [0.0, 0.0, basemva, 0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 0.0, 1.0, 1.0, 1.0, 1.0, 0.0, 0.0, 1.1, 0.9, 1.1, 0.9, 0.0, 0.0]

        if line[2] == ' ' and bus_str.has_key(line[6:14].strip() + str(float(line[14:18]))) and bus_str.has_key(line[19:27].strip() + str(float(line[27:31]))):

            if line[3:6] != "   " and not owner_str.has_key(line[3:6].strip()):		#If the owner is not in the list, add it to the list with a new number
                owner_nbr = owner_nbr + 1
                owner_str[line[3:6].strip()] = owner_nbr	#A new dictionnary entry
                intgar[2] = owner_nbr				#The new owner number
            elif line[3:6] != "   " and owner_str.has_key(line[3:6].strip()):
                intgar[2] =  owner_str[line[3:6].strip()]	#Get the owner number
            elif line[3:6] == "   ":
                intgar[2] = 1					#The default owner

            iarg = bus_str[line[6:14].strip() + str(float(line[14:18]))]		#From bus
            jarg = bus_str[line[19:27].strip() + str(float(line[27:31]))]		#To bus
            intgar[8] = iarg						#The winding one side

#            if iarg==jarg:
#               psspy.progress_output(1,"",[0,0])           # 显示
#               psspy.progress("Trans error"+line+"\n\n")
#               psspy.progress(str(iarg)+'*'+str(jarg)+'*'+line[6:14].strip() + str(float(line[14:18])) +'***'+ line[19:27].strip()+ str(float(line[27:31]))+"\n")
#               psspy.progress("\n"+line[6:14]+'*****'+line[19:27]+"\n")              
#               psspy.progress_output(6,"",[0,0])           # 关闭显示

            if isfloat(line[14:18]): realari[4] = float(line[14:18])	#Winding one nominal voltage

            if line[18] == '1': intgar[1] = iarg	#The metered end is IARG
            elif line[18] == '2': intgar[1] = jarg	#The metered end is JARG
            else: #if line[18] == ' ':			#If the metered end is left blank, BPA chooses
                if bus_owner_nbr[iarg] == bus_owner_nbr[jarg]: intgar[1] = iarg		#If both ends have the same owner, IARG is the metered end
                elif bus_owner_nbr[iarg] == intgar[2]: intgar[1] = jarg			#Else the end that has a different owner from the line is the metered end
                elif bus_owner_nbr[jarg] == intgar[2]: intgar[1] = iarg

            if isfloat(line[27:31]): realari[7] = float(line[27:31])	#Winding two nominal voltage

            if isfloat(line[33:37]): 
            	 if  float(line[33:37])>1e-6:   # 如果是零
            	     realari[8] = float(line[33:37])	#The Rate A in MVA
            	 else:    realari[8] = basemva
            else: 	  	realari[8] = basemva 	
            	 
            if line[31] != ' ': ckt = line[31]		#The circuit identifier
            else: ckt = '1'				#If none is mentionned, the default is 1

            if isfloat(line[38:44]): realari[0] = getfloatvalue(line[38:44], 5)			#Nominal transformer resistance
            	
            if isfloat(line[44:50]):
                 if float(line[44:50])!= 0.0:
                     realari[1] = getfloatvalue(line[44:50], 5) # Nominal transformer reactance
                 else: realari[1] = 1e-5                        # 变压器阻抗标幺值默认1e-5的数值       
            else: realari[1] = 1e-5                        # 变压器阻抗标幺值默认1e-5的数值  
            	    
            	       	
            if isfloat(line[50:56]): realari[15] = getfloatvalue(line[50:56], 5)		#The magnetizing conductance

            if isfloat(line[56:62]): realari[16] = getfloatvalue(line[56:62], 5)		#The magnetizing susceptance

            if line[1] == ' ':
                if isfloat(line[62:67]):
                    realari[3] = getfloatvalue(line[62:67], 2)	                #The winding one ratio/voltage
                    if abs(realari[3]-realari[4])/realari[4]>0.6: 
                    	 realari[3]=realari[4]   # 如果分接头差别太大，则认为是错误数据，然后填额定电压
                    	 raise('error1')                  	
                else: realari[3] = realari[4]
                if isfloat(line[67:72]):
                    realari[6] = getfloatvalue(line[67:72], 2)	                #The winding two ratio/voltage
                    if abs(realari[6]-realari[7])/realari[7]>0.6: 
                    	 realari[6]=realari[7]	 
                    	 raise('error1')
                else: realari[6] = realari[7]

            else:  # 是调相器
                if isfloat(line[62:67]): realari[5] = getfloatvalue(line[62:67],2)	#The winding one phase shift angle F5.2需修改
                realari[3] = realari[4]						#The winding one ratio/voltage
                realari[6] = realari[7]						#The winding two ratio/voltage

            #It seems the format for the date in the file 2006eh-tmp-1 is not the same as in the user manual
            #if line[77:80] != "   ": intgar[0] = 0		#If there is an out of service date, the status is offline

            intgar[11], intgar[9], realari[17], realari[18], intgar[6], realari[19], realari[20] = GetRCard(bpa_str_ar, line[6:14].strip() + str(float(line[14:18])), line[19:27].strip() + str(float(line[27:31])), bus_str)

            if line[32] == ' ':
                if DEBUG:
                    ierr, realaro  = psspy.two_winding_data(iarg, jarg, ckt, intgar, realari, "")	#The API to load transformer data in PSS/E
                trans_counter += 1

            elif line[32] == '1':	#If the transformer is section number 1
                multi_intgar = [intgar[1], 0, 0, 0, 0, 0, 0, 0, 0, 0]
                section_nbr = 2		#To search for section number 2
                to_bus, dummy_nbr, bus_nbr, bus_str, bus_owner_nbr, owner_nbr, multi_intgar = GetSectionLine(bpa_str_ar, iarg, jarg, section_nbr, dummy_nbr, bus_nbr, bus_str, bus_owner_nbr, owner_nbr, multi_intgar, line[31], base_mva)

                intgar[1] = iarg	#The metered end is IARG
                ierr, realaro  = psspy.two_winding_data(iarg, to_bus, ckt, intgar, realari, "")		#The API used to load two-winding data in PSS/E
                if multi_intgar[1] != 0: ierr = psspy.multi_section_line_data(iarg, jarg, "&" + ckt, multi_intgar)	#The API used to load multi-section line in PSS/E

    return owner_str, owner_nbr, bus_nbr, bus_str, bus_owner_nbr, trans_counter

# ----------------------------------------------------------------------------------------------------
def GetPlusCard(bpa_str_ar, load_id_ar, machine_id_ar, bus_str, owner_str, owner_nbr, base_mva, scale_str, bus_zone_str):
    """Gets the bus supplement data from BPA."""

    for line in bpa_str_ar:
        #line = line.lstrip() 
        line = line.rstrip('\n')
        if line == "": continue		#Continue if the line is blank

        #--------------#
        #Default values#
        #--------------#
        load_intgar = [1, 1, 1, 1]
        machine_intgar = [0, 0, 0, 0, 0]
        load_realar = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
        machine_realar = [0.0, 0.0, 9999.0, -9999.0, 9999.0, -9999.0, base_mva, 0.0, 1.0, 0.0, 0.0, 0.0, 1.0, 1.0, 1.0, 1.0]
        plant_realar = [1.0, 1.0]

        if line[0] == '+' and line[2] == ' ':	#If it is a + card
            line = line + ' '*(77-len(line))	#To pad each line with spaces up to 77 records

            if line[3:6] != "   " and not owner_str.has_key(line[3:6].strip()):	#If the owner is not in the list, add it to the list with a new number
                owner_nbr = owner_nbr + 1
                owner_str[line[3:6].strip()] = owner_nbr	#A new dictionnary entry
                load_intgar[3] = owner_nbr			#The new owner number
            elif line[3:6] != "   " and owner_str.has_key(line[3:6].strip()):
                load_intgar[3] =  owner_str[line[3:6].strip()]	#Get the owner number
            elif line[3:6] == "   ":
                load_intgar[3] = 1				#The default owner

            iarg = bus_str[line[6:14].strip() + str(float(line[14:18]))]	#The from bus number

            if line[18:20] == "*I" or line[18:20] == "01":			#If it is a constant current load
                if isfloat(line[20:25]): load_realar[2] = float(line[20:25])	#Constant current load P
                if isfloat(line[25:30]): load_realar[3] = float(line[25:30])	#Constant current load Q
            elif line[18:20] == "*P" or line[18:20] == "02":			#If it is a constant power load
                if isfloat(line[20:25]): load_realar[0] = float(line[20:25])	#Constant power load P
                if isfloat(line[25:30]): load_realar[1] = float(line[25:30])	#Constant power load Q

            if isfloat(line[30:34]): load_realar[4] = float(line[30:34])	#Constant impedance load P
            if isfloat(line[34:38]): load_realar[5] = float(line[34:38])	#Constant impedance load Q

            if isfloat(line[42:47]): machine_realar[0] = float(line[42:47])	#Generation P
            if isfloat(line[47:52]): machine_realar[1] = float(line[47:52])	#Generation Q

            if line[1] != 'A' and line[1] != 'F' and line[1] != 'I' and line[1] != 'P': type_code = 'N'	#Non-industrial load or generation
            elif line[1] == 'F' or line[1] == 'I' or line[1] == 'P': type_code = 'I'			#Industrial load
            elif line[1] == 'A' and line[18:20] == "*I" or line[18:20] == "01": type_code = 'X'		#If it is a +A card with constant current load
            elif line[1] == 'A' and line[18:20] == "*P" or line[18:20] == "02": type_code = 'Y'		#If is is a +A card with constant power load

            scale_ar = GetScaleData(scale_str, line[3:6].strip(), bus_zone_str[iarg], type_code)	#To get the scale factors

            load_realar[0] = load_realar[0] * scale_ar[0]		#To scale the constant power load P
            load_realar[1] = load_realar[1] * scale_ar[1]		#To scale the constant power load Q
            load_realar[2] = load_realar[2] * scale_ar[2]		#To scale the constant current load P
            load_realar[3] = load_realar[3] * scale_ar[3]		#To scale the constant current load Q
            load_realar[4] = load_realar[4] * scale_ar[4]		#To scale the constant impedance load P
            load_realar[5] = load_realar[5] * scale_ar[5]		#To scale the constant impedance load Q
            machine_realar[0] = machine_realar[0] * scale_ar[6]		#To scale the generation P
            machine_realar[1] = machine_realar[1] * scale_ar[7]		#To scale the generation Q

            if load_realar[0] != 0.0 or load_realar[1] != 0.0 or load_realar[2] != 0.0 or load_realar[3] != 0.0 or load_realar[4] != 0.0 or load_realar[5] != 0.0:
                load_id_ar[iarg] = load_id_ar[iarg] + 1		#To increment the load CKT
                ierr = psspy.load_data(iarg, str(load_id_ar[iarg]), load_intgar, load_realar)	#Loads the load data in PSS/E

            if machine_realar[0] != 0.0 or machine_realar[1] != 0.0 or machine_realar[4] != 0.0:	#If there is a generator
                machine_id_ar[iarg] = machine_id_ar[iarg] + 1		#To increment the machine CKT
                machine_intgar[0] = 1					#The machine status
                machine_intgar[1] = owner_nbr				#The owner number
                if machine_id_ar[iarg] == 1:				#If it is the first generator on this bus
                    ierr = psspy.plant_data(iarg, 0, plant_realar)	#First load the plant data
                ierr = psspy.machine_data(iarg, str(machine_id_ar[iarg]), machine_intgar, machine_realar)	#After add the generator data

    return owner_str, owner_nbr

# ----------------------------------------------------------------------------------------------------
def GetICard(bpa_str_ar, area_str, area_slack_nbr):
    """Gets the inter-area transfer data from BPA."""

    inter_nbr = 48			#In order to start to the character '1' in the transfer ID

    for line in bpa_str_ar:		#This loop is used to get the inter-area transfer data
        #line = line.lstrip() 
        line = line.rstrip('\n')
        if line == "": continue		#Continue if it is a blank line

        #--------------#
        #Default values#
        #--------------#
        ia = 1				#Default area number
        ja = 1				#Default area number
        realar = [0.0]			#Default transfer

        if line[0] == 'I' and line[2] == ' ':		#If it is an I card
            line = line + ' '*(34-len(line))		#To pad each line with spaces up to 34 records
            for (x,y), z in area_str.iteritems():
                if x == line[3:13].strip(): ia = z	#The 'from' area number
                if x == line[14:24].strip(): ja = z	#The 'to' area number

            if isfloat(line[26:34]): realar[0] = float(line[26:34])	#The amount of MW in the transfer

            inter_nbr = inter_nbr + 1	#To increment the transfer ID
            if inter_nbr == 58: inter_nbr = 65

            if inter_nbr == 123:	#After the 'Z' character, PSS/E do not accept the other ASCII characters
                psspy.alert(" WARNING: MAXIMUM NUMBER OF INTER-AREA TRANSFERS REACHED\n")
                return

            ierr = psspy.transfer_data(1, ia, ja, chr(inter_nbr), realar[0])	#The transfer data API

# ----------------------------------------------------------------------------------------------------
def GetRZCard(bpa_str_ar):
    """Gets the line series compensation from BPA."""

    for line in bpa_str_ar:
        #line = line.lstrip() 
        line = line.rstrip('\n')
        if line == "": continue		#Continue if the line is blank

        if line[0] == 'R' and line[1] == 'Z':	#If it is an RZ card
            line = line + ' '*(60-len(line))	#To pad each line with spaces up to 60 records

            #Found no equivalent way to model this series compensation device in PSS/E
            psspy.alert(" WARNING: THE RZ CARD GOING FROM BUS [" + line[6:18] + "] TO BUS [" + line[19:31] + "] CANNOT BE IMPORTED\n")

# ----------------------------------------------------------------------------------------------------
def GetZCard(bpa_str_ar, zone_str):
    """Gets the renamed zone names from BPA."""

    zone_str_temp = {}		#A dictionnary containing the modified zone names

    for line in bpa_str_ar:
        #line = line.lstrip() 
        line = line.rstrip('\n')
        if line == "": continue

        if line[0] == 'Z':
            line = line + ' '*(77-len(line))	#To pad each line with spaces up to 77 records

            for x, y in zone_str.iteritems():
                if x == line[3:5].strip(): zone_str_temp[line[5:7].strip()] = y
                elif x == line[8:10].strip(): zone_str_temp[line[10:12].strip()] = y
                elif x == line[13:15].strip(): zone_str_temp[line[15:17].strip()] = y
                elif x == line[18:20].strip(): zone_str_temp[line[20:22].strip()] = y
                elif x == line[23:25].strip(): zone_str_temp[line[25:27].strip()] = y
                elif x == line[28:30].strip(): zone_str_temp[line[30:32].strip()] = y
                elif x == line[33:35].strip(): zone_str_temp[line[35:37].strip()] = y
                elif x == line[38:40].strip(): zone_str_temp[line[40:42].strip()] = y
                elif x == line[43:45].strip(): zone_str_temp[line[45:47].strip()] = y
                elif x == line[48:50].strip(): zone_str_temp[line[50:52].strip()] = y
                elif x == line[53:55].strip(): zone_str_temp[line[55:57].strip()] = y
                elif x == line[58:60].strip(): zone_str_temp[line[60:62].strip()] = y
                elif x == line[63:65].strip(): zone_str_temp[line[65:67].strip()] = y
                elif x == line[68:70].strip(): zone_str_temp[line[70:72].strip()] = y
                elif x == line[73:75].strip(): zone_str_temp[line[75:77].strip()] = y
                else: zone_str_temp[x] = y

            zone_str.clear()			#To clear the dictionnary
            zone_str.update(zone_str_temp)	#To copy the contents of the temporary dictionnary

    return zone_str

# ----------------------------------------------------------------------------------------------------
def GetDZCard(bpa_str_ar, zone_str):
    """Gets the deleted zone names from BPA and remove connected buses and equipment."""

    for line in bpa_str_ar:
        #line = line.lstrip() 
        line = line.rstrip('\n')
        if line == "": continue		#Continue if the line is blank

        if line[0] == 'D' and line[1] == 'Z':	#If it is a DZ card
            line = line + ' '*(5-len(line))	#To pad each line with spaces up to 5 records
            if line[3:5].strip() in zone_str:
                ierr = psspy.bsys(0, 0, [0.0,0.0], 0, [], 0, [], 0, [], 1, [zone_str[line[3:5].strip()]])	#A bus subsystem of this zone
                ierr = psspy.extr(0, 0, [1, 0])						#To delete this subsystem
                ierr = psspy.bsys(0, 1, [0.0, 9999.], 0, [], 0, [], 0, [], 0, [])	#To bring back the original unfiltered view
                del zone_str[line[3:5].strip()]						#To delete the zone

            else: psspy.alert(" WARNING: THE ZONE [" +  line[3:5].strip() + "] SPECIFIED IN THE DZ CARD CANNOT BE FOUND\n")	#If the zone is not found

    return zone_str

# ----------------------------------------------------------------------------------------------------
def GetDCOwners(bpa_str_ar, from_bus, to_bus):
    """Gets the owners of the DC buses for the two-terminal DC lines in BPA."""

    owner_1 = ""	#The owner of bus 'converter 1'
    owner_2 = ""	#The owner of bus 'converter 2'

    for line in bpa_str_ar:
        #line = line.lstrip() 
        line = line.rstrip('\n')
        if line == "": continue		#Continue if the line is blank

        if line[0] == 'B' and line[1] == 'D' and line[2] == ' ':	#If it is a BD card
            if from_bus == (line[6:14].strip() + str(float(line[14:18]))): owner_1 = line[3:6].strip()
            elif to_bus == (line[6:14].strip() + str(float(line[14:18]))): owner_2 = line[3:6].strip()

    return owner_1, owner_2

# ----------------------------------------------------------------------------------------------------
def GetTwoTermConv(bpa_str_ar, from_bus, to_bus, dc_line_nbr, flow_flag, bus_str):
    """Gets the converter data for two-terminal lines in BPA."""

    #--------------#
    #Default values#
    #--------------#
    intgar_1 = [0, 0, 0, 0, 0]
    realari_1 = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
    intgar_2 = [0, 0, 0, 0, 0]
    realari_2 = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]

#    write_file = open('d:\\123.dat', 'w')
    
    for line in bpa_str_ar:
        #line = line.lstrip() 
        line = line.rstrip('\n')
        if line == "": continue			#Continue if the line is blank

        if line[0] == 'B' and line[1] == 'D' and line[2] == ' ':	#If it is a BD card
            line = line + ' '*(62-len(line))	#To pad each line with spaces up to 62 records
#            write_file.write(line+'\n')
            if from_bus == line[6:14].strip() + str(float(line[14:18])):	#If the bus is converter 1
                if flow_flag: cnvflg_1 = 1	#If the flow is positive, the converter 1 is a rectifier
                else: cnvflg_1 = 2		#If the flow is negative, the converter 1 is an inverter

                if bus_str.has_key(line[50:58].strip() + str(float(line[58:62]))): intgar_1[0] = bus_str[line[50:58].strip() + str(float(line[58:62]))]	#Converter 1 bus number

                if isint(line[23:25]): intgar_1[1] = int(line[23:25])		#Number of bridges in series
                if isfloat(line[30:35]): realari_1[0] = getfloatvalue(line[30:35],1)	#Minimum firing angle F5.1需修改
                 
                if isfloat(line[35:40]):   realari_1[1] = getfloatvalue(line[35:40],1)	#Maximum firing angle    F5.1需修改
#                if realari_1[1]>90: realari_1[1]=90                             #BPA给定的数据会超过90°，在PSSE里面分Converter等的区别，可能是输出bug 
                       

            elif to_bus == line[6:14].strip() + str(float(line[14:18])):	#If the bus is converter 2
                if flow_flag: cnvflg_2 = 2	#If the flow is positive, the converter 2 is an inverter
                else: cnvflg_2 = 1		#If the flow is negative, the converter 2 is a rectifier

                if bus_str.has_key(line[50:58].strip() + str(float(line[58:62]))): intgar_2[0] = bus_str[line[50:58].strip() + str(float(line[58:62]))]	#Converter 2 bus number

                if isint(line[23:25]): intgar_2[1] = int(line[23:25])		#Number of bridges in series
                if isfloat(line[30:35]): realari_2[0] = getfloatvalue(line[30:35],1)	#Minimum firing angle F5.1需修改
                
                if isfloat(line[35:40]): realari_2[1] = getfloatvalue(line[35:40],1)	#Maximum firing angle F5.1需修改
#                if realari_2[1]>180: realari_2[1]=90                           #PSSE输出的逆变侧的触发角不允许超过100°，可能是psse输出函数bug

        if line[0] == 'T' and line[1] == ' ' and line[2] == ' ':		#If it is a T card
            line = line + ' '*(80-len(line))	#To pad each line with spaces up to 80 records

            if from_bus == line[6:14].strip() + str(float(line[14:18])) or from_bus == line[19:27].strip() + str(float(line[27:31])):
                if isfloat(line[38:44]): realari_1[2] = getfloatvalue(line[38:44],5)	#Commutating resistance  F6.5需修改
                if isfloat(line[44:50]): realari_1[3] = getfloatvalue(line[44:50],5)	#Commutating reactance   F6.5需修改

                realari_1[4] = float(line[14:18])				#Primary base voltage PSSE中需要的是pu，BPA给定的是kV
#                write_file.write(str(realari_1[4])+line[14:18]+'\n')

                if isfloat(line[67:72]): realari_1[5] = getfloatvalue(line[67:72],2)	#Transformer ratio F5.2需修改
                if isfloat(line[62:67]): realari_1[6] = getfloatvalue(line[62:67],2)	#Tap setting      F5.2需修改     BPA给定的是kV

                realari_1[5] = realari_1[5]/realari_1[4]	#Transformer ratio
                realari_1[6] = realari_1[6]/realari_1[4]	#Tap setting          BPA给定的是kV

            elif to_bus == line[6:14].strip() + str(float(line[14:18])) or to_bus == line[19:27].strip() + str(float(line[27:31])):
                if isfloat(line[38:44]): realari_2[2] = float(line[38:44])	#Commutating resistance
                if isfloat(line[44:50]): realari_2[3] = float(line[44:50])	#Commutating reactance

                realari_2[4] = float(line[14:18])				#Primary base voltage PSSE中需要的是pu，BPA给定的是kV

                if isfloat(line[67:72]): realari_2[5] = getfloatvalue(line[67:72],2)	#Transformer ratio   F5.2需修改 节点2分接头位置
                if isfloat(line[62:67]): realari_2[6] = getfloatvalue(line[62:67],2)	#Tap setting       F5.2需修改   BPA给定的是kV，节点1分接头位置

                realari_2[5] = realari_2[5]/realari_2[4]
                realari_2[6] = realari_2[6]/realari_2[4]                        #转化为标幺值
                

        if line[0] == 'R' and line[1] == ' ' and line[2] == ' ':		#If it is a R card
            line = line + ' '*(67-len(line))	#To pad each line with spaces up to 67 records

            if from_bus == line[6:14].strip() + str(float(line[14:18])) or from_bus == line[19:27].strip() + str(float(line[27:31])):
                realari_1[4] = float(line[14:18])				#Primary base voltage PSSE中需要的是pu，BPA给定的是kV
                if isfloat(line[45:50]): realari_1[7] = float(line[45:50])	#Maximum tap setting   节点1分解头位置，kV 
                if isfloat(line[50:55]): realari_1[8] = float(line[50:55])	#Minimum tap setting   节点1分解头位置，kV 
                if isint(line[55:57]) and int(line[55:57]) != 0: realari_1[9] = (realari_1[7] - realari_1[8])/int(line[55:57])	#Tap step  每一个抽头的电压量

                realari_1[7] = realari_1[7]/realari_1[4]	
                realari_1[8] = realari_1[8]/realari_1[4]
                realari_1[9] = realari_1[9]/realari_1[4]

            elif to_bus == line[6:14].strip() + str(float(line[14:18])) or to_bus == line[19:27].strip() + str(float(line[27:31])):
                realari_2[4] = float(line[14:18])
                if isfloat(line[45:50]): realari_2[7] = getfloatvalue(line[45:50],2)	#Maximum tap setting F5.2需修改
                if isfloat(line[50:55]): realari_2[8] = getfloatvalue(line[50:55],2)	#Minimum tap setting F5.2需修改
                if isint(line[55:57]) and int(line[55:57]) != 0: realari_2[9] = (realari_2[7] - realari_2[8])/int(line[55:57])	#Tap step

                realari_2[7] = realari_2[7]/realari_2[4]
                realari_2[8] = realari_2[8]/realari_2[4]                        #转化为标幺值
                realari_2[9] = realari_1[9]/realari_2[4]
                

    ierr,realaro = psspy.two_term_dc_convr_data(cnvflg_1, dc_line_nbr, intgar_1, realari_1, "")	#To load the converter 1 in PSS/E
    ierr,realaro = psspy.two_term_dc_convr_data(cnvflg_2, dc_line_nbr, intgar_2, realari_2, "")	#To load the converter 2 in PSS/E

#    write_file.write(str(dc_line_nbr)+' '+str(cnvflg_1)+' '+str(intgar_1)+' '+str(realari_1)+'\n')
#    write_file.write(str(dc_line_nbr)+' '+str(cnvflg_2)+' '+str(intgar_2)+' '+str(realari_2)+'\n')
  
#    write_file.close()
# ----------------------------------------------------------------------------------------------------
def GetTwoTermLine(bpa_str_ar, bus_str):
    """Gets the two-terminal DC line data from BPA."""

    dc_line_nbr = 0	#The two-terminal DC line number

    for line in bpa_str_ar:
        #line = line.lstrip() 
        line = line.rstrip('\n')
        print line
        
        if line == "": continue		#Continue if the line is blank

        if line[0] == 'L' and line[1] == 'D' and line[2] == ' ':	#If it is a LD card
            line = line + ' '*(78-len(line))	#To pad each line with spaces up to 78 records

            #--------------#
            #Default values#
            #--------------#
            intgar = [1, 0]
            realari = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]

            dc_line_nbr = dc_line_nbr + 1	#To increment the DC line number

            if isfloat(line[56:61]): realari[0] = float(line[56:61])	#Scheduled power demand
            if isfloat(line[61:66]): realari[1] = float(line[61:66])	#Scheduled DC voltage
            if isfloat(line[37:43]): realari[4] = float(line[37:43])	#DC line resistance

            if line[55] == 'I': realari[5] = 0			#Compounding resistance = 0 if Vdc is controlled at the inverter
            elif line[55] == 'R': realari[5] = realari[4]	#Compounding resistance = Rdc if Vdc is controlled at the rectifier

            if line[18] == '1':				#If the metered side is converter 1
                if realari[0] >= 0.0: metrar = 'R'	#The metered side is the rectifier if the flow is positive
                else: metrar = 'I'			#The metered side is the inverter if the flow is negative

            elif line[18] == '2':			#If the metered side is converter 2
                if realari[0] >= 0.0: metrar = 'I'	#The metered side is the inverter if the flow is positive
                else: metrar = 'R'			#The metered side is the rectifier if the flow is negative

            elif line[18] == ' ':			#If the metered side is left blank, BPA chooses
                owner_1, owner_2 = GetDCOwners(bpa_str_ar, line[6:14].strip() + str(float(line[14:18])), line[19:27].strip() + str(float(line[27:31])))
                if owner_1 == owner_2:		#If both DC buses have the same owner
                    if realari[0] >= 0.0: metrar = 'R'	#The metered side is the rectifier if the flow is positive
                    else: metrar = 'I'			#The metered side is the inverter if the flow is negative
                elif owner_1 == line[3:6].strip():  #If the owner of converter 1 is the same as the line
                    if realari[0] >= 0.0: metrar = 'I'	#The metered side is the inverter if the flow is positive
                    else: metrar = 'R'			#The metered side is the rectifier if the flow is negative
                elif owner_2 == line[3:6].strip():  #If the owner of converter 2 is the same as the line
                    if realari[0] >= 0.0: metrar = 'R'	#The metered side is the rectifier if the flow is positive
                    else: metrar = 'I'			#The metered side is the inverter if the flow is negative

            ierr, realaro = psspy.two_terminal_dc_line_data(dc_line_nbr, intgar, realari, metrar)	#The API to load two-terminal DC line in PSS/E

            GetTwoTermConv(bpa_str_ar, line[6:14].strip() + str(float(line[14:18])), line[19:27].strip() + str(float(line[27:31])), dc_line_nbr, realari[0] >= 0.0, bus_str)	#To get the 2 converters data

# ----------------------------------------------------------------------------------------------------
def GetConvTrans(bpa_str_ar, conv_intgar, conv_realari, bus_name):
    """Gets the T and R card corresponding to the converter bus specified."""

    for line in bpa_str_ar:
        #line = line.lstrip() 
        line = line.rstrip('\n')
        if line == "": continue		#Continue if the line is blank

        if line[0] == 'T' and line[1] == ' ' and line[2] == ' ':	#If it is a T card
            line = line + ' '*(80-len(line))	#To pad each line with spaces up to 80 records

            if bus_name == line[6:14].strip() + str(float(line[14:18])) or bus_name == line[19:27].strip() + str(float(line[27:31])):
                if isfloat(line[67:72]): conv_realari[3] = float(line[67:72])	#Transformer ratio
                if isfloat(line[38:44]): conv_realari[4] = float(line[38:44])	#Commutating resistance
                if isfloat(line[44:50]): conv_realari[5] = float(line[44:50])	#Commutating reactance

                conv_realari[6] = float(line[14:18])				#Winding one base voltage

                if isfloat(line[62:67]): conv_realari[7] = float(line[62:67])	#Tap setting

        if line[0] == 'R' and line[1] == ' ' and line[2] == ' ':		#If it is a R card
            line = line + ' '*(67-len(line))	#To pad each line with spaces up to 67 records

            if bus_name == line[6:14].strip() + str(float(line[14:18])) or bus_name == line[19:27].strip() + str(float(line[27:31])):
                if isfloat(line[45:50]): conv_realari[8] = float(line[45:50])	#Maximum tap setting
                if isfloat(line[50:55]): conv_realari[9] = float(line[50:55])	#Minimum tap setting
                if isint(line[55:57]) and int(line[55:57]) != 0: conv_realari[10] = (conv_realari[8] - conv_realari[9])/int(line[55:57])	#Tap step

    return conv_intgar, conv_realari

# ----------------------------------------------------------------------------------------------------
def GetMultiTermConv(bpa_str_ar, bus_str, zone_str, zone_nbr, owner_str, owner_nbr, bus_nbr, area_str, area_slack_nbr, area_nbr):
    """Gets the converter data for multi-terminal lines and creates DC buses."""

    dc_bus_nbr = 0		#The DC bus number for multi-terminal lines
    dc_bus_str = {}		#The dictionnary to contain DC bus names and numbers
    dc_bus_owner = ["DEFAULT"]	#An array of the owners of each DC bus

    for line in bpa_str_ar:
        #line = line.lstrip() 
        line = line.rstrip('\n')
        if line == "": continue		#Continue if the line is blank
        print line

        if line[0] == 'B' and line[1] == 'M' and line[2] == ' ':	#If it is a BM card
            line = line + ' '*(80-len(line))	#To pad each line with spaces up to 80 records
            
            #--------------#
            #Default values#
            #--------------#
            conv_intgar = [0, 0, 0, 0]
            conv_realari = [0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 1.0, 1.5, 0.51, 0.00625, 0.0, 1.0]
            bus_intgar = [0, 0, 0, 0, 0]
            bus_realari = 0.0

            if not dc_bus_str.has_key(line[6:14].strip() + str(float(line[14:18]))):		#If the bus is not in the list, add it
                dc_bus_nbr = dc_bus_nbr + 1
                dc_bus_str[line[6:14].strip() + str(float(line[14:18]))] = dc_bus_nbr		#A new dictionnary entry
            else:
                dc_bus_nbr = dc_bus_str[line[6:14].strip() + str(float(line[14:18]))]		#Get the bus number

            if line[3:6] != "   " and not owner_str.has_key(line[3:6].strip()):	#If the owner is not in the list, add it to the list with a new number
                owner_nbr = owner_nbr + 1
                owner_str[line[3:6].strip()] = owner_nbr		#A new dictionnary entry
                bus_intgar[3] = owner_nbr				#The new owner number
            elif line[3:6] != "   " and owner_str.has_key(line[3:6].strip()):
                bus_intgar[3] =  owner_str[line[3:6].strip()]		#Get the owner number
            elif line[3:6] == "   ":
                bus_intgar[3] = 1					#The default owner

            dc_bus_owner.insert(dc_bus_nbr + 1, line[3:6].strip())	#Insert the owner number of this bus

            if line[18:20] != "  ":
                if not zone_str.has_key(line[18:20].strip()):		#If the zone is not in the list, add it
                    zone_nbr = zone_nbr + 1
                    zone_str[line[18:20].strip()] = zone_nbr		#A new dictionnary entry
                    bus_intgar[2] = zone_nbr				#The new zone number
                else:
                    bus_intgar[2] =  zone_str[line[18:20].strip()]	#Get the zone number

            area_flag, bus_flag, area_nbr, swing_bus_nbr, bus_str, area_str, area_slack_nbr = GetACard(bpa_str_ar, line[18:20].strip(), bus_str, bus_nbr, area_str, area_slack_nbr, area_nbr)

            if area_flag == True:		#To assign the area number to the DC buses
                bus_intgar[1] = area_nbr
            else:				#The default area number
                bus_intgar[1] = 1

            if line[62] != ' ':			#If it is not a passive DC bus
                conv_intgar[0] = bus_str[line[50:58].strip() + str(float(line[58:62]))]	#Converter bus number
                bus_intgar[0] = conv_intgar[0]						#Converter bus number

                if isint(line[23:25]): conv_intgar[1] = int(line[23:25])		#Number of bridges in series
                if isfloat(line[69:75]):
                    if float(line[69:75]) != 0.0: conv_realari[0] = float(line[69:75])	#Scheduled power
                    elif isfloat(line[75:80]): conv_realari[0] = float(line[75:80])	#Scheduled voltage

                if line[62] == 'R' and isfloat(line[30:35]): conv_realari[1] = float(line[30:35])	#Minimum firing angle

                elif line[62] == 'I' and isfloat(line[66:69]) or line[62] == 'M' and isfloat(line[66:69]):
                    conv_realari[1] = float(line[66:69])			#Minimum extinction angle

                if isfloat(line[35:40]): conv_realari[2] = float(line[35:40])	#Maximum firing angle

                conv_intgar, conv_realari = GetConvTrans(bpa_str_ar, conv_intgar, conv_realari, line[6:14].strip() + str(float(line[14:18])))	#To get the corresponding T and R parameters

                ierr,realaro = psspy.multi_term_dc_convr_data(1, conv_intgar, conv_realari)	#The API to load multi-terminal DC line converter data in PSS/E

            ierr,realaro = psspy.multi_term_dc_bus_data(1, dc_bus_nbr, bus_intgar, bus_realari, line[6:14].strip())	#The API to load multi-terminal DC line bus data in PSS/E

    return dc_bus_str, dc_bus_owner, owner_str, owner_nbr, zone_str, zone_nbr, area_str, area_slack_nbr, area_nbr

# ----------------------------------------------------------------------------------------------------
def GetMultiTermLink(bpa_str_ar, dc_bus_str, dc_bus_owner):
    """Gets the multi-terminal link data from the BPA file."""

    for line in bpa_str_ar:
        #line = line.lstrip() 
        line = line.rstrip('\n')
        if line == "": continue		#Continue if the line is blank

        if line[0] == 'L' and line[1] == 'M' and line[2] == ' ':	#If it is a LM card
            line = line + ' '*(80-len(line))	#To pad each line with spaces up to 80 records

            #--------------#
            #Default values#
            #--------------#
            intgar = 0
            realar = [0.0, 0.0]

            ibus = dc_bus_str[line[6:14].strip() + str(float(line[14:18]))]	#The from bus
            jbus = dc_bus_str[line[19:27].strip() + str(float(line[27:31]))]	#The to bus

            if line[18] == '1': intgar = ibus		#The metered end is IBUS
            elif line[18] == '2': intgar = jbus		#The metered end is JBUS
            elif line[18] == ' ':			#If the metered end is left blank, BPA chooses
                if dc_bus_owner[ibus] == dc_bus_owner[jbus]: intgar = ibus	#If both ends have the same owner, IBUS is the metered end
                elif dc_bus_owner[ibus] == line[3:6].strip(): intgar = jbus	#Else the end that has a different owner from the line is the metered end
                elif dc_bus_owner[jbus] == line[3:6].strip(): intgar = ibus
                else: intgar = ibus 

            if isfloat(line[37:43]): realar[0] = float(line[37:43])	#DC link resistance
            if isfloat(line[43:49]): realar[1] = float(line[43:49])	#DC link inductance

            ierr = psspy.multi_term_dc_link_data(1, ibus, jbus, '1', intgar, realar)	#The API used to load multi-terminal DC line link data in PSS/E

# ----------------------------------------------------------------------------------------------------
def GetMultiTermLine(bpa_str_ar, bus_str, zone_str, zone_nbr, owner_str, owner_nbr, area_str, area_slack_nbr, area_nbr):
    """Gets the multi-terminal lines from the BPA file."""

    for line in bpa_str_ar:
        #line = line.lstrip() 
        line = line.rstrip('\n')
        if line == "": continue		#Continue if the line is blank

        if line[0] == 'L' and line[1] == 'M' and line[2] == ' ':	#If it is a LM card
            line = line + ' '*(80-len(line))	#To pad each line with spaces up to 80 records

            #--------------#
            #Default values#
            #--------------#
            intgari = 1
            realar = 0.0

            ierr, intgaro = psspy.multi_term_dc_line_data(1, intgari, realar)	#The API used to load multi-terminal DC line data in PSS/E

            #To get the converter, DC bus and DC link data
            dc_bus_str, dc_bus_owner, owner_str, owner_nbr, zone_str, zone_nbr, area_str, area_slack_nbr, area_nbr = GetMultiTermConv(bpa_str_ar, bus_str, zone_str, zone_nbr, owner_str, owner_nbr, bus_nbr, area_str, area_slack_nbr, area_nbr)
            GetMultiTermLink(bpa_str_ar, dc_bus_str, dc_bus_owner)

            return zone_str, zone_nbr, owner_str, owner_nbr, area_str, area_slack_nbr, area_nbr

    return zone_str, zone_nbr, owner_str, owner_nbr, area_str, area_slack_nbr, area_nbr

# ----------------------------------------------------------------------------------------------------
def Revise_busname(bus_str):
    
    for bus_name in bus_str:
        if str(bus_name).find("'")>=0:
            iarg=bus_str[bus_name]
            bus_newname=(str(bus_name).split("'"))[0]+'_'
            ierr =psspy.bus_data_2(iarg,[_i,_i,_i,_i],[_f,_f,_f], bus_newname)
        elif str(bus_name).find('"')>=0:
            iarg=bus_str[bus_name]
            bus_newname=(str(bus_name).split('"'))[0]+'1_'
            ierr =psspy.bus_data_2(iarg,[_i,_i,_i,_i],[_f,_f,_f], bus_newname)
        
# ----------------------------------------------------------------------------------------------------
def GetLOCard(bpa_str_ar1, bus_str, bus_nbr, DEBUG):

    seq_line_str = []                            #An array to contain all Seq_branch line data
    for line in bpa_str_ar1:
        line = line.rstrip('\n')
        if line == "":continue
        if line[0] == 'L' and line[1] == 'O' and line[2] != '+' : seq_line_str.append(line)

    for line in seq_line_str:
        line = line + ' '*(80-len(line))         #To pad each line with spaces up to 80 records


        seq_line_realar = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
        if bus_str.has_key(line[4:12].strip() + str(float(line[12:16]))) and bus_str.has_key(line[18:26].strip() + str(float(line[26:30]))): #稳定中有些节点可能潮流中没有
            iarg = bus_str[line[4:12].strip() + str(float(line[12:16]))]		#From bus
            jarg = bus_str[line[18:26].strip() + str(float(line[26:30]))]		#To bus
        else:
            psspy.progress_output(1,"",[0,0])           # 显示
            psspy.progress("Line termianl bus "+line[4:12]+' or '+line[18:26].strip()+" is not well defined\n\n")
            psspy.progress_output(6,"",[0,0])           # 关闭显示
            #raise('aaaa')                               #若有节点不存在，终止
            continue

        if line[32] != ' ': ckt = line[32]		#The circuit identifier
        else: ckt = '1'				#If none is mentionned, the default is 1

        if isfloat(line[35:42]): seq_line_realar[0] = getfloatvalue(line[35:42],4)			# zero sequence branch resistance #  F7.4
        if isfloat(line[42:49]): seq_line_realar[1] = getfloatvalue(line[42:49],4)			# zero sequence branch reactance  #  F7.4
        #若仅一侧有并联电抗器，则为不对称支路，应填Ga0、Ba0、Gb0、Bb0。如无电抗器，或双侧有相同的并联电抗器，则为对称零序支路， Ga0= Gb0、Ba0= Bb0，此时
        #Ga0 、Ba0要填，Gb0、Bb0不填，或填Gb0=0、 Bb0=0,则重新自动取Gb0= Ga0、Bb0= Ba0。
        if isfloat(line[63:70]) or isfloat(line[70:77]):
            seq_line_realar[5] = getfloatvalue(line[63:70],4)                             #zero sequence active line shunt at bus J end   #  F7.4
            seq_line_realar[6] = getfloatvalue(line[70:77],4)                             #zero sequence reactive line shunt at bus J end #  F7.4
            if isfloat(line[49:56]): seq_line_realar[3] = getfloatvalue(line[49:56],4)                             #zero sequence active line shunt at bus I end   #  F7.4
            if isfloat(line[56:63]): seq_line_realar[4] = getfloatvalue(line[56:63],4)                             #zero sequence reactive line shunt at bus I end #  F7.4
        else:
            if isfloat(line[49:56]):
                seq_line_realar[5] = getfloatvalue(line[49:56],4)
                seq_line_realar[3] = seq_line_realar[5]
            if isfloat(line[56:63]):
                seq_line_realar[2] = getfloatvalue(line[56:63],4)*2



        ierr = psspy.seq_branch_data(iarg, jarg, ckt, seq_line_realar)                 #Loads the data in PSS/E

# ----------------------------------------------------------------------------------------------------
def GetLOZCard(bpa_str_ar1):

    seq_scale_str=[]

    for line in bpa_str_ar1:         #Loop over every line of the BPA file

        line = line.rstrip('\n')
        if line == "": continue      #To continue if it is a blank line

        if line[0] == 'L' and line[1] == 'O' and line[3] == 'Z':	#If the line is a LO-Z card
            line = line + ' '*(80-len(line))	#To pad each line with spaces up to 80 records
            scale_str.append(line)

    return seq_scale_str				#Returns an array of LO-Z cards


# ----------------------------------------------------------------------------------------------------
"""def GetSeqScaleData(seq_scale_str, iarg_zone_name, jarg_zone_name):

    

    for line in seq_scale_str:
        seq_scale_ar = [1.0, 1.0, 1.0, 1.0, 0.0, 0.0]
        branch_seq_mark = 0                                        #是否使用seq_scale_ar的标志                   
        
        if line[5:49] == "":                                #缺省
            if isfloat(line[50:55]):seq_scale_str[0] = seq_scale_str[0] * getfloatvalue(line[50:55],4)       #KG0，线路零序电导相对于正序电导的倍数
            if isfloat(line[55:60]):seq_scale_str[1] = seq_scale_str[1] * getfloatvalue(line[55:60],4)       #KB0，线路零序电纳相对于正序电纳的倍数
            if isfloat(line[60:65]):seq_scale_str[2] = seq_scale_str[2] * getfloatvalue(line[60:65],4)       #KGC0，线路对地支路零序电导相对于对地支路正序电导的倍数
            if isfloat(line[65:70]):seq_scale_str[3] = seq_scale_str[3] * getfloatvalue(line[65:70],4)       #KBC0，线路对地支路零序电纳相对于对地支路正序电纳的倍数
            if isfloat(line[70:75]):
                seq_scale_str[4] = seq_scale_str[4] * getfloatvalue(line[70:75],1)       #VMIN，电压等级限制最小值（kV）
            else:
                seq_scale_str[4] = 0
            if isfloat(line[75:80]):
                seq_scale_str[5] = seq_scale_str[5] * getfloatvalue(line[75:80],1)       #VMAX，电压等级限制最大值（kV）
            else:
                seq_scale_str[5] = 99999
            branch_seq_mark = 1

        elif (iarg_zone_name == line[5:7].strip() or iarg_zone_name ==line[8:10].strip() or iarg_zone_name ==line[11:13].strip()\
        or iarg_zone_name == line[14:16].strip() or iarg_zone_name == line[17:19].strip() or iarg_zone_name == line[20:22].strip()\
        or iarg_zone_name == line[23:25].strip() or iarg_zone_name == line[26:28].strip() or iarg_zone_name == line[29:31].strip()\
        or iarg_zone_name == line[32:34].strip() or iarg_zone_name == line[35:37].strip() or iarg_zone_name == line[38:40].strip()\
        or iarg_zone_name == line[41:43].strip() or iarg_zone_name == line[44:46].strip() or iarg_zone_name == line[47:49].strip())\
        and (jarg_zone_name == line[5:7].strip() or jarg_zone_name ==line[8:10].strip() or jarg_zone_name ==line[11:13].strip()\
        or jarg_zone_name == line[14:16].strip() or jarg_zone_name == line[17:19].strip() or jarg_zone_name == line[20:22].strip()\
        or jarg_zone_name == line[23:25].strip() or jarg_zone_name == line[26:28].strip() or jarg_zone_name == line[29:31].strip()\
        or jarg_zone_name == line[32:34].strip() or jarg_zone_name == line[35:37].strip() or jarg_zone_name == line[38:40].strip()\
        or jarg_zone_name == line[41:43].strip() or jarg_zone_name == line[44:46].strip() or jarg_zone_name == line[47:49].strip()):   #节点ij都在该卡中

            if isfloat(line[50:55]):seq_scale_str[0] = seq_scale_str[0] * getfloatvalue(line[50:55],4)       #KG0，线路零序电导相对于正序电导的倍数
            if isfloat(line[55:60]):seq_scale_str[1] = seq_scale_str[1] * getfloatvalue(line[55:60],4)       #KB0，线路零序电纳相对于正序电纳的倍数
            if isfloat(line[60:65]):seq_scale_str[2] = seq_scale_str[2] * getfloatvalue(line[60:65],4)       #KGC0，线路对地支路零序电导相对于对地支路正序电导的倍数
            if isfloat(line[65:70]):seq_scale_str[3] = seq_scale_str[3] * getfloatvalue(line[65:70],4)       #KBC0，线路对地支路零序电纳相对于对地支路正序电纳的倍数
            if isfloat(line[70:75]):
                seq_scale_str[4] = seq_scale_str[4] * getfloatvalue(line[70:75],1)       #VMIN，电压等级限制最小值（kV）
            else:
                seq_scale_str[4] = 0
            if isfloat(line[75:80]):
                seq_scale_str[5] = seq_scale_str[5] * getfloatvalue(line[75:80],1)       #VMAX，电压等级限制最大值（kV）
            else:
                seq_scale_str[5] = 99999
            branch_seq_mark = 1

    return seq_scale_ar, branch_seq_mark"""


# ----------------------------------------------------------------------------------------------------
def GetSeqLineData(bpa_str_ar1, bus_str, bus_nbr, iarray, jarray, bckt, seq_scale_str):

    branch_mark = 0  
    ierr, brnchs = psspy.abrncount(-1,1,3,2,1)       #获取线路总数
    
    for i in range(0,brnchs):

        iarg = iarray[i]
        jarg = jarray[i]
        ckt = bckt[i]

        ierr, base = psspy.busdat(iarg ,'BASE')                #电压等级
        ierr, iarg_zone_name = psspy.busint(iarg, 'ZONE')      #节点所在分区
        ierr, jarg_zone_name = psspy.busint(jarg, 'ZONE')
        
        
        ierr, charg = psspy.brndat(iarg,jarg,ckt,'CHARG')                #Total branch charging.
        ierr, chargz = psspy.brndat(iarg,jarg,ckt,'CHARGZ')              #Zero sequence charging.
        ierr, rx = psspy.brndt2(iarg,jarg,ckt,'RX')                       #Branch resistance and reactance in per unit.
        ierr, rxz = psspy.brndt2(iarg,jarg,ckt,'RXZ')                     #Zero sequence branch resistance and reactance.
        ierr, ishnt = psspy.brndt2(iarg,jarg,ckt,'ISHNT')                 #Admittance of line shunt at IBUS end of the branch.
        ierr, jshnt = psspy.brndt2(iarg,jarg,ckt,'JSHNT')                 #Admittance of line shunt at JBUS end of the branch.
        ierr, ishntz = psspy.brndt2(iarg,jarg,ckt,'ISHNTZ')               #Zero sequence admittance of line shunt at IBUS end of the branch.
        ierr, jshntz = psspy.brndt2(iarg,jarg,ckt,'JSHNTZ')               #Zero sequence admittance of line shunt at JBUS end of the branch.

        seq_line_realar = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
        
        if chargz == 0 and rxz.real == 0 and rxz.imag == 0 and ishntz.real == 0 and ishntz.imag == 0 and jshntz.real == 0 and jshntz.imag == 0:
            #判断零序参数都为0，即无LO卡使用默认值

#            seq_scale_ar, branch_seq_mark = GetSeqScaleData(seq_scale_str, iarg_zone_name, jarg_zone_name) #LO-Z卡
            for line in seq_scale_str:
                seq_scale_ar = [1.0, 1.0, 1.0, 1.0, 0.0, 0.0]
                branch_mark = 0                                        #是否使用seq_scale_ar的标志 
                if line[5:49] == "":                                #缺省
                    if isfloat(line[50:55]):seq_scale_str[0] = seq_scale_str[0] * getfloatvalue(line[50:55],4)       #KG0，线路零序电导相对于正序电导的倍数
                    if isfloat(line[55:60]):seq_scale_str[1] = seq_scale_str[1] * getfloatvalue(line[55:60],4)       #KB0，线路零序电纳相对于正序电纳的倍数
                    if isfloat(line[60:65]):seq_scale_str[2] = seq_scale_str[2] * getfloatvalue(line[60:65],4)       #KGC0，线路对地支路零序电导相对于对地支路正序电导的倍数
                    if isfloat(line[65:70]):seq_scale_str[3] = seq_scale_str[3] * getfloatvalue(line[65:70],4)       #KBC0，线路对地支路零序电纳相对于对地支路正序电纳的倍数
                    if isfloat(line[70:75]):
                        seq_scale_str[4] = seq_scale_str[4] * getfloatvalue(line[70:75],1)       #VMIN，电压等级限制最小值（kV）
                    else:
                        seq_scale_str[4] = 0
                    if isfloat(line[75:80]):
                        seq_scale_str[5] = seq_scale_str[5] * getfloatvalue(line[75:80],1)       #VMAX，电压等级限制最大值（kV）
                    else:
                        seq_scale_str[5] = 99999
                    branch_mark = 1
                elif (iarg_zone_name == line[5:7].strip() or iarg_zone_name ==line[8:10].strip() or iarg_zone_name ==line[11:13].strip()\
                or iarg_zone_name == line[14:16].strip() or iarg_zone_name == line[17:19].strip() or iarg_zone_name == line[20:22].strip()\
                or iarg_zone_name == line[23:25].strip() or iarg_zone_name == line[26:28].strip() or iarg_zone_name == line[29:31].strip()\
                or iarg_zone_name == line[32:34].strip() or iarg_zone_name == line[35:37].strip() or iarg_zone_name == line[38:40].strip()\
                or iarg_zone_name == line[41:43].strip() or iarg_zone_name == line[44:46].strip() or iarg_zone_name == line[47:49].strip())\
                and (jarg_zone_name == line[5:7].strip() or jarg_zone_name ==line[8:10].strip() or jarg_zone_name ==line[11:13].strip()\
                or jarg_zone_name == line[14:16].strip() or jarg_zone_name == line[17:19].strip() or jarg_zone_name == line[20:22].strip()\
                or jarg_zone_name == line[23:25].strip() or jarg_zone_name == line[26:28].strip() or jarg_zone_name == line[29:31].strip()\
                or jarg_zone_name == line[32:34].strip() or jarg_zone_name == line[35:37].strip() or jarg_zone_name == line[38:40].strip()\
                or jarg_zone_name == line[41:43].strip() or jarg_zone_name == line[44:46].strip() or jarg_zone_name == line[47:49].strip()):   #节点ij都在该卡中
                    if isfloat(line[50:55]):seq_scale_str[0] = seq_scale_str[0] * getfloatvalue(line[50:55],4)       #KG0，线路零序电导相对于正序电导的倍数
                    if isfloat(line[55:60]):seq_scale_str[1] = seq_scale_str[1] * getfloatvalue(line[55:60],4)       #KB0，线路零序电纳相对于正序电纳的倍数
                    if isfloat(line[60:65]):seq_scale_str[2] = seq_scale_str[2] * getfloatvalue(line[60:65],4)       #KGC0，线路对地支路零序电导相对于对地支路正序电导的倍数
                    if isfloat(line[65:70]):seq_scale_str[3] = seq_scale_str[3] * getfloatvalue(line[65:70],4)       #KBC0，线路对地支路零序电纳相对于对地支路正序电纳的倍数
                    if isfloat(line[70:75]):
                        seq_scale_str[4] = seq_scale_str[4] * getfloatvalue(line[70:75],1)       #VMIN，电压等级限制最小值（kV）
                    else:
                        seq_scale_str[4] = 0
                    if isfloat(line[75:80]):
                        seq_scale_str[5] = seq_scale_str[5] * getfloatvalue(line[75:80],1)       #VMAX，电压等级限制最大值（kV）
                    else:
                        seq_scale_str[5] = 99999
                    branch_mark = 1           

            if branch_mark == 1 and base >= seq_scale_ar[4] and base <= seq_scale_ar[5]:                #使用LO-Z卡

                seq_line_realar[0] = rxz.ral * seq_scale_ar[0]
                seq_line_realar[1] = rxz.imag * seq_scale_ar[1]
                seq_line_realar[3] = ishnt.real * seq_scale_ar[2]     
                seq_line_realar[4] = ishnt.imag * seq_scale_ar[3]     
                seq_line_realar[5] = jshnt.real * seq_scale_ar[2]    
                seq_line_realar[6] = jshnt.imag * seq_scale_ar[3]     
                ierr = psspy.seq_branch_data(iarg, jarg, ckt, seq_line_realar)                 #Loads the data in PSS/E
                

            elif charg == 0:              #E卡
                seq_line_realar[0] = rx.real * 0.8
                seq_line_realar[1] = rx.imag * 0.313
                seq_line_realar[3] = ishnt.real * 0.7       #zero sequence active line shunt at bus I end   #  F7.4
                seq_line_realar[4] = ishnt.imag * 0.68      #zero sequence reactive line shunt at bus I end #  F7.4
                seq_line_realar[5] = jshnt.real * 0.7         #zero sequence active line shunt at bus J end   #  F7.4
                seq_line_realar[6] = jshnt.imag * 0.68       #zero sequence reactive line shunt at bus J end #  F7.4
                ierr = psspy.seq_branch_data(iarg, jarg, ckt, seq_line_realar)                 #Loads the data in PSS/E
            else:                           #L卡
                seq_line_realar[0] = rxz.real * 0.8
                seq_line_realar[1] = rxz.imag * 0.313
                seq_line_realar[2] = charg * 0.68
                seq_line_realar[3] = ishnt.real * 0.7
                seq_line_realar[5] = jshnt.real * 0.7 
                ierr = psspy.seq_branch_data(iarg, jarg, ckt, seq_line_realar)                 #Loads the data in PSS/E
    

# ----------------------------------------------------------------------------------------------------
def GetLOPlusCard(bpa_str_ar1, bus_str, bus_nbr, DEBUG):
    
    seq_line_str = []                            #An array to contain all Seq_branch line data
    for line in bpa_str_ar1:
        line = line.rstrip('\n')
        if line == "":continue
        if line[0:3] == 'LO+': seq_line_str.append(line)

    for line in seq_line_str:
        line = line + ' '*(80-len(line))         #To pad each line with spaces up to 80 records

        if bus_str.has_key(line[4:12].strip() + str(float(line[12:16]))) and bus_str.has_key(line[18:26].strip() + str(float(line[26:30]))): #稳定中有些节点可能潮流中没有
            ibus = bus_str[line[4:12].strip() + str(float(line[12:16]))]		#From bus
            jbus = bus_str[line[18:26].strip() + str(float(line[26:30]))]		#To bus

            if line[32] != ' ': ckt = line[32]		#The circuit identifier
            else: ckt = '1'				#If none is mentionned, the default is 1

            ierr, ishntz = psspy.brndt2(ibus, jbus, ckt,'ISHNTZ')
            ierr, jshntz = psspy.brndt2(ibus, jbus, ckt,'JSHNTZ')
 
            if isfloat(line[35:42]):
                ishntz = float(line[35:42]) + ishntz.imag            #XL1，线路前侧高抗对应的等值零序电抗（pu） F7.0
            if isfloat(line[42:49]):
                jshntz = float(line[42:49]) + jshntz.imag            #XL2，线路后侧高抗对应的等值零序电抗（pu） F7.0

            ierr = psspy.seq_branch_data(ibus, jbus, ckt, [_f,_f,_f,_f,ishntz,_f,jshntz])    

# ----------------------------------------------------------------------------------------------------
def GetLMCard(bpa_str_ar1, bus_str):
    seq_line_str = []                            #An array to contain all Seq_branch line data
    for line in bpa_str_ar1:
        line = line.rstrip('\n')
        if line == "":continue
        if line[0:2] == 'LM': seq_line_str.append(line)

    for line in seq_line_str:
        line = line + ' '*(80-len(line))         #To pad each line with spaces up to 80 records

        realar = [0.0, 0.0]

        if bus_str.has_key(line[4:12].strip() + str(float(line[12:16]))) and bus_str.has_key(line[18:26].strip() + str(float(line[26:30])))\
        and bus_str.has_key(line[35:43].strip() + str(float(line[43:47]))) and bus_str.has_key(line[49:57].strip() + str(float(line[57:61]))):
            #稳定中有些节点可能潮流中没有
            ibus = bus_str[line[4:12].strip() + str(float(line[12:16]))]		#ibus
            jbus = bus_str[line[18:26].strip() + str(float(line[26:30]))]		#jbus
            kbus = bus_str[line[35:43].strip() + str(float(line[43:47]))]		#kbus
            lbus = bus_str[line[49:57].strip() + str(float(line[57:61]))]		#lbus

            if line[32] != ' ': ickt1 = line[32]		#The circuit identifier
            else: ickt1 = '1'				#If none is mentionned, the default is 1

            if line[32] != ' ': ickt2 = line[32]		#The circuit identifier
            else: ickt2 = '1'				#If none is mentionned, the default is 1

            if isfloat(line[66:73]): realar[0] = getfloatvalue(line[66:73],4)     #Rm，互感电阻（pu） F7.4
            if isfloat(line[73:80]): realar[1] = getfloatvalue(line[73:80],4)     #Xm，互感电阻（pu） F7.4

            ierr = psspy.seq_mutual_data(ibus, jbus, ickt1, kbus, lbus, ickt2, realar)
        

# ----------------------------------------------------------------------------------------------------
def GetXRCard(bpa_str_ar1, bus_str):

    seq_line_str = []                            #An array to contain all Seq_branch line data
    for line in bpa_str_ar1:
        line = line.rstrip('\n')
        if line == "":continue
        if line[0:2] == 'XR': seq_line_str.append(line)

    for line in seq_line_str:
        line = line + ' '*(80-len(line))         #To pad each line with spaces up to 80 records

        realar = [0.0, 0.0]

        if bus_str.has_key(line[4:12].strip() + str(float(line[12:16]))):   #稳定中有些节点可能潮流中没有
            ibus = bus_str[line[4:12].strip() + str(float(line[12:16]))]

            if isfloat(line[21:28]):
                realar[0] = getfloatvalue(line[21:28],4)   #R0，对地支路零序电阻（pu） F7.4
        
            if isfloat(line[28:35]):
                realar[1] = getfloatvalue(line[28:35],4)   #X0，对地支路零序电抗（pu） F7.4
            fixid = '1'

            ierr = psspy.seq_fixed_shunt_data(ibus, fixid, realar)
# ----------------------------------------------------------------------------------------------------
def GerXOCard(bpa_str_ar1, bus_str):

    seq_line_str = []                            #An array to contain all Seq_branch line data
    for line in bpa_str_ar1:
        line = line.rstrip('\n')
        if line == "":continue
        if line[0:2] == 'XO': seq_line_str.append(line)

    for line in seq_line_str:
        line = line + ' '*(80-len(line))         #To pad each line with spaces up to 80 records

        realar = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0]

        if bus_str.has_key(line[4:12].strip() + str(float(line[12:16]))) and bus_str.has_key(line[18:26].strip() + str(float(line[26:30]))):#稳定中有些节点可能潮流中没有
            ibus = bus_str[line[4:12].strip() + str(float(line[12:16]))]
            jbus = bus_str[line[18:26].strip() + str(float(line[26:30]))]   

            if line[33] != ' ': ickt = line[33]		#The circuit identifier
            else: ickt = '1'				#If none is mentionned, the default is 1

            if line[31] == 1:                                  #SID=1 变压器零序电抗接于A侧
                if isfloat(line[21:28]):
                    realar[0] = getfloatvalue(line[37:44],4)   #X0，变压器零序电抗（pu） F7.4
                    realar[1] = getfloatvalue(line[44:51],4)   #R0，变压器零序电阻（pu） F7.4
            elif line[31] == 2:                                #SID=1 变压器零序电抗接于B侧
                if isfloat(line[21:28]):
                    realar[2] = getfloatvalue(line[37:44],4)   #X0，变压器零序电抗（pu） F7.4
                    realar[3] = getfloatvalue(line[44:51],4)   #R0，变压器零序电阻（pu） F7.4
            elif line[31] == 3:                                #SID=1 变压器零序电抗接于A, B之间
                if isfloat(line[21:28]):
                    realar[2] = getfloatvalue(line[37:44],4)   #X0，变压器零序电抗（pu） F7.4
                    realar[3] = getfloatvalue(line[44:51],4)   #R0，变压器零序电阻（pu） F7.4

            intgar = 8                                                                              #是否为8??
        
            ierr = psspy.seq_two_winding_data(ibus, jbus, ickt, intgar, realar)    

# ----------------------------------------------------------------------------------------------------
DEBUG = 1
if DEBUG:
    
    bpa_file = OpenFile() # 打开指定文件

if bpa_file:					#If the file opened successfully
    #psspy.progress_output(6,"",[0,0])          # 不显示
    #psspy.progress_output(1,"",[0,0])           # 显示
    if DEBUG:
        psspy.progress("\n\nFile opened successfully, starting the conversion:\n")

    bpa_str = bpa_file.read()			#The string containing the text file, to use the find() function
    bpa_file.seek(- bpa_file.tell(), 1)		#To position back at the beginning
    bpa_str_ar = bpa_file.readlines()		#The array that is containing all the lines of the BPA file

    basemva = GetMVA(bpa_file, bpa_str)		#To get the MVA base
    
    titl1, titl2 = GetTitles(bpa_file, bpa_str)	#To get the title data

    if DEBUG:
        psspy.progress("\n-Title 1: " + titl1)
        psspy.progress("\n-Title 2: " + titl2 + "\n")
        ierr = psspy.newcas(basemva, 'change', ' ')	#To create the new case
        psspy.progress("\n-Base MVA: " + str(basemva)) # 输出基准容量

#    psspy.progress_output(6,"",[0,0])          # 不显示
   
    scale_str = GetPCard(bpa_str_ar)		#To get the P cards for scaling

    area_str, area_nbr, area_slack_nbr, owner_str, owner_nbr, bus_owner_nbr, bus_str, bus_nbr, zone_str, \
                                            zone_nbr, load_id_ar, machine_id_ar, bus_zone_str = GetBCard(bpa_str_ar, basemva, scale_str, DEBUG) # 节点

#    owner_str, owner_nbr = GetPlusCard(bpa_str_ar, load_id_ar, machine_id_ar, bus_str, owner_str, owner_nbr, basemva, scale_str, bus_zone_str)  # 移到前面来，为了保证增加了的节点及时添加

    owner_str, owner_nbr, dummy_nbr, bus_nbr, bus_str, bus_owner_nbr, line_counter, iarray, jarray, bckt = GetLCard(bpa_str_ar, owner_str, owner_nbr, \
                                                            bus_str, bus_nbr, bus_owner_nbr, basemva, DEBUG)                                    # 线路

    GetLPlusCard(bpa_str_ar, owner_str, owner_nbr, bus_str, bus_nbr, bus_owner_nbr, basemva, DEBUG)            # 线路高抗参数

    owner_str, owner_nbr, bus_nbr, bus_str, bus_owner_nbr, trans_counter = GetTCard(bpa_str_ar, owner_str, owner_nbr, bus_str, bus_nbr, \
                                                                     bus_owner_nbr, dummy_nbr, basemva, DEBUG) # 变压器

    owner_str, owner_nbr = GetPlusCard(bpa_str_ar, load_id_ar, machine_id_ar, bus_str, owner_str, owner_nbr, basemva, scale_str, bus_zone_str)  

    GetTwoTermLine(bpa_str_ar, bus_str)     # 直流尚存在问题，先忽略
    zone_str, zone_nbr, owner_str, owner_nbr, area_str, area_slack_nbr, area_nbr = GetMultiTermLine(bpa_str_ar, bus_str, zone_str, zone_nbr, owner_str, owner_nbr, area_str, area_slack_nbr, area_nbr)
    zone_str = GetZCard(bpa_str_ar, zone_str)	#To get Z cards
    zone_str = GetDZCard(bpa_str_ar, zone_str)	#To get DZ cards

    for x, y in zone_str.iteritems():
        ierr = psspy.zone_data(y, x)	#The API used to load zone data in PSS/E

    for (x,y), z in area_str.iteritems():
        ierr = psspy.area_data(z, area_slack_nbr[z - 1], [y, 0.0], x)	#The API used to load zone data in PSS/E

    GetICard(bpa_str_ar, area_str, area_slack_nbr)	#To get the I cards

    for x, y in owner_str.iteritems():
        ierr = psspy.owner_data(y, x)	#The API used to load owner data in PSS/E

    GetRZCard(bpa_str_ar)	#To get RZ cards

    bpa_file1 = OpenFile()              # 打开指定文件(.swi)
    if bpa_file1:
    	ierr = psspy.newseq()           #生成序阻抗矩阵
    	if DEBUG:
            psspy.progress("\n\nFile opened successfully, starting the conversion:\n")
        bpa_str1 = bpa_file1.read()			#The string containing the text file, to use the find() function
        bpa_file1.seek(- bpa_file1.tell(), 1)		#To position back at the beginning
        bpa_str_ar1 = bpa_file1.readlines()		#The array that is containing all the lines of the BPA file

#        GetLOCard(bpa_str_ar1, bus_str, bus_nbr, DEBUG)           #To get LOcard
        

        seq_scale_str = GetLOZCard(bpa_str_ar1)

        GetSeqLineData(bpa_str_ar1, bus_str, bus_nbr, iarray, jarray, bckt, seq_scale_str)      #计算无LO卡时线路零序阻抗

        GetLOPlusCard(bpa_str_ar1, bus_str, bus_nbr, DEBUG)                                           #读入LO+卡

        GetLMCard(bpa_str_ar1, bus_str)

        GetXRCard(bpa_str_ar1, bus_str)

        GerXOCard(bpa_str_ar1, bus_str)
        

        bpa_file1.close()	#To close the BPA file

    psspy.progress_output(1,"",[0,0])           # 显示
    psspy.progress("\n\nConversion completed\n")
    Revise_busname(bus_str)     #修正母线名中的特殊字符'
    psspy.case_title_data('BPA2PSSE',"")

    bpa_file.close()	#To close the BPA file

    psspy.progress("bus num: " + str(bus_nbr) + "\n")
    psspy.progress("line num: " + str(line_counter) + "\n")
    psspy.progress("trans num: " + str(trans_counter) + "\n")

#读发电机及控制系统的数据
