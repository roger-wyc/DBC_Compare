import os
import sys
import pprint
import time
import xlrd
import xlwt
import re
import copy
import tkinter
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import xlsxwriter


def getMessageFromDBC(line, pattern):
    tempMsgLine = re.findall(pattern, line)

    tempMessage = {}
    tempMessage['ID'] = hex(int(tempMsgLine[0][1]))
    tempMessage['Cycle_Time'] = 0
    tempMessage['Message_Name'] = tempMsgLine[0][2]
    tempMessage['Message_Length'] = int(tempMsgLine[0][3])
    tempMessage['Sender'] = tempMsgLine[0][4]
    tempMessage['Signals'] = {}

    return tempMessage


def getSignalFromDBC(line, pattern):
    tempSignalLine = re.findall(pattern, line)

    tempSignal = {}
    tempSignal['Signal_Name'] = tempSignalLine[0][1]
    tempSignal['Start-Bit'] = int(tempSignalLine[0][2])
    tempSignal['Length'] = int(tempSignalLine[0][3])
    tempSignal['Byte-Order'] = int(tempSignalLine[0][4])
    tempSignal['Value-Type'] = tempSignalLine[0][5]
    tempSignal['Factor'] = float(tempSignalLine[0][6])
    tempSignal['Offset'] = float(tempSignalLine[0][7])
    tempSignal['Minimum'] = float(tempSignalLine[0][8])
    tempSignal['Maximum'] = float(tempSignalLine[0][9])
    tempSignal['Unit'] = tempSignalLine[0][10]
    tempSignal['InvalidVlue'] = 'NA'
    tempSignal['InitVlue'] = 'NA'
    tempSignal['ValueTab'] = 'NA'

    return tempSignal


def updateMessageCycletimeFromDBC(canMessages, dbcfile, pattern):
    dbcfile.seek(0)
    for line in dbcfile:        
        if (re.search(pattern, line) != None):
            tempCycleLine = re.findall(pattern, line)
            tempMessageCycletime = {}
            tempMessageCycletime['ID'] = hex(int(tempCycleLine[0][3]))
            tempMessageCycletime['Cycle_Time'] = int(tempCycleLine[0][4])
            for msg in canMessages:
                if (canMessages[msg]['ID'] == tempMessageCycletime['ID']):
                    canMessages[msg]['Cycle_Time'] = tempMessageCycletime['Cycle_Time']   


def updateValueTabFromDBC(canMessages, dbcfile, pattern):
    dbcfile.seek(0)
    for line in dbcfile:        
        if (re.search(pattern, line) != None):
            tempDesc = re.findall(pattern, line)
            for message in canMessages:
                for signal in canMessages[message]['Signals']:
                    if canMessages[message]['Signals'][signal]['Signal_Name'] == tempDesc[0][2]:
                        canMessages[message]['Signals'][signal]['ValueTab'] = tempDesc[0][3]


def updateSignalInitialValueFromDBC(canMessages, dbcfile, pattern):
    dbcfile.seek(0)
    for line in dbcfile:  
        if re.search(pattern, line) != None:
            tempDesc = re.findall(pattern, line)
            for message in canMessages:
                for signal in canMessages[message]['Signals']:
                    if canMessages[message]['Signals'][signal]['Name'] == tempDesc[0][4]:
                        canMessages[message]['Signals'][signal]['InitVlue'] = tempDesc[0][3]


def updateSignalInvalidValueFromDBC(canMessages, dbcfile, pattern):
    dbcfile.seek(0)
    for line in dbcfile:
        if re.search(pattern, line) != None:
            tempDesc1 = re.findall(pattern, line)
            for message in canMessages:
                for signal in canMessages[message]['Signals']:
                    if canMessages[message]['Signals'][signal]['Name'] == tempDesc1[0][4]:
                        canMessages[message]['Signals'][signal]['InvalidVlue'] = tempDesc1[0][5]


def ParseDBC(DBCFileName):
    canMessages = {}
    cycletimePattern = r'(BA_)\s*("GenMsgCycleTime")\s*(BO_)\s*(\d*)\s*(\d*);\n'
    messagePattern = r'(BO_)\s*(\d*)\s*(\w*)\s*:\s*(\d*)\s*(\w*)\s*\n'
    signalPattern = r'(SG_)\s*(\w*\s*\w+)\s*:\s*(\d*)\s*\|\s*(\d*)\s*@\s*(\d*)\s*(\+|\-)\s*\(\s*(-?\d*\.*\d*)\s*,\s*(-?\d*\.*\d*)\s*\)\s*\[\s*(-?\d*\.*\d*)\s*\|\s*(-?\d*\.*\d*)\s*]\s*"(.*)"\s+([\w*,?]*)\s*\n+'
    # valueTabPattern = r'(VAL_)\s*\d*\s*([\w\d]+)\s*(.+);'
    valueTabPattern = r'(VAL_)\s*(\d*)\s*(\w*)\s*(.*)\s*;\n'     
    invalidVluePattern = r'(BA_)\s*("GenSigInvalidValue")\s*(SG_)\s*(\d*)\s*(\w*)\s*"(.*)";\n'
    initVluePattern = r'(BA_)\s*("GenSigStartValue")\s*(SG_)\s*(\d*)\s*(\w*)\s*(.*);\n'

    
    DBCFile = open(DBCFileName, 'r')
    DBCFile.seek(0)
    CurrentMsgName = ''
    for line in DBCFile:
        if (re.search(messagePattern, line) != None):
            tempMessage = getMessageFromDBC(line, messagePattern)  
            CurrentMsgName = tempMessage['Message_Name']
            canMessages[CurrentMsgName] = tempMessage  # Add temporary message to CAN message list
        elif (re.search(signalPattern, line) != None):
            tempSignal = getSignalFromDBC(line, signalPattern)
            canMessages[CurrentMsgName]['Signals'][tempSignal['Signal_Name']] = tempSignal  # Add temporary signal to current message of CAN message list


    updateMessageCycletimeFromDBC(canMessages, DBCFile, cycletimePattern)
    updateValueTabFromDBC(canMessages, DBCFile,valueTabPattern)
    updateSignalInitialValueFromDBC(canMessages, DBCFile, initVluePattern)
    updateSignalInvalidValueFromDBC(canMessages, DBCFile, invalidVluePattern)

    DBCFile.close()

    return canMessages
  

 
def chooseinputfile1():
    dbcfilename = filedialog.askopenfilename(title='Open Excel File', filetypes=[('DBC', '*.dbc'), ('All Files', '*')])
    inputfilepath1.set(dbcfilename)

def chooseinputfile2():
    dbcfilename = filedialog.askopenfilename(title='Open Excel File', filetypes=[('DBC', '*.dbc'), ('All Files', '*')])
    inputfilepath2.set(dbcfilename)

def chooseoutputfile():
    xlsxfilename = filedialog.askdirectory()
    outputfilepath.set(xlsxfilename)


def checkpath():
    if not os.path.exists(inputfilepath1.get()):
        messagebox.showerror('Error','The Old DBC File is not existed!')
        return False
    if not os.path.exists(inputfilepath2.get()):
        messagebox.showerror('Error','The New DBC File is not existed!')
        return False
    if not os.path.isdir(outputfilepath.get()):
        messagebox.showerror('Error','The Output Path is not existed!')
        return False
    return True


def Cmp_Message(msg1, msg2):
    diff_attr = {}
    # for I in InfoName:
    #     if msg1[I] != msg2[I]:
    #         diff_attr.update({I:[msg1[I], msg2[I]]})

    if msg1['ID'] != msg2['ID']:
        diff_attr.update( {'ID':[ msg1['ID'], msg2['ID'] ] })
    if msg1['Cycle_Time'] != msg2['Cycle_Time']:
        diff_attr.update( {'Cycle_Time':[ msg1['Cycle_Time'], msg2['Cycle_Time'] ] })
    if msg1['Message_Name'] != msg2['Message_Name']:
        diff_attr.update( {'Message_Name':[ msg1['Message_Name'], msg2['Message_Name'] ] })
    if msg1['Message_Length'] != msg2['Message_Length']:
        diff_attr.update( {'Message_Length':[ msg1['Message_Length'], msg2['Message_Length'] ] })

    if diff_attr == {}:
        return True
    else:
        return diff_attr

def Cmp_Signal(sg1, sg2):
    diff_attr = {}
    # for I in InfoName:
    #     if sg1[I] != sg2[I]:
    #         diff_attr[I] = [ sg1[I], sg2[I] ]

    if sg1['Signal_Name'] != sg2['Signal_Name']:
        diff_attr.update( {'Signal_Name':[ sg1['Signal_Name'], sg2['Signal_Name'] ] })
    if sg1['Start-Bit'] != sg2['Start-Bit']:
        diff_attr.update( {'Start-Bit':[ sg1['Start-Bit'], sg2['Start-Bit'] ] })
    if sg1['Length'] != sg2['Length']:
        diff_attr.update( {'Length':[ sg1['Length'], sg2['Length'] ] })
    if sg1['Value-Type'] != sg2['Value-Type']:
        diff_attr.update( {'Value-Type':[ sg1['Value-Type'], sg2['Value-Type'] ] })
    if sg1['Factor'] != sg2['Factor']:
        diff_attr.update( {'Factor':[ sg1['Factor'], sg2['Factor'] ] })
    if sg1['Offset'] != sg2['Offset']:
        diff_attr.update( {'Offset':[ sg1['Offset'], sg2['Offset'] ] })
    if sg1['Minimum'] != sg2['Minimum']:
        diff_attr.update( {'Minimum':[ sg1['Minimum'], sg2['Minimum'] ] })
    if sg1['Maximum'] != sg2['Maximum']:
        diff_attr.update( {'Maximum':[ sg1['Maximum'], sg2['Maximum'] ] })

    if diff_attr == {}:
        return True
    else:
        return diff_attr


def Cmp_CMX(AllMsg_1, AllMsg_2):
    del_message = copy.deepcopy(AllMsg_1)
    add_message = copy.deepcopy(AllMsg_2)
    diff_msg = {}
    for msg2 in AllMsg_2:
        for msg1 in AllMsg_1:
            if msg2 == msg1:
                del del_message[msg1]
                del add_message[msg2]
                rst = Cmp_Message(AllMsg_1[msg1], AllMsg_2[msg2])
                diff_msg[msg1] = {}
                if rst != True:
                    diff_msg[msg1]['Diff_Msg'] = rst
                del_sg = copy.deepcopy(AllMsg_1[msg1]['Signals'])
                add_sg = copy.deepcopy(AllMsg_2[msg2]['Signals'])
                for sg1 in AllMsg_1[msg1]['Signals']:
                    for sg2 in AllMsg_2[msg2]['Signals']:
                        if sg1 == sg2:
                            rst_sg = Cmp_Signal(AllMsg_1[msg1]['Signals'][sg1], AllMsg_2[msg2]['Signals'][sg2])
                            if rst_sg != True:
                                if 'Diff_Signals' in diff_msg[msg1]:
                                    diff_msg[msg1]['Diff_Signals'][sg1] = rst_sg
                                else:
                                    diff_msg[msg1]['Diff_Signals'] = {}
                                    diff_msg[msg1]['Diff_Signals'][sg1] =  rst_sg
                            del del_sg[sg1]
                            del add_sg[sg2]
                if del_sg:
                    diff_msg[msg1]['del_signal'] = del_sg
                if add_sg:
                    diff_msg[msg2]['add_signal'] = add_sg

    _diff_msg = copy.deepcopy(diff_msg)
    for m in _diff_msg:
        if diff_msg[m] == {}:
            del diff_msg[m]
    return diff_msg, del_message, add_message


def GenCmpCANMtx(diff_msg, del_msg, add_msg, path):
    OutFile = xlwt.Workbook()

    # Record Deleted Messages
    OutSheet = OutFile.add_sheet('Deleted Messages')
    i = 1
    OutSheet.write(0, 1, 'Message Name')
    OutSheet.write(0, 0, 'Message ID')
    for m in del_msg:
        OutSheet.write(i, 1, str(del_msg[m]['Message_Name']))
        OutSheet.write(i, 0, str(del_msg[m]['ID']))
        i += 1

    # Record Added Messages
    OutSheet = OutFile.add_sheet('Added Messages')
    OutSheet.write(0, 0, 'Message ID')
    OutSheet.write(0, 1, 'Message Name')   
    i = 1

    for m in add_msg:
        OutSheet.write(i, 0, str(add_msg[m]['ID']))
        OutSheet.write(i, 1, str(add_msg[m]['Message_Name']))
        i += 1

    # Record Modified Messages
    OutSheet = OutFile.add_sheet('Modified Messages')
    OutSheet.write(0, 0, 'Message Name')
    OutSheet.write(0, 1, 'Message Differences')
    OutSheet.write(0, 2, 'Added Signals')
    OutSheet.write(0, 3, 'Deleted Signals')
    OutSheet.write(0, 4, 'Signal Differences')

    i = 1
    for m in diff_msg:
        start_i = i
        if 'Diff_Msg' in diff_msg[m]:
           s = ''
           j = 1
           for k in diff_msg[m]['Diff_Msg']:
               s += str(j)+'. '+ str(k)+': '+ str(diff_msg[m]['Diff_Msg'][k])+ ';  '
               j += 1
           OutSheet.write(i, 1, str(s))
           i += 1

        if 'add_signal' in diff_msg[m]:
           j = 1
           for k in diff_msg[m]['add_signal']:
               s = str(j)+'. '+str(k)+';  '
               j += 1
               OutSheet.write(i, 2, str(s))
               i += 1

        if 'del_signal' in diff_msg[m]:
           j = 1
           for k in diff_msg[m]['del_signal']:
               s = str(j)+'. '+str(k)+';  '
               j += 1
               OutSheet.write(i, 3, str(s))
               i += 1

        if 'Diff_Signals' in diff_msg[m]:
           j = 1
           for sg in diff_msg[m]['Diff_Signals']:
               s = str(j)+'. '+str(sg)+': '
               j += 1
               for a in diff_msg[m]['Diff_Signals'][sg]:
                   s += ' --'+str(a)+': '+str(diff_msg[m]['Diff_Signals'][sg][a])+';  '
               OutSheet.write(i, 4, str(s))
               i += 1
                  
        OutSheet.write_merge(start_i, i-1, 0, 0, str(m))

    OutFile.save(path)


def generateCT():
    if not checkpath(): 
        return
    
    try:
        CANList1 = ParseDBC(inputfilepath1.get())
        # pprint.pprint(CANList1)
        CANList2 = ParseDBC(inputfilepath2.get())
        # pprint.pprint(CANList2)

        diff_msg, del_msg, add_msg = Cmp_CMX(CANList1, CANList2)
        
        GenCmpCANMtx(diff_msg, del_msg, add_msg, os.path.join(outputfilepath.get(), outputfilename.get() + r'.xls'))

        messagebox.showinfo('Tip','Generate Successfully!')        
    except Exception as e:
        messagebox.showerror('Error','Generate Failed!\n' + 'Reason:'+ str(e))


#################################################################################################################

if __name__ == '__main__':
    window = tkinter.Tk()
    window.title('DBC Compare Tool')
    window.geometry('900x150')  # 设定窗口的大小(长 * 宽),这里的乘是小x
    window.resizable(0,0)       # 防止用户调整尺寸

    inputfilepath1 = StringVar()
    inputfilepath2 = StringVar()
    outputfilepath = StringVar()
    outputfilename = StringVar()
    inputfilepath1.set('')
    inputfilepath2.set('')
    outputfilepath.set('')
    outputfilename.set('')


    Label(window,text = "Old DBC:").grid(row = 0, column = 0, sticky = "w")
    Entry(window, textvariable = inputfilepath1, width = 100).grid(row = 0, column = 2)
    Button(window, text = "Select", command = chooseinputfile1).grid(row = 0, column = 4)

    Label(window,text = "New DBC:").grid(row = 1, column = 0, sticky = "w")
    Entry(window, textvariable = inputfilepath2, width = 100).grid(row = 1, column = 2)
    Button(window, text = "Select", command = chooseinputfile2).grid(row = 1, column = 4)

    Label(window,text = "Output Path:").grid(row = 2, column = 0, sticky = "w")
    Entry(window, textvariable = outputfilepath, width = 100).grid(row = 2, column = 2)
    Button(window, text = "Select", command = chooseoutputfile).grid(row = 2, column = 4)

    Label(window,text = "Output File Name:").grid(row = 3, column = 0, sticky = "w")
    Entry(window, textvariable = outputfilename, width = 50).grid(row = 3, column = 2, sticky = "w")

    Button(window, text = "Generate", command = generateCT).grid(row = 4, column = 2)

    window.mainloop()
