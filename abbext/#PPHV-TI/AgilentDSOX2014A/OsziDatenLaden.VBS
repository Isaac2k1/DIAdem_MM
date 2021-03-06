'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2014-05-11 16:01:04
'-- Author: Rene Irion / rene.irion@ch.abb.com
'-- Comment: Load Data from Agilent DSO-X 2014A
'-- 
'-- This script will load the saved data from the PPHV-TI Oscilloscope DSO-X 2014a
'-- Data specification
'-- needs to be saved as .CSV file
'-- the last digits needs to be a number
'--
'-- example: Test_1.csv, Test_2.csv
'--
'-- For a correct import you need the following data plugin "CSV_DataImport4Agilent.uri"
'-- Stored in: "C:\DIAdem\abbext\#PPHV-TI\AgilentDSOX2014A\"
'-- 
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.

'-- Declaration
Dim StandardFolder, StandardName, StandardBegin, StandardEnd, StandardChannelNo
Dim count_a, count_b
Dim MyFolders()
Dim RawDataName
Dim DataStart, DataEnd
Dim ChannelNo
Dim ChannelName()

'-- Standard settings
StandardFolder = "C:\Users\chreiri\Desktop\Messungen_Oszi\Test\"
StandardName = "scope_"
StandardBegin = 1
StandardEnd = 1
StandardChannelNo = 1


'-- Script
Call UserInput

Call DataHandling


'-------------------------------------------------------------------------------
Sub UserInput
  '-- Info
  msgbox("This script will load measurement data from Agilent DSO-X 2014A, all data needs to be saved as NAME_XX.csv (XX are consecutive numbers).")
  
  '-- Location of the data
  ReDim MyFolders(1)
  MyFolders(0) = InputBox("Enter the data location","Folder",StandardFolder) 

  '-- Name of the Raw Data
  RawDataName = InputBox("Enter name of the measurement data without number or suffix. eg. ""Test_1.csv"" --> ""Test_""","Data name",StandardName)
  
  '-- Define the first measurement to load
  DataStart = InputBox("Enter the number of the FIRST measurement","START",StandardBegin) 
  
  '-- Define the last measurement to load
  DataEnd = InputBox("Enter the number of the LAST measurement","END",StandardEnd)
  
  End Sub
  
  
  '-------------------------------------------------------------------------------
Sub DataHandling
 
For count_a = DataStart to DataEnd
  
  'Load
  Call DataFileLoad(MyFolders(0) & RawDataName & count_a & ".csv","CSV_DataImport4Agilent","Load")
  
  Select Case DataStart
  
  Case  "" & count_a & ""
    '-- Number
    ChannelNo = data.Root.ActiveChannelGroup.Channels.Count
    ReDim ChannelName(ChannelNo) 
    
    for count_b = 2 to ChannelNo
      ChannelName(count_b) = InputBox("Name of channel " & count_b-1 & ": ","Name of Channel " & count_b-1, Data.Root.ChannelGroups(GroupDefaultGet).Channels(count_b).Properties("name").Value) 
    next
    
  Case Else
  
  End Select 
  
  'Preparation
  for count_b = 2 to ChannelNo
    'Rename
    Data.Root.ChannelGroups(GroupDefaultGet).Channels(count_b).Name = ChannelName(count_b)
    
    'Numeric2Waveform
    Call ChnToWfChn("[" & GroupDefaultGet & " ]/x-axis","[" & GroupDefaultGet & "]/" & ChannelName(count_b),0,"WfXRelative")
  
    'Offset
    Call ChnOffset("[" & GroupDefaultGet & "]/" & ChannelName(count_b),"[" & GroupDefaultGet & "]/" & ChannelName(count_b),50,"mean value offset")
  
  next

  Call Data.Root.ChannelGroups(GroupDefaultGet).Channels.Remove("x-axis")

Next


End Sub