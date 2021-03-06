'-------------------------------------------------------------------------------
'-- VBS-Script-File
'-- Created: 2009-08-07 
'-- Authors:	Rene Irion 
'				Jonas Schwammberger
'-- Version: 1.9

'-- Purpose: Main file of the DIAdem2Excel Application.
'-- History:
'     1.01 Alpha Bugfixing 
'     1.02 Alpha Bugfixing
'     1.03 Alpha Hotfix for processing with a 6x Polynom
'     1.04 Alpha Hotfix and Bugfix for processing with a 6x Polynom
'     1.05 Alpha Polynom can have any level.
'     1.06 Alpha developing Smoothtravelcurve
'     1.07 Alpha developing Smoothtravelcurve 
'-------------------------------------------------------------------------------
Option Explicit 
   
Dim Shots()
Dim ShotType()
Dim ShotOperation()
Dim numberOfShots

Dim MechChannels()          'Channel name
Dim MechNewChnName()        'New channel name
Dim MechColumn()            'column in sdf file
Dim MechMultVal()           'factor
Dim MechStartVal()          'start value
Dim MechStroke()            'stroke value
Dim MechApplPolynoms()      'apply polynom
Dim MechDiagram()           'diagram
Dim MechChannelType()
Dim MechNumbOfChn           'anzahl der kanäle


Dim PowChannels()           'analog zu mech
Dim PowNewChnName()
Dim PowColumn()
Dim PowMultVal()
Dim PowStartVal()
Dim PowStroke()
Dim PowRemInductive()
Dim PowApplPolynoms()
Dim PowDiagram()            'defines on what diagram a channel should be.
Dim PowChannelType()
Dim PowNumbOfChn

Dim Polynom()
Dim PolNumber

Dim EnergyChannel()         'contains energychanneldescription

'Dim TriggerChannel

'SerialNumber               'these variables are global because the GUI needs them
'firstPowerShot

'configPath
'Inductance
'resistance

'outputFolder               'this global variable holds the path of the output folder

dim myFolders()             'holds all files and folders used by this programm

'dim correctiveOffset        'offset used by CropChannels()
'dim samplingRate            'sampling rate to make all channels equidistant
'dim threshold               'threshold level for finding the rising edge used by CropChannels()

dim smoothParam()		        'parameter array for the Smooth Travel Curve dialog

dim errorLog
'-------------------------------------------------------------------------------
'use functions defined in DIAdem2Excel_functions.VBS and Output_functions.VBS
call init()

Call ScriptCmdAdd(myFolders(0)&"\ClearTravelSpikes.vbs")
Call ScriptCmdAdd(myFolders(0)& "..\equix\equix.vbs") 'Must be in the same Path as this File
Call ScriptCmdAdd(myFolders(1))
Call ScriptCmdAdd(myFolders(2))
Call ScriptCmdAdd(myFolders(8))

call initClearTravel
globaldim("spikeFactor")
globaldim("smoothFactor")
globaldim("smoothCurve")
spikeFactor = globalFactor
smoothFactor = globalFactor
smoothCurve = 1               'true

Call CallDialog()
'-------------------------------------------------------------------------------
'summary: initialize variables
sub init()

  Redim myFolders(8)
    myFolders(0) = "C:\DIAdem\abbext\DIAdem2Excel\"               'folder with all sourcefiles
    myFolders(1) = myFolders(0)&"DIAdem2Excel_functions.VBS"      'VBS Function collection
    myFolders(2) = myFolders(0)&"Output_Functions.vbs"            'VBS Output Function collection
    myFolders(3) = myFolders(0)&"DIAdem2ExcelWind.sud"            'dialog file
    myFolders(4) = myFolders(0)&"Report.TDR"                      'Powershot Report template
    myFolders(5) = myFolders(0)&"Report.TDR"                      'Mechanical shot Report template
    myFolders(6) = myFolders(0)&"ExanpleREPORT.TDR"               'Example plot template
    myFolders(7) = myFolders(0)&"ClearTravelREPORT.TDR"
    myFolders(8) = myFolders(0)&"WindowsZip.VBS"
    
  GlobalDim("correctiveOffset")
  GlobalDim("samplingRate")
  GlobalDim("threshold")
  samplingRate = 0.0001
  correctiveOffset = 0.002
  threshold = 0.02
  
  GlobalDim("SerialNumber")
  GlobalDim("firstPowerShot")
  
  GlobalDim("runningFlag")
  runningFlag = 0
  
  GlobalDim("configPath")
  GlobalDim("Inductance")
  Inductance = 0.000007
  GlobalDim("Resistance")
  Resistance = 0.002
  
  GlobalDim("outputFolder")
  
  GlobalDim("Version")
    Version = "1.9"
  GlobalDim("TriggerChannel")
    TriggerChannel = "VS"
  
  redim EnergyChannel(1)
    EnergyChannel(0) = false
  
  ReDim MechChannels(1)
    MechChannels(0) = TriggerChannel
  MechNumbOfChn = 0
  
  ReDim PowChannels(1) 
    PowChannels(0) = TriggerChannel
  PowNumbOfChn = 0
  
  numberOfShots = -1                                              'has to be -1 to indicate that no cfg file has been read
  
  errorLog = ""
  
  Redim Polynom(0)
  PolNumber = 0
  
end sub
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
sub initClearTravel()
  'Init global variables from ClearTravelSpikes.vbs
  globalFactor = 9
  globalResponse= ""
  globalSkipAll = false  
end sub
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'summary: Calls Dialogs. Any communication between the scripts and the dialogs happens here.
'parameter: none
'output: none
sub CallDialog()
  dim repeat
  dim return
  dim debug
  
  if runningFlag = 0 Then
    do
      repeat = false
      debug = runningFlag
      
      'call window
      runningFlag = 1
      return = Suddlgshow("ChannelSetup", myFolders(3))
      
      if  return = "IDCancel" then
        runningFlag = 0
        Exit Do
      Else
      
        'load configuration file
        call LoadCfg(configPath)
      
        'check if there are powershots. if not, we don't need to display the next window
        firstPowerShot = GetFirstPowerShot()
        debug = runningFlag
      
        if firstPowerShot <> "" then
        
          globalResponse = "testvalue"
        
          do while globalResponse = "testvalue"
            debug = runningFlag
           'load example reports
            Data.Root.Clear()

            Call LoadExampleShot(firstPowerShot,SerialNumber)
            Call LoadExampleReport(firstPowerShot,SerialNumber)
        
            debug = runningFlag
        
           'show inductiveresistance window
            if Suddlgshow("InductiveResistance",myFolders(3)) ="IDOk" and globalResponse = "apply" then
              debug = runningFlag
             runningFlag = 0
             call Export()
           else
             'user wants to go back to dialog ChannelSetup
             debug = runningFlag
             repeat = true
           end if
          loop
          
        else
          runningFlag = 0
          call Export()
        end if
      End if
      
    loop while repeat = true
  end if
  
end sub
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'summary: Load Shot, Crop Channels, save it in Excel, multiply channel, set offset
'   calculate Polynom, clear Arc voltage and save it in .sdf file
'parameter: none
'output: none
sub Export()
  dim ShotName        'shotname to import
  dim outputFile
  dim outputName
  dim duration        'duration of the shot to cut out
  dim i               'indexer
  dim eqGroup         'channelgroup with equidistant values
  dim travelStart     'travelcurve start value according to the Operation
  dim travelMult
  dim fso
  dim debug
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  
  'Export shots
  for i = 0 to numberOfShots
    dim headerWritten
    dim j
    
    outputFile = SerialNumber & "_" & shots(i)
    outputName = SerialNumber & "/" & shots(i)
    ShotName = "l" & SerialNumber & "-" & shots(i) & ".rbd"
    eqGroup ="l" & SerialNumber & "-" & shots(i) & ".rbe"
    
    'determine operation time and travel begin/end
    If ShotOperation(i) = "O" Then
      duration = 0.15
    Elseif ShotOperation(i) = "C" Then
      duration = 0.15
    Elseif ShotOperation(i) = "CO" Then
      duration = 0.3
    ElseIf ShotOperation(i) = "OCO" Then
      duration = 0.5
      travelStart = 0
    End If
    
    Call Data.Root.Clear()
    
    if LoadShot(ShotName) = true Then      
      'determine power or mech shot
      If ShotType(i) = "P" Then
        Call CreateEquidistantChannels(ShotName,PowChannels,PowNumbOfChn,samplingRate)
        Call CropChannels(eqGroup,PowChannels,PowNumbOfChn,duration,threshold,samplingRate,correctiveOffset)
        Call CreateExcelFile(outputFile,eqGroup,PowChannels,PowNumbOfChn)
        
        'check if all shots exist
        headerWritten = false
        
        'multiply,start value and polynom
        'skip VS channel
    for j = 1 to PowNumbOfChn
      if Data.Root.ChannelGroups.Item(eqGroup).Channels.Exists(PowChannels(j)) = false then
        if headerWritten = false then
          headerWritten = true
          errorLog = errorLog & "Error in Shot: "& ShotName & vbCr
        End if
        
        errorLog = errorLog &"    Couldn't find channel: " & PowChannels(j) & VbCr
        
      Else
        if PowStroke(j) <> "NULL" Then
          travelMult = val(powMultVal(j))
          Call MultiplyChannel(eqGroup,PowChannels(j),travelMult)
          call chnoffset(eqGroup&"/"&PowChannels(j),eqGroup&"/"&PowChannels(j),0,"min. value offset")
  
          'set polynom
          if PowApplPolynoms(j) = "true" then
            Call CalculateTravel(eqGroup,PowChannels(j),polynom)
            call chnoffset(eqGroup&"/"&PowChannels(j),eqGroup&"/"&PowChannels(j),0,"min. value offset")
          end if
  
          Call ShowSmoothDialog(eqGroup,PowChannels(j))
          call chnoffset(eqGroup&"/"&PowChannels(j),eqGroup&"/"&PowChannels(j),0,"min. value offset")
        Else
          Call MultiplyChannel(eqGroup,PowChannels(j),VAL(PowMultVal(j)))
  
          if PowStartVal(j) <> "NULL" Then
            call chnoffset(eqGroup&"/"&PowChannels(j),eqGroup&"/"&PowChannels(j),0,"first value offset")
            call chnoffset(eqGroup&"/"&PowChannels(j),eqGroup&"/"&PowChannels(j),VAL(PowStartVal(j)),"free offset")
          end if
        End If
      end if
    next
        
        'remove inductive offset                   
        for j = 0 to PowNumbOfChn
          if PowRemInductive(j) = "true" and cno(eqGroup&"/"&PowChannels(j)) > 0 and cno(eqGroup&"/"&GetCurrent()) > 0 then
            Call RemoveInductivResistantParts(eqGroup&"/"&PowChannels(j),eqGroup&"/"&GetCurrent(),samplingRate,Inductance, Resistance)
          End if
        next
        
        'Calculate energy if requested
        if EnergyChannel(0) = true then
          call CalculateEnergy(eqGroup,EnergyChannel(1),EnergyChannel(3),EnergyChannel(4))
          
          'EnergyChannel(2) column
          'add energychannel
          PowNumbOfChn = PowNumbOfChn + 1

          Call CreateSDFFile(outputFile,false,eqGroup,PowNewChnName,PowChannels,PowNumbOfChn,powColumn,mechNewChnName,mechChannels,mechNumbOfChn,mechColumn,samplingRate)
          'Call CreateSDFFile(outputFile,PowNewChnName,eqGroup,PowChannels,PowNumbOfChn,samplingRate)
        Else
          Call CreateSDFFile(outputFile,false,eqGroup,PowNewChnName,PowChannels,PowNumbOfChn,powColumn,mechNewChnName,mechChannels,mechNumbOfChn,mechColumn,samplingRate)
          'Call CreateSDFFile(outputFile,PowNewChnName,eqGroup,PowChannels,PowNumbOfChn,samplingRate)
        end if
        
        'rename
        'skip trigger signal
        For j = 1 to PowNumbOfChn
          if Data.Root.ActiveChannelGroup.Channels.Exists(PowChannels(j)) = true Then
            Data.Root.ChannelGroups(GroupDefaultGet).Channels(PowChannels(j)).Name = PowNewChnName(j)
          End if
        Next
        
        call CreatePowerReport(myFolders(4), outputFile, outputName, eqGroup, PowChannels,PowNewChnName,PowNumbOfChn,PowDiagram)
        
      ElseIf ShotType(i) = "M" Then
        debug = Mechchannels(0)
        
        Call CreateEquidistantChannels(ShotName,MechChannels,MechNumbOfChn,samplingRate)
        Call CropChannels(eqGroup,MechChannels,MechNumbOfChn,duration,threshold,samplingRate,correctiveOffset)
        Call CreateExcelFile(outputFile,eqGroup,MechChannels,MechNumbOfChn)
       
        
        'check if all shots exist
        headerWritten = false
        
        'skip "VS" channel
        for j = 1 to MechNumbOfChn
          if Data.Root.ChannelGroups.Item(eqGroup).Channels.Exists(MechChannels(j)) = false then
            if headerWritten = false then
              errorLog = errorLog & "Error in Shot: "& ShotName & VbCr
              headerWritten = true
            End if
            errorLog = errorLog &"    Couldn't find channel: " & MechChannels(j) & VbCr
          Else
            if MechStroke(j) <> "NULL" Then
'              If ShotOperation(i) = "O" Then
'                travelStart = 0
'                travelMult = VAL(MechMultVal(j))
'              Elseif ShotOperation(i) = "C" Then
'                travelStart = VAL(MechStroke(j))
 '               travelMult = VAL(MechMultVal(j)) ' removed: factor -1 #TW 21.06.10
 '             Elseif ShotOperation(i) = "CO" Then
 '               travelStart = VAL(MechStroke(j))
 '               travelMult = VAL(MechMultVal(j)) ' removed: factor -1 #TW 21.06.10
 '             ElseIf ShotOperation(i) = "OCO" Then
 '               travelStart = 0
 '               travelMult = VAL(MechMultVal(j))
  '            End If
  
              travelMult = VAL(MechMultVal(j))
              Call MultiplyChannel(eqGroup,MechChannels(j),travelMult)
              call chnoffset(eqGroup&"/"&MechChannels(j),eqGroup&"/"&MechChannels(j),0,"min. value offset")
            
'              if travelStart <> "NULL" Then
'                call chnoffset(eqGroup&"/"&MechChannels(j),eqGroup&"/"&MechChannels(j),0,"first value offset")
'                call chnoffset(eqGroup&"/"&MechChannels(j),eqGroup&"/"&MechChannels(j),travelStart,"free offset")
'              end if
            
              'set polynom
              if MechApplPolynoms(j) = "true" then
                Call CalculateTravel(eqGroup,MechChannels(j),polynom)
              end if 
            
            Else        
              Call MultiplyChannel(eqGroup,MechChannels(j),VAL(MechMultVal(j)))
           
              if MechStartVal(j) <> "NULL" Then
                call chnoffset(eqGroup&"/"&MechChannels(j),eqGroup&"/"&MechChannels(j),0,"first value offset")
                call chnoffset(eqGroup&"/"&MechChannels(j),eqGroup&"/"&MechChannels(j),VAL(MechStartVal(j)),"free offset")
              end if
            End If
          end if
        next
        
        Call CreateSDFFile(outputFile,true,eqGroup,PowNewChnName,PowChannels,PowNumbOfChn,powColumn,mechNewChnName,mechChannels,mechNumbOfChn,mechColumn,samplingRate)
        'Call CreateSDFFile(outputFile,true,eqGroup)
        
        'rename new channels
        'skip Trigger Signal
        For j = 1 to MechNumbOfChn
          if Data.Root.ActiveChannelGroup.Channels.Exists(MechChannels(j)) = true Then
            Data.Root.ChannelGroups(GroupDefaultGet).Channels(MechChannels(j)).Name = MechNewChnName(j)
          end if
        Next
        
        Call CreateMechReport(myFolders(5), outputFile, outputName, eqGroup, MechChannels,MechNewChnName,MechNumbOfChn,MechDiagram)
        
      end if 
    Else
      errorLog = errorLog & "\n couldn't load shot: " & ShotName
    End if
  next
  
  'Add files to ZIP-Archive
  if numberOfShots > -1 then
    Call NewZip(outputFolder & SerialNumber & "_" & shots(0) & "-" & shots(numberOfShots) & "_SDF.zip", true)
    Call NewZip(outputFolder & SerialNumber & "_" & shots(0) & "-" & shots(numberOfShots) & "_Excel.zip", true)
    Call NewZip(outputFolder & SerialNumber & "_" & shots(0) & "-" & shots(numberOfShots) & "_PDF.zip", true)
  end if
  Call Pause(1)
  for i = 0 to numberOfShots
    outputFile = SerialNumber & "_" & shots(i)
    If fso.FileExists(outputFolder & "Excel\" & outputFile & ".xls") Then
      Call WindowsZip(outputFolder & "Excel\" & outputFile & ".xls", outputFolder & SerialNumber & "_" & shots(0) & "-" & shots(numberOfShots) & "_Excel.zip")        
    End If
    If fso.FileExists(outputFolder & "SDF\" & outputFile & ".sdf") Then
      Call WindowsZip(outputFolder & "SDF\" & outputFile & ".sdf", outputFolder & SerialNumber & "_" & shots(0) & "-" & shots(numberOfShots) & "_SDF.zip")        
    End If
    If fso.FileExists(outputFolder & "PDF\" & outputFile & ".pdf") Then
      Call WindowsZip(outputFolder & "PDF\" & outputFile & ".pdf", outputFolder & SerialNumber & "_" & shots(0) & "-" & shots(numberOfShots) & "_PDF.zip")        
    End If
  next
  
  'write error report
  Call WriteErrorReport(errorLog)
  
  Call MsgBox("All Done!",vbOKOnly,"ALL DONE!")
end sub
'-------------------------------------------------------------------------------



'-------------------------------------------------------------------------------
'summary: load a configuration file
'parameter cfgFile:
'output: none
Function LoadCFG(cfgFile)
  Dim BoolSuccess
  Dim fso, txt
  Dim line
  dim i
  dim debug
  
  'init
  i = 0  
  BoolSuccess = true
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set txt = fso.OpenTextFile(cfgFile, 1, 0)
  
  'reinitialize these variables to avoid exceptions
  ReDim MechChannels(1)
    MechChannels(0) = TriggerChannel
  
  ReDim PowChannels(1) 
    PowChannels(0) = TriggerChannel
 
  MechNumbOfChn = 0
  PowNumbOfChn = 0
  numberOfShots = -1
  
  'read configuration file
  Do While Not txt.AtEndOfStream = True
    line = txt.ReadLine
      
    If Left(line, 1) = "#" Then
      'comment line

    ElseIf Trim(line) = "" Then
      'empty line
      
    ElseIf Left(line, 11) = "Resistance:" Then
      Resistance = Trim(Right(line, Len(line) - 11))
      
    ElseIf Left(line, 11) = "Inductance:" Then
      Inductance = Trim(Right(line, Len(line) - 11))

    ElseIf Left(line, 7) = "Serial:" Then
      SerialNumber = Trim(Right(line, Len(line) - 7))

    'if serial number not 4 characters long, trim it.
    If Len(SerialNumber) < 4 Then _
      SerialNumber = String(4 - Len(SerialNumber), "0") & SerialNumber

    ElseIf Left(line, 5) = "Mech:" Then    'mechanical channel list
      MechNumbOfChn = MechNumbOfChn + 1
        
      ReDim Preserve MechChannels(MechNumbOfChn)
      ReDim Preserve MechNewChnName(MechNumbOfChn)
      ReDim Preserve MechColumn(MechNumbOfChn)
      ReDim Preserve MechMultVal(MechNumbOfChn)
      ReDim Preserve MechStartVal(MechNumbOfChn)
      ReDim Preserve MechStroke(MechNumbOfChn)
      ReDim Preserve MechApplPolynoms(MechNumbOfChn)
      Redim Preserve MechDiagram(MechNumbOfChn)
      redim preserve MechChannelType(MechNumbOfChn)
      dim MechDef                         'whole Mech channel definition
        
      'now distribute the values to the arrays
      MechDef = Trim(Right(line, Len(line) - 5))
      
      'set name to array      
      MechChannels(MechNumbOfChn) = Left(MechDef,InStr(MechDef, ";")-1)      
      MechDef = Mid(MechDef,InStr(MechDef, ";")+1)    'cut out name of channel
        
      'rinse and repeat for the other arrays

      MechNewChnName(MechNumbOfChn) = Left(MechDef,InStr(MechDef, ";")-1)
      MechDef = Mid(MechDef,InStr(MechDef, ";")+1)
     
      MechColumn(MechNumbOfChn) = Left(MechDef,InStr(MechDef, ";")-1)
      MechDef = Mid(MechDef,InStr(MechDef, ";")+1)
      
      MechMultVal(MechNumbOfChn) = Left(MechDef,InStr(MechDef, ";")-1)
      MechDef = Mid(MechDef,InStr(MechDef, ";")+1)
        
      MechStartVal(MechNumbOfChn) = Left(MechDef,InStr(MechDef, ";")-1)
      MechDef = Mid(MechDef,InStr(MechDef, ";")+1)
        
      MechStroke(MechNumbOfChn) = Left(MechDef,InStr(MechDef, ";")-1)
      MechDef = Mid(MechDef,InStr(MechDef, ";")+1)
       
      MechApplPolynoms(MechNumbOfChn) = Left(MechDef,InStr(MechDef, ";")-1)
      MechDef = Mid(MechDef,InStr(MechDef, ";")+1)
        
      MechDiagram(MechNumbOfChn) = Left(MechDef,InStr(MechDef, ";")-1)
      MechDef = Mid(MechDef,InStr(MechDef, ";")+1)
      
      MechChannelType(MechNumbOfChn) = MechChannels(MechNumbOfChn)  'Mechdef   ääääääääääääääääääääääääää
    ElseIf Left(line, 4) = "Pow:" Then     'power channel list
      PowNumbOfChn = PowNumbOfChn +1
        
      ReDim Preserve PowChannels(PowNumbOfChn)
      ReDim Preserve PowNewChnName(PowNumbOfChn)
      Redim Preserve PowColumn(PowNumbOfChn)
      ReDim Preserve PowMultVal(PowNumbOfChn)
      ReDim Preserve PowStartVal(PowNumbOfChn)
      Redim Preserve PowStroke(PowNumbOfChn)
      ReDim Preserve PowRemInductive(PowNumbOfChn)
      ReDim Preserve PowApplPolynoms(PowNumbOfChn)
      Redim Preserve PowDiagram(PowNumbOfChn)
      redim preserve PowChannelType(PowNumbOfChn)
      Dim PowDef                       'power channel definition
            
      PowDef = Trim(Right(line, Len(line) - 4))
        
      'channel name
      PowChannels(PowNumbOfChn) = Left(PowDef,InStr(PowDef, ";")-1)
      PowDef = Mid(PowDef,InStr(PowDef, ";")+1)
       
      'new channel name
      PowNewChnName(PowNumbOfChn) = Left(PowDef,InStr(PowDef, ";")-1)
      PowDef = Mid(PowDef,InStr(PowDef, ";")+1)
      
      PowColumn(PowNumbOfChn) = Left(PowDef,InStr(PowDef, ";")-1)
      PowDef = Mid(PowDef,InStr(PowDef, ";")+1)
      
      'multiplication value
      PowMultVal(PowNumbOfChn) = Left(PowDef,InStr(PowDef, ";")-1)
      PowDef = Mid(PowDef,InStr(PowDef, ";")+1)
      
      'start value
      PowStartVal(PowNumbOfChn) = Left(PowDef,InStr(PowDef, ";")-1)
      PowDef = Mid(PowDef,InStr(PowDef, ";")+1)
       
      'stroke (only for travel channels)
      PowStroke(PowNumbOfChn) = Left(PowDef,InStr(PowDef, ";")-1)
      PowDef = Mid(PowDef,InStr(PowDef, ";")+1)
       
      'remove inductance (contains "true" if inductance should be removed)
      PowRemInductive(PowNumbOfCHn) = Left(PowDef,InStr(PowDef, ";")-1)
      PowDef = Mid(PowDef,InStr(PowDef, ";")+1)
        
      PowApplPolynoms(PowNumbOfCHn) = Left(PowDef,InStr(PowDef, ";")-1)
      PowDef = Mid(PowDef,InStr(PowDef, ";")+1)
        
      'apply polynom (contains "true" if polynom should be applied)
      PowDiagram(PowNumbOfChn) = Left(PowDef,InStr(PowDef, ";")-1)
      PowDef = Mid(PowDef,InStr(PowDef, ";")+1)
      
      PowChannelType(PowNumbOfChn) = PowDef' PowDef      ääääääääääääääääääääääääääää
      
    ElseIf Left(line,5) = "Shot:" Then
      numberOfShots = numberOfShots + 1
        
      ReDim Preserve Shots(numberOfShots)
      ReDim Preserve ShotType(numberOfShots)
      ReDim Preserve ShotOperation(numberOfShots)
      Dim ShotDef
          
      'cut out "Shot:"
      ShotDef = Trim(Right(line, Len(line) - 5))
        
      'distribute values to the arrays
      Shots(numberOfShots) = Left(ShotDef,InStr(ShotDef, ";")-1)
      ShotDef =Mid(ShotDef,InStr(ShotDef, ";")+1)
          
      ShotType(numberOfShots) = Left(ShotDef,InStr(ShotDef, ";")-1)
      ShotDef =Mid(ShotDef,InStr(ShotDef, ";")+1)
        
      ShotOperation(numberOfShots) = ShotDef
           
      If Len(Shots(numberOfShots)) < 4 Then _
        Shots(numberOfShots) = String(4 - Len(Shots(numberOfShots)), "0") & Shots(numberOfShots) 'make it four char long
    ElseIf Left(line, 8) = "polynom:" Then
      polNumber = polNumber + 1
      Redim Preserve Polynom(polNumber)
      Polynom(polNumber-1) =  VAL(Right(line,Len(line) - InStr(line, ";")))
    ElseIf Left(line, 7) = "Energy:" Then
      ReDim EnergyChannel(4)
      'NOTE: The Energy Channel definition is ALWAYS at the end of the CFG file.
      ' the Energy Channel Definition is appended at the end of all ohter powerchanneldefinitions without
      ' increasing the "PowNumbOfChn" variable, so the programm doesn't try to import and process
      ' because the Energy has to be calculated AFTER the Channels have been imported and processed.
      EnergyChannel(0) = true
        
      'cut out "Energy:"
      ShotDef = Trim(Right(line, Len(line) - 7))
      EnergyChannel(1) = Left(ShotDef,InStr(ShotDef, ";")-1)
      ShotDef =Mid(ShotDef,InStr(ShotDef, ";")+1)
        
      EnergyChannel(2) = Left(ShotDef,InStr(ShotDef, ";")-1)
      ShotDef =Mid(ShotDef,InStr(ShotDef, ";")+1)
        
      EnergyChannel(3) = Left(ShotDef,InStr(ShotDef, ";")-1)
      ShotDef =Mid(ShotDef,InStr(ShotDef, ";")+1)
        
      EnergyChannel(4) = ShotDef
        
      redim preserve PowChannels(numberOfShots + 1)
      redim preserve PowNewChnName(numberOfShots + 1)
      redim preserve PowColumn(numberOfShots + 1)
      redim preserve PowMultVal(numberOfShots + 1)
      redim preserve PowStartVal(numberOfShots + 1)
      redim preserve PowStroke(numberOfShots + 1)
      redim preserve PowRemInductive(numberOfShots + 1)
      redim preserve PowApplPolynoms(numberOfShots + 1)
      redim preserve PowDiagram(numberOfShots + 1) 
      
      debug = EnergyChannel(0)
      debug = EnergyChannel(1)
      debug = EnergyChannel(2)
      debug = EnergyChannel(3)
      debug = EnergyChannel(4)
        
      'add energychannel
        
    Else
      On Error Resume Next  'in case of a bad shot definition
          
      If Err.Number <> 0 Then    'an error occured
        MsgBox "Corrupted Data found" & vbCRlf & _
            "The data in question is: " & temp & vbCrLf, 16, "Corrupted Data"
            boolSuccess = True
            Exit Do
      End If

      On Error GoTo 0 'no exception handling after this point
   End If
  Loop

  txt.Close
  Set fso = Nothing
  
  LoadCFG = BoolSuccess
end Function
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'summary: Shows a dialog to user to smooth travelCurve
'parameter
'output: none
Sub ShowSmoothDialog(ChannelGroup, ChannelName)
	dim dlgReturn
	dim testGroup
	dim channels()
	dim template
	dim run

	run = true
	testGroup = "SmoothTravel"
	dlgReturn = ""

	redim channels(1)
	channels(0) = testGroup&"/"&"TravelCurve"
	channels(1) = testGroup&"/"&"TravelCurve_original"

	if globalSkipAll = false Then

		'copy curves
		Data.Root.ChannelGroups.Add(testGroup)
		call chnCopy(ChannelGroup&"/"&ChannelName,channels(0))
		call chnCopy(ChannelGroup&"/"&ChannelName,channels(1))

		'load report and set curves
		Call PicLoad(myFolders(7))
		Call GraphObjOpen("2D-Axis1")
      Call GraphObjOpen("2DObj5_Curve1")
        D2CChnYName      = channels(1)
      Call GraphObjClose("2DObj5_Curve1")
      Call GraphObjOpen("2DObj5_Curve2")
        D2CChnYName      = channels(0)
      Call GraphObjClose("2DObj5_Curve2")
		Call GraphObjClose("2D-Axis1")
    call chnCopy(ChannelGroup&"/"&ChannelName,channels(0))

		globalFactor = spikeFactor
		call ClearTravelSpikes(channels(0),channels(1),globalfactor,samplingrate) 

		Call Picupdate(0)

		do while(run)
			call Suddlgshow("RemoveSpikes",myFolders(3))
			if globalResponse = "apply" Then
				call ClearTravelSpikes(ChannelGroup&"/"&ChannelName,channels(1),globalfactor,samplingrate)
				run = false
			Elseif globalResponse  = "skip" Then
				run = false
			Elseif globalResponse = "skipall" Then
				globalskipall = true
				Data.Root.ChannelGroups.Remove(testGroup)
				run = false
			Elseif globalResponse = "testvalue" Then
				'run through loop again
        call chnCopy(ChannelGroup&"/"&ChannelName,channels(0))
				call ClearTravelSpikes(channels(0),channels(1),globalfactor,samplingrate) 
				Call Picupdate(0)      
			end if
		loop

		spikeFactor = globalFactor

		if not globalskipall and smoothCurve = 1 then
			run = true
			globalFactor = smoothFactor
			call chnCopy(ChannelGroup&"/"&ChannelName,channels(1))
			call ChnSmooth(channels(1),channels(0),globalFactor,"maxNumber")
			Call Picupdate(0)   

			do while(run)
				call Suddlgshow("SmoothCurve",myFolders(3))
				if globalResponse = "apply" Then
					call ChnSmooth(channels(1),ChannelGroup&"/"&ChannelName,globalFactor,"maxNumber")
					Data.Root.ChannelGroups.Remove(testGroup)
					run = false
				Elseif globalResponse  = "skip" Then
					Data.Root.ChannelGroups.Remove(testGroup)
					run = false
				Elseif globalResponse = "skipall" Then
					globalskipall = true
					Data.Root.ChannelGroups.Remove(testGroup)
					run = false
				Elseif globalResponse = "testvalue" Then
					'run through loop again
          call chnCopy(ChannelGroup&"/"&ChannelName,channels(0))
					call ChnSmooth(channels(1),channels(0),globalFactor,"maxNumber")
					Call Picupdate(0)      
				end if
			loop

			smoothFactor = globalFactor
		end if    
	end if
end sub
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'summary:
'parameter cfgFile:
'output: none
Function GetFirstPowerShot()
  dim ShotNumber
  dim i
  
  i = numberOfShots
  ShotNumber = ""
  
  for i = 0 to numberOfShots
    if ShotType(i) = "P" then
      if ShotNumber = "" then
        ShotNumber = Shots(i)
      end if
    end if
  next
  
  GetFirstPowerShot = ShotNumber
end function
'-------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
'summary:
'parameter base:
Sub CalculateEnergy(strGroup,strEnergyChannel,strCurrent,strInj)
  dim i
  dim debug
  
  Data.Root.ChannelGroups(strGroup).Activate()
  call ChnAlloc(strEnergyChannel, ChnLength(strArc), 1, ChnValueDataType(strArc))

  for i = 1 to chnlength(strArc)
    CHD(i,strGroup&"/"&strEnergyChannel) = CHD(i,strGroup&"/"&strArc)*CHD(i,strGroup&"/"&strInj)
  next
  
end sub
'--------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'summary: loads and crops an example shot
'parameter exampleShot:
'parameter osc: oscillation number
'output: none
Sub LoadExampleShot(exampleShot,osc)
  dim shot
  dim eqShot
  Dim Channels()
  Dim exampleGroup
  dim debug
  
  shot = "l"&osc&"-"&exampleShot&".rbd"
  eqShot = "l"&osc&"-"&exampleShot&".rbe"
  
  'original channels
  redim Channels(3)
   Channels(0) = TriggerChannel
   Channels(1) = GetArcVoltage
   Channels(2) = GetCurrent
   Channels(3) = "arc voltage original"

  If Data.Root.ChannelGroups.Exists(eqshot) = true Then
    ChnDel(eqshot&"/"&Channels(1))
    Call ChnCopy(eqshot&"/"&Channels(3),eqshot&"/"&Channels(1))
    Call RemoveInductivResistantParts(eqshot&"/"&Channels(1),eqshot&"/"&Channels(2),samplingRate,Inductance, Resistance)
  Else
    Call Data.Root.Clear()
    call LoadShot(shot)
    
    Call CreateEquidistantChannels(shot,channels,2,samplingRate)
    Call ChnCopy(eqshot&"/"&Channels(1),eqshot&"/"&Channels(3))
    debug = Channels(0)
    debug = Channels(1)
    debug = Channels(2)
    debug = Channels(3)
    Call RemoveInductivResistantParts(eqshot&"/"&Channels(1),eqshot&"/"&Channels(2),samplingRate,Inductance, Resistance)
  End If
  
end sub
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'summary: Creates the report for the example shot
'parameter
'parameter
'output: none
'note: this procedure is intended to be called after LoadExampleShot()
sub LoadExampleReport(exampleShot,osc)
  dim shot
  dim eqShot
  Dim Channels()
  Dim eqChannels()
  Dim duration
  dim exampleGroupName
  
  shot = "l"&osc&"-"&exampleShot&".rbd"
  eqShot = "l"&osc&"-"&exampleShot&".rbe"
  
  redim Channels(2)
   Channels(0) = shot&"/"&TriggerChannel
   Channels(1) = shot&"/"&GetArcVoltage()
   Channels(2) = shot&"/"&GetCurrent()
  
  redim eqChannels(2)
   eqChannels(0) = eqshot&"/"&TriggerChannel
   eqChannels(1) = eqshot&"/"&GetArcVoltage()
   eqChannels(2) = eqshot&"/"&GetCurrent()
  
  'display exampleReport
  Call PicLoad(myFolders(0)&"/exampleREPORT.TDR")              '... PicFile 
  Call PicUpdate(0)                       '... PicDoubleBuffer
  
'------------------- Curve and axis definition ---------------------
Call GraphObjOpen("2D-Axis1")
  '------------------- Curve list -------------------------------
  Call GraphObjOpen("2DObj2_Curve1")
    D2CChnYName      = eqChannels(1)
    D2CAxisPairNo    = 1

  Call GraphObjClose("2DObj2_Curve1")
 
  Call GraphObjOpen("2DObj2_Curve2")
    D2CChnYName      = eqChannels(2)
    D2CAxisPairNo    = 2
  Call GraphObjClose("2DObj2_Curve2")
  '------------------- Position ---------------------------------
Call GraphObjClose("2D-Axis1")

 Call PicUpdate(0)                       '... PicDoubleBuffer 
end sub
'-------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
'summary:
'parameter base:
function GetArcVoltage()
  dim i
  
  for i = 0 to PowNumbOfChn
      if PowChannelType(i) = "arc_voltage" then
        exit for
      end if
      if i = pownumbofchn then
        getArcVoltage = "noArcVoltage"
        msgbox("no Arc Voltage, i will crash now...")
        exit function
      end if  
  next
  
  GetArcVoltage = PowChannels(i)
end Function
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
'summary:
'parameter base:
function GetCurrent()
  dim i
  for i = 0 to PowNumbOfChn
      if PowChannelType(i) = "I-Shunt" then
        exit for
      end if
      if i = pownumbofchn then
        getCurrent = "noCurrent"
        exit function
      end if      
  next
  
  GetCurrent = PowChannels(i)
End function
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
'summary:
'parameter base:
function GetTRV()
  dim i
  for i = 0 to PowNumbOfChn
      if PowChannelType(i) = "TRV" then
        exit for
      end if
  next
  
  GetTRV = PowChannels(i)
End function
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
'summary:
'parameter base:
Sub CreateMechReport(templatePath, outFile, outName, strGroup, ChannelNames,NewChannelNames,MaxNumbOfChn,DiagramIndex)
  Dim boolSuccess                         'true if function was successful
  Dim intSheetNo                          'sheet
  Dim intSheetNoMax
  Dim intCurveIndex
  Dim intCurveNoMax
  dim i
  dim actualChannelNr
  Dim ActualChannelIndex()
  
  Dim fso
  Dim folder

  Set fso = CreateObject("Scripting.FileSystemObject")

  If (fso.FolderExists(outputFolder & "\PDF")) Then

  Else
    Set folder = fso.CreateFolder(outputFolder & "\PDF")
  End If
  
  redim ActualChannelIndex(0)
  actualChannelNr = 0
  
  'get actual existing channels. if it exists, write the index in actualChannelIndex()
  For i = 0 to MaxNumbOfChn
    if Data.Root.ActiveChannelGroup.Channels.Exists(NewChannelNames(i)) = true Then
      actualChannelNr = actualChannelNr + 1
      redim preserve ActualChannelIndex(actualChannelNr)
      ActualChannelIndex(actualChannelNr-1) = i
    end if
  next
  
  'Set Title
  Call Data.Root.Properties.Item("description").Value("Report of mechanical shot: " & outName)
  
  'get maximum number of sheets
  intSheetNoMax = 0
  
  For i = 0 to MaxNumbOfChn
    if DiagramIndex(i) > intSheetNoMax then
      intSheetNoMax = DiagramIndex(i)
    end if
  Next
  dim debug
  
  debug = intSheetNoMax
  
  'loop over all report sheets
  For intSheetNo = 1 to intSheetNoMax               'intSheetNoMax = max. no. of diagrams
    intCurveIndex = 0
    
    If intSheetNo = 1 Then                        'Load Template
      Call PicLoad(templatePath)
    Else
      Call PicFileAppend(templatePath)            'Add Sheet from template
    End If    
     
    Call PicUpdate(0)                             '... PicDoubleBuffer 
     
    '------------- open plot -----------------
    Call GraphObjOpen("2DAxis1")                  '2DAxis --> Name of the diagramm
                                                  '1      --> Because it is the first diagram on this report

    '------------------- Curve list -------------------------------
    For i = 1 to MaxNumbOfChn                        'intCurveNoMax = max. no. of curves on the second y-axis
    
    'D2LegTxtTypeA(i) = "Free text"
    'D2LegTxtFreeA(i) = NewChannelNames(i)
    
      If DiagramIndex(i) = 0 Then
        Call  GraphObjNew("2D-Curve","Curve_" & intCurveIndex)    'Plot the travel curve on the first y-axis   
        Call GraphObjOpen("Curve_" & intCurveIndex)              
          D2CChnExpand     = 1                                
          D2CChnYName      = NewChannelNames(i)           'Array name of travel curve
          D2CAxisPairNo    = 1
        Call GraphObjClose("Curve_" & intCurveIndex)           
       
        'Call GraphObjOpen("2DYAxis4_1")
        '  D2AxisYTxt       = NewChannelNames(i)
        'Call GraphObjClose("2DYAxis4_1")                  
        
        'Call PicUpdate(0)  'maybe delete that
        intCurveIndex = intCurveIndex + 1
                  
      Elseif VAL(DiagramIndex(i)) = intSheetNo Then
        Call  GraphObjNew("2D-Curve","Curve_" & intCurveIndex)     'Plot other curves on the second y-axis
        Call GraphObjOpen("Curve_" & intCurveIndex)            
          D2CChnExpand     = 1                              
          D2CChnYName      = NewChannelNames(i)                               'Array name of other curves
          D2CAxisPairNo    = 2
        Call GraphObjClose("Curve_" & intCurveIndex)           

        intCurveIndex = intCurveIndex + 1
      Else
        
      End If
        
    Next

    '------------- close plot -----------------
    Call GraphObjClose("2DAxis1")
              
    '------------- update plot -----------------
    Call PicUpdate(0)
    
    '------------- rename sheet -----------------
    Call GraphSheetNGet(intSheetNo)
    Call GraphSheetRename(GraphSheetName,"Page " & intSheetNo)
    
  Next
  
  Call PicPDFExport(outputFolder & "\PDF\" & outFile & ".pdf",0)             'Save as PDF
  
  Call GraphDeleteAll()

end Sub
'--------------------------------------------------------------------------------