'-------------------------------------------------------------------------------
'-- VBS-Script-File
'-- Created: 2010-08-02 
'-- Authors: Jonas Schwammberger
'-- Version: 1.0.0

'-- Purpose: Main file of the TestRecord Application. 
'-- History:
'-------------------------------------------------------------------------------
Option Explicit

'Note: theoretically all those variables declared here should be available
'	for other scripts added with "ScriptCmdAdd()" and shouldn't have to be
'	declared global with the "GlobalDim()" function. But for some reason DIAdem
'	doesn't run stable. To be on the save side,if you need to use variables in functions
'	saved in other files, declare them global.

'dim rootFolder
'dim functionLibrary		'name of the Function Library of the TestRecord Application
'dim FrontTemplate			'name of the Frontpage Report Template
'dim DataTemplate			  'name of the Datapage Report Template
'dim PicTemplate			  'name of the Picture Report Template
'dim iResultSelected    'used in GUI, varaible contains index of selected item in lbResults
'dim oGUIParam          
'dim Version            
dim sudName				      'name of the .sud file
dim searchDialog			  'name of the search dialog in the .sud
dim TableDialog         'name of the select channels dialog in the .sud
dim cfgPath             'filepath of configuration file with all datatable standards
'-------------------------------------------------------------------------------



'-------------------------------------------------------------------------------
'summary: initializes global and local variables
'parameter: none
'output: none
Sub init()
	Globaldim("rootFolder")
		rootFolder = "C:\DIAdem\abbext\TestRecord"
	Globaldim("functionLibrary")
		functionLibrary = "TestRecordFunctions.vbs"
	Globaldim("FrontTemplate")
		FrontTemplate = rootFolder&"\FrontREPORT.TDR"
  Globaldim("DataTemplate")
		DataTemplate = rootFolder&"\ContentREPORT.TDR"  
  Globaldim("PicTemplate")
		PicTemplate = rootFolder&"\PictureREPORT.TDR"
  Globaldim("ConfigFile")
    ConfigFile =  rootFolder&"\TableStandards.txt"
  Globaldim("iResultSelected")
    iResultSelected = 0
  Globaldim("oGUIParam")
    set oGUIParam = new Reportparameter
    set oGUIParam.oDataFinder  = Navigator.ConnectDataFinder(Navigator.Display.CurrDataProvider.Name)
    oGUIParam.oDataFinder.Results.MaxCount = 1000
    dim sel(0)
    oGuiParam.SelectedShots = sel
    oGuiParam.SelectionCount = 0
  Globaldim("Version")
    Version = "V 1.0.0"
    
	'init "local" var
	sudName = "TestRecord.sud"
	searchDialog = "TestRecord"
  TableDialog = "SelectChannels"
  cfgPath = rootFolder&"\TableStandards.txt"
  Call ScriptCmdAdd(rootFolder&"\"&functionLibrary)
end sub
'-------------------------------------------------------------------------------



'-------------------------------------------------------------------------------
'summary: Open dialog form the .sud file.
'parameter: none
'output: none
Sub OpenDialog()
  dim DialogResponse
  call Readconfig(cfgpath)
  
  'only run through de programm only when the datafinder is bacqui
  DialogResponse = "IDOK"
  if Navigator.Display.CurrDataProvider.Name = "acqui@CH-W-PTHX001" then
      'do while (DialogResponse = "IDOK")
        if Suddlgshow(searchDialog,rootFolder&"/"&sudName,oGUIParam) ="IDOk" then
          Call SearchForTestShifts()
	      end if
      'loop
  else
    call msgbox("WRONG DATAFINDER!"&vblf&"Use BACQUI DATAFINDER. Look in: "&rootFolder&"\Documentation.doc for help")
	end if
  
end sub
'-------------------------------------------------------------------------------



'-------------------------------------------------------------------------------
class ReportParameter
  Dim oDataFinder               'The DataFinder used (use a Server Edition)
  Dim SelectedShots             'shots array which were selected
  Dim SelectionCount            'Number of Shots
  
  Dim StandardTable             'Standard Datatables for Tabledialog
  Dim StandardCount
  Dim TestShiftStart
  dim TestShiftEnd
  Dim TestShift
  Dim TestShiftCount
  
  Dim TestShiftProperties
  Dim Propertiescount
  
  Dim PDFPath                   'viariables for SetPicture Dialog
  Dim PDFHeading
  Dim PDFCount
end class
'-------------------------------------------------------------------------------



'-------------------------------------------------------------------------------
'summary: Open dialog form the .sud file.
'parameter: none
'output: none
Sub ReadConfig(cfgfile)
  Dim fso
  Dim txt
  Dim line
  Dim CurrentStandard:CurrentStandard = ""
  Dim CFGStandard
  Dim CFGParameter
  Dim table()
  Dim Tablecount:Tablecount = -1
  Dim Standards()
  Dim StandardCount
  dim debug, i
  
  'init standards
  StandardCount = 0
  redim standards(StandardCount)
  set standards(standardcount) = new standard
  
  set fso = CreateObject("Scripting.FileSystemObject")
  set txt =  fso.OpenTextFile(cfgFile, 1, 0)
  
  'read file
  Do while txt.AtEndOfStream = false
    line = txt.ReadLine
    
    'skip empty and comment lines
    If "#"=left(line,1) Or " " = left(line,1) Then
    
    Elseif line <> "" Then
    	CFGStandard = Left(line,instr(line,",")-1)
    	line = mid(line,instr(line,",")+1)
    	CFGParameter = line
        
        if CurrentStandard = CFGStandard or  CurrentStandard = "" then
            Tablecount = tablecount + 1
            redim preserve table(tablecount)
            Table(tablecount) = CFGParameter
            
            CurrentStandard= CFGStandard
        Else
            Call Standards(Standardcount).init(Table,TableCount,CurrentStandard)
            Standardcount = Standardcount +1
            Redim preserve Standards(Standardcount)
            set Standards(Standardcount) = new Standard
            
            Tablecount = 0
            redim preserve table(tablecount)
            Table(tablecount) = CFGParameter
            CurrentStandard= CFGStandard
        End if
    End If
  loop
  
  Call Standards(Standardcount).init(Table,TableCount,CurrentStandard)
  Standardcount = Standardcount +1
  Redim preserve Standards(Standardcount)
  set Standards(Standardcount) = new Standard
  
  oGuiparam.StandardTable = Standards
  oGuiparam.StandardCount = Standardcount-1

  'add to reportparameters
  txt.Close
  Set fso = Nothing
End Sub
'-------------------------------------------------------------------------------



'-------------------------------------------------------------------------------
class Standard
    Dim StandardName
    Dim Table()
    Dim TableCount
    
    sub init(DataTable,DataTableCount,DataStandardName)
      dim i
      StandardName = DatastandardName
      TableCount = DataTableCount
      redim Table(Tablecount)
      
      for i = 0 to TableCount
        Table(i) = DataTAble(i)
      Next
    End sub
End Class
'-------------------------------------------------------------------------------



'-------------------------------------------------------------------------------
'searches through testshits and then ask the user how his datatable should look like
'it has to be in the same file as the startthread, otherwise it doesn't work properly
sub SearchForTestShifts()
    dim i
    dim j
    Dim identical:identical = true               'false if shot is not in the same testshift
    Dim serial
    dim Shotnr
    dim comparecount: comparecount = 9
    Dim comparePropNames(): redim comparePropNames(comparecount)
    Dim compareProp():redim compareProp(comparecount)
    Dim Remarks
    Dim PropertyCount:Propertycount = 0
    Dim Properties():Redim Properties(Propertycount)
    dim debug
    dim TestShift():redim TestShift(0)
    Dim TestShiftCount:TestShiftCount = 0
    
    comparePropNames(0) = "CLIENT"
    comparePropNames(1) = "MANUFACTURER"
    comparePropNames(2) = "TEST_OBJECT"
    comparePropNames(3) = "U_RATED__KV__"
    comparePropNames(4) = "I_SC_RATED__KA__"
    comparePropNames(5) = "F_RATED__HZ__"
    comparePropNames(6) = "KIND_OF_TEST"
    comparePropNames(7) = "STANDARD"
    comparePropNames(8) = "TEST_CURRENT__KA__"
    comparePropNames(9) = "F_RATED__HZ__"

    'init compareprop
    oGUIParam.SelectedShots(0).GetRBA()
    serial = Data.Root.ActiveChannelGroup.Properties.Item("Name").Value
    serial = left(serial,len(serial)-4)
    serial = mid(serial,2)
    serial = left(serial,instr(serial,"-")-1)
    Shotnr= Data.Root.ActiveChannelGroup.Properties.Item("Name").Value
    Shotnr= left(Shotnr,len(Shotnr)-4)
    Shotnr = mid(Shotnr,instr(shotnr,"-")+1)
    
    for j = 0 to comparecount
        if Data.Root.ActiveChannelGroup.Properties.Exists(comparePropNames(j)) then
            compareprop(j) = Data.Root.ActiveChannelGroup.Properties.Item(comparePropNames(j)).Value
        Else
            compareprop(j) = ""
        end if
    next
    
    call GetallProperties(Properties,PropertyCount)
    
    TestShift(0) = oGUIParam.SelectedShots(0).RBAPath
    oGuiparam.testshiftstart = oGUIParam.SelectedShots(0).RBAPath
    oGuiparam.testshiftend = oGUIParam.SelectedShots(0).RBAPath
    Data.Root.Clear()

    'check if shots have consistent testshift parameters. if not put the shots in a different pdf
    for i = 1 to oGUIParam.SelectionCount
        'get remarks value
        oGUIParam.SelectedShots(i).GetRBA()
        serial = Data.Root.ActiveChannelGroup.Properties.Item("Name").Value
        serial = left(serial,len(serial)-4)
        serial = mid(serial,2)
        serial = left(serial,instr(serial,"-")-1)
        debug = serial
        Remarks = Data.Root.ActiveChannelGroup.Properties.Item("REMARKS").Value
        Shotnr= Data.Root.ActiveChannelGroup.Properties.Item("Name").Value
        Shotnr= left(Shotnr,len(Shotnr)-4)
        Shotnr = mid(Shotnr,instr(shotnr,"-")+1)
        debug = Shotnr
        
        'get all properties
        call GetallProperties(Properties,PropertyCount)
        
        'check if compareprop are the same
        for j = 0 to comparecount
            if Data.Root.ActiveChannelGroup.Properties.Exists(comparePropNames(j)) then
                if compareprop(j) <> Data.Root.ActiveChannelGroup.Properties.Item(comparePropNames(j)).Value then
                    identical = false
                end if
            End if
        next

        if identical = false then
            'call gui
            oGuiparam.TestShiftProperties = Properties
            oGuiparam.Propertiescount = PropertyCount
            oGuiparam.TestShift = TestShift
            oGuiparam.TestShiftCount = TestShiftCount
            if suddlgshow("SetPicture",rootFolder & "/TestRecord.sud",oGUIParam) = "IDOk" then
            End if
            
            if suddlgshow(TableDialog,rootFolder&"/TestRecord.sud",oGUIParam) = "IDOk" then
            end if
            
            'reinit variables
            identical = true
            PropertyCount = 0
            oGuiparam.testshiftstart = oGUIParam.SelectedShots(i).RBAPath
            oGuiparam.testshiftend = oGUIParam.SelectedShots(i).RBAPath
            debug = oGuiparam.testshiftstart
            debug = oGuiparam.testshiftend
            TestShiftCount = -1
            redim TestShift(0)
            oGuiparam.TestShift = TestShift
            oGuiparam.TestShiftCount = TestShiftCount
            
            'reinit compareproperties
            oGUIParam.SelectedShots(i).GetRBA()
            for j = 0 to comparecount
                if Data.Root.ActiveChannelGroup.Properties.Exists(comparePropNames(j)) then
                    compareprop(j) = Data.Root.ActiveChannelGroup.Properties.Item(comparePropNames(j)).Value
                end if
            next
            i = i -1
            call Data.Root.Clear()
        else
            TestShiftcount = TestShiftCount + 1
            redim preserve TestShift(TestShiftCount)
            debug = oGUIParam.SelectedShots(i).RBAPath
            TestShift(TestShiftCount) = oGUIParam.SelectedShots(i).RBAPath
            oGuiparam.testshiftend = oGUIParam.SelectedShots(i).RBAPath
        end if
        
        Data.Root.Clear()  
    next
    
    oGuiparam.TestShiftProperties = Properties
    oGuiparam.Propertiescount = PropertyCount
    oGuiparam.TestShift = TestShift
    oGuiparam.TestShiftCount = TestShiftCount
    if suddlgshow("SetPicture",rootFolder & "/TestRecord.sud",oGUIParam) = "IDOk" then
    End if
    
    if suddlgshow(TableDialog,rootFolder&"/TestRecord.sud",oGUIParam) = "IDOk" then
    end if
End sub
'-------------------------------------------------------------------------------



'-------------------------------------------------------------------------------
call init()
call Opendialog()