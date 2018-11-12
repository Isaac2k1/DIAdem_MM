'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 06/19/2006 15:57:24
'-- Author: Kaan Oenen
'-- Comment: Data Plugin for Rebadas report files
'-------------------------------------------------------------------------------
'History
'v1.0
'
'v1.1
'Date property in group properties are refined. " " characters are trimmed and it is assured that _
'it has the following structure: YYYY-MM-DD. The rule for building date is updated according to RBD Data plugin _
'FileSize property is added to Root properties.
'v1.2
'Error handling added to TIME property (determination of weather TIME has hour,min,sec info)
'v1.3
'Test file naming convention implemented. According to the decision taken at the meeting on 7th Sept 2006
'in Oerlikon, the test files will be named: 
'Lssss[ssss]-tttt.rbX   - L -> Labor Prefix: l, m or v
'                       - X -> File-Type: 'd' -> .rbd -> DATA- or 'a' -> .rba -> ASCII-File
'                       - ssss oder ssssssss -> short or long Serial-Number
'                       - tttt -> Test-Number
'v1.4
'Error handling added, see Function : "is_rba_format"
'"Time" property added as string (hh:mm:ss) to channelgroup properties
'v1.5
'DATE and TIME Property were saved as string variables. Now they are put in one DATE_TIME Field as datatype "datetime"

Option Explicit


Sub ReadStore(File)

  'debug
  'call startlog()
  'end debug

'Get the name of the file loaded
 Dim LoadedFileName,LoadedFileExt, ShotNo, SerialNo, Labor, intPos, sDataPlugin
      LoadedFileName = File.Info.FileName         'e.g: "l2220-0222"
      LoadedFileExt  = File.Info.Extension        'e.g: ".rba"
      sDataPlugin    = "rebadas rba"

 Call is_rba_format(File,LoadedFileName,LoadedFileExt,sDataPlugin)
      
      'Get last 4 characters of the file name and save it as shot no
      ShotNo = Right(LoadedFileName,4)
      'Convert ShotNo(string) into number
      ShotNo = Cdbl(ShotNo)
      
      'Get Laboratory name
      Labor = Mid(LoadedFileName,1,1)
        Select Case Labor
          Case "l" 
            Labor = "Leistung"
          Case "m"
            Labor = "Mechanik"
          Case "v"
            Labor = "Hochspannung"
        End Select 
      
      intPos = InStr(2, LoadedFileName, "-")
      
      If intPos = 6 Then
          SerialNo = Cdbl(Mid(LoadedFileName,2,4))
      Elseif intPos = 10 Then
          SerialNo = Cdbl(Mid(LoadedFileName,2,8))
      Else
          RaiseError("The file is not a report file")
      End if        

'----Data Plugin for Report Files---

  'Define the line feed constant in ASCII file
  File.Formatter.LineFeeds        = vbLf
  'Define if you want to ignore empty lines in ASCII file
  File.Formatter.IgnoreEmptyLines = true
  'Define if you want to trim " " character
  File.Formatter.TrimCharacters   = " "
  'Define the delimiter between values in your ASCII file
  File.Formatter.Delimiters       = ""
  'Define Thousand Separator
  File.Formatter.ThousandSeparator= "."
  'Define the time format of your ASCII file
  File.Formatter.TimeFormat      = "YYYY-MM-DD hh:mm:ss"
  
  
  'Use the name of the file loaded as ChannelGroup's name(Report name)
  Dim ChannelGroupReport : Set ChannelGroupReport = Root.ChannelGroups.Add(LoadedFileName&".rba")
  'Add the Labor typte to Channelgroup properties
  Call Root.Properties.Add("Labor", Labor )
  Call Root.Properties.Add("Serial No", SerialNo )
  Call Root.Properties.Add("Shot No", ShotNo )
  Call Root.Properties.Add("FileSize", File.Size) 
  
  Dim PropertyName, PropertyValue, I, PreSign, Unit, RealNumber,iDashPos1,iDashPos2,YYear,MMonth,DDay,HHour,MMin,SSec
  Dim iKolonPos1,iKolonPos2
  
  For I = 1 to 6
    File.SkipLine()
    PropertyName = File.GetCharacters(16)    
    PropertyName = UCase(Trim(PropertyName))
    PropertyValue= UCase(File.GetNextStringValue(eString))
    If PropertyValue <> "" Then 
        'check if date is at least 8 characters
        If PropertyName = "DATE" and Len(PropertyValue) >= 8   Then
          
            iDashPos1 = InStr(PropertyValue,"-")
            iDashPos2 = InStrRev(PropertyValue,"-")
         
            If iDashPos1=0 or iDashPos2= 0 Then
              PropertyValue = "Problem with Date: no - used to seperate Year-Month-Day"
              'Call WriteToLog(PropertyValue)
              'Call WriteToLog(PropertyValue&" ReportName: "&LoadedFileName)
            Else
            
              
              YYear     = TRim(Left(PropertyValue,iDashPos1-1))
              MMonth    = Trim(Mid(PropertyValue,iDashPos1+1,iDashPos2-iDashPos1-1))
              DDay      = Trim(Right(PropertyValue,Len(PropertyValue)-iDashPos2))
            
              If Isnumeric(YYear) Then
                If Len(YYear) = 2 and YYear < 70 Then
                  YYear = 2000+YYear
                ElseIf Len(YYear) = 2 and YYear <= 99 Then
                  YYear = YYear + 1900
                Elseif Len(YYear) = 4 Then
                  YYear = Yyear
                End if    
              Else  
                Raiseerror(YYear&" is not a valid year expression")  
              End if  
            
              If not Isnumeric(MMonth) Then
                Raiseerror(MMonth&" is not a numeric expression")
              Elseif Cdbl(MMonth) < 1 or Cdbl(MMonth) > 12 Then 
                Raiseerror(MMonth&"is not a valid month expression")     
              Elseif Len(MMonth)= 1 Then
                'Assure that MMonth is two characters
                MMonth = 0&MMonth
              End if
            
              If not Isnumeric(DDay) Then
                Raiseerror(DDay&" is not a numeric expression")
              Elseif Cdbl(DDay) < 1 or Cdbl(DDay) > 31 Then 
                Raiseerror(DDay&" is not a valid day expression")
              Elseif Len(DDay)= 1 Then
                'Assure that DDay is two characters
                DDay = 0&DDay
              End if
              
            End if
          
          'convert to Datetime datatype and add to properties
          if HHour <> "" AND YYear <> "" then
            Call ChannelGroupReport.Properties.Add("DATE_TIME", CDate(YYear&"-"&MMonth&"-"&DDay&" "&HHour&":"&MMin&":"&SSec))
          end if
        
          'do not add DATE property
          PropertyName = ""
          
        End if      
        
        If PropertyName = "TIME"    Then
          
          iKolonPos1 = InStr(PropertyValue,":")
          iKolonPos2 = InStrRev(PropertyValue,":")
            If iKolonPos1 = iKolonPos2 Then               'Supposedly there is only hour and min information 
                HHour      = Trim(Left(PropertyValue,iKolonPos1-1))
                MMin       = Trim(Mid(PropertyValue,iKolonPos1+1,Len(PropertyValue)-iKolonPos1))  
                SSec       = 0
            Else
                HHour      = Trim(Left(PropertyValue,iKolonPos1-1))
                MMin       = Trim(Mid(PropertyValue,iKolonPos1+1,iKolonPos2-iKolonPos1-1))
                SSec       = Trim(Right(PropertyValue,Len(PropertyValue)-iKolonPos2))
            End if

            If not Isnumeric(HHour) Then
                Raiseerror(HHour&" is not a numeric expression")
            Elseif Cdbl(HHour) < 0 or Cdbl(HHour) > 24 Then 
                Raiseerror(HHour&" is not a valid hour expression")     
            End if
          
            If not Isnumeric(MMin) Then
                Raiseerror(MMin&" is not a numeric expression")
            Elseif Cdbl(MMin) < 0 or Cdbl(HHour) > 59 Then 
                Raiseerror(MMin&" is not a valid minute expression")     
            End if

            If not Isnumeric(SSec) Then
                Raiseerror(SSec&" is not a numeric expression")
            Elseif Cdbl(SSec) < 0 or Cdbl(SSec) > 59 Then 
                Raiseerror(SSec&" is not a valid second(s) expression")     
            End if
          
          'function CreateTime() can be used when you have to allocate a time more accurate than Seconds.
          'Call Root.Properties.Add("datetime", CreateTime(Cint(YYear),Cint(MMonth),Cint(DDay),Cint(HHour),Cint(MMin),Cint(SSec),0,0,0))
		      Call Root.Properties.Add("datetime", CDate(YYear&"-"&MMonth&"-"&DDay&" "&HHour&":"&MMin&":"&SSec))
          
		      'convert to Datetime datatype and add to properties
          if HHour <> "" AND YYear <> "" then
            Call ChannelGroupReport.Properties.Add("DATE_TIME", CDate(YYear&"-"&MMonth&"-"&DDay&" "&HHour&":"&MMin&":"&SSec))
          end if
		      
          'do not add TIME property.
          PropertyName = ""
        End if
        
        'ensure "DATE_TIME" is of datatype "date"
        If "DATE_TIME" = PropertyName then
        
          'if diadem can convert Value into date, add property, else skip it
          if isDate(PropertyValue) then
            PropertyValue = CDate(PropertyValue)
          Else
            PropertyName = ""
          End if
        end if
      
      if PropertyName <> "" then
        Call ChannelGroupReport.Properties.Add(PropertyName, PropertyValue)
      end if
      
      File.SkipLine()
    Else
      Call ChannelGroupReport.Properties.Add(PropertyName, "NoValue")
      File.SkipLine()
    End if  
  Next  

  While (File.Position <> File.Size)
    
    PreSign = UCase(File.GetNextStringValue(eString))
    File.SkipLine()
    
    If Presign = "2" Then
      PropertyName = UCase(Trim(File.GetCharacters(16)))
      Unit =  UCase(Trim(File.GetCharacters(8)))
      PropertyName = PropertyName&"( "&Unit&" )"
      RealNumber = UCase(Trim(File.GetCharacters(8)))
        
        If not IsNumeric(RealNumber) Then
            PropertyValue = RealNumber
        Else
            PropertyValue = Cdbl(RealNumber)
        End If
    
    Elseif Presign = "1" Then  
        PropertyName = UCase(Trim(File.GetCharacters(16)))
        RealNumber = UCase(Trim(File.GetNextStringValue(eString)))
        If not IsNumeric(RealNumber) Then
            PropertyValue = RealNumber
        Else
            PropertyValue = Cdbl(RealNumber)
        End If
    Else  
      'Get first 16 char. of line and assign it to PropertyName
      PropertyName = File.GetCharacters(16)
      'Trim " " components of the String both from left and right sides
      PropertyName = UCase(Trim(PropertyName)) 
 
      'Get next String Value and assign it to PropertyValue
      PropertyValue= UCase(File.GetNextStringValue(eString))
        'Select Case PropertyName
        '  Case "DATE"
        '    PropertyValue = DateValue(PropertyValue)
        '  Case "TIME"
        '    PropertyValue = TimeValue(PropertyValue)
        'End Select
    End if

    'Add Property to the ChannelGroup
    Call ChannelGroupReport.Properties.Add(PropertyName, PropertyValue)
    File.SkipLine()
  Wend

'Call ChannelGroup.Channels.AddImplicitChannel("NoChannels", 1, 1, 5, eI16) 'For Report File we don't need to create Channels
'---------------------------------------------------

'-------------------------------------
'----End of Data Plugin for Report Files---

  'debug
  'call endlog()
  'end debug
End Sub

'-------------------------------------------------------------------------------
' Check whether the file has "rba" format 
' File            : Object to access the file
' LoadedFileName  : Filename without extension
' LoadedFileExt   : Extension of File
'-------------------------------------------------------------------------------
Public Function is_rba_format(File,sFileName,sFileExt,sPluginName)
  
  Dim prefix : prefix = Left(sFileName,1)
  
  If ( File.Size < 128 ) Then RaiseError(sFileName&sFileExt & " is not a "&sPluginname&" file ! File size is too small!")
  If sFileExt <> ".rba" Then Raiseerror(sFileName&sFileExt &" is not a "&sPluginName&" file !")
  If (prefix <>"l") and (prefix <>"m") and (prefix <>"v") Then Raiseerror(sFileName&sFileExt &" is not a "&sPluginName&" file ! Labor prefix is missing in filename")

End Function

dim fso
dim file

sub StartLog()
  Set fso = CreateObject("Scripting.FileSystemObject")
  set file = fso.CreateTextFile("C:\dev\DIAdem\TestRecort\Development\debug\Debug.txt")
  
  'call msgbox("ALIVE")
End sub

Public sub WriteToLog(sErrorString)
  'old debugging function
  'Dim intHandleNew, strTextNew, intErrorNew
  'Dim intHandleOld, strMyTextOld, intErrorOld
  'Dim fso
  'set fso = CreateObject("Scripting.FileSystemObject")
  '    fso.CopyFile "C:\Program Files\National Instruments\DIAdem 10.0\Addinfo\RPT_Log.txt", "C:\Program Files\National Instruments\DIAdem 10.0\Addinfo")
  'set fso = nothing

 'Call FileCopy("C:\Program Files\National Instruments\DIAdem 10.0\Addinfo\RPT_Log.txt","C:\Program Files\National Instruments\DIAdem 10.0\Addinfo\RPT_Log_old.txt")
  '  intHandleNew = TextFileOpen("C:\Program Files\National Instruments\DIAdem 10.0\Addinfo\RPT_Log.txt",tfCreate) 
  '  intTextNew= TextfileWriteLn(intHandleNew,sErrorString)
  '  intHandleOld = TextFileOpen("C:\Program Files\National Instruments\DIAdem 10.0\Addinfo\RPT_Log_old.txt",tfRead)
  'Do While Not TextFileEOF(intHandleOld) 
  '  strTextNew = TextFileReadLn(intHandleOld)
  '  intTextNew= TextfileWriteLn(intHandleNew,strTextNew)
  'loop 

  '  intErrorNew = TextFileClose(intHandleNew)
  '  intErrorOld = TextFileClose(intHandleOld)
  'call FileDelete("C:\Program Files\National Instruments\DIAdem 10.0\Addinfo\RPT_Log_old.txt", 1)
  'end old debugging function
  
  file.Writeline(sErrorString)
End sub



sub endLog()
  file.Close()
  set fso = nothing
End sub
