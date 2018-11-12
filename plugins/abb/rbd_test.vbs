'version history
'v1.0
'
'v1.1
'FileSize property added to root properties
'v1.2
'Custom channel properties are added to X-Channel too
'"Time" property is added to channel group properties
'datetime (StorageTime in File level) is added to Root(file) properties which makes it possible
'to search for the date of the file in Navigator
'v1.3
'Timebases are added to channel properties (both x,y channels)
'v1.4
'Important mistake found and corrected. look at line no 91-93

Option Explicit 
Const HDRBUFFERSIZE = 1024
const PLUGINNAME      = "ABB_LV_RBD"
const PLUGINLONGNAME  = "ABB_LV_RBD"
'-------------------------------------------------------------------------------
' Data Plugin to read data from a ABB_LV_RBD file
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
' Main entry point. This procedure is called to execute the script
'-------------------------------------------------------------------------------
Sub ReadStore(oFile)
  Dim   oChannelGroup,oFHDR,IndBeg,ChBegRec,ChInfBeg,ADCBits,iChnHDR
  Dim   sgFilename : sgFilename = oFile.Info.Filename & oFile.Info.Extension
  Dim   oaChnHDRs(),lChnCount,oChn,oBlockX,oBlockY,oDAChnX,oDAChnY,YMax,YScale,YOffset
  Dim   aTUnit : aTUnit = Array(1,60,3600,3600e3,3600e6,3600e9,3600e12) 'h, min, s, ms, us, ns, ps
  Dim   TimebaseA,TimebaseB,TimebaseAe,TimebaseBe,iValue,RawValue,oXChn,oYChn,ScaledValue

  oFile.Formatter.ByteOrder = eBigEndian
  
'Create Txt debugging File
Const ForWriting = 2
Const ForAppending = 8
   Dim fso,f
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.OpenTextFile("c:\rbd_dbg.txt", ForWriting, True)
   f.Close





  If ( Not IsABB_LV_RBDFormat(oFile,sgFilename,PLUGINLONGNAME) ) Then RaiseError(sgFilename & " is not a "&PLUGINLONGNAME&" file !")
  Set oChannelGroup = Root.ChannelGroups.Add(oFile.Info.FileName&"_data") 
  oFile.Position = 0
  Set oFHDR = new RBD_FILE_HEADER
  
  oFHDR.ReadFileHeader(oFile)
  
  IndBeg = HDRBUFFERSIZE-((HDRBUFFERSIZE-11-oFHDR.HeadLen)\6)*6+1
  
  Set f = fso.OpenTextFile("c:\rbd_dbg.txt", ForAppending, True)
  f.WriteLine("IndBeg: " &IndBeg)

  Call oChannelGroup.Properties.Add("IndBeg", IndBeg)
  Call Root.Properties.Add("FileSize", oFile.Size)    
  Call Root.Properties.Add("datetime", CreateTime(oFHDR.YYear,oFHDR.MMonth,_
       oFHDR.DDay,oFHDR.HHour,oFHDR.MMin,oFHDR.SSec,0,0,0))
  
  oFile.Position = IndBeg
  ReDim oaChnHDRs(oFHDR.TotA+oFHDR.TotD)
  lChnCount = 0
  For iChnHDR = 1 To oFHDR.TotA+oFHDR.TotD
    ChBegRec = oFile.GetNextBinaryValue(eI16)
    ChInfBeg = oFile.GetNextBinaryValue(eI16)
    ADCBits  = oFile.GetNextBinaryValue(eI16)
    
    f.WriteLine("Ch "&iChnHDR)
    f.WriteLine("ChBegRec: " &ChBegRec)
    f.WriteLine("ChInfBeg: " &ChInfBeg)
    f.WriteLine("ADCBits: " &ADCBits)
    f.WriteLine("ChStartPos: " &(HDRBUFFERSIZE*(ChBegRec-1)+4*ChInfBeg))
    
    f.WriteBlankLines 1

    If (ADCBits = 0 ) Then ADCBits = 12 End If
    If (ChBegRec > 0 ) Then
      Set oaChnHDRs(lChnCount)         = new RBD_CHN_HEADER
      oaChnHDRs(lChnCount).ADCBits     = ADCBits
      oaChnHDRs(lChnCount).ChnNumber   = iChnHDR
      oaChnHDRs(lChnCount).ChnStartPos = HDRBUFFERSIZE*(ChBegRec-1)+4*ChInfBeg
      lChnCount = lChnCount+1
    End If
  Next
  f.Close
  Call oChannelGroup.Properties.Add("TotalAvailableChannels",lChnCount)
  ReDim Preserve oaChnHDRs(lChnCount-1)
  For Each oChn in oaChnHDRs
    oFile.Position = oChn.ChnStartPos
    oChn.ReadChanHeader(oFile)
  Next
  '--------------------------------------------------
  ' group properties
  '--------------------------------------------------
  Call oFHDR.MapToTDM(oChannelGroup)
  '--------------------------------------------------
  ' process x and y-channels
  '--------------------------------------------------
  YMax = 65520
  For Each oChn in oaChnHDRs
    ' process x-channel
    TimebaseA  = oChn.TimeBase1*aTUnit(2)/aTUnit(oChn.UnitTb1-1)
    TimebaseB  = oChn.TimeBase2*aTUnit(2)/aTUnit(oChn.UnitTb2-1)   
    TimebaseAe = TimebaseA * oChn.Itb12 + oChn.Timestart
    TimebaseBe = TimebaseB * (oChn.Itb21-oChn.Itb12) + TimebaseAe
    oFile.Position = oChn.ChnStartPos + oChn.Dlen
    Set oBlockX         = oFile.GetBinaryBlock()
    oBlockX.BlockWidth  = 4
    oBlockX.BlockLength = oChn.TotP
    Set oDAChnX         = oBlockX.Channels.Add(oChn.ChannelName&"_X", eU32)
    oDAChnX.Formatter.Bitmask = (2^(32-oChn.ADCBits))-1
    ' There can be up to three different rates used 
    ' to acquire one channel. If there is more than 
    ' one, we have to read value by value
    
    'If ( oChn.TotP <= oChn.Itb12 ) Then  !!!!!!!!! This line is wrong !!!!!! Correct one is below 
    'modification v1.4 (in If statement : LastX instead of channel length)
    If ( oChn.LastX <= oChn.Itb12 ) Then 
      oDAChnX.Factor = TimebaseA
      oDAChnX.Offset = oChn.TimeStart
      Set oXChn = oChannelGroup.Channels.AddDirectAccessChannel(oDAChnX)
    Else
      Set oXChn = oChannelGroup.Channels.Add(oChn.ChannelName&"_Xx", eR64)
      For iValue = 1 To oChn.TotP
        RawValue = CLng(oDAChnX.Values(iValue))
        If     ( RawValue <= oChn.Itb12 ) Then
          ScaledValue = oChn.TimeStart + RawValue * TimebaseA
        Elseif ( RawValue <= oChn.Itb21 ) Then
          ScaledValue = TimebaseAe + (RawValue-oChn.Itb12) * TimebaseB  
        Else
          ScaledValue = TimebaseBe + (RawValue-oChn.Itb21) * TimebaseA
        End If
        oXChn.Values(iValue) = ScaledValue
      Next
    End If 
    ' process y-channel
    oFile.Position = oChn.ChnStartPos + oChn.Dlen
    Set oBlockY         = oFile.GetBinaryBlock()
    oBlockY.BlockWidth  = 4
    oBlockY.BlockLength = oChn.TotP
    Set oDAChnY         = oBlockY.Channels.Add(oChn.ChannelName&"_Y", eU16)
    oDAChnY.Formatter.Bitmask = Not (2^(16-oChn.ADCBits))-1
    YScale = oChn.InputR * oChn.ScaleF * oChn.CalibF * oChn.Attenuator / YMax
    YOffset= YMax * (1000.+oChn.InOffset) / 2000.0 +  0.5
    oDAChnY.Factor =YScale
    oDAChnY.Offset = YOffset*YScale*(-1)
    Set oYChn = oChannelGroup.Channels.AddDirectAccessChannel(oDAChnY)
    '--------------------------------------------------
    ' Channel properties
    '--------------------------------------------------
    oChn.MapToTDM(oYChn)
    oChn.MapToTDM(oXChn)
    
    Call oYChn.Properties.Add("YScale", YScale)
    Call oYChn.Properties.Add("YOffset", YOffset)
    Call oYChn.Properties.Add("SR1__Hz__", 1/TimebaseA)
    Call oYChn.Properties.Add("SR2__Hz__", 1/TimebaseB)
    Call oYChn.Properties.Add("TbA__s__", TimebaseA)
    Call oYChn.Properties.Add("TbB__s__", TimebaseB)
    
    Call oXChn.Properties.Add("Unit_String", "s")
    Call oXChn.Properties.Add("SR1__Hz__", 1/TimebaseA)
    Call oXChn.Properties.Add("SR2__Hz__", 1/TimebaseB)
    Call oXChn.Properties.Add("TbA__s__", TimebaseA)
    Call oXChn.Properties.Add("TbB__s__", TimebaseB)
  Next
End Sub
'-------------------------------------------------------------------------------
' Check whether the file has ABB_LV_RBD format 
' oFile         : Object to access the file
' sgFilename    : Filename without path but with extension
' sgPluginname  : Long name / description of the plugin. Typically this is 
'                 a descriptive name for the file format
'-------------------------------------------------------------------------------
Function IsABB_LV_RBDFormat(oFile,sgFilename,sgPluginname)
  Dim   oFHDR, ReadError
  'Set the initial value of the function to False
  IsABB_LV_RBDFormat = False
  
  'modification v1.5
  '----Start checking the format
  'If filesize is smaller than 2048 bytes abort loading file and write an error in log file
  If ( oFile.Size < 2048 ) Then RaiseError(sgFilename & " is not a "&sgPluginname&" file !")
  
  'Read the header of file
  Set oFHDR = new RBD_FILE_HEADER
  If (oFHDR.ReadFileHeader(oFile) <> True) Then RaiseError (oFHDR.ReadFileHeader(oFile))
  '----End of checking format
  
  'If there is no errors found on fileformat then function returns true as final value
  IsABB_LV_RBDFormat = True
End Function

'###############################################################################
'
'###############################################################################
Class RBD_FILE_HEADER
  '--------------------------------------------------
  ' Call MapToTDM with oLevel set to the appropriate
  ' root/group/channel object
  '--------------------------------------------------
  Public Function MapToTDM(oLevel)   ' map variables to TDM properties
    Call oLevel.Properties.Add("vKey", vKey)
    Call oLevel.Properties.Add("TotA", TotA)
    Call oLevel.Properties.Add("TotD", TotD)
    Call oLevel.Properties.Add("HeadLen", HeadLen)
    Call oLevel.Properties.Add("Acquistring", Acquistring)   
    Call oLevel.Properties.Add("Plotstring", Plotstring)
    Call oLevel.Properties.Add("TR_String", TR_String)
    Call oLevel.Properties.Add("Time", TimeString)
    Call oLevel.Properties.Add("Max_Time", Max_Time)
    Call oLevel.Properties.Add("AcquiCode", AcquiCode)
    Call oLevel.Properties.Add("PlotCode", PlotCode)
    Call oLevel.Properties.Add("TR_DefaultCode", TR_DefaultCode)
    Call oLevel.Properties.Add("Date", DDay&"."&MMonth&"."&YYear)
  End Function
  '--------------------------------------------------
  ' Definition of properties
  '--------------------------------------------------
  Public vKey          ,TotA          ,TotD          ,HeadLen       ,AcquiCode     
  Public PlotCode      ,TR_DefaultCode,AcquiString   ,Plotstring    ,TR_String     
  Public Dday          ,Mmonth        ,Yyear         ,TimeString    ,Max_Time      
  
  Public SizeOfStructure, IndBeg, HHour, MMin, SSec
  '--------------------------------------------------
  '@
  '--------------------------------------------------
  Private Sub Class_Initialize   ' Setup Initialize event.
    Call Initialize
  End Sub
  '--------------------------------------------------
  '@
  '--------------------------------------------------
  Private Sub Class_Terminate

  End Sub
  '--------------------------------------------------
  ' read structure from file 
  '--------------------------------------------------
  Public Function ReadFileHeader(oFile)
    Dim  K,StartPosition__
    StartPosition__ = oFile.Position
    vKey           = oFile.GetNextBinaryValue(eI32)
    call OpenFileAndAppend("vKey: " &vKey)
    TotA           = oFile.GetNextBinaryValue(eI16)
    call OpenFileAndAppend("TotA: " &TotA)
    TotD           = oFile.GetNextBinaryValue(eI16)
    call OpenFileAndAppend("TotD: " &TotD)
    HeadLen        = oFile.GetNextBinaryValue(eI16)
    'call OpenFileAndAppend("HeadLen: " &HeadLen)
    AcquiCode      = oFile.GetNextBinaryValue(eI16)
    call OpenFileAndAppend("AcquiCode: " &AcquiCode)
    PlotCode       = oFile.GetNextBinaryValue(eI16)
    call OpenFileAndAppend("PlotCode: " &PlotCode)
    TR_DefaultCode = oFile.GetNextBinaryValue(eI16)
    call OpenFileAndAppend("TR_DefaultCode: " &TR_DefaultCode)
    AcquiString    = oFile.GetCharacters(30)
    Plotstring     = oFile.GetCharacters(30)
    TR_String      = oFile.GetCharacters(30)
    Dday           = oFile.GetNextBinaryValue(eI16)
    call OpenFileAndAppend("Dday: " &Dday)
    Mmonth         = oFile.GetNextBinaryValue(eI16)
    call OpenFileAndAppend("Mmonth: " &Mmonth)
    Yyear          = oFile.GetNextBinaryValue(eI16)
    TimeString     = oFile.GetCharacters(8)
    call OpenFileAndAppend("TimeString: " &TimeString)
    Max_Time       = oFile.GetNextBinaryValue(eR64)
    call OpenFileAndAppend("Max_Time: " &Max_Time)
    SizeOfStructure = oFile.Position-StartPosition__
    ReadFileHeader = True
    'Reference : rbs/app/gettestinfo.ftn (Lines 262-266)
      If YYear < 70 Then
        YYear = 2000 + YYear
      ElseIf YYear <= 99 Then
        YYear = YYear + 1900
      End if
    call OpenFileAndAppend("YYear: " &YYear)
    'Modification v1.5
    'Check if Timestring is in proper format and divide it into HHour,MMin,SSec
    TimeString = Replace(TimeString," ","0")
    If(Mid(TimeString,3,1) <> ":" ) Then
     TimeString = "?!:?!:?!"
     HHour = 0
     MMin  = 0
     SSec  = 0
    ElseIf(InStrRev(TimeString,":") <> 6) Then
     TimeString = Left(TimeString,5) & ":00"
     HHour = cInt(Left(TimeString, 2))
     MMin  = cInt(Mid(TimeString,4,2))
     SSec  = 0
    Else
     HHour = cInt(Left(TimeString, 2))
     MMin  = cInt(Mid(TimeString,4,2))
     SSec  = cInt(Mid(TimeString,7,2))
    End If
    
    AcquiString = Trim(AcquiString)
    call OpenFileAndAppend("AcquiString: " &AcquiString)
    Plotstring  = Trim(Plotstring)
    call OpenFileAndAppend("Plotstring: " &Plotstring)
    TR_String   = Trim(TR_String)
    call OpenFileAndAppend("TR_String: " &TR_String)

    Call WriteBlankLines()

  End Function
  '--------------------------------------------------
  '@
  '--------------------------------------------------
  Public Function Initialize()
    Dim  K
    vKey           = NULL : TotA           = NULL : TotD           = NULL : HeadLen        = NULL : AcquiCode      = NULL
    PlotCode       = NULL : TR_DefaultCode = NULL : AcquiString    = NULL : Plotstring     = NULL : TR_String      = NULL
    Dday           = NULL : Mmonth         = NULL : Yyear          = NULL : TimeString     = NULL : Max_Time       = NULL
    HHour          = NULL : MMin           = NULL : SSec           = NULL

  End Function
  Public Function OpenFileAndAppend(Line)
   Const ForAppending = 8
   Dim fso,f
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.OpenTextFile("c:\rbd_dbg.txt", ForAppending)
   f.WriteLine(Line)
   f.Close
  End Function

  Public Function WriteBlankLines()
   Const ForAppending = 8
   Dim fso,f
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.OpenTextFile("c:\rbd_dbg.txt", ForAppending)
   f.WriteBlankLines 2
   f.Close
  End Function


End Class
'###############################################################################
'
'###############################################################################
Class RBD_CHN_HEADER
  '--------------------------------------------------
  ' Call MapToTDM with oLevel set to the appropriate
  ' root/group/channel object
  '--------------------------------------------------
  Public Function MapToTDM(oLevel)   ' map variables to TDM properties
    Call oLevel.Properties.Add("Unit_String", ChannelUnit)
    Call oLevel.Properties.Add("ADC Bits", ADCBits)
    Call oLevel.Properties.Add("DLen", DLen)
    Call oLevel.Properties.Add("TotP",TotP)
    Call oLevel.Properties.Add("LastX",LastX)
    Call oLevel.Properties.Add("ChannelStart",ChnStartPos)
    Call oLevel.Properties.Add("Attenuator",Attenuator)
    Call oLevel.Properties.Add("CalibF",CalibF)
    Call oLevel.Properties.Add("InOffset",InOffset)
    Call oLevel.Properties.Add("InputR",InputR)
    Call oLevel.Properties.Add("Itb12",Itb12)
    Call oLevel.Properties.Add("Itb21",Itb21)
    Call oLevel.Properties.Add("ScaleF",ScaleF)
    Call oLevel.Properties.Add("TimeStart",TimeStart)
    Call oLevel.Properties.Add("TriggerPoint",TriggerPoint)
    Call oLevel.Properties.Add("ChnNumber",ChnNumber)
    'Call oLevel.Properties.Add("YScale",YScale)
    'Call oLevel.Properties.Add("YOffset",YOffset)

  End Function
  '--------------------------------------------------
  ' Definition of properties
  '--------------------------------------------------
  Public TotP   ,LastX       ,Empty_1     ,Dlen        ,ChannelName 
  Public ChannelUnit ,CalibF      ,ScaleF      ,InputR      ,InOffset    
  Public Attenuator  ,TimeBase1   ,TimeBase2   ,TimeStart   ,Itb12       
  Public Itb21       ,TriggerPoint,UnitTb1     ,UnitTb2
  Public SizeOfStructure

  Public  ADCBits,ChnStartPos,ChnNumber
  '--------------------------------------------------
  '@
  '--------------------------------------------------
  Private Sub Class_Initialize   ' Setup Initialize event.
    Call Initialize
  End Sub
  '--------------------------------------------------
  '@
  '--------------------------------------------------
  Private Sub Class_Terminate

  End Sub
  '--------------------------------------------------
  ' read structure from file 
  '--------------------------------------------------
  Public Function ReadChanHeader(oFile)
    Dim  Line,K,StartPosition__,f
    Dim sTotP,sLastX,sDlen,sChannelName,sChannelUnit,sCalibF,sScaleF,sInputR,sInOffset
    Dim sAttenuator,sTimeBase1,sTimeBase2,sTimeStart,sItb12,sItb21,sTriggerPoint,sUnitTb1,sUnitTb2
    

    StartPosition__ = oFile.Position
    TotP            = oFile.GetNextBinaryValue(eI32)   ' Total Number of saved Points in a Channel
    LastX           = oFile.GetNextBinaryValue(eI32)   ' Maximum X-Value  (Sample Number)
    Dlen            = oFile.GetNextBinaryValue(eI32)   ' Length of Channel-Description
    ChannelName     = oFile.GetCharacters(16)   ' Name(s) of the channel(s)
    ChannelUnit     = oFile.GetCharacters(8)   ' Unit(s) of the channel(s)
    CalibF          = oFile.GetNextBinaryValue(eR32)   ' CalibrationFactor
    ScaleF          = oFile.GetNextBinaryValue(eR32)   ' Scaling-Factor Geber [Units/Volt]
    InputR          = oFile.GetNextBinaryValue(eR32)   ' Inputrange of TR in Volts
    InOffset        = oFile.GetNextBinaryValue(eI16)   ' Offset of Input in Promille
    Attenuator      = oFile.GetNextBinaryValue(eI16)   ' Attenenuator of the Channel as 1/xx
    TimeBase1       = oFile.GetNextBinaryValue(eR32)   ' Timebase1
    TimeBase2       = oFile.GetNextBinaryValue(eR32)   ' Timebase2
    TimeStart       = oFile.GetNextBinaryValue(eR64)   ' Start in relation to absolut time Zero
    Itb12           = oFile.GetNextBinaryValue(eI32)   ' Pointnumber for switching Timebases from 1 to 2
    Itb21           = oFile.GetNextBinaryValue(eI32)   ' Pointnumber for switching Timebases from 2 to 1
    TriggerPoint    = oFile.GetNextBinaryValue(eI32)   ' Pointnumber where TR has been triggered
    UnitTb1         = oFile.GetNextBinaryValue(eByte)   ' Unit Timebase1 as [1..7] hour,min,s,ms,us,ns,ps
    UnitTb2         = oFile.GetNextBinaryValue(eByte)   ' Unit Timebase2 as [1..7] hour,min,s,ms,us,ns,ps
    SizeOfStructure = oFile.Position-StartPosition__
    ReadChanHeader  = True
    ChannelName     = Trim(RemoveNonASCII(ChannelName))
    ChannelUnit     = Trim(ChannelUnit)
    

    sTotP            = "TotP: "             &TotP            
    sLastX           = "LastX: "            &LastX           
    sDlen            = "Dlen: "             &Dlen            
    sChannelName     = "ChannelName: "      &ChannelName     
    sChannelUnit     = "ChannelUnit: "      &ChannelUnit     
    sCalibF          = "CalibF: "           &CalibF          
    sScaleF          = "ScaleF: "           &ScaleF          
    sInputR          = "InputR: "           &InputR          
    sInOffset        = "InOffset: "         &InOffset        
    sAttenuator      = "Attenuator: "       &Attenuator      
    sTimeBase1       = "TimeBase1: "        &TimeBase1       
    sTimeBase2       = "TimeBase2: "        &TimeBase2       
    sTimeStart       = "TimeStart: "        &TimeStart       
    sItb12           = "Itb12: "            &Itb12           
    sItb21           = "Itb21: "            &Itb21           
    sTriggerPoint    = "TriggerPoint: "     &TriggerPoint    
    sUnitTb1         = "UnitTb1: "          &UnitTb1         
    sUnitTb2         = "UnitTb2: "          &UnitTb2         
    
    
    call OpenFileAndAppend(sTotP)
    call OpenFileAndAppend(sLastX)
    call OpenFileAndAppend(sDlen)
    call OpenFileAndAppend(sChannelName)
    call OpenFileAndAppend(sChannelUnit)
    call OpenFileAndAppend(sCalibF)
    call OpenFileAndAppend(sScaleF)
    call OpenFileAndAppend(sInputR)
    call OpenFileAndAppend(sInOffset)
    call OpenFileAndAppend(sAttenuator)
    call OpenFileAndAppend(sTimeBase1)
    call OpenFileAndAppend(sTimeBase2)
    call OpenFileAndAppend(sTimeStart)
    call OpenFileAndAppend(sItb12)
    call OpenFileAndAppend(sItb21)
    call OpenFileAndAppend(sTriggerPoint)
    call OpenFileAndAppend(sUnitTb1)
    call OpenFileAndAppend(sUnitTb2)
        
    call WriteBlankLines()        
    
  End Function
  '--------------------------------------------------
  ' RemoveNonASCII
  '--------------------------------------------------
  Private Function RemoveNonASCII(sgString)
    Dim sgBuffer  : sgBuffer = sgString
    Dim K
    For K = 0 To 31
      sgBuffer = Replace(sgBuffer,Chr(K)," ")
    Next
    RemoveNonASCII = sgBuffer
  End Function
  '--------------------------------------------------
  '@
  '--------------------------------------------------
  Public Function Initialize()
    Dim  K
    TotP    = NULL : LastX        = NULL : Empty_1      = NULL : Dlen         = NULL : ChannelName  = NULL
    ChannelUnit  = NULL : CalibF       = NULL : ScaleF       = NULL : InputR       = NULL : InOffset     = NULL
    Attenuator   = NULL : TimeBase1    = NULL : TimeBase2    = NULL : TimeStart    = NULL : Itb12        = NULL
    Itb21        = NULL : TriggerPoint = NULL : UnitTb1      = NULL : UnitTb2      = NULL
  End Function
  
'Public Function CreateFile()
'   Const ForWriting = 2
'   Dim fso,f
'   Set fso = CreateObject("Scripting.FileSystemObject")
'   Set f = fso.OpenTextFile("c:\rbd_dbg.txt", ForWriting, True)
'   f.Close
'End Function

Public Function OpenFileAndAppend(Line)
   Const ForAppending = 8
   Dim fso,f
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.OpenTextFile("c:\rbd_dbg.txt", ForAppending)
   f.WriteLine(Line)
   f.Close
End Function

Public Function WriteBlankLines()
   Const ForAppending = 8
   Dim fso,f
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.OpenTextFile("c:\rbd_dbg.txt", ForAppending)
   f.WriteBlankLines 2
   f.Close
End Function

End Class

