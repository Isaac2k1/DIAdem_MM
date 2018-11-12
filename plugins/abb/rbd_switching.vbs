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

Option Explicit 
Const HDRBUFFERSIZE = 1024
const   PLUGINNAME      = "ABB_LV_RBD"
const   PLUGINLONGNAME  = "ABB_LV_RBD"
'-------------------------------------------------------------------------------
' Data Plugin to read data from a ABB_LV_RBD file
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
' Main entry point. This procedure is called to execute the script
'-------------------------------------------------------------------------------
Sub ReadStore(oFile)
  Dim   oChannelGroup,oFHDR,LastRead,BlockNumber,Offset,ADCBits,iChnHDR
  Dim   sgFilename : sgFilename = oFile.Info.Filename & oFile.Info.Extension
  Dim   oaChnHDRs(),lChnCount,oChn,oBlockX,oBlockY,oDAChnX,oDAChnY,YMax,YScale,YOffset
  Dim   aTUnit : aTUnit = Array(1,60,3600,3600e3,3600e6,3600e9,3600e12) 'h, min, s, ms, us, ns, ps
  Dim   TimebaseA,TimebaseB,TimebaseAe,TimebaseBe,iValue,RawValue,oXChn,oYChn,ScaledValue

  oFile.Formatter.ByteOrder = eBigEndian

  If ( Not IsABB_LV_RBDFormat(oFile,sgFilename,PLUGINLONGNAME) ) Then RaiseError(sgFilename & " is not a "&PLUGINLONGNAME&" file !")
  Set oChannelGroup = Root.ChannelGroups.Add(oFile.Info.FileName&"_data") 
  oFile.Position = 0
  Set oFHDR = new RBD_FILE_HEADER
  oFHDR.ReadFromFile(oFile)

  LastRead = HDRBUFFERSIZE - (CInt((HDRBUFFERSIZE-11-oFHDR.HeadLen)/6)*6+1)
  Call oChannelGroup.Properties.Add("LastRead", LastRead)
  Call Root.Properties.Add("FileSize", oFile.Size)    
  Call Root.Properties.Add("datetime", CreateTime(Cint(oFHDR.YYear),Cint(oFHDR.MMonth),Cint(oFHDR.DDay),Cint(oFHDR.HHour),Cint(oFHDR.MMin),Cint(oFHDR.SSec),0,0,0))

  oFile.Position = LastRead+1
  ReDim oaChnHDRs(oFHDR.TotA+oFHDR.TotD)
  lChnCount = 0
  For iChnHDR = 1 To oFHDR.TotA+oFHDR.TotD
    BlockNumber = oFile.GetNextBinaryValue(eI16)
    Offset      = oFile.GetNextBinaryValue(eI16)
    ADCBits     = oFile.GetNextBinaryValue(eI16)
    If ( 0 <> ADCBits ) Then 
      Set oaChnHDRs(lChnCount)         = new RBD_CHN_HEADER
      oaChnHDRs(lChnCount).ADCBits     = ADCBits
      oaChnHDRs(lChnCount).ChnStartPos = (HDRBUFFERSIZE*(BlockNumber-1))+((Offset+2)*4)
      oaChnHDRs(lChnCount).ChnNumber   = iChnHDR
      lChnCount = lChnCount + 1
    End If
  Next
  Call oChannelGroup.Properties.Add("TotalAvailableChannels", lChnCount)
  ReDim Preserve oaChnHDRs(lChnCount-1)
  For Each oChn in oaChnHDRs
    oFile.Position = oChn.ChnStartPos - 12
    oChn.ReadFromFile(oFile)
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
    oBlockX.BlockLength = oChn.ChnLength
    Set oDAChnX         = oBlockX.Channels.Add(oChn.ChannelName&"_X", eU32)
    oDAChnX.Formatter.Bitmask = (2^(32-oChn.ADCBits))-1
    ' There can be up to three different rates used 
    ' to acquire one channel. If there is more than 
    ' one, we have to read value by value
'    Dim tot,s1,s2
'    tot = oChn.ChnLength
'    s1  = oChn.Itb12   
'    s2  = oChn.Itb21
    If ( oChn.LastX <= oChn.Itb12 ) Then 
      'If oDAChnX.Name = "I-Shunt ~_X" Then
'
'      oDAChnX.Factor = TimebaseAv
'      oDAChnX.Offset = oChn.TimeStart
'      Else
      oDAChnX.Factor = TimebaseA
      oDAChnX.Offset = oChn.TimeStart
      'End if

      Set oXChn = oChannelGroup.Channels.AddDirectAccessChannel(oDAChnX)
    Else
      Set oXChn = oChannelGroup.Channels.Add(oChn.ChannelName&"_Xx", eR64)
      For iValue = 1 To oChn.ChnLength
        RawValue = CLng(oDAChnX.Values(iValue))
        If     ( RawValue <= oChn.Itb12 ) Then
          ScaledValue = oChn.TimeStart + RawValue * TimebaseA
        Elseif ( RawValue <= oChn.Itb21 ) Then
          ScaledValue = TimebaseAe + (RawValue-oChn.Itb12) * TimebaseB  
        Else
          ScaledValue = TimebaseBe + (RawValue-oChn.Itb21) * TimebaseA
        End If
        oXChn.Values(iValue) = ScaledValue
        'oXChn.Values(iValue) = RawValue      
      Next
    End If 
    ' process y-channel
    oFile.Position = oChn.ChnStartPos + oChn.Dlen
    Set oBlockY         = oFile.GetBinaryBlock()
    oBlockY.BlockWidth  = 4
    oBlockY.BlockLength = oChn.ChnLength
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
  Dim   oFHDR
  IsABB_LV_RBDFormat = False

  If ( oFile.Size < 128 ) Then RaiseError(sgFilename & " is not a "&sgPluginname&" file !")

  Set oFHDR = new RBD_FILE_HEADER
  oFHDR.ReadFromFile(oFile)

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
  
  Public SizeOfStructure, LastRead, HHour, MMin, SSec
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
  Public Function ReadFromFile(oFile)
    Dim  K,StartPosition__
    StartPosition__ = oFile.Position
    vKey           = oFile.GetNextBinaryValue(eI32)
    TotA           = oFile.GetNextBinaryValue(eI16)
    TotD           = oFile.GetNextBinaryValue(eI16)
    HeadLen        = oFile.GetNextBinaryValue(eI16)
    AcquiCode      = oFile.GetNextBinaryValue(eI16)
    PlotCode       = oFile.GetNextBinaryValue(eI16)
    TR_DefaultCode = oFile.GetNextBinaryValue(eI16)
    AcquiString    = oFile.GetCharacters(30)
    Plotstring     = oFile.GetCharacters(30)
    TR_String      = oFile.GetCharacters(30)
    Dday           = oFile.GetNextBinaryValue(eI16)
    Mmonth         = oFile.GetNextBinaryValue(eI16)
    Yyear          = oFile.GetNextBinaryValue(eI16)
    TimeString     = oFile.GetCharacters(8)
    Max_Time       = oFile.GetNextBinaryValue(eR64)
    SizeOfStructure = oFile.Position-StartPosition__
    ReadFromFile = True
    'Reference : rbs/app/gettestinfo.ftn (Lines 262-266)
      If YYear < 70 Then
        YYear = 2000 + YYear
      ElseIf YYear <= 99 Then
        YYear = YYear + 1900
      Else
        YYear = YYear
      End if
    'Check if Timestring is in proper format and divide it into HHour,MMin,SSec
    If (Mid(TimeString,3,1) = ":") and (Mid(TimeString,6,1)=":") Then
      HHour = Left(TimeString, 2)
      MMin  = Mid(TimeString,4,2)
      SSec  = Right(TimeString,2)
    Else
      RaiseError("Timestring has no valid format")
    End if

    AcquiString = Trim(AcquiString)
    Plotstring  = Trim(Plotstring)
    TR_String   = Trim(TR_String)
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
    Call oLevel.Properties.Add("TotP",ChnLength)
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
  Public ChnLength   ,LastX       ,Empty_1     ,Dlen        ,ChannelName 
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
  Public Function ReadFromFile(oFile)
    Dim  K,StartPosition__
    StartPosition__ = oFile.Position
    ChnLength    = oFile.GetNextBinaryValue(eI32)   ' Total Number of saved Points in a Channel
    LastX        = oFile.GetNextBinaryValue(eI32)   ' Maximum X-Value  (Sample Number)
    Empty_1      = oFile.GetNextBinaryValue(eI16)   ' N/A
    Dlen         = oFile.GetNextBinaryValue(eI16)   ' Length of Channel-Description
    ChannelName  = oFile.GetCharacters(16)   ' Name(s) of the channel(s)
    ChannelUnit  = oFile.GetCharacters(8)   ' Unit(s) of the channel(s)
    CalibF       = oFile.GetNextBinaryValue(eR32)   ' CalibrationFactor
    ScaleF       = oFile.GetNextBinaryValue(eR32)   ' Scaling-Factor Geber [Units/Volt]
    InputR       = oFile.GetNextBinaryValue(eR32)   ' Inputrange of TR in Volts
    InOffset     = oFile.GetNextBinaryValue(eI16)   ' Offset of Input in Promille
    Attenuator   = oFile.GetNextBinaryValue(eI16)   ' Attenenuator of the Channel as 1/xx
    TimeBase1    = oFile.GetNextBinaryValue(eR32)   ' Timebase1
    TimeBase2    = oFile.GetNextBinaryValue(eR32)   ' Timebase2
    TimeStart    = oFile.GetNextBinaryValue(eR64)   ' Start in relation to absolut time Zero
    Itb12        = oFile.GetNextBinaryValue(eI32)   ' Pointnumber for switching Timebases from 1 to 2
    Itb21        = oFile.GetNextBinaryValue(eI32)   ' Pointnumber for switching Timebases from 2 to 1
    TriggerPoint = oFile.GetNextBinaryValue(eI32)   ' Pointnumber where TR has been triggered
    UnitTb1      = oFile.GetNextBinaryValue(eByte)   ' Unit Timebase1 as [1..7] hour,min,s,ms,us,ns,ps
    UnitTb2      = oFile.GetNextBinaryValue(eByte)   ' Unit Timebase2 as [1..7] hour,min,s,ms,us,ns,ps
    SizeOfStructure = oFile.Position-StartPosition__
    ReadFromFile = True
    ChannelName = Trim(RemoveNonASCII(ChannelName))
    ChannelUnit = Trim(ChannelUnit)
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
    ChnLength    = NULL : LastX        = NULL : Empty_1      = NULL : Dlen         = NULL : ChannelName  = NULL
    ChannelUnit  = NULL : CalibF       = NULL : ScaleF       = NULL : InputR       = NULL : InOffset     = NULL
    Attenuator   = NULL : TimeBase1    = NULL : TimeBase2    = NULL : TimeStart    = NULL : Itb12        = NULL
    Itb21        = NULL : TriggerPoint = NULL : UnitTb1      = NULL : UnitTb2      = NULL
  End Function
End Class

