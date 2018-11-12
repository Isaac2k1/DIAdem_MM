'' ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'' Sub ReadStore(oFile)
   
'' Source-File -=> C:\DIAdem\plugins\abb\rbd.vbs
'' First Issue -=> 2006-08-10.15:47:04 ; Kaan Oenen
'' Last Update -=> 2007-07-31.12:29:04 ; a41081 -> M.Subotic
'' ======================================================================
'' Op.-System --=> nt4/w2k/wXP
'' Version -----=> 1.9.0
'' Reviewed ----=> 
'' ======================================================================
'' Description -=> ???
   
'' Rem.:-Naming of test Files: According to the decision taken at the me-
''       eting on 7th Sept 2006 in Oerlikon, the test files will be named
''       Lssss[ssss]-tttt.rbX with:
''       - L -> Labor Prefix: l, m or v
''       - X -> Type: 'd' -> .rbd -> DATA- or 'a' -> .rba -> ASCII-File
''       - ssss oder ssssssss -> short or long Serial-Number
''       - tttt -> Test-Number
   
'' Input  Parameter: 
''  oFile -> ???
   
'' Output Parameter: None
   
'' Used Subroutines: None
   
'' Error   Handling: None
'' ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'' Version  History: C:\DIAdem\plugins\abb\rbd.his
'' ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

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
'Important mistake found and corrected:
'if(oChn.TotP <= oChn.iTb12) changed to if(oChn.LastX <= oChn.iTb12)
'v1.5
'The variable names are adapted to "storeread.f" for the common naming
'of rebadas and dataplug-in variables
'Creation of a Log File for Debugging (only for developer mode)
'Checking the test file format had some changes as following:
'IsABB_LV_RBDformat function has changed its name to IsABB_RBDformat
'Minimum RBD file size allowed to be loaded by DataPlugin is 2048 bytes 
'If file header can not be read, loading file aborts
'v1.6
'When processing x and y values , File pointer is positioned at 
'oFile.Position = oChn.ChnBegPos+oChn.dLen+12 See Lines 113 and 145
'If dBug mode is not on, then the log file is not opened for writing Lines 49-50
'v1.7
'yOffset is truncated with Fix function and save as integer (see line 164)
'Group-Naming changed from xxx_data to xxx.rbd
'v1.8
'- Added Handling for new rebadas Data-Model with multiple xMax-Blocks
'- Naming of X-Channels cahnged from Name_X to Name_A or _AB or _ABA
'  depending on the Number of Time-Bases and/or Switching-Points


option explicit 
const hbSiz          = 1024 ' Record-Size: Bytes
const cbSiz          =  256 ' Record-Size: Int*4 Words = hbSiz/4
const PLUGINNAME     = "ABB_RBD"
const PLUGINLONGNAME = "ABB_RBD"
'-------------------------------------------------------------------------------
' Data Plugin to read data from a ABB_RBD file
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
' Open a Log-File for Debuging
'-------------------------------------------------------------------------------
dim fso,f,dBug : dBug = false
If dBug then
 set fso = CreateObject("Scripting.FileSystemObject")
 set f   = fso.OpenTextFile("C:\DIAdem\plugins\rbd-dbg.log",2,True) ' 2 -> Open for writing
end if
'-------------------------------------------------------------------------------
' Main entry point. This procedure is called to execute the script
'-------------------------------------------------------------------------------
sub ReadStore(oFile)
 dim oChannelGroup,oofHdr,IndBeg,ChnBegRec,ChnInfBeg,ADCBits,iChn,iHlp
 dim sgFilename : sgFilename = oFile.Info.Filename & oFile.Info.Extension
 dim ChnHDRs(),ChnCnt,oChn,oBlockX,oBlockY,oDAChnX,oDAChnY,YMax,YScale,YOffset
 dim tUnit : tUnit = Array(1,60,3600,3600e3,3600e6,3600e9,3600e12) 'h, min, s, ms, us, ns, ps
 dim TbA,TbB,TbAe,TbBe,xMax,xOfs,ePnt,xEnt,xPnt,oXChn,oYChn,tAbs
 
 oFile.formatter.ByteOrder = eBigEndian
 
 if(Not IsABB_RBDformat(oFile,sgFilename,PLUGINLONGNAME)) then
  RaiseError(sgFilename & " is not a "&PLUGINLONGNAME&" file !")
 end if
 
 set oChannelGroup = Root.ChannelGroups.Add(oFile.Info.FileName&".rbd")
 oFile.Position = 0
 set oofHdr = new RBD_FILE_HEADER
 oofHdr.ReadFileHeader(oFile)
 
 IndBeg = hbSiz-((hbSiz-11-oofHdr.HdrLen)\6)*6+1
 call oChannelGroup.Properties.Add("IndBeg",IndBeg)
 call Root.Properties.Add("FileSize",oFile.Size)
 call Root.Properties.Add("datetime",CreateTime(oofHdr.YYear,oofHdr.MMonth,_
      oofHdr.DDay,oofHdr.HHour,oofHdr.MMin,oofHdr.SSec,0,0,0))
 
 oFile.Position = IndBeg-1 ' Last Byte before IndBeg
 if(dBug) then
  f.WriteLine("IndBeg = " & IndBeg & " : fPos = " & oFile.Position & vbCRLF)
 end if
 redim ChnHDRs(oofHdr.TotA+oofHdr.TotD)
 ChnCnt = 0 ' Count used Channels
 
 for iChn=1 To oofHdr.TotA+oofHdr.TotD
  iHlp = oFile.Position  ' Save current File-Position
  ChnBegRec = oFile.GetnextBinaryValue(eI16)
  
  if(ChnBegRec<0) then   ' New Data-Model and ChnBegRec > 32767
   oFile.Position = iHlp ' Restore old File-Position
   ChnBegRec =-oFile.GetnextBinaryValue(eI32)
   iHlp      = oFile.GetnextBinaryValue(eI16) ' 100*ChnInfBeg+ADCBits
   ChnInfBeg = iHlp\100
   ADCBits   = iHlp-100*ChnInfBeg
   if(dBug) then
    f.WriteLine("ChNum = " & iChn & ": ChnBegRec = " & ChnBegRec _
    & " : iHlp = 100*ChnInfBeg+ADCBits = " & iHlp _
    & " : ChnInfBeg = " & ChnInfBeg & " -> ADCBits = " & ADCBits)
   end if
  else                   ' Old Data-Model or  ChnBegRec < 32768
   ChnInfBeg = oFile.GetnextBinaryValue(eI16)
   ADCBits   = oFile.GetnextBinaryValue(eI16)
   if(dBug and ChnBegRec>0) then
    f.WriteLine("ChNum = " & iChn & ": ChnBegRec = " & ChnBegRec _
    & " : ChnInfBeg = " & ChnInfBeg & " : ADCBits = " & ADCBits)
   end if
  end if
  
  if(ADCBits=0) then ADCBits = 12 end if
  
  if(ChnBegRec>0) then
   set ChnHDRs(ChnCnt)       = new RBD_CHN_HEADER
   ChnHDRs(ChnCnt).ADCBits   = ADCBits
   ChnHDRs(ChnCnt).ChnNumber = iChn
   ChnHDRs(ChnCnt).ChnBegRec = ChnBegRec
   ChnHDRs(ChnCnt).ChnBegPos = hbSiz*(ChnBegRec-1)+(ChnInfBeg-1)*4 ' Last Byte before ChnBeg
   ChnCnt = ChnCnt+1
  end if
 next
 
 call oChannelGroup.Properties.Add("TotC",ChnCnt)
 redim preserve ChnHDRs(ChnCnt-1)
 for each oChn in ChnHDRs
  oFile.Position = oChn.ChnBegPos
  if(dBug) then
   f.WriteLine(vbCRLF & "ChNum = " & oChn.ChnNumber _
   & ": ChnBegPos = " & oChn.ChnBegPos)
  end if
  oChn.ReadChanHeader(oFile)
 next
 
 ' Map Group-Properties
 call oofHdr.MapToTDM(oChannelGroup)
 
 ' Process X- and Y-Channels
 YMax = 65520
 for each oChn in ChnHDRs
  ' Process X-Channel
  TbA  = oChn.TbA*tUnit(2)/tUnit(oChn.uTbA-1) ' [s]
  TbB  = oChn.TbB*tUnit(2)/tUnit(oChn.uTbB-1) ' [s] 
  TbAe = TbA*oChn.iTb12+oChn.tStart
  TbBe = TbB*(oChn.iTb21-oChn.iTb12)+TbAe
  oFile.Position      = oChn.ChnBegPos+oChn.dLen+12
  set oBlockX         = oFile.GetBinaryBlock()
  oBlockX.BlockWidth  = 4
  oBlockX.BlockLength = oChn.TotP
  set oDAChnX         = oBlockX.Channels.Add(oChn.ChnName&"_A",eU32)
  oDAChnX.formatter.Bitmask = (2^(32-oChn.ADCBits))-1
  xMax = (2^(32-oChn.ADCBits))-1 'Highest Sample-Nr per Block
  xOfs = 0 'Offset for first xMax-Block
  
  ' There can be up to three different Sampling-Rates  used  to  acquire
  ' a Channel. If there is more than one, we have to read Value-by-Value
  
  if(oChn.LastX <= oChn.iTb12) then
   if(oChn.LastX <= xMax) then ' Single xMax-Block
    oDAChnX.Factor = TbA
    oDAChnX.Offset = oChn.tStart
    set oXChn = oChannelGroup.Channels.AddDirectAccessChannel(oDAChnX)
   else                         ' Multiple xMax-Blocks
    set oXChn = oChannelGroup.Channels.Add(oChn.ChnName&"_A",eR64)
    for xEnt = 1 To oChn.TotP
     ePnt = CLng(oDAChnX.Values(xEnt)) : xPnt = ePnt+xOfs
     oXChn.Values(xEnt) = oChn.tStart+xPnt*TbA
     if(ePnt=xMax) then xOfs = xOfs+xMax 'Offset for next xMax-Block
    next
   end if
  else
   if(oChn.iTb21 = oChn.iTb12) then
    set oXChn = oChannelGroup.Channels.Add(oChn.ChnName&"_AB",eR64)
   else
    set oXChn = oChannelGroup.Channels.Add(oChn.ChnName&"_ABA",eR64)
   end if
   
   if(oChn.LastX <= xMax) then ' Single xMax-Block
    for xEnt = 1 To oChn.TotP
     xPnt = CLng(oDAChnX.Values(xEnt))
     If    ( xPnt <= oChn.iTb12) then
      oXChn.Values(xEnt) = oChn.tStart+xPnt*TbA
     elseif( xPnt <= oChn.iTb21) then
      oXChn.Values(xEnt) = TbAe+(xPnt-oChn.iTb12)*TbB  
     else
      oXChn.Values(xEnt) = TbBe+(xPnt-oChn.iTb21)*TbA
     end if
    next
   else                         ' Multiple xMax-Blocks
    for xEnt = 1 To oChn.TotP
     ePnt = CLng(oDAChnX.Values(xEnt)) : xPnt = ePnt+xOfs
     If    ( xPnt <= oChn.iTb12) then
      oXChn.Values(xEnt) = oChn.tStart+xPnt*TbA
     elseif( xPnt <= oChn.iTb21) then
      oXChn.Values(xEnt) = TbAe+(xPnt-oChn.iTb12)*TbB  
     else
      oXChn.Values(xEnt) = TbBe+(xPnt-oChn.iTb21)*TbA
     end if
     if(ePnt=xMax) then xOfs = xOfs+xMax 'Offset for next xMax-Block
    next
   end if
  end if
  
  ' Process Y-Channel
  oFile.Position = oChn.ChnBegPos+oChn.dLen+12
  set oBlockY         = oFile.GetBinaryBlock()
  oBlockY.BlockWidth  = 4
  oBlockY.BlockLength = oChn.TotP
  set oDAChnY         = oBlockY.Channels.Add(oChn.ChnName&"_Y", eU16)
  oDAChnY.formatter.Bitmask = Not (2^(16-oChn.ADCBits))-1
  YScale = oChn.InRange * oChn.ScaleF * oChn.CalibF * oChn.Attenuator / YMax
  YOffset= Fix(YMax*(1000.0+oChn.InOffset)/2000.0+0.5)
  oDAChnY.Factor = YScale
  oDAChnY.Offset = YOffset*YScale*(-1)
  set oYChn = oChannelGroup.Channels.AddDirectAccessChannel(oDAChnY)
  
  ' Map common Chn-Properties
  call oChn.MapToTDM(oXChn)
  
  ' Map special xChn-Properties
  call oXChn.Properties.Add("ChnUnit","s")
  call oXChn.Properties.Add("SrA_Hz",1/TbA)
  call oXChn.Properties.Add("SrB_Hz",1/TbB)
  call oXChn.Properties.Add("TbA_s",TbA)
  call oXChn.Properties.Add("TbB_s",TbB)
  
  ' Map common Chn-Properties
  call oChn.MapToTDM(oYChn)
  
  ' Map special yChn-Properties
  call oYChn.Properties.Add("YScale",YScale)
  call oYChn.Properties.Add("YOffset",YOffset)
  call oYChn.Properties.Add("SrA_Hz",1/TbA)
  call oYChn.Properties.Add("SrB_Hz",1/TbB)
  call oYChn.Properties.Add("TbA_s",TbA)
  call oYChn.Properties.Add("TbB_s",TbB)
 next
end sub

'-------------------------------------------------------------------------------
' Check whether the file has ABB_RBD format 
' oFile         : Object to access the file
' sgFilename    : Filename without path but with extension
' sgPluginname  : Long name / description of the plugin. Typically this is 
'                 a descriptive name for the file format
'-------------------------------------------------------------------------------
function IsABB_RBDformat(oFile,sgFilename,sgPluginname)
 dim   oofHdr, ReadError
 'set the initial value of the function to False
 IsABB_RBDformat = False
 
 'modification v1.5
 '----Start checking the format
 'If filesize is smaller than 2048 bytes abort loading file and write an error in log file
 if(oFile.Size < 2048) then RaiseError(sgFilename & " is not a "&sgPluginname&" file !")
 
 'Read the header of file
 set oofHdr = new RBD_FILE_HEADER
 if(oofHdr.ReadFileHeader(oFile) <> True) then RaiseError("File Header can not be read !")
 '----End of checking format
 
 'If there is no errors found on fileformat then function returns true as final value
 IsABB_RBDformat = True
end function

'###############################################################################
'
'###############################################################################
Class RBD_FILE_HEADER
 '---------------------------------------
 ' Map Variables to TDM-Properties of the
 ' appropriate Level: root/group/channel
 '---------------------------------------
 Public function MapToTDM(oLevel)
  call oLevel.Properties.Add("vKey",vKey)
  call oLevel.Properties.Add("TotA",TotA)
  call oLevel.Properties.Add("TotD",TotD)
  call oLevel.Properties.Add("HdrLen",HdrLen)
  call oLevel.Properties.Add("AcquiString",AcquiString)   
  call oLevel.Properties.Add("PlotString",PlotString)
  call oLevel.Properties.Add("TrString",TrString)
  call oLevel.Properties.Add("MaxTime",MaxTime)
  call oLevel.Properties.Add("AcquiCode",AcquiCode)
  call oLevel.Properties.Add("PlotCode",PlotCode)
  call oLevel.Properties.Add("TrCode",TrCode)
  call oLevel.Properties.Add("TestDate",TestDate)
  call oLevel.Properties.Add("TestTime",TestTime)
 End function
 
 'Define Data-Portal Properties
 Public vKey,HdrLen,TotA,TotD
 Public AcquiCode,PlotCode,TrCode
 Public AcquiString,PlotString,TrString
 Public TestDate,TestTime,MaxTime
 
 'Define internal used Properties
 Public IndBeg,YYear,MMonth,DDay,HHour,MMin,SSec
 '--------------------------------------------------
 '@
 '--------------------------------------------------
 Private Sub Class_Initialize   ' setup Initialize event.
  call InitFileHeader()
 end sub
 
 '--------------------------------------------------
 '@
 '--------------------------------------------------
 Private Sub Class_Terminate
  ' Do nothing
 end sub
 
 '--------------------------------------------------
 ' read structure from file 
 '--------------------------------------------------
 Public function ReadFileHeader(oFile)
  dim ChnBegPos
  ChnBegPos      = oFile.Position
  vKey           = oFile.GetnextBinaryValue(eI32)
  TotA           = oFile.GetnextBinaryValue(eI16)
  TotD           = oFile.GetnextBinaryValue(eI16)
  HdrLen         = oFile.GetnextBinaryValue(eI16)
  AcquiCode      = oFile.GetnextBinaryValue(eI16)
  PlotCode       = oFile.GetnextBinaryValue(eI16)
  TrCode         = oFile.GetnextBinaryValue(eI16)
  AcquiString    = oFile.GetCharacters(30)
  PlotString     = oFile.GetCharacters(30)
  TrString       = oFile.GetCharacters(30)
  DDay           = oFile.GetnextBinaryValue(eI16)
  MMonth         = oFile.GetnextBinaryValue(eI16)
  YYear          = oFile.GetnextBinaryValue(eI16)
  TestTime       = oFile.GetCharacters(8)
  MaxTime        = oFile.GetnextBinaryValue(eR64)
  ReadFileHeader = True
  
  'Reference : rbs/app/gettestinfo.ftn
  if(YYear < 70) then
   YYear = 2000 + YYear
  elseif(YYear <= 99) then
   YYear = YYear + 1900
  end if
  
  'Create TestDate and force it to YYYY-MM-DD format
  if(MMonth>9) then
   TestDate = YYear & "-" & MMonth
  else
   TestDate = YYear & "-0" & MMonth
  end if
  
  if(DDay>9) then
   TestDate = TestDate & "-" & DDay
  else
   TestDate = TestDate & "-0" & DDay
  end if
  
  'Check the TestTime and force it to hh:mm:ss format
  TestTime = Replace(TestTime," ","0")
  if(Mid(TestTime,3,1) <> ":") then
   TestTime = "?!:?!:?!"
   HHour = 0
   MMin  = 0
   SSec  = 0
  elseif(InStrRev(TestTime,":") <> 6) then
   TestTime = Left(TestTime,5) & ":00"
   HHour = cInt(Left(TestTime, 2))
   MMin  = cInt(Mid(TestTime,4,2))
   SSec  = 0
  else
   HHour = cInt(Left(TestTime, 2))
   MMin  = cInt(Mid(TestTime,4,2))
   SSec  = cInt(Mid(TestTime,7,2))
  end if
  
  AcquiString = Trim(AcquiString)
  PlotString  = Trim(PlotString)
  TrString   = Trim(TrString)
 End function
  
 '--------------------------------------------------
 '@
 '--------------------------------------------------
 Public function InitFileHeader()
  vKey = NULL : TotA = NULL : TotD = NULL : HdrLen = NULL
  AcquiCode   = NULL : PlotCode   = NULL : TrCode   = NULL
  AcquiString = NULL : PlotString = NULL : TrString = NULL
  YYear       = NULL : MMonth     = NULL : DDay     = NULL
  HHour       = NULL : MMin       = NULL : SSec     = NULL
  TestDate  = NULL : TestTime = NULL : MaxTime  = NULL
 End function
End Class

'###############################################################################
'
'###############################################################################
Class RBD_CHN_HEADER
 '--------------------------------------------------
 ' call MapToTDM with oLevel set to the appropriate
 ' root/group/channel object
 '--------------------------------------------------
 Public function MapToTDM(oLevel)   ' map variables to TDM properties
  call oLevel.Properties.Add("ChnUnit",ChnUnit)
  call oLevel.Properties.Add("ADCBits",ADCBits)
  call oLevel.Properties.Add("dLen",dLen)
  call oLevel.Properties.Add("TotP",TotP)
  call oLevel.Properties.Add("LastX",LastX)
  call oLevel.Properties.Add("Attenuator",Attenuator)
  call oLevel.Properties.Add("CalibF",CalibF)
  call oLevel.Properties.Add("InOffset",InOffset)
  call oLevel.Properties.Add("InRange",InRange)
  call oLevel.Properties.Add("iTb12",iTb12)
  call oLevel.Properties.Add("iTb21",iTb21)
  call oLevel.Properties.Add("ScaleF",ScaleF)
  call oLevel.Properties.Add("tStart",tStart)
  call oLevel.Properties.Add("TriggerPoint",TriggerPoint)
  call oLevel.Properties.Add("ChnNumber",ChnNumber)
  'call oLevel.Properties.Add("YScale",YScale) : Will be mapped in Main: ReadStore(oFile)
  'call oLevel.Properties.Add("YOffset",YOffset)                  - " -
  '?'call oLevel.Properties.Add("SrA_Hz",1/TbA)
  '?'call oLevel.Properties.Add("SrB_Hz",1/TbB)
  '?'call oLevel.Properties.Add("TbA_s",TbA)
  '?'call oLevel.Properties.Add("TbB_s",TbB)
 End function
 
 'Define Data-Portal Properties
 Public TotP,LastX,dLen,ChnName,ChnUnit,ADCBits ',TbA,TbB
 Public CalibF,ScaleF,InRange,InOffset,Attenuator
 Public TbA,TbB,uTbA,uTbB,tStart
 Public iTb12,iTb21,TriggerPoint,ChnNumber
 
 'Define internal used Properties
 Public ChnBegRec,ChnBegPos
 
 '--------------------------------------------------
 '@
 '--------------------------------------------------
 Private Sub Class_Initialize   ' setup Initialize event.
  call InitChanHeader()
 end sub
 '--------------------------------------------------
 '@
 '--------------------------------------------------
 Private Sub Class_Terminate
  ' Do nothing
 end sub
 '--------------------------------------------------
 ' read structure from file 
 '--------------------------------------------------
 Public function ReadChanHeader(oFile)
  dim K,ChnBegPos
  
  ChnBegPos = oFile.Position
  TotP         = oFile.GetnextBinaryValue(eI32)   ' Total Number of saved Points in a Channel
  LastX        = oFile.GetnextBinaryValue(eI32)   ' Maximum X-Value  (Sample Number)
  if(dBug) then
   f.WriteLine("fPos = " & oFile.Position _
   & " : TotP = " & TotP & " : LastX = " & LastX)
  end if
  dLen         = oFile.GetnextBinaryValue(eI32)   ' Length of Channel-Description
  ChnName      = oFile.GetCharacters(16)          ' Name(s) of the channel(s)
  ChnUnit  = oFile.GetCharacters(8)           ' Unit(s) of the channel(s)
  CalibF       = oFile.GetnextBinaryValue(eR32)   ' CalibrationFactor
  ScaleF       = oFile.GetnextBinaryValue(eR32)   ' Scaling-Factor Geber [Units/Volt]
  InRange      = oFile.GetnextBinaryValue(eR32)   ' Input-Range of TR in Volts
  InOffset     = oFile.GetnextBinaryValue(eI16)   ' Offset of Input in Promille
  Attenuator   = oFile.GetnextBinaryValue(eI16)   ' Attenenuator of the Channel as 1/xx
  TbA    = oFile.GetnextBinaryValue(eR32)   ' TbA
  TbB    = oFile.GetnextBinaryValue(eR32)   ' TbB
  tStart    = oFile.GetnextBinaryValue(eR64)   ' Start in relation to absolut time Zero
  iTb12        = oFile.GetnextBinaryValue(eI32)   ' Pointnumber for switching Timebases from 1 to 2
  iTb21        = oFile.GetnextBinaryValue(eI32)   ' Pointnumber for switching Timebases from 2 to 1
  TriggerPoint = oFile.GetnextBinaryValue(eI32)   ' Pointnumber where TR has been triggered
  uTbA      = oFile.GetnextBinaryValue(eByte)  ' Unit TbA as [1..7] hour,min,s,ms,us,ns,ps
  uTbB      = oFile.GetnextBinaryValue(eByte)  ' Unit TbB as [1..7] hour,min,s,ms,us,ns,ps
  ChnName      = Trim(RemoveNonASCII(ChnName))
  ChnUnit  = Trim(ChnUnit)
  ReadChanHeader = True
 End function
 
 '--------------------------------------------------
 ' RemoveNonASCII
 '--------------------------------------------------
 Private function RemoveNonASCII(sgString)
  dim sgBuffer  : sgBuffer = sgString
  dim K
  for K = 0 To 31
   sgBuffer = Replace(sgBuffer,Chr(K)," ")
  next
  RemoveNonASCII = sgBuffer
 End function
 
 '--------------------------------------------------
 '@
 '--------------------------------------------------
 Public function InitChanHeader()
  dim  K
  TotP = NULL : LastX = NULL : dLen = NULL : ChnName = NULL : ChnUnit = NULL
  CalibF       = NULL : ScaleF       = NULL : InRange       = NULL : InOffset     = NULL
  Attenuator   = NULL : TbA    = NULL : TbB    = NULL : tStart    = NULL : iTb12        = NULL
  iTb21        = NULL : TriggerPoint = NULL : uTbA      = NULL : uTbB      = NULL
 End function
End Class
