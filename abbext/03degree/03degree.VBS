'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2018-10-08 13:50:48
'-- Author: Adrian Kress
'-- Comment: changed to accept only channels named exactly 'Dreigrad' or 'Nullgrad'
'-------------------------------------------------------------------------------
'Forces the explicit declaration of all the variables in a script.
Option Explicit
dim s_type, CHdrei, CHnull, spannung, group, my_channel_0_time, my_channel_0,my_channel_3_time, my_channel_3, i, dif3, dif0, nEvents 
dim Channel3Length, Channel3SR, Channel0Length, Channel0SR, timeCH, sGnameCH, active_group
Dim startSin, startTan, pp1, pp2, diff, lowV, highV, sGname, sCurveName, oDig00, sNumber
Dim sON,sOFF,sSINrec,sDECH,iCurves,iCurve,sFormula,sFormula1,sSymbols,sValues,sUnit,Pts,t0,dt
Dim A,D,O

if Data.Root.ChannelGroups.Count = 0 then
 msgbox("No data available. We will provide demo data for Dreigrad and Nullgrad pulses.")
 Data.Root.ChannelGroups.Add("Demo_data")
 'msgbox("active channel group: "&Data.Root.ActiveChannelGroup.Name)
end if
active_group = Data.Root.ActiveChannelGroup.Name 'Save active group for convert to waveform

group = "group"
if not Data.Root.ChannelGroups.Exists(group) then
  Data.Root.ChannelGroups.Add(group)
end if
Call UIAutoRefreshSet(False)
'msgbox("group created?")
'msgbox(Data.Root.ActiveChannelGroup.Name)

'>>> select your 'Spannung' here
spannung = 0.8

'>>> select your between "increasing", "decreasing" or "both"
s_type = "increasing"

'>>> Propose creation of channels Dreigrad and Nullgrad if both do NOT exist.
if NOT(Data.Root.ActiveChannelGroup.Channels.Exists("Dreigrad")) AND NOT(Data.Root.ActiveChannelGroup.Channels.Exists("Nullgrad")) then
'msgbox("Generate pulses!")
  call generate_pulses
end if 

'>>> get properties of channels Dreigrad and Nullgrad
if Data.Root.ActiveChannelGroup.Channels.Exists("Dreigrad") then
  set CHdrei   = Data.Root.ActiveChannelGroup.Channels("Dreigrad")
  Channel3Length = CHdrei.Properties("length").Value
  Channel3SR     = Data.Root.ActiveChannelGroup.Channels("Dreigrad").Properties("wf_increment").Value

'= Case ======= Dreigrad ==========
  ChnEventResultList = Null 
  ChnEventResultList = ChnEventDetectionSlope( , CHdrei, spannung, s_type)

'convert input waveforms into x- and y-channel
  Set my_channel_3_time = Data.Root.ChannelGroups(group).Channels.Add("Dreigrad_x", DataTypeChnFloat64)
  Set my_channel_3      = Data.Root.ChannelGroups(group).Channels.Add("Dreigrad"  , DataTypeChnFloat64)

  nEvents = UBound(ChnEventResultList,1)
'msgbox(nEvents)
  For i = 1 to nEvents  'UBound(ChnEventResultList,1)
    my_channel_3_time.Values(i) = ChnEventResultList(i,2)        'Fill channel with data
  Next

  dif3 = my_channel_3_time.Values(2) - my_channel_3_time.Values(1)

  call ChnLinGenImp("my_equi3",Channel3Length, 0, Channel3SR) 

  For i = 1 to nEvents-1
    dif3 = my_channel_3_time.Values(i+1) - my_channel_3_time.Values(i)
    my_channel_3.Values(i) = 1/dif3/120 'Division by 120, the number of teeth on 3 deg sensor wheel
  Next
'Linear mapping to create a channel which is equidistant to the existing Dreigrad and Nullgrad channels
  Call CHNMAPLINCALC(my_channel_3_time,my_channel_3,Data.Root.ActiveChannelGroup.Channels("my_equi3"),"Result_dreigrad",1,"f[bound.slope]",NOVALUE,"analogue")
  Call ChnToWfChn(Data.Root.ActiveChannelGroup.Channels("my_equi3"),Data.Root.ActiveChannelGroup.Channels("Result_dreigrad"))
  Call ChnDelete(Data.Root.ActiveChannelGroup.Channels("my_equi3"))
End if

'=== Case ======== Nullgrad ============
if Data.Root.ActiveChannelGroup.Channels.Exists("Nullgrad") then
  set CHnull   = Data.Root.ActiveChannelGroup.Channels("Nullgrad")
  Channel0Length = CHnull.Properties("length").Value
  'msgbox("Channel0Length="&Channel0Length)
  Channel0SR     = Data.Root.ActiveChannelGroup.Channels("Nullgrad").Properties("wf_increment").Value
  call ChnLinGenImp("my_equi0",Channel0Length, 0, Channel0SR) 

'--- Event detection ---
  ChnEventResultList = Null 
  ChnEventResultList = ChnEventDetectionSlope(, CHnull, spannung, s_type)
'convert input waveforms into x- and y-channel
  Set my_channel_0_time = Data.Root.ChannelGroups(group).Channels.Add("Nullgrad_x", DataTypeChnFloat64)
  Set my_channel_0      = Data.Root.ChannelGroups(group).Channels.Add("Nullgrad"  , DataTypeChnFloat64)

  nEvents = UBound(ChnEventResultList,1)
  if nEvents<3 then 
    msgbox("Number of Nullgrad pluses: "& nEvents &" Minimum number is 3, consider increase of recording time.")
  else
    For i = 1 to nEvents  'UBound(ChnEventResultList,1)
      my_channel_0_time.Values(i) = ChnEventResultList(i,2)        'Fill channel with data
    Next
 
    'dif0 = my_channel_0_time.Values(2) - my_channel_0_time.Values(1)

    'call ChnLinGenImp("my_equi0",Channel0Length, 0, Channel0SR) 
    For i = 1 to nEvents-1
      dif0 = my_channel_0_time.Values(i+1) - my_channel_0_time.Values(i)
      my_channel_0.Values(i) = 1/dif0  'Nullgrad does not need /120
    Next
  'Linear mapping to create a channel which is equidistant to the existing Dreigrad and Nullgrad channels
    Call CHNMAPLINCALC(my_channel_0_time,my_channel_0,Data.Root.ActiveChannelGroup.Channels("my_equi0"),"Result_nullgrad",1,"f[bound.slope]",NOVALUE,"analogue")
    Call ChnToWfChn(Data.Root.ActiveChannelGroup.Channels("my_equi0"),Data.Root.ActiveChannelGroup.Channels("Result_nullgrad"))
    Call ChnDelete(Data.Root.ActiveChannelGroup.Channels("my_equi0"))
  End if 'from else case of nEvents<3
End if 
'------------------------------
'msgbox("remove")
Call Data.Root.ChannelGroups.Remove(group)

'-------------------------------------------------------------------------------
Sub generate_pulses()
iCurves = 2
sGname = "Geno"
'-------------------------------------------------------------------------------
'Call Data.Root.Clear() ' don't do that if these channels shall be added and you might need other content
Call Data.Root.ChannelGroups(active_group).Activate()
highV =  1  '2
lowV  = -1  '0.5

 sON       = "(abs(0<sin((A+(ch(""t"")*D)) * 2*pi/120*Ch(""t"")+pp1)) * abs(0>sin((A+(ch(""t"")*D)) * 2*pi/120*Ch(""t"")+pp2)) + O)"
 sFormula1 = "sin( (A+(ch(""t"")*D)) * 2*pi*Ch(""t"")) * " & sON

for iCurve = 1 to iCurves
    call set3deg(iCurve,sSymbols,sValues,sUnit,Pts,t0,dt)
    call ChnLinGenimp("t"&iCurve,Pts,t0,dt,"s") ' generates x channel
    call ChnLinGenimp("t",Pts,t0,dt,"s") ' generates another copy of x channel, just for the calculate command
    'define sFormula for the calculate command
    sFormula =   "ch(""Geno-A" & iCurve & """)=" & sFormula1 ' prepare proper entry in data portal
    sCurveName = "ch(""Geno-A" & iCurve & """)"
    'msgbox sFormula
    call  calculate(sFormula,sSymbols,sValues,sUnit) ' calculate the formula
    
    'create square wave
    Set oDig00 = Data.Root.ActiveChannelGroup.Channels.Add("Digi-A" &iCurve, DataTypeFloat64)
    For i = 1 to Pts
      if Data.Root.ActiveChannelGroup.Channels("Geno-A" &iCurve)(i) > 0.5 then  ' 2.5 if analog input is used
         oDig00(i) = highV
      else
         oDig00(i) = lowV
      end if
    next
    'the x channel is now no longer needed
    call Chndelete("t")
next 

'Convert to waveforms (taken from Function_Library.vbs)
for i=1 to iCurves
  sNumber = Str(i,"d")
  timeCH = "t" & sNumber
   'Data.Root.ActiveChannelGroup.Channel("t" & sNumber)
  sGnameCH = sGname & "-A" & sNumber
  'msgbox(sGnameCH)
  'sCurveName = Data.Root.ActiveChannelGroup.Channels(sGnameCH) ' "/"&sGname&"-A"&sNumber
  Call ChnToWfChn(timeCH,sGnameCH,0,"WfXRelative") '... X,ChnNoStr,XChnDelete,WfXStartTimeMode 
next

'Convert digital channels to waveforms (taken from Function_Library.vbs)
sGname = "Digi"
for i=1 to iCurves
  sNumber = Str(i,"d")
  sCurveName = active_group & "/"&sGname&"-A"&sNumber
  Call ChnToWfChn("[1]/t"&sNumber,sCurveName,0,"WfXRelative") '... X,ChnNoStr,XChnDelete,WfXStartTimeMode 
next

'Rename the synthetic channels to whatever is needed.
Data.Root.ActiveChannelGroup.Channels("Digi-A1").Name = "Dreigrad"
Data.Root.ActiveChannelGroup.Channels("Digi-A2").Name = "Nullgrad"
End Sub

Sub set3deg(iCurve,sSymbols,sValues,sUnit,Pts,t0,dt)
'      sSymbols = array("Ampl" , "Start"     ,  "Stop"         , "pp1"     , "pp2"     , "sUnit", "O"   , "Pts"  , "t0" , "dt")
       sSymbols = array("C"    , "A"         ,  "D"            , "pp1"     , "pp2"     , "sUnit", "O"   , "Pts"  , "t0" , "dt")
select case iCurve
 Case 1 sValues = array("0"    , 6000        , -1000           , 0         , 0         , "V"    ,  1    , 15000 ,  0.0 , 5E-6)
 Case 2 sValues = array("0"    , 6000        , -1000           , 61/120*pi , 62/120*pi , "V"    ,  0    , 15000 ,  0.0 , 5E-6)

End Select
   sUnit        = array( "A"   ,      "1"    ,      "1"    ,        "A",    "A"    , "A"    , "A"   ,   "1",  "s",  "s")
' Ampl:  Scaling of all curves
' Start: phase angle of switch point from AC to linear decay
' Stop:  phase angle of switch point from linear decay to AC
' pp1:   slope factor of linear decay
' pp2:   offset of linear decay
Pts = sValues(7)
t0  = sValues(8)
dt  = sValues(9)
End Sub




