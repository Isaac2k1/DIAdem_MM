'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2011-10-14 12:27:23
'-- Author: Marco Mailand
'-- Comment: new definition with sinwave and linear decay
'-------------------------------------------------------------------------------
' 3 degree and zero degree pulses
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
'Call ScriptInclude("Function_Library")
Dim startSin, startTan, pp1, pp2, diff, lowV, highV, sGname, i, sCurveName, oDig00, sNumber
Dim sON,sOFF,sSINrec,sDECH,iCurves,iCurve,sFormula,sFormula1,sSymbols,sValues,sUnit,Pts,t0,dt
Dim A,D,O
'-------------------------------------------------------------------------------
iCurves = 2
sGname = "Geno"
'-------------------------------------------------------------------------------
Call Data.Root.Clear()
Call Data.Root.ChannelGroups.Add("DIAdem_reference").Activate()
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
    Set oDig00 = Data.Root.ChannelGroups(1).Channels.Add("Digi-A" &iCurve, DataTypeFloat64)
    For i = 1 to Pts
      if Data.Root.ChannelGroups(1).Channels("Geno-A" &iCurve)(i) > 0.5 then  ' 2.5 if analog input is used
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
  sCurveName = "[1]/"&sGname&"-A"&sNumber
  Call ChnToWfChn("[1]/t"&sNumber,sCurveName,0,"WfXRelative") '... X,ChnNoStr,XChnDelete,WfXStartTimeMode 
next

'Convert digital channels to waveforms (taken from Function_Library.vbs)
sGname = "Digi"
for i=1 to iCurves
  sNumber = Str(i,"d")
  sCurveName = "[1]/"&sGname&"-A"&sNumber
  Call ChnToWfChn("[1]/t"&sNumber,sCurveName,0,"WfXRelative") '... X,ChnNoStr,XChnDelete,WfXStartTimeMode 
next

'Rename the synthetic channels to whatever is needed.
Data.Root.ChannelGroups(1).Channels("Digi-A1").Name = "Dreigrad"
Data.Root.ChannelGroups(1).Channels("Digi-A2").Name = "Nullgrad"

Sub set3deg(iCurve,sSymbols,sValues,sUnit,Pts,t0,dt)
'      sSymbols = array("Ampl" , "Start"     ,  "Stop"         , "pp1"     , "pp2"     , "sUnit", "O"   , "Pts"  , "t0" , "dt")
       sSymbols = array("C"    , "A"         ,  "D"            , "pp1"     , "pp2"     , "sUnit", "O"   , "Pts"  , "t0" , "dt")
select case iCurve
 Case 1 sValues = array("0"    , 6000        , -1000           , 0         , 0         , "V"    ,  1    , 100000 ,  0.0 , 5E-7)
 Case 2 sValues = array("0"    , 6000        , -1000           , 61/120*pi  , 62/120*pi  , "V"    ,  0    , 100000 ,  0.0 , 5E-7)

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

