'-------------------------------------------------------------------------------
'-- Created:        2013-03-20 12:10:02
'-- Author:         R.Irion 
'-- Comment:        This script will integrate a channel
'--                 End of the integration intervall (current zero) must be selected by a crosshair-curve-cursor
'--                 Beginn of the integration intervall is selected by arcing-time
'-- 
'-------------------------------------------------------------------------------
Option Explicit

'declaration
dim ArcTime, ArcTimePoint
dim chn_name, chn_lenght, chn_samplerate, chn_abs
dim Cursor_1_X

'input (manual by user)
chn_name = InputBox("Enter name of current channel","Input","I-TO") 
ArcTime = InputBox("Enter arcing time [ms]","Input","") 
Cursor_1_X = View.ActiveSheet.Cursor.P1

'input (channel)
chn_lenght = Data.Root.ChannelGroups(GroupDefaultGet).Channels(chn_name).Properties("length").Value
chn_samplerate = Data.Root.ChannelGroups(GroupDefaultGet).Channels(chn_name).Properties("wf_increment").Value

'Convert arcing time from [ms] to [s]
ArcTime = ArcTime * 10^-3

'make a copy from original channel
Call ChnCopy("[" & GroupDefaultGet & "]/" & chn_name, "[" & GroupDefaultGet & "]/" & chn_name & "_original")

'calculate points to delete, round it to value without decimals
ArcTimePoint = Cursor_1_X - ArcTime / chn_samplerate
ArcTimePoint = round(ArcTimePoint,0)

'delete points after current zero and replace with 0
Call DataAreaDel("[" & GroupDefaultGet & "]/" & chn_name, Cursor_1_X, chn_lenght)
Call DataAreaInsert0("[" & GroupDefaultGet & "]/" & chn_name, Cursor_1_X, chn_lenght-Cursor_1_X+1)

'delete points before arcing time and replace with 0
Call DataAreaDel("[" & GroupDefaultGet & "]/" & chn_name, 1, ArcTimePoint)
Call DataAreaInsert0("[" & GroupDefaultGet & "]/" & chn_name, 1, ArcTimePoint)

'Calculate absolut value of the channel
Call Calculate ("Ch(""[" & GroupDefaultGet & "]/" & chn_name & """)=Abs(Ch(""[" & GroupDefaultGet & "]/" & chn_name & """))")

'create channel i(t)^2
Call ChnMul("[" & GroupDefaultGet & "]/" & chn_name,"[" & GroupDefaultGet & "]/" & chn_name,"[" & GroupDefaultGet & "]/" & chn_name & "^2")

'Integrate
Call ChnIntegrate("","[" & GroupDefaultGet & "]/" & chn_name, "[" & GroupDefaultGet & "]/" & chn_name & "_integrated-i(t)")
Call ChnIntegrate("","[" & GroupDefaultGet & "]/" & chn_name & "^2", "[" & GroupDefaultGet & "]/" & chn_name & "_integrated-i(t)^2")

'display values
Call MsgBoxDisp("Maximum [i(t)*s]: " & Str(Val(ChnPropGet("[" & GroupDefaultGet & "]/" & chn_name & "_integrated-i(t)" ,"maximum"))*10^3,"d.dd") & " As")
Call MsgBoxDisp("Maximum [i(t)^2*s]: " & Str(Val(ChnPropGet("[" & GroupDefaultGet & "]/" & chn_name & "_integrated-i(t)^2" ,"maximum")),"d.dd") & " MA^2s")

'delete not longer used channels
Call Data.Root.ChannelGroups(GroupDefaultGet).Channels.Remove(chn_name)
Call Data.Root.ChannelGroups(GroupDefaultGet).Channels.Remove(chn_name & "^2")
Data.Root.ChannelGroups(GroupDefaultGet).Channels(chn_name & "_original").Name = "" & chn_name &""
