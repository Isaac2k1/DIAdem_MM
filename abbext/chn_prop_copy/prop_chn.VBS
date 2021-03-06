'-------------------------------------------------------------------------------
'-- VBS-Script-Datei
'-- Neu erstellt am 31.10.2008 08:55:39
'-- Ersteller: Mathias Knaak
'-- Kommentar: This function allocates two new channels "Property Name" and "Property Value"
'and writes all group properties to these channels
'-------------------------------------------------------------------------------
Option Explicit  'Erzwingt die explizite Deklaration aller Variablen in einem Skript.

'allocate new channels
call chnalloc ("Property Name",100,,datatypestring,,groupdefaultget)
call chnalloc ("Property Value",100,,datatypestring,,groupdefaultget)

' set targetgroup to default group
TargetGroupIndex = groupdefaultget
Dim intLoop, group_name
group_name=groupname(groupdefaultget)
' loop over all properties in the default group
For intLoop = 1 to GroupPropCount(TargetGroupIndex)
call msgbox("intLoop="&intLoop&" Property name="&group_name)
 'write name and value of the properties to the channels
  cht(intLoop,group_name&"/Property Name")=GroupPropNameGet(TargetGroupIndex ,intLoop)
  cht(intLoop,group_name&"/Property Value")=GroupPropGet(TargetGroupIndex , GroupPropNameGet(TargetGroupIndex ,intLoop))
next
