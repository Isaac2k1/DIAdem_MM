'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2014-06-04 16:06:29
'-- Author: MM
'-- Comment: copy_channel_names
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
'allocate new channels
call chnalloc ("Property Name",100,,datatypestring,,groupdefaultget)
call chnalloc ("Property Value",100,,datatypestring,,groupdefaultget)

' set targetgroup to default group
TargetGroupIndex = groupdefaultget
Dim intLoop, group_name
group_name=groupname(groupdefaultget)
' loop over all properties in the default group
For intLoop = 10 to 10 'GroupPropCount(TargetGroupIndex)
call msgbox("intLoop="&intLoop&" Property name="&group_name)
 'write name and value of the properties to the channels
  cht(intLoop,group_name&"/Property Name")=GroupPropNameGet(TargetGroupIndex ,intLoop)
  cht(intLoop,group_name&"/Property Value")=GroupPropGet(TargetGroupIndex , GroupPropNameGet(TargetGroupIndex ,intLoop))
next

Data.Root.ChannelGroups(1).Channels("S1M1C1").Name = "S1M1C1_I-coil_ON"
