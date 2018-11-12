'-------------------------------------------------------------------------------
'-- VBS-Script-File
'-- Created: 2009-08-07 
'-- Authors:	Rene Irion 
'				Jonas Schwammberger
'-- Version: 1.9

'-- Purpose: Function collection of DIAdem2Excel. Here to Reuse.
'-- History:
'     1.01 Alpha Bugfixing
'     1.02 Alpha Bugfixing
'     1.03 Alpha Hotfix for processing with a 6x Polynom
'     1.04 Alpha Hotfix and Bugfix for processing with a 6x Polynom
'     1.05 Alpha Polynom can have any level.
'     1.06 Alpha developing Smoothtravelcurve
'     1.07 Alpha developing Smoothtravelcurve
'-------------------------------------------------------------------------------
Option Explicit  'Erzwingt die explizite Deklaration aller Variablen in einem Script.

'Call ScriptCmdAdd(myFolders(0)& "..\equix\equix.vbs") 'Must be in the same Path as this File

'--------------------------------------------------------------------------------
'summary: Find and load a single Shot
'parameter shot: string variable which defines the shot to load.
'output: true if successful  
Function LoadShot(shot)
  Dim BoolSuccess             'true if function was successful
  Dim MyDataFinder            'data finder object
  Dim AdvancedQuery           'query object
  
  BoolSuccess = true          'false if an error occured in the function
  
  'Connect to DataFinder
  Set MyDataFinder = Navigator.ConnectDataFinder(Navigator.Display.CurrDataProvider.Name)
  Set AdvancedQuery=Navigator.CreateQuery(eAdvancedQuery)
  
  'Define the result type (file)
  AdvancedQuery.ReturnType=eSearchFile
  Call AdvancedQuery.Conditions.RemoveAll()
  
  'Search for Binary Data File "rbd"
  Call AdvancedQuery.Conditions.Add(eSearchFile,"fileName","=",shot)
  
  'Search only for "data"-Files #TW 11.05.2010
  Call AdvancedQuery.Conditions.Add(eSearchFile,"fullpath","=","*data*")
  
  MyDataFinder.Search(AdvancedQuery)

  'check if the datafinder has found some data
  If MyDataFinder.Results.Count = 0 Then
    BoolSuccess = false
  
  ElseIf MyDataFinder.Results.Count = 1 Then
    Navigator.LoadData(MyDataFinder.Results) 'Load a single rbd File
      
 'if there is more than one result, use the first
  ElseIf MyDataFinder.Results.Count > 1 Then
    Navigator.LoadData(MyDataFinder.Results.Item(1))
  End If
  
  Set MyDataFinder = nothing
  Set AdvancedQuery = nothing
  
  LoadShot = BoolSuccess
 End Function
'--------------------------------------------------------------------------------



'--------------------------------------------------------------------------------
'summary: Make Channels Equidistant
'parameter strChannels: array of channel strings
'paraneter numberOfChannels: number of channels in array
'parameter samplingRate:
'output true if successful
sub CreateEquidistantChannels(strGroup,strChannels(),numberOfChannels,samplingRate)
  dim iXYChNo       '2d array which contains all x and y channel names
  dim iGroupindex   'no idea
  dim i             'indexer
  dim channelNumb   'actual channel number existing in this group
  
  'activate this group
  Data.Root.ChannelGroups(strGroup).Activate()
  
  channelNumb = -1
  
  for i = 0 to numberOfChannels
      if Data.Root.ActiveChannelGroup.Channels.Exists(strChannels(i) & "_A" ) = true Then
        channelNumb = channelNumb +1
      end if
      if Data.Root.ActiveChannelGroup.Channels.Exists(strChannels(i) & "_AB" ) = true Then
        channelNumb = channelNumb +1
      end if
  next
  
  'get number of channels in the group
  Redim iXYChNo(channelNumb,1)
  Redim iGroupindex(channelNumb)

  dim j
  
  j = 0
  For i = 0 To numberOfChannels
  'if activechannelgroup is 
    if Data.Root.ActiveChannelGroup.Channels.Exists(strChannels(i) & "_A" ) = true Then
      iGroupindex(j) = ChnGroup(CNo( strGroup&"/"&strChannels(i) & "_A" ))  'Group of the channel
      iXYChNo(j,0)  = CNo( strGroup&"/"&strChannels(i) & "_A" )  'index of x channel
      iXYChNo(j,1)  = CNo( strGroup&"/"&strChannels(i) & "_Y" )  'index of y channel
      j = j+1
    end if
    if Data.Root.ActiveChannelGroup.Channels.Exists(strChannels(i) & "_AB" ) = true Then
      iGroupindex(j) = ChnGroup(CNo( strGroup&"/"&strChannels(i) & "_AB" ))  'Group of the channel
      iXYChNo(j,0)  = CNo( strGroup&"/"&strChannels(i) & "_AB" )  'index of x channel
      iXYChNo(j,1)  = CNo( strGroup&"/"&strChannels(i) & "_Y" )  'index of y channel
      j = j+1
    end if
  Next
  
  Call equix(iGroupindex,iXYChNo,1,samplingRate)
  
End sub
'--------------------------------------------------------------------------------



'--------------------------------------------------------------------------------
'summary: removes Inductive and Resistance error.
'  formula was copied form Frederik Lundqvist's excel sheet.
'parameter sourceChannel: string of the sourceChannel to remove inductiveresistance from
'parameter currentChannel: channel with the Current waveform
'parameter samplingRate:
'parameter Inductance: Inductance value to remove the error
'parameter Resistance: Resistance value to remove the error
sub RemoveInductivResistantParts(sourceChannel,currentChannel,samplingRate,Inductance, Resistance)
  Dim RowCounter
  Dim InductivePart
  Dim ResistivePart
  Dim inductanceRemoved
  dim offset 
  
  if chnLength(sourceChannel) > chnLength(currentChannel) then
    msgbox("not possible to remove inductive and resistant parts, because current channel contains not egnough datapoints.")
    exit sub
  end if

  if ChnPropGet(currentChannel,"tStart") < ChnPropGet(sourceChannel,"tStart") then
    offset = int(ChnPropGet(sourceChannel,"tStart") / samplingRate)
  else
    offset = 0
  end if
  
  'it starts with the second value, 
  for RowCounter = 2 to chnLength(sourceChannel)
  
    'inductance * (currentValue - lastCurrentvalue) / sampling rate
    InductivePart = Inductance * (CHD(RowCounter + offset,currentChannel) - CHD(RowCounter - 1 + offset,currentChannel)) / samplingRate
      
    'Resistance * currentValue
    ResistivePart = Resistance * CHD(RowCounter + offset,currentChannel)
    
    'write value to the new channel
    inductanceRemoved = CHD(RowCounter,sourceChannel) - InductivePart - ResistivePart 
    CHD(RowCounter,sourceChannel) = inductanceRemoved
  next
 
end sub
'--------------------------------------------------------------------------------



'--------------------------------------------------------------------------------
'summary: Crop the Channels according to the Trigger-Signal of the VS channel
'     vs channel gets an offset correction.
'parameter strGroup: string of the channelgroup
'parameter strChannels: array of channel strings
'parameter numberOfChannels: number of channels in array
'parameter duration: duration of the shot
'parameter threshold: 
'parameter samplingRate: number of samplings
'parameter correction: corrective offset
sub CropChannels(strGroup,strChannels(),numberOfChannels, duration, threshold,samplingRate,correctiveOffset)
  dim intStartLine    'start of the trigger signal rise
  dim intEndLine      'no idea
  dim i               'indexer variable of for loop
  dim startLinePos    'line of positive trigger signal rise
  dim startLineNeg    'line of negative trigger signal rise
  
  'activate this group

  Data.Root.ChannelGroups(strGroup).Activate()
  
  'offset correction
  'original: call chnoffset(strChannels(0),strChannels(0),0,"first value offset")
  'call chnoffset(strChannels(0),strChannels(0),0,"first value offset")
  
  'Find the rising edge of the VS channel
  startLinePos = ChnFind("Ch("""&strChannels(0)&""")>" & CStr(threshold))
  startLineNeg = ChnFind("Ch("""&strChannels(0)&""")<" & CStr(-1*threshold))
  
  'take the first rise found. The ChnFind() returns 0 if it didn't find anything
  If startLineNeg < 2 Then
    if startLinePos > 2 Then
      intStartLine = startLinePos
    else
      intStartLine = 21
    end if
  Elseif startLineNeg < startLinePos Then
    intStartLine = startLineNeg
  elseif startLinePos > 2 Then
    intStartLine = startLinePos
  else
    intStartLine= startLineNeg
  End if
  
  'try to tweak here if the exported area doesn't quite fit
  intStartLine = intStartLine - CLng(correctiveOffset / samplingRate) 'remove a corrective offset of 0.002s
  intEndLine = intStartLine + CLng(duration / samplingRate)
  
  For i = 0 To NumberOfChannels 
    dim curChannel  'current channel in string
    curChannel = strGroup&"/"&strChannels(i)
    
  'Crop the channel
    if Data.Root.ActiveChannelGroup.Channels.Exists(strChannels(i)) = true then
      If ChnLength(CNo(curChannel)) - intEndLine > 0 Then 
        Call DataBlDel(curChannel,intEndLine, ChnLength(CNo(curChannel)) - intEndLine) 'remove the data at the end
        'Else... no need to crop, as intEndLine is after the channel's end
      end if

      'remove the data at the beginning
      Call DataBlDel(curChannel,1, intStartLine)
    end if
  Next
  
End sub
'--------------------------------------------------------------------------------



'--------------------------------------------------------------------------------
'summary: multiplies a channel with a specific value
'parameter strGroup: string of the channelgroup
'parameter channel: channel name as string
'parameter multValue: 
Sub MultiplyChannel(strGroup,channel,multValue)
  dim i
  
  Data.Root.ChannelGroups.Item(strGroup).Activate()
  
  if Data.Root.ActiveChannelGroup.Channels.Exists(channel) = true Then
    for i = 1 to chnLength(strGroup&"/"&channel)
      CHD(i,strGroup&"/"&channel) = CHD(i,strGroup&"/"&channel) * multValue
    next
  end if
  
end sub
'--------------------------------------------------------------------------------



'--------------------------------------------------------------------------------
'summary: calculates the Travel curve from degree to meter
'parameter strGroup: string of the channelgroup
'parameter channel: channel name as string
'parameter poly: array with all polynom factors
'parameter polNumber: number of polynom factors in array
Sub CalculateTravel(strGroup,channel,poly())
  dim travelInMm
  dim degree
  dim i
  dim j
  dim polNumber
 ' dim debug                                                debug deactivated
  
  polNumber = ubound(poly)
  
  Data.Root.ChannelGroups(strGroup).Activate()
  
  if Data.Root.ActiveChannelGroup.Channels.Exists(channel) = true then
    if polNumber > 0 Then
      for i = 1 to chnLength(channel)
        degree = CHD(i,channel)
        travelInMm = poly(0)
      
        for j = 1 to polNumber-1
          travelInMm = travelInMm + poly(j) * degree^j
          'debug = poly(j) * degree^j
        next
        
        CHD(i,channel) = travelInMm
      next
    end if
  end if
  
end sub
'--------------------------------------------------------------------------------



'--------------------------------------------------------------------------------
'summary: returns the "powered" value
'parameter base:
'parameter exponent:
function  Power(base,exponent)
  dim i
  dim output
  output = base
  
  if exponent > 1 then
    for i = 2 to exponent
      output = output * base
    next
  Elseif exponent = 1 then
    output = base
  Else
    output = 1
  end if

  Power = output
end function
'--------------------------------------------------------------------------------




