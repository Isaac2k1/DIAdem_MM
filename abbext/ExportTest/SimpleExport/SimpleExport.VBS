'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2013-01-23 13:36:49
'-- Author: Alex Huang
'-- Comment: Simple exporting
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  dim outputFolder 
  outputFolder = "C:\Users\chalhua\Desktop\test export"
  
  dim shotName 
  shotName = "2526-0898"    'ie "2526-0121"  xxxx-xxxx 
  
  Dim Channels(6)
  Channels(1) = "VS"
  Channels(2) = "Travel Drive"
  Channels(3) = "P-Comp"
  Channels(4) = "P-Heiz"
  Channels(5) = "arc voltage"
  Channels(6) = "I-Shunt"
  
  dim samplingRate
  samplingRate = 0.0001
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  Call ScriptCmdAdd("C:\DIAdem\abbext\equix\equix.vbs")
  
  dim chnNum 
  chnNum = Ubound(Channels)
 
  dim shotFile
  shotFile = "l" & shotName & ".rbd"
  
  dim RbeFile
  RbeFile = "l" & shotName & ".rbe"
  
  Call Data.Root.Clear()
  LoadShot(shotFile)
  call CreateEquidistantChannel(shotFile, Channels, chnNum, samplingRate)
  call CreateSDFFile()

  
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
Function CreateEquidistantChannel(strGroup,strChannels(),numberOfChannels,samplingRate)
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
      if Data.Root.ActiveChannelGroup.Channels.Exists(strChannels(i) & "_ABA" ) = true Then
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
      iXYChNo(j,1)  = iXYChNo(j,0) +1
      j = j+1
    end if
    if Data.Root.ActiveChannelGroup.Channels.Exists(strChannels(i) & "_AB" ) = true Then
      iGroupindex(j) = ChnGroup(CNo( strGroup&"/"&strChannels(i) & "_AB" ))  'Group of the channel
      iXYChNo(j,0)  = CNo( strGroup&"/"&strChannels(i) & "_AB" )  'index of x channel
      iXYChNo(j,1)  = iXYChNo(j,0) +1
      j = j+1
    end if
    if Data.Root.ActiveChannelGroup.Channels.Exists(strChannels(i) & "_ABA" ) = true Then
      iGroupindex(j) = ChnGroup(CNo( strGroup&"/"&strChannels(i) & "_ABA" ))  'Group of the channel
      iXYChNo(j,0)  = CNo( strGroup&"/"&strChannels(i) & "_ABA" )  'index of x channel
      iXYChNo(j,1)  = iXYChNo(j,0) +1
      j = j+1
    end if
  Next
  
  Call equix(iGroupindex,iXYChNo,1,samplingRate)
End Function

Sub CreateSDFFile()
  dim file, line, fso
  dim i, j, k
  dim milliseconds

  Data.Root.ChannelGroups(ShotFile).Activate()
  Set fso = CreateObject("Scripting.FileSystemObject")  
  Set file = fso.CreateTextFile(outputFolder &"\" & shotName & ".sdf", True)
  
  line = "#Time_[s]: "
  for i = 1 to chnNum 'actualChnNr
    line = line & Replace(channels(i), " ", "_") & ": "
  next  
  call file.WriteLine(line)
  
    for i = 1 to chnlength(RbeFile&"/"&channels(1))
    
    milliseconds = (i-1) * samplingRate
    dim Number 
    Number = milliseconds
    dim sign
    for k = 0 to 100
       if abs(Number) >= 10 then
         Number = Number / 10
         sign = "+"
       elseif abs(Number) < 1 and Number <> 0 then
         Number = Number * 10
         sign = "-"
       else 
         exit for
       end if 
     next  
        
     Number = round ( Number , 3)
     Number = FormatNumber ( Number , 3)
     if k < 10 and k > 0.1 then
       sign = sign & "0"
     end if
        
     if k = 0 then
       sign = "+0"
     end if
        
    line = Number & "E" & sign & k & " "
    
    'write one complete line
    for j = 1 to chnNum

      Number = CHD(i,RbeFile&"/" & channels(j))
      if isnumeric ( Number) then
        for k = 0 to 100
         if abs(Number) >= 10 then
            Number = Number / 10
            sign = "+"
          elseif abs(Number) < 1 and Number <> 0 then
            Number = Number * 10
            sign = "-"
          else 
            exit for
          end if 
        next
          
        Number = round ( Number , 3)
        Number = FormatNumber ( Number , 3)
        if k < 10 and k > 0.1 then
          sign = sign & "0"
        end if
          
        if k = 0 then
          sign = "+0"
        end if
          
        line = line & Number & "E" & sign & k & " "
       else
        line = line & "0.000E+00 "
       end if
           
    next

    Trim(line)
    file.WriteLine(line)
    line = ""
  next 
  call file.Close()
End Sub