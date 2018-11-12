'------------------------ Author -------------------------------------------------------------------------------------------
''   Janick Schmid
'
'------------------------ Creation Date ------------------------------------------------------------------------------------
''   2014-03 to 2015-08
'
'------------------------ Description --------------------------------------------------------------------------------------
'
' >>> This class (QueryController.vbs) is part of the "Load Test" program to be used in Diadem
' >>> This class is called by Load_Test.SUD
' >>> DataFinder is a global variable created in the LoadTest GUI
' >>> All scripts and classes are located in the same folder - Load Test
' >>> The detailed description of all the functions is also located in the Load Test folder
'---------------------------------------------------------------------------------------------------------------------------
'
'
Class QueryController
	private oQuerySearch
	private lCondition
	private lSelectedShots
	private lSelectedChannels
	
	Public Sub Class_Initialize
		Set lCondition = CreateObject("System.Collections.ArrayList")
		Set lSelectedShots = CreateObject("System.Collections.ArrayList")
		Set lSelectedChannels = CreateObject("System.Collections.ArrayList")
		Set oQuerySearch = new QuerySearch
	End Sub
	
	Public Function initSearch(returnType)
		lCondition.Clear
		oQuerySearch.intitSearch(returnType)
	End Function
	
	Public Function clearSelectedShots()
		lSelectedChannels.Clear 
	End Function
  
	Public Function clearSelectionLists()
		lSelectedShots.Clear 
		clearSelectedShots()
	End Function
	
	Public Function createCondition(vType, vProperty, vOperator, vValue)
		dim oCondition
		
		Set oCondition = new Condition

		Call oCondition.setType(vType)
		Call oCondition.setProperty(vProperty)
		Call oCondition.setValue(vValue)
		Call oCondition.setOperator(vOperator)
		
		lCondition.Add(oCondition)
		
	End Function
  
	Public Function startSearch()
		If lCondition.Count < 1 Then
			MsgBox("No Conditions given, script stoppped! Pleas try again.")
			Exit Function
		End If
			Call oQuerySearch.addCondition(lCondition)
			Call oQuerySearch.searchFiles()
	End Function
	
	Public Function addShotSelection(shotName)
		lSelectedShots.Add(shotName)
	End Function
	
	Public Function addChannelSelection(channel)
		lSelectedChannels.Add(channel)
	End Function
  
	Public Function loadResults(headFile, channelFile)
		dim shotName, oElement
		 
		For each shotName in lSelectedShots
			lCondition.Clear
			oQuerySearch.intitSearch(eSearchFile)
      
      ' checks if it' s the 2017 version
			if (ProgramVersionName = 2017) then
        Call Controller.createCondition(eSearchFile,"Filename","=",shotName&channelFile)
      else
        Call Controller.createCondition(eSearchFile,"File name","=",shotName&channelFile)
      end if
      
			startSearch()
			
			If Data.Root.ChannelGroups.Exists(shotName) Then 'KEMA
				Data.Root.ChannelGroups.Remove(shotName)
			End If 
			
			
			oQuerySearch.loadResults()
			
			If Data.Root.ChannelGroups.Exists(shotName&".rba") Then
				Call GroupPropCopy(Data.Root.ChannelGroups(shotName&".rba").Properties("Index").Value, Data.Root.ChannelGroups(shotName&".rbd").Properties("Index").Value)
				Call Data.Root.ChannelGroups.Remove(shotName&".rba")
			End If
			
		Next  
		
	End Function
	
	Public Function removeUnselectedChannels(lChannelEnd)    
		dim lChannels, channel,shotName, i, channelEnd
		
		Set lChannels =  CreateObject("System.Collections.ArrayList")
		For i = 1 to LBChannelResults.Items.Count
			lChannels.Add(LBChannelResults.Items(i).Text)
		Next
		
		For each channel in lSelectedChannels
			lChannels.Remove(channel)
		Next
		
		For each shotName in lSelectedShots
			For each channel in lChannels
				channel = Replace(channel,"*","x")	'DIAdem converts * to x
				channel = Replace(channel,"/","\")	'DIAdem converts / to \
					If lChannelEnd.Count < 1 Then
						If Data.Root.ChannelGroups(shotName).Channels.Exists(channel) Then
							Data.Root.ChannelGroups(shotName).Channels.Remove(channel)
						End If
					End If
				For each channelEnd in lChannelEnd
					If Data.Root.ChannelGroups(shotName&".rbd").Channels.Exists(channel&channelEnd)Then
						Data.Root.ChannelGroups(shotName&".rbd").Channels.Remove(channel&channelEnd)
					End If
				Next
			Next
		  
		Next
	
	End Function
  
	Public Function convertToEquix()
		dim channelNumberXY(), groupIndex(), shotName, channel, i, arraySize
		
		For each shotName in lSelectedShots
			i = 0
			
			arraySize = Data.Root.ChannelGroups(shotName&".rbd").Channels.Count/2 'only x-Channel of x/y channels needed
			
			Redim groupIndex(arraySize-1)
			Redim channelNumberXY(arraySize-1,1)
			
			For channel = 1 to arraySize
				channelNumberXY(i,0) = CInt(Data.Root.ChannelGroups(shotName&".rbd").Channels(channel*2-1).GetReference(eRefTypeNumber))
				channelNumberXY(i,1) = CInt(Data.Root.ChannelGroups(shotName&".rbd").Channels(channel*2).GetReference(eRefTypeNumber))
				groupIndex(i) = CInt(Data.Root.ChannelGroups(shotName&".rbd").Properties("Index").Value)
				i = i+1
			Next
			
			' ABB user function
			Call equix(groupIndex,channelNumberXY,0,"") 
			
			Data.Root.ChannelGroups.Remove(shotName&".rbd")
		Next
    
	End Function
  
	Public Function removeEmptyGroups()
		dim group
		For each group in Data.Root.ChannelGroups
			If group.Channels.Count < 1 Then
				Data.Root.ChannelGroups.Remove(group.Name)
			End If
		Next
	End Function
	
	'to combine the Channel and the Time to waveform, because ABB SE data contains unnecessary Time channels.
	Public Function combineToWave()
		dim shot, i
		For each shot in lSelectedShots
			For i = 1 to Data.Root.ChannelGroups(shot).Channels.Count/2
			Call ChnToWfChn(Data.Root.ChannelGroups(shot).Channels.Item(i).GetReference(eRefTypeNameIndex),Data.Root.ChannelGroups(shot).Channels.Item(i+1).GetReference(eRefTypeNameIndex),1,"WfXRelative")
			Next
		Next
	End Function
  
End Class