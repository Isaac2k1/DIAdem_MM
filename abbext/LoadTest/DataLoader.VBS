'Name: 			DataLoader.vbs
'Author: 		Janick Schmid
'Last Edited: 	04-06-2014
'by: 			Janick Schmid
'Version:		0.9
'Description:
'Load the Results from Data Finder into the data portal from DIAdem
'addShotProperty allows to create a new group in the data portal where 
'the new data will be saved
'######################################################################

Class DataLoader
	
	'Set group in data portal, create new if doesn't exist
	Public Function checkShotGroup(shotName)
		
		If NOT Data.Root.ChannelGroups.Exists(shotName&".rba") Then
			Data.Root.ChannelGroups.Add(shotName&".rba").Activate
		End If
		
	End Function
	'load the channels in the results which matches the parameter
	Public Function loadChannel(channelName)
		dim oChannel
		
		For each oChannel in DataFinder.Results
			If Left(oChannel.Name,InStr(oChannel.Name,"_")-1) = channelName Then' load channel_A and channel_Y
				Call Navigator.LoadData(oChannel)
			End if
		Next 
		
	End Function
	'loads all files from the results into the data portal
	Public Function loadResult()
		dim oElement
		
		For each oElement in DataFinder.Results
			Call Navigator.LoadData(oElement)
		Next
		
	End Function
	
End Class
