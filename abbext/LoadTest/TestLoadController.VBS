'Name: 			TestLoadController.vbs
'Author: 		Janick Schmid
'Last Edited: 	03-07-2014
'by: 			Janick Schmid
'Version:		1.0
'Description:
'Save and update the selected shots and channels.
'controls QuerySearch and DataLoader to search and load data
'######################################################################

Class TestLoadController
	
	private oQuerySearch
	private oDataLoader
	private lSelectedShots
	private lSelectedChannels
	
	
	Public Sub Class_Initialize
		Set oQuerySearch = new QuerySearch
		Set oDataLoader = new DataLoader
		Set lSelectedShots = CreateObject("System.Collections.ArrayList")
		Set lSelectedChannels = CreateObject("System.Collections.ArrayList")
	End Sub
	
	Public Function initShotSearch(serialNumber,shotNumber,searchType)	
		
		Select Case searchType
		Case "single"
			Call oQuerySearch.singleShotQuery(serialNumber,shotNumber)
		Case "multi"
			Call oQuerySearch.multiShotQuery(serialNumber,shotNumber)
		End Select
		
	End Function
	
	Public Function addShotFilter(filterProperty, filterValue) 
		Call oQuerySearch.addShotFilter(filterProperty, filterValue)
	End Function
	
	Public Function startShotSearch()
		Call oQuerySearch.search()
	End Function
	
	Public Function searchChannel(shotNumber)
		oQuerySearch.searchChannel(shotNumber)
	End Function
	'RENAME FUNCTION
	Public Function updateSingleShot(shotName)
		lSelectedShots.Clear
		lSelectedShots.Add(shotName)
	End Function
	
	Public Function updateSelectedShots(Shots)
		Dim i
		
		lSelectedShots.Clear
		
		For i = 1 To Shots.Count
			lSelectedShots.Add(Shots.Item(i).Text)
		Next
		
	End Function
	
	Public Function updateSelectedChannels(Channels)
		Dim i
		
		lSelectedChannels.Clear
		
		For i = 1 To Channels.Count
			lSelectedChannels.Add(Channels.Item(i).Text)
		Next
		
	End Function
	
	Public Function loadChannels(Channels)
		Dim shotName, i, e
		
		
		updateSelectedChannels(Channels)
		
		For each shotName in lSelectedShots
		e = 1
			oQuerySearch.searchChannelFile(shotName)
			Call oDataLoader.loadResult()
			Call oQuerySearch.searchProperty(shotName)
			Call oDataLoader.loadResult()
			Call oDataLoader.checkShotGroup(shotName)
			For i = 1 to GroupChnCount(GroupIndexGet(shotName&".rbd"))
				For each channelName in lSelectedChannels
					If Left(Data.GetChannel(shotName&".rbd/["&i&"]").Name,InStr(Data.GetChannel(shotName&".rbd/["&i&"]").Name,"_")-1) = channelName Then
						Call ChnCopyExt(shotName&".rbd/["&i&"]", GroupIndexGet(shotName&".rba"), e)	
						e = e+1
					End If
				Next
			Next
			Call GroupDel(GroupIndexGet(shotName&".rbd"))
		Next
		
	End Function
	
	Public Function loadAllChannels(Channels)
		dim i, shotName
		
		updateSelectedChannels(Channels)
		
		For each shotName in lSelectedShots
			oQuerySearch.searchChannelFile(shotName)
			Call oDataLoader.loadResult()
			Call oQuerySearch.searchProperty(shotName)
			Call oDataLoader.loadResult()
			Call oDataLoader.checkShotGroup(shotName)
			For i = 1 to GroupChnCount(GroupIndexGet(shotName&".rbd"))
				Call ChnCopyExt(shotName&".rbd/["&i&"]", GroupIndexGet(shotName&".rba"), i)
			Next
			Call GroupDel(GroupIndexGet(shotName&".rbd"))
		Next
		
	End Function

	
End Class