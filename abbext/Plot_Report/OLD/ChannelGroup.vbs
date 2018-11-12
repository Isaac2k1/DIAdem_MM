Class ChannelGroup
	private groupName
	private groupTitle
	private groupDate
	private lSelectedChannels
	
	Public Sub Class_Initialize 
		Set lSelectedChannels = CreateObject("System.Collections.ArrayList")
	End Sub
	
	Public Function setGroupName(VgroupName)
		groupName = VgroupName
	End Function
	
	Public Function setGroupDate(VgroupDate)
		groupDate = VgroupDate
	End Function
	
	Public Function setGroupTitle(vGroupTitle)
		groupTitle = vGroupTitle
	End Function
	
	Public Function getGroupTitle()
		getGroupTitle = groupTitle
	End Function
	
	Public Function getGroupName()
		getGroupName = groupName
	End Function
	
	Public Function getGroupDate()
		getGroupDate = groupDate
	End Function
	
	Public Function addSelectedChannel(channelName)
		lSelectedChannels.Add(channelName)
	End Function
	
	Public Function removeSelectedChannel(channelName)
		lSelectedChannels.Remove(channelName)
	End Function
	
	Public Function getSelectedChannels()
		getSelectedChannels =  lSelectedChannels
	End Function
	
End Class