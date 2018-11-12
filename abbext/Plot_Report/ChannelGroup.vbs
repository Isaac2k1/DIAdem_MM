'contains information about a DIAdem channel group
Class ChannelGroup
	private groupName
	private groupTitle
	private groupDate
	
	Public Function setGroupName(VgroupName)
		groupName = VgroupName
	End Function
	
	Public Function setGroupDate(VgroupDate)
		groupDate = VgroupDate
	End Function
	
	Public Function setGroupTitle(VgroupTitle)
		groupTitle = VgroupTitle
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
	
End Class