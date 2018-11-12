Class DataPlotter
	
	private lSettings
	private lChannelGroups
	
	private currentChannelGroup
	
	Public Sub Class_Initialize 
		Set lChannelGroups = CreateObject("System.Collections.ArrayList")
		Set lSettings = CReateObject("System.Collections.ArrayList")
	End Sub
	
	Public Function changeChannelGroup(groupName)
	
		'Put the Group Object back into the list
		If NOT IsEmpty(currentChannelGroup)  Then
			lChannelGroups.Add(currentChannelGroup)
		End If
		
		'load Group Object and remove from list
		If Data.Root.ChannelGroups.Exists(groupName) Then
			Set currentChannelGroup = lChannelGroups.Item(lChannelGroups.IndexOf(groupName))
			lChannelGroups.Remove(currentChannelGroup)
			Data.Root.ChannelGroups(groupName).Activate()
		Else ' when Group Object doesn't exist - create new	/ könnte eventuell noch verbessert(verschönert werden)
			Set currentChannelGroup = new ChannelGroups
			currentChannelGroup.setGroupName(groupName)
			
			If NOT Data.Root.ChannelGroups(groupName).Properties.Exists("TestDate")Then 
				currentChannelGroup.setGroupDate(InputBox("No Value for DATE_TIME found!"&Chr(13)&"Please enter Value","No Properties","NOVALUE"))
			Else
				currentChannelGroup.setGroupDate(Data.Root.ChannelGroups(groupName).Properties("TestDate").Value)
			End If
			
			If InStr(groupName,".") = 0 Then
				currentChannelGroup.setGroupTitle(groupName)
			Else
				currentChannelGroup.setGroupTitle(Left(groupName,InStr(groupName,".")-1))
			End If
			
		
			Data.Root.ChannelGroups.Add(groupName).Activate() ' what, maby remove
		End If
	End Function
	
	Public Function getCurrentGroupName()
		getCurrentGroupName = currentChannelGroup.getGroupTitle()
	End Function
	
	Public Function getCurrentGroupDate()
		getCurrentGroupDate = currentChannelGroup.getGroupDate()
	End Function
	
End Class