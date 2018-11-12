Class Detail
	private shotName
	private shotTitle
	private shotDate
	private pageLogo
	private pageTitle
	private channelName
	private lDetailSettings
	
	Public Sub Class_Initialize 
		Set lDetailSettings = CreateObject("System.Collections.ArrayList")
	End Sub
	
	Public Function setShotName(VshotName)
		shotName = VshotName
	End Function
	
	Public Function getShotName()
		getShotName = shotName
	End Function
	
	
	Public Function setShotTitle(VshotTitle)
		shotTitle = VshotTitle
	End Function
	
	Public Function getShotTitle()
		getShotTitle = shotTitle
	End Function

	
	Public Function setShotDate(VshotDate)
		shotDate = VshotDate
	End Function
	
	Public Function getShotDate()
		getShotDate = shotDate
	End Function
	
	
	Public Function setPageLogo(VpageLogo)
		pageLogo = VpageLogo
	End Function
	
	Public Function getPageLogo()
		getPageLogo = pageLogo
	End Function
	
	Public Function setPageTitle(VpageTitle)
		pageTitle = VpageTitle
	End Function
	
	Public Function getPageTitle()
		getPageTitle = pageTitle
	End Function
	
	Public Function setChannelName(VchannelName)
		channelName = VchannelName
	End Function
	
	Public Function getChannelName()
		getChannelName = channelName
	End Function
	
	Public Function setDetailSettings(lSettings)
		Set lDetailSettings = lSettings
	End Function

	Public Function getDetailSettings()
		Set getDetailSettings = lDetailSettings
	End Function

End Class