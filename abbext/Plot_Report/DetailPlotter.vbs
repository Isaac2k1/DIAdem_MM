Class DetailPlotter
	
	private oDetail
	
	
	Public Function initDetailPlot(shotName, shotTitle, shotDate, pageLogo, pageTitle)
		Set oDetail = new Detail
		
		oDetail.setShotName(shotName)
		oDetail.setShotTitle(shotTitle)
		oDetail.setShotDate(shotDate)
		oDetail.setPageLogo(pageLogo)
		oDetail.setPageTitle(pageTitle)
	End Function	
  
  Public Function setDetailSettings(lSettings)
		oDetail.setDetailSettings(lSettings)
  End Function
  
	Public Function createDetailPlot(channelName)
		oDetail.setChannelName(channelName)
		Call SUDDlgShow("Plot_Detail",CurrentScriptPath&"Plot_Report.sud",oDetail)
	End Function
	
End Class