Class Grid
	
	private width
	private height
	private marginLeft
	private marginTop
	private marginBottom
	private marginRight
	
	Public Function setHeight(Vheight)
		height = Vheight	
	End Function
	
	Public Function setWidth(Vwidth)
		width = Vwidth
	End Function
	
	Public Function setMarginLeft(VmarginLeft)
		marginLeft = VmarginLeft
	End Function
	
	Public Function setMarginBottom(VmarginBottom)
		marginBottom = VmarginBottom
	End Function
	
	Public Function getWidth()
		getWidth = width
	End Function
	
	Public Function getHeight()
		getHeight = height
	End Function
	
	Public Function getMarginLeft()
		getMarginLeft = marginLeft
	End Function
	
	Public Function getMarginTop()
		getMarginTop = marginTop
	End Function
	
	Public Function getMarginBottom()
		getMarginBottom = marginBottom
	End Function
	
	Public Function getMarginRight()
		getMarginRight = marginRight
	End Function
	
	Public Function calculateMargins()
		marginRight = 100 - marginLeft - width
		marginTop = 100 - marginBottom - height
	End Function
	
End Class