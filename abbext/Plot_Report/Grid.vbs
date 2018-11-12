'The Grid contains the information for the background (the raster) and the size of the curve space
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
		calculateMargins()
		getMarginTop = marginTop
	End Function
	
	Public Function getMarginBottom()
		getMarginBottom = marginBottom
	End Function
	
	Public Function getMarginRight()
		calculateMargins()
		getMarginRight = marginRight
	End Function
	
	Private Function calculateMargins() 'Paper is 100% in width and height
		marginRight = 100 - marginLeft - width
		marginTop = 100 - marginBottom - height
	End Function
	
End Class