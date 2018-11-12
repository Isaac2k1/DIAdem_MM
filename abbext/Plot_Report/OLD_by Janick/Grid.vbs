Class Grid
	
	private posX
	private posY
	private length
	private height
	private mleft
	private mtop
	private mbottom
	private mright
	
	Public Function setPosX(VposX)
		posX = VposX
	End Function
	
	Public Function setPosY(VposY)
		posY = VposY
	End Function
	
	Public Function setHeight(Vheight)
		height = Vheight	
	End Function
	
	Public Function setLength(Vlength)
		length = Vlength
	End Function
	
	Public Function calculateMissing()
		mleft = posX
		mright = 100 - posX - length
		mbottom = posY
		mtop = 100 - posY - height
	End Function
	
	Public Function getPosX()
		getPosX = posX
	End Function
	
	Public Function getPosY()
		getPosY = posY
	End Function
	
	Public Function getLength()
		getLength = length
	End Function
	
	Public Function getHeight()
		getHeight = height
	End Function
	
	Public Function getmLeft()
		getMleft = mleft
	End Function
	
	Public Function getmTop()
		getMtop = mtop
	End Function
	
	Public Function getmBottom()
		getMbottom = mbottom
	End Function
	
	Public Function getmRight()
		getMright = mright
	End Function
	
End Class