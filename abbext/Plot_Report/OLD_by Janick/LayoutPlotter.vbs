Class LayoutPlotter
	
	private lRaster
	private oGrid
	private unit
	private paperSize
	private rasterSize
	private rasterAmount
	private plotDuration
	
	Public Sub Class_Initialize
		Set lRaster = CreateObject("System.Collections.ArrayList")
		Call PicDelete()
	End Sub
	
	Public Function setUnit(Vunit)
		unit = Vunit
	End Function
	
	Public Function getUnit()
		getUnit = unit
	End Function
	
	Public Function setPaperSize(VpaperSize)
		paperSize = VpaperSize
	End Function
	
	Public Function getPaperSize()
		getPaperSize = paperSize
	End Function
	
	Public Function getRasterAmount()
		getRasterAmount = rasterAmount
	End Function
	
	Public Function getPlotDuration()
		getPlotDuration = plotDuration
	End Function
	
	Public Function getRaster(id)
		Set getRaster = lRaster(id-1) 
	End Function
  
	Public Function getAllRaster()
		Set getAllRaster = lRaster
	End Function
	
	Public Function preparePlot()
		
		Call Report.Sheets.Remove("Layout")
		Call Report.Sheets.Insert("Layout", 1)
		Call Report.Sheets.Remove(2)
		Call GraphObjNew("2D-Axis","Grid")   'Creates a new 2D axis system
		Call GraphObjNew("FreeText","Place")
		Call GraphObjNew("FreeText","Laboratory")
		Call GraphObjNew("FreeText","TTest")
		Call GraphObjNew("FreeText","Test")
		Call GraphObjNew("FreeText","TDate")
		Call GraphObjNew("FreeText","Date")
		Call GraphObjNew("FreeText","Save")
	End Function

	Public Function setGrid(xPos,yPos,length, height)
		Set oGrid = new Grid
		Call oGrid.setPosX(xPos)
		Call oGrid.setPosY(yPos)
		Call oGrid.setHeight(height)
		Call oGrid.setLength(length)
		Call oGrid.calculateMissing()
		lRaster.Clear
	End Function
	
	Public Function drawGrid()
		dim height, posX, posY
		height = oGrid.getHeight()
		posX = oGrid.getPosX()
		posY = oGrid.getPosY()
		
		'Grid
		Call GraphObjOpen("Grid")
			'Position
			D2AxisTop        = oGrid.getmTop()      
			D2AxisRight       =oGrid.getmRight()
			D2AxisBottom     =oGrid.getmBottom()
			D2AxisLeft       =oGrid.getmLeft()
  
			'Layout
			D2AxisDisp(1) = "Grid"
			D2AxisDisp(2) = "Grid"
			D2AxisDispType = "Frame"
			D2AxisLineWidth = "min"
			D2AxisGridColor = "grey"
			
			'Scaling
			D2AxisScaledIn = "units per unit length"
			D2UseCommonXChn = TRUE
			
			'X-Axis
			Call GraphObjOpen(D2AxisXNam(1))
				D2AxisXBegin = 0
				D2AxisXEnd = 80
		
				'Scaling
				D2AxisXScaleType = "manual"
				D2AxisXMiniTick = 0
				D2AxisXTick = 1
				D2AxisXTickAuto = TRUE
				D2AxisXUnitPreset = "ms"
				D2AxisYOffOrigin = "AxisBegin"
				
				D2AxisXSize = 0
				D2AxisXTickType = "no Ticks"
			Call GraphObjClose(D2AxisXNam(1))
			
			'Y-Axis
			Call GraphObjOpen(D2AxisYNam(1))
				D2AxisYBegin = 0
				D2AxisYEnd = 20
		
				'Scaling
				D2AxisYScaleType = "manual"
				D2AxisYMiniTick = 0
				D2AxisYTick = 1
				D2AxisYTickAuto = FALSE
				D2AxisXOffOrigin = "AxisBegin"
				
				D2AxisYSize = 0
				D2AxisYTickType = "no Ticks"
			Call GraphObjClose(D2AxisYNam(1))
			
		Call GraphObjClose("Grid") 

		'Header Text
		Call GraphObjOpen("Place")
			TxtPosX = posX+1
			TxtPosY= 2+posY+height
			TxtTxt = "ABB"    
			TxtFont = "Arial"
			TxtSize = 4
			TxtBold = FALSE
			TxtRelPos = "rigth"
		Call GraphObjClose("Place")
			
		Call GraphObjOpen("Laboratory")
			TxtPosX = posX+9
			TxtPosY= 1.7+posY+height
			TxtTxt = "TESTING LABORATORY BADEN"    
			TxtFont = "Arial"
			TxtSize = 2.4
			TxtBold = TRUE
			TxtRelPos = "rigth"
		Call GraphObjClose("Laboratory")
		
		Call GraphObjOpen("TTest")
			TxtPosX = posX+30
			TxtPosY= 1.7+posY+height
			TxtTxt = "TEST:"    
			TxtFont = "Arial"
			TxtSize = 2
			TxtBold = FALSE
			TxtRelPos = "rigth"
		Call GraphObjClose("TTest")
		
		Call GraphObjOpen("Test")
			TxtPosX = posX+40
			TxtPosY= 1.7+posY+height
			TxtTxt = "VALUE"    
			TxtFont = "Arial"
			TxtSize = 2
			TxtBold = FALSE
			TxtRelPos = "rigth"
		Call GraphObjClose("Test")		
		
		Call GraphObjOpen("TDate")
			TxtPosX = posX+50
			TxtPosY= 1.7+posY+height
			TxtTxt = "DATE:"    
			TxtFont = "Arial"
			TxtSize = 2
			TxtBold = FALSE
			TxtRelPos = "rigth"
		Call GraphObjClose("TDate")	
    
		Call GraphObjOpen("Date")
			TxtPosX = posX+60
			TxtPosY= 1.7+posY+height
			TxtTxt = "VALUE"    
			TxtFont = "Arial"
			TxtSize = 2
			TxtBold = FALSE
			TxtRelPos = "rigth"
		Call GraphObjClose("Date")
		
		Call GraphObjOpen("Save")
			TxtPosX = 90
			TxtPosY= 90
			TxtTxt = "VALUE"    
			TxtFont = "Arial"
			TxtSize = 2
			TxtBold = FALSE
			TxtRelPos = "rigth"
		Call GraphObjClose("Save")
		
		Call PicUpdate()
	End Function
	
	Public Function addRaster(factor,amount)
		dim oRaster
		
		Set oRaster = new Raster	
		
		Call oRaster.setAmount(amount)
		Call oRaster.setFactor(factor)
		
		rasterAmount = rasterAmount + amount
		rasterSize = oGrid.getLength/rasterAmount
		plotDuration = plotDuration + oRaster.getDuration
		
		lRaster.add(oRaster)
	End Function
	
	Public Function changeRaster(factor,amount,id)
		dim oRaster

		Set oRaster = lRaster(id-1)
    
		rasterAmount = rasterAmount - oRaster.getAmount()
		rasterAmount = rasterAmount + amount
		
		Call oRaster.setAmount(amount)
		Call oRaster.setFactor(factor)
		
		rasterSize = oGrid.getLength/rasterAmount
		plotDuration = plotDuration + oRaster.getDuration
	End Function
	
	Public Function removeRaster(id)
		dim oRaster
		Set oRaster = lRaster(id-1)
		
		rasterAmount = rasterAmount - oRaster.getAmount()
		rasterSize = oGrid.getLength/rasterAmount
		plotDuration = plotDuration - oRaster.getDuration()

		lRaster.RemoveAt(id-1)
	End Function
	
	Public Function drawRaster()
		dim raster, startValue, number, duration, scale, i, startPoint, posY
		startPoint = oGrid.getPosX()
		startValue = 0
		number = 1
		posY = oGrid.getPosY()
		
		preparePlot()
		drawGrid()
		
		Call GraphObjOpen("Grid")
				Call GraphObjOpen(D2AxisXNam(1))
					D2AxisXEnd = rasterAmount
				Call GraphObjClose(D2AxisXNam(1))
		Call GraphObjClose("Grid") 
		
		Call GraphObjOpen("Save")
			TxtTxt = lRaster.Count
		Call GraphObjClose("Save")
		
		For Each raster in lRaster
			scale = rasterSize*raster.getAmount()
			duration = raster.getDuration()
			
			Call GraphObjNew("2D-Axis","Raster_"&number)   'Creates a new 2D axis system
			Call GraphObjOpen("Raster_"&number)
				'Position
				D2AxisTop        =100 - posY - oGrid.getHeight()
				D2AxisRight       =100 - startPoint - scale
				D2AxisBottom     =posY
				D2AxisLeft       =startPoint
				
				'Layout
				D2AxisSystem = "one system" 'WICHTIG
				'D2AxisDisp(1) = "Grid"
				'D2AxisDisp(2) = "Grid"
				D2AxisDispType = "Axis"
				'D2AxisHide(1) = TRUE
				D2AxisLineWidth = "min"
				
				'Scaling
				D2AxisScaledIn = "units per unit length"
				'D2MultipleScal = TRUE
				D2UseCommonXChn = TRUE
				
				'X-Axis
				Call GraphObjOpen(D2AxisXNam(1))
					D2AxisXBegin = startValue
					D2AxisXEnd = startValue + duration
					
					'Scaling
					D2AxisXScaleType = "manual"
					D2AxisXMiniTick = 0
					'D2AxisXOrigin = 0
					D2AxisXTick = duration+startValue
					D2AxisXTickAuto = FALSE
					D2AxisXUnitPreset = unit
					D2AxisYOffOrigin = "AxisBegin"
					
					
					'graphic
					D2AxisXSize = 0
					D2AxisXTickSize = 0
				Call GraphObjClose(D2AxisXNam(1))
			  
				For i = 1 to 9
					
					
					'Y-Axis
					Call GraphObjOpen(D2AxisYNam(i))
						D2AxisYBegin = -500
						D2AxisYEnd = 500
					
						'Scaling
						D2AxisYScaleType = "manual"
						D2AxisYMiniTick = 0
						'D2AxisXOrigin = 0
						D2AxisYTick = 200
						D2AxisYTickAuto = FALSE
						'D2AxisXUnitPreset = "ms"
						D2AxisXOffOrigin = "AxisBegin"
						
						'graphic
						D2AxisYSize = 0
						D2AxisYTickSize = 0
						D2AxisXOffset = 0
					Call GraphObjClose(D2AxisYNam(i))
			  
					If i < 9 Then
						GraphObjYAxisNew("left")
					End If
				Next
				
			Call GraphObjClose("Raster_"&number)
      
			'Unit
			Call GraphObjNew("FreeText","TRasterUnit_"&number)
			Call GraphObjOpen("TRasterUnit_"&number)
				TxtPosX = startPoint+0.5
				TxtPosY= posY-2
				TxtTxt = raster.getFactor()&" "&unit&" / []"    
				TxtFont = "Arial"
				TxtSize = 2
				TxtBold = FALSE
				TxtRelPos = "rigth"
			Call GraphObjClose("TRasterUnit_"&number)

			Call GraphObjNew("Arrow","UnitLine_"&number)
			Call GraphObjOpen("UnitLine_"&number)
				ArrowLineColor = "black"
				ArrowSymbolEnd = "NoArrow"
				ArrowSymbolBegin ="NoArrow"
				ArrowPTY(1)=posY
				ArrowPTY(2)=posY-4
				ArrowPTX(1)=startPoint
				ArrowPTX(2)=startPoint
				ArrowLineWidth = "min"
			Call GraphObjClose("UnitLine_"&number)
			
			startPoint = startPoint + scale
			startValue = startValue + duration
			number = number + 1
		Next
    
		Call PicUpdate() 
		
	End Function
	
End Class