Class LayoutPlotter
	
	private listSegments
	private oGrid
	private oLayout
	
	private unitSize
	private unitAmount
	private segmentAmount
	private plotDuration
	
	private MAX_CHANNELS
	
	private lSettings
	private lChannelGroups
	private lSelectedChannels '-->next plot will take the same
	
	private currentChannelGroup
	
	Public Sub Class_Initialize
		Set listSegments = CreateObject("System.Collections.ArrayList")
		MAX_CHANNELS = 18
		Set lChannelGroups = CreateObject("System.Collections.ArrayList")
		Set lSettings = CReateObject("System.Collections.ArrayList")
		Set lSettings = CReateObject("System.Collections.ArrayList")
		Call PicDelete()
	End Sub
	
	Public Function getTimeUnit()
		getTimeUnit = oLayout.getTimeUnit()
	End Function
	
	Public Function getSegmentAmount()
		getSegmentAmount = unitAmount
	End Function
	
	Public Function getPlotDuration()
		getPlotDuration = plotDuration
	End Function
	
	Public Function getSegment(id)
		Set getSegment = listSegments(id-1) 
	End Function
  
	Public Function getAllSegments()
		Set getAllSegments = listSegments
	End Function
	
	Public Function setLayout(timeUnit, pageFormat, pageLogo, pageTitle)
		Set oLayout = new Layout
		Call oLayout.setTimeUnit(timeUnit)
		Call oLayout.setPageFormat(pageFormat)
		Call oLayout.setPageLogo(pageLogo)
		Call oLayout.setPageTitle(pageTitle)
	End Function
	
	Public Function setGrid(marginLeft, marginBottom, gridWidth, gridHeight)
		Set oGrid = new Grid
		Call oGrid.setMarginLeft(marginLeft)
		Call oGrid.setMarginBottom(marginBottom)
		Call oGrid.setHeight(gridHeight)
		Call oGrid.setWidth(gridWidth)
		Call oGrid.calculateMargins()
		listSegments.Clear
	End Function

	Public Function drawLayout()
		dim height, posX, posY
		
		height = oGrid.getHeight()
		posX = oGrid.getMarginLeft()
		posY = oGrid.getMarginBottom()
		
		PicPageOrient = "landscape"
		
		Call Report.Sheets.Remove("Layout")
		Call Report.Sheets.Insert("Layout", 1)
		Call Report.Sheets.Remove(2)
		
		'Header Text
		Call GraphObjNew("FreeText","Place")
		Call GraphObjOpen("Place")
			TxtPosX = posX+1
			TxtPosY= 2+posY+height
			TxtTxt = oLayout.getPageLogo()  
			TxtFont = "Arial"
			TxtSize = 4
			TxtBold = FALSE
			TxtRelPos = "rigth"
		Call GraphObjClose("Place")
		
		Call GraphObjNew("FreeText","Laboratory")
		Call GraphObjOpen("Laboratory")
			TxtPosX = posX+9
			TxtPosY= 1.7+posY+height
			TxtTxt = oLayout.getPageTitle()   
			TxtFont = "Arial"
			TxtSize = 2.4
			TxtBold = TRUE
			TxtRelPos = "rigth"
		Call GraphObjClose("Laboratory")
		
		Call GraphObjNew("FreeText","TTest")
		Call GraphObjOpen("TTest")
			TxtPosX = posX+30
			TxtPosY= 1.7+posY+height
			TxtTxt = "TEST:"    
			TxtFont = "Arial"
			TxtSize = 2
			TxtBold = FALSE
			TxtRelPos = "rigth"
		Call GraphObjClose("TTest")
		
		Call GraphObjNew("FreeText","Test")
		Call GraphObjOpen("Test")
			TxtPosX = posX+40
			TxtPosY= 1.7+posY+height
			TxtTxt = "VALUE"    
			TxtFont = "Arial"
			TxtSize = 2
			TxtBold = FALSE
			TxtRelPos = "rigth"
		Call GraphObjClose("Test")		
		
		Call GraphObjNew("FreeText","TDate")
		Call GraphObjOpen("TDate")
			TxtPosX = posX+50
			TxtPosY= 1.7+posY+height
			TxtTxt = "DATE:"    
			TxtFont = "Arial"
			TxtSize = 2
			TxtBold = FALSE
			TxtRelPos = "rigth"
		Call GraphObjClose("TDate")	
		
		Call GraphObjNew("FreeText","Date")
		Call GraphObjOpen("Date")
			TxtPosX = posX+60
			TxtPosY= 1.7+posY+height
			TxtTxt = "VALUE"    
			TxtFont = "Arial"
			TxtSize = 2
			TxtBold = FALSE
			TxtRelPos = "rigth"
		Call GraphObjClose("Date")
	End Function
	
	Public Function drawGrid()
		dim height, posX, posY
		height = oGrid.getHeight()
		posX = oGrid.getMarginLeft()
		posY = oGrid.getMarginBottom()
		
		'Grid
		Call GraphObjNew("2D-Axis","Grid")   'Creates a new 2D axis system
		Call GraphObjOpen("Grid")
			'Position
			D2AxisTop        = oGrid.getMarginTop()      
			D2AxisRight       =oGrid.getMarginRight()
			D2AxisBottom     =posY
			D2AxisLeft       =posX
  
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
				D2AxisXEnd = 60
		
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
		
		Call PicUpdate()
	End Function
	
	Public Function addSegment(factor,amount)
		dim oSegment
		
		Set oSegment = new Segment	
		
		Call oSegment.setUnitAmount(amount)
		Call oSegment.setUnitFactor(factor)
		
		unitAmount = unitAmount + amount
		unitSize = oGrid.getWidth/unitAmount
		plotDuration = plotDuration + oSegment.getDuration
		
		listSegments.add(oSegment)
		segmentAmount = listSegments.Count
	End Function
	
	Public Function changeSegment(factor,amount,id)
		dim oSegment

		Set oSegment = listSegments(id-1)
    
		unitAmount = unitAmount - oSegment.getUnitAmount
		unitAmount = unitAmount + amount
		plotDuration = plotDuration - oSegment.getDuration
		
		Call oSegment.setUnitAmount(amount)
		Call oSegment.setUnitFactor(factor)
		
		unitSize = oGrid.getWidth/unitAmount
		plotDuration = plotDuration + oSegment.getDuration
	End Function
	
	Public Function removeSegment(id)
		dim oSegment
		Set oSegment = listSegments(id-1)
		
		unitAmount = unitAmount - oSegment.getUnitAmount()
		unitSize = oGrid.getWidth/unitAmount
		plotDuration = plotDuration - oSegment.getDuration()

		listSegments.RemoveAt(id-1)
		segmentAmount = listSegments.Count
	End Function
	
	Public Function drawSegments()
		dim i, segment, startTime, segmentNumber, segmentDuration, segmentSize,  segmentStartPositionX, segmentStartPositionY
		
		startTime = 0
		segmentNumber = 1
		segmentStartPositionY = oGrid.getMarginBottom()
		segmentStartPositionX = oGrid.getMarginLeft()
		
		Call GraphObjOpen("Grid")
				Call GraphObjOpen(D2AxisXNam(1))
					D2AxisXEnd = unitAmount
				Call GraphObjClose(D2AxisXNam(1))
		Call GraphObjClose("Grid") 
		
		'use segmentAmount
		'Call GraphObjOpen("Save")
		'	TxtTxt = listSegments.Count
		'Call GraphObjClose("Save")
		
		For Each segment in listSegments
			segmentSize = unitSize*segment.getUnitAmount()
			segmentDuration = segment.getDuration()
			
			Call GraphObjNew("2D-Axis","Segment_"&segmentNumber)   'Creates a new 2D axis system
			Call GraphObjOpen("Segment_"&segmentNumber)
				'Position
				D2AxisTop        =100 - segmentStartPositionY - oGrid.getHeight()
				D2AxisRight       =100 - segmentStartPositionX - segmentSize
				D2AxisBottom     =segmentStartPositionY
				D2AxisLeft       =segmentStartPositionX
				
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
					D2AxisXBegin = startTime
					D2AxisXEnd = startTime + segmentDuration
					
					'Scaling
					D2AxisXScaleType = "manual"
					D2AxisXMiniTick = 0
					'D2AxisXOrigin = 0
					D2AxisXTick = segmentDuration+startTime
					D2AxisXTickAuto = FALSE
					D2AxisXUnitPreset = oLayout.getTimeUnit
					D2AxisYOffOrigin = "AxisBegin"
					
					'graphic
					D2AxisXSize = 0
					D2AxisXTickSize = 0
				Call GraphObjClose(D2AxisXNam(1))
			  
				For i = 1 to MAX_CHANNELS
					
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
			  
					If i < MAX_CHANNELS Then
						GraphObjYAxisNew("left")
					End If
				Next
				
			Call GraphObjClose("Segment_"&segmentNumber)
      
			'Unit
			Call GraphObjNew("FreeText","TSegmentUnit_"&segmentNumber)
			Call GraphObjOpen("TSegmentUnit_"&segmentNumber)
				TxtPosX = segmentStartPositionX+0.5
				TxtPosY= segmentStartPositionY-2
				TxtTxt = segment.getUnitFactor()&" "&oLayout.getTimeUnit&" / []"    
				TxtFont = "Arial"
				TxtSize = 2
				TxtBold = FALSE
				TxtRelPos = "rigth"
			Call GraphObjClose("TSegmentUnit_"&segmentNumber)
			
			Call GraphObjNew("Arrow","UnitLine_"&segmentNumber)
			Call GraphObjOpen("UnitLine_"&segmentNumber)
				ArrowLineColor = "black"
				ArrowSymbolEnd = "NoArrow"
				ArrowSymbolBegin ="NoArrow"
				ArrowPTY(1)=segmentStartPositionY
				ArrowPTY(2)=segmentStartPositionY-4
				ArrowPTX(1)=segmentStartPositionX
				ArrowPTX(2)=segmentStartPositionX
				ArrowLineWidth = "min"
			Call GraphObjClose("UnitLine_"&segmentNumber)
			
			segmentStartPositionX = segmentStartPositionX + segmentSize
			startTime = startTime + segmentDuration
			segmentNumber = segmentNumber + 1
		Next
    
		Call PicUpdate() 
		
	End Function
	
	Public Function changeChannelGroup(groupName)
			
		Set currentChannelGroup = new ChannelGroup
		currentChannelGroup.setGroupName(groupName)
		
		If Report.Sheets.Exists(groupName) Then
		  Report.Sheets(groupName).Activate()
		  Call GraphObjOpen("Test")
			currentChannelGroup.setGroupTitle(TxtTxt)
		  Call GraphObjClose("Test")	
		  
		  Call GraphObjOpen("Date")
			currentChannelGroup.setGroupDate(TxtTxt)
		  Call GraphObjClose("Date")
		Else
		Call  Report.Sheets.Copy("Layout",groupName,2)
				
				If NOT Data.Root.ChannelGroups(groupName).Properties.Exists("TestDate")Then 
					currentChannelGroup.setGroupDate(InputBox("No Value for TestDate found!"&Chr(13)&"Please enter Value","No Properties","NOVALUE"))
				Else
					currentChannelGroup.setGroupDate(Data.Root.ChannelGroups(groupName).Properties("TestDate").Value)
				End If
				
				If InStr(groupName,".") = 0 Then
					currentChannelGroup.setGroupTitle(groupName)
				Else
					currentChannelGroup.setGroupTitle(Left(groupName,InStr(groupName,".")-1))
				End If
		  
		  Call GraphObjOpen("Test")
			TxtTxt =  currentChannelGroup.getGroupTitle()
		  Call GraphObjClose("Test")	
		  
		  Call GraphObjOpen("Date")
			TxtTxt =  currentChannelGroup.getGroupDate()
		  Call GraphObjClose("Date")
		  
			End If
		
		If NOT Data.Root.ChannelGroups.Exists(currentChannelGroup.getGroupTitle()&".plot") Then
		  Data.Root.ChannelGroups.Add(currentChannelGroup.getGroupTitle()&".plot")
		End If
		
		Data.Root.ChannelGroups(currentChannelGroup.getGroupTitle()&".plot").Activate
		
		'Call draw Function
		Call PicUpdate()
	End Function
	
End Class