'LayoutPlotter.VBS
'_____________________________
'Date: 2015-08-??
'Author: Adrian Kress
'Version old: N/A
'Version new: New 1.0
'Comment: Initial script
'
'
'
'
'Date: 2017-04-21
'Author: Adrian Kress
'Version old: 1.0
'Version new: 1.1
'Comment: L89;   checks first if a sheet "Layout" exists
'        L606;  checks if a sheet "currentChannelGroup.getGroupName()" exists
'_____________________________
'Date:  2017-07-12
'Author:  Adrian Kress
'Version old: 1.1
'Version new: 1.2
'Comment: L781;  updated name and unit placement in plot
'_____________________________
'Date:  2017-07-25
'Author:  Adrian Kress 
'Version old: 1.2
'Version new: 1.3
'Comment: L123;  update ABB Logo
'_____________________________
'Date:  2017-08-08
'Author:  Adrian Kress 
'Version old: 1.3
'Version new: 1.4
'Comment: L783;   updated name and unit positioning
'_____________________________
'Date:  2017-11-27
'Author:  Adrian Kress 
'Version old: 1.4
'Version new: 1.5
'Comment: Lxxx;   updated the postion of the text
'_____________________________
'Date:  2018-07-05
'Author:  Philip Streit
'Version old: 1.5
'Version new: 1.6
'Comment: L541:   added If statement to check the shot and set the title accordingly - still in progress
'         L001;   updated Header
'_____________________________
'Date:  2018-07-06
'Author:  Philip Streit
'Version old: 1.6
'Version new: 1.7
'Comment: L549:   new method for setting Name and shot on the Plot
'_____________________________
'Date:  2019-09-19
'Author:  Adrian Kress 
'Version old: 1.7
'Version new: 1.8
'Comment: L824;   added get_point(channel_name): calculates the last y-point of the channel 'channel_name' to fix S:2018-M-079/1328 C: U_Aux_BG1_101_10
'_____________________________
'
'Descpription
'creates a Plot
'__________________________


Class LayoutPlotter
	
	private listSegments
	private oGrid
	private oLayout
	
	private unitSize  'in % of the paper length (100%)
	private unitAmount
	private segmentAmount
	private plotDuration 'in time
	private gridStart
	
	private MAX_CHANNELS  'max channels per plot
	
	private lSettings
	private lChannelGroups
	private lSelectedChannels '-->next plot will take the same
	
	private currentChannelGroup
	
	private oDetailPlotter 
	
	Public Sub Class_Initialize
		MAX_CHANNELS = 18
		Set listSegments = CreateObject("System.Collections.ArrayList")
		Set lChannelGroups = CreateObject("System.Collections.ArrayList")
		Set lSettings = CReateObject("System.Collections.ArrayList")
		Set lSelectedChannels = CReateObject("System.Collections.ArrayList")
		Set oDetailPlotter = new DetailPlotter
		Set lDetailSetting = CreateObject("System.Collections.ArrayList")
		Call PicDelete()
	End Sub
	
	Public Function getTimeUnit()
		getTimeUnit = oLayout.getTimeUnit()
	End Function
	
  
	Public Function getMaxChannels()
		getMaxChannels = MAX_CHANNELS
	End Function
	
	Public Function getUnitAmount()
		getUnitAmount = unitAmount
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
	
	Public Function setGrid(marginLeft, marginBottom, gridWidth, gridHeight, VgridStart)
		Set oGrid = new Grid
		Call oGrid.setMarginLeft(marginLeft)
		Call oGrid.setMarginBottom(marginBottom)
		Call oGrid.setHeight(gridHeight)
		Call oGrid.setWidth(gridWidth)
		gridStart = VgridStart
		listSegments.Clear
	End Function

	
  'draw Title, Description, Date and ShotNumber 
	Public Function drawLayout()
		dim height, posX, posY
		
		height = oGrid.getHeight()
    'position on 2D axis
		posX = oGrid.getMarginLeft()
		posY = oGrid.getMarginBottom()
		
		PicPageOrient = "landscape"
   
    If Report.Sheets.Exists("Layout") Then
      Call Report.Sheets.Remove("Layout")
    End If   
		Call Report.Sheets.Insert("Layout", 1)
		PicDefByIdent = 1
		
		'Header Text
		Select Case oLayout.getPageLogo
			Case "ABB"
				Call GraphObjNew("FreeGraph","Place")
				Call GraphObjOpen("Place")
					MtaPosX = posX+1
					MtaPosY= 1+posY+height
					MtaRelPos = "r-top"
					MtaFileName = CurrentScriptPath&"ABB.png" 
					MtaHeight = 2
					MtaWidth = 5.4
				Call GraphObjClose("Place")
			Case "PEHLA"
				Call GraphObjNew("FreeGraph","Place")
				Call GraphObjOpen("Place")
					MtaPosX = posX+1
					MtaPosY= 1+posY+height
					MtaRelPos = "r-top"
					MtaFileName = CurrentScriptPath&"PEHLA.png" 
					MtaHeight = 3
					MtaWidth = 9
				Call GraphObjClose("Place")
		End Select
		
		Call GraphObjNew("FreeText","Laboratory")
		Call GraphObjOpen("Laboratory")
			TxtPosX = posX+12
			TxtPosY= 1+posY+height
			TxtTxt = oLayout.getPageTitle()   
			TxtFont = "Arial"
			TxtSize = 2.4
			TxtBold = TRUE
			TxtRelPos = "r-top"
		Call GraphObjClose("Laboratory")
		
		Call GraphObjNew("FreeText","TTest")
		Call GraphObjOpen("TTest")
			TxtPosX = posX+35
			TxtPosY= 1+posY+height
			TxtTxt = "TEST:"    
			TxtFont = "Arial"
			TxtSize = 2
			TxtBold = FALSE
			TxtRelPos = "r-top"
		Call GraphObjClose("TTest")
		
		Call GraphObjNew("FreeText","Test")
		Call GraphObjOpen("Test")
		'"Value" is a placeholder and is being changed as the Next Shot is loaded.
			TxtPosX = posX+40
		' ^moved a bit to the left to guarantee more space
			TxtPosY= 1+posY+height
			TxtTxt = "VALUE"    
			TxtFont = "Arial"
			TxtSize = 2
			TxtBold = FALSE
			TxtRelPos ="r-top"
		Call GraphObjClose("Test")		
		
		Call GraphObjNew("FreeText","TDate")
		Call GraphObjOpen("TDate")
			TxtPosX = posX+60
			TxtPosY= 1+posY+height
			TxtTxt = "DATE:"    
			TxtFont = "Arial"
			TxtSize = 2
			TxtBold = FALSE
			TxtRelPos = "r-top"
		Call GraphObjClose("TDate")	
		
		Call GraphObjNew("FreeText","Date")
		Call GraphObjOpen("Date")
			TxtPosX = posX+65
			TxtPosY= 1+posY+height
			TxtTxt = "VALUE"    
			TxtFont = "Arial"
			TxtSize = 2
			TxtBold = FALSE
			TxtRelPos = "r-top"
		Call GraphObjClose("Date")
	End Function
	
	Public Function drawGrid()
		dim height, posX, posY
		height = oGrid.getHeight()
		posX = oGrid.getMarginLeft()
		posY = oGrid.getMarginBottom()
		
		'Grid
		Call GraphObjNew("2D-Axis","Grid")   'Creates a new 2D axis system
		'objects to be drawn must be opened and closed after definition
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
			
		'objects must be closed after definition	
		Call GraphObjClose("Grid") 
		
		Call PicUpdate()
	End Function
	
	
	Public Function addSegment(factor,amount)
		dim oSegment
		
		Set oSegment = new Segment	
		
		Call oSegment.setUnitAmount(amount)
		Call oSegment.setUnitFactor(factor)
		
    'update the total amount, size and duration
		unitAmount = unitAmount + amount
		unitSize = oGrid.getWidth/unitAmount
		plotDuration = plotDuration + oSegment.getDuration
		
		listSegments.add(oSegment)
		segmentAmount = listSegments.Count
	End Function
	
	Public Function changeSegment(factor,amount,id)
		dim oSegment

		Set oSegment = listSegments(id-1)
    
    'remove the current segment from total amount
		unitAmount = unitAmount - oSegment.getUnitAmount
		plotDuration = plotDuration - oSegment.getDuration
		
		Call oSegment.setUnitAmount(amount)
		Call oSegment.setUnitFactor(factor)
		
    'add the current segment to total amount with new values
		unitAmount = unitAmount + amount
		unitSize = oGrid.getWidth/unitAmount
		plotDuration = plotDuration + oSegment.getDuration
	End Function
	
	Public Function removeSegment(id)
		dim oSegment
		Set oSegment = listSegments(id-1)
		
    'update the total amount, size and duration
		unitAmount = unitAmount - oSegment.getUnitAmount()
		unitSize = oGrid.getWidth/unitAmount
		plotDuration = plotDuration - oSegment.getDuration()

		listSegments.RemoveAt(id-1)
		segmentAmount = listSegments.Count
	End Function
  	
	Public Function drawSegments()
		dim i, segment, startTime, segmentNumber, segmentDuration, segmentSize,  segmentStartPositionX, segmentStartPositionY
		
		startTime = gridStart
		segmentNumber = 1
    
    'start position from the grid
		segmentStartPositionY = oGrid.getMarginBottom()
		segmentStartPositionX = oGrid.getMarginLeft()
		
    'set raster amount on the grid : 1 unit = 1 raster
		Call GraphObjOpen("Grid")
				Call GraphObjOpen(D2AxisXNam(1))
					D2AxisXEnd = unitAmount
				Call GraphObjClose(D2AxisXNam(1))
		Call GraphObjClose("Grid") 
		
		For Each segment in listSegments
			segmentSize = unitSize*segment.getUnitAmount()
			segmentDuration = segment.getDuration() ' in time unit (s/ms/..)
			
			Call GraphObjNew("2D-Axis","Segment_"&segmentNumber)   'Creates a new 2D axis system
			Call GraphObjOpen("Segment_"&segmentNumber)
				'Position
				D2AxisTop        =100 - segmentStartPositionY - oGrid.getHeight()
				D2AxisRight       =100 - segmentStartPositionX - segmentSize
				D2AxisBottom     =segmentStartPositionY
				D2AxisLeft       =segmentStartPositionX
				
				'Layout
				D2AxisSystem = "one system" 'IMPORTANT
				D2AxisDispType = "Axis"
				D2AxisLineWidth = "min"
				
				'Scaling
				D2AxisScaledIn = "units per unit length"
				D2UseCommonXChn = False
				
				'X-Axis
				Call GraphObjOpen(D2AxisXNam(1))
					D2AxisXBegin = startTime
					D2AxisXEnd = startTime + segmentDuration
					
					'Scaling
					D2AxisXScaleType = "manual"
					D2AxisXMiniTick = 0
					D2AxisXTick = startTime + segmentDuration ' -> no tick
					D2AxisXTickAuto = FALSE
					D2AxisXUnitPreset = oLayout.getTimeUnit
					D2AxisYOffOrigin = "AxisBegin"
					
					'graphic
					D2AxisXSize = 0
					D2AxisXTickSize = 0
				Call GraphObjClose(D2AxisXNam(1))
        
        'every channel needs an y-axis
				For i = 1 to MAX_CHANNELS
					
					'Y-Axis
					Call GraphObjOpen(D2AxisYNam(i))
            'height of 1000 points
						D2AxisYBegin = -500
						D2AxisYEnd = 500
					
						'Scaling
						D2AxisYScaleType = "manual"
						D2AxisYMiniTick = 0
						D2AxisYTick = 200
						D2AxisYTickAuto = FALSE
						D2AxisXOffOrigin = "AxisBegin"
						
						'graphic
						D2AxisYSize = 0
						D2AxisYTickSize = 0
						D2AxisXOffset = 0
					Call GraphObjClose(D2AxisYNam(i))
			  
					If i < MAX_CHANNELS Then
            'creates new y-axis
						GraphObjYAxisNew("left")
					End If
				Next
				
			Call GraphObjClose("Segment_"&segmentNumber)
      
      
      'Segment description
			'Unit
			Call GraphObjNew("FreeText","TSegmentUnit_"&segmentNumber)
			Call GraphObjOpen("TSegmentUnit_"&segmentNumber)
				TxtPosX = segmentStartPositionX+0.5
				TxtPosY= segmentStartPositionY-2
				TxtTxt = segment.getUnitFactor()&" "&oLayout.getTimeUnit&" / DIV"    
				TxtFont = "Arial"
				TxtSize = 1.5
				TxtBold = FALSE
				TxtRelPos = "rigth"
			Call GraphObjClose("TSegmentUnit_"&segmentNumber)
			'Line
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
			
      'starting point and tim for next segment
			segmentStartPositionX = segmentStartPositionX + segmentSize 
			startTime = startTime + segmentDuration 
      
			segmentNumber = segmentNumber + 1
		Next
    
		Call PicUpdate() 
		
	End Function
	
	
	Public Function changeChannelGroup(groupName)
		dim channel, i, n
			
		Set currentChannelGroup = new ChannelGroup
		currentChannelGroup.setGroupName(groupName)
		
		LBChannels.Items.RemoveAll
		For each channel in Data.Root.ChannelGroups(groupName).Channels
			Call LBChannels.Items.Add(channel.Name,1)
		Next    
		
		If Report.Sheets.Exists(groupName) Then 'already exist
			Report.Sheets(groupName).Activate()
			Call GraphObjOpen("Test")
				currentChannelGroup.setGroupTitle(TxtTxt)
			Call GraphObjClose("Test")	
			
			Call GraphObjOpen("Date")
				currentChannelGroup.setGroupDate(TxtTxt)
			Call GraphObjClose("Date")

			segmentAmount = Data.Root.ChannelGroups(currentChannelGroup.getGroupName()&".plot").Properties.Item("segmentAmount").Value
			
			'read the segment Parameters from the Plot, show them in the GUI
			loadSegements()
			
			drawLayout()
			drawGrid()
			drawSegments()
			
			
			For each channel in Data.Root.ChannelGroups(currentChannelGroup.getGroupName()&".plot").Channels
				Call LBChannelsSelected.Items.Add(channel.Name,1)
				oLayoutPlotter.setSettings(channel)
			Next
  
			drawGroupChannels()
  
		Else
			
			If NOT Data.Root.ChannelGroups(groupName).Properties.Exists("TestDate")Then 
				currentChannelGroup.setGroupDate(InputBox("No Value for TestDate found!"&Chr(13)&"Please enter Value","No Properties","NOVALUE"))
			Else
				currentChannelGroup.setGroupDate(Data.Root.ChannelGroups(groupName).Properties("TestDate").Value)
			End If
			 
			'remove any .rbe additions from the name of the channel group
			If InStr(groupName,".") = 0 Then
				currentChannelGroup.setGroupTitle(Right(groupName,Len(groupName)-1))
			Else
          		If (IsNumeric(Left(groupName,1))) Then
					currentChannelGroup.setGroupTitle(Left(Right(groupName,Len(groupName)),InStr(Right(groupName,Len(groupName)),".")-1))
                	' ^ rewrite this line of code because it is very hard to read and some argument are unnecessary.
                	'currentChannelGroup.setGroupTitle(groupName), 
				Else
					currentChannelGroup.setGroupTitle(Left(Right(groupName,Len(groupName)-1),InStr(Right(groupName,Len(groupName)-1),".")-1))
				End If
			End If
       
			'modify Title if Pehla shot
			If oLayout.getPageLogo() = "PEHLA" Then
				dim pehla, pTitle, pDate, pShot, pLen, a, b, c
				
				pDate = currentChannelGroup.getGroupDate()  
				pehla = Right(Left(pDate,4),2)
				pTitle = currentChannelGroup.getGroupTitle()			
				
				pShot = Right(pTitle,Len(pTitle)-InStr(pTitle,"-"))
				pehla = pehla & Right(Left(pTitle,InStr(pTitle,"-")-1),3) & "Ba / "
				
				'if campaign-field exists then set pTitle to campaign value
            	'to check if there are any attributes availabele to set as pTitle
				If Data.Root.ChannelGroups(groupName).Properties.Exists("PEHLA") Then
					'PEHLA
					pehla = Data.Root.ChannelGroups(groupName).Properties("PEHLA").Value & "/"
					a = Data.Root.ChannelGroups(groupName).Properties("sourceoriginalname").Value
					b = Right(a, 8)
					c = Left(b, 4)
					pShot = c
				ElseIf Data.Root.ChannelGroups(groupName).Properties.Exists("TNS_CAMPAIGN") Then
					' ^ use ElseIf written together otherwise you have to add a closing End If down below
					'TNS_CAMPAIGN
					pehla = Data.Root.ChannelGroups(groupName).Properties("TNS_CAMPAIGN").Value & " / "
					a = Data.Root.ChannelGroups(groupName).Properties("sourceoriginalname").Value
					b = Right(a, 8)
					c = Left(b, 4)
					pShot = c
				Else
					'If no additional attributes are given
					pTitle = currentChannelGroup.getGroupTitle()
      			End if

				pLen = pehla
          
				'test number, remove leading 0 --> 00nn0 = nn0
				For i = 1 to Len(pLen) 
					If InStr(pehla,"0") = i Then
						pehla = Right(pehla,Len(pehla)-1) 
						i = i -1
					Else
						Exit For
					End If

				Next
          
				pLen = pShot
				
				'shot number remove leading 0
				For i = 1 to Len(pLen) 
					If InStr(pShot,"0") = i Then
						pShot = Right(pShot,Len(pShot)-1)
						i = i -1
					Else
						Exit For
					End If
				Next
          
				currentChannelGroup.setGroupTitle(pehla&pShot)
          
			'create ABB Title
			Elseif oLayout.getPageLogo() = "ABB" Then
				dim aTitle, aShot,  aLen
				
				aTitle = currentChannelGroup.getGroupTitle()
				
				aTitle = Replace(aTitle,"-","/")
				
				aShot = Right(aTitle,Len(aTitle)-InStr(aTitle,"/"))
				
				aTitle = Left(aTitle,Instr(aTitle,"/"))
				
				aTitle = Replace(aTitle,"/"," / ")
				
				aLen = aTitle
				
				'test number, remove leading 0 --> 00nn0 = nn0
				For i = 1 to Len(aLen) 
					If InStr(aTitle,"0") = i Then
						aTitle = Right(aTitle,Len(aTitle)-1) 
						i = i -1
					Else
						Exit For
					End If
				
				Next
				
				aLen = aShot
				
				'shot number remove leading 0
				For i = 1 to Len(aLen) 
					If InStr(aShot,"0") = i Then
						aShot = Right(aShot,Len(aShot)-1)
						i = i -1
					Else
						Exit For
					End If
				Next
          
				currentChannelGroup.setGroupTitle(aTitle&aShot)
          
			End If
			
			
			
			
			If Data.Root.ChannelGroups.Exists(currentChannelGroup.getGroupName()&".plot") Then
				Call Data.Root.ChannelGroups.Remove(currentChannelGroup.getGroupName()&".plot")
			End If
			
			Call Data.Root.ChannelGroups.Add(currentChannelGroup.getGroupName()&".plot")
			Call Data.Root.ChannelGroups(currentChannelGroup.getGroupName()&".plot").Properties.Add("segmentAmount",segmentAmount)
		  
			drawGroupChannels()
			'add previous selection to selection list in GUI
			If lSelectedChannels.Count > 0 AND LBChannels.Items.Count > 0 Then
				For i = 0 to lSelectedChannels.Count-1
					For  n = 1 to LBChannels.Items.Count
						if lSelectedChannels(i) = LBChannels.Items(n).Text Then
							Call LBChannelsSelected.Items.Add(LBChannels.Items(n).Text,1)
						End If
					Next
				Next
			End If
      
		End If
		 
		Data.Root.ChannelGroups(currentChannelGroup.getGroupName()&".plot").Activate
		addSelectionToPlotgroup()
		convertPlotgroupChannels() 
		plotShotChannels()
		
		Call PicUpdate()
		'Detail Plot
		Call oDetailPlotter.initDetailPlot(currentChannelGroup.getGroupName(),currentChannelGroup.getGroupTitle(),currentChannelGroup.getGroupDate(),oLayout.getPageLogo(),oLayout.getPageTitle())

	End Function
	
	
	'draw new plot for the current channel group
	Public Function drawGroupChannels()
		If Report.Sheets.Exists(currentChannelGroup.getGroupName()) Then
		  Call Report.Sheets.Remove(currentChannelGroup.getGroupName())
    end if
 
		Call GraphSheetCopy("Layout",currentChannelGroup.getGroupName(),2)
		PicDefByIdent = 1
    
		Call GraphObjOpen("Test")
			TxtTxt =  currentChannelGroup.getGroupTitle()
		Call GraphObjClose("Test")	
		  
		Call GraphObjOpen("Date")
			TxtTxt =  currentChannelGroup.getGroupDate()
		Call GraphObjClose("Date")
    
		Data.Root.ChannelGroups(currentChannelGroup.getGroupName()&".plot").Properties.Item("segmentAmount").Value= segmentAmount
		Call PicUpdate()
	End Function
	
	Public Function loadSegements()
		dim segmentNumber, segmentFactor, gridSegementAmount, gridLength, segmentDuration, unitSize, segmentUnitAmount
		
		listSegments.Clear
		unitAmount = 0
		unitSize = 0
		plotDuration = 0
		
		Call GraphObjOpen("Grid")
			gridLength = 100 - D2AxisLeft - D2AxisRight
			Call GraphObjOpen(D2AxisXNam(1))
				gridSegementAmount = D2AxisXEnd
			Call GraphObjClose(D2AxisXNam(1))
		Call GraphObjClose("Grid") 
      
		unitSize = gridLength/gridSegementAmount
    
		For segmentNumber = 1 to segmentAmount
      
			Call GraphObjOpen("TSegmentUnit_"&segmentNumber)
				segmentFactor = TxtTxt  
			Call GraphObjClose("TSegmentUnit_"&segmentNumber)
			
			segmentFactor = (Left(segmentFactor,InStr(segmentFactor," ms / DIV")-1))
			
			Call GraphObjOpen("Segment_"&segmentNumber)
				segmentDuration = 100 - D2AxisRight - D2AxisLeft 
			Call GraphObjClose("Segment_"&segmentNumber)
			
			segmentUnitAmount = segmentDuration/unitSize
			
			Call addSegment(segmentFactor,segmentUnitAmount)
			
		Next
		
		refreshLBSegment()
		LBSegment.Selection = LBSegment.Items.Count
	End Function
	
	Public Function addSelectionToPlotgroup()
		dim i, source
		If LBChannelsSelected.Items.Count > 0 Then
			
			Data.Root.ChannelGroups(currentChannelGroup.getGroupName()&".plot").Channels.RemoveAll
			lSelectedChannels.Clear
			For i = 1 to LBChannelsSelected.Items.Count
					Call Data.Root.ChannelGroups(currentChannelGroup.getGroupName()&".plot").Activate
					Set source = Data.Root.ChannelGroups(currentChannelGroup.getGroupName()).Channels.Item(LBChannelsSelected.Items(i).Text) 
					Call Data.Root.ChannelGroups(currentChannelGroup.getGroupName()&".plot").Channels.AddChannel(source)
					lSelectedChannels.Add(LBChannelsSelected.Items(i).Text)
			Next
		End If
	End Function
  
	Public Function convertPlotgroupChannels()
		dim channel, channelHigh, channelLow, channelAmplitutde
		dim factor, scal, off, offstep, plotScal
		
		For each channel in Data.Root.ChannelGroups(currentChannelGroup.getGroupName()&".plot").Channels
			channelHigh = channel.Properties("maximum").Value
			channelLow = channel.Properties("minimum").Value
		
			If channelHigh < 0 Then
				channelHigh = channelHigh * -1
			End If
      
			If channelLow < 0 Then
				channelLow = channelLow * -1
			End If
      
			If channelHigh > channelLow Then
				channelAmplitutde = channelHigh
			ElseIf channelHigh < channelLow Then
				channelAmplitutde = channelLow
			End If
      
			channelAmplitutde = Round(channelAmplitutde)
		
			factor = 1
			scal = 1.0
			off = 0.0
			offstep = 50
		
				'convert all channels to similar size, between 50 and 150
				If channelAmplitutde > 150 OR channelAmplitutde  < 50 Then
					factor =  (Left(channelAmplitutde,1)+1)*10^(Len(channelAmplitutde)-1)/100
					Call ChnLinScale(channel.GetReference(eRefTypeIndexIndex),channel.GetReference(eRefTypeIndexIndex),1/factor,0)
				End If
			
			Call channel.Properties.Add("PlotFactor",factor)
			Call channel.Properties.Add("PlotScal",scal)
			Call channel.Properties.Add("PlotOff",off)
			Call channel.Properties.Add("PlotOffstep",offstep)  
			
			
			loadChannelSettings(channel)
			
		Next
		
		For each channel in Data.Root.ChannelGroups(currentChannelGroup.getGroupName()&".plot").Channels
		  plotScal = channel.Properties.Item("PlotScal").Value
		  Call ChnLinScale(channel.GetReference(eRefTypeIndexIndex),channel.GetReference(eRefTypeIndexIndex),plotScal,0)
		Next
		
	End Function
	
	Public Function loadChannelSettings(channel)
	  dim setting
	  
		If NOT lSettings.Count = 0 Then
		  For each setting in lSettings
			If setting.getName() = channel.Name Then
			
				channel.Properties.Item("PlotScal").Value = CDbl(setting.getScal())
				channel.Properties.Item("PlotOff").Value = CDbl(setting.getOff())
				channel.Properties.Item("PlotOffstep").Value = CInt(setting.getOffstep())
			End If
		  Next
		End If
	End Function

' >>> calculates the last y-point of the channel 'channel_name'
Function get_point(channel_name)
  get_point = (oGrid.getMarginBottom() + oGrid.getHeight()/2 + getLastPoint(channel_name)* offsetSize) + 4.5 * plotOff
End Function
  
Public Function plotShotChannels()
  dim channel, factor, plotOffStep,plotOff, plotScal, segment,i, offsetSize, last_point
  
  i = 1
  
  'size of one data point
  offsetSize = oGrid.getHeight()/1000
  
  For each channel in Data.Root.ChannelGroups(currentChannelGroup.getGroupName()&".plot").Channels
    factor = channel.Properties.Item("PlotFactor").Value
    plotOffStep = channel.Properties.Item("PlotOffStep").Value
    plotOff = channel.Properties.Item("PlotOff").Value
    plotScal = channel.Properties.Item("PlotScal").Value
    
    ' set last point outside of this function
    last_point = get_point(channel.Name)
    'segmentAmount is the number of time axes 
    For segment = 1 To segmentAmount
        Call GraphObjOpen("Segment_"&segment)  
          Call  GraphObjNew("2D-Curve","Curve_"&segment&i) 'Creates a new curve
          Call GraphObjOpen("Curve_"&segment&i)
            D2CurveColor = "black"
            D2CAxisPairNo = i
            D2CChnYName = channel.GetReference(eRefTypeNameIndex)
            
          Call GraphObjClose("Curve_"&segment&i)
          Call GraphObjOpen(D2AxisYNam(i))
            D2AxisYBegin = -500 - plotOffStep*plotOff
            D2AxisYEnd = 500 - plotOffStep*plotOff
          Call GraphObjClose(D2AxisYNam(i))
        Call GraphObjClose("Segment_"&segment)
      Next
    
    
      Call GraphObjNew("FreeText","TName_"&i)
      Call GraphObjOpen("TName_"&i)
        TxtPosX = 79
        TxtPosY= last_point '(oGrid.getMarginBottom() + oGrid.getHeight()/2 + getLastPoint(channel.Name)* offsetSize) + 4.5 * plotOff 'plotOff*plotOffStep*offsetSize+getLastPoint(channel.Name)*offsetSize
        TxtTxt = channel.Name   
        TxtFont = "Arial"
        TxtSize = 2
        TxtBold = FALSE
        TxtRelPos = "right"
      Call GraphObjClose("TName_"&i)
      
      Call GraphObjNew("FreeText","TUnit_"&i)
      Call GraphObjOpen("TUnit_"&i)
        TxtPosX = 98
        TxtPosY= last_point '(oGrid.getMarginBottom() + oGrid.getHeight()/2 + getLastPoint(channel.Name)* offsetSize) + 4.5 * plotOff 'plotOff*plotOffStep*offsetSize+getLastPoint(channel.Name)*offsetSize
        TxtTxt = factor*50/plotScal&" "&channel.Properties.Item("unit_string").Value&" / DIV"     
        TxtFont = "Arial"
        TxtSize = 2
        TxtBold = FALSE
        TxtRelPos = "left"
      Call GraphObjClose("TUnit_"&i)
      
      i = i+ 1
		Next
		'DIAdem function to refresh the content of the report
		Call PicUpdate()
    
	End Function
	
	'get the value of the last point of the current channel
	Public Function getLastPoint(channelName)
		dim channelUnit, waveStep, pointsUnit, endPoint
		
		endPoint = Data.Root.ChannelGroups(currentChannelGroup.getGroupName()&".plot").Channels(channelName).Properties("length").Value
		
		channelUnit = Data.Root.ChannelGroups(currentChannelGroup.getGroupName()&".plot").Channels(channelName).Properties("wf_xunit_string").Value
		waveStep = Data.Root.ChannelGroups(currentChannelGroup.getGroupName()&".plot").Channels(channelName).Properties("wf_increment").Value
		
		
		pointsUnit = 1/waveStep
		
		'convert to match the unit s
		Select Case channelUnit
		  Case "s"
			'no conversion needed
		  Case "ms"
			pointsUnit = 1000*pointsUnit
		  Case "µs"
			pointsUnit = 1000000*pointsUnit
		  Case Else
			MsgBox("Time Unit not supported!")
			getLastPoint = 0
			Exit Function
		End Select
		
		Select Case oLayout.getTimeUnit
		  Case "s"
			'no conversion needed
		  Case "ms"
			pointsUnit = pointsUnit/1000
		  Case "µs"
			pointsUnit = pointsUnit/1000000
	   End Select
		
		If pointsUnit*plotDuration > endPoint Then
		  getLastPoint = Data.Root.ChannelGroups(currentChannelGroup.getGroupName()&".plot").Channels(channelName).Values(endPoint)
		Else
		  getLastPoint = Data.Root.ChannelGroups(currentChannelGroup.getGroupName()&".plot").Channels(channelName).Values(pointsUnit*plotDuration)
		End If
	
	End Function
	
	Public Function setSize(channelName, factor)
		dim origscal, setting, scal, channel, i, cr, plotOff, plotOffStep, offsetSize, newYPos

		Set channel = Data.Root.ChannelGroups(currentChannelGroup.getGroupName()&".plot").Channels(channelName)
    
		offsetSize = oGrid.getHeight()/1000
	
		origscal = channel.Properties.Item("PlotScal").Value
		plotOffStep = channel.Properties.Item("PlotOffStep").Value
		plotOff = channel.Properties.Item("PlotOff").Value
    
		'get the channel index in the group
		i =InStr(channel.GetReference(eRefTypeIndexIndex),"/[") 
		cr = Right(channel.GetReference(eRefTypeIndexIndex),Len(channel.GetReference(eRefTypeIndexIndex))-i-1)
		i = CInt(Left(cr,Len(cr)-1))
  
		Call ChnLinScale(channel.GetReference(eRefTypeIndexIndex),channel.GetReference(eRefTypeIndexIndex),1/origscal,0)

		scal = (channel.Properties.Item("PlotFactor").Value*50)/factor

		Call ChnLinScale(channel.GetReference(eRefTypeIndexIndex),channel.GetReference(eRefTypeIndexIndex),scal,0)
  
		channel.Properties.Item("PlotScal").Value = scal

		newYPos = oGrid.getMarginBottom() + oGrid.getHeight()/2+plotOff*plotOffStep*offsetSize+getLastPoint(channel.Name)*offsetSize
		
		'check if text goes out of the plot
		If newYPos > (oGrid.getMarginBottom()+ 0 + oGrid.getHeight()) OR newYPos < oGrid.getMarginBottom() + 0 Then ' 0 --> show VBS that this is a number !
			'revert the previous conversion
			Call ChnLinScale(channel.GetReference(eRefTypeIndexIndex),channel.GetReference(eRefTypeIndexIndex),1/scal,0)
			Call ChnLinScale(channel.GetReference(eRefTypeIndexIndex),channel.GetReference(eRefTypeIndexIndex),origscal,0)
			
			MsgBox("Value too small")
			channel.Properties.Item("PlotScal").Value = origscal
		Else
	
			Call GraphObjOpen("TUnit_"&i)
				TxtTxt = channel.Properties.Item("PlotFactor").Value*50/scal&" "&channel.Properties.Item("unit_string").Value&" / DIV" 
				TxtPosY =  oGrid.getMarginBottom() + oGrid.getHeight()/2+plotOff*plotOffStep*offsetSize+getLastPoint(channel.Name)*offsetSize
			Call GraphObjClose("TUnit_"&i)
      
			Call GraphObjOpen("TName_"&i)
				TxtPosY =  oGrid.getMarginBottom() + oGrid.getHeight()/2+plotOff*plotOffStep*offsetSize+getLastPoint(channel.Name)*offsetSize
			Call GraphObjClose("TName_"&i)
			
    
	
		End If
		
		Call PicUpdate() 
    
		setSettings(channel)
	End Function
	
	Public Function setOffset(channelName, direction, factor)
		dim i,  c, off, setting, channel, offstep, offsetSize, cr, currPosY
    
		Set channel = Data.Root.ChannelGroups(currentChannelGroup.getGroupName()&".plot").Channels(channelName)
    
		'get the channel index in the group
		i =InStr(channel.GetReference(eRefTypeIndexIndex),"/[") 
		cr = Right(channel.GetReference(eRefTypeIndexIndex),Len(channel.GetReference(eRefTypeIndexIndex))-i-1)
		i = CInt(Left(cr,Len(cr)-1))
		
		off = channel.Properties("plotOff").Value
		offstep =  channel.Properties("PlotOffstep").Value
		offsetSize = oGrid.getHeight()/1000
    
		Call GraphObjOpen("TName_"&i)
				currPosY = TxtPosY
			Call GraphObjClose("TName_"&i)
      
		'check if new position is on the plot
		If direction = "down" AND currPosY - offstep*offsetSize*factor >= 0+oGrid.getMarginBottom() Then
			
			for c = 1 to segmentAmount
				Call GraphObjOpen("Segment_"&c)
					Call GraphObjOpen(D2AxisYNam(i))        
					D2AxisYBegin = D2AxisYBegin + offstep*factor
					D2AxisYEnd = D2AxisYEnd + offstep*factor
					Call GraphObjClose(D2AxisYNam(i))
				Call GraphObjClose("Segment_"&c)
			next
			
			Call GraphObjOpen("TName_"&i)
				TxtPosY= TxtPosy - offstep*offsetSize*factor
			Call GraphObjClose("TName_"&i)
			
			Call GraphObjOpen("TUnit_"&i)
				TxtPosY= TxtPosy - offstep*offsetSize*factor
			Call GraphObjClose("TUnit_"&i)
			
		channel.Properties("PlotOff").Value = off - 1*factor
			
		'check if new position is on the plot
		Elseif direction = "up" AND  currPosY+offstep*offsetSize*factor <= 0+oGrid.getMarginBottom()+oGrid.getHeight()  Then
			
			for c = 1 to segmentAmount
				Call GraphObjOpen("Segment_"&c)
					Call GraphObjOpen(D2AxisYNam(i))
						D2AxisYBegin = D2AxisYBegin - offstep*factor
						D2AxisYEnd = D2AxisYEnd - offstep*factor
					Call GraphObjClose(D2AxisYNam(i))
				Call GraphObjClose("Segment_"&c)
			Next
			
			Call GraphObjOpen("TName_"&i)
				TxtPosY= TxtPosy + offstep*offsetSize*factor
			Call GraphObjClose("TName_"&i)
			
			Call GraphObjOpen("TUnit_"&i)
				TxtPosY= TxtPosy + offstep*offsetSize*factor
			Call GraphObjClose("TUnit_"&i)
			
			channel.Properties("PlotOff").Value = off + 1*factor
			
		End If
		
		Call PicUpdate
    
		setSettings(channel)
		
	End Function
	
	
	Public Function setSettings(channel)
		dim setting, index, sett
		index = 0
    
		For each setting in lSettings
		  If setting.getName() = channel.Name Then
			lSettings.RemoveAt(index)     
			Exit For
		  End If
		  index = index + 1
		Next
		
		Set sett = new Setting
		
		sett.setName(channel.Name)
		sett.setOff(channel.Properties.Item("PlotOff").Value)
		sett.setScal(channel.Properties.Item("PlotScal").Value)
		sett.setOffstep(channel.Properties.Item("PlotOffstep").Value)
			
		lSettings.Add(sett)
  
	End Function
	
	Public Function loadSettings()
		dim objFSO, objFile, readPath, readName, readType,  setting, currLine, savedValue, returnValue
		Set objFSO=CreateObject("Scripting.FileSystemObject")
		
		returnValue = FileNameGet("ANY","FileRead","", "*.txt")
		readPath = FileDlgDir
		readName = FileDlgFile
		readType  = FileDlgExt
	  
	  
		If returnValue = "IDCancel" Then
			Exit Function
		End If
	  
		Set objFile = objFSO.OpenTextFile(readPath&readName&readType)

		listSegments.Clear
		lSettings.Clear
		plotDuration = 0
		unitAmount = 0
		  
			If objFile.ReadLine = "#Segments" Then
				currLine = objFile.ReadLine
				Do Until currLine = "#Settings"
					' create Segments
					savedValue = Split(currLine,":")
					currLine = objFile.ReadLine
			  
			  
			  
					Call addSegment(savedValue(0),savedValue(1))
			  
					refreshLBSegment()
					LBSegment.Selection = LBSegment.Items.Count
			  
				Loop 
			
				Do Until objFile.AtEndOfStream
					currLine = objFile.ReadLine
			  
					'create Settings
			  
					savedValue = Split(currLine,":")
			  
					Set setting = new Setting

					setting.setName(savedValue(0))
					setting.setScal(savedValue(1))
					setting.setOffstep(savedValue(2))
					setting.setOff(savedValue(3))
					lSettings.Add(setting)
			  
				Loop
				MsgBox("Layout parameters will be loaded")
			Else
			MsgBox("ERROR!: File doesn't start as expected")
		End If
    
		objFile.Close
    
	End Function
	
	Public Function saveSettings()
		dim objFSO,objFile, savePath, saveName, saveType, segment, setting, returnValue
		Set objFSO=CreateObject("Scripting.FileSystemObject")
		'open file selection dialog box for destination filename
		'returns OK or cancel
		returnValue = FileNameGet("ANY","FileWrite","", "*.txt")
		'returns directory
		savePath = FileDlgDir
		'returns filename
		saveName = FileDlgFile
		'returns file extension
		saveType = FileDlgExt
		
		If returnValue = "IDCancel" Then
			Exit Function
		End If
		
		If objFSO.FileExists(savePath&saveName&saveType) Then
			If Not MSgBox("File already exist, Overwrite?",4,"Save Layout") = 6 Then
				Exit Function
			End If
		End If
		'create new file, true is to overwrite
		'see: https://msdn.microsoft.com/en-us/library/5t9b5c0c(v=vs.84).aspx
		Set objFile = objFSO.CreateTextFile(savePath&saveName&saveType,True)
		'
		objFile.Write "#Segments" & vbCrLf
		
		For each segment in listSegments
			objFile.Write segment.getUnitFactor()&":"&segment.getUnitAmount()&vbCrLf
		Next
		
		objFile.Write "#Settings" & vbCrLf
		
		For each setting in lSettings
			objFile.Write setting.getName()&":"&setting.getScal()&":"&setting.getOffstep()&":"&setting.getOff()&vbCrLf
		Next
		
		objFile.Close
		MsgBox("Layout parameters have been saved")
		
	End Function
	
	
	Public Function callDetailPlot(mode)
		'Note: O1 is a DIAdem global variable and is set in the Detail Plot Dialogue
		dim i
    
		oDetailPlotter.setDetailSettings(lDetailSetting)
		Select Case mode
			Case 0
				Call oDetailPlotter.createDetailPlot("0")
				If Not O1 = "0" Then
					oLayoutPlotter.saveDetailSettings(O1)
				End If
			Case 1
				For i = 1 to LBChannels.MultiSelection.Count
				Call oDetailPlotter.createDetailPlot(LBChannels.MultiSelection.Item(i).Text)
				  
				If Not O1 = "0" Then
					oLayoutPlotter.saveDetailSettings(O1)
				End If
			Next
			
		End Select
		
		Report.Sheets.Item(currentChannelGroup.getGroupName() ).Activate
		
	End Function
	
	
	Public Function saveDetailSettings(channelName)
		dim channel, element, oDetailSetting
		
		Set channel = Data.Root.ChannelGroups(currentChannelGroup.getGroupName()).Channels(channelName)
		
		For each element in lDetailSetting
			If element.getName() = channel.Name Then
				lDetailSetting.Remove(element)
				Exit For
			End If
		Next
		
		Set oDetailSetting = new DetailSetting
		
		oDetailSetting.setName(channel.Name)
		oDetailSetting.setTimeUnit(channel.Properties("DetailPlot_TimeUnit").Value)
		oDetailSetting.setStartTime(channel.Properties("DetailPlot_StartTime").Value)
		oDetailSetting.setDuration(channel.Properties("DetailPlot_Duration").Value)
		oDetailSetting.setMinValue(channel.Properties("DetailPlot_MinValue").Value)
		oDetailSetting.setMaxValue(channel.Properties("DetailPlot_MaxValue").Value)
		
		lDetailSetting.Add(oDetailSetting)
		
	End Function
	
End Class
