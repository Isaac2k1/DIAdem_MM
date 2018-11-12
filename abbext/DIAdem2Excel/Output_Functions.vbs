'-------------------------------------------------------------------------------
'-- VBS-Script-File
'-- Created: 2009-08-07 
'-- Authors:	Rene Irion 
'				Jonas Schwammberger
'-- Version: 1.9

'-- Purpose: Main file of the DIAdem2Excel Application.
'-- History:

'-------------------------------------------------------------------------------
option explicit


'-------------------------------------------------------------------------------
'summary: safing a .pdf mechanical report to a specific location
'parameter templatePath: path of template .TDR file
'parameter outFile: file to safe .pdf file
'parameter strGroup: group with the mechanical channels
'output: none
Sub CreateMechReport(templatePath, outFile, outName, strGroup, ChannelNames,NewChannelNames,MaxNumbOfChn,DiagramIndex)
  Dim boolSuccess                         'true if function was successful
  Dim intSheetNo                          'sheet
  Dim intSheetNoMax
  Dim intCurveIndex
  Dim intCurveNoMax
  dim i
  dim actualChannelNr
  Dim ActualChannelIndex()
  
  Dim fso
  Dim folder

  Set fso = CreateObject("Scripting.FileSystemObject")

  If (fso.FolderExists(outputFolder & "\PDF")) Then

  Else
    Set folder = fso.CreateFolder(outputFolder & "\PDF")
  End If
  
  redim ActualChannelIndex(0)
  actualChannelNr = 0
  
  'get actual existing channels. if it exists, write the index in actualChannelIndex()
  For i = 0 to MaxNumbOfChn
    if Data.Root.ActiveChannelGroup.Channels.Exists(NewChannelNames(i)) = true Then
      actualChannelNr = actualChannelNr + 1
      redim preserve ActualChannelIndex(actualChannelNr)
      ActualChannelIndex(actualChannelNr-1) = i
    end if
  next
  
  'Set Title
  Call Data.Root.Properties.Item("description").Value("Report of mechanical shot: " & outName)
  
  'get maximum number of sheets
  intSheetNoMax = 0
  
  For i = 0 to MaxNumbOfChn
    if DiagramIndex(i) > intSheetNoMax then
      intSheetNoMax = DiagramIndex(i)
    end if
  Next

  'loop over all report sheets
  For intSheetNo = 1 to intSheetNoMax               'intSheetNoMax = max. no. of diagrams
    intCurveIndex = 0
    
    If intSheetNo = 1 Then                        'Load Template
      Call PicLoad(templatePath)
    Else
      Call PicFileAppend(templatePath)            'Add Sheet from template
    End If    
     
    Call PicUpdate(0)                             '... PicDoubleBuffer 
     
    '------------- open plot -----------------
    Call GraphObjOpen("2DAxis1")                  '2DAxis --> Name of the diagramm
                                                  '1      --> Because it is the first diagram on this report

    '------------------- Curve list -------------------------------
    For i = 1 to MaxNumbOfChn                        'intCurveNoMax = max. no. of curves on the second y-axis
    
    'D2LegTxtTypeA(i) = "Free text"
    'D2LegTxtFreeA(i) = NewChannelNames(i)
    
      If DiagramIndex(i) = 0 Then
        Call  GraphObjNew("2D-Curve","Curve_" & intCurveIndex)    'Plot the travel curve on the first y-axis   
        Call GraphObjOpen("Curve_" & intCurveIndex)              
          D2CChnExpand     = 1                                
          D2CChnYName      = NewChannelNames(i)           'Array name of travel curve
          D2CAxisPairNo    = 1
        Call GraphObjClose("Curve_" & intCurveIndex)           
       
        'Call GraphObjOpen("2DYAxis4_1")
        '  D2AxisYTxt       = NewChannelNames(i)
        'Call GraphObjClose("2DYAxis4_1")                  
        
        Call PicUpdate(0)  'maybe delete that
        intCurveIndex = intCurveIndex + 1
                  
      Elseif VAL(DiagramIndex(i)) = intSheetNo Then
        Call  GraphObjNew("2D-Curve","Curve_" & intCurveIndex)     'Plot other curves on the second y-axis
        Call GraphObjOpen("Curve_" & intCurveIndex)            
          D2CChnExpand     = 1                              
          D2CChnYName      = NewChannelNames(i)                               'Array name of other curves
          D2CAxisPairNo    = 2
        Call GraphObjClose("Curve_" & intCurveIndex)           
        
        'Call GraphObjOpen("2DYAxis4_2")
        '  D2AxisYTxt       = NewChannelNames(i)
        'Call GraphObjClose("2DYAxis4_2")     
        
        Call PicUpdate(0) 'maybe delete that
        intCurveIndex = intCurveIndex + 1
      Else
        
      End If
        
    Next

    '------------- close plot -----------------
    Call GraphObjClose("2DAxis1")
              
    '------------- update plot -----------------
    Call PicUpdate(0)
    
    '------------- rename sheet -----------------
    Call GraphSheetNGet(intSheetNo)
    Call GraphSheetRename(GraphSheetName,"Page " & intSheetNo)
    
  Next
  
  Call PicPDFExport(outputFolder & "\PDF\" & outFile & ".pdf",0)             'Save as PDF
  
  Call GraphDeleteAll()

end Sub
'-------------------------------------------------------------------------------



'-------------------------------------------------------------------------------
'summary: safing a .pdf power report to a specific location
'parameter templatePath: path of template .TDR file
'parameter outPath: file to safe .pdf file
'parameter strGroup: group with the power channels
Sub CreatePowerReport(templatePath, outPath, outputName, strGroup, ChannelNames,NewChannelNames,MaxNumbOfChn,DiagramIndex)
  Dim boolSuccess                         'true if function was successful
  Dim intSheetNo
  Dim intSheetNoMax
  Dim intCurveIndex
  Dim intCurveNoMax
  Dim i
  Dim fso
  Dim folder
  dim actualChannelNr
  Dim ActualChannelIndex()

  Set fso = CreateObject("Scripting.FileSystemObject")

  If (fso.FolderExists(outputFolder & "\PDF")) Then

  Else
    Set folder = fso.CreateFolder(outputFolder & "\PDF")

  End If
  
  'get actual existing channels. if it exists, write the index in actualChannelIndex()
  For i = 0 to MaxNumbOfChn
    if Data.Root.ActiveChannelGroup.Channels.Exists(NewChannelNames(i)) = true Then
      actualChannelNr = actualChannelNr + 1
      redim preserve ActualChannelIndex(actualChannelNr)
      ActualChannelIndex(actualChannelNr-1) = i
    end if
  next
  
  'Set Title
  Call Data.Root.Properties.Item("description").Value("Report of power shot: " & outputName)
  
  'get maximum number of sheets
  intSheetNoMax = 0
  
  For i = 0 to MaxNumbOfChn
    if DiagramIndex(i) > intSheetNoMax then
      intSheetNoMax = DiagramIndex(i)
    end if
  Next

  'loop over all report sheets
  For intSheetNo = -1 to intSheetNoMax             'intSheetNoMax = max. no. of diagrams
    intCurveIndex = 0
    
    If intSheetNo = -1 Then                        'Load Template
      Call PicLoad(templatePath)
    Else
      Call PicFileAppend(templatePath)            'Add Sheet from template
    End If
	  
    Call PicUpdate(0)                             '... PicDoubleBuffer 

    '------------- open plot -----------------
    Call GraphObjOpen("2DAxis1")                  '2DAxis --> Name of the diagramm
    
    '------------------- Curve list -------------------------------
    For i = 1 to MaxNumbOfChn                     'intCurveNoMax = max. no. of curves on the second y-axis
        if intSheetNo = -1 Then 
        if VAL(DiagramIndex(i)) = 0 Then
          Call  GraphObjNew("2D-Curve","Curve_" & intCurveIndex)    'Plot the travel curve on the first y-axis   
          Call GraphObjOpen("Curve_" & intCurveIndex)              
            D2CChnExpand     = 1                                
            D2CChnYName      = NewChannelNames(i)                      'Array name of travel curve
            D2CAxisPairNo    = 1
          Call GraphObjClose("Curve_" & intCurveIndex)
          
          'Call GraphObjOpen("2DYAxis4_1")
          '  D2AxisYTxt       = NewChannelNames(i)
          'Call GraphObjClose("2DYAxis4_1") 
          
          Call PicUpdate(0)'maybe delete that
          intCurveIndex = intCurveIndex + 1
        end if
        
        if VAL(DiagramIndex(i)) = -2 Then
          Call  GraphObjNew("2D-Curve","Curve_" & intCurveIndex)     'Plot other curves on the second y-axis
          Call GraphObjOpen("Curve_" & intCurveIndex)            
            D2CChnExpand     = 1                              
            D2CChnYName      = NewChannelNames(i)                               'Array name of other curves
            D2CAxisPairNo    = 2
          Call GraphObjClose("Curve_" & intCurveIndex)
          
          'Call GraphObjOpen("2DYAxis4_2")
          '  D2AxisYTxt       = NewChannelNames(i)
          'Call GraphObjClose("2DYAxis4_2") 
          
          Call PicUpdate(0)'maybe delete that
          intCurveIndex = intCurveIndex + 1
        end if
        
      Elseif intSheetNo = 0 then
        if VAL(DiagramIndex(i)) = -1 Then
          Call  GraphObjNew("2D-Curve","Curve_" & intCurveIndex)    'Plot the travel curve on the first y-axis   
          Call GraphObjOpen("Curve_" & intCurveIndex)              
            D2CChnExpand     = 1                                
            D2CChnYName      = NewChannelNames(i)                      'Array name of travel curve
            D2CAxisPairNo    = 1
          Call GraphObjClose("Curve_" & intCurveIndex)
          
          'Call GraphObjOpen("2DYAxis4_1")
          '  D2AxisYTxt       = NewChannelNames(i)
          'Call GraphObjClose("2DYAxis4_1") 
          
          Call PicUpdate(0)'maybe delete that
          intCurveIndex = intCurveIndex + 1
        end if
        
        if VAL(DiagramIndex(i)) = -3 Then
          Call  GraphObjNew("2D-Curve","Curve_" & intCurveIndex)     'Plot other curves on the second y-axis
          Call GraphObjOpen("Curve_" & intCurveIndex)            
            D2CChnExpand     = 1                              
            D2CChnYName      = NewChannelNames(i)                               'Array name of other curves
            D2CAxisPairNo    = 2
          Call GraphObjClose("Curve_" & intCurveIndex)
          
          Call PicUpdate(0)'maybe delete that
          intCurveIndex = intCurveIndex + 1
        end if
        
      Else
        If VAL(DiagramIndex(i)) = 0 Then
          Call  GraphObjNew("2D-Curve","Curve_" & intCurveIndex)    'Plot the travel curve on the first y-axis   
          Call GraphObjOpen("Curve_" & intCurveIndex)              
            D2CChnExpand     = 1                                
            D2CChnYName      = NewChannelNames(i)           'Array name of travel curve
            D2CAxisPairNo    = 1
          Call GraphObjClose("Curve_" & intCurveIndex)
          
          'Call GraphObjOpen("2DYAxis4_1")
          '  D2AxisYTxt       = NewChannelNames(i)
          'Call GraphObjClose("2DYAxis4_1") 
          
          Call PicUpdate(0)'maybe delete that
          intCurveIndex = intCurveIndex + 1
          
        Elseif VAL(DiagramIndex(i)) = intSheetNo Then
          Call  GraphObjNew("2D-Curve","Curve_" & intCurveIndex)     'Plot other curves on the second y-axis
          Call GraphObjOpen("Curve_" & intCurveIndex)            
            D2CChnExpand     = 1                              
            D2CChnYName      = NewChannelNames(i)                               'Array name of other curves
            D2CAxisPairNo    = 2
          Call GraphObjClose("Curve_" & intCurveIndex)
          
          'Call GraphObjOpen("2DYAxis4_2")
          '  D2AxisYTxt       = NewChannelNames(i)
          'Call GraphObjClose("2DYAxis4_2") 
          
          Call PicUpdate(0)'maybe delete that
          intCurveIndex = intCurveIndex + 1
        End If
      end if
    Next

    '------------- close plot -----------------
    Call GraphObjClose("2DAxis1")        
    '------------- update plot -----------------
    Call PicUpdate(0)
    '------------- rename sheet -----------------
    Call GraphSheetNGet(intSheetNo+2)
    Call GraphSheetRename(GraphSheetName,"Page " & intSheetNo+2)
  Next
  
  Call PicPDFExport(outputFolder & "\PDF\" & outPath & ".pdf",0)             'Save as PDF
  Call GraphDeleteAll()
end Sub
'-------------------------------------------------------------------------------



'-------------------------------------------------------------------------------
'summary: Export the data to a excel file
'parameter outFile: path to safe .xls file to
'parameter strGroup: group name of the channels
'parameter strChannels: channels to export
'parameter numberOfChannels: number of channels in strChannels
Sub CreateExcelFile(outFile,strGroup,strChannels(),numberOfChannels)
  dim i
  Dim fso
  Dim folder
  Dim actualChnNr
  Dim actualChn()
  'dim EXCELChnCount     'הההההההה
  'dim ExcelExpSheetChn  'ההההההההה
  'dim ExcelExpChn()     'ההההההההה
 ' msgbox("blubb")
  actualChnNr = 0
  redim actualChn(0)
  
  'create output folder if it doesn't exist
  Set fso = CreateObject("Scripting.FileSystemObject")
  
 ' dim strBleistift
  'strBleistift = "outfile: " &outFile&" strGroup: "&strGroup&" numberOfChannels: "& numberOfChannels
 ' for i=1 to numberOfChannels
 '   strBleistift=strBleistift & " Channel" & i & ": " & strChannels(i)
 ' next
 ' msgbox(strBleistift)
  
  If (fso.FolderExists(outputFolder & "\Excel")) Then
  Else
    Set folder = fso.CreateFolder(outputFolder & "\Excel")
  End If
  
  'msgbox(folder)
  
  Data.Root.ChannelGroups(strGroup).Activate()
  
  'get actual number of Channels
  for i = 1 to numberOfChannels
    if Data.Root.ActiveChannelGroup.Channels.Exists(strChannels(i)) = true Then
      actualChnNr = actualChnNr+1
      redim preserve actualChn(actualChnNr)
      actualChn(actualChnNr-1) = strChannels(i)
    end if    
  next
  
  EXCELChnCount = actualChnNr
  ExcelExpSheetChn = "DIAdem"       'this variable HAS to be set. Otherwise the export doesn't work
  
 ' redim preserve ExcelExpChn(actualChnNr + 1)  'ההההההההההההההה
  
  For i = 0 To actualChnNr          'we're skipping "VS"
    ExcelExpChn(i+1) = CNo(strGroup&"/"&actualChn(i))
  Next
  
  Call ExcelExport(outputFolder & "\Excel\" & outFile & ".xls", "DIAdem", 0, "")
  
End Sub
'--------------------------------------------------------------------------------


'-------------------------------------------------------------------------------
'summary: Safe an .sdf file to the specific location.
'  it is also putting milliseconds as first column.
'output: none
Sub CreateSDFFile(outFile,isMech,strGroup,PowNewChnName,PowChannels,PowNumbOfChn,powColumn,mechNewChnName,mechChannels,mechNumbOfChn,mechColumn,samplingRate)
  Dim fso                     'FileSystemObject handle
  Dim file                    'sdf file
'  Dim intErr                  'maybe not needed
  Dim i,j                     'index
  Dim line                    'one string line
  Dim folder                  'path to SDF
  Dim boolSuccess             'true if function was successful
'  Dim actualChnNr             
  dim channelLength           'length of the channels
  dim maxColumn               
  dim milliseconds
  
  Dim strChannels()           'names of Channels
  Dim strNewChannelNames()    'name in sdf
  Dim isZeros()               'empty column as boolean
  
  dim debug
  
 ' actualChnNr = 0

  'create output folder if it doesn't exist
  Set fso = CreateObject("Scripting.FileSystemObject")  
  If not (fso.FolderExists(outputFolder & "\SDF")) Then
    Set folder = fso.CreateFolder(outputFolder & "\SDF")
  End If
  
  Data.Root.ChannelGroups(strGroup).Activate()
  
  'create sdf file
  Set file = fso.CreateTextFile(outputFolder & "\SDF\" & outFile & ".sdf", True)
  
  maxColumn = 0

  for i=1 to PowNumbOfChn                                   'find highest column number 
    if val(powColumn(i)) > maxColumn then
      maxColumn = val(powColumn(i))
    end if
  next

  redim strChannels(maxColumn)                              'redefine and clear
  redim strNewChannelNames(maxColumn)
  redim isZero(maxColumn)

  for i = 1 to maxColumn                                    'set default values to each column
    isZero(i) = true
    strChannels(i) = "none"
    strNewChannelNames(i) = "empty"
  next

  'sort for new columns
  for i = 2 to maxColumn    
    for j = 1 to PowNumbOfChn
      if val(powColumn(j)) = i Then
        strChannels(i) = PowChannels(j)
        strNewChannelNames(i) = PowNewChnName(j)
        'msgbox(strNewChannelNames(i) & " in column " & i)
        isZero(i) = true
      end if
    next
  next
  
  'paste mech definition if it is a mech shot
  if isMech = true then
    for i = 1 to mechNumbOfChn
      strChannels(val(MechColumn(i))) = MechChannels(i)
      strNewChannelNames(val(MechColumn(i))) = MechNewChnName(i)
      
      'use channelvalues instead of zeros if channel exists
      if Data.Root.ActiveChannelGroup.Channels.Exists(MechChannels(i)) = true then
        isZero(val(MechColumn(i))) = false
      end if
    next
  else
    for i = 2 to maxColumn 'actualChnNr
      'use channelvalues instead of zeros if channel exists
      if Data.Root.ActiveChannelGroup.Channels.Exists(strChannels(i)) = true then
        'msgbox(strChannels(i) & " exist")
        isZero(i) = false
      end if
    next
  end if
  
  'write header, skip vs channel and get actual number of Channels
  line = "#Time_[s]: "
  for i = 2 to maxColumn 'actualChnNr
    line = line & strNewChannelNames(i) & " "
  next  
  file.WriteLine(line)
    
  line = ""
  
  if isMech = true then
    channelLength = chnLength(strGroup&"/"&MechChannels(0))
  Else
    channelLength = chnLength(strGroup&"/"&PowChannels(0))
  end if
  
  'write channel values to file
  for i = 1 to channelLength
    
    milliseconds = (i-1) * samplingRate
    line = milliseconds & " "
    
    'write one complete line
    for j = 2 to maxColumn

      debug = strChannels(j)
      if isZero(j) Then
        line = line & "0 "
      else
        line = line & CHD(i,strGroup&"/"&strChannels(j)) & " "
      end if

    next

    Trim(line)
    file.WriteLine(line)
    line = ""
  next
  
  file.Close()
End Sub
'-------------------------------------------------------------------------------



'-------------------------------------------------------------------------------
'summary: Export the data to a excel file
'parameter outFile: path to safe .xls file to
'parameter strGroup: group name of the channels
'parameter strChannels: channels to export
'parameter numberOfChannels: number of Channels in strChannels
sub WriteErrorReport(fileContent)
  dim fso
  dim file
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set file = fso.CreateTextFile(outputFolder & "\ERROR.txt", True)
  file.WriteLine(fileContent)
  file.Close()
end sub
'-------------------------------------------------------------------------------