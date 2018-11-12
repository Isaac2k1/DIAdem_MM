Option Explicit 
'===============================================================================
' >>> STANDARD DataPlugin for AMO-SDF V3 files <<<
' Version       :   1.5 
' Created       :   February 28, 2005
' Last Change   :   February 15, 2011
'-------------------------------------------------------------------------------
' Version history :
' 1.0   created
' 1.01  fix: dDeltaT/CDbl(10.^9) >>> dDeltaT/CDbl(10.^8) 
' 1.1   support for minor changes in the header added (irmen@amo.de)
' 1.2   support for digital channels added (irmen@amo.de)
' 1.3   function GetMinMax added, reducedData processing disabled (irmen@amo.de)
' 1.4   mapping of ascii header to channel properties enhanced, adding 
'       properties to channel not to group (irmen@amo.de)
' 1.5   added more comments, speed improvements (schoenitz@amotronics.de)
'-------------------------------------------------------------------------------

'===============================================================================
' ReadStore
'-------------------------------------------------------------------------------
' Main function called by DIAdem
' oFile:       file object for accessing the file which is currently processed
' bIndexing:   determines whether the file is currently being indexed or
'              loaded: TRUE =  Indexing, FALSE = Loading
'-------------------------------------------------------------------------------
Sub ReadStore(oFile, bIndexing)
  
  '>>> IMPORTANT NOTE <<<
  'The following line has to be disabled in the DataPlugin AMO_SDF_V3 
  '(standard) and has to be enabled in the DataPlugin AMO_SDF_V3_EXTENDED.
  'This prevents the DataFinder from indexing the SDF files twice.
  'Explanation for RaiseError: Calling RaiseError without a message text tells 
  'the DataFinder that the current file should not be indexed with this 
  'DataPlugin.
  'if bIndexing then call RaiseError
  
  Dim oChannelGroup,oBlock,oBlkChannel,oUSIChannel
  Dim lHeaderVersion,lCard,lModule,lChannel,lIsDigital,dYOffset,dYGain
  Dim lSampleCount,sgChannelName, lPrecision, lOffset, lFastScanRecScans
  Dim dTimeZero,dDeltaT,lReduction,sgFilename,lIndex
  Dim lValue,lChannelMin,lChannelMax
  Dim dChannelMin,dChannelMax,fCopyReduced,sgUnitName,dChannelIdx
  Dim sgChannelNameBase
  Dim sgOrigin,sgExperimentDate, sgProject, sgInputRange, sgDescription
  Dim sgUnitX, sgUnitY
  Dim sgCoupling, sgTermination, sgMode, sgChannelId

  ' get file information
  sgFilename = oFile.Info.Filename & oFile.Info.Extension
  ' set general file formatting
  oFile.Formatter.ByteOrder = eLittleEndian

  ' create new channel group
  Set oChannelGroup = Root.ChannelGroups.Add(oFile.Info.FileName) 

  ' read meta data from file header
  call AMOReadHeader(oFile,oChannelGroup,lHeaderVersion,lCard,lModule,_
                     lChannel,lIsDigital,dYOffset,dYGain,lSampleCount,_
                     dTimeZero,dDeltaT,sgChannelName,sgUnitName,lPrecision,_
                     sgOrigin,sgExperimentDate,sgProject,sgInputRange,_
                     sgDescription,sgUnitX,sgUnitY,sgCoupling,sgTermination,_
                     sgMode,sgChannelId,lOffset,lFastScanRecScans)
  
  ' process digital and ADC channel data differently
  if ( lIsDigital > 0 ) Then 
    ' digital => each bit represents one channel
    sgChannelNameBase= sgChannelName & "_" 'base name for all channels
    ' add digital channels
    For dChannelIdx = 0 to (lPrecision - 1)
      ' define data block in file starting at the current position
      Set oBlock = oFile.GetBinaryBlock()
      oBlock.BlockLength = lSampleCount '# of samples per channel
      sgChannelName= sgChannelNameBase & dChannelIdx 'channel name

      ' define new channel in block
      Set oBlkChannel = oBlock.Channels.Add(sgChannelName, eU16)
      ' set waveform parameters for x axis
      call MakeWaveformEx(oBlkChannel,dTimeZero,dDeltaT,"Time","s")
     
      ' set channel formatting parameters
      ' use bit mask, factor and offset for extracting the digital channel
      oBlkChannel.Formatter.Bitmask= 2 ^ dChannelIdx
      oBlkChannel.Factor = 1. / (2 ^ dChannelIdx)
      oBlkChannel.Offset = 0 
      
      ' add channel to channel group
      Set oUSIChannel = oChannelGroup.Channels.AddDirectAccessChannel(oBlkChannel) 
      ' add unit to channel properties
      call oUSIChannel.Properties.Add("unit_string",sgUnitName)
    Next 'digital channel
    '-------------------------------------------
  else
    '-------------------------------------------
   ' ADC => only one analog channel
   ' define data block in file starting at the current position
    Set oBlock = oFile.GetBinaryBlock()
    oBlock.BlockLength = lSampleCount '# of samples (important because of 
                                      'additionally stored reduced data for
                                      'faster browsing in Saturn Studio II)

    ' take FASTSCAN mode into account while defining new channel in block
    if ( lFastScanRecScans > 0 ) Then
      ' FASTSCAN mode = different datatype, factor
      Set oBlkChannel = oBlock.Channels.Add(sgChannelName, eU32)
      ' set channel formatting parameters
      ' (will be used to automatically convert stored data into physical unit)
      oBlkChannel.Factor = dYGain / lFastScanRecScans 'mean of all scans
      oBlkChannel.Offset = dYOffset
      '>>> dDeltaT = dDeltaT / 4.0 '???
    else
      ' ordinary scan mode
      Set oBlkChannel = oBlock.Channels.Add(sgChannelName, eU16)
      ' set channel formatting parameters
      ' (will be used to automatically convert stored data into physical unit)
      oBlkChannel.Factor = dYGain
      oBlkChannel.Offset = dYOffset
    end if

    ' set waveform parameters for x axis
    call MakeWaveformEx(oBlkChannel,dTimeZero,dDeltaT,"Time","s")
    
    ' add channel to channel group
    set oUSIChannel = oChannelGroup.Channels.AddDirectAccessChannel(oBlkChannel) 
    ' add unit to channel properties
    call oUSIChannel.Properties.Add("unit_string",sgUnitName)

    ' shift file pointer to end of channel (eU16 = 2 bytes)
    oFile.Position = oFile.Position + 2*lSampleCount

    ' set channel characteristic properties which speed up loading time
    call GetMinMax(oFile, dYOffset,dYGain, dChannelMin, dChannelMax)
    call oUSIChannel.Properties.Add("minimum",dChannelMin)
    call oUSIChannel.Properties.Add("maximum",dChannelMax)
    call oUSIChannel.Properties.Add("novaluekey","No")
    call oUSIChannel.Properties.Add("monotony","not monotone")

    '-------------------------------------------
  end if 'digital or ADC

  ' add Properties to channel and not to group
  Dim prefix
  prefix = "amo_"
  call oUSIChannel.Properties.Add(prefix & "Origin", sgOrigin)
  call oUSIChannel.Properties.Add(prefix & "Project", sgProject)
  call oUSIChannel.Properties.Add(prefix & "ExperimentDate", sgExperimentDate)
  call oUSIChannel.Properties.Add(prefix & "InputRange", sgInputRange)
  call oUSIChannel.Properties.Add(prefix & "UnitX", sgUnitX)
  call oUSIChannel.Properties.Add(prefix & "UnitY", sgUnitY)
  call oUSIChannel.Properties.Add(prefix & "Coupling", sgCoupling)
  call oUSIChannel.Properties.Add(prefix & "Termination", sgTermination)
  call oUSIChannel.Properties.Add(prefix & "Mode", sgMode)
  call oUSIChannel.Properties.Add(prefix & "ChannelId", sgChannelId)
  call oUSIChannel.Properties.Add(prefix & "Description", sgDescription)
  call oUSIChannel.Properties.Add(prefix & "Precision", lPrecision)
  call oUSIChannel.Properties.Add(prefix & "RangeOffset", lOffset)
  call oUSIChannel.Properties.Add(prefix & "Offset", dYOffset)
  call oUSIChannel.Properties.Add(prefix & "Gain", dYGain)
  call oUSIChannel.Properties.Add(prefix & "FastScanRecords", lFastScanRecScans)

End Sub 'ReadStore
'-------------------------------------------------------------------------------
' END OF MAIN FUNCTION
'===============================================================================


'===============================================================================
' AMOReadHeader
'-------------------------------------------------------------------------------
' Function for processing the meta data contained in the AMO header.
'-------------------------------------------------------------------------------
Function AMOReadHeader(oFile,oChannelGroup,lHeaderVersion,lCard,lModule,_
         lChannel,lIsDigital,dYOffset,dYGain,lSampleCount,dTimeZero,dDeltaT,_
         sgChannelName,sgUnitName,lPrecision,sgOrigin,sgExperimentDate,_
         sgProject,sgInputRange,sgDescription,sgUnitX,sgUnitY,sgCoupling,sgTermination,sgMode,_
         sgChannelId, lOffset, lFastScanRecScans)
  Dim lPreviewIncluded,sgHeader,lDelPos
  Dim dMeasClk, lClock1,lClock2, lFlags
  Dim lToken,sgaTokens,sgName,sgValue,dStartTime,dTimeShift,sgaNames(),sgaValues()
  
  ' initialization
  AMOReadHeader = False
  sgChannelName = "Y-Channel"
  
  ' read header version from file
  lHeaderVersion = oFile.GetNextBinaryValue(eByte)
    
  ' reject file if not correct version (processing of file will stop here)
  if lHeaderVersion <> 3 then call RaiseError

  ' read precision from file
  lPrecision = oFile.GetNextBinaryValue(eByte)
  
  ' depending on precision proceed with old or new version of SDFV3 format
  if ( lPrecision = 0 ) Then
    '-------------------------------------------------
    ' OLDER SDFV3 FORMAT !!! -> less properties and different data types
    ' skip empty bytes and preset missing properties
    oFile.GetCharacters(2)
    sgCoupling = ""
    sgTermination = ""
    sgMode = ""
    lOffset = 0
    
    ' proceed with reading meta data
    lModule = oFile.GetNextBinaryValue(eU32) 'module number
    lChannel = oFile.GetNextBinaryValue(eU32) 'channel  number
    oFile.GetCharacters(4) 'skip 4 empty bytes (alignment)

    ' scaling (factor and offset which have to be applied to all stored data)
    dYOffset = oFile.GetNextBinaryValue(eR64) 'voltage offset in Volts
    dYGain = oFile.GetNextBinaryValue(eR64) 'gain specified as a factor

    lSampleCount = CLng(oFile.GetNextBinaryValue(eU32)) 'number of samples
    oFile.GetCharacters(4) 'skip 4 empty bytes (alignment)
    
    ' support for a few new values in the header:
    ' uiMeasClk contains the sample rate in in terms of 100MHz clock cycles
    ' Note: in this older SDFV3 version uiMeasClk is a 64-bit unsinged integer!
    ' scanrate = uiMeasClk/1024
    lClock1 = oFile.GetNextBinaryValue(eU32) 'part 1 of 64 bit
    lClock2 = oFile.GetNextBinaryValue(eU32) 'part 2 of 64 bit
    dDeltaT = CDbl(lClock1)/1024. + (CDbl(lClock2)/1024.)/CDbl(2^32)
    dDeltaT = dDeltaT/CDbl(10.^8) '100MHz cycles
  else
    '-------------------------------------------------
    ' CURRENT SDFV3 FORMAT (still with btVersion= 3)
    lOffset     = oFile.GetNextBinaryValue(eByte) 'range offset
    lFlags      = oFile.GetNextBinaryValue(eByte) 'modus
    ' add the flags as properties
    if ( lFlags AND 1 ) Then
      sgCoupling = "AC"
    else
      sgCoupling = "DC"
    end if

    if ( lFlags AND 2 ) Then
      sgTermination = "1 MOhm"
    else
      sgTermination = "50 Ohm"
    end if
    
    if ( lFlags AND 4 ) Then
      sgMode = "single"
    else
      sgMode = "differential"
    end if
    
    'read and add further properties
    lCard = oFile.GetNextBinaryValue(eByte) 'card number
    lModule = oFile.GetNextBinaryValue(eByte) 'module number
    lChannel = oFile.GetNextBinaryValue(eByte) 'channel number
    lIsDigital = oFile.GetNextBinaryValue(eByte) '0 = ADC channel, 
                                                  '1 = digital (1 channel per bit)
    oFile.GetCharacters(8) 'skip 4 empty bytes (alignment)
    
    ' scaling (factor and offset which have to be applied to all stored data)
    dYOffset = oFile.GetNextBinaryValue(eR64) 'voltage offset in Volts
    dYGain = oFile.GetNextBinaryValue(eR64) 'gain specified as a factor

    lSampleCount = CLng(oFile.GetNextBinaryValue(eU32)) 'number of samples
    oFile.GetCharacters(4) 'skip 4 empty bytes (alignment)
    
    dMeasClk = oFile.GetNextBinaryValue(eR64) 'sample period in sec
    dDeltaT = dMeasClk 'sample period in sec
  end if 'lPrecision = 0 (old or new SDFV3 format)

  '-------------------------------------------------
  'read rest of meta data for old and new SDFV3 files
  dStartTime  = oFile.GetNextBinaryValue(eR64) 'time of first measurement point
                                               'after trigger
  dTimeShift  = oFile.GetNextBinaryValue(eR64) 'internal value - don't use!
  dTimeZero   = dStartTime 'do not use timeshift here!
  
  lFastScanRecScans = CLng(oFile.GetNextBinaryValue(eU32)) '# of recorded scans
                                                           'in FASTSCAN mode
  lPreviewIncluded = CLng(oFile.GetNextBinaryValue(eI32)) 'preview flag
                                                          '0 = no preview included
                                                          '1 = 1000 points preview

  '-------------------------------------------------
  'read misc. meta data for channel in ASCII format
  sgHeader = oFile.GetCharacters(256)
  ' Parse text to get single tokens. Tokens are seperated by a LF (0A) Byte.
  ' Each Token is given as a name/value pair with "Tab" (09) as a delimiter.
  sgaTokens = Split(sgHeader,Chr(10)) 'separation of tokens 
  ReDim sgaNames(Ubound(sgaTokens))
  ReDim sgaValues(Ubound(sgaTokens))
  For lToken = Lbound(sgaTokens) To Ubound(sgaTokens) 
    lDelPos = InStr(1,sgaTokens(lToken),Chr(9)) 'separation of name and value
    if ( 1 < lDelPos ) Then
      sgName  = Left(sgaTokens(lToken),lDelPos-1)
      sgValue = Right(sgaTokens(lToken),Len(sgaTokens(lToken))-lDelPos)
      sgaNames(lToken)  = sgName
      sgaValues(lToken) = Trim(sgValue)
    end if 
  Next
  '-------------------------------------------------
  ' Try to get more meta information :
  ' Row   Label           Contents      Example
  '  0    -               Headerinfo    "Saturn Data-Export"
  '  1    -               date & time   "25.08 - 15:16"
  '  2    "Project:"      Project name  "Saturn Export (no Name)"
  '  3    "Channel:"      channel name  "S1M1C1"
  '  4    "Name:"         xxxxxxx       ""
  '  5    "Input-Range:"  voltage range "2,00 V"
  '  6    "time"          xxxx
  '  7    "s"             xxxxx
  '-------------------------------------------------
  if ( 7 <= UBound(sgaTokens) ) Then
    if ( 0 < Len(sgaTokens(0)) ) Then sgOrigin = sgaTokens(0)
    if ( 0 < Len(sgaTokens(1)) ) Then sgExperimentDate = sgaTokens(1)
    if ( 0 < Len(sgaValues(2)) ) Then sgProject = sgaValues(2)
    if ( 0 < Len(sgaValues(3)) ) Then sgChannelId = sgaValues(3)
    ' use ChannelId as presetting for channel name in case name has not been set by user
    if ( 0 < Len(sgaValues(3)) ) Then sgChannelName = sgaValues(3)
    if ( 0 < Len(sgaValues(4)) ) Then sgChannelName = sgaValues(4)
    if ( 0 < Len(sgaValues(5)) ) Then sgInputRange = sgaValues(5)
    if ( 0 < Len(sgaValues(6)) ) Then sgDescription = sgaValues(6)
    if ( 0 < Len(sgaNames(7)) ) Then  sgUnitX = sgaNames(7)
    if ( 0 < Len(sgaValues(7)) ) Then  sgUnitY = sgaValues(7)
    sgUnitName = CStr(sgaValues(7))
  end if
 
  AMOReadHeader = True
End Function 'AMOReadHeader
'-------------------------------------------------------------------------------


'===============================================================================
' MakeWaveformEx
'-------------------------------------------------------------------------------
' Rev 1.0, February 4, 2005
' Append waveform information to an existing channel. DIAdem and LabVIEW 
' automatically load this information and create waveform signals. 
' If needed there are functions to create an additional explicit time channel
' if needed. Data should be provided in units of "seconds".
'-------------------------------------------------------------------------------
Function  MakeWaveformEx(oChannel,dOffset,dIncrement,sgXName,sgXUnit)
  MakeWaveformEx = False
  On Error Resume Next
  call oChannel.Properties.Add("wf_start_offset",CDbl(dOffset))
  call oChannel.Properties.Add("wf_increment",CDbl(dIncrement))
  call oChannel.Properties.Add("wf_xname",sgXName)
  call oChannel.Properties.Add("wf_xunit_string",sgXUnit)
  call oChannel.Properties.Add("wf_samples",CLng(1)) 'same as # of Y channels
                                                     'only needed if several
                                                     'scans in one channel
  call oChannel.Properties.Add("wf_time_pref","relative")
  call oChannel.Properties.Add("wf_start_time",CreateTime(0,1,1,0,0,0,0,0,0))
  if ( 0 = Err.Number ) Then MakeWaveformEx = True
  On Error Goto 0
End Function


'===============================================================================
' GetMinMax
'-------------------------------------------------------------------------------
' Function for evaluating minimum and maximum of an ADC channel
' (setting these parameters speeds up loading time)
'-------------------------------------------------------------------------------
Function GetMinMax( oFile, dYOffset, dYGain, dChannelMin, dChannelMax )
  Dim lPosition
  ' remember file pointer position
  lPosition = oFile.Position
  ' move file pointer to position of the last file section which contains the 
  ' highest data reduction level (only two values = minimum and maximum)
  oFile.Position = oFile.Size - 4
  ' read values from file 
  dChannelMin     = oFile.GetNextBinaryValue(eU16)
  dChannelMax     = oFile.GetNextBinaryValue(eU16)
  ' scale values
  dChannelMin    = CDbl(dYOffset) + CDbl(dChannelMin) * CDbl(dYGain)
  dChannelMax    = CDbl(dYOffset) + CDbl(dChannelMax) * CDbl(dYGain)
  ' reset file pointer to original position
  oFile.Position = lPosition
  GetMinMax= True
End Function 'GetMinMax
'-------------------------------------------------------------------------------

