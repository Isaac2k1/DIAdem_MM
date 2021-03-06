'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2009-06-14 15:30:39
'-- Author: Marco.Mailand@ch.abb.com
'-- Comment: My first VBS code
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
Dim oMyGrp, oMyChn, oDigOUT, i , oAna00, oAna90, oDig00 , oDig90, imax, Digital_position, position, positionstep
Set oMyGrp = Data.Root.ChannelGroups.Add("MyChnGroup")
Set oMyChn = oMyGrp.Channels.Add("Digital_Angle", DataTypeFloat64)
Set oDig00 = oMyGrp.Channels.Add("oDig00", DataTypeFloat64)
Set oDig90 = oMyGrp.Channels.Add("oDig90", DataTypeFloat64)
Set oAna00 = Data.Root.ChannelGroups(1).Channels("Travel dig 1")
Set oAna90 = Data.Root.ChannelGroups(1).Channels("Travel dig 2")
'since the two digital signals have been measured with analog signals due to lack of 
'connector hardware we do digitalization in software
'the threshold is set to 2.5 and not to TTL-levels to make it really roboust
'For i = 1 to Data.Root.ChannelGroups(1).Channels("Digital 0 Grad").Properties("length").Value
For i = 1 to oAna00.Properties("length").Value
  if oAna00(i) > 2.5 then 
    oDig00(i) = 1
  else
    oDig00(i) = 0
  end if
  if oAna90(i) > 2.5 then 
    oDig90(i) = 1
  else
    oDig90(i) = 0
  end if
Next
'this is the initial position of the very first measurement point
'the real value could be the last point of the absolut position from previous measurement
position = 0
'the pitch of the used aperture plate was one degree, Laser beam diameter was a half degree
'resulting in a quarter degree resulotion
'positionstep should be an input parameter since the pitch may change later (use e.g. ScaleF)
'positionstep = 0.25  ' one quarter degree for a sensor disc with 360 slits
positionstep = 0.141732283464567  ' sensor disc with 635 slits 360/635/4=0.141732283464567
'loop thru all points to find the rising and falling slopes for increment and decrement 
'of the position counter. This loop is probably very slow.
'For i = 2 to Data.Root.ChannelGroups(1).Channels("Digital 0 Grad").Properties("length").Value-1
For i = 2 to oAna00.Properties("length").Value-1
'either digital output can initiate incrementing or decrementing
'the sign is dependant on the state of the other digital input
'these 8 lines can certainly be improved, please send me your result :-)
'mailto:marco.mailand@ch.abb.com
  if oDig00(i-1) < oDig00(i) AND oDig90(i)=0 then position = position - positionstep
  if oDig00(i-1) > oDig00(i) AND oDig90(i)=1 then position = position - positionstep
  if oDig00(i-1) < oDig00(i) AND oDig90(i)=1 then position = position + positionstep
  if oDig00(i-1) > oDig00(i) AND oDig90(i)=0 then position = position + positionstep
  if oDig90(i-1) < oDig90(i) AND oDig00(i)=0 then position = position + positionstep
  if oDig90(i-1) > oDig90(i) AND oDig00(i)=1 then position = position + positionstep
  if oDig90(i-1) < oDig90(i) AND oDig00(i)=1 then position = position - positionstep
  if oDig90(i-1) > oDig90(i) AND oDig00(i)=0 then position = position - positionstep
  Data.Root.ChannelGroups(2).Channels("Digital_Angle")(i) = position
  'this is a very nasty workaround to set the very first result channel point to zero
  'the line on its own ("Data.Root.ChannelGroups(2).Channels("Digital_Angle")(i-1) = 0")
  'does NOT work outside the For loop and there is no obvious reason why
  if i = 2 then Data.Root.ChannelGroups(2).Channels("Digital_Angle")(i-1) = 0
Next
'get an x-axis from the first channel
Call WfChnToChn("[1]/micronor")     '... ChnNoStr 
'reconstruct the just converted x-y-axis pair back into a waveform, because DIAdem deletes 
'the source channel and cannot be switched to no delete mode
Call ChnToWfChn("[1]/micronor_A_generated","[1]/micronor",0)       '... X,ChnNoStr,XChnDelete 
'now generate three waveforms by means of the just created x-axis
Call ChnToWfChn("[1]/micronor_A_generated","[2]/Digital_Angle",0)  '... X,ChnNoStr,XChnDelete 
Call ChnToWfChn("[1]/micronor_A_generated","[2]/oDig00",0)         '... X,ChnNoStr,XChnDelete 
Call ChnToWfChn("[1]/micronor_A_generated","[2]/oDig90",0)         '... X,ChnNoStr,XChnDelete 
'the x-axis is no longer needed -> trash it
Call Data.Root.ChannelGroups(1).Channels.Remove("micronor_A_generated")
'remove the offset of the analog measurements
Call ChnOffset("[1]/micronor","/micronor",100,"mean value offset") '... Y,E,ChnOffsetValue,ChnOffsetMode 
Call ChnOffset("[1]/poti","/poti",100,"mean value offset")         '... Y,E,ChnOffsetValue,ChnOffsetMode 
'scaling of the voltage channels for degree output
Call MatScalMul("[1]/micronor",79.5/3.96,0)               '... ChnNoStr,MatScalVal,MatScalIP 
Call MatScalMul("[1]/poti",79.5/3.2,0)                    '... ChnNoStr,MatScalVal,MatScalIP 
'now nename the scaled channels, move the digital signals to the first channel group and remove
'now empty second channel group
Data.Root.ChannelGroups(1).Channels("MatrixScalarMultiplicatedZ1").Name = "micronor_scaled"
Data.Root.ChannelGroups(1).Channels("MatrixScalarMultiplicatedZ2").Name = "poti_scaled"
Call Data.Move(Data.Root.ChannelGroups(2).Channels("Digital_Angle"),Data.Root.ChannelGroups(1).Channels,7)
Call Data.Move(Data.Root.ChannelGroups(2).Channels("oDig00"),Data.Root.ChannelGroups(1).Channels,8)
Call Data.Move(Data.Root.ChannelGroups(2).Channels("oDig90"),Data.Root.ChannelGroups(1).Channels,9)
Call Data.Root.ChannelGroups.Remove(2)
