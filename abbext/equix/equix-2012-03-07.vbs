'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2006/11/02 13:27:02
'-- Author: Kaan Oenen/Mathias Knaak/...
'-- Version: 1.9.3
'-- Comment: Creating equidistant x axis for channels with 1 or 2 sampling frequency 
'            (!!!! using DataPlugin rbd and Linear Mapping tool)
'-------------------------------------------------------------------------------
'
'For details of some functions used in this script it may be necessary to look at
'user command file abb_user_commands.vbs in \abb folder 
'Following user functions are used: CopyProp(), delete_repeats_then_sort()

'Version History
'v1.0
'Takes the X and Y channels(scaled x and y values) as input and copies them(incl.properties) to new channel group
'Uses CHNLINGENIMP function to generate an implicit channel(new x range) which has high sampling freq.
'New channel has a sampling rate of TbB
'Maps x and y channels to new x range using CHNMAPLINCALC function (Linear Mapping tool in Analysis panel)
'Result is an interpolated Y Channel mapped to new x range
'Converts X and Y Channels into a waveform channel
'Properties of the original channel is transferred to result waveform channel
'Temporarily created channels are deleted after converting to equidistant
'
'Function delete_repeats_then_sort() included, which deletes repeating values in X channel and sorts
'x and y channels so that x channel becomes monoton increasing
'v1.1
'equix() can take more then 1 x-y channel pairs as input
'v1.2
'Second input of the function is x-y channel numbers (instead of indexes) from this version on.
'v1.3
'First input of the function is an array of Channelgroup indexes of the channels selected. 
'v1.4
'It is possible to select "auto x-step" or "custom x-step" for implicit x channel generation
'Two Input added to the function. iGenType       : implicit x channel generation type (auto,custom x-step)
'                                 CustomTimeStep : Difference between two sequential x values in case of Custom x-stepping
'v1.5
'The way to copy files to new channel group has changed. The revision could be seen at line no 119,120
'v1.6
'Getting X and Y Channelno's using Groupindex/Channelindex realised. Line 125-126
'Waveform channel naming is changed. Waveform channels could be identified by "#" at the end of their names. e.g : Trave
'Deleting repeated values in x channel in order to make the channel monoton increasing, is not done anymore. Such files
'will be accepted to be corrupt files and the waveform conversion wont' be done.
'v1.7
'Mode operations are not used anymore, since these operations dont deliver reliable results if real numbers used.
'Channel numbers are acquired directly from the array which was passed to function equix().
'Autodetection of small sampling rate (fast sampling): MinV(TbACopy,TbBCopy) See Line 187
'v1.8
'After the waveform conversion script runs, the values of the program variables are set to NULL (no valid data)
'v1.9
'New GroupName-Extension changed from "_data_# " to ".rbe"
'Channel names changed to remove "_Y" at or near the end
'The new ChannelName checked by DIAdem and modified if required
'v1.9.1
'For each Input-File a new Channel-Group is created adding 1,2,... at the End-of-Name if required
'
'v1.9.2 (Mathias Knaak)
'properties from rbd-group copied to new rbe-group

'v1.9.3 (Mathias Knaak)
' x-Axis Unit added



'Dim xy(0,1)
'xy(0,0) = 3
'xy(0,1) = 4
'xy(1,0) = 3
'xy(1,1) = 4

'xy(2,0) = 1
'xy(2,1) = 2
'xy(3,0) = 3
'xy(3,1) = 4
'xy(4,0) = 5
'xy(4,1) = 6
'xy(2,0) = 1
'xy(2,1) = 2

'Dim aGroup(0)
'aGroup(0) = 2
'aGroup(1) = 1
'aGroup(2) = 3
'aGroup(3) = 3
'aGroup(4) = 3

'Call equix(aGroup,xy,0,0.00032)

'-------------------------------------------------------------------------------
'Function 
'equix()
'v1.9.1
'Description
'   This function creats equidistant x axis for channels with 1 or 2 sampling frequency 
'   (!!!! using DataPlugin rbd and Linear Mapping tool)
'
'   Input
'   iGroupindex() : Indexes of the channelgroups (in correct order), where the channels you want to convert are placed
'   iXYChNo()     : 2 dimension array,  1st dimension (unlimited size)       : number of x-y channel pairs
'                                  2nd dimension (size = 1 (2elements)) : x,y channel no's in DataPortal
'   iGenType      : Implicit X Channel generation method. 0 = Auto or 1 = x-step
'   CustomTimeStep: Difference between two sequential x values in case of Custom x-stepping
'   
'   Output
'   Equidistant waveforms of the input channels in a new channel group
'-------------------------------------------------------------------------------

Function equix(iGroupindex(),iXYChNo(),iGenType,CustomTimeStep)

  'if the second dimension of the array is not 1 Then the x-y channel pairs ChannelNo-Array is not a valid array
  If Ubound(iXYChNo,2) <> 1 Then
    Call Msgbox("You have a wrong dimensioned array as input. The script will be terminated"&vbCrLf&vbCrLf& _
                  " Make sure equix( ) has the following inputs"&vbCrLf&vbCrLf& _
                  " iGroupindex()  --> Index of the channelgroups, where the channels you want to convert are placed"&vbCrLf&vbCrLf& _ 
                  " iXYChindex() --> 2 dimension array"&vbCrLf& _
                  "                             1st dimension (unlimited size)              : number of x-y channel pairs"&vbCrLf& _
                  "                             2nd dimension (size = 1 (2elements)) : x,y channel no's in Dataportal"&vbCrLf& _
                  " iGenType      : Implicit X Channel generation method. Auto(0) or Custom(1) x-step")
    Call Autoquit()
  End if
  
  Dim Tb1Copy,Tb2Copy,sChnName,iChnNoPortal,iChn,sXChnIndexP,sYChnIndexP,iXChnNoPortal,iYChnNoPortal
  Dim SrcGrpName,TgtGrpName,iXChNoCopy,iYChNoCopy,m,n,blnFound,intCount,iXChnNo,iYChnNo,iGrp,ChNum
  Dim sXChnNameCopy,sYChnNameCopy,sXChnIndexCopy,sYChnIndexCopy,TgtGrpIndex,sChnNameYinterp
  Dim TbACopy,TbBCopy,LastXCopy,tStart,yScale,yOffset,iTb12,iTb21,iFirstSampleNo,ChnNoXimp
  Dim t_initial,t_final,t_total,ts1,ts2,t_totalA,t_totalB,t_totalC,SmallSR,TbAeCopy,TbBeCopy
  Dim t_unit,i
  Dim MeineVar_Debug

  iChn = 0
  ChNum = Ubound(iXYChNo,1)
  
  While(iChn<=ChNum)
   'get the group name of the Source-Channel
   SrcGrpName = GroupPropGet(iGroupindex(iChn),"Name")
   'Call MsgBox(SrcGrpName & ": Chn" & iChn & " '" & ChnPropGet(iXYChNo(iChn,0),"name") & "'")
   
   If(Not ChnPropExist(iXYChNo(iChn,0),"TbA_s")) Then
    Call MsgBox("Can't convert Channel: " & SrcGrpName & "/" & ChnPropGet(iXYChNo(iChn,0),"name"))
    iChn = iChn+1 'Try next Channel
   Else
    'Create generic Name of the Target-Group
    n = Len(SrcGrpName)
    m = InStr(SrcGrpName,".rbd")
    If(m>0) Then 'Replace .rbd by .rbe
     TgtGrpName = Left(SrcGrpName,m-1)&".rbe"
    Else         'Append .eqx
     TgtGrpName = SrcGrpName&".eqx"
    End If
    
    n = 0
    m = TgtGrpName
    blnFound = True
    While(blnFound)
     blnFound = FALSE
     'Check all Groups for already existing TgtGrpName
     For intCount = 1 to GroupCount
      If(GroupName(intCount) = TgtGrpName) Then
       blnFound = TRUE
       Exit For
      End If
     Next
     
     If(blnFound) Then
      n = n+1
      TgtGrpName = m&n
     End If
    Wend
    
    Call GROUPCREATE(TgtGrpName,0)
    call grouppropcopy(groupindexget(SrcGrpName),groupindexget(TgtGrpName)) 'copy properties from source group

    'get the group index of new channel group created 
    TgtGrpIndex = GroupIndexGet(TgtGrpName)
    
    'Specify the last generated group as default
    Call GROUPDEFAULTSET(TgtGrpIndex)
    iGrp = iGroupindex(iChn)
    
    Do While(iGroupindex(iChn)=iGrp)
     'Get Channel-Numbers of Input-Channels in Data-Portal 
     iXChnNo           = iXYChNo(iChn,0)                                      
     iYChnNo           = iXYChNo(iChn,1)
     sXChnNameCopy     = ChnPropGet(iXChnNo,"name")  
     sYChnNameCopy     = ChnPropGet(iYChnNo,"name")
     
     'Copy X- and Y-Channels to the last Positions in the new Channel-Group
     Call ChnCopyExt(iXChnNo, TgtGrpIndex, GroupChnCount(TgtGrpIndex)+1)
     Call ChnCopyExt(iYChnNo, TgtGrpIndex, GroupChnCount(TgtGrpIndex)+1)
     
     'Create Chngroup-Index/ChnName-String for copied Channels
     sXChnIndexCopy = "["&TgtGrpIndex&"]/"&sXChnNameCopy
     sYChnIndexCopy = "["&TgtGrpIndex&"]/"&sYChnNameCopy
     
     iXChNoCopy     = Val(ChnPropGet(sXChnIndexCopy ,"number")) 'get channel number of new X channel in DataPortal
     iYChNoCopy     = Val(ChnPropGet(sYChnIndexCopy ,"number")) 'get channel number of new Y channel in DataPortal
     
     t_unit         = ChnPropGet(iXChNoCopy,"ChnUnit")
    
     TbACopy        = cDbl(ChnPropGet(iXChNoCopy,"TbA_s"))      'get timebase1 (time dif.between two samples for fs1)
     TbBCopy        = cDbl(ChnPropGet(iXChNoCopy,"TbB_s"))      'get timebase2 (time dif.between two samples for fs2)
     LastXCopy      = cDbl(ChnPropGet(iXChNoCopy,"LastX"))      'get total number of desired points (LastX)
     tStart         = cDbl(ChnPropGet(iXChNoCopy,"tStart"))     'get tStart
     yScale         = cDbl(ChnPropGet(iYChNoCopy,"yScale"))
     yOffset        = cDbl(ChnPropGet(iYChNoCopy,"yOffset"))
     iTb12          = cDbl(ChnPropGet(iYChNoCopy,"iTb12"))
     iTb21          = cDbl(ChnPropGet(iYChNoCopy,"iTb21"))
   
     If ChnVal(1, iXChNoCopy) = tStart Then
      iFirstSampleNo = 0
     Elseif ChnVal(1, iXChNoCopy) = tStart + TbACopy Then 
      iFirstSampleNo = 1
     End if

    

    
     'Define parameters for channels with 1 or 2 sampling rate
     
     '...and false
     if(TbACopy <> TbBCopy) and (iTb21 = iTb12)  then
 '    msgbox(ixchnno & " case 1")
        TbAeCopy = TbACopy * iTb12 + tStart
        TbBeCopy = TbBCopy * (iTb21-iTb12) + TbAeCopy
        
        'Calculate first X value
       t_initial = tStart + iFirstSampleNo * TbACopy 
             'Calculate X value for switching point fs1-->fs2
       ts1       = tStart + iTb12 * TbACopy
       'Calculate X value for switching point fs2-->fs1
       ts2       = TbAeCopy + (iTb21-iTb12) * TbBCopy 
       'Calculate last X value
       t_final   = TbBeCopy + (LastXCopy-iTb21) * TbACopy
       'Calculate total acquisition time
       t_total   = t_final - t_initial
             'Calculate time till first switching point    
       t_totalA  = ts1 - t_initial
       'Calculate time between first and second switching points
       t_totalB = ts2 - ts1
       'Calculate time from switching point 2 till the end of channel
       t_totalC = t_final - ts2
      
             'Find out which sampling rate is the small one
       SmallSR = MinV(TbACopy,TbBCopy)
      
             'Assign auxillary variables for an easier mode operation (by assuring integer operation)
       L1 = Fix(SmallSR*(10^(9))) 
       L2 = Fix(t_total*(10^(9)))
       
       'Calculate (t_total)mod(TbBCopy) 
       Call FormulaCalc("L3 := (L2)Mod(L1)")
      
             'Depending on Auto or Custom X steps, input parameters for x channel generation will change
       Select Case iGenType      
         Case 0   'Auto x-step                     
           'Assign start value of the implicit channel which will be generated
           GhdStartVal = t_initial
           'Calculate the planned length of the implicit channel which will be generated
           GhdChnLength = 1 + Trunc(L2/L1) + (-1)*(L3>0)    '(L3>0): True = -1 False = 0
           'Specifiy the difference between two sequential values in implicit channel
           GHdStep = SmallSR
         Case 1   'Custom x-step
           GhdStartVal   = t_initial
           GhdChnLength  = 1 + Trunc(t_total/CustomTimeStep) + (-1)*(Frac(t_total/CustomTimeStep)<>0)     
           GHdStep       = CustomTimeStep
       End Select

    'If there are two sample rates in the channel
    elseif(TbACopy <> TbBCopy) and (iTb21 > iTb12) and (LastXCopy >= iTb21)  Then      
'    msgbox(ixchnno & " case 2")
       TbAeCopy = TbACopy * iTb12 + tStart
       TbBeCopy = TbBCopy * (iTb21-iTb12) + TbAeCopy
       
       'Calculate first X value
       t_initial = tStart + iFirstSampleNo * TbACopy 
       'Calculate X value for switching point fs1-->fs2
       ts1       = tStart + iTb12 * TbACopy
       'Calculate X value for switching point fs2-->fs1
       ts2       = TbAeCopy + (iTb21-iTb12) * TbBCopy 
       'Calculate last X value
       t_final   = TbBeCopy + (LastXCopy-iTb21) * TbACopy
       'Calculate total acquisition time
       t_total   = t_final - t_initial

       
       'Find out which sampling rate is the small one
       SmallSR = MinV(TbACopy,TbBCopy)
       
       'Assign auxillary variables for an easier mode operation (by assuring integer operation)
       L1 = Fix(SmallSR*(10^(9))) 
       L2 = Fix(t_total*(10^(9)))
       
       'Calculate (t_total)mod(TbBCopy) 
       Call FormulaCalc("L3 := (L2)Mod(L1)")
       
       'Depending on Auto or Custom X steps, input parameters for x channel generation will change
       Select Case iGenType      
         Case 0   'Auto x-step                     
           'Assign start value of the implicit channel which will be generated
           GhdStartVal = t_initial
           'Calculate the planned length of the implicit channel which will be generated
           GhdChnLength = 1 + Trunc(L2/L1) + (-1)*(L3>0)    '(L3>0): True = -1 False = 0
           'Specifiy the difference between two sequential values in implicit channel
           GHdStep = SmallSR
         Case 1   'Custom x-step
           GhdStartVal   = t_initial
           GhdChnLength  = 1 + Trunc(t_total/CustomTimeStep) + (-1)*(Frac(t_total/CustomTimeStep)<>0)     
           GHdStep       = CustomTimeStep
       End Select
       
     'If there is only one sampling rate in the channel
     ElseIf(TbACopy = TbBCopy) and (iTb21 = iTb12) Then  
'    msgbox(ixchnno & " case 3")
       'Calculate first X value
       t_initial = tStart + iFirstSampleNo * TbACopy
       'Calculate last X value
       t_final   = tStart + LastXCopy * TbACopy
       'Calculate total acquisition time
       t_total   = t_final - t_initial
       
       'Depending on Auto or Custom X steps, input parameters for x channel generation will change
      
       Select Case iGenType
         Case 0
           If ChnVal(1, iXChNoCopy) = tStart Then                        
             GhdStartVal    = tStart
             GhdChnLength   = LastXCopy+1
             GHdStep        = TbACopy     
        ' Changed by MK, comparison of (ChnVal(1, iXChNoCopy) was never equal to (tStart + TbACopy)
           Elseif ((ChnVal(1, iXChNoCopy) - (tStart + TbACopy)) < (Tbacopy/10)) Then         'if the first sampleno is 0 , then the length of implicit channel should be incremented by 1
             GhdStartVal    = tStart + TbACopy
             GhdChnLength   = LastXCopy
             GHdStep        = TbACopy
           End if
         Case 1
           GhdStartVal     = t_initial
           GhdChnLength    = 1 + Trunc(t_total/CustomTimeStep) + (-1)*(Frac(t_total/CustomTimeStep)<>0)
           GHdStep         = CustomTimeStep
       End Select
     Else
       Call AutoQuit("Check if TbA,TbB,switching-poinsts,LastX values are consistent!"& vbCrLf&_
                   "Script will be terminated and no waveform channels will be created!")
     End if
     'Modify channel name to remove extra '_Y' at or near the end
     m = Len(sYChnNameCopy): n = InStrRev(sYChnNameCopy,"_Y")
     If m = n+1 Then ' Remove '_Y' at the very end
       sYChnNameCopy = Left(sYChnNameCopy,n-1)
     ElseIf m > n+1 Then ' Remove '_Y' at the end but preserve succeeding characters (usually 1, 2, ...)
       sYChnNameCopy = Left(sYChnNameCopy,n-1)&Right(sYChnNameCopy,m-n-1)
     End if
     sChnNameYinterp = sYChnNameCopy
     
     'Check if there is a channel named the same as "sChnNameYinterp"
     Call ChnNameChk(TgtGrpIndex,sChnNameYinterp)
     If GHdChnName <> sChnNameYinterp Then
      sChnNameYinterp = GHdChnName
     End If
     
     GHdChnName  = sXChnNameCopy&"_generated"          'get X channel name and modify it
     
     'Delete repeating values in X channel and sort the channel so that it becomes monoton increasing
     'Call delete_repeats_then_sort(iXChNoCopy,iYChNoCopy)  
     'Generate an implicit channel(new x range) with linear definition    'input: Chn name,  chn length,  start value,step
     ChnNoXimp = Val(ChnPropGet(CHNLINGENIMP(GHdChnName,GhdChnLength,GhdStartVal,GHdStep),"number")) '... GHDCHNNAME,GHDCHNLENGTH,GHDSTARTVAL,GHDSTEP 
     'Call CHNLINGENIMP(GHdChnName,GhdChnLength,GhdStartVal,GHdStep) '... GHDCHNNAME,GHDCHNLENGTH,GHDSTARTVAL,GHDSTEP 
     'Map x and y channels to new x range. input: x chn, y chn, chn with new x range, result chn,...

    
     Call CHNMAPLINCALC(iXChNoCopy,iYChNoCopy,ChnNoXimp,"/"&sChnNameYinterp,1,"f[bound.slope]",NOVALUE,"analogue") '... XW,Y,X1,E,MAPLINNOVINTERP,MAPLINEXTTYPE,MAPLINBDRYVAL,MAPLINFCTTYPE   
     'Convert X and Y Channels into a waveform channel
     Call CHNTOWFCHN(ChnNoXimp,"/"&sChnNameYinterp,0) '... X,CHNNOSTR,XCHNDELETE 
     'Copy source channel properties to result waveform channel
     Call CopyProp(iYChnNo,"/"&sChnNameYinterp)

     'Get Unit from "ChnUnit" to "Unit"
      chndim ("["&TgtGrpIndex&"]/" &"[" &iChn+4 &"]") = ChnPropGet("["&TgtGrpIndex&"]/" &"[" &iChn+4 &"]","ChnUnit")
      call chnpropset("["&TgtGrpIndex&"]/" &"[" &iChn+4 &"]","wf_xunit_string",t_unit)


'offset adjustment
  'undo shifting of the messurement 
  call chnoffset("["&TgtGrpIndex&"]/" &"[" &iChn+4 &"]","["&TgtGrpIndex&"]/" &"[" &iChn+4 &"]",-0.001*yOffset*yScale,"free offset")
  call chnpropset("["&TgtGrpIndex&"]/" &"[" &iChn+4 &"]","YOffset",0) 'delete offset information in the adjusted Wave
  
  'try to reduce offset by bad calibration
  dim offset10ms
  dim offset20ms
  dim peakToPeak
  dim stepPerMS    
  dim noise
  
  stepPerMS = 1/(1000 * TbACopy)    
  peakToPeak = abs(cDbl(ChnPropGet("["&TgtGrpIndex&"]/" &"[" &iChn+4 &"]","Maximum"))) + abs(cDbl(ChnPropGet("["&TgtGrpIndex&"]/" &"[" &iChn+4 &"]","Minimum")))
  
  'try to find a possible offset at the beginnig
  'calculate meanvalues and noise peak-to-peak value at the beginning
  offset10ms = mean("["&TgtGrpIndex&"]/" &"[" &iChn+4 &"]",1000*tstart*stepPerMS,1000*tstart*stepPerMS+10*stepPerMS)
  offset20ms = mean("["&TgtGrpIndex&"]/" &"[" &iChn+4 &"]",1000*tstart*stepPerMS,1000*tstart*stepPerMS+20*stepPerMS)
  noise = abs(findPeakToPeak("["&TgtGrpIndex&"]/" &"[" &iChn+4 &"]",0,10))
  
  if abs(offset20ms-offset10ms) < noise and abs(offset20ms) < 0.01 * peakToPeak then
    call chnoffset("["&TgtGrpIndex&"]/" &"[" &iChn+4 &"]","["&TgtGrpIndex&"]/" &"[" &iChn+4 &"]",-1*(offset20ms+offset10ms)/2,"free offset")
  else  'try it at the end    
    'calculate meanvalues and noise peak-to-peak at the end
    offset10ms = mean("["&TgtGrpIndex&"]/" &"[" &iChn+4 &"]",lastXcopy-10*stepPerMS,lastXcopy)
    offset20ms = mean("["&TgtGrpIndex&"]/" &"[" &iChn+4 &"]",lastXcopy-20*stepPerMS,lastXcopy)
    noise = abs(findPeakToPeak("["&TgtGrpIndex&"]/" &"[" &iChn+4 &"]",lastXCopy - 10,lastXCopy))
    
    if abs(offset20ms-offset10ms) < noise and abs(offset20ms) < 0.01 * peakToPeak then
      call chnoffset("["&TgtGrpIndex&"]/" &"[" &iChn+4 &"]","["&TgtGrpIndex&"]/" &"[" &iChn+4 &"]",-1*(offset20ms+offset10ms)/2,"free offset")
    end if 
  end if      
  
     'Get Unit from "ChnUnit" to "Unit"
      chndim ("["&TgtGrpIndex&"]/" &"[" &iChn+4 &"]") = ChnPropGet("["&TgtGrpIndex&"]/" &"[" &iChn+4 &"]","ChnUnit")
      call chnpropset("["&TgtGrpIndex&"]/" &"[" &iChn+4 &"]","wf_xunit_string",t_unit)
      
     'Delete temporarily created channels
     Call CHNDELETE("'["&TgtGrpIndex&"]/["&(GroupChnCount(TgtGrpIndex)-3)&"]' - '["&TgtGrpIndex&"]/["&(GroupChnCount(TgtGrpIndex)-1)&"]'") '... CLPSOURCE

     iChn = iChn+1 'Next Channel
     If(iChn>ChNum) Then Exit Do
      Loop
     End if
  Wend
  'Set program values to NULL (no valid data) so that the values are reset
  'Wrong assignment may occur if they are not reset, because program variables are valid until DIAdem is closed.
  GhdStartVal  = NULL
  GhdChnLength = 0  
  GHdStep      = NULL
  GHdChnName   = NULL
End Function


function findPeakToPeak(channel,startXValue,stopXValue)
  dim min
  dim max
  
  min =  1e9
  max = -1e9
  
  for i = startXValue + 1 to stopXValue 
    if CHD(i,channel) < min then min = CHD(i,channel)
    if CHD(i,channel) > max then max = CHD(i,channel) 
  next
  
  findPeakToPeak = max - min
end function

Function mean(channel,startXValue,stopXValue) 
    dim area
    dim i
    int stepWidth 
    
    stepWidth = (stopXValue-startXValue)/10000            'to limit the calulations
    if stepWidth < 1 then stepWidth = 1                 

    area = 0
    
    for i = startXValue + 1 to stopXValue step stepWidth
      area = area + CHD(i,channel)                        
    next
    
    mean = stepWidth * area / (stopXValue - startXValue)
end function