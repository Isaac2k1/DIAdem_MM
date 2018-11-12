'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2006/11/20 08:52:48
'-- Author: Kaan Oenen /Updates: Mathias Knaak
'-- Comment: ABB user commands: This file includes frequently used functions
'            created for global usage in all scripts 
'   Last Update: 2008-07-28.09:19:10
'   Version: 1.0.0
'   Reviewed:	
'7.7.2008: Mathias Knaak : rbd_unit_rename added
'11.7.2008: Mathias Knaak: remove_digital_channels added
'16.7.2008: Mathias Knaak: group_desc_write added
'-------------------------------------------------------------------------------

Option Explicit




'-------------------------------------------------------------------------------
'Sub delete_repeats_then_sort()
'   
'Description 
'   This procedure finds out the repeating values in the X channel, deletes them 
'   (also corresponding values in y channel) and sorts the X channel ascending
'   so that it becomes monoton increasing (also sorts corresponding values in y channel)
'
'   Input
'   CALCXChn : the channel number of x channel 
'   CALCYChn : the channel number of y channel
'
'------------------------------------------------------------------------------- 
Public Sub delete_repeats_then_sort(CALCXChn, CALCYChn)
  
  dim XChName, XChNameDelta, ChnNoDelta, Repeats, I, x, y0, y1

  'CALCXChn    = 1                                          'define the channel number of x channel
  'CALCYChn    = 2                                          'define the channel number of y channel
  XChName = ChnPropGet(CALCXChn,"name")                     'get x channel name
  XChNameDelta= XChName&"_Delta"                            'create a name for delta channel

  Call ChnMultipleSort(CALCXChn, CALCYChn, 0, 1)            'sort channels according to ascending x values
  Call ChnDeltaCalc(CALCXChn, XChNameDelta)                 'calculate the difference between adjacent values in x channel
  ChnNoDelta = Val(ChnPropGet(XChNameDelta,"number"))

  'iChNoDelta = CNo(XChNameCopy&"_Delta")
  'Tell user which points repeat
  Repeats = ""
  For I = 1 to Chnlength(ChnNoDelta) 
    If ChnVal(I,ChnNoDelta) = 0 Then
      'Check if Y values are same for repeating x
      x  = ChnVal(I,CALCXChn)
      y0 = ChnVal(I,CALCYChn)
      y1 = ChnVal(I+1,CALCYChn)
      If Y0 <> Y1 Then
        Call AutoQuit("Corrupted rbd file: In Channel "&XChName&" there are repeating x-y paris with variable y values (x/y :"&x&"/"&y0&","&x&"/"&y1&"). "&_ 
                      XChName&" can not be converted into waveform channel. The script will be terminated. Exclude that channel and try again")
      End if
      Repeats = Repeats&", "&ChnVal(I,CALCXChn)
    End if
  Next
  If Repeats <> "" Then
    Call msgbox("There are x values repeating in channel: "&XChName&" which causes non-monotonous x channel.These points are x = "&Repeats&_
                " These disturbing points will be removed from the channel(without modifying original rbd file)")
  End if

  L1 = CALCXChn
  L2 = CNo(XChNameDelta)

  ChnLength(L2) = ChnLength(L2) + 1                         'increase delta channels length by 1
  ChDX(ChnLength(L2), L2) = Null                            'assign Null to last value of delta channel
  Call FormulaCalc("Ch(L1):= Ch(L1) + NoValue*(Ch(L2)=0)")  'find out repeating values in x channel and replace them with NOVALUEs
  Call ChnDel(L2)                                           'delete the temporary delta channel
  Call ChnNoVHandle(CALCXChn, CALCYChn, "Delete", "X", 1, 0)'get rid of NOVALUES in x channel and corresponding values in y channel
  'Call View.LoadLayout(AutoActPath & "Remove Repeats.TDV")
  'Call WndShow("VIEW")

End Sub
'-------------------------------------------------------------------------------
'------------------------------------------------------------------------------- 
'Sub CopyProp()
'
'Description
'   This prodecure copies all the properties that are defined by user from the source channel to target channel     
'    
'   Input : SourceChnNo : The channel whose properties you want to copy
'           TargetChnNo : The target channel
'------------------------------------------------------------------------------- 
Sub CopyProp(SourceChnNo,TargetChnNo)
 Dim intLoop,sPropName,sPropValue
 
 'TargetChnIndex = CNo("(ISHUNT)_X")
 For intLoop = 1 to ChnPropCount(SourceChnNo)
  Call ChnPropInfoGet(SourceChnNo,ChnPropNameGet(SourceChnNo,intLoop))
  If (Not PropIsFixed ) Then 'Property is user defined
    sPropName   = ChnPropNameGet(SourceChnNo,intLoop)
    'sPropValue  = Val(ChnPropGet(SourceChnNo,sPropName))
    'DataType of ScaleF unkown in both original _Y Channel and the corresp.
    'Equidistan-channel. Also in the original Channel it is diplayed with 6
    'decimal Digits while in the Equidistant Channel it is rounded to 5 Digits.
    'Next two Lines could not fix the Problem.
    sPropValue  = ChnPropGet(SourceChnNo,sPropName)
    Call ChnPropInfoGet(SourceChnNo,sPropName)
    Call ChnPropCreate(TargetChnNo,sPropName,DataType)
    'msgbox(DatatypeAsText(DataType)&" "&sPropName&"="&sPropValue) 
    Call ChnPropSet(TargetChnNo,sPropName,sPropValue) 
  Else
    intLoop = intLoop + 1
  End If
 Next

End Sub
'------------------------------------------------------------------------------- 

'-------------------------------------------------------------------------------
'Sub SortIntArrayAsc()
'
'Description
'   This procedure takes an integer array with randomly ordered numbers and outputs
'   the same array with the elements in ascending order
'      
'   Input
'   intArr() : integer array with random order
'
'   Output
'   The same array sorted in ascending order
'-------------------------------------------------------------------------------
Sub SortIntArrayAsc(intArr())

  Dim Save,N,Q
  
  'Check adjacent elements in the array
  Do until N = Ubound(intArr)
    'If a leading element is bigger than the following element
    If intArr(N) > intArr(N+1) Then
      'Buffer the leading element (bigger one)
      Save        = intArr(N)
      'Assign the leading element with smaller value
      intArr(N)   = intArr(N+1)
      'Assign the following element with bigger value
      intArr(N+1) = Save
      'Go back to the beginning of the array to check through again
      N           = 0
    Else
      N = N+1
    End if
  Loop
  
  'Display the elements of re-ordered Array
  'For Q = 0 to Ubound(intArr)
  '  msgbox(intArr(Q))
  'Next

End Sub
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'Function
'ChnStringtoChnNumbers()
'
'Description
'   This function takes channel numbers string as input and outputs a channel number array
'   
'   Input
'   ChannelString(string)             : Channel string e.g: ("1,3-5,7")
'
'   Output
'   ChannelNoArray(array of integers) : Array of Channel numbers e.g: (1,3,4,5,7)
'-------------------------------------------------------------------------------
Function ChnStringtoChnNumbers(ChannelString)

  Dim asBuffer(),sBuffer,iKommaPos,I,J,K,L,iMinusPos,iLenAsBuffer,ShotNo(),iRangeStart,iRangeEnd

  sBuffer = ChannelString
  I = 0
  'If there is no comma, then there is only one range of channels e.g: "1-5"
  If InstrRev(sBuffer, ",") = 0 Then
    Redim Preserve asBuffer(0)
    asBuffer(0) = sBuffer
  Else
  'If the elements are seperated with comma, there could be single channels / range of channels  
    While (Not InStrRev(sBuffer,",") = 0)
      'msgbox sbuffer
      Redim Preserve asBuffer(I)
      iKommaPos = InStrRev(sBuffer,",")
      asBuffer(I) = Mid(sBuffer,iKommaPos+1,Len(sBuffer)-iKommaPos) 
      'Save each elements/ranges seperated by ","
      sBuffer = Left(sBuffer,iKommaPos-1)
      I = I +1
    Wend
    Redim Preserve asBuffer(I)
    asBuffer(I) = sBuffer
  End If

  K = -1
  'Analize each element if there is a range in it e.g : "2-7" (channels numbers 2 to 7)
  For J=0 To I
    iMinusPos = InStr(asBuffer(J),"-")
    iLenAsBuffer = Len(asBuffer(J))
    
    'If there is no range indicator "-", then it is a single channel number element
    If iMinusPos = 0 Then
      K = K+1
      Redim Preserve ShotNo(K)
      ShotNo(K) = Cint(asBuffer(J))
    Else
    'If there is a range indicator "-", find out where the range starts and ends
      iRangeStart = Cint(Left(asBuffer(J),iMinusPos-1))
      iRangeEnd = Cint(Right(asBuffer(J),iLenAsBuffer-iMinusPos))
      
      'Create the intermediate channel numbers in the range
      For L = iRangeStart To iRangeEnd
        K = K+1
        Redim Preserve ShotNo(K)
        'Save channel numbers in the array
        ShotNo(K) = L
      Next
    End If
  Next

  'Sort the integer array in ascending order
  'Call SortIntArrayAsc(ShotNo)
  'Return value = sorted array
  ChnStringtoChnNumbers = ShotNo
End Function
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'Function 
'Check_numeric_string_syntax()
'v1.0
'Description
'   This function checks if a string if it is a valid numeric expression
'   It allows the string to have following characters "0-9" and "."   
'
'   Input
'   InputString : The string to be checked if it is numeric
'
'   Output
'   True  : String is a valid numeric expression
'   False : Strin is not a valid numeric expression 
'-------------------------------------------------------------------------------

Function Check_numeric_string_syntax(InputString)

'Check the syntax of SNo/TNo string
  Dim A : A=1
  
  Check_numeric_string_syntax = True
  
  'Check all the characters in the string
  Do While A <= Len(InputString)
    sChar = Mid(InputString,A,1)
    'Check if the character is one of the allowed characters
    If (sChar <> "1" ) And (sChar <> "2" ) And (sChar <> "3" ) And (sChar <> "4" ) And (sChar <> "5" ) And _
       (sChar <> "6" ) And (sChar <> "7" ) And (sChar <> "8" ) And (sChar <> "9" ) And (sChar <> "0" ) And _
       (sChar <> "." ) Then
    
    'Return false when a non-allowed char. found
    Check_numeric_string_syntax = False
    Exit Do
    
    End If
    A = A +1
  Loop

End Function
'-------------------------------------------------------------------------------




'-------------------------------------------------------------------------------

'''''''''''''
' sub rbd_unit_rename()
'
'
'Copies Unit description for each channel in .rbd Groups from "ChnUnit" to "unit" 
'
' V1.0 Mathias Knaak 7.7.08
'
'Input: needs no input, checks all groups with "rbd" in the group name


sub rbd_unit_rename()
dim grp_number,chn_number

'check each group 
for grp_number=1 to groupcount 
 'if Groupname contains "rbd" ...
  if instr(1,grouppropvalget(grp_number,"Name"),"rbd",1)<>0 then  
    '...go to each channel        
    for chn_number=1 to groupchncount(grp_number)                 
    'Copy Unit from "ChnUnit" to "Unit"
      chndim ("["&grp_number&"]/" &"[" &chn_number&"]") = ChnPropGet("["&grp_number&"]/" &"[" &chn_number&"]","ChnUnit")
    next 
  
  end if

next  'grp_number=1 to groupcount 

end sub

'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'
' sub remove_digital_channels
'
'removes digital channels from rbd-group
'Properties "TotA" in group description and "ChnNumber" in each channel needed
'
' "TotA" contains the number of analog channels in the transient recorder. 
' If the channel number "ChnNumber" is greater than "TotA", the channel is a digital channel.
'
' M.Knaak  11.07.2008
'
'Input: Group number, the group has to be an *.rbd group


sub remove_digital_channels (group_number)  'needs group number 

dim chn_number,chn_count, analog_count
chn_number=1
'Number of all channels in the group
chn_count=groupchncount(group_number) 

'determine number of analog channels
if grouppropexist(group_number,"TotA") then
   analog_count=grouppropvalget(group_number,"TotA")
else
  'if "TotA" is missing, warn user and exit sub 
  call msgbox("Number of analog channels cannot be read from group description.")
  exit sub
end if

'if the group is an *.rbd group
if instr(1,grouppropvalget(group_number,"Name"),"rbd",1)<>0 then
  'check each channel in the group
  while chn_number <= groupchncount(group_number)
    'if the channel is a digital channel, delete this channel
    if chnpropvalget("["&group_number&"]/["&chn_number&"]","ChnNumber") > analog_count then
        chndelete("["&group_number&"]/["&chn_number&"]")
    else
    'if the channel is an analog channel, go to next channel
    chn_number=chn_number+1  
    end if
  wend
else 
  ' warn user if the group is not an rbd-group
  call msgbox("Group is not an *.rbd-group."&CHR(13)&"No channels deleted.")
  exit sub
end if

end sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' group_desc_write
'
' M.Knaak 16.7.2008
'
' copies serial number and shot number to group description
' deletes labcode and data type, replaces "-" with "/"
'
' Input: groupindex

sub group_desc_write(group_index)

dim group_desc, desc_pos,shot_desc                    'Gruppenname, Gruppenindex

               'remove first character and .rbd from group description
          
                group_desc = grouppropget(group_index,"name")
                desc_pos=len(group_desc)
                
                'remove first character
                group_desc = right(group_desc,desc_pos-1) 
                
                'remove .rb*
                desc_pos=instr(group_desc,".rb")
                if desc_pos > 0 then group_desc = left(group_desc,desc_pos-1) 'remove ".rb*"
                
                'exception handling for shot 20080020:
                if left(group_desc,4)=8020 then
                  group_desc=replace(group_desc,"8020","20080020")
                end if
                
                'change Description for PEHLA-Plots
                desc_pos=instr(group_desc,"-")
                if desc_pos=9 then 'PEHLA-Shot found
                  desc_pos=len(group_desc)
                  shot_desc=right(group_desc,4)
                  'remove leading "0" in shot-number
                  while left(shot_desc,1)="0" 
                    shot_desc=right(shot_desc,len(shot_desc)-1)
                  wend
                  group_desc = mid(group_desc,3,7)&shot_desc
                  desc_pos=len(group_desc)
                  group_desc=left(group_desc,2)&right(group_desc,desc_pos-3) 
                'replace "-" with "Ba/"
                group_desc=replace(group_desc,"-","Ba/") 
                else
                'replace "-" with "/"
                group_desc=replace(group_desc,"-","/") 
                end if  'PEHLA-Shot
                'write new description to "description"
                call grouppropset(group_index,"description",group_desc) 'Gruppenname in "desription" sichern


end sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' change_menu
'
' M.Knaak 14.10.2008
'writes user mode to user config file

sub change_menu(user_type)

 Dim tfh,menu_mode
      ' open text file, read formula from text file
      tfh = TextFileOpen("C:\Diadem\configs\user_mode.txt",TFCreate or tfwrite or tfansi)
      If TextFileError(tfh) = 0 Then
        call Textfileseek(tfh,1)
        call Textfilewriteln(tfh,user_type)
      else msgbox "error"
      end if
    textfileclose(tfh)

'call scriptstart(autoactpath &"abbext_switch.vbs")
call scriptstart("C:\DIAdem\abbext\abbext.vbs")

if user_type="dvlp" then
call msgbox ("User mode changed to DEVELOPER!",64,"User Mode changed")
call msgbox("Developer functions are not validated."&CHR(13)&"Functions may have errors or give wrongs results!",48,"Warning") 
elseif user_type="user" then call msgbox ("User mode changed to USER!",64,"User Mode changed")
end if
end sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DateEnglish
'
' M. Averes 12.02.2009
' Wird z.B. benutzt in ABB2009.tdrm
'   Date: @@DateEnglish(CurrDateTimeReal)@@
function DateEnglish(DateVal)
  Dim iMonth, sMonth
  iMonth = RTP(DateVal,"t")
  select case iMonth
    case  1 : sMonth = "January"
    case  2 : sMonth = "February"
    case  3 : sMonth = "March"
    case  4 : sMonth = "April"
    case  5 : sMonth = "May"
    case  6 : sMonth = "June"
    case  7 : sMonth = "July"
    case  8 : sMonth = "August"
    case  9 : sMonth = "September"
    case 10 : sMonth = "October"
    case 11 : sMonth = "November"
    case 12 : sMonth = "December"
  end select
  DateEnglish = sMonth & " " & str(CurrDatetimereal, "#dd") & ", " & str(CurrDatetimereal, "#YYYY")
end function

