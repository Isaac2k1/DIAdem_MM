'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2006/11/28 13:14:05
'-- Author: Kaan Oenen
'-- Comment: Main script for the sub-menu "covert to waveform" in ABB menu 
'-------------------------------------------------------------------------------
'
Option Explicit

Globaldim "xstep,ChnNoString,CustomTimeStep"
Dim ChnNoArray,I

'Already loaded by abbext.vbs
'Call ScriptCmdAdd(AutoActPath & "..\equidistant_x\equix.vbs")

If SUDDlgShow("Dlg1",AutoActPath & "num2wf.sud") = "IDOk" Then
  
  ChnNoArray = ChnStringtoChnNumbers(ChnNoString)
  Dim NoOfSel : NoOfSel = UBound(ChnNoArray)
  Dim iXYChNo(),iGroupindex()
  Redim Preserve iXYChNo(NoOfSel,1)
  REdim Preserve iGroupindex(NoOfSel)
  For I = 0 to NoOfSel
    'Assign Chnnumber to x value
    iXYChNo(I,0)    = ChnNoArray(I)
    'Assign Chnnumber+1 to y value
    iXYChNo(I,1)    = ChnNoArray(I)+1
    iGroupindex(I)  = ChnGroup(ChnNoArray(I))
  Next
  Call equix(iGroupindex,iXYChNo,xstep,CustomTimeStep)
    
'Else
  ' Display message.
  'Call MsgBoxDisp("You have clicked <Cancel>. The script is finished.")
End if

'Never remove: Reset of the Script-Engine and Menu-Crash would be forced
'Call ScriptCmdRemove(AutoActPath & "..\equidistant_x\equix.vbs")
