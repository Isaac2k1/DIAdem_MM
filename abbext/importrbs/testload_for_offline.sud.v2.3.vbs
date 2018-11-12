'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2006/12/21 10:50:16
'-- Author: Kaan Oenen
'-- Comment: Main Script for "Load an old Test" Menu
'-------------------------------------------------------------------------------

testload

Sub testload()
  Globaldim "K"
  Globaldim "lab_code"
  Globaldim "sDataName"
  Globaldim "sReportName"
   
  If SudDlgShow("Dlg1",AutoActPath & "offline_v2.3.sud") = "IDOk" Then
    'For I = 0 to K  'Ubound(sDataName)
    '  NameString = NameString&" "&sDataName(I)&" "&sReportName(I)
    'Next
    'Call msgbox("Selected lab_code: "&lab_code&" and "&NameString) 
  End if
    
End sub
'-------------------------------------------------------------------------------