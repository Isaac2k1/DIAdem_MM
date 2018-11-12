'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2006/12/11 11:07:30
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------

GlobalDim "LastSelectionText,LastSelectionInd,LastSelectionType"

Call UserVarCompile(AutoActPath & "gvar_eva_abb.vas")
Call SUDDlgShow("Dlg1",AutoActPath & "edit_v1.1.sud")
'msgbox T1
'If SUDDlgShow("Dlg2","Edit") = "IDOk" Then
  'msgbox basic_eva_
  'msgbox doc_and_updates_
  'msgbox spec_eva_
'End if