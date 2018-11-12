'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2006/12/11 11:07:30
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------

Option Explicit



evaludef

Sub evaludef()

  GlobalDim "LastSelectionText,LastSelectionInd,LastSelectionType"
  GlobalDim "questions"
  GlobalDim "K,add_eva,selected_ch_ind"
  
  'Set the User Script Folder to \abb
  AUTODRVUSER = AutoActPath
  Call UserVarCompile("gvar_eva_abb_v1.6.vas")

on error resume next
  call SUDDlgShow("Dlg1",AutoActPath & "edit_v1.8.sud")
End sub  
  
  'msgbox T1
  'If SUDDlgShow("Dlg2","Edit") = "IDOk" Then
    'msgbox basic_eva_
    'msgbox doc_and_updates_
    'msgbox spec_eva_
  'End if