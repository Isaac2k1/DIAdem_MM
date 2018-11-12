'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2006/12/11 11:07:30
'-- Author: Kaan Oenen/Mathias Knaak
'   Last Update: 2008-01-01.00.00.00
'   Version: 1.0.0
'   Reviewed:
'   Comment:  calls evaluations dialog to modify and run evaluations.
'             for batch processing use "batch_processing_v*.*.vbs"    
'-------------------------------------------------------------------------------

Option Explicit



evaludef

Sub evaludef()

  GlobalDim "LastSelectionText,LastSelectionInd,LastSelectionType"
  GlobalDim "questions"
  GlobalDim "K,add_eva,selected_ch_ind"
  
  'Set the User Script Folder to \abb
  AUTODRVUSER = AutoActPath
  Call UserVarCompile("gvar_eva_abb.vas")

'show View-Window

'call scriptcmdadd(autoactpath&"importrbs_for_evaludef.vbs")

'call importrbs_for_evaludef("2517/358")

'call SUDDlgShow("Dlg_docu_string",AutoActPath & "edit_v1.8.sud","")



'on error resume next
  'call SUDDlgShow("Dlg1",AutoActPath & "edit_v1.8.sud","auto_eva")
  call SUDDlgShow("Dlg1",AutoActPath & "edit.sud","")
End sub  




draw_mark_lines



  
  'msgbox T1
  'If SUDDlgShow("Dlg2","Edit") = "IDOk" Then
    'msgbox basic_eva_
    'msgbox doc_and_updates_
    'msgbox spec_eva_
  'End if