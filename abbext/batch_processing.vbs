'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2008-08-29 13:14:00
'-- Author: Mathias Knaak
'-- Comment: Batch processing for evaluation routines
'   Last Update: 2008-01-01.00.00.00
'   Version: 1.0.0
'   Reviewed:
'-------------------------------------------------------------------------------
Option Explicit  'Erzwingt die explizite Deklaration aller Variablen in einem Skript.

Call scriptcmdadd(autoactpath&"importrbs_for_evaludef.vbs")

 GlobalDim "LastSelectionText,LastSelectionInd,LastSelectionType"
  GlobalDim "questions"
  GlobalDim "K,add_eva,selected_ch_ind"
 
 ' If Not ItemInfoGet("labcode") then

      Globaldim "labcode,loadmode,load_digital_chn,load_error_files"
      labcode = 1         'default is power lab
      loadmode = 2        'default is equidist-x
      load_digital_chn=0  'default is not to load digital channels
      load_error_files=0  'default is to load standard files
  
  'end if  
 
 
  
  'Set the User Script Folder to \abb
  AUTODRVUSER = AutoActPath
  Call UserVarCompile("gvar_eva_abb.vas")



View.loadlayout(autoactpath&"view_for_evaluation.TDV")
'  View.Sheets(1).Cursor.X1 = t_start
'  View.Sheets(1).Cursor.X2 = t_end
  Call WndShow("VIEW","MAXIMIZE")



call SUDDlgShow("Dlg1",AutoActPath & "batch_process_dlg.sud")








