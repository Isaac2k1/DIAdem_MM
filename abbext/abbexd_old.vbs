'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 06/29/2006 09:33:28
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
Option Explicit

'-------------------------------------------------------------------------------
'Register ABB User-Commands Files at start-up: Those Files contain
'Functions that are added to  the  standdard  DIAdem-Functionality
'and must be localized in the same Directory as this Start-Script.
'-------------------------------------------------------------------------------
Call ScriptCmdAdd(AutoActPath & "abb_user_commands.vbs")
'Call ScriptCmdAdd(AutoActPath & "abb_eva_sub_collection.vbs")
Call ScriptCmdAdd(AutoActPath & "equix\equix.vbs")
Call ScriptCmdAdd(AutoActPath & "importrbs\importrbs.vbs")
'Call ScriptCmdAdd(AutoActPath & "report\ABB_Report_Script.vbs")

'-------------------------------------------------------------------------------
'Create ABB menu in all DIAdem panels
'-------------------------------------------------------------------------------
Dim DDWindows(4), Window

DDWindows(0) = "NAVIGATOR"
DDWindows(1) = "VIEW"
DDWindows(2) = "ANALYSIS"
DDWindows(3) = "REPORT"
DDWindows(4) = "SCRIPT"

'Create Strings for [MenuItemFct] : menu item functions for ABB Menu
Dim ABB01 : ABB01 = "Call Scriptstart("""  & AutoActPath & "importrbs\testload.vbs"")"
Dim ABB02 : ABB02 = "Call Scriptstart("""  & AutoActPath & "equix\num2wf.vbs"")"
Dim ABB03 : ABB03 = "Call Scriptstart("""  & AutoActPath & "Travel_comp\Travel_compare.vbs"")"
'Dim ABB04 : ABB04 = "Call Scriptstart("""  & AutoActPath & "report\Load_Report.vbs"")"

'Create Strings for [MenuItemFct] : menu item functions for Developer Menu
Dim DEV01 : DEV01 = "Call Scriptstart("""  & AutoActPath & "evaludef_for_edit-v1.9.vbs"")"
'Dim DEV02 : DEV02 = "Call ReportExLoad1()"
Dim DEV02 : DEV02 = "Call Scriptstart("""  & AutoActPath & "batch_processing_v1.0.vbs"")"
Dim DEV03 : DEV03 = "Call Scriptstart("""  & AutoActPath & "Kanalvergleich\Kanalvergleich.vbs"")"
'Dim DEV03 : DEV03 = "Call ReportExLoad2()"
'Dim DEV04 : DEV04 = "Call ReportExLoad3()"

'Restores the default settings
Call MenuReset()

For each Window in DDWindows
  
  'This call will populate variables such as MenuItemCount
  Call MenuItemCountGet(Window, "M")
  
  'Add ABB Menu
  'This call creates the main entry in the drop down menu area in DIAdem
  Call MenuItemInsert(Window, cstr(MenuItemCount), "Popup", "&ABB")
  'This call adds a new menu item to the end of the main menu bar in all panels
  'Call MenuItemAdd(Window,"M","MenuItem","ABB")
  Call MenuItemInsert(Window, cstr(MenuItemCount)& ".1" ,"MENUITEM","&Load Test",ABB01)
  Call MenuItemInsert(Window, cstr(MenuItemCount)& ".2" ,"MENUITEM","Convert to &Waveform (equidistant x)",ABB02)
  Call MenuItemInsert(Window, cstr(MenuItemCount)& ".5" ,"MENUITEM","Travel compare",ABB03) 
  'Call MenuItemInsert(Window, cstr(MenuItemCount)& ".4" ,"MENUITEM","Load &Reportfile",ABB04) 
  
  'Add Developer Menu
  'This call creates the main entry in the drop down menu area in DIAdem
  Call MenuItemInsert(Window, cstr(MenuItemCount+1), "Popup", "&Developer")
  'This call adds a new menu item to the end of the main menu bar in all panels
  'Call MenuItemAdd(Window,"M","MenuItem","ABB")
  Call MenuItemInsert(Window, cstr(MenuItemCount+1)& ".1" ,"MENUITEM","&evaludef",DEV01)
  Call MenuItemInsert(Window, cstr(MenuItemCount+1)& ".2" ,"MENUITEM","Evaluation: batch processing",DEV02)
  Call MenuItemInsert(Window, cstr(MenuItemCount+1)& ".3" ,"MENUITEM","Kanalvergleich",DEV03)
  'Call MenuItemInsert(Window, cstr(MenuItemCount+1)& ".2" ,"MENUITEM","Create &Report1",DEV02)
  'Call MenuItemInsert(Window, cstr(MenuItemCount+1)& ".3" ,"MENUITEM","Create &Report2",DEV03)
    'Call MenuItemInsert(Window, cstr(MenuItemCount+1)& ".4" ,"MENUITEM","Create &Report3",DEV04) 

Next
