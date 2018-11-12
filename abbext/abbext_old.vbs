'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 06/29/2006 09:33:28
'-- Author: Markus Averes
'-- Comment: Modified Kaans version
'   Last Update: 2008-07-28.09:23:00
'   Version: 1.0.0
'   Reviewed:	
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

'Create Strings for [MenuItemFct] : menu item functions
Dim ABB01 : ABB01 = "Call Scriptstart("""  & AutoActPath & "importrbs\testload.vbs"")"
Dim ABB02 : ABB02 = "Call Scriptstart("""  & AutoActPath & "equix\num2wf.vbs"")"
'Dim ABB03 : ABB03 = "Call Scriptstart("""  & AutoActPath & "evaludef_for_edit.vbs"")"
'Dim ABB04 : ABB04 = "Call ReportExLoad()"
'Dim ABB04 : ABB04 = "Call Scriptstart("""  & AutoActPath & "report\Load_Report.vbs"")"
Dim ABB05 : ABB05 = "Call Scriptstart("""  & AutoActPath & "Travel_comp\Travel_compare.vbs"")"

'Restores the default settings
Call MenuReset()

For each Window in DDWindows
  'This call will populate variables such as MenuItemCount
  Call MenuItemCountGet(Window, "M")
  'This call creates the main entry in the drop down menu area in DIAdem
  Call MenuItemInsert(Window, cstr(MenuItemCount), "Popup", "&ABB")
  'This call adds a new menu item to the end of the main menu bar in all panels
  'Call MenuItemAdd(Window,"M","MenuItem","ABB")
  Call MenuItemInsert(Window, cstr(MenuItemCount)& ".1" ,"MENUITEM","&Load Testdata",ABB01)
  Call MenuItemInsert(Window, cstr(MenuItemCount)& ".2" ,"MENUITEM","Convert to &Waveform (equidistant x)",ABB02)
  'Call MenuItemInsert(Window, cstr(MenuItemCount)& ".3" ,"MENUITEM","&evaludef",ABB03)
 ' Call MenuItemInsert(Window, cstr(MenuItemCount)& ".4" ,"MENUITEM","Load &Reportfile",ABB04) 
  Call MenuItemInsert(Window, cstr(MenuItemCount)& ".5" ,"MENUITEM","Travel compare",ABB05) 
Next
