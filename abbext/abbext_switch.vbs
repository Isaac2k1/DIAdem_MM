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
Dim DEV04 : DEV04 = "change_menu(""user"")"
Dim DEV05 : DEV05 = "change_menu(""dvlp"")"
Dim DEV06 : DEV06 = "Call Scriptstart("""  & AutoActPath & "Tangentcursor\TangentCursor_Init.vbs"")"
Dim DEV07 : DEV07 = "Call Scriptstart("""  & AutoActPath & "Tangentcursor\stop_tangent_cursor.vbs"")"
Dim DEV08 : DEV08 = "Call Scriptstart("""  & AutoActPath & "testload_ludvika\load_test_ludvika.vbs"")"
'Dim DEV04 : DEV04 = "Call ReportExLoad3()"

'Restores the default settings
Call MenuReset()

 Dim tfh,menu_mode
      ' open text file, read formula from text file
      tfh = TextFileOpen("C:\Diadem\configs\user_mode.txt", tfRead)
      If TextFileError(tfh) = 0 Then
        menu_mode=Textfilereadln(tfh)
      end if
    textfileclose(tfh)

'msgbox menu_mode


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
  Call MenuItemInsert(Window, cstr(MenuItemCount)& ".3" ,"MENUITEM","Travel compare",ABB03) 
  'Call MenuItemInsert(Window, cstr(MenuItemCount)& ".4" ,"MENUITEM","Load &Reportfile",ABB04) 
  
  if menu_mode="dvlp" then 
  
      'Add Developer Menu
      'This call creates the main entry in the drop down menu area in DIAdem
      'Call MenuItemInsert(Window, cstr(MenuItemCount+1), "Popup", "&Developer")
      'This call adds a new menu item to the end of the main menu bar in all panels
      'Call MenuItemAdd(Window,"M","MenuItem","ABB")
      
      'auskommentiert
      'Call MenuItemInsert(Window, cstr(MenuItemCount)& ".4" ,"SEPARATOR","","")
      'Call MenuItemInsert(Window, cstr(MenuItemCount)& ".5" ,"MENUITEM","Following functions are not validated","")
      'Call MenuItemInsert(Window, cstr(MenuItemCount)& ".6" ,"MENUITEM","and only for development:","")
      'Call MenuItemInsert(Window, cstr(MenuItemCount)& ".7" ,"SEPARATOR","","")
      'Call MenuItemInsert(Window, cstr(MenuItemCount)& ".8" ,"MENUITEM","Evaluation (beta, not validated)",DEV01)
      'Call MenuItemInsert(Window, cstr(MenuItemCount)& ".9" ,"MENUITEM","Batch processing for evaluation (beta, not validated)",DEV02)
      'Call MenuItemInsert(Window, cstr(MenuItemCount)& ".10" ,"MENUITEM","Calibration tool (beta, not validated)",DEV03)
      'Call MenuItemInsert(Window, cstr(MenuItemCount)& ".11" ,"MENUITEM","View with Tangent Cursor",DEV06)
      'Call MenuItemInsert(Window, cstr(MenuItemCount)& ".12" ,"MENUITEM","Stop Tangent Cursor",DEV07)
      'Call MenuItemInsert(Window, cstr(MenuItemCount)& ".13" ,"SEPARATOR","","")
      'Call MenuItemInsert(Window, cstr(MenuItemCount)& ".14" ,"MENUITEM","Remove developer menu",DEV04)
      'Call MenuItemInsert(Window, cstr(MenuItemCount+1)& ".2" ,"MENUITEM","Create &Report1",DEV02)
      'Call MenuItemInsert(Window, cstr(MenuItemCount+1)& ".3" ,"MENUITEM","Create &Report2",DEV03)
      'Call MenuItemInsert(Window, cstr(MenuItemCount+1)& ".4" ,"MENUITEM","Create &Report3",DEV04) 
  
      Call MenuItemInsert(Window, cstr(MenuItemCount)& ".4" ,"SEPARATOR","","")
      Call MenuItemInsert(Window, cstr(MenuItemCount)& ".5" ,"Popup","Developer functions (not validated)","")
      'Call MenuItemInsert(Window, cstr(MenuItemCount)& ".6" ,"MENUITEM","and only for development:","")
      'Call MenuItemInsert(Window, cstr(MenuItemCount)& ".7" ,"SEPARATOR","","")
      Call MenuItemInsert(Window, cstr(MenuItemCount)& ".5.1" ,"MENUITEM","Evaluation",DEV01)
      Call MenuItemInsert(Window, cstr(MenuItemCount)& ".5.2" ,"MENUITEM","Batch processing for evaluation",DEV02)
      Call MenuItemInsert(Window, cstr(MenuItemCount)& ".5.3" ,"MENUITEM","Calibration tool",DEV03)
      Call MenuItemInsert(Window, cstr(MenuItemCount)& ".5.4" ,"MENUITEM","View with Tangent Cursor",DEV06)
      Call MenuItemInsert(Window, cstr(MenuItemCount)& ".5.5" ,"MENUITEM","Stop Tangent Cursor",DEV07)
      Call MenuItemInsert(Window, cstr(MenuItemCount)& ".5.6" ,"MENUITEM","Load Test from Ludvika",DEV08)
      Call MenuItemInsert(Window, cstr(MenuItemCount)& ".5.7" ,"SEPARATOR","","")
      Call MenuItemInsert(Window, cstr(MenuItemCount)& ".5.8" ,"MENUITEM","Remove developer menu",DEV04)
      
  
  
  elseif menu_mode="user" then 
      Call MenuItemInsert(Window, cstr(MenuItemCount)& ".4" ,"SEPARATOR","","")
      Call MenuItemInsert(Window, cstr(MenuItemCount)& ".5" ,"MENUITEM","Add developer menu",DEV05)
   
  end if 'if menu_mode="dvlp" then 

Next


