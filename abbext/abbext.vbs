'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 06/29/2006 09:33:28
'-- Modified on 28.03.2013 10:10:00 by Markus
'-- Report function repaired (directory C:\DIAdem\abbext\Report 
'-- Comment: massive changes in 2018-04-19 to execute this file at each startup.
'-------------------------------------------------------------------------------
Option Explicit
'-------------------------------------------------------------------------------
'Register ABB User-Commands Files at start-up: Those Files contain
'Functions that are added to  the  standard  DIAdem-Functionality
'and must be localized in the same Directory as this Start-Script.
'-------------------------------------------------------------------------------
Dim DDWindows(4), Window, tfh, tfhError, menu_mode, strFile
strFile = "C:\Diadem\configs\user_mode.txt"
Call BarManager.Reset ' this is the new command to reset DIAdem menus, for versions since 2015
' open text file for user_mode, read value in or create default file for user
tfh = TextFileOpen(strFile,eTextFileAttributeRead)
menu_mode=Textfilereadln(tfh)
'msgbox("menu mode:"& menu_mode)
tfhError = TextFileClose(tfh)
  If tfh = -1 Then 'if file does not exist, create it
    tfh = TextFileOpen(strFile,eTextFileAttributeCreate OR eTextFileAttributeWrite OR eTextFileAttributeANSI)
    tfhError = TextFileWriteLn(tfh,"user") ' initially set it to user mode
    tfhError = TextFileClose(tfh)
    Call MsgboxDisp("user mode was unknown, created file for user mode","MB_OK",,,5) 
    End If
tfh = TextFileOpen(strFile,eTextFileAttributeRead)
menu_mode=Textfilereadln(tfh)
'msgbox("menu mode:"& menu_mode)
tfhError = TextFileClose(tfh)

Call ScriptCmdAdd(AutoActPath & "abb_user_commands.vbs")
Call ScriptCmdAdd(AutoActPath & "equix\equix.vbs")
Call ScriptCmdAdd(AutoActPath & "importrbs\importrbs.vbs")
'-------------------------------------------------------------------------------
'Create ABB menu in all DIAdem panels
'-------------------------------------------------------------------------------
  
    DDWindows(0) = "NAVIGATOR"
    DDWindows(1) = "VIEW"
    DDWindows(2) = "ANALYSIS"
    DDWindows(3) = "REPORT"
    DDWindows(4) = "SCRIPT"
    
    'Create Strings for [MenuItemFct] : menu item functions for ABB Menu
    Dim ABB01 : ABB01 = "Call Scriptstart("""  & AutoActPath & "LoadTest\Load_Test.VBS"")"
    Dim ABB02 : ABB02 = "Call Scriptstart("""  & AutoActPath & "importrbs\testload.vbs"")"
    Dim ABB03 : ABB03 = "Call Scriptstart("""  & AutoActPath & "equix\num2wf.vbs"")"
    Dim ABB04 : ABB04 = "Call Scriptstart("""  & AutoActPath & "Travel_comp\Travel_compare.vbs"")"
    Dim ABB05 : ABB05 = "Call Scriptstart("""  & AutoActPath & "ExportTest\ExportTest.vbs"")"
    Dim ABB06 : ABB06 = "Call Scriptstart("""  & AutoActPath & "Plot_Report\Plot_Report.vbs"")"
    
    'Create Strings for [MenuItemFct] : menu item functions for Developer Menu
    Dim DEV01 : DEV01 = "Call Scriptstart("""  & AutoActPath & "evaludef_for_edit.vbs"")"
    Dim DEV02 : DEV02 = "Call Scriptstart("""  & AutoActPath & "batch_processing.vbs"")"
    Dim DEV03 : DEV03 = "Call Scriptstart("""  & AutoActPath & "Kanalvergleich\Kanalvergleich.vbs"")"
    Dim DEV04 : DEV04 = "Call Scriptstart("""  & AutoActPath & "Tangentcursor\TangentCursor_Init.vbs"")"
    Dim DEV05 : DEV05 = "Call Scriptstart("""  & AutoActPath & "Tangentcursor\stop_tangent_cursor.vbs"")"
    Dim DEV06 : DEV06 = "Call Scriptstart("""  & AutoActPath & "testload_ludvika\load_test_ludvika.vbs"")"
    Dim DEV07 : DEV07 = "Call Scriptstart("""  & AutoActPath & "chn_prop_copy\prop_chn.vbs"")"
    Dim DEV08 : DEV08 = "Call Scriptstart("""  & AutoActPath & "laser_time_shift\time_shift_of_laser_measurement.vbs"", ""manual"")"
    Dim DEV09 : DEV09 = "Call Scriptstart("""  & AutoActPath & "testload_kema\load_test_kema.vbs"")"
    Dim DEV10 : DEV10 = "Call Scriptstart("""  & AutoActPath & "TestRecord\TestRecord.vbs"")"
    Dim DEV11 : DEV11 = "Call Scriptstart("""  & AutoActPath & "calc_travel\calc_travel.vbs"")"
    Dim DEV12 : DEV12 = "Call Scriptstart("""  & AutoActPath & "Integrate_channel_with_arcing-time\Integrate_with_arcing-time.VBS"")"
    Dim DEV13 : DEV13 = "Call Scriptstart("""  & AutoActPath & "#PPHV-TI\AgilentDSOX2014A\OsziDatenLaden.VBS"")"
    Dim DEV14 : DEV14 = "Call Scriptstart("""  & AutoActPath & "#PPHV-TI\UFD-AMDET\AMDET_2.2.VBS"", ""InitializingSingle"")"
    Dim DEV15 : Dev15 = "Call Scriptstart("""  & AutoActPath & "#PPHV-TI\UFD-AMDET\AMDET_2.2.VBS"", ""InitializingAll"")"
    Dim DEV16 : Dev16 = "Call Scriptstart("""  & AutoActPath & "Temp_Update\Temp_Update.VBS"")"
    Dim DEV17 : Dev17 = "Call Scriptstart("""  & AutoActPath & "03degree\03degree.VBS"")"
    Dim DEV18 : Dev18 = "Call Scriptstart("""  & AutoActPath & "LoadTest\Load_Test_MOCKUP.VBS"")"
    Dim DEV19 : DEV19 = "change_menu(""user"")"
    Dim DEV20 : DEV20 = "change_menu(""dvlp"")"
       

For each Window in DDWindows
      'This call will populate variables such as MenuItemCount
      Call MenuItemCountGet(Window, "M")
      'Add ABB Menu
      'This call creates the main entry in the drop down menu area in DIAdem
      Call MenuItemInsert(Window, cstr(MenuItemCount), "Popup", "&ABB")
      'This call adds a new menu item to the end of the main menu bar in all panels
      'Call MenuItemAdd(Window,"M","MenuItem","ABB")
      Call MenuItemInsert(Window, cstr(MenuItemCount)& ".1" ,"MENUITEM","&Load Test Data",ABB01)
      Call MenuItemInsert(Window, cstr(MenuItemCount)& ".2" ,"MENUITEM", "Load Test Data [old]",ABB02)
      Call MenuItemInsert(Window, cstr(MenuItemCount)& ".3" ,"MENUITEM","Convert to &Waveform (equidistant x)",ABB03)
      Call MenuItemInsert(Window, cstr(MenuItemCount)& ".4" ,"MENUITEM","Travel compare",ABB04) 
      Call MenuItemInsert(Window, cstr(MenuItemCount)& ".5" ,"MENUITEM","Export Test",ABB05)
      Call MenuItemInsert(Window, cstr(MenuItemCount)& ".6" ,"MENUITEM","Create Plot",ABB06)
      
      if menu_mode="dvlp" then 
          'Add Developer Menu
          'This call creates the main entry in the drop down menu area in DIAdem
          'Call MenuItemInsert(Window, cstr(MenuItemCount+1), "Popup", "&Developer")
          'This call adds a new menu item to the end of the main menu bar in all panels
          'Call MenuItemAdd(Window,"M","MenuItem","ABB")
      
          Call MenuItemInsert(Window, cstr(MenuItemCount)& ".7"      , "SEPARATOR","","")
          Call MenuItemInsert(Window, cstr(MenuItemCount)& ".8"      , "Popup","Developer functions (not validated)","")
          Call MenuItemInsert(Window, cstr(MenuItemCount)& ".8.1"    , "MENUITEM","Evaluation",DEV01)
          Call MenuItemInsert(Window, cstr(MenuItemCount)& ".8.2"    , "MENUITEM","Batch processing for evaluation",DEV02)
          Call MenuItemInsert(Window, cstr(MenuItemCount)& ".8.3"    , "MENUITEM","Calibration tool",DEV03)
          Call MenuItemInsert(Window, cstr(MenuItemCount)& ".8.4"    , "MENUITEM","Start Tangent Cursor",DEV04)
          Call MenuItemInsert(Window, cstr(MenuItemCount)& ".8.5"    , "MENUITEM","Stop Tangent Cursor",DEV05)
                                                                       
          Call MenuItemInsert(Window, cstr(MenuItemCount)& ".8.6"    , "Popup", "Load Test from another lab","") 
          Call MenuItemInsert(Window, cstr(MenuItemCount)& ".8.6.1"  , "MENUITEM","Load Test from Ludvika",DEV06)
          Call MenuItemInsert(Window, cstr(MenuItemCount)& ".8.6.2"  , "MENUITEM","Load Test from Kema",DEV09)
                                                                       
          Call MenuItemInsert(Window, cstr(MenuItemCount)& ".8.7"    , "MENUITEM","Copy Properties to channels",DEV07)
          Call MenuItemInsert(Window, cstr(MenuItemCount)& ".8.8"    , "MENUITEM","TestRecord",DEV10)
          Call MenuItemInsert(Window, cstr(MenuItemCount)& ".8.9"    , "MENUITEM","degree -> m",DEV11) 
                                                                       
          Call MenuItemInsert(Window, cstr(MenuItemCount)& ".8.10"   , "Popup", "Interrupter Development","") 
          Call MenuItemInsert(Window, cstr(MenuItemCount)& ".8.10.1" , "MENUITEM","Laser Time Shift",DEV08)
          Call MenuItemInsert(Window, cstr(MenuItemCount)& ".8.10.2" , "MENUITEM","Integrate i(t) with arcing-time",DEV12)
          Call MenuItemInsert(Window, cstr(MenuItemCount)& ".8.10.3" , "MENUITEM","Load data from AgilentDSOX2014A",DEV13)
          Call MenuItemInsert(Window, cstr(MenuItemCount)& ".8.10.4" , "MENUITEM","UFD-AMDET (Single Shot)",DEV14)
          Call MenuItemInsert(Window, cstr(MenuItemCount)& ".8.10.5" , "MENUITEM","UFD-AMDET (Data Portal Evaluation incl. TestLoad)",DEV15)
          Call MenuItemInsert(Window, cstr(MenuItemCount)& ".8.11"   , "MENUITEM", "Temp Data Update",DEV16) 
          Call MenuItemInsert(Window, cstr(MenuItemCount)& ".8.12"   , "MENUITEM", "Dreigrad - Nullgrad",DEV17) 
          Call MenuItemInsert(Window, cstr(MenuItemCount)& ".8.14"   , "MENUITEM", "Test_Load_MOCKUP",DEV18) 
		  
        'Report function only for test engineers (directory C:\DIAdem\abbext\Report is needed & activation of lines 54,55,113,114 & 115, ask Markus Averes or Marco Mailand)
        'Call MenuItemInsert(Window, cstr(MenuItemCount)& ".6.12" ,"Popup", "Plot","")                                  'MA
        'Call MenuItemInsert(Window, cstr(MenuItemCount)& ".6.12.1" ,"MENUITEM", "Generalplot",DEV15)                   'MA
        'Call MenuItemInsert(Window, cstr(MenuItemCount)& ".6.12.2" ,"MENUITEM", "Detailedplot",DEV16)                  'MA
        ' To implement new MenuItem insert new Call above this comment and
        ' increase the MenuItemCount of the subsequent SEPARATOR and Remove... accordingly.        
          
          Call MenuItemInsert(Window, cstr(MenuItemCount)& ".8.14" ,"SEPARATOR","","")
          Call MenuItemInsert(Window, cstr(MenuItemCount)& ".8.15" ,"MENUITEM","Remove developer menu",DEV19)
      
      elseif menu_mode="user" then 
          Call MenuItemInsert(Window, cstr(MenuItemCount)& ".7" ,"SEPARATOR","","")
          Call MenuItemInsert(Window, cstr(MenuItemCount)& ".8" ,"MENUITEM","Add developer menu",DEV20)
      
      end if ' menu_mode="dvlp" 
  Next
'End If