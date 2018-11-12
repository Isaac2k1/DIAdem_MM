'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2006/12/21 10:50:16
'-- Author: Kaan Oenen (Update Mathias Knaak)
'-- Comment: Main Script for "Load an old Test" Menu
'-------------------------------------------------------------------------------

'v2.7 select if digital channels should be loaded
'v2.8 select if error data should be loaded

'-------------------------------------------------------------------------------
'v3.0 Loading single channels
'
'30.03.2010: Fabio Maurizio
'new variable viewtemplate added in order to load the different kinds of templates

'07.07.2010: Fabio Maurizio
'two new variables added for the new function: "load single channels"
'
'-------------------------------------------------------------------------------

Call testload()

'-------------------------------------------------------------------------------
'Sub procedure
'testload()
'compatible with dialogbox "Offline_v2.8"
'
'Description
'
'-------------------------------------------------------------------------------
'Variables
'K: Number of files
  Call Globaldim("load_channel")


Sub testload()
  
  if not iteminfoget ("view_template") then
  
    Globaldim "labcode,loadmode,load_digital_chn,load_error_files,dp_clear_mode,view_template,load_channel,channel_name,channel_name2,channel_option"
    labcode = 1         'default is power lab
    loadmode = 2        'default is load equidistant X
    load_digital_chn=0  'default is not to load digital channels
    load_error_files=0  'default is to load standard files
    view_template = 0   'default is to not load the view template
    load_channel = 0    'default is to not load single channels
    channel_name = ""   'default is an empty textbox
    channel_name2 = ""  
    channel_option = 0  'default only one textbox
   
  end if
  
  'Already loaded by abbext.vbs
  'Call ScriptCmdAdd(AutoActPath & "..\equidistant_x\equix.vbs") 'Must be in the same Path as this File
  'Call ScriptCmdAdd(AutoActPath & "importrbs.vbs") 'Must be in the same Path as this File
  
  If SudDlgShow("Dlg1",AutoActPath & "offline.sud") = "IDOk" Then
    'For I = 0 to K  'Ubound(sDataName)
    '  NameString = NameString&" "&sDataName(I)&" "&sReportName(I)
    'Next
    
    
    'Loading the different view templates for every kind of test
    '-------------------------------------------------------------------------------------
    
    If view_template = 1 Then
      
      'ONLY FOR HIGH POWER
      '2(ABB or PEHLA)x2(mechanical or electrical)x4(cycle code)+2(3s-operation)=18 templates 
      If labcode = 1 Then
      
        If len(Data.Root.ChannelGroups(Groupcount).Name) < 18  Then   'keine Pehlanummer sondern ABB Entwicklung
          If FilEx("C:\DIAdem\user_data\layouts\Leistungslabor\View_ABB_"&Data.Root.ChannelGroups(Groupcount).Properties("TESTING_CODE").Value&"_"&Data.Root.ChannelGroups(Groupcount).Properties("CYCLE_CODE").Value&"_"&Data.Root.ChannelGroups(Groupcount).Properties("CIRCUIT_CODE").Value&".TDV") Then
           Call View.LoadLayout("C:\DIAdem\user_data\layouts\Leistungslabor\View_ABB_"&Data.Root.ChannelGroups(Groupcount).Properties("TESTING_CODE").Value&"_"&Data.Root.ChannelGroups(Groupcount).Properties("CYCLE_CODE").Value&"_"&Data.Root.ChannelGroups(Groupcount).Properties("CIRCUIT_CODE").Value&".TDV")
           Call msgbox("View template for ABB "&Data.Root.ChannelGroups(Groupcount).Properties("TESTING_CODE").Value&" "&Data.Root.ChannelGroups(Groupcount).Properties("CYCLE_CODE").Value&" "&Data.Root.ChannelGroups(Groupcount).Properties("CIRCUIT_CODE").Value&" loaded!")
          Elseif FilEx("C:\DIAdem\user_data\layouts\Leistungslabor\View_ABB_"&Data.Root.ChannelGroups(Groupcount).Properties("TESTING_CODE").Value&"_"&Data.Root.ChannelGroups(Groupcount).Properties("CYCLE_CODE").Value&".TDV") Then
           Call View.LoadLayout("C:\DIAdem\user_data\layouts\Leistungslabor\View_ABB_"&Data.Root.ChannelGroups(Groupcount).Properties("TESTING_CODE").Value&"_"&Data.Root.ChannelGroups(Groupcount).Properties("CYCLE_CODE").Value&".TDV")
           Call msgbox("View template for ABB "&Data.Root.ChannelGroups(Groupcount).Properties("TESTING_CODE").Value&" "&Data.Root.ChannelGroups(Groupcount).Properties("CYCLE_CODE").Value&" loaded and this will be saved also with the correct ciucuit code name!")
           Call View.SaveLayout("C:\DIAdem\user_data\layouts\Leistungslabor\View_ABB_"&Data.Root.ChannelGroups(Groupcount).Properties("TESTING_CODE").Value&"_"&Data.Root.ChannelGroups(Groupcount).Properties("CYCLE_CODE").Value&"_"&Data.Root.ChannelGroups(Groupcount).Properties("CIRCUIT_CODE").Value&".TDV")
          Else
           Call View.LoadLayout("C:\DIAdem\user_data\layouts\Leistungslabor\View_ABB_default.TDV")
           Call msgbox("View default template for ABB loaded and saved with the correct name!")
           Call View.SaveLayout("C:\DIAdem\user_data\layouts\Leistungslabor\View_ABB_"&Data.Root.ChannelGroups(Groupcount).Properties("TESTING_CODE").Value&"_"&Data.Root.ChannelGroups(Groupcount).Properties("CYCLE_CODE").Value&"_"&Data.Root.ChannelGroups(Groupcount).Properties("CIRCUIT_CODE").Value&".TDV")
          End if
        End if 'Entwicklung
        
        If len(Data.Root.ChannelGroups(Groupcount).Name) > 17  Then   'Pehlanummer
      
          If FilEx("C:\DIAdem\user_data\layouts\Leistungslabor\View_PEHLA_"&Data.Root.ChannelGroups(Groupcount).Properties("TESTING_CODE").Value&"_"&Data.Root.ChannelGroups(Groupcount).Properties("CYCLE_CODE").Value&"_"&Data.Root.ChannelGroups(Groupcount).Properties("CIRCUIT_CODE").Value&".TDV") Then
           Call View.LoadLayout("C:\DIAdem\user_data\layouts\Leistungslabor\View_PEHLA_"&Data.Root.ChannelGroups(Groupcount).Properties("TESTING_CODE").Value&"_"&Data.Root.ChannelGroups(Groupcount).Properties("CYCLE_CODE").Value&"_"&Data.Root.ChannelGroups(Groupcount).Properties("CIRCUIT_CODE").Value&".TDV")
           Call msgbox("View template for PEHLA "&Data.Root.ChannelGroups(Groupcount).Properties("TESTING_CODE").Value&" "&Data.Root.ChannelGroups(Groupcount).Properties("CYCLE_CODE").Value&" "&Data.Root.ChannelGroups(Groupcount).Properties("CIRCUIT_CODE").Value&" loaded!")
          Elseif FilEx("C:\DIAdem\user_data\layouts\Leistungslabor\View_PEHLA_"&Data.Root.ChannelGroups(Groupcount).Properties("TESTING_CODE").Value&"_"&Data.Root.ChannelGroups(Groupcount).Properties("CYCLE_CODE").Value&".TDV") Then
           Call View.LoadLayout("C:\DIAdem\user_data\layouts\Leistungslabor\View_PEHLA_"&Data.Root.ChannelGroups(Groupcount).Properties("TESTING_CODE").Value&"_"&Data.Root.ChannelGroups(Groupcount).Properties("CYCLE_CODE").Value&".TDV")
           Call msgbox("View template for PEHLA "&Data.Root.ChannelGroups(Groupcount).Properties("TESTING_CODE").Value&" "&Data.Root.ChannelGroups(Groupcount).Properties("CYCLE_CODE").Value&" loaded and this will be saved also with the correct ciucuit code name!")
           Call View.SaveLayout("C:\DIAdem\user_data\layouts\Leistungslabor\View_PEHLA_"&Data.Root.ChannelGroups(Groupcount).Properties("TESTING_CODE").Value&"_"&Data.Root.ChannelGroups(Groupcount).Properties("CYCLE_CODE").Value&"_"&Data.Root.ChannelGroups(Groupcount).Properties("CIRCUIT_CODE").Value&".TDV")
          Else
           Call View.LoadLayout("C:\DIAdem\user_data\layouts\Leistungslabor\View_PEHLA_default.TDV")
           Call msgbox("View default template for PEHLA loaded and saved with the correct name!")
           Call View.SaveLayout("C:\DIAdem\user_data\layouts\Leistungslabor\View_PEHLA_"&Data.Root.ChannelGroups(Groupcount).Properties("TESTING_CODE").Value&"_"&Data.Root.ChannelGroups(Groupcount).Properties("CYCLE_CODE").Value&"_"&Data.Root.ChannelGroups(Groupcount).Properties("CIRCUIT_CODE").Value&".TDV")
          End if
        End if 'Pehla
  
      End if ' labcode = 1
   
  End if 'view template = 1
    
    '---------------------------------------------------------------------------------------
    'end of loading templates
    
    
  End if 'Window offline.sud
  
  'Never remove: Reset of the Script-Engine and Menu-Crash would be forced
  'Call ScriptCmdRemove(AutoActPath & "..\equidistant_x\equix.vbs") 'Must be in the same Path as this File
  'Call ScriptCmdRemove(AutoActPath & "importrbs.vbs")
End sub
'-------------------------------------------------------------------------------