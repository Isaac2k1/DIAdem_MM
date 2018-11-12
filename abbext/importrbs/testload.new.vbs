'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2006/12/21 10:50:16
'-- Author: Kaan Oenen (Update Mathias Knaak)
'-- Comment: Main Script for "Load an old Test" Menu
'-------------------------------------------------------------------------------

'v2.7 select if digital channels should be loaded
'v2.8 select if error data should be loaded

'30.03.2010: Fabio Maurizio
'new variable viewtemplate added in order to load the different kinds of templates

'07.07.2010: Fabio Maurizio
'two new variables added for the new function: "load single channels"

Call testloadnew()

'-------------------------------------------------------------------------------
'Sub procedure
'testload.new()
'compatible with dialogbox "Offline_v2.8"
'
'Description
'
'-------------------------------------------------------------------------------
'Variables
'K: Number of files
  Call Globaldim("load_channel")


Sub testloadnew()
  
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
  
  If SudDlgShow("Dlg1",AutoActPath & "offline.new.sud") = "IDOk" Then
    'For I = 0 to K  'Ubound(sDataName)
    '  NameString = NameString&" "&sDataName(I)&" "&sReportName(I)
    'Next
    
    
    'Loading the different view templates for every kind of test
    '-------------------------------------------------------------------------------------
    
    If view_template = 1 Then
      
      'ONLY FOR HIGH POWER
      '2(ABB or PEHLA)x2(mechanical or electrical)x4(cycle code)+2(3s-operation)=18 templates 
    If labcode = 1 Then
      
      If (Data.Root.ChannelGroups(1).Properties("VERSUCHS_TYP").Value="ENTWICKLUNG" or Data.Root.ChannelGroups(1).Properties("VERSUCHS_TYP").Value="ABNAHME ABB"or Data.Root.ChannelGroups(1).Properties("VERSUCHS_TYP").Value="TYPPRUEFUNG ABB"or Data.Root.ChannelGroups(1).Properties("VERSUCHS_TYP").Value="TYPPRUEFUNG"or Data.Root.ChannelGroups(1).Properties("VERSUCHS_TYP").Value="UNTERHALT")  Then
        If Data.Root.ChannelGroups(1).Properties("TESTING_CODE").Value="ELECTRICAL" Then 
          
          If Data.Root.ChannelGroups(1).Properties("CYCLE_CODE").Value="O" Then
            Call View.LoadLayout("I:\Com\DIAdem\DIAdem\user_data\layouts\Leistungslabor\View_ABB_electrical_O.TDV")
            Call msgbox("View template for ABB Electrical O loaded!")
          End if
          If Data.Root.ChannelGroups(1).Properties("CYCLE_CODE").Value="C" Then
            Call View.LoadLayout("I:\Com\DIAdem\DIAdem\user_data\layouts\Leistungslabor\View_ABB_electrical_C.TDV")
            Call msgbox("View template for ABB Electrical C loaded!")
          End if
          If Data.Root.ChannelGroups(1).Properties("CYCLE_CODE").Value="CO" Then
            If Data.Root.ChannelGroups(1).Properties("CIRCUIT_CODE").Value="DIRECT 1 PHASE" Then
              Call View.LoadLayout("I:\Com\DIAdem\DIAdem\user_data\layouts\Leistungslabor\View_ABB_electrical_CO_direct1Phase.TDV")
              Call msgbox("View template for ABB Electrical CO direct1Phase loaded!")
            Else
              Call View.LoadLayout("I:\Com\DIAdem\DIAdem\user_data\layouts\Leistungslabor\View_ABB_electrical_CO.TDV")
              Call msgbox("View template for ABB Electrical CO loaded!")
            End if
          End if
          If Data.Root.ChannelGroups(1).Properties("CYCLE_CODE").Value="O-CO" Then
            Call View.LoadLayout("I:\Com\DIAdem\DIAdem\user_data\layouts\Leistungslabor\View_ABB_electrical_O-CO.TDV")
            Call msgbox("View template for ABB Electrical O-CO loaded!")
          End if
          If Data.Root.ChannelGroups(1).Properties("CYCLE_CODE").Value="SHORT 3 SEC" Then
            Call View.LoadLayout("I:\Com\DIAdem\DIAdem\user_data\layouts\Leistungslabor\View_ABB_electrical_3sO.TDV")
            Call msgbox("View template for ABB Electrical 3sO loaded!")
          End if
          
        End if 'electrical
        If Data.Root.ChannelGroups(1).Properties("TESTING_CODE").Value="MECHANICAL TB" Then 
          
          If Data.Root.ChannelGroups(1).Properties("CYCLE_CODE").Value="O" Then
            Call View.LoadLayout("I:\Com\DIAdem\DIAdem\user_data\layouts\Leistungslabor\View_ABB_mechanical_O.TDV")
            Call msgbox("View template for ABB Mechanical O loaded!")
          End if
          If Data.Root.ChannelGroups(1).Properties("CYCLE_CODE").Value="C" Then
            Call View.LoadLayout("I:\Com\DIAdem\DIAdem\user_data\layouts\Leistungslabor\View_ABB_mechanical_C.TDV")
            Call msgbox("View template for ABB Mechanical C loaded!")
          End if
          If Data.Root.ChannelGroups(1).Properties("CYCLE_CODE").Value="CO" Then
            If Data.Root.ChannelGroups(1).Properties("CIRCUIT_CODE").Value="DIRECT 1 PHASE" Then
              Call View.LoadLayout("I:\Com\DIAdem\DIAdem\user_data\layouts\Leistungslabor\View_ABB_mechanical_CO_direct1Phase.TDV")
              Call msgbox("View template for ABB Mechanical CO direct1Phase loaded!")
            Else
              Call View.LoadLayout("I:\Com\DIAdem\DIAdem\user_data\layouts\Leistungslabor\View_ABB_mechanical_CO.TDV")
              Call msgbox("View template for ABB Mechanical CO loaded!")
            End if
          End if
          If Data.Root.ChannelGroups(1).Properties("CYCLE_CODE").Value="O-CO" Then
            Call View.LoadLayout("I:\Com\DIAdem\DIAdem\user_data\layouts\Leistungslabor\View_ABB_mechanical_O-CO.TDV")
            Call msgbox("View template for ABB Mechanical O-CO loaded!")
          End if
          
        End if 'mechanical
      End if 'Entwicklung
      '-----
      If (Data.Root.ChannelGroups(1).Properties("VERSUCHS_TYP").Value="TYPPRUEFUNG PEHLA" Or Data.Root.ChannelGroups(1).Properties("VERSUCHS_TYP").Value="ABNAHME PEHLA") Then
        If Data.Root.ChannelGroups(1).Properties("TESTING_CODE").Value="ELECTRICAL" Then 
          
          If Data.Root.ChannelGroups(1).Properties("CYCLE_CODE").Value="O" Then
            Call View.LoadLayout("I:\Com\DIAdem\DIAdem\user_data\layouts\Leistungslabor\View_PEHLA_electrical_O.TDV")
            Call msgbox("View template for PEHLA Electrical O loaded!")
          End if
          If Data.Root.ChannelGroups(1).Properties("CYCLE_CODE").Value="C" Then
            Call View.LoadLayout("I:\Com\DIAdem\DIAdem\user_data\layouts\Leistungslabor\View_PEHLA_electrical_C.TDV")
            Call msgbox("View template for PEHLA Electrical C loaded!")
          End if
          If Data.Root.ChannelGroups(1).Properties("CYCLE_CODE").Value="CO" Then
            Call View.LoadLayout("I:\Com\DIAdem\DIAdem\user_data\layouts\Leistungslabor\View_PEHLA_electrical_CO.TDV")
            Call msgbox("View template for PEHLA Electrical CO loaded!")
          End if
          If Data.Root.ChannelGroups(1).Properties("CYCLE_CODE").Value="O-CO" Then
            Call View.LoadLayout("I:\Com\DIAdem\DIAdem\user_data\layouts\Leistungslabor\View_PEHLA_electrical_O-CO.TDV")
            Call msgbox("View template for PEHLA Electrical O-CO loaded!")
          End if
          If Data.Root.ChannelGroups(1).Properties("CYCLE_CODE").Value="SHORT 3 SEC" Then
            Call View.LoadLayout("I:\Com\DIAdem\DIAdem\user_data\layouts\Leistungslabor\View_PEHLA_electrical_3sO.TDV")
            Call msgbox("View template for PEHLA Electrical 3sO loaded!")
          End if
          
        End if 'electrical
        If Data.Root.ChannelGroups(1).Properties("TESTING_CODE").Value="MECHANICAL TB" Then 
          
          If Data.Root.ChannelGroups(1).Properties("CYCLE_CODE").Value="O" Then
            Call View.LoadLayout("I:\Com\DIAdem\DIAdem\user_data\layouts\Leistungslabor\View_PEHLA_mechanical_O.TDV")
            Call msgbox("View template for PEHLA Mechanical O loaded!")
          End if
          If Data.Root.ChannelGroups(1).Properties("CYCLE_CODE").Value="C" Then
            Call View.LoadLayout("I:\Com\DIAdem\DIAdem\user_data\layouts\Leistungslabor\View_PEHLA_mechanical_C.TDV")
            Call msgbox("View template for PEHLA Mechanical C loaded!")
          End if
          If Data.Root.ChannelGroups(1).Properties("CYCLE_CODE").Value="CO" Then
            Call View.LoadLayout("I:\Com\DIAdem\DIAdem\user_data\layouts\Leistungslabor\View_PEHLA_mechanical_CO.TDV")
            Call msgbox("View template for PEHLA Mechanical CO loaded!")
          End if
          If Data.Root.ChannelGroups(1).Properties("CYCLE_CODE").Value="O-CO" Then
            Call View.LoadLayout("I:\Com\DIAdem\DIAdem\user_data\layouts\Leistungslabor\View_PEHLA_mechanical_O-CO.TDV")
            Call msgbox("View template for PEHLA Mechanical O-CO loaded!")
          End if
          
        End if 'mechanical
      End if 'Pehla
  
  End if ' labcode = 1
   
    'ONLY FOR MECHANIC
      '4(cycle code)templates 
   If labcode = 3 Then
          If Data.Root.ChannelGroups(1).Properties("TIMINGCODE").Value="O" Then
            Call View.LoadLayout("I:\Com\DIAdem\DIAdem\user_data\layouts\Mechanic\Mechanic_O.TDV")
            Call msgbox("View template for Cyclecode O loaded!")
          End if
          If Data.Root.ChannelGroups(1).Properties("TIMINGCODE").Value="C" Then
            Call View.LoadLayout("I:\Com\DIAdem\DIAdem\user_data\layouts\Mechanic\Mechanic_C.TDV")
            Call msgbox("View template for Cyclecode C loaded!")
          End if
          If Data.Root.ChannelGroups(1).Properties("TIMINGCODE").Value="CO" Then
            Call View.LoadLayout("I:\Com\DIAdem\DIAdem\user_data\layouts\Mechanic\Mechanic_CO.TDV")
            Call msgbox("View template for Cyclecode CO loaded!")
          End if
          If Data.Root.ChannelGroups(1).Properties("TIMINGCODE").Value="O-CO" Then
            Call View.LoadLayout("I:\Com\DIAdem\DIAdem\user_data\layouts\Mechanic\Mechanic_O-CO.TDV")
            Call msgbox("View template for Cyclecode O-CO loaded!")
          End if
  
    End if 'labcode = 3
  
  End if 'view template = 1
    
    '---------------------------------------------------------------------------------------
    'end of loading templates
    
    
  End if 'Window offline.new.sud
  
  'Never remove: Reset of the Script-Engine and Menu-Crash would be forced
  'Call ScriptCmdRemove(AutoActPath & "..\equidistant_x\equix.vbs") 'Must be in the same Path as this File
  'Call ScriptCmdRemove(AutoActPath & "importrbs.vbs")
End sub
'-------------------------------------------------------------------------------