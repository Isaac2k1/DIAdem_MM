'-------------------------------------------------------------------------------
'-- VBS-Script-File
'-- Writer: René Irion
'-- Summary:
'-- This programm correct the time shift of the LASER measurement system (Keyence).
'-- Information about the used equation you will find in the english user manual on page 6-22.
'--
'-- All needed information must be stored with the following names in the properties of a shot.
'-- LM_count         [Possible Values: 1, 2, ...] 
'-- LM_SetIdent      [The settings of both systems are identically: 0=No, 1=Yes]
'-- 
'-- Description of the array
'-- LaserMeasurement_SystemNumber_Function_ChannelNumber
'-- 
'-- System 1:
'-- LM_1_sampletime  [µs]
'-- LM_1_channel_1   [Channel name]
'-- LM_1_channel_2   [Channel name]
'-- LM_1_filter_1    [1=Average, 2=Low Pass, 3=High Pass]
'-- LM_1_filter_2    [1=Average, 2=Low Pass, 3=High Pass]
'-- LM_1_average_1   [Possible Values: 1, 4, 16, 64, 256, 1024, 4096, 16384, 65536, 262144] 
'-- LM_1_average_2   [Possible Values: 1, 4, 16, 64, 256, 1024, 4096, 16384, 65536, 262144] 
'-- 
'-- Optional information in a shot
'-- LM_1_VanHand     [µs]
'--
'-- System 2:
'-- LM_2_sampletime  [µs]
'-- LM_2_channel_1   [Channel name]
'-- LM_2_channel_2   [Channel name]
'-- LM_2_filter_1    [1=Average, 2=Low Pass, 3=High Pass]
'-- LM_2_filter_2    [1=Average, 2=Low Pass, 3=High Pass]
'-- LM_2_average_1   [Possible Values: 1, 4, 16, 64, 256, 1024, 4096, 16384, 65536, 262144] 
'-- LM_2_average_2   [Possible Values: 1, 4, 16, 64, 256, 1024, 4096, 16384, 65536, 262144] 
'-- 
'-- Optional information in a shot
'-- LM_2_VanHand     [µs] 
'--
'-- History
'-- 2014-10-23
'-- Single channel mode added. Now its posible to use only one channel from a measurement system
'-- eg to realize 3 measurements
'--
'-- 2014-08-06
'-- Version 2 finalized
'-- Some additional scripting specially for the UFD project.
'-- The lines 146 can be deleted when the next analysis script is working
'-- The sub process "PropertyChange()" is neccessary for further work
'--
'-- 2013-01-06
'-- Implement "LM_VanHand" in a better way
'-- Improve the script for a second, third... identical laser measurement system (loop)
'-- Also the old shots with the former settings can be manipulated with this script
'-- 
'-- 2012-10-10
'-- Add experimental time shift "LM_VanHand"
'-- Experimental value to reduce the time shift. Be careful,
'-- will work fine only when both laser has the same setting!
'-- René Irion
'-- If there any questions, dont hesistate to ask me --> Rene.Irion@ch.abb.com
'-- 
'-- 2009-12-07
'-- First Programm
'-- René Irion
'-- Automatic execution through "importrbs.vbs"
'-------------------------------------------------------------------------------
Option Explicit

'declaration of independent variables
  'Abbreviation of the measurement system
  Dim Measurement(2)
  Measurement(0) = "LM_"
  Measurement(1) = "count"
  Measurement(2) = "SetIdent"
  
  'Name
  Dim MeasurementProperty(1)
  MeasurementProperty(0) = Measurement(0) & Measurement(1)      'Channel name for "Number of measurements"
  MeasurementProperty(1) = Measurement(0) & Measurement(2)      'Channel name for "Identical settings"
  
  'Value
  Dim MeasurementPropertyValue(1)
  MeasurementPropertyValue(0) = "0"                             'Number of measurements                      
  MeasurementPropertyValue(1) = "0"                             'Identical settings 0=No, 1=Yes
  
  'OK flag
  Dim MeasurementPropertyValueOK(1)
  MeasurementPropertyValueOK(0) = "0"                             
  MeasurementPropertyValueOK(1) = "0"                             
  
  'Ignore shot flag
  Dim ShotIgnore
  ShotIgnore = 0

  'Manual Mode flag
  Dim ManualMode
  ManualMode = 0
  
  'Flag for open the config window
  Dim OpenConfigWindow
  OpenConfigWindow = 0

  'Blubb...
  Dim AllOrOne_start, AllOrOne_end

  Dim count_a, count_b, count_c
  count_a = 0
  count_b = 0
  count_c = 0
  
  'declaration of dependent variables
  If Not Iteminfoget("NameMeasurementProperty") Then
    GlobalDim ("NameMeasurementProperty()")
  End If
  
  If Not Iteminfoget("NameProperty") Then
    GlobalDim("NameProperty()")
  End If
  
  If Not Iteminfoget("NamePropertyValue") Then
    GlobalDim(" NamePropertyValue()")
  End If
  
  If Not Iteminfoget("NamePropertyValueOK") Then
    GlobalDim(" NamePropertyValueOK()")
  End If
  
  Dim TimeShift()
  Dim FactorFilter()
  Dim FactorRefresh()
  Dim points_to_delete()
  Dim x_increment()
  Dim SingleChannelMode() 
  
'-------------------------------------------------------------------------------
'Programms runs in automatic mode
'It means that the programm would be started when loading an test
Sub automatic()

'Exists the property "LM_channel1" and "LM_channel2"?
'We need to check this because a lot of old tests doesnt have the new settings
Call PropertyExistsSingleSys()

'Exists the property "LM_count" and "LM_SetIdent"?
Call PropertyExistsMultipleSys()

'When both cannels doesn't exist is nothing to do
If MeasurementPropertyValueOK(0) = 0 and MeasurementPropertyValueOK(1) = 0 then
  
Else
  
  'Test for UFD Programming
  'Delete after successful development of Rebeccas Script
 
  If FilEx("C:\DIAdem\user_data\scripts\Test_Kanalerstellung.VBS") Then
    Call Scriptstart("C:\DIAdem\user_data\scripts\Test_Kanalerstellung.VBS")
  Else
    
  End If
     
  
  'declaration of dependent variables
  Call redeclaration()

  'Read all properties
  Call PropertyRead()
  
  'Check plausibility of properties
  Call PropertyPlausibility()
  
  'When something is wrong this command will open a dialog
  Call PropertyMissing()
  
  'Skip this shot when user is pressing the cancel button
  If ShotIgnore = 0 Then

    'calculate the time shift
    Call TimeShiftCalculate()
  
    'delete the time shift
    Call TimeShiftDelete()

  End If

End If

End Sub

'-------------------------------------------------------------------------------
'Programms runs in manual mode
'It means that the programm would be started when using the ABB menu
Sub manual()

'Flag for open the dialog
ManualMode = 1

'How many systems are used?
  MeasurementPropertyValue(0) = InputBox("Enter the number of measurement systems","How many systems are used?","1") 


'declaration of dependent variables
  Call redeclaration()

for count_a = 0 to MeasurementPropertyValue(0) - 1
  'Set variables to some general values
  NamePropertyValue(count_a, 0) = "Laser A"  'channel 1
  NamePropertyValue(count_a, 1) = "Laser B"  'channel 2
  NamePropertyValue(count_a, 2) = "100"      'sampling rate
  NamePropertyValue(count_a, 3) = "1"        'Filter of channel 1
  NamePropertyValue(count_a, 4) = "1"        'Filter of channel 2
  NamePropertyValue(count_a, 5) = "1"        'Averaging times of channel 1
  NamePropertyValue(count_a, 6) = "1"        'Averaging times of channel 2
  NamePropertyValue(count_a, 7) = "548"      '

  'Ignore the plausibility check
  For count_b = 0 to 7
    NamePropertyValueOK(count_a, count_b) = 1
  Next

Next


'ask to use the settings for all laser measurements
'when yes --> loop over all groups with this settings
'when no --> only the active group will be corrected
If MsgBox("Use the following settings for all LASER measurements in the existing groups in 'data portal'?" & vbCrLf & _
  " " & vbCrLf & _
  "Yes: All LASER measurements will be corrected with this settings" & vbCrLf & _
  "No: Only the LASER measurements in the active group will be corrected with this settings", 4, "Use settings for all LASER measurements?") = 6 Then		'4 --> vbYesNo; 6 --> Yes
  
  'start and end value for loop over all groups
  AllOrOne_start = 1
  AllOrOne_end = GroupCount
  
Else
  'start and end value for one loop (only active group)
  AllOrOne_start = GroupDefaultGet
  AllOrOne_end = GroupDefaultGet
  
End If


'open the dialog
Call PropertyMissing()

  'Skip this shot when user is pressing the cancel button
  If ShotIgnore = 0 Then

    'loop over all groups
    For count_c = AllOrOne_start to AllOrOne_end
      Call GroupDefaultSet(count_c)

      'calculate the time shift
      Call TimeShiftCalculate()
  
      'delete the time shift
      Call TimeShiftDelete()

    Next
  
  End If

End Sub

'-------------------------------------------------------------------------------
'summary:   Check if the old property "LM_channel1" and "LM_channel2" exists and
'           fix it for the new version
Sub PropertyExistsSingleSys()
  
  'generate new settings
  If Data.Root.ChannelGroups(GroupDefaultGet).Properties.Exists("LM_channel_1") or Data.Root.ChannelGroups(GroupDefaultGet).Properties.Exists("LM_channel_2") Then
    Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_count", "1", DataTypeString)
    Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_SetIdent", "0", DataTypeString)
  
    'channel 1
    If Data.Root.ChannelGroups(GroupDefaultGet).Properties.Exists("LM_channel_1") Then
      Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_1_channel_1", Data.Root.ChannelGroups(GroupDefaultGet).Properties("LM_channel_1").Value , DataTypeString)
      Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Remove("LM_channel_1")
    End If
  
    'channel 2
    If Data.Root.ChannelGroups(GroupDefaultGet).Properties.Exists("LM_channel_2") Then
      Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_1_channel_2", Data.Root.ChannelGroups(GroupDefaultGet).Properties("LM_channel_2").Value , DataTypeString)
      Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Remove("LM_channel_2")
    End If
  
    'filter 1
    If Data.Root.ChannelGroups(GroupDefaultGet).Properties.Exists("LM_filter_1") Then
      Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_1_filter_1", Data.Root.ChannelGroups(GroupDefaultGet).Properties("LM_filter_1").Value , DataTypeString)
      Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Remove("LM_filter_1")
    End If
    
    'filter 2
    If Data.Root.ChannelGroups(GroupDefaultGet).Properties.Exists("LM_filter_2") Then
      Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_1_filter_2", Data.Root.ChannelGroups(GroupDefaultGet).Properties("LM_filter_2").Value , DataTypeString)
      Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Remove("LM_filter_2")
    End If  
  
    'average 1
    If Data.Root.ChannelGroups(GroupDefaultGet).Properties.Exists("LM_average_1") Then
      Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_1_average_1", Data.Root.ChannelGroups(GroupDefaultGet).Properties("LM_average_1").Value , DataTypeString)
      Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Remove("LM_average_1")
    End If  
  
    'average 2
    If Data.Root.ChannelGroups(GroupDefaultGet).Properties.Exists("LM_average_2") Then
      Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_1_average_2", Data.Root.ChannelGroups(GroupDefaultGet).Properties("LM_average_2").Value , DataTypeString)
      Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Remove("LM_average_2")
    End If  
  
    'sampletime
    If Data.Root.ChannelGroups(GroupDefaultGet).Properties.Exists("LM_sampletime") Then
      Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_1_sampletime", Data.Root.ChannelGroups(GroupDefaultGet).Properties("LM_sampletime").Value , DataTypeString)
      Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Remove("LM_sampletime")
    End If  
    
    'VanHand
    If Data.Root.ChannelGroups(GroupDefaultGet).Properties.Exists("LM_VanHand") Then
      Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Add("LM_1_VanHand", Data.Root.ChannelGroups(GroupDefaultGet).Properties("LM_VanHand").Value , DataTypeString)
      Call Data.Root.ChannelGroups(GroupDefaultGet).Properties.Remove("LM_VanHand")
    End If  
  
  End If
  
End Sub

'-------------------------------------------------------------------------------
'summary:   Exists the new property "LM_count" and "LM_SetIdent"
Sub PropertyExistsMultipleSys()

'Loop over both properties
For count_a = 0 to 1

  'when property exists, then set the OK Flag
  '0 --> property doesn't exists --> skip this shot
  '1 --> property exists
  
  If Data.Root.ChannelGroups(GroupDefaultGet).Properties.Exists(MeasurementProperty(count_a)) Then
    MeasurementPropertyValue(count_a) = Data.Root.ActiveChannelGroup.Properties(MeasurementProperty(count_a)).Value
    MeasurementPropertyValueOK(count_a) = 1
  Else
    MeasurementPropertyValueOK(count_a) = 0
  End If
  
next

End Sub

'-------------------------------------------------------------------------------
'redeclaration of dependent variables
Sub redeclaration()

  'redeclaration of dependent variables
  GlobalReDim ("NameMeasurementProperty(" & MeasurementPropertyValue(0)-1 & "," & 4 & ")")
  GlobalReDim ("NameProperty(" & MeasurementPropertyValue(0)-1 & "," & 7 & ")")
  GlobalReDim ("NamePropertyValue(" & MeasurementPropertyValue(0)-1 & "," & 7 & ")")
  GlobalReDim ("NamePropertyValueOK(" & MeasurementPropertyValue(0)-1 & "," & 7 & ")")
  
  ReDim TimeShift(MeasurementPropertyValue(0)-1, 3)
  ReDim FactorFilter(MeasurementPropertyValue(0)-1, 1)
  ReDim FactorRefresh(MeasurementPropertyValue(0)-1, 1)
  ReDim points_to_delete(MeasurementPropertyValue(0)-1, 3)
  ReDim x_increment(MeasurementPropertyValue(0)-1, 3)
  ReDim SingleChannelMode(MeasurementPropertyValue(0)-1)
  
    For count_a = 0 to (MeasurementPropertyValue(0)-1)
      NameMeasurementProperty(count_a, 0) = "_channel_"
      NameMeasurementProperty(count_a, 1) = "_sampletime"         
      NameMeasurementProperty(count_a, 2) = "_filter_"         
      NameMeasurementProperty(count_a, 3) = "_average_"
      NameMeasurementProperty(count_a, 4) = "_VanHand"

      'Name of the property
      NameProperty(count_a, 0) = Measurement(0) & count_a + 1 & NameMeasurementProperty(count_a, 0) & "1"   'channel 1
      NameProperty(count_a, 1) = Measurement(0) & count_a + 1 & NameMeasurementProperty(count_a, 0) & "2"   'channel 2
      NameProperty(count_a, 2) = Measurement(0) & count_a + 1 & NameMeasurementProperty(count_a, 1)         'sampling rate
      NameProperty(count_a, 3) = Measurement(0) & count_a + 1 & NameMeasurementProperty(count_a, 2) & "1"   'Filter of channel 1
      NameProperty(count_a, 4) = Measurement(0) & count_a + 1 & NameMeasurementProperty(count_a, 2) & "2"   'Filter of channel 2
      NameProperty(count_a, 5) = Measurement(0) & count_a + 1 & NameMeasurementProperty(count_a, 3) & "1"   'Averaging times of channel 1
      NameProperty(count_a, 6) = Measurement(0) & count_a + 1 & NameMeasurementProperty(count_a, 3) & "2"   'Averaging times of channel 2
      NameProperty(count_a, 7) = Measurement(0) & count_a + 1 & NameMeasurementProperty(count_a, 4)         'Experimental value to reduce the time shift of both channels

      'Set defined values
      For count_b = 0 to 7
        NamePropertyValue(count_a, count_b) = "0"
        NamePropertyValueOK(count_a, count_b) = "0"

        Select Case count_b
          Case 0, 1, 2, 3
            TimeShift(count_a, count_b) = 0
            points_to_delete(count_a, count_b) = 0
            x_increment(count_a, count_b) = 0

          Case Else

        End Select
        
        Select Case count_b
          Case 0, 1
            FactorFilter(count_a, count_b) = 0
            FactorRefresh(count_a, count_b) = 0

          Case Else
          
        End Select
        
      Next
      
    Next

end sub

'-------------------------------------------------------------------------------
'summary:   Read all properties of a shot
Sub PropertyRead()

'when property exists, read out the value
'else set it to 0

If MeasurementPropertyValue(1) = 0 then
'read all
  For count_a = 0 to (MeasurementPropertyValue(0)-1)
    For count_b = 0 to 7
      If Data.Root.ChannelGroups(GroupDefaultGet).Properties.Exists(NameProperty(count_a, count_b)) Then
        NamePropertyValue(count_a, count_b) = Data.Root.ChannelGroups(GroupDefaultGet).Properties(NameProperty(count_a, count_b)).Value
      Else
        NamePropertyValue(count_a, count_b) = "0"
      End If
    Next
  Next

Else
  If MeasurementPropertyValue(1) = 1 then
  'identical settings
    
  'read the settings of the first measurement
    For count_b = 0 to 7
      If Data.Root.ChannelGroups(GroupDefaultGet).Properties.Exists(NameProperty(0, count_b)) Then
        NamePropertyValue(0, count_b) = Data.Root.ChannelGroups(GroupDefaultGet).Properties(NameProperty(0, count_b)).Value
      Else
        NamePropertyValue(0, count_b) = "0"
      End If
    Next

  'set the settings of the other measurement systems to the value from the first one
    For count_a = 1 to (MeasurementPropertyValue(0)-1)
      For count_b = 2 to 7
        NamePropertyValue(count_a, count_b) = NamePropertyValue(0, count_b)
      Next
    Next

  'read all channel names
    For count_a = 1 to (MeasurementPropertyValue(0)-1)
      For count_b = 0 to 1
        If Data.Root.ChannelGroups(GroupDefaultGet).Properties.Exists(NameProperty(count_a, count_b)) Then
          NamePropertyValue(count_a, count_b) = Data.Root.ChannelGroups(GroupDefaultGet).Properties(NameProperty(count_a, count_b)).Value
        Else
          NamePropertyValue(count_a, count_b) = "0"
        End If
      Next
    Next
  
  Else
  
  End if 

End if

end sub


'-------------------------------------------------------------------------------
'summary:   Checks the plausibility of the read values
Sub PropertyPlausibility()

For count_a = 0 to MeasurementPropertyValue(0)-1
  For count_b = 0 to 7
    Select Case count_b
    Case 0, 1
      'channel exist
      If Data.Root.ChannelGroups(GroupDefaultGet).Channels.Exists(NamePropertyValue(count_a, count_b)) Then
        NamePropertyValueOK(count_a, count_b) = 1                          'channel exists
      Else
        NamePropertyValueOK(count_a, count_b) = 0                          'channel don't exists
      End If
          
    Case 2
      'plausibility of sample rate
      Select Case NamePropertyValue(count_a, count_b)
        Case "20", "50", "100", "200", "500", "1000"
          NamePropertyValueOK(count_a, count_b) = 1           'Value its possible
        
        Case Else
          NamePropertyValueOK(count_a, count_b) = 0           'Values must be checked
         
      End Select

    Case 3, 4
      'plausibility of filter  
      Select Case NamePropertyValue(count_a, count_b)
        Case "1", "2", "3"
          NamePropertyValueOK(count_a, count_b) = 1           'Value its possible
        
        Case Else
          NamePropertyValueOK(count_a, count_b) = 0           'Values must be checked
          
        End Select
      
    Case 5, 6
      'plausibility of average time  
      If NamePropertyValue(count_a, count_b - 2) = 1 Then
        Select Case NamePropertyValue(count_a, count_b)
          Case "1", "4", "16", "64", "256", "1024", "4096", "16384", "65536", "262144" 
            NamePropertyValueOK(count_a, count_b) = 1           'Value its possible
       
          Case Else
            NamePropertyValueOK(count_a, count_b) = 0           'Values must be checked
        
        End Select 
      End If
    
    Case 7
      NamePropertyValueOK(count_a, count_b) = 1                 'Optional
      
    End select

  Next

Next


For count_a = 0 to MeasurementPropertyValue(0)-1
  If Data.Root.ChannelGroups(GroupDefaultGet).Properties.Exists(NameProperty(count_a, 1)) and Data.Root.ChannelGroups(GroupDefaultGet).Properties.Exists(NameProperty(count_a, 4)) and Data.Root.ChannelGroups(GroupDefaultGet).Properties.Exists(NameProperty(count_a, 6)) Then
    SingleChannelMode(count_a) = 0

  Else
    SingleChannelMode(count_a) = 1  
    NamePropertyValueOK(count_a, 1) = 1
    NamePropertyValueOK(count_a, 4) = 1
    NamePropertyValueOK(count_a, 6) = 1
  End If
Next


End Sub

'-------------------------------------------------------------------------------
'summary:   Show the configuration window when a property is missing
Sub PropertyMissing()

'Will ignored during manual mode
For count_a = 0 to (MeasurementPropertyValue(0)-1)
  For count_b = 0 to 7
    If NamePropertyValueOK(count_a, count_b) = 0 Then
      OpenConfigWindow = 1
    End If
  Next
Next

'When somtheing is wrong in the config or ManualMode is active, show config window
If OpenConfigWindow or ManualMode = 1 Then

  'open the config window
  For count_a = 0 to (MeasurementPropertyValue(0)-1)
    Call SUDDlgShow("Dlg1", AutoActPath & "time_shift_of_laser_measurement.SUD", count_a)  
  Next
  
  If DlgState = "IDOk" Then
    ShotIgnore = 0
    
  Else
    ShotIgnore = 1
  
  End If

End If

End Sub


'-------------------------------------------------------------------------------
'summary:   Calculate the time shift of both laser
'parameter: 
Sub TimeShiftCalculate()

For count_a = 0 to (MeasurementPropertyValue(0)-1)
  
  For count_b = 0 to 1
  
    'Set factor for the euqation
    Select Case NamePropertyValue(count_a, count_b + 3)
      Case 1
        FactorFilter(count_a, count_b) = 1
    
      Case Else
        FactorFilter(count_a, count_b) = 0
    
      End Select
  
   'Set factor for the euqation
     Select Case NamePropertyValue(count_a, count_b + 5)
      Case 16384
        FactorRefresh(count_a, count_b) = 4
    
     Case 65536
        FactorRefresh(count_a, count_b) = 16
    
      Case 262144
        FactorRefresh(count_a, count_b) = 64
      
    Case Else
        FactorRefresh(count_a, count_b) = 1
      
    End Select

    'Calculate the time shift
    TimeShift(count_a, count_b) = ( FactorFilter(count_a, count_b) * NamePropertyValue(count_a, count_b + 5) + FactorRefresh(count_a, count_b) + 4 ) * NamePropertyValue(count_a, 2) 
  
    'Conversion to [s]
    TimeShift(count_a, count_b) = TimeShift(count_a, count_b) * 10^-6

    'VanHand Conversion to [s]
    TimeShift(count_a, count_b + 2) = NamePropertyValue(count_a, 7) * 10^-6 

  Next
    
Next

End Sub


'-------------------------------------------------------------------------------
'summary:   Correct the measurement channels
'parameter: 
Sub TimeShiftDelete()

For count_a = 0 to (MeasurementPropertyValue(0)-1)

  For count_b = 0 to 1

    If NamePropertyValue(count_a, count_b) <> "0" Then
    
      'rename channel
      Data.Root.ChannelGroups(GroupDefaultGet).Channels(NamePropertyValue(count_a, count_b)).Name = NamePropertyValue(count_a, count_b) & " - original"
      Data.Root.ChannelGroups(GroupDefaultGet).Channels(NamePropertyValue(count_a, count_b) & " - original").Properties("wf_xname").Value = NamePropertyValue(count_a, count_b) & " - original" & "_A_generated"

      'read x-increment value, calculate the points to delete, rount points to delete without decimals
      x_increment(count_a, count_b) = Data.Root.ChannelGroups(GroupDefaultGet).Channels(NamePropertyValue(count_a, count_b) & " - original").Properties("wf_increment").Value
      points_to_delete(count_a, count_b) = TimeShift(count_a, count_b) /  x_increment(count_a, count_b)
      points_to_delete(count_a, count_b) = round(points_to_delete(count_a, count_b),0)
    
      'Correct with calculated values
      If points_to_delete(count_a, count_b) >= 1 then
        Call ChnCopy("[" & GroupDefaultGet & "]/" & NamePropertyValue(count_a, count_b) & " - original", "[" & GroupDefaultGet & "]/" & NamePropertyValue(count_a, count_b) & " - corrected")
        Call ChnMove("[" & GroupDefaultGet & "]/" & NamePropertyValue(count_a, count_b) & " - corrected", GroupDefaultGet, Data.Root.ChannelGroups(GroupDefaultGet).Channels(NamePropertyValue(count_a, count_b) & " - original").Properties("groupindex").Value + 1)
        Call DataAreaDel("[" & GroupDefaultGet & "]/" & NamePropertyValue(count_a, count_b) & " - corrected",1,points_to_delete(count_a, count_b),NamePropertyValue(count_a, count_b) & " - corrected")
        Data.Root.ChannelGroups(GroupDefaultGet).Channels(NamePropertyValue(count_a, count_b) & " - corrected").Properties("wf_xname").Value = NamePropertyValue(count_a, count_b) & " - corrected" & "_A_generated"

      End IF
        
      'correct with experimental values    
      If TimeShift(count_a, count_b + 2) <> "0" Then
        points_to_delete(count_a, count_b + 2) = TimeShift(count_a, count_b + 2) /  x_increment(count_a, count_b)
        points_to_delete(count_a, count_b + 2) = round(points_to_delete(count_a, count_b + 2),0)

        If points_to_delete(count_a, count_b + 2) >= 1 then
          Call ChnCopy("[" & GroupDefaultGet & "]/" & NamePropertyValue(count_a, count_b) & " - original", "[" & GroupDefaultGet & "]/" & NamePropertyValue(count_a, count_b) & " - experimental")
          Call ChnMove("[" & GroupDefaultGet & "]/" & NamePropertyValue(count_a, count_b) & " - experimental", GroupDefaultGet, Data.Root.ChannelGroups(GroupDefaultGet).Channels(NamePropertyValue(count_a, count_b) & " - original").Properties("groupindex").Value + 2)        
          Call DataAreaDel("[" & GroupDefaultGet & "]/" & NamePropertyValue(count_a, count_b) & " - experimental",1,points_to_delete(count_a, count_b + 2),NamePropertyValue(count_a, count_b) & " - experimental")
          Data.Root.ChannelGroups(GroupDefaultGet).Channels(NamePropertyValue(count_a, count_b) & " - experimental").Properties("wf_xname").Value = NamePropertyValue(count_a, count_b) & " - experimental" & "_A_generated"

        End IF
   
      End if  

    Else
    
    End If

  Next
  
Next

End Sub


