'-------------------------------------------------------------------------------
'-- VBS-Script-Datei
'-- Neu erstellt am 2009-12-09 15:24:21
'-- Ersteller: 
'-- Kommentar: 
'-------------------------------------------------------------------------------
Option Explicit  'Erzwingt die explizite Deklaration aller Variablen in einem Script.

'Call scriptstart(AutoActPath & "time_shift_of_laser_measurement.VBS", "manual")

Call scriptstart(AutoActPath & "time_shift_of_laser_measurement.VBS", "automatic")


'Call for abbext.vbs
'manual
'Dim DEV11 : DEV11 = "Call Scriptstart(""" & AutoActPath & "laser_time_shift\time_shift_of_laser_measurement.vbs"", ""manual"")"

'automatic
'Dim DEV11 : DEV11 = "Call Scriptstart(""" & AutoActPath & "laser_time_shift\time_shift_of_laser_measurement.vbs"", ""automatic"")"
