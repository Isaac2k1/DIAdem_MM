'-------------------------------------------------------------------------------
'-- VBS-Script-Datei
'-- Neu erstellt am 27.08.2008 08:36:07
'-- Ersteller: 
'-- Kommentar: 
'-------------------------------------------------------------------------------

Option Explicit  'Erzwingt die explizite Deklaration aller Variablen in einem Skript.


If Not ItemInfoGet("sPathDocuments") Then
  Call GlobalDim("sPathDocuments")
  Call GlobalDim("sPathData")
End If
sPathDocuments = AutoActPath
sPathData   = AutoActPath & "..\Data\"


' Funktionen aus "travel_functions" einbinden

Call ScriptCmdAdd(AutoActPath&"travel_functions.VBS")
Call ScriptCmdAdd(AutoActPath&"abb_eva_sub_collection.VBS")

'define variables

call globaldim ("comp_chn,ref_chn,scale_y,timeshift,offset_shift,t_start,t_end")
call globaldim ("time_scale")
t_start=0.1
t_end=0.2
timeshift=0
offset_shift=0
scale_y=1
time_scale=1


'call Dialog for Program

call suddlgshow("Dlg1", sPathDocuments & "Kanalvergleich")


