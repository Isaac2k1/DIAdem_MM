'-------------------------------------------------------------------------------
'-- VBS-Script-Datei
'-- Neu erstellt am 15.10.2008 09:45:33
'-- Ersteller: 
'-- Kommentar: 
'-------------------------------------------------------------------------------
Option Explicit  'Erzwingt die explizite Deklaration aller Variablen in einem Skript.

View.Events.OnCursorChanged = ""

if groupindexget("__TangentCursor__")<>0 then
  call groupdel(groupindexget("__TangentCursor__"))
end if

msgbox "Tangent Cursor cleared"