'-------------------------------------------------------------------------------
'-- VBS-Script-Datei
'-- Neu erstellt am 16.09.2008 08:05:06
'-- Ersteller: 
'-- Kommentar: 
'-------------------------------------------------------------------------------
Option Explicit  'Erzwingt die explizite Deklaration aller Variablen in einem Skript.
 
Dim tfh, iTxtRows, sTemp, iLoop,text_string
  ' open text file
  tfh = TextFileOpen(autoactpath & "\formula_files\formula_1.txt", tfRead)
  If TextFileError(tfh) = 0 Then
    text_string=Textfilereadln(tfh)
  end if

  msgbox text_string