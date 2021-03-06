'-------------------------------------------------------------------------------
'-- VBS-Script-Datei
'-- Neu erstellt am 22.06.2010 09:22:04
'-- Ersteller: Nathan Rice
'--            modified by Tobias Wild 22.06.10 for NI DIAdem
'-- Kommentar: Provides functions to handle Zip-Files
'-------------------------------------------------------------------------------
Option Explicit  'Erzwingt die explizite Deklaration aller Variablen in einem Script.

Function WindowsUnZip(sUnzipFileName, sUnzipDestination)
  'This script is provided under the Creative Commons license located
  'at http://creativecommons.org/licenses/by-nc/2.5/ . It may not
  'be used for commercial purposes with out the expressed written consent
  'of NateRice.com
  
  Dim oUnzipFSO
  
  Set oUnzipFSO = CreateObject("Scripting.FileSystemObject")
 
  If Not oUnzipFSO.FolderExists(sUnzipDestination) Then
    oUnzipFSO.CreateFolder(sUnzipDestination)
  End If

  With CreateObject("Shell.Application")
       .NameSpace(sUnzipDestination).Copyhere .NameSpace(sUnzipFileName).Items
  End With

  Set oUnzipFSO = Nothing
End Function

Function WindowsZip(sFile, sZipFile)
  'This script is provided under the Creative Commons license located
  'at http://creativecommons.org/licenses/by-nc/2.5/ . It may not
  'be used for commercial purposes with out the expressed written consent
  'of NateRice.com

'  Dim oZipShell
  Dim oZipFSO
  Dim oZipApp
  Dim oZipFile
  Dim sZipFileCount
  Dim aFileName
  Dim sFileName
  Dim sDupe
  Dim sFileNameInZip

'  Set oZipShell = CreateObject("WScript.Shell")
  Set oZipFSO = CreateObject("Scripting.FileSystemObject")
  
  If Not oZipFSO.FileExists(sZipFile) Then
    Call NewZip(sZipFile, False)
  End If

  Set oZipApp = CreateObject("Shell.Application")
  
  sZipFileCount = oZipApp.NameSpace(sZipFile).items.Count

  aFileName = Split(sFile, "\")
  sFileName = (aFileName(Ubound(aFileName)))
  
  'listfiles
  sDupe = False
  For Each sFileNameInZip In oZipApp.NameSpace(sZipFile).items
    If LCase(sFileName) = LCase(sFileNameInZip) Then
      sDupe = True
      Exit For
    End If
  Next
  
  Set oZipFile = oZipApp.NameSpace(sZipFile)
  If Not sDupe Then ' do not process duplicate files
    oZipFile.Copyhere sFile

    'Keep script waiting until Compressing is done
    On Error Resume Next
    sLoop = 0
    Do While oZipFile.Items.Count <= sZipFileCount
      Call Pause(0.1)
      sLoop = sLoop + 1
    Loop
    On Error GoTo 0
  End If
  Set oZipFSO = Nothing
  Set oZipApp = Nothing
  Set oZipFile = Nothing
End Function

Sub NewZip(sNewZip, sOverwrite)
  'This script is provided under the Creative Commons license located
  'at http://creativecommons.org/licenses/by-nc/2.5/ . It may not
  'be used for commercial purposes with out the expressed written consent
  'of NateRice.com

  Dim oNewZipFSO
  Dim oNewZipFile
  
  Set oNewZipFSO = CreateObject("Scripting.FileSystemObject")
  Set oNewZipFile = oNewZipFSO.CreateTextFile(sNewZip, sOverwrite)
    
  oNewZipFile.Write Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, 0)
  
  oNewZipFile.Close
  Set oNewZipFSO = Nothing
  Set oNewZipFile = Nothing

  Call Pause(0.2)
End Sub
