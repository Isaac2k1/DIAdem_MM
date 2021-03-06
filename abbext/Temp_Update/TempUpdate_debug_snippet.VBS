'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2018-07-04 16:14:03
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
  dim oFSO, oFile, currLine, dataArray
  Dim addShots, addSerial, addTestNumbers
dim PATH, FOLDER, FILE
dim dataPool, serialNumber, count, maxshot, intCount
dim Start, Ende

PATH = "\\CH-S-0008257\acqui\@tmp\"
FOLDER = "\DATA\CHANNEL\"
FILE = "campaign.ini"
dataPool = "CH-W-PTHX109"
 maxshot = 0
  serialNumber = "#"
  
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  If NOT oFSO.FileExists(PATH & dataPool & "/"  & FILE) Then
    MsgBox("No usable data aviable")
  '  Exit Sub
  End If
  
  
  ' set serial nr. and finds out the highest shot
  Set oFile = oFSO.OpenTextFile(PATH & dataPool & "/" & FILE)
  addSerial = false

      
      
  Do Until oFile.AtEndOfStream 
    currLine = oFile.ReadLine()
    msgbox(currLine)
    If Trim(currLine) = "" then
    msgbox("empty line")
      'empty line
    ElseIf currLine = "[REBADAS_NO]" Then
        addSerial = true
    ElseIf addSerial then
      
      dataArray = Split(currLine, "=")
      If maxshot < Cint(dataArray(0)) then
        maxshot = Cint(dataArray(0))
      msgbox(Cint(dataArray(0)))
      End If
      If (Left(currLine,1)= 0) then
        serialNumber = Replace(dataArray(1), Chr(34), "")
      End If
  
    End if
  loop

      