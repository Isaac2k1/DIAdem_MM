'-------------------------------------------------------------------------------
' Name         : ---
' Copyright(c) : National Instruments Ireland Resources Limited
' Author       : NI
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.

If Not ItemInfoGet("sPathDocuments") Then
  Call GlobalDim("sPathDocuments")
  Call GlobalDim("sPathData")
End If

If Not ItemInfoGet("sTCursorChnName") Then
  Call GlobalDim("sTCursorChnName(3)")
End If

Dim sChnName, GroupIndex

sPathDocuments = AutoActPath
sPathData   = AutoActPath & "..\Data\"


'Call DATADELALL(1)                      '... HEADERDEL
'Call DATAFILELOAD(sPathData & "Example_Data.TDM","","") '... DATAFILENAME,FILEIMPORTFILTER,IMPORTACTION

if groupindexget("__TangentCursor__")<>0 then
  call groupdel(groupindexget("__TangentCursor__"))
end if

GroupIndex = GroupIndexGet(GroupCreate("__TangentCursor__"))
sChnName = ChnAlloc("X_TangentenCursor",2,,,,GroupIndex)
sTCursorChnName(0) = sChnName(0)
sChnName = ChnAlloc("Y_TangentenCursor",2,,,,GroupIndex)
sTCursorChnName(1) = sChnName(0)
sChnName = ChnAlloc("TangentenSlope",2,,,,GroupIndex)
sTCursorChnName(2) = sChnName(0)


ScriptCmdAdd(sPathDocuments & "TangentCursor_Event.VBS")

Call View.LoadLayout(sPathDocuments & "TangentCursor")
View.AutoRefresh = TRUE

'view.Sheets("Blatt 1").Areas("Area : 1").DisplayObj.Curves.item(1).XChannelName   ="[1]/[1]"
'view.Sheets("Blatt 1").Areas("Area : 1").DisplayObj.Curves.item(1).YChannelName   ="[1]/[2]"
'view.Sheets("Blatt 1").Areas("Area : 1").DisplayObj.Curves.item(2).XChannelName   =sTCursorChnName(0)
'view.Sheets("Blatt 1").Areas("Area : 1").DisplayObj.Curves.item(2).YChannelName   =sTCursorChnName(1)
'view.Sheets("Blatt 1").Areas("Area : 2").DisplayObj.Columns.item(1).ChannelName   =sTCursorChnName(0)
'view.Sheets("Blatt 1").Areas("Area : 2").DisplayObj.Columns.item(2).ChannelName   =sTCursorChnName(1)
View.Events.OnCursorChanged = "TangentCursor"
'view.Sheets("Blatt 1").Cursor.P1 = ChnLength("[1]/[1]") / 3 'Set cursor to the end of the first third of the x-axis
'view.Sheets("Blatt 1").Cursor.P2 = ChnLength("[1]/[1]") / 2 'Set cursor at the halfway point of the x-axis

view.Sheets("Blatt 1").Areas("Area : 1").DisplayObj.DoubleBuffered = true

WndShow "View"



