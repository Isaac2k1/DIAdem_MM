' ----------------------------------------------------------------------------------------
' Name         : ---
' Author       : National Instruments Ireland Resources Limited
' Comment      : ---                                                  
' ----------------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.

If Not ItemInfoGet("sPathDocuments") Then
  Call GlobalDim("sPathDocuments")
  Call GlobalDim("sPathData")
End If

If Not ItemInfoGet("oTCursorChnName") Then
  Call GlobalDim("oTCursorChnName(2)")
End If

Dim oGroup, oGroupChns

sPathDocuments = AutoActPath
sPathData   = AutoActPath & "..\Data\"


Call Data.Root.Clear

Call DATAFILELOAD(sPathData & "Example_Data.TDM","","") '... DATAFILENAME,FILEIMPORTFILTER,IMPORTACTION

set oGroup = Data.Root.ChannelGroups.Add("__TangentCursor__")
call oGroup.Activate
set oGroupChns = oGroup.Channels
set oTCursorChnName(0) = oGroupChns.Add("X_TangentCursor",DataTypeChnFloat64)
set oTCursorChnName(1) = oGroupChns.Add("Y_TangentCursor",DataTypeChnFloat64)

ScriptCmdAdd(sPathDocuments & "TangentCursor_Event.VBS")

Call View.LoadLayout(sPathDocuments & "TangentCursor")

view.Sheets("Sheet 1").Areas("Area : 1").DisplayObj.Curves.item(1).XChannelName   ="[1]/[1]"
view.Sheets("Sheet 1").Areas("Area : 1").DisplayObj.Curves.item(1).YChannelName   ="[1]/[2]"
view.Sheets("Sheet 1").Areas("Area : 1").DisplayObj.Curves.item(2).XChannelName   =oTCursorChnName(0).Name
view.Sheets("Sheet 1").Areas("Area : 1").DisplayObj.Curves.item(2).YChannelName   =oTCursorChnName(1).Name
view.Sheets("Sheet 1").Areas("Area : 2").DisplayObj.Columns.item(1).ChannelName   =oTCursorChnName(0).Name
view.Sheets("Sheet 1").Areas("Area : 2").DisplayObj.Columns.item(2).ChannelName   =oTCursorChnName(1).Name

call AddUserCommandToEvent("View.Events.OnCursorChanged", "TangentCursor")

dim oTempChn : set oTempChn = Data.GetChannel("[1]/[1]")
view.Sheets("Sheet 1").Cursor.P1 = oTempChn.Size / 3 'Set cursor to the end of the first third of the x-axis
view.Sheets("Sheet 1").Cursor.P2 = oTempChn.Size / 2 'Set cursor at the halfway point of the x-axis

view.Sheets("Sheet 1").Areas("Area : 1").DisplayObj.DoubleBuffered = true

WndShow "View"



