'-------------------------------------------------------------------------------
' Name         : ---
' Copyright(c) : National Instruments Ireland Resources Limited
' Author       : NI
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.

GlobalDim("oVIEWTCursorM")

Set oVIEWTCursorM = new VIEW_Tangent_Cursor

Sub TangentCursor(oCursor)

  if (cno(sTCursorChnName(0)) <= 0 ) or _
     (cno(sTCursorChnName(1)) <= 0 ) then
    View.Events.OnCursorChanged = ""
    exit sub
  end if
  ' Calculates tangent
  Call oVIEWTCursorM.OnCursorHasChanged(oCursor)

End Sub

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
Class VIEW_Tangent_Cursor
  '--------------------------------------------------
  ' Definition of properties
  '--------------------------------------------------
  Public oCommanderDialog
  Public sgDlgHomeDirM
  '--------------------------------------------------
  '@
  '-------------------------------------------------- 
  Private Sub Class_Initialize   
  End Sub 
  '--------------------------------------------------
  '@
  '--------------------------------------------------
  Private Sub Class_Terminate 
  End Sub
  '--------------------------------------------------
  ' 
  '--------------------------------------------------
  Public Function OnCursorHasChanged(oCursor)
    Dim   XMin,XMax,YMin,YMax,ChnNoX,ChnNoY,ChnSlope,sgIntersectionMask
    Dim   dXS(3),dYS(3),ICount
    Dim XP1, XP2, YP1, YP2, XS, YS, M1
    Dim XBegin, XEnd, YBegin, YEnd
    If ( oCursor.X1 <= oCursor.X2 ) Then
      XP1 = oCursor.X1
      YP1 = oCursor.Y1 
      XP2 = oCursor.X2
      YP2 = oCursor.Y2 
    Else
      XP1 = oCursor.X2
      YP1 = oCursor.Y2 
      XP2 = oCursor.X1
      YP2 = oCursor.Y1 
    End If 

    XBegin = oCursor.Sheet.Areas(1).DisplayObj.XBegin
    XEnd   = oCursor.Sheet.Areas(1).DisplayObj.XEnd
    YBegin = oCursor.Sheet.Areas(1).DisplayObj.YBegin
    YEnd   = oCursor.Sheet.Areas(1).DisplayObj.YEnd

    sgIntersectionMask = ""
    ' Y-axis
    If ( Intersection(XBegin,YBegin,0.0,YEnd-YBegin,XP1,YP1,XP2-XP1,YP2-YP1,M1,XS,YS) ) Then
      If ( 0.<=M1 And 1.>=M1 ) Then 
        sgIntersectionMask = sgIntersectionMask & "+"
        dXS(ICount) = XS
        dYS(ICount) = YS
        ICount      = ICount + 1
      Else
        sgIntersectionMask = sgIntersectionMask & "-"
      End If 
    End If 
    ' Y-axis 
    If ( Intersection(XEnd,YBegin,0.0,YEnd-YBegin,XP1,YP1,XP2-XP1,YP2-YP1,M1,XS,YS) ) Then
      If ( 0.<=M1 And 1.>=M1 ) Then 
        sgIntersectionMask = sgIntersectionMask & "+"
        dXS(ICount) = XS
        dYS(ICount) = YS
        ICount      = ICount + 1
      Else
        sgIntersectionMask = sgIntersectionMask & "-"
      End If 

    End If 
    ' X-axis 
    If ( Intersection(XBegin,YBegin,XEnd-XBegin,0.0,XP1,YP1,XP2-XP1,YP2-YP1,M1,XS,YS) ) Then
      If ( 0.<=M1 And 1.>=M1 ) Then 
        sgIntersectionMask = sgIntersectionMask & "+"
        dXS(ICount) = XS
        dYS(ICount) = YS
        ICount      = ICount + 1
      Else
        sgIntersectionMask = sgIntersectionMask & "-"
      End If 

    End If 
    ' X-axis
    If ( Intersection(XBegin,YEnd,XEnd-XBegin,0.0,XP1,YP1,XP2-XP1,YP2-YP1,M1,XS,YS) ) Then
      If ( 0.<=M1 And 1.>=M1 ) Then 
        sgIntersectionMask = sgIntersectionMask & "+"
        dXS(ICount) = XS
        dYS(ICount) = YS
        ICount      = ICount + 1
      Else
        sgIntersectionMask = sgIntersectionMask & "-"
      End If 
    End If 
    XP1 = dXS(0)
    XP2 = dXS(1)

    YP1 = dYS(0)
    YP2 = dYS(1)

    ChnNoX = CNo(sTCursorChnName(0))
    ChnNoY = CNo(sTCursorChnName(1))
    ChnSlope = CNo(sTCursorChnName(2))
    CHDX(1,ChnNoX) = XP1
    CHD(2,ChnNoX) = XP2
    CHDX(1,ChnNoY) = YP1
    CHD(2,ChnNoY) = YP2
    'msgbox chnslope&" "&XP1&" "&XP2
    if XP2<>XP1 then  CHD(1,ChnSlope) = (YP2-YP1)/(XP2-XP1)
    'view.Refresh()

  End Function 

  '-------------------------------------------------
  ' Calculates the intersection point of two straight line:
  ' The function returns True if the two lines intersect
  '-------------------------------------------------
  Function  Intersection(X10,Y10,DX1,DY1,X20,Y20,DX2,DY2,M1,XS,YS)
    Dim   Q1,Q2
    Intersection = False
    M1 = 0.
    XS = 0.
    YS = 0.
    If ( DX1=0 And DY1=0 ) Then Exit Function
    If ( DX2=0 And DY2=0 ) Then Exit Function
    Q1 = DX1*DY2 - DY1*DX2
    ' Parallel straight line? 
    If ( 0 = Q1 ) Then Exit Function
    Q2 = (Y10-Y20)*DX2 - (X10-X20)*DY2
    M1 = Q2/Q1
    XS = X10 + M1 * DX1
    YS = Y10 + M1 * DY1
    Intersection = True
  End Function

End Class

