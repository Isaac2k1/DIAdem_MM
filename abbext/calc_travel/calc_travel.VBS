'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2011-12-09 09:10:35
'-- Author: 
'-- Comment: Calculates a Curve from degree to a metric format
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
'-------------------------------------------------------------------------------

GlobalDim "ChnString"         'string with all selected channels
GlobalDim "factor"            'scale factor to get m or mm
GlobalDim "unit"              'metric unit
Dim i                         'index used by the for-loop
dim fullName                  'name of the new channel

Dim H : H = 6                 'Polynom...
Globaldim "dPolynom"          '...ugly, but it works
Globalredim "dPolynom("&H&")"

If SUDDlgShow("Dlg1",AutoActPath & "calc_travel.sud") = "IDOk" Then 'if OK-Button was pressed

  For i = 0 to UBound(ChnString) - 1    'for each selected channel

    fullName = split(ChnString(i),"/")(0) & "/" & split(ChnString(i),"/")(1)&"_[" & unit & "] "     

    call ChnCopy(ChnString(i),fullName)                     'copy everything from the old to the new channel
    call ChnOffset(fullName,fullName,0,"max. value offset") 'shifts the entire curve to negative y-values
    call CalculateTravel(fullname)                          'calculates the new curve
    call chnpropset(fullname,"unit_string",unit)            'set the correct unit to the curve properties
  Next

  msgbox("Done!")

End if

'--------------------------------------------------------------------------------
'summary: calculates the Travel curve from degree to meter and invert the curve
'parameter fullname: channel group and name as string
'needs additional the global varables dPolynom and factor
'the input curve needs to be completly negative, otherwise the calculation doesn't work correctly
Sub CalculateTravel(fullname)
  dim travelInMm                            'new value of point
  dim degree                                'old value of point
  dim i, j                                  'indices for for-loops
  dim polyDim                               'dimension of polynom
  
  for i = UBound(dPolynom) to 0 step -1     'check how much calculation steps are needed for each datapoint
    if dPolynom(i) <> 0 then                'seach the highes non zero level of the polynom
      polyDim = i
      exit for
    end if      
  next

  for i = 1 to chnLength(fullname)          'calculates new value for each datapoint
    degree = -CHD(i,fullname)               'old datapoint is degree, mirrored at the x-Axis
    travelInMm = dPolynom(0)                'startvalue is A_0
  
    for j = 1 to polyDim                    'for each coefficient from polynom
      travelInMm = travelInMm + dPolynom(j) * degree^j
    next
    
    CHD(i,fullname) = travelInMm/factor     'new Value is scaled by factor
  next
 
end sub
'--------------------------------------------------------------------------------