'-------------------------------------------------------------------------------
'-- Authors: Jonas Schwammberger 
'-- Version: 1.9
'
'-- Purpose: Collection of Functions to Clear Travel Spikes.
'-- History:
'     1.01 Alpha Bugfixing 
'     1.02 Alpha Bugfixing
'     1.03 Alpha Hotfix for processing with a 6x Polynom
'     1.04 Alpha Hotfix and Bugfix for processing with a 6x Polynom
'     1.05 Alpha Polynom can have any level.
'     1.06 Alpha developing Smoothtravelcurve
'     1.07 Alpha developing Smoothtravelcurve 
'-------------------------------------------------------------------------------
Option Explicit  'Erzwingt die explizite Deklaration aller Variablen in einem Script.

GlobalDim("globalFactor")      'these global variables are used in the GUI.
GlobalDim("globalResponse")
GlobalDim("globalSkipAll")
'--------------------------------------------------------------------------------
'summary: clears travel spikes
'         I have no Idea what this algorythm is actually calculating, i copied
'         the code from the old Excel sheet. m Frederik Lundqvist might be the one' 
'         who could understand it.
'parameter travelCurve:
'parameter factor:
'output: none
sub ClearTravelSpikes(travelCurve,travelCurveOriginal,factor,Timestep)
	Dim Percentage
	Dim limitDer
  
	'check factor
	if factor < 0 then
		factor = 0
	elseif factor > 10 then
		factor = 10
	end if
  
  'back up travel curve
  
	percentage = (10 - factor) / 10
  LimitDer = FindMaxDerivate(travelCurve) * Percentage * 1.1
  
  'smooth travel curve
  Dim i
	Dim a, b, SumX, SumY, SumXY, SumX2
	Dim Iterations
	Dim RowCounter 'maybe not needed
	Dim Counter
	Dim AbsDiff
  
  For i = 2 to ChnLength(travelCurve)
    'We calculate the derivative of 10 points spread over the last 2 ms.
    a = CHD(i-1,travelCurve)
    b = 0
  
    If i > 0.003 / Timestep + 10 Then
      SumX = 0
      SumY = 0
      SumXY = 0
      SumX2 = 0
      Iterations = 0
    
      'counter is +20, because the script was copied from the excel sheet where the columns
      'start at 20
      For Counter = i - 0.003 / Timestep To i - 1 Step 0.003 / Timestep / 3
        SumX = SumX + Counter + 19
        SumY = SumY + CHD(Counter,travelCurve)
        SumXY = SumXY + (Counter + 19) * CHD(counter,travelCurve)
        SumX2 = SumX2 + (Counter + 19) ^ 2
        Iterations = Iterations + 1
      Next 
      
      a = (SumY * SumX2 - SumX * SumXY) / (Iterations * SumX2 - SumX * SumX)
      b = (Iterations * SumXY - SumX * SumY) / (Iterations * SumX2 - SumX * SumX)
      
      AbsDiff = abs((CHD(i,travelCurveOriginal) - CHD(i-1,travelCurve)) / Timestep - b)
    
      If AbsDiff < LimitDer / Timestep Then
        CHD(i,travelCurve) = CHD(i,travelCurveOriginal)
      Else
        CHD(i,travelCurve) = a + b * i
      End If
    
    End If
  
  Next

end sub
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
'summary: Find maximum Derivate in a channel
'parameter channel:
'output: maximum Derivate in a channel
function FindMaxDerivate(channel)
	Dim RowCounter
	Dim MaxDer
	Dim AbsDiff

	MaxDer = 0

	for RowCounter = 21 to chnLength(channel)
		AbsDiff =  abs(CHD(RowCounter,channel) - CHD(RowCounter-1,channel))
		If AbsDiff > MaxDer Then
			MaxDer = AbsDiff
		End If
	next

	FindMaxDerivate = MaxDer

end function
'--------------------------------------------------------------------------------