'-------------------------------------------------------------------------------
'-- Authors: Jonas Schwammberger 
'-- Version: 1.9.1
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
'summary: clears travel spikes.
'parameter travelCurve:
'parameter factor:
'output: none
sub ClearTravelSpikes(travelCurve,travelCurveOriginal,factor,Timestep)

dim i
For i = 2 to chnlength(travelCurveOriginal)-2
  if ((abs(CHD(i, travelCurveOriginal) - CHD(i-1 , travelCurveOriginal)) > factor * 0.001) and (abs(CHD(i, travelCurveOriginal) - CHD(i+1 , travelCurveOriginal)) > factor *0.001)) then
    dim j, sum
    sum = 0
    for j =1 to 5
      if (j >= i) or (j + i)>= chnlength(travelCurve) then
        exit for
      end if
      sum = sum + CHD(i-j, travelCurve) + CHD(i+j, travelCurveOriginal)
    next
    CHD(i, travelCurve)= sum / 2 /(j - 1)
  end if
next

end sub
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
'summary: Find maximum Derivate in a channel
'parameter channel:
'output: maximum Derivate in a channel
'function FindMaxDerivate(channel)
'	Dim RowCounter
'	Dim MaxDer
'	Dim AbsDiff
'
'	MaxDer = 0
'
'	for RowCounter = 21 to chnLength(channel)
'		AbsDiff =  abs(CHD(RowCounter,channel) - CHD(RowCounter-1,channel))
'		If AbsDiff > MaxDer Then
'			MaxDer = AbsDiff
'		End If
'	next
'
'	FindMaxDerivate = MaxDer
'
'end function
'--------------------------------------------------------------------------------
