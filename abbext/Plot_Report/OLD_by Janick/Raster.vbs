Class Raster
  
	private factor     	'ms per raster  
	private amount  	'raster
	private duration   	
  
	Public Sub Class_Initialize
		factor = 0
		amount = 0
		duration = 0
	End Sub
  
	Private Function calculateMissing()
		duration = factor * amount
	End Function
	
	Public Function setFactor(factorV)
		factor = factorV
		calculateMissing()
	End Function
	
	Public Function setAmount(amountV)
		amount = amountV
		calculateMissing()
	End Function
	
	Public Function getFactor()
		getFactor = factor
	End Function
	
	Public Function getAmount()
		getAmount = amount 
	End Function
	
	Public Function getDuration()
		getDuration = duration 
	End Function
  
End Class