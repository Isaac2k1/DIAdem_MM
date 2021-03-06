'Contains the data of one time segment on the whole raster
Class Segment
  
	private unitFactor     	'time per unit e.g. 50ms/[]
	private unitAmount  	'amount of units --> [x]
	private duration   	
  
	Public Sub Class_Initialize
		unitFactor = 0
		unitAmount = 0
		duration = 0
	End Sub
  
	Private Function calculateDuration()
		duration = unitFactor * unitAmount
	End Function
	
	Public Function setUnitFactor(VunitFactor)
		unitFactor = VunitFactor
		calculateDuration()
	End Function
	
	Public Function setUnitAmount(VunitAmount)
		unitAmount = VunitAmount
		calculateDuration()
	End Function
	
	Public Function getUnitFactor()
		getUnitFactor = unitFactor
	End Function
	
	Public Function getUnitAmount()
		getUnitAmount = unitAmount 
	End Function
	
	Public Function getDuration()
		getDuration = duration 
	End Function
  
End Class