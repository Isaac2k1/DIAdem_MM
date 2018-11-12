Class DetailSetting
	
	private name
	private timeUnit
	private startTime
	private duration
	private minValue
	private maxValue
	
	Public Function setName(Vname)
		name = Vname
	End Function
	
	Public Function setTimeUnit(VtimeUnit)
		timeUnit = VtimeUnit
	End Function
	
	Public Function setStartTime(VstartTime)
		startTime = VstartTime
	End Function
	
	Public Function setDuration(Vduration)
		duration = Vduration
	End Function
	
	Public Function setMinValue(VminValue)
		minValue = VminValue
	End Function
	
	Public Function setMaxValue(VmaxValue)
		maxValue = VmaxValue
	End Function
	
	Public Function getName()
		getName = name
	End Function
	
	Public Function getTimeUnit()
		getTimeUnit = timeUnit
	End Function
	
	Public Function getStartTime()
		getStartTime = startTime
	End Function
	
	Public Function getDuration()
		getDuration = duration
	End Function
	
	Public Function getMinValue()
		getMinValue = minValue
	End Function
	
	Public Function getMaxValue()
		getMaxValue = maxValue
	End Function
	
End Class