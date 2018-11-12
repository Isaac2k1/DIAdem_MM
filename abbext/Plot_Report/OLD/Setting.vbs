Class Setting
	private name
	private factor
	private off
	private offstep
	private scal
	
	Public Function setName(vname)
		name = vname
	End Function
	
	Public Function setFactor(vfactor)
		factor = vfactor
	End Function
	
	Public Function setOff(voff)
		off = voff
	End Function
	
	Public Function setOffstep(voffstep)
		offstep = voffstep
	End Function
	
	Public Function setScal(vscal)
		scal = vscal
	End Function
	
	Public Function getName()
		getName = name
	End Function
	
	Public Function getFactor()
		getFactor = factor
	End Function
	
	Public Function getOff()
		getOff = off
	End Function
	
	Public Function getOffstep()
		getOffstep = offstep
	End Function
	
	Public Function getScal()
		getScal = scal
	End Function
	
End Class