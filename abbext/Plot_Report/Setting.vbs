'stores setting of the current channel which will be used by the next channel group
Class Setting
	private name
	private off
	private offstep
	private scal
	
	Public Function setName(vname)
		name = vname
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