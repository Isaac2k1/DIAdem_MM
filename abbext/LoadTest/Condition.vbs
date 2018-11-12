'------------------------ Author -------------------------------------------------------------------------------------------
''   Janick Schmid
'
'------------------------ Creation Date ------------------------------------------------------------------------------------
''   2014-03 to 2015-08
'
'------------------------ Description --------------------------------------------------------------------------------------
'
' >>> This class (Condition.vbs) is part of the "Load Test" program to be used in Diadem
' >>> This class is called by Load_Test.SUD
' >>> The objects of the class are used by class "QuerySearch" to store condition parameter
' >>> All scripts and classes are located in the same folder - Load Test
' >>> The detailed description of all the functions is also located in the Load Test folder
'---------------------------------------------------------------------------------------------------------------------------
'
'
Class Condition
	private pType
	private pProperty
	private pValue
	private pOperator
	
	Public Function setType(Vtype)
		pType = Vtype
	End Function
	
	Public Function setProperty(Vproperty)
		pProperty = Vproperty
	End Function
	
	Public Function setValue(Vvalue)
		pValue =  Vvalue
	End Function
  
	Public Function setOperator(Voperator)
		pOperator = Voperator
	End Function
	
	Public Function getType()
		getType = pType
	End Function
	
	Public Function getValue()
		getValue = pValue
	End Function 
	
	Public Function getProperty()
		getProperty = pProperty
	End Function
  
	Public Function getOperator()
		getOperator = pOperator
	End Function
  
End Class
