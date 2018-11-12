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
' >>> The class allows to use the "AdvancedQuery" of DIAdem with Conditions for the search
' >>> All scripts and classes are located in the same folder - Load Test
' >>> The detailed description of all the functions is also located in the Load Test folder
'
' >>> DataFinder is a global variable - created in the Load Test GUI
'---------------------------------------------------------------------------------------------------------------------------
'
'
'
Class QuerySearch
  
	private oAdvancedQuery
	private conditionNumber
	
	' >>> Object from DIAdem to create the Advanced Search Query
	Public Sub Class_Initialize
		Set oAdvancedQuery = DataFinder.CreateQuery(eAdvancedQuery)
	End Sub
	
	Public function intitSearch(returnType) 'reset Variables and remove data from previous search
		Call oAdvancedQuery.Conditions.RemoveAll
		oAdvancedQuery.ReturnType = returnType		'eSearchFile / eSearchChannelGroup / eSearchChannel
		conditionNumber = 1
	End function
	
	Public Function addCondition(lCondition) 'receive list with Conditions Objects
		dim condition, queryLogic
		
		For each condition in lCondition
			Call oAdvancedQuery.Conditions.Add( condition.getType(), condition.getProperty(),condition.getOperator(), condition.getValue())
			If conditionNumber = 1 Then 'create Query Logic --> 3 Conditions = "C1 AND C2 AND C..."
				queryLogic = "C1"
			Else
				queryLogic = queryLogic &" AND C"&conditionNumber
			End If
			conditionNumber = conditionNumber + 1
		Next
		oAdvancedQuery.Conditions.Logic=(queryLogic)
    
	End Function
	
	Public Function searchFiles()
		Call DataFinder.Search(oAdvancedQuery)
		Call DataFinder.Results.Sort()
	End Function
	
	Public Function loadResults()
		dim oElement
		
		For each oElement in DataFinder.Results
			If Data.Root.ChannelGroups.Exists(oElement.Name) Then
				Data.Root.ChannelGroups.Remove(oElement.Name)
			End If
			Call Navigator.LoadData(oElement)
		Next
	End Function
	
End Class
