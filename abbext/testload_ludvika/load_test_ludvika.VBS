'-------------------------------------------------------------------------------
'-- VBS-Script-Datei
'-- Neu erstellt am 29.10.2008 10:47:58
'-- Ersteller: 
'-- Kommentar: 
'-------------------------------------------------------------------------------
Option Explicit  'Erzwingt die explizite Deklaration aller Variablen in einem Skript.

load_ludvika

sub load_ludvika

dim file_name, serial_nr,shot_group_index, prop_name, prop_value

serial_nr=inputbox("Please enter Serial-Number from Ludvika:","Enter Serial-Number")
if serial_nr = "" then exit sub

'serial_nr="00699934"

file_name=serial_nr&".001"

If not FileFilterExist("Nicolet_Team") Then
  Call msgbox ("Please install the data Plugin `Nicolet_Team` from NI-Website!")
'  exit sub
'else msgbox "continue"
end if

'datadelall

'Connect to DataFinder
          Dim MyDataFinder, AdvancedQuery,Element,count, added_groups,group_nr
          added_groups=0
          Set MyDataFinder = Navigator.ConnectDataFinder(Navigator.Display.CurrDataProvider.Name)
          'Define the query type (advanced)
          Set AdvancedQuery=Navigator.CreateQuery(eAdvancedQuery)
          'Define the result type (file)
          AdvancedQuery.ReturnType=eSearchFile
          
          Call UIAutoRefreshSet(True)
          
                          
                   
              Call AdvancedQuery.Conditions.RemoveAll()          
            'Search for  Data File 
            Call AdvancedQuery.Conditions.Add(eSearchFile,"fileName","=",file_name)
            Call AdvancedQuery.Conditions.Add(eSearchFile,"fullpath","<>","*RECINFO*")
             Call MyDataFinder.Search(AdvancedQuery)
             
            If(MyDataFinder.Results.Count = 0) Then
              msgbox "No test data found for "&file_name
              exit sub
            Elseif(MyDataFinder.Results.Count > 0) Then 'Load single or multiple rbd Files
             For Each Element In MyDataFinder.Results
              'call msgbox (Element)
              Call Navigator.LoadData(Element)
              added_groups=added_groups+1
             next
            end if


          for group_nr=groupcount-added_groups+1 to groupcount
              call groupdefaultset(group_nr)
              for count=cnoxget(group_nr,1) to cnoxget(group_nr,groupchncount(group_nr))
             '   msgbox count
                if cn(count)<>"Time" then
                  call chntowfchn("Time",count,False)
                end if   
              next
              call chndel("Time")
        next              
        
        if added_groups>1 then
          for group_nr=groupcount-added_groups+2 to groupcount
              call groupdefaultset(group_nr)
              for count=cnoxget(group_nr,1) to cnoxget(group_nr,groupchncount(group_nr))
                call chncopyext(count,groupcount-added_groups+1,0)
              next
          'call groupdel(group_nr)
          next              
        
        for group_nr=groupcount to groupcount-added_groups+2 step -1
          call groupdel(group_nr)
        next
        end if
        
        'generate properties-group


'load properties from Ascii file
call asciiconfigload(autoactpath&"ludvika_properties.stp")
call asciiassocset("*.dat",autoactpath&"ludvika_properties.stp")
        
        
      Set MyDataFinder = Navigator.ConnectDataFinder(Navigator.Display.CurrDataProvider.Name)
          'Define the query type (advanced)
          Set AdvancedQuery=Navigator.CreateQuery(eAdvancedQuery)
          'Define the result type (file)
          AdvancedQuery.ReturnType=eSearchFile
          
          Call UIAutoRefreshSet(True)
          
                          
                   
              Call AdvancedQuery.Conditions.RemoveAll()          
            'Search for  Data File 
            Call AdvancedQuery.Conditions.Add(eSearchFile,"fileName","=","*.dat")
            Call AdvancedQuery.Conditions.Add(eSearchFile,"fullpath","=","*"&serial_nr&"*")
             Call MyDataFinder.Search(AdvancedQuery)
             
            If(MyDataFinder.Results.Count = 0) Then
              msgbox "No test data found for "&serial_nr
              exit sub
            Elseif(MyDataFinder.Results.Count = 1) Then 'Load single or multiple rbd Files
             For Each Element In MyDataFinder.Results
              'call msgbox (Element)
              'CmdNoDialogDisp=True
              Call Navigator.LoadData(Element)
              'call asciiload(Element,0)
              next
            end if

      Groupname(groupdefaultget)="Properties"
      
      
      
        'write properties to group with shot data
shot_group_index=groupindexget("Properties")-1

for count=1 to cl("Properties/[1]")
  prop_name=cht(count,"Properties/[1]")
  if cht(count,"Properties/[4]")<>"" then prop_name=prop_name&"__"&cht(count,"Properties/[4]")
  prop_value=cht(count,"Properties/[3]")
prop_name=grouppropcreate(shot_group_index,prop_name,DataTypeString)
call grouppropset(shot_group_index,prop_name,prop_value)
next
call groupdel(groupindexget("Properties"))

        
        
        
        
        
        
        
end sub