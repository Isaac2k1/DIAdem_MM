'-------------------------------------------------------------------------------
'-- VBS-Script-Datei
'-- Neu erstellt am 06.05.2011 09:40
'-- Ersteller: M. Averes
'-- Kommentar: Entstand aus testload_ludvika
'-------------------------------------------------------------------------------
Option Explicit  'Erzwingt die explizite Deklaration aller Variablen in einem Skript.

load_kema

sub load_kema
dim file_name, serial_nr,shot_group_index, prop_name, prop_value

serial_nr=inputbox("Please enter Serial-Number from KEMA (e.g. 264):","Enter Serial-Number")
if serial_nr = "" then exit sub
'serial_nr="264"
file_name=serial_nr&".sid"

If not FileFilterExist("KEMA") Then
  Call msgbox ("Please install the data Plugin `KEMA` from C:\DIAdem\plugins\kema!")
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
  next 'Element
end if

'rename groupname
Data.Root.ChannelGroups(groupcount).Name = "Kema-" & serial_nr

end sub