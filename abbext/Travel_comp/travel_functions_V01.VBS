'-------------------------------------------------------------------------------
'-- VBS-Script-Datei
'-- Neu erstellt am 19.06.2008 16:39:23
'-- Ersteller: Mathias Knaak
'-- Kommentar: Funktionen für die Benutzung beim Wegvergleich
'   Last Update: 2008-07-29.11:23:50
'   Version: 1.3.0
'   Reviewed: 
'-------------------------------------------------------------------------------

' Description: Functions for usage in travel_compare.vbs

' History
' v1.0
' first complete and usable version
' v1.1
' v1.2
' as defined in IEC 62271-100, tolerance Band is only drawn when contacts are open. 
' tolerance test is only done in range with open contacts
' CS-Channel is scaled to 75% of maximum travel


'call create_tolerance_band ("ref",0,200,1)




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'load_channel 
'lädt die ausgewählten Kanäle mit den Einstellungen aus dem Dialog
'return value is boolean, true if channel correctly loaded
'
'Input: None 
'Output: True, if channels loaded successfully
'
function load_channel()

'set load_channel=TRUE, will be set to "FALSE" if an error occurs
load_channel=true

'if CS Signal is used, the user has to select 4 channels
if ((chn_ref="" or chn_comp="" or cs_chn_ref="" or cs_chn_comp="") and use_cs=1) then
  call msgbox ("Please select all channels.")
  load_channel=false
  exit function
end if

'user has to select two Travel channels
if ((chn_ref="" or chn_comp="") and use_cs=0) then
  call msgbox ("Please select both travel channels.")
  load_channel=false
  exit function
end if

' all selected input-channels have to be waveform-channels
' all channels have to be "equidist-X"
' if channels are loaded as "Load as it is", they are stored as numeric channels

if chnwfkey(chn_ref)=false then
  call msgbox("Reference channel is not an Waveform-Channel!"&Chr(13)&"Load Channels as equidist-X"&Chr(13)&"Please check groups and try again.",16)
  load_channel=false
  exit function
end if

if chnwfkey(chn_comp)=false then
  call msgbox("Compared channel is not an Waveform-Channel!"&Chr(13)&"Load Channels as equidist-X"&Chr(13)&"Please check groups and try again.",16)
  load_channel=false
  exit function
end if

if use_cs=1 then
    if chnwfkey(cs_chn_ref)=false then
      call msgbox("Reference CS-channel is not an Waveform-Channel!"&Chr(13)&"Load Channels as equidist-X"&Chr(13)&"Please check groups and try again.",16)
      load_channel=false
      exit function
    end if

    if chnwfkey(cs_chn_comp)=false then
      call msgbox("Compared CS-channel is not an Waveform-Channel!"&Chr(13)&"Load Channels as equidist-X"&Chr(13)&"Please check groups and try again.",16)
      load_channel=false
      exit function
    end if
end if

'delete existing Travel groups

if groupindexget("Travel_ref") <> 0 then groupdel(groupindexget("Travel_ref"))

if groupindexget("Travel_comp") <> 0 then groupdel(groupindexget("Travel_comp"))


'generate group "Travel_ref" for reference curve

  call Group_generate (chn_ref,"ref")   'Gruppe "Travel_ref" für die Referenzkurve erzeugen
  
    if clean_ref=1 then                 ' Daten bereinigen
      call clean ("ref",Val(clean_ref_limit))
    end if

    if remove_offset_ref = 1 then       'Offset entfernen
      call offset_remove("ref",offset_auto_ref,val(r_offset_man_val))
    end if

    if scale_ref=1 then                 ' skalieren  
      call scale_y("ref", Val(scale_val_ref))
    end if

  call backup_travel("ref")     ' Backup des Travel-Kanals erstellen

 
  call Group_generate(chn_comp,"comp")    'Gruppe "Travel_comp" für die Vergleichskurve erzeugen

   if clean_comp=1 then                   ' Daten bereinigen
      clean "comp",Val(clean_comp_limit)
    end if

    if remove_offset_comp = 1 then            'Offset entfernen
      offset_remove  "comp",offset_auto_comp,val(c_offset_man_val)
    end if

    if scale_comp=1 then                  ' skalieren  
      scale_y "comp", Val(scale_val_comp)
    end if

    backup_travel "comp"                  ' Backup des Travel-Kanals erstellen
    
    'copy Travel Channels for other plots
    Call Chncopy("Travel_comp/Travel_unshifted","Travel_comp/Travel_2")     'Travelkurve für den zweiten Plot erstellen
    Call Chncopy("Travel_comp/Travel_unshifted","Travel_comp/Travel_3")     'Travelkurve für den dritten Plot erstellen
end function
'''''''''''''''''''''''''''''''''''''''




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Group_generate
'Kanalgruppe erzeugen, Travelkanal kopieren, Eigenschaften aus *.rba übernehmen
'
'Input: Inputchannel, Name of Travelgroup "Travel_"&Name


Sub Group_generate(Inputchannel, Name)     

dim group_name,group_index,group_desc, desc_pos                    'Gruppenname, Gruppenindex

group_name = ChnPropGet(Inputchannel, "groupname")  
group_index= groupindexget(group_name)              

' bei korrekt geladenen Daten enthält diese Gruppe alle Eigenschaftes des Versuchs

'Kopie der Gruppe mit den Versuchseigenschaften erstellen, 
'Gruppe wird nach den aktuellen Gruppen eingereiht

'call groupcopy(group_index, 0)    

'remove first character and .rbe from group description
group_desc = grouppropget(group_index,"name")
desc_pos=len(group_desc)
group_desc = right(group_desc,desc_pos-1) 'remove first character
desc_pos=instr(group_desc,".rbe")
if desc_pos > 0 then group_desc = left(group_desc,desc_pos-1) 'remove ".rbe"
group_desc=replace(group_desc,"-","/") 
call grouppropset(group_index,"description",group_desc) 'Gruppenname in "desription" sichern

call Groupcreate("Travel_"&Name)
call Grouppropcopy(group_index,groupcount)

Call Chncopy(Inputchannel,"Travel_"&Name&"/Travel")       'Copy Data to new Channel "Travel"  
End Sub
'''''''''''''''''''''''''''''''''''''''




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'offset_remove
'Offset entfernen, automatisch oder manuell
' 
'Input: Name, auto or manuell [0|1], offset value

Sub offset_remove(Name, auto, offset_value) 'Channelgroup "Name", 

dim channelname
channelname="Travel_"&Name&"/Travel"

  select case auto

    case "1"      'Offset automatisch korrigieren, minimaler y-Wert wird auf null gesetzt

      Call ChnOffset(channelname, channelname ,0,"min. value offset")

    case "0"      'Offset manuell korrigieren, offset_value wird addiert

      Call ChnOffset(channelname, channelname ,offset_value,"free offset")

  end select

End Sub
'''''''''''''''''''''''''''''''''''''''




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'clean
'Entfernt Störungen aus dem Messignal
'Bildet die Ableitung des Messignals, bei Steigung grösser "threshold" wird das Messignal als fehlerhaft 
'angenommen und dort linear interpoliert
'
'Input: Name, threshold


sub clean(Name, threshold)    'Gruppenname, Grenzwert für Störentfernung

dim channelname,group_name, diff_chn, NV_chn,diff_len

group_name="Travel_"&Name
channelname=group_name&"/Travel"
T1=group_name&"/Travel_NV"
T2=channelname
T4=channelname&"_Diff_Y"

diff_chn=channelname&"_Diff_Y"
NV_chn="Travel_NV"

L1=threshold

' differentiate Travelchannel
Call ChnDifferentiate("",channelname,"",group_name&"/Travel_Diff_Y") 
'select channel group     
Call GroupDefaultSet(groupindexget(group_name))                            
' calculate absolute Values of differentiated channel
Call channel_absolute(group_name&"/Travel_Diff_Y") 

'create new channel 
Call Chnalloc(NV_chn)
'remove distorted values: if derivation is greater than threshold, Values will be replaced with NOVALUES
Call ChnCalculate("Ch(T1) = Ch(T2) + CTNV(Ch(T4)>L1) ")    
'generate time channel and rename channel (bug in DIAdem)
Call ChnFromWfXGen(group_name&"/Travel_Diff_Y",group_name&"/Travel_time") '... Y,E            
Call ChnPropValSet(group_name&"/noname","name","travel_time")                                
'interpolate NoValues
Call ChnNovHandle(group_name&"/Travel_time",group_name&"/Travel_NV","Interpolate","XY",1,0,0) 
'create Waveformchannel
Call ChnToWfChn(group_name&"/travel_time",group_name&"/Travel_NV",0)            
'copy to Travel-Channel
Call ChnCopy(group_name&"/Travel_NV",group_name&"/Travel")                      
'insert one "zero" for original channel length (at differentiation one value was "lost") 
Call ChnAreaInsert0(group_name&"/Travel", 1, 1)                                 


End Sub
'''''''''''''''''''''''''''''''''''''''




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'scale_y
'skaliert die Messwerte auf einen vorgegeben Maximalwert
'
'Input: Name, maximum Value

sub scale_y(name, Max_Value)    'Gruppenname, neuer Maximalwert

dim channelname,group_name,scalefactor
group_name="Travel_"&Name
channelname=group_name&"/Travel"

scalefactor= Max_Value / Cmax(channelname)

Call ChnLinScale(channelname,channelname,scalefactor,0) 'skalieren auf neuen Maximalwert

end sub
'''''''''''''''''''''''''''''''''''''''




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'backup_travel
'
'kopiert den Travel-Kanal in einen Backup-Kanal

Sub backup_travel(Name)     'Name der Kanalgruppe
dim group_name
group_name="Travel_"&Name          
Call Chncopy(group_name&"/Travel",group_name&"/Travel_unshifted")       'Copy Data to new Channel "Travel_unshifted"  
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'create_tolerance_band
'erzeugt zwei neue Kurven mit 5% Abstand zur Referenzkurve
'Verschiebung der Kurven mit "shift" möglich

sub create_tolerance_band(name, shift, max_way, show)   'Name der Kanalgruppe, Offset des Toleranzbandes(+/- 5%), Gesamthub, 
                                                        ' show=1 : Kanäle erzeugen    show=0 : Kanäle löschen
dim channelname,group_name, pos_tol, neg_tol
group_name="Travel_"&Name
channelname=group_name&"/Travel"


' Wert für die Verschiebung gegenüber dem Originalkanal bestimmen
pos_tol=(0.05*max_way)*(1+shift/5)
neg_tol=(0.05*max_way)*(-1+shift/5)

Call GroupDefaultSet(groupindexget(group_name))  

'Toleranzband löschen, falls es bereits besteht
If CNo(group_name&"/Travel_pos_tol") <> 0 then
  call chndel (group_name&"/Travel_pos_tol")
  call chndel (group_name&"/Travel_neg_tol")
end if

if show = 1 then        'wenn Kanäle erzeugt werden sollen, wird der Travel-Kanal kopiert und verschoben

Dim width_read, width_calc
width_read=0
  If scale_tol=0 Then
      width_calc=round(chnpropget("Travel_ref/travel","Maximum")-chnpropget("Travel_ref/travel","Minimum"),3)
      call Msgbox("Measured Travel width is: "&width_calc&chr(13)&"Please enter Travel width for tolerance band!")	
  End If


  Call ChnCopy(group_name&"/Travel",group_name&"/Travel_pos_tol")
  Call ChnCopy(group_name&"/Travel",group_name&"/Travel_neg_tol")

  Call ChnOffset(group_name&"/Travel_pos_tol", group_name&"/Travel_pos_tol" ,pos_tol,"free offset")
  Call ChnOffset(group_name&"/Travel_neg_tol", group_name&"/Travel_neg_tol" ,neg_tol,"free offset")


  'remove tolerance Band when contact is closed
  'to insert novalues, channels have to be converted into numeric channels and vice versa

  view.autorefresh=False 'do not refresh view

      if use_cs=1 then      'if CS-Signal is used
        'convert channels to numeric channels
        dim time_chn_1, time_chn_2
        time_chn_1= ChnFromWfXGen(group_name&"/Travel_pos_tol","Travel_pos_time") '... Y,E 
        time_chn_2= ChnFromWfXGen(group_name&"/Travel_neg_tol",group_name&"/Travel_neg_time") '... Y,E 
        call wfchntochn(group_name&"/Travel_pos_tol")
        call wfchntochn(group_name&"/Travel_neg_tol")
        
    
          dim length,count
          'count values in the channel
          length=CL(group_name&"/CS")
          count=1
          'check each value in CS-Channel; if contacts are closed, remove tolerance band
          for count=1 to length 
            if chd(count,group_name&"/CS")>50 then 
              chd(count,group_name&"/Travel_pos_tol")=NOVALUE
              chd(count,group_name&"/Travel_neg_tol")=NOVALUE
            end if
          next
        'convert numeric channels back to Waveform-Channels
         call chntowfchn(time_chn_1,group_name&"/Travel_pos_tol",1)
         call chntowfchn(time_chn_2,group_name&"/Travel_neg_tol",1)
         'call chntowfchn((cno(group_name&"/Travel_pos_tol")+2),group_name&"/Travel_pos_tol",1)
         'call chntowfchn((cno(group_name&"/Travel_neg_tol")+2),group_name&"/Travel_neg_tol",1)
      
      end if  ' if use_cs=1 then 

  view.autorefresh=True 'refresh view
  call view.refresh()

end if  'if show = 1 then   

end sub
'''''''''''''''''''''''''''''''''''''''




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'shift_x
'verschiebt einen Waveform-Kanal in x-Richtung um einen vorgegebenen Wert


sub shift_x(Name,shift_value)     'kompletter Name, Wert, um den verschoben werden soll
dim xValue,step_value
xValue = Chnpropvalget(Name,"wf_start_offset")  'alten x-Offset einlesen
xValue =xValue + shift_value                    'neuen x-Offset berechnen   
step_value=Chnpropvalget(Name,"wf_increment")   'x-Schrittweite einlesen
Call Chnwfpropset(Name,"travel_time",,xValue,step_value)  'neue Waveform-Parameter schreiben
end sub
'''''''''''''''''''''''''''''''''''''''




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'find_cs
'findet Zeitpunkte der Kontakttrennung im CS-Kanal und Richtung der Bewegung
'
'erzeugt folgende Ergebniskanäle:
'CS_switch_number: fortlaufende Nummer der gefundenen Schalthandlung
'CS_switch_time: Zeitpunkt der gefundenen Schalthandlung
'CS_switch_mode: fortlaufende Nummer und Richtung (close/open) der gefundenen Schalthandlung

sub find_cs(name)

dim cs_name, group_name, chn_pos
dim chn_diff_max
dim switch_number
dim diff_ch_name

switch_number=1

group_name="Travel_"&Name
cs_name=group_name&"/CS"
T5= "Travel_"&Name&"/CS_Diff_Y"

Call GroupDefaultSet(groupindexget(group_name))   

Call ChnDifferentiate("",cs_name,"",group_name&"/CS_Diff_Y")      '  Ableitung des CS-Kanals bilden
Call ChnFromWfXGen(group_name&"/CS_Diff_Y",group_name&"/CS_time") ' dazugehörigen Zeitkanal erzeugen 
Call ChnPropValSet(group_name&"/noname","name","CS_time")         ' und umbenennen

'neue Kanäle anlegen
Call chnalloc("CS_switch_number",1000)
Call chnalloc("CS_switch_time",1000)
Call chnalloc("CS_switch_mode",1000,,datatypestring)

call chncharacter("CS_diff_Y")    'Kennwerte (Max/Min) des Kanals berechnen

dim max_pos, max_neg

max_pos=cmax(group_name&"/cs_diff_y")
max_neg=cmin(group_name&"/cs_diff_y")
max_neg=abs(max_neg)

R1 = 0.5 * MaxV(max_pos, max_neg)     'Grenzwert für Schaltvorgang: Mehr als 0.5 der maximalen Steigung im Kanal
chn_pos=1

Do while chn_pos < chnlength(cs_name)   'so lange das Ende des Kanals nicht erreicht ist

  chn_pos = chnfind("Abs(Ch(T5)) > R1",chn_pos) '...springe zum nächsten Schaltvorgang (Steigung grösser R1)
 
  if chn_pos = 0 then                   ' wenn nichts gefunden/Kanalende erreicht, dann Funktion beenden
    exit do
  end if
 
    
  chd(switch_number,group_name&"/CS_switch_number")=switch_number           'Nummer des Schaltvorgangs speichern
  chd(switch_number,group_name&"/CS_switch_time")=chdwfx(chn_pos,cs_name)   'Zeitpunkt bestimmen und speichern
  
  'bei positiver Steigung "close" speichern
  
  if chd(chn_pos,group_name&"/cs_diff_y") > 0 then  
    cht(switch_number,group_name&"/CS_switch_mode")="("&switch_number& ") close"
  end if

  'bei negativer Steigung "open" speichern

  if chd(chn_pos,group_name&"/cs_diff_y") <= 0 then  
    cht(switch_number,group_name&"/CS_switch_mode")="("&switch_number& ") open"
  end if
    

 switch_number = switch_number + 1
 

  chn_pos = chn_pos + 50      'um 2,5 ms weiter springen, um Fehler/Doppelerkennungen zu vermeiden

Loop 
end sub
'''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'load_cs
'kopiert channel-selection "CS" in die Kanalgruppe, skaliert den CS-Kanal auf die Hälfte des Hubs

sub load_cs(Inputchannel,Name) 'zu ladender "CS"-Kanal mit komplettem Namen, "ref" oder "comp"
dim group_name
group_name="Travel_"&Name
Call Chncopy(Inputchannel,group_name&"/CS")
scalefactor=150

if grouppropexist(groupindexget(group_name),"Antriebs_Hub__MM__") then 
    scalefactor= 0.75 * grouppropvalget(groupindexget(group_name),"Antriebs_Hub__MM__") / Cmax(Inputchannel)
    Call ChnLinScale(group_name&"/CS",group_name&"/CS",scalefactor,0) 'skalieren auf neuen Maximalwert
  elseif grouppropexist(groupindexget(group_name),"GESAMT_HUB__MM__") then 
    scalefactor= 0.75 * grouppropvalget(groupindexget(group_name),"GESAMT_HUB__MM__") / Cmax(Inputchannel)
    Call ChnLinScale(group_name&"/CS",group_name&"/CS",scalefactor,0) 'skalieren auf neuen Maximalwert
  else 
    call msgbox("`Antriebs_Hub__MM__` could not be found. Use default scaling for `CS`-signal.")
end if

Call Chncopy(group_name&"/CS",group_name&"/CS_unshifted")       'Copy Data to new Channel "CS_unshifted"  
Call Chncopy(group_name&"/CS",group_name&"/CS_2")       'Copy Data to new Channel "CS_2"  
Call Chncopy(group_name&"/CS",group_name&"/CS_3")       'Copy Data to new Channel "CS_2"  
end sub
'''''''''''''''''''''''''''''''''''''''




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'report_set
'ruft Report auf
'aktualisiert die Grenzen der X-Achse im Report
' 

sub report_set(sel_report)    

const text_passed="tolerance band test passed"
const text_failed="tolerance band test failed"

call picdelete()

'for tolerance band test, all tested channels need the same sample rate.
'if the sample rate is not the same, tolerance band cannot be checked

dim sample_rate_correct

if valequal(chnpropget("Travel_comp/Travel", "wf_increment"),chnpropget("Travel_ref/Travel", "wf_increment")) then
  sample_rate_correct=true
else
  sample_rate_correct=false
  call msgbox("Channels do not have the same samplerate."&CHR(13)&"Tolerance test is not possible.")
end if

select case sel_report

      case "0"  'single "O"
        
        Call Picload(sPathData&"ABB_plot_O_010.tdr")
        
        Call GraphObjOpen("2DAxis1")
          Call GraphObjOpen("2DXAxis25_1")
              D2AxisXBegin     =x1pos_1
              D2AxisXEnd       =x2pos_1
          Call GraphObjClose("2DXAxis25_1")
        Call GraphObjClose("2DAxis1")

      if show_tolerance =1 and sample_rate_correct then    
        Call GraphObjOpen("Text1")
        If check_tolerance ("Travel_comp/Travel",x1pos_1,x2pos_1) then
            txttxt=text_passed  
          else
            txttxt=text_failed  
        end if
        Call GraphObjclose("Text1")
     end if   
        

      case "1"    'single "C"
        
        Call Picload(sPathData&"ABB_plot_C_011.tdr")
        
        Call GraphObjOpen("2DAxis1")
          Call GraphObjOpen("2DXAxis25_1")
              D2AxisXBegin     =x1pos_1
              D2AxisXEnd       =x2pos_1
          Call GraphObjClose("2DXAxis25_1")
        Call GraphObjClose("2DAxis1")
     
     if show_tolerance =1 and sample_rate_correct then    
      Call GraphObjOpen("Text1")
        If check_tolerance ("Travel_comp/Travel",x1pos_1,x2pos_1) then
            txttxt=text_passed  
          else
            txttxt=text_failed  
        end if
        Call GraphObjclose("Text1")
     end if   

      case "2"  '"CO" one plot

        
        Call Picload(sPathData&"ABB_plot_CO_012.tdr")
        
        Call GraphObjOpen("2DAxis1")
          Call GraphObjOpen("2DXAxis25_1")
              D2AxisXBegin     =x1pos_1
              D2AxisXEnd       =x2pos_1
          Call GraphObjClose("2DXAxis25_1")
        Call GraphObjClose("2DAxis1")

    if show_tolerance =1 and sample_rate_correct then    
      Call GraphObjOpen("Text1")
        If check_tolerance ("Travel_comp/Travel",x1pos_1,x2pos_1) then
            txttxt=text_passed  
          else
            txttxt=text_failed  
        end if
        Call GraphObjclose("Text1")
    end if

      case "3"      '"CO" two plots

        Call Picload(sPathData&"ABB_plot_C_O_013.tdr")
        
          Call GraphObjOpen("2DAxis2")
            Call GraphObjOpen("2DXAxis26_1")
                D2AxisXBegin     =x1pos_1
                D2AxisXEnd       =x2pos_1
            Call GraphObjClose("2DXAxis26_1")
          Call GraphObjClose("2DAxis2")

          Call GraphObjOpen("2DAxis1")
            Call GraphObjOpen("2DXAxis25_1")
                D2AxisXBegin     =x1pos_2
                D2AxisXEnd       =x2pos_2
            Call GraphObjClose("2DXAxis25_1")
          Call GraphObjClose("2DAxis1")

if show_tolerance=1 and sample_rate_correct then
      Call GraphObjOpen("Text1")
        If check_tolerance ("Travel_comp/Travel",x1pos_1,x2pos_1) then
            txttxt=text_passed  
          else
            txttxt=text_failed  
        end if
        Call GraphObjclose("Text1")

       Call GraphObjOpen("Text26")
         If check_tolerance ("Travel_comp/Travel_2",x1pos_2,x2pos_2) then
            txttxt=text_passed  
          else
            txttxt=text_failed  
        end if
        Call GraphObjclose("Text26")
end if


      case "4"    '"O-CO" two plots

        
        Call Picload(sPathData&"ABB_plot_O_CO_014.tdr")

          Call GraphObjOpen("2DAxis2")
            Call GraphObjOpen("2DXAxis26_1")
                D2AxisXBegin     =x1pos_1
                D2AxisXEnd       =x2pos_1
            Call GraphObjClose("2DXAxis26_1")
          Call GraphObjClose("2DAxis2")

        
        Call GraphObjOpen("2DAxis1")
          Call GraphObjOpen("2DXAxis25_1")
              D2AxisXBegin     =x1pos_2
              D2AxisXEnd       =x2pos_2
          Call GraphObjClose("2DXAxis25_1")
        Call GraphObjClose("2DAxis1")
  
    if show_tolerance =1 and sample_rate_correct then    
      Call GraphObjOpen("Text1")
        If check_tolerance ("Travel_comp/Travel",x1pos_1,x2pos_1) then
            txttxt=text_passed  
          else
            txttxt=text_failed  
        end if
        Call GraphObjclose("Text1")

      Call GraphObjOpen("Text26")
         If check_tolerance ("Travel_comp/Travel_2",x1pos_2,x2pos_2) then
            txttxt=text_passed  
          else
            txttxt=text_failed  
        end if
        Call GraphObjclose("Text26")
  end if

case "5"    '"O-C-O" three plots

        
        Call Picload(sPathData&"ABB_plot_O_C_O_015.tdr")

          Call GraphObjOpen("2DAxis2")
            Call GraphObjOpen("2DXAxis26_1")
                D2AxisXBegin     =x1pos_1
                D2AxisXEnd       =x2pos_1
            Call GraphObjClose("2DXAxis26_1")
          Call GraphObjClose("2DAxis2")

        
        Call GraphObjOpen("2DAxis1")
          Call GraphObjOpen("2DXAxis25_1")
              D2AxisXBegin     =x1pos_2
              D2AxisXEnd       =x2pos_2
          Call GraphObjClose("2DXAxis25_1")
        Call GraphObjClose("2DAxis1")
  
        Call GraphObjOpen("2DAxis3")
          Call GraphObjOpen("2DXAxis27_1")
              D2AxisXBegin     =x1pos_3
              D2AxisXEnd       =x2pos_3
          Call GraphObjClose("2DXAxis27_1")
        Call GraphObjClose("2DAxis3")

    if show_tolerance =1 and sample_rate_correct then    
      Call GraphObjOpen("Text1")
        If check_tolerance ("Travel_comp/Travel",x1pos_1,x2pos_1) then
            txttxt=text_passed  
          else
            txttxt=text_failed  
        end if
        Call GraphObjclose("Text1")

      Call GraphObjOpen("Text26")
         If check_tolerance ("Travel_comp/Travel_2",x1pos_2,x2pos_2) then
            txttxt=text_passed  
          else
            txttxt=text_failed  
        end if
        Call GraphObjclose("Text26")
  
       Call GraphObjOpen("Text28")
         If check_tolerance ("Travel_comp/Travel_3",x1pos_3,x2pos_3) then
            txttxt=text_passed  
          else
            txttxt=text_failed  
        end if
        Call GraphObjclose("Text28")
  
  end if



end select

Call GraphSheetNGet(1)
Call GraphSheetRename(GraphSheetName, "Travel_comparison_"&grouppropget(groupindexget("Travel_ref"),"description")&"_"&grouppropget(groupindexget("Travel_comp"),"description"))


  If ABB_PEHLA = 1 Then
    MasterlayoutLink = "C:\DIAdem\abbext\Travel_comp\Data\" & "PEHLA_Plot_010.TDRM"
  Else
  select case labcode
    case 1
      MasterlayoutLink = "C:\DIAdem\abbext\Travel_comp\Data\" & "ABB_Plot_010.TDRM"
    case 2
      MasterlayoutLink = "C:\DIAdem\abbext\Travel_comp\Data\" & "ABB_Plot_110.TDRM"
    case 3
      MasterlayoutLink = "C:\DIAdem\abbext\Travel_comp\Data\" & "ABB_Plot_210.TDRM"
    case else
      msgbox("Error. Not able to interprete labcode=" & labcode)
  end select
  End If

call Reportmasterupd

Call Picupdate

end sub
'''''''''''''''''''''''''''''''''''''''





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'check_tolerance
'überprüft, ob ein bestimmter Zeitbereich einer Kurve innerhalb der Toleranzgrenzen liegt.

function check_tolerance(Name,t_start,t_end) 'kompletter Kanalname, Start- und Endzeit

dim i_count,i_end, inc, start_offset

inc = chnpropget(Name, "wf_increment")              'Schrittweite bestimmen
start_offset = chnpropget(Name, "wf_start_offset")  'x-Offset bestimmen

'Umrechnung der Zeitwerte in numerische Schritte

i_count = trunc((t_start - start_offset)/inc)
i_end = trunc((t_end - start_offset)/inc) + 1

'Offset zwischen den zu vergleichenden Kanälen bestimmen
i_chn_offset = trunc((chnpropget("Travel_ref/Travel_pos_tol", "wf_start_offset") - start_offset)/inc)

check_tolerance=True

Do while  i_count < i_end     'jeden Zeitpunkt überprüfen, ob der Kanal das Toleranzband über- oder unterschreitet

  if CHD(i_count - i_chn_offset,"Travel_ref/Travel_neg_tol") > CHD(i_count,Name) then
   check_tolerance=False
    exit do
  end if
  if CHD(i_count - i_chn_offset,"Travel_ref/Travel_pos_tol") < CHD(i_count,Name) then
    check_tolerance=False
    exit do
  end if

  i_count =i_count + 1

Loop

end function
'''''''''''''''''''''''''''''''''''''''




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' check_data
'
'überprüft die geladenen Daten im Datenportal


function check_data()

check_data=True

'if groupcount > 2 then
'  if Msgbox("This script does not need more than 2 groups in the data portal."&Chr(13)&"Delete all other groups?",36,"Travel Compare") = 6 then
'    while groupcount > 2            'wenn mehr als 2 Gruppen im Datenportal geladen sind, werden die letzten Gruppen gelöscht
'      call groupdel (groupcount)    ' bis nur noch 2 Gruppen übrig sind 
'    wend
'  end if
'end if

'delete existing Travel groups

if groupindexget("Travel_ref") <> 0 then groupdel(groupindexget("Travel_ref"))

if groupindexget("Travel_comp") <> 0 then groupdel(groupindexget("Travel_comp"))



if groupcount < 2 then 
  call msgbox("Please check data groups!"&Chr(13)&"Please load 2 tests and try again.",16)
  check_data=false
  exit function
  end if

'if instr(1,grouppropvalget(1,"Name"),"rbe",1)=0 then
'  call msgbox("First group is not an rbe-file!"&Chr(13)&"Load Channels as equidist-X"&Chr(13)&"Please check groups and try again.",16)
'  check_data=false
'  exit function
'end if
    
'if instr(1,grouppropvalget(2,"Name"),"rba",1)=0 then
'  call msgbox("Second group is not an rba-file!"&Chr(13)&"Please check groups and try again.",16)
 ' check_data=false
 ' exit function
'end if

'if instr(1,grouppropvalget(2,"Name"),"rbe",1)=0 then
'  call msgbox("Second group is not an rbe-file!"&Chr(13)&"Load Channels as equidist-X"&Chr(13)&"Please check groups and try again.",16)
'  check_data=false
'  exit function
'end if

'if instr(1,grouppropvalget(4,"Name"),"rba",1)=0 then
'  call msgbox("Fourth group is not an rba-file!"&Chr(13)&"Please check groups and try again.",16)
'  check_data=false
'  exit function
'end if

'if grouppropvalget(1,"TESTING_CODE")<>"MECHANICAL TB" or grouppropvalget(2,"TESTING_CODE")<>"MECHANICAL TB" then
'  if msgbox("Testing code is not `mechanical TB`"&Chr(13)&"Please check loaded data!"&Chr(13)&"Continue?",36)="VBNo" then check_data=false
'end if


end function

'''''''''''''''''''''''''''''''''''''''



'------------------------------------------------------------------------------
'
'
'sub channel_absolute(channelname)
'
'calculates the absolute value of each value in a channel
'
'Abs-Function in Calculator is not reliable

sub channel_absolute(channelname)
dim length,count
length=CL(channelname)

for count=1 to length 
  if chd(count,channelname)<0 then chd(count,channelname)=-chd(count,channelname)
next

end sub







