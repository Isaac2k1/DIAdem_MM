'-------------------------------------------------------------------------------
'-- VBS-Script-Datei
'-- 24.07.2008 
'-- Ersteller: Mathias Knaak  
'-- Kommentar: "travel compare" to compare, check and plot Travel curves
'   Last Update: 2008-07-29.11:23:30
'   Version: 1.3.0
'   Reviewed:
'-------------------------------------------------------------------------------
'
'description: "travel compare" to compare, check and plot travel curves
' used in "ABB"-Menu
'
'History:
'
'v1.0 first complete program version, uses .rbe and rba. files
'
'v1.1 uses only *.rbe groups
'
'v1.2 tolerance band and tolerance test only in open CS-position
'
'v1.3 manual y-shift added, 
'     load channel: "offset remove" after "clean"
'     manual timeshift "fine" changed from steps of 1 ms to 0.25 ms 




'
'
'used subroutines: travel_functions.vbs, travel_compare_dialog.sud
                    ' all files have to be in the same directory 




Option Explicit  'Erzwingt die explizite Deklaration aller Variablen in einem Skript.


If Not ItemInfoGet("sPathDocuments") Then
  Call GlobalDim("sPathDocuments")
  Call GlobalDim("sPathData")
End If

If not iteminfoget ("labcode") then
  call autoquit("An error occured, please empty your Dataportal and load Data again with ABB Test Load!")
End if
  
sPathDocuments = AutoActPath
sPathData   = AutoActPath & "\Data\"


' Funktionen aus "travel_functions" einbinden

Call ScriptCmdAdd(AutoActPath&"travel_functions.VBS")


''''''''''''''''''''''''''''''''
' Variablendeklaration
' 
' Chn_ref, Chn_comp                             Kanalname für den zu ladenden Kanal
' remove_offset_ref, remove_offset_comp         Offset entfernen ja/nein
' offset_auto_ref, offset_auto_comp             Offset automatisch/manuell entfernen
' r_offset_man_val, c_offset_man_val            Wert der manuellen Offsetkorrektur
' clean_ref, clean_comp                         Messwerte bereinigen ja/nein
' clean_ref_limit, clean_comp_limit             Grenzwert der clean-Funktion                    
' scale_ref, scale_comp                         Messwerte skalieren ja/nein
' scale_val_ref, scale_val_comp                 Endwert der Skalierung

If Not ItemInfoGet("remove_offset_ref") Then

  Call Globaldim ("Chn_ref, remove_offset_ref, offset_auto_ref, r_offset_man_val, clean_ref, clean_ref_limit, scale_ref, scale_val_ref")
  Call Globaldim ("Chn_comp, remove_offset_comp, offset_auto_comp, c_offset_man_val, clean_comp, clean_comp_limit, scale_comp, scale_val_comp")


  Chn_ref=""
  remove_offset_ref=1
  offset_auto_ref=1
  r_offset_man_val=0
  clean_ref=0
  clean_ref_limit=8000
  scale_val_ref=0

  Chn_comp=""
  remove_offset_comp=1
  offset_auto_comp=1
  c_offset_man_val=0
  clean_comp=0
  clean_comp_limit=8000
  scale_val_comp=0


  ' cs_chn_ref, cs_chn_comp                       Kanalname für den zu ladenden CS-Kanal

  Call Globaldim ("Cs_chn_ref, Cs_chn_comp")

  cs_chn_ref=""
  cs_chn_comp=""

  ' X-Werte für die erste Grafik

  Call Globaldim ("x1Pos_1, xwidth_1, x2pos_1")

  x1Pos_1 = 0.05
  x2Pos_1 = 0.15
  xwidth_1 = 0.1

  ' X-Werte für die zweite Grafik

  Call Globaldim ("x1Pos_2, xwidth_2, x2pos_2")

  x1Pos_2 = 0.1
  x2Pos_2 = 0.3
  xwidth_2 = 0.2

  ' X-Werte für die dritte Grafik

  Call Globaldim ("x1Pos_3, xwidth_3, x2pos_3")

  x1Pos_3 = 0.1
  x2Pos_3 = 0.3
  xwidth_3 = 0.2



  Call Globaldim ("timeshift_count, timeshift_view, show_tolerance,scale_tol,shift_tol,edit_preview_plot")
  Call Globaldim ("timeshift_count_2, timeshift_view_2, report_type, use_cs,ABB_PEHLA")
  Call Globaldim ("timeshift_count_3, timeshift_view_3")
  timeshift_count=1
  timeshift_count_2=1
  timeshift_view=""
  timeshift_view_2=""
  timeshift_count_3=1
  timeshift_view_3=""  
  scale_tol=200
  shift_tol=0
  show_tolerance=0
  edit_preview_plot=1
  report_type=0
end if

scale_ref=0
scale_comp=0
use_cs=1
ABB_PEHLA=0






dim savename

if check_data=true then   ' start program only with correct data groups 

   if Suddlgshow("chn_select", sPathDocuments & "Travel_compare_dialog")="IDOk" then
    call report_set(report_type)                             'create report
    
    'create filename for PDF-Export    
    savename=grouppropget(groupindexget("Travel_ref"),"description")&"__"&grouppropget(groupindexget("Travel_comp"),"description")&".PDF"
    
    ' / is not allowed in filename, "/" will be replaced by "_"
    savename=replace (savename,"/","_")
    savename="C:\Travel_comparison_"&savename
    
    'Call FileSave-Dialog 
    If (FileNameGet("ANY", "FileWrite",savename,"PDF",,,"Export PDF") = "IDOk") Then ' Dialog closed with Ok
      Call PicPDFExport(FileDlgName)  'Save as PDF
      Call Msgbox("PDF-Report saved in "&Filedlgname)
      Call ExtProgram("",Filedlgname) 'Show saved PDF in external Viewer
    End If
  
  
  
  end if

end if


