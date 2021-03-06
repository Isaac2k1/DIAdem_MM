'-------------------------------------------------------------------------------
'-- VBS-Script-Datei
'-- Neu erstellt am 12.02.2009 13:05:05
'-- Ersteller: Markus Averes
'-- Kommentar: Wird z.B. benutzt in ABB2009.tdrm (script muss in Einstellungen -> Erweiterungen -> Optionen -> Anwenderbefehle bekannt sein
'             Date: @@DateEnglish(CurrDateTimeReal)@@
'-------------------------------------------------------------------------------
Option Explicit  'Erzwingt die explizite Deklaration aller Variablen in einem Skript.
function DateEnglish(DateVal)
  Dim iMonth, sMonth
  iMonth = RTP(DateVal,"t")
  select case iMonth
    case  1 : sMonth = "January"
    case  2 : sMonth = "February"
    case  3 : sMonth = "March"
    case  4 : sMonth = "April"
    case  5 : sMonth = "May"
    case  6 : sMonth = "June"
    case  7 : sMonth = "July"
    case  8 : sMonth = "August"
    case  9 : sMonth = "September"
    case 10 : sMonth = "October"
    case 11 : sMonth = "November"
    case 12 : sMonth = "December"
  end select
  DateEnglish = sMonth & " " & str(CurrDatetimereal, "#dd") & ", " & str(CurrDatetimereal, "#YYYY")
end function