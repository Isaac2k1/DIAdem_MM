'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2018-03-29
'-- Author: Philip Streit, Daniel Morris, Adrian Kress
'-- Comment: 
' >>> This class (Shot.vbs) is part of the Load_Test.SUD program to be used in Diadem.
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.

Class Shot

    private File_Name
    private TNS_Nr
    private shot_Path
   
    'Definition of the seters
    Public Function setFileName(VFileName)
        File_Name = VFileName
    End Function

    Public Function setTNSNr(VtnsNr)
        TNS_Nr = VtnsNr
    End Function

    Public Function setShotPath(VshotPath)
        shot_Path = VshotPath
    End Function

'Definition of geters
    Public Function getFileName()
        getFileName = File_Name
    End Function

    Public Function getTNSNr()
        getTNSNr = TNS_Nr
    End Function

    Public Function getShotPath()
        getShotPath = shot_Path
    End Function
    
End Class

