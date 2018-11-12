Data Plugins rbd.vbs and rba.vbs should always be copies of the files
with the biggest version number. 

It results in the double existence of the Data Plugins! (one with the
the most actual version number and rbd.vbs or rba.vbs)

This system guarantees to have a non-changing filename of the Data Plugin to make
sure that the batch file for preparing DIAdem works properly (Prepare-
DIAdem.bat, PrepareDIAdem_german_WXP.bat)

To see the content of batch files go to :
I:\HG\Lab\13 Software\DIAdem
