'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2006/12/11 11:07:30
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------

Option Explicit



printeva

Sub printeva()
  Dim intHandleWrite,intErr,I,J,x,EvaInd,EvaName,InpVarListStr,EvaTitleStr,EvaQStr,PasteString
  GlobalDim "LastSelectionText,LastSelectionInd,LastSelectionType"
  GlobalDim "questions"
  GlobalDim "K,add_eva,selected_ch_ind"

  'Call UserVarReset(AutoActPath & "gvar_eva_abb_v1.3.vas")
  Call UserVarCompile(AutoActPath & "gvar_eva_abb_v1.3.vas","append")
  
  'Create the text file which you want to write the results
  intHandleWrite = TextFileOpen("C:\Program Files\National Instruments\DIAdem 10.0\abb\eva_print.txt", tfCreate or tfWrite or tfANSI)
  If intHandleWrite = 0 Then
      'Error handling
      Call MsgError(intHandleWrite)
  Else
    
    
    For J = 1 to 26
      
      EvaQStr = ""
      x = VEnum("basic_eva_ind_",J)
      EvaInd  = Trunc(Val(VEnum("basic_eva_ind_",J)))
      EvaName = VEnum("basic_eva_nam_",J)
      InpVarListStr = VEnum("basic_eva_str_",J)
    
      EvaTitleStr = EvaInd &VbCrLf& EvaName &VbCrLf& InpVarListStr &VbCrLf

      L3 = Trunc(Val(VEnum("basic_eva_totq_",J)))
      For I = 1 To L3
        'Save input(question) indexes in auxillary long integer vector LV1(max:15)
        LV1(I) = Trunc(Val(VEnum("basic_eva_q" & I &"_",J)))
        EvaQStr = EvaQStr & VEnum("question_",LV1(I)) & VbCRLF
      Next  
    
      PasteString = EvaTitleStr & EvaQStr
      'Write the title line to file
      intErr= TextfileWriteLn(intHandleWrite, PasteString)
      
      If intErr <> 0 Then
          Call MsgError(intErr)                           'Error handling
      End If

    Next
    
    
    
    intErr = TextFileclose(intHandleWrite)                   'Close file
      If intErr <> 0 Then
        Call MsgError(intErr)                              'Error handling
      End If

  End if
End sub  
  
