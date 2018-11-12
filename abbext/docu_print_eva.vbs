'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2006/12/11 11:07:30
'-- Author: Kaan Oenen 
'-- Comment: This Scripts prints the docu and updates definitions to a txt file
'            including evaluation index, name, input variables string and
'            evaluation elements (questions)
'-------------------------------------------------------------------------------

Option Explicit



printeva

Sub printeva()
  Dim intHandleWrite,intErr,I,J,x,EvaInd,EvaName,InpVarListStr,EvaTitleStr,EvaQStr,PasteString


  'Call UserVarReset(AutoActPath & "gvar_eva_abb_v1.3.vas")
  Call UserVarCompile(AutoActPath & "gvar_eva_abb_v1.4.vas","append")
  
  'Create the text file which you want to write the results
  intHandleWrite = TextFileOpen("C:\Program Files\National Instruments\DIAdem 10.0\abb\eva_doc_print1.txt", tfCreate or tfWrite or tfANSI)
  If intHandleWrite = 0 Then
      'Error handling
      Call MsgError(intHandleWrite)
  Else
    
    
    For J = 1 to 12
      
      EvaQStr = ""
      EvaInd  = Trunc(Val(VEnum("doc_and_up_ind_",J)))
      EvaName = VEnum("doc_and_up_nam_",J)
      InpVarListStr = VEnum("doc_and_up_str_",J)
    
      EvaTitleStr = EvaInd &VbCrLf& EvaName &VbCrLf& InpVarListStr &VbCrLf

      L3 = Trunc(Val(VEnum("doc_and_up_totq_",J)))
      For I = 1 To L3
        'Save input(question) indexes in auxillary long integer vector LV1(max:15)
        LV1(I) = Trunc(Val(VEnum("doc_and_up_q" & I &"_",J)))
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
  
