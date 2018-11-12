'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2006/12/21 10:50:16
'-- Author: Kaan Oenen
'-- Comment: Main Script for "Load an old Test" Menu
'-------------------------------------------------------------------------------

testload


'-------------------------------------------------------------------------------
'Sub procedure
'testload()
'compatible with dialogbox "Offline_v2.4"
'
'Description
'
'-------------------------------------------------------------------------------
'Variables
'K: Number of files


Sub testload()
  
  Dim I, ConvertListGroupInd(),NoOfChn(),CustomTimeStep
  Globaldim "K,lab_code,sDataName,sReportName,eqmode,NoOfRBDLoaded,NoOfRBALoaded,SerialNo,ShotNo"
  eqmode = 0

  If SudDlgShow("Dlg1",AutoActPath & "offline_v2.4.sud") = "IDOk" Then

    Select Case eqmode
      Case 1  'equidistant x axis requested by user
        I = 0 : TotalNoOfChn = 0
        'Get Group and Channel Indexes for equix conversion
        'Define which channelgroup to start depending on how many files were loaded
        iStartAtGroup = GroupCount - NoOfRBDLoaded - NoOfRBALoaded + 1
        iLoop = iStartAtGroup
        'Go through each Channelgroup which were loaded, and find out the indexes of rbd files
        Do While GroupCount >= iLoop
          If GroupChnCount(iLoop) <> 0 Then 'It is not a report file, but an rbd file          
            'Get the SerialNo/TestNo from DataPortal
            SerialNoCheck = Mid(GroupName(iLoop),2,4)
            TestNoCheck = Mid(GroupName(iLoop),7,4)
            
            If SerialNoCheck = SerialNo Then 'SerialNo of ChannelGroup in DataPortal matches with the one from dialogbox "offline" 
              
              For J = 1 to 2*(K+1) '2*K: Number of files expected to be loaded (rba+rbd)
                If ShotNo(J-1) = TestNoCheck Then 'TestNo found
                    I = I +1
                    Redim Preserve ConvertListGroupInd(I)    'List of Channelgroup indexes to be converted
                    ConvertListGroupInd(I) = iLoop
                    Redim Preserve NoOfChn(I)                'Total Number of Channels to be converted in the Channelgroup
                    NoOfChn(I) = (GroupChnCount(iLoop))/2
                End if  'ShotNo(J) = TestNoCheck    
              Next
                
            End if  'SerialNoCheck = SerialNo
          
          End if  'GroupChnCount(iLoop) <> 0
          iLoop = iLoop+1
        Loop
        
        'Calculate total number of channels to be converted
        For P = 1 to I
          TotalNoOfChn = TotalNoOfChn + NoOfChn(P)
        Next
        'Create array iXYChNO(?,1) which includes indexes of x-y channel pairs
        Dim iXYChNO()
        Redim iXYChNO(TotalNoOfChn-1,1)

        'Create iGroupindex() and iXYChNo() arrays as the input for sub-prodecure equix()
        For T = 1 to I 'Ubound(ConvertListGroupInd)
          TotalIndex = TotalIndex + NoOfChn(T)          
          Redim Preserve iGroupindex(TotalIndex)
          For Q = 1 to NoOfChn(T)
            iGroupindex(TotalIndex-Q) = ConvertListGroupInd(T)   'iGroupindex() first element in zero position
            iXYChNO((TotalIndex-Q),0) = 2*NoOfChn(T)-(2*Q-1)     'Assign x channel index as 1,3,5,7...
            iXYChNO((TotalIndex-Q),1) = 2*NoOfChn(T)-(2*Q-2)     'Assign y channel index as 2,4,6,8...
          Next          
        Next        
        
        'For A = 0 to TotalIndex-1
        '  msgbox iGroupindex(A) & "/" & iXYChNO(A,0) & "," & iXYChNO(A,1)
        'Next       


        'Error handling
        'If I <> NoOfRBDLoaded Then 'The number of files assigned to be converted are bigger than number of RBD files loaded 
         ' msgbox "Conversion to equidistant x can not be done automatically, please do it manually from the ABB Menu"
        'Else       
            Call equix(iGroupindex,iXYChNo,0,CustomTimeStep)
        'End if
    End Select
    


    'For I = 0 to K  'Ubound(sDataName)
    '  NameString = NameString&" "&sDataName(I)&" "&sReportName(I)
    'Next
    'Call msgbox("Selected lab_code: "&lab_code&" and "&NameString) 
  End if
    
End sub
'-------------------------------------------------------------------------------