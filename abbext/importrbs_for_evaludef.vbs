'   Last Update: 2008-07-31.09:37:52
'   Version: 2.6.1
'   Reviewed:	





  '
  'M. Averes 25.04.2008: 
  'Don't load files from error directories (ToDo: window to ask, if user want also to load files from here) 
  
  'Mathias Knaak 7.7.2008
  'rbd_unit_rename() added, copies unit description in rbd-groups

  'v2.4  M.Knaak 10.7.2008:
  'group properties copied to rbd/rbe-Groups
  'does not load digital channels if not selected

  'v2.5  M.Knaak 11.7.2008
  'copies properties from rba-Groups, does NOT create *.rba-Groups

  'v2.6 M.Knaak 14.7.2008
  'select, if standard data or error data should be loaded
  
  ' M.Knaak 16.7.2008
  ' writes serial number and shot number into group description

'defined in "batch_processing_v1.0.vbs"

'call globaldim ("loadmode, labcode,load_error_files,load_digital_chn")
'
'loadmode=2
'labcode=1
'load_error_files=0
'load_digital_chn=0


'call importrbs_for_evaludef("2517/358-365")

Sub importrbs_for_evaludef(Input_string,load_code)

'load_code=val(load_code)

'Dim This : Set This = Load
  'Here you find the actions to be taken when the user clicks OK
  'v1.4
  'GUI-less query
  'v1.6
  'detailed comments added
  'v2.0
  'v2.1
  'laboratory selection added
  'v2.2
  'Previous and Next buttons added
  'v2.3
  'Previous and Next buttons removed (5th Jan 2007, decided together with Marco Mailand) 
  'Syntax check and error handling for SerialNo/TestNo input field had big changes its If-ElseIf structure

  Call Globaldim("sDataName")
  Call Globaldim("sReportName")
  Call Globaldim("labcode")
  Call Globaldim("SerialNo")
  Call Globaldim("ShotNo")
  Call GlobalReDim("ShotNo(1)")
  Call Globaldim("K")
  'if not iteminfoget("shot_count") then call 
  call globaldim("shot_count")
'  Call Globaldim("load_digital_chn")
  Dim sCurrVal,sBuffer,iLenCurrVal,iSlashPos,asBuffer(),I,J,iChn,iKommaPos,iXYChNo()
  Dim iMinusPos,iRangeStart, iRangeEnd,iLenAsBuffer,A,sChar,AllowLoad,iGroupindex()
  Dim GpNum,ChCnt,ChNum
  shot_count=0
  
  'As default allow to load files
  AllowLoad = 1
      
    'Save the SNo/TNo string in a variable
    sCurrVal = Input_string
    'Find the position of "/" in SNo/TNo string
    iSlashPos = InStr(sCurrVal, "/")
    'Save the length of SNo/TNo string
    iLenCurrVal = Len(sCurrVal)
    
'-------------- Start of User Input Check (SerialNo) -----------------------------------------------

    'Check the syntax of SNo/TNo string
      A=1
      Do While A <= Len(sCurrVal)
        sChar = Mid(sCurrVal,A,1)
        If (sChar <> "1" ) And (sChar <> "2" ) And (sChar <> "3" ) And (sChar <> "4" ) And (sChar <> "5" ) And (sChar <> "6" ) And _
           (sChar <> "7" ) And (sChar <> "8" ) And (sChar <> "9" ) And (sChar <> "0" ) And (sChar <> "," ) And (sChar <> "-" ) And (sChar <> "/" ) Then
          msgbox "You can enter only numbers from 1 to 9 and ' / '   ' , '   ' - '  "
          'Do not allow to load files
          AllowLoad = 0
          Exit Do
        End If
        A = A +1        
      Loop
    
    'If the user didn't enter anything to Editbox1, warn user
    If sCurrVal = "" Then
      msgbox "Please enter the Serial No / Test No!"
      'Do not allow to load files
      AllowLoad = 0
'      EditBox1.SetFocus
    
    'Check if the user selected laboratory type
    Elseif labcode<1 Or labcode>3 Then
      msgbox "Please select laboratory!"
      AllowLoad = 0
    
    'Check if the user selected laboratory type
    Elseif loadmode <1 Or loadmode>3 Then
      msgbox "Please select Load-Mode"
      AllowLoad = 0
      
    'Check if "/" is used to seperate Serial No and Test No
    Elseif iSlashPos = 0 Then
      msgbox "Please use '/' To seperate SerialNo/Shotno"
      AllowLoad = 0
    
    'Check if the user entered TestNo after "/"
    Elseif iSlashPos = iLenCurrVal Then
      msgbox "Please enter TestNo after '/'" 
      AllowLoad = 0
      
    Elseif iSlashPos = 1 Then
      msgbox "Invalid SerialNo (it must have 1, 2, 3, 4 or 8 digits)!"
      AllowLoad = 0
      
    Elseif iSlashPos <= 5 Then
      'Save 4 digit Serial No
      SerialNo = Left(sCurrVal,iSlashPos-1)     
      
      'Precede as much "0" as needed to create SerialNo string with length of 4
      Dim G : G = 5-iSlashPos
      Do While G > 0 
        SerialNo="0"&SerialNo
        g = g-1
      Loop        
      'Save the part of the string after "/"
      sBuffer = Mid(sCurrVal,iSlashPos+1,iLenCurrVal-iSlashPos)
    
    Elseif iSlashPos < 9 Then
      msgbox "Invalid SerialNo (it must have 1, 2, 3, 4 or 8 digits)!"
      AllowLoad = 0
      
    Elseif iSlashPos = 9 Then
      'Save 8 digit Serial No (PEHLA)
      SerialNo = Left(sCurrVal,8)
      'Save the part of the string after "/"
      sBuffer = Mid(sCurrVal,iSlashPos+1,iLenCurrVal-iSlashPos)
    
    Elseif iSlashPos > 9 Then
        msgbox "Invalid SerialNo (it must have 1, 2, 3, 4 or 8 digits)!"
        AllowLoad = 0
    
    End If  'sCurrVal = ""
    
    T3 = Input_string  'Assign user input to auxiliary variable
'-------------- End of User Input Check (SerialNo) -----------------------------------------------        

'-------------- Start of Analysis of TestRange given by User ---------------------------------------        

      I = 0
      'Find out if there is a range seperated with ","
      If InstrRev(sBuffer, ",") = 0 Then  'only one TestNo or range exist
        Redim Preserve asBuffer(0)
        asBuffer(0) = sBuffer
      Else  'There is a range of tests
        
        While (Not InStrRev(sBuffer,",") = 0)
          Redim Preserve asBuffer(I)
          iKommaPos = InStrRev(sBuffer,",")
          asBuffer(I) = Mid(sBuffer,iKommaPos+1,Len(sBuffer)-iKommaPos) 
          sBuffer = Left(sBuffer,iKommaPos-1)
          I = I +1
        Wend
          'Save the strings seperated by "," into asBuffer()
          'eg: If sBuffer = "1-5,10" Then asBuffer(0) = "10" , asBuffer(1) = "1-5"
          Redim Preserve asBuffer(I)
          asBuffer(I) = sBuffer
      End If 'InstrRev(sBuffer, ",") = 0
      
      'Find out if there is elements in asBuffer that defines a range with "-"
      K = -1
      For J=0 To I
        'Save the position of "-" in the string
        iMinusPos = InStr(asBuffer(J),"-")
        'Save the length of the string
        iLenAsBuffer = Len(asBuffer(J))
        
        'If there is no range defined by "-" then it is a single test number
        If iMinusPos = 0 Then
          K = K+1
          Globalredim "Preserve ShotNo("&K&")"
          'Redim Preserve ShotNo(K)
            'G = Len(asBuffer(J))
            If iLenAsBuffer > 4 Then
              msgbox "You entered a test no with more than 4 digits"
              'Do not allow loading files
              AllowLoad = 0
            Else
              'Append as much "0" as needed to create 4 digit ShotNo
              Do While iLenAsBuffer < 4
                asBuffer(J) = "0"&asBuffer(J)
                iLenAsBuffer = iLenAsBuffer + 1
              Loop
              'Save single ShotNos in array ShotNo()
              ShotNo(K) = asBuffer(J)
            End If
        Else  'There is a range
          'Save range start and range end
          iRangeStart = Left(asBuffer(J),iMinusPos-1)
          iRangeEnd = Right(asBuffer(J),iLenAsBuffer-iMinusPos)
          'Check range conditions
          If Cint(iRangeStart) > Cint(iRangeEnd) Then
            msgbox "Enter the TestNo range like : TestNoMin-TestNoMax"
            'Do not allow loading files
            AllowLoad = 0
          Elseif (Len(iRangeStart) > 4) Or (Len(iRangeEnd) > 4) Then
            msgbox "TestNo range can only be between 1...9999"
            'Do not allow loading files
            AllowLoad = 0
          Else
            'Go through the range to save single ShotNos  
            For L = iRangeStart To iRangeEnd
              K = K+1
              'Add an element to ShotNo Array
              Globalredim "Preserve ShotNo("&K&")"
              Do While Len(L) < 4
                L = "0"&L
              Loop
              'Save ShotNos
              ShotNo(K) = L
            Next
          End If 'iRangeStart > iRangeEnd 
        End If 'iMinusPos = 0 
      Next
'-------------- End of Analysis of TestRange given by User ---------------------------------------
      
      'Define prefix depending on users Labarotory selection
      Select Case labcode
        Case 1          'high power laboratory was selected
          prefix = "l"  
        Case 2          'high voltage laboratory was selected
          prefix = "v"
        Case 3          'mechanic laboratory was selected
          prefix = "m"
      End Select
      
      'Load the files if it is allowed ( if no syntax error found in user input)
      
      If(AllowLoad) Then
          Globalredim "sDataName("&K&")"
          Globalredim "sReportName("&K&")"
          Globaldim "NoOfRBDLoaded"
          Globaldim "NoOfRBALoaded"
          dim test_time,rba_load_count
          NoOfRBDLoaded = 0
          NoOfRBALoaded = 0
          
          'Create the strings to use for file search
          For J = 0 To K
            sDataName(J)   = prefix&SerialNo&"-"&ShotNo(J)&"*.rbd"
            sReportName(J) = prefix&SerialNo&"-"&ShotNo(J)&".rba"
          Next
      
          'Dialog.OK
          'Call DATADELALL(1)                      '... HEADERDEL 
              
          'Connect to DataFinder
          Dim MyDataFinder, AdvancedQuery
          Set MyDataFinder = Navigator.ConnectDataFinder(Navigator.Display.CurrDataProvider.Name)
          'Define the query type (advanced)
          Set AdvancedQuery=Navigator.CreateQuery(eAdvancedQuery)
          'Define the result type (file)
          AdvancedQuery.ReturnType=eSearchFile
          
          Call UIAutoRefreshSet(true)
          
          dim load_start, load_end
          
          shot_count=K
          if load_code=-1 or load_code>K then
            load_start=0
            load_end=K
          elseif load_code=>-1 and load_code<=K then
             load_start=load_code
             load_end=load_code 
          end if
          
          For J= load_start To load_end
                    
                   
              Call AdvancedQuery.Conditions.RemoveAll()          
            'Search for Binary Data File "rbd"
            Call AdvancedQuery.Conditions.Add(eSearchFile,"fileName","=",sDataName(J))

            'select if "data" files or "error" files should be loaded
            if load_error_files=0 then
              Call AdvancedQuery.Conditions.Add(eSearchFile,"fullpath","=","*data*")          
            elseif load_error_files=1 then
              Call AdvancedQuery.Conditions.Add(eSearchFile,"fullpath","=","*error*")          
            end if
            Call MyDataFinder.Search(AdvancedQuery)
             
            If(MyDataFinder.Results.Count = 0) Then
             msgbox "No test data found for "&sDataName(J)
            Elseif(MyDataFinder.Results.Count > 0) Then 'Load single or multiple rbd Files
             For Each Element In MyDataFinder.Results
              Call Navigator.LoadData(Element)
              'remove digital channels (M.Knaak 10.7.2008)
              if load_digital_chn=0 then call remove_digital_channels(groupcount) 
              
              if load_error_files=1 then
                groupname(groupcount)=left(groupname(groupcount),instr(groupname(groupcount),"rbd")-2)&"_ERROR.rbd"
              end if      
              'write group desription, see "abb_user_commands.vbs" M.Knaak, 16.7.2008
              call group_desc_write(groupcount)
              
               'Equidistant Channels
              If(loadmode>1) Then
               GpNum = GroupCount
               ChCnt = GroupChnCount(GpNum)
               ChNum = ChnNoMax-ChCnt+1 'first Channel to be converted
               Redim iGroupindex(ChCnt/2-1), iXYChNo(ChCnt/2-1,1)
               
               For iChn = 1 To ChCnt/2
                iGroupindex(iChn-1) = GpNum
                iXYChNo(iChn-1,0)   = ChNum
                iXYChNo(iChn-1,1)   = ChNum+1
                ChNum = ChNum+2
               Next
               
               Call equix(iGroupindex,iXYChNo,0,CustomTimeStep)
               If(loadmode=2) Then Call GroupDel(GpNum)
              End If

              
             Next 'Each Element In MyDataFinder.Results
            NoOfRBDLoaded=MyDataFinder.Results.Count
              If (MyDataFinder.Results.Count > 1) then
               msgbox "There are more than one rbd file with the name "&sDataName(J)&" All " &MyDataFinder.Results.Count &" files are loaded!"
             End if
            End If 'MyDataFinder.Results.Count = 0
           
            
            'Search for ASCII report files "rba"
            Call AdvancedQuery.Conditions.RemoveAll()
            Call AdvancedQuery.Conditions.Add(eSearchFile,"fileName","=",sReportName(J))
            MyDataFinder.Search(AdvancedQuery)
            
            If MyDataFinder.Results.Count = 0 Then
            call  msgbox ("No test report found for "&sReportName(J)&chr(13)&"Could not copy properties.")
            Elseif MyDataFinder.Results.Count = 1 Then 'Load a single rba File
              Navigator.LoadData(MyDataFinder.Results)
              'Save the number of files loaded
              NoOfRBALoaded = NoOfRBALoaded +1    
              
              for rba_load_count=1 to NoOfRBDLoaded
                'Copy properties from rba to rbe file (M.Knaak, 10.7.2008)
                call grouppropcopy(groupcount, groupcount-rba_load_count)
                'write group desription, see "abb_user_commands.vbs" M.Knaak, 16.7.2008
                call group_desc_write(groupcount-rba_load_count)
                if grouppropexist(groupcount-rba_load_count,"TestTime") then    
                  if grouppropexist(groupcount-rba_load_count,"TIME") then
                    call grouppropset(groupcount-rba_load_count,"TIME",grouppropget(groupcount-rba_load_count,"TestTime"))
                  elseif grouppropexist(groupcount-rba_load_count,"TIME_OF_TEST") then
                    call grouppropset(groupcount-rba_load_count,"TIME_OF_TEST",grouppropget(groupcount-rba_load_count,"TestTime"))
                  else call msgbox("Could not set `Time` in *.rba-file.")
                  end if
                else call msgbox("Could not read `TestTime` from *.rbd-file.")
                end if 
              next
              'remove *.rba group (M.Knaak)
              call groupdel(groupcount)
            Elseif MyDataFinder.Results.Count > 1 Then 'Load multiple rba Files
              Navigator.LoadData(MyDataFinder.Results)
              NoOfRBALoaded = NoOfRBALoaded + MyDataFinder.Results.Count
              'Warn user
              msgbox "There are more than one report file with the name "&sReportName(J)&" All files are loaded!"&chr(13)&"No Properties copied from *.rba to *.rbe group."
            End If           
          Next
          
          'Call MsgboxDisp("The requested file(s) are loaded.(total " &(NoOfRBDLoaded+NoOfRBALoaded)&_
          '                 " files. Number of RBD Files: " &NoOfRBDLoaded& " Number of RBD Files: " &NoOfRBDLoaded&" )."&_
          '                 " Click 'Exit' if you are done. Or Enter another SerialNo/TestNo and Click 'Load' ",MBNO_BUTTON,MsgTypeNote,,5)
          
          call rbd_unit_rename() 'copies unit in all rbd-groups from "chnunit" to "unit" 
'          Call Dialog.OK()
      End If
End Sub 
