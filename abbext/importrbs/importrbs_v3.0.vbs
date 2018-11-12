'   Last Update: 2008-09-08.15:10:30
'   Version: 3.0
'   Reviewed:	
'   Author: Kaan Oenen, Markus Averes, Mathias Knaak
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

  'M.Knaak 8.9.2008
  'bugfix: copies group desciption to rbd AND rbe groups, if "load both" is selected
  
  'M.Knaak 3.11.2008
  'Exception handling for special PEHLA-Test-No. 20080020 (saved as 8020)

  'M.Knaak 26.11.2008
  'Group Property "Time" added, if it does not exist
  
  'R.Irion 15.12.2009
  'Lines 370-374 & 385 added to start the script for Laser measurement correction plus several comment lines
  
  'F.Maurizio 31.03.2010
  'add group properties if they do not exist
  
  'F.Maurizio 07.07.2010
  ' "load single channels" function added

Sub importrbs(Dialog,EditBox1)
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
'  Call Globaldim("load_digital_chn")
  Dim sCurrVal,sBuffer,iLenCurrVal,iSlashPos,asBuffer(),I,J,iChn,iKommaPos,iXYChNo()
  Dim iMinusPos,iRangeStart, iRangeEnd,iLenAsBuffer,A,sChar,AllowLoad,iGroupindex()
  Dim GpNum,ChCnt,ChNum

  
  'As default allow to load files
  AllowLoad = 1
      
    'Save the SNo/TNo string in a variable
    sCurrVal = EditBox1.Text
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
      EditBox1.SetFocus
    
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
    
    'Exception handling for special serial number 20080020 saved as 8020 (PEHLA-Shot with 4 digit serial no.) 
    if serialno=20080020 then serialno=8020
    
    
    
    T3 = EditBox1.Text  'Assign user input to auxiliary variable
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
          dim test_time,rba_load_count,load_count, z
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
          Dim MyDataFinder, AdvancedQuery, w, ww
          Set MyDataFinder = Navigator.ConnectDataFinder(Navigator.Display.CurrDataProvider.Name)
          'Define the query type (advanced)
          Set AdvancedQuery=Navigator.CreateQuery(eAdvancedQuery)
          'Define the result type (file)
          AdvancedQuery.ReturnType=eSearchFile
          
          Call UIAutoRefreshSet(True)
          
          For J= 0 To K
                    
                   
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
              
              if InStr(channel_name,"_Y") <> 0 or InStr(channel_name,"_A") <> 0 or InStr(channel_name2,"_Y") <> 0 or InStr(channel_name2,"_A") <> 0 then
                msgbox "The '_Y' or '_A' character sequence is not possible. Whole channel group will be loaded."
                load_channel = 0
              end if
              
              if load_channel = 1 then  'Falls man einzelne Kanäle laden will (zuerst alle geladen, dann werden die nicht gesünschten gelöscht) F.M.
                ww = ChnNoMax
                w= ChnNoMax - GroupChnCount(GroupCount)
                While w < ww 
                  w=w+1
                  if channel_option = 0 then
                    if InStr(ChnName(w),channel_name) = 0 then
                      Call ChnDel(ChnName(w))
                      w=w-1        'Wenn das Programm ein Kanal löscht nummeriert es die restliche neu:
                      ww=ww-1      'Darum muss man die Zählvariable w und die Maximalvariable ww um 1 verkleinern. F.M. 07.07.2010
                    end if
                  end if
                  if channel_option = 1 then
                    if InStr(ChnName(w),channel_name) = 0 and InStr(ChnName(w),channel_name2) = 0 then
                      Call ChnDel(ChnName(w))
                      w=w-1        
                      ww=ww-1      
                    end if
                  end if
                  if channel_option = 2 then
                    if InStr(ChnName(w),channel_name) = 0 or InStr(ChnName(w),channel_name2) <> 0 then
                      Call ChnDel(ChnName(w))
                      w=w-1        
                      ww=ww-1      
                    end if
                  end if
                Wend
              
                if ChnNoMax = 0 then 
                  msgbox "No channel found which satisfies these criteria!"
                  GroupDel(GroupCount)
                  Exit sub
                end if
                
              end if 'load single channel
              
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
              load_count=NoOfRBDLoaded
              if loadmode=3 then load_count=load_count*2
              for rba_load_count=1 to load_count
                'Copy properties from rba to rbe file (M.Knaak, 10.7.2008)
                call grouppropcopy(groupcount, groupcount-rba_load_count)
                'write group desription, see "abb_user_commands.vbs" M.Knaak, 16.7.2008
                call group_desc_write(groupcount-rba_load_count)
                if grouppropexist(groupcount-rba_load_count,"TestTime") then
                  'add group property "TIME" if it does not exist  M.Knaak, 26.11.08
                  if not grouppropexist(groupcount-rba_load_count,"TIME") then call grouppropcreate(groupcount-rba_load_count,"TIME",Datatypestring)
                  call grouppropset(groupcount-rba_load_count,"TIME",grouppropget(groupcount-rba_load_count,"TestTime"))
                  else call msgbox("Could not read `TestTime` from *.rbd-file.")
                end if

                'add group properties if they do not exist  F.Maurizio, 31.03.10
                if not grouppropexist(groupcount-rba_load_count,"VERSUCHS_BEZ_") then
                   call grouppropcreate(groupcount-rba_load_count,"VERSUCHS_BEZ_",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"VERSUCHS_BEZ_","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"PRUEFKR_ID") then
                   call grouppropcreate(groupcount-rba_load_count,"PRUEFKR_ID",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"PRUEFKR_ID","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"ANLAGE_ING_") then
                   call grouppropcreate(groupcount-rba_load_count,"ANLAGE_ING_",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"ANLAGE_ING_","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"PRUEFZELLE") then
                   call grouppropcreate(groupcount-rba_load_count,"PRUEFZELLE",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"PRUEFZELLE","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"STEUERPLATZ") then
                   call grouppropcreate(groupcount-rba_load_count,"STEUERPLATZ",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"STEUERPLATZ","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"DATE") then
                   call grouppropcreate(groupcount-rba_load_count,"DATE",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"DATE","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"TESTING_CODE") then
                   call grouppropcreate(groupcount-rba_load_count,"TESTING_CODE",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"TESTING_CODE","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"CYCLE_CODE") then
                   call grouppropcreate(groupcount-rba_load_count,"CYCLE_CODE",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"CYCLE_CODE","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"CIRCUIT_CODE") then
                   call grouppropcreate(groupcount-rba_load_count,"CIRCUIT_CODE",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"CIRCUIT_CODE","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"BEMERKUNG") then
                   call grouppropcreate(groupcount-rba_load_count,"BEMERKUNG",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"BEMERKUNG","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"PROJEKTNAME") then
                   call grouppropcreate(groupcount-rba_load_count,"PROJEKTNAME",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"PROJEKTNAME","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"KAMMERN") then
                   call grouppropcreate(groupcount-rba_load_count,"KAMMERN",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"KAMMERN","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"KUNDE") then
                   call grouppropcreate(groupcount-rba_load_count,"KUNDE",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"KUNDE","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"V_LEITER") then
                   call grouppropcreate(groupcount-rba_load_count,"V_LEITER",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"V_LEITER","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"HERSTELLER") then
                   call grouppropcreate(groupcount-rba_load_count,"HERSTELLER",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"HERSTELLER","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"VERSUCHS_TYP") then
                   call grouppropcreate(groupcount-rba_load_count,"VERSUCHS_TYP",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"VERSUCHS_TYP","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"VERSUCHS_NR") then
                   call grouppropcreate(groupcount-rba_load_count,"VERSUCHS_NR",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"VERSUCHS_NR","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"VARIANTE") then
                   call grouppropcreate(groupcount-rba_load_count,"VARIANTE",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"VARIANTE","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"ZEICHNUNGS_NR") then
                   call grouppropcreate(groupcount-rba_load_count,"ZEICHNUNGS_NR",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"ZEICHNUNGS_NR","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"KONTO") then
                   call grouppropcreate(groupcount-rba_load_count,"KONTO",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"KONTO","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"PRUEF_NORM") then
                   call grouppropcreate(groupcount-rba_load_count,"PRUEF_NORM",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"PRUEF_NORM","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"SCHICHTTYP") then
                   call grouppropcreate(groupcount-rba_load_count,"SCHICHTTYP",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"SCHICHTTYP","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"BEOBACHTER_1") then
                   call grouppropcreate(groupcount-rba_load_count,"BEOBACHTER_1",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"BEOBACHTER_1","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"BEOBACHTER_2") then
                   call grouppropcreate(groupcount-rba_load_count,"BEOBACHTER_2",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"BEOBACHTER_2","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"KIND_OF_OBJECT") then
                   call grouppropcreate(groupcount-rba_load_count,"KIND_OF_OBJECT",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"KIND_OF_OBJECT","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"SERIAL_NO_TO") then
                   call grouppropcreate(groupcount-rba_load_count,"SERIAL_NO_TO",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"SERIAL_NO_TO","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"U_NENN__KV__") then
                   call grouppropcreate(groupcount-rba_load_count,"U_NENN__KV__",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"U_NENN__KV__","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"IK_NENN__KA__") then
                   call grouppropcreate(groupcount-rba_load_count,"IK_NENN__KA__",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"IK_NENN__KA__","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"FREQUENZ__HZ__") then
                   call grouppropcreate(groupcount-rba_load_count,"FREQUENZ__HZ__",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"FREQUENZ__HZ__","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"SOLL_PRUEFSTROM__KA__") then
                   call grouppropcreate(groupcount-rba_load_count,"SOLL_PRUEFSTROM__KA__",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"SOLL_PRUEFSTROM__KA__","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"MEDIUM") then
                   call grouppropcreate(groupcount-rba_load_count,"MEDIUM",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"MEDIUM","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"PF_IST__BAR__") then
                   call grouppropcreate(groupcount-rba_load_count,"PF_IST__BAR__",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"PF_IST__BAR__","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"ANTRIEB") then
                   call grouppropcreate(groupcount-rba_load_count,"ANTRIEB",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"ANTRIEB","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"SERIAL_NO_DRIVE") then
                   call grouppropcreate(groupcount-rba_load_count,"SERIAL_NO_DRIVE",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"SERIAL_NO_DRIVE","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"MOTOR_SPG__V__") then
                   call grouppropcreate(groupcount-rba_load_count,"MOTOR_SPG__V__",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"MOTOR_SPG__V__","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"MOTORSPG_TYP") then
                   call grouppropcreate(groupcount-rba_load_count,"MOTORSPG_TYP",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"MOTORSPG_TYP","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"TRIPPING__V__") then
                   call grouppropcreate(groupcount-rba_load_count,"TRIPPING__V__",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"TRIPPING__V__","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"HUB_N1__MM__") then
                   call grouppropcreate(groupcount-rba_load_count,"HUB_N1__MM__",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"HUB_N1__MM__","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"HUB_N2__MM__") then
                   call grouppropcreate(groupcount-rba_load_count,"HUB_N2__MM__",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"HUB_N2__MM__","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"GESAMT_HUB__MM__") then
                   call grouppropcreate(groupcount-rba_load_count,"GESAMT_HUB__MM__",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"GESAMT_HUB__MM__","*****")
                end if
                if not grouppropexist(groupcount-rba_load_count,"ANTRIEBS_HUB__MM__") then
                   call grouppropcreate(groupcount-rba_load_count,"ANTRIEBS_HUB__MM__",Datatypestring)
                   call grouppropset(groupcount-rba_load_count,"ANTRIEBS_HUB__MM__","*****")
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

          ' This is the script to correct laser measurement data. Complete script execution is omitted if variables 
          ' §LM_channel_1 or 2 do not exist in RBA-file. Script checks for these two variables only.
          Call scriptstart(AutoActPath & "..\laser_time_shift\time_shift_of_laser_measurement.VBS", "automatic")
          
          ' Insert here new scripts for automatic data manipulation after shotload.

          
          Next
          
          'Call MsgboxDisp("The requested file(s) are loaded.(total " &(NoOfRBDLoaded+NoOfRBALoaded)&_
          '                 " files. Number of RBD Files: " &NoOfRBDLoaded& " Number of RBD Files: " &NoOfRBDLoaded&" )."&_
          '                 " Click 'Exit' if you are done. Or Enter another SerialNo/TestNo and Click 'Load' ",MBNO_BUTTON,MsgTypeNote,,5)
          
          call rbd_unit_rename() 'copies unit in all rbd-groups from "chnunit" to "unit" 

          ' Insert here new scripts for automatic data manipulation after testload.
          
          Call Dialog.OK()
          
      End If
  
End Sub 
