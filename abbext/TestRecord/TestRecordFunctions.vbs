'-------------------------------------------------------------------------------
'-- VBS-Script-File
'-- Created: 2010-08-02 
'-- Authors: Jonas Schwammberger
'-- Version: 1.0.0

'-- Purpose: Functions called by the GUI of TestRecord
'-- History:
'-------------------------------------------------------------------------------
Option Explicit
Dim SerialNumber
'-------------------------------------------------------------------------------
'add diadem2Excel funtion library
Call ScriptCmdAdd("C:\DIAdem\abbext\equix\equix.vbs")
'-------------------------------------------------------------------------------



'-------------------------------------------------------------------------------
'summary: 
sub GetAllProperties(Prop,Propertycount)
    Dim i
    dim j
    dim found
    
    found = false
    
    'if it's the first call, add all properties
    if PropertyCount = 0 then
        redim Prop(Data.root.ActiveChannelGroup.Properties.count-1)
        PropertyCount = Data.root.Properties.count-1
        
        for i = 0 to PropertyCount
            Prop(i) = Data.root.ActiveChannelGroup.Properties(i+1).Name
        next
    else
        for i = 1 to Data.root.ActiveChannelGroup.Properties.count
            found = false
            
            for j = 0 to Propertycount
                if Prop(j) = Data.root.ActiveChannelGroup.Properties(i).Name then
                    found = true
                End if
            next
            
            if found = false then
                Propertycount = Propertycount + 1
                redim preserve Prop(Propertycount)
                Prop(Propertycount) = Data.root.ActiveChannelGroup.Properties(i).Name
            End if
        next
    end if
End sub
'-------------------------------------------------------------------------------



'-------------------------------------------------------------------------------
'summary:
'parameter: none
'output: none
sub CreateRecord(Serial,Testshift,Shotcount,Datatable,Datacount,PDFPath,PDFHeading,PDFCount)
    Dim i
    Dim PDF
    Dim Boxresponse
    
    set PDF = new TestRecordPDF
    call PDF.init(serial)
    
    'set shots
    for i = 0 to Shotcount
      call PDF.AddShot(TestShift(i))
    next
    
    'Set Datatable
    Call PDF.setDatatable(Datatable,Datatable,dataCount)
    
    'add picture
    for i = 0 to PDFCount
      call pdf.addPicture(PDFHeading(i),PDFPath(i))
    next
    
    'only write pdf, if user has selected an output path
    if PDF.getFilePathUser <> false then
      Call PDF.WritePDF()
    End if
end sub
'-------------------------------------------------------------------------------



'-------------------------------------------------------------------------------
class FrontPage
  Dim TextGroup
  dim Column1
  Dim Column2
  dim PropertyName
  dim PropertyCount
  dim Template
  'dim chnLen: chnLen = 13
  '-------------------------------------------
  
  
  '-------------------------------------------
  sub initVariables(FrontpageTemplate)
    TextGroup = "ReportData"
    Template = FrontPageTemplate
    PropertyCount = 24
    redim PropertyName(PropertyCount)
    redim column1(PropertyCount)
    redim column2(PropertyCount)
    
    PropertyName(0)= "name"
    PropertyName(1) = ""
    PropertyName(2)= "DATE_TIME"
    PropertyName(3)= "CLIENT"
    PropertyName(4)= "MANUFACTURER"
    PropertyName(5) = ""
    PropertyName(6) = ""
    PropertyName(7)= "PROJECT_NAME"
    PropertyName(8)= "TEST_OBJECT"
    PropertyName(9) = "TO_VERSION"
    PropertyName(10) = "DRAWING_NUMBER"
    PropertyName(11)= "KIND_OF_OBJECT"
    PropertyName(12)= "U_RATED__KV__"
    PropertyName(13)= "I_SC_RATED__KA__"
    PropertyName(14)= "F_RATED__HZ__"
    PropertyName(15)= "MEDIUM"
    PropertyName(16)= "FILLING_PRESSURE_BA__"
    PropertyName(17) = "RATED_PROPERTY_HERE"
    PropertyName(18) = ""
    PropertyName(19) = ""
    PropertyName(20)= "KIND_OF_TEST"
    PropertyName(21)= "STANDARD"
    PropertyName(22)= "TEST_CURRENT__KA__"
    PropertyName(23)= "TEST_CIRCUIT_ID"
    PropertyName(24) = "REVISION_PROPERTY_HERE"

    'set column1
    Column1(0) = "Test number"
    Column1(1) = ""
    Column1(2) = "Date of test"
    Column1(3) = "Client"
    Column1(4) = "Manufacturer"
    Column1(5) = ""
    Column1(6) = ""
    Column1(7) = "Project name"
    Column1(8) = "Test object"
    Column1(9) = "Variant"
    Column1(10) = "Drawing number"
    Column1(11) = "Kind of object"
    Column1(12) = "U-rated [kV]"
    Column1(13) = "I-sc-rated [kA]"
    Column1(14) = "f-rated [Hz]"
    Column1(15) = "Medium"
    Column1(16) = "Filling Pressure [Ba]"
    Column1(17) = "Rated Pressure [Ba]"
    Column1(18) = ""
    Column1(19) = ""
    Column1(20) = "Kind of test"
    Column1(21) = "Standard"
    Column1(22) = "Test current [kA]"
    Column1(23) = "Test circuit ID"
    Column1(24) = "Revision"
  end sub
  '-------------------------------------------
  
  
  '-------------------------------------------
  sub WriteFrontPage(StartPage,MaxPage,RBAShot)
    dim i
    dim propertyval
    
    'load RBA Data
    Data.Root.Clear()
    
    'Create Frontpage and group
    Call PicLoad(FrontTemplate)
    call GroupCreate(TextGroup, StartPage, false)
    
    call DatafileLoad(RBAShot)
    
    'set column2
    for i = 0 to PropertyCount
      'if Property should be a space, do nothing
      if PropertyName(i) = "" then
        column2(i) = ""
      else
        'check if everything is fine with this property
        if Data.Root.ActiveChannelGroup.Properties.Exists(PropertyName(i)) then
          propertyval =  Data.Root.ActiveChannelGroup.Properties.Item(PropertyName(i)).Value
          if  propertyval = "*****" then
            column2(i)= "-"
          else
            column2(i)= propertyval
          end if
        else
          column2(i)= "-"
        end if
      end if
    next
    
    'field description needs extra handling
    dim descr: descr = column2(0)
    descr = left(descr,instr(descr,"-")-1)  'cut out shotnumber
    descr = mid(descr,2)                  'cut out lab letter
    column2(0) = descr
    
    'field dateTime needs extra handling. remove time string
    column2(2) = Left(Column2(2),10)
    
    'add Channels to group
    call chnalloc(Column1(0),PropertyCount,1,DataTypeString,"Text",StartPage,0)
    call chnalloc(Column2(0),PropertyCount,1,DataTypeString,"Text",StartPage,0)
    
    'add channel values
    for i = 1 to PropertyCount
      CHT(i,TextGroup&"/"&Column1(0)) = Column1(i)
      CHT(i,TextGroup&"/"&Column2(0)) = Column2(i)
    next
    
    Groupdel(StartPage+1)
    
    'FILL TABLE
    Call GraphObjOpen("2DTable1")
      D2TabChnName(1)  = TextGroup&"/"&Column1(0)
      D2TabChnName(2)  = TextGroup&"/"&Column2(0)
    Call GraphObjClose("2DTable1")
    Call GraphObjOpen("Text3")
      TxtTxt = StartPage&" / "&MaxPage
    Call GraphObjClose("Text3")
    Call PicUpdate(0)
    Call GraphSheetNGet(StartPage)
    Call GraphSheetRename(GraphSheetName,"Page " & StartPage)
    Call GraphObjOpen("Text4")
      TxtTxt           = SerialNumber
    Call GraphObjClose("Text4")
  end sub
  '-------------------------------------------

  '-------------------------------------------
  Function GetNumberOfPages()
    GetNumberOfPages = 1
  end Function
  '-------------------------------------------
end class
'-------------------------------------------------------------------------------



'-------------------------------------------------------------------------------
'summary: Creates a Datapages
class DataPage
  dim ChannelGroup
  dim ShotCount
  dim Columncount
  dim RBAShots()
  dim PropertyCount
  dim PropertyHeading()
  dim PropertyName()
  dim PropertyContents()
  dim DataPageHeading
  dim template
  Dim OscNumb
  '-------------------------------------------
  
  
  '-------------------------------------------
  sub initVariables(PageTemplate,OscNumber)
    ChannelGroup = "DataChannels"
    Shotcount = 0
    Columncount = 5
    PropertyCount = 1         'property 1 is always Shotnr, 2,property 0 is empty because of diadem
    Redim RBAShots(0)
    redim PropertyHeading(0)
    redim PropertyName(0)
    redim PropertyContents(0)
    template = PageTemplate
    
    OscNumb = OscNumber
  end sub
  '-------------------------------------------
  
  
  '-------------------------------------------
  sub AddShot(RBAPath)
    Shotcount = Shotcount +1
    redim preserve RBAShots(ShotCount)
    RBAShots(Shotcount) = RBAPath
  end sub
  '-------------------------------------------
  
  
  '-------------------------------------------
  sub AddProperty(PropertyH,PropertyN)
      'there are maximum 22 properties in the table,
      '3 are reserved for shotnr and testresults
      if PropertyCount < 19 then
      Propertycount = Propertycount +1
    
      redim preserve PropertyHeading(Propertycount)
      redim preserve PropertyName(Propertycount)
    
      PropertyHeading(Propertycount) = PropertyH
      PropertyName(Propertycount) = PropertyN
    end if
  end sub
  '-------------------------------------------
  
  
  '-------------------------------------------
  Sub WriteDataPages(StartPage,MaxPage)
    dim i
    dim j
    dim thisGroup
    dim propertyval
    Dim ShotsRead         'is the number of shots read in. if it is as high as column number, Write datapage
    ShotsRead = 1
    thisGroup = ChannelGroup&StartPage
    
    'Add Shotnr property
    PropertyHeading(1) = "Osc - Nr.            " & OscNumb
    PropertyName(1) = "Name"
    
    'add Cycle Property
    Propertycount = Propertycount +1
    redim preserve PropertyHeading(Propertycount):PropertyHeading(Propertycount) = "Cycle"
    redim preserve PropertyName(Propertycount): PropertyName(Propertycount) = "CYCLE_CODE"
    
    'Add Testresult Property
    Propertycount = Propertycount +1
    redim preserve PropertyHeading(Propertycount):PropertyHeading(Propertycount) = "Test Result"
    redim preserve PropertyName(Propertycount): PropertyName(Propertycount) = "REMARKS"
    
    'do nothing when there are no shots. for loop will fall through too.
    if ShotCount > 0 then
    
      'create Channel group and Datachannels for the page
      call GroupCreate(thisGroup, StartPage, false)
      For j = 1 to PropertyCount
        call chnalloc(PropertyHeading(j),ColumnCount,1,DataTypeString,"Text",StartPage,0)
      next
    end if
    
    'fill Datachannels
    'iterate through all shots
    for i = 1 to ShotCount
      call DatafileLoad(RBAshots(i))
      
      'iterate through all properties
      for j = 1 to PropertyCount
        'check if everything is fine with this property
        if Data.Root.ActiveChannelGroup.Properties.Exists(PropertyName(j)) then
          propertyval =  Data.Root.ActiveChannelGroup.Properties.Item(PropertyName(j)).Value
          if  propertyval = "*****" then
            CHT(ShotsRead,thisGroup&"/"&PropertyHeading(j)) = "-"
          else
            CHT(ShotsRead,thisGroup&"/"&PropertyHeading(j)) = propertyval
          end if
        else
           CHT(ShotsRead,thisGroup&"/"&PropertyHeading(j)) = "-" 
        end if
      next
      
      'Shotnr and Test Results need special processing
      Dim Shotnr
      Dim TestResult
      Shotnr = CHT(ShotsRead,thisGroup&"/"&PropertyHeading(1))
      Shotnr = mid(Shotnr,instr(Shotnr,"-")+1)
      Shotnr = left(Shotnr,4)
      TestResult = CHT(ShotsRead,thisGroup&"/"&PropertyHeading(Propertycount))
      
      'if shot passed(P), if shot failed(N),everything else(-)
      Select case(TestResult)
        Case "VALID TEST","GEHALTEN"
          TestResult = "P"
        Case "FAILURE OF TEST PLANT","ANLAGE-FEHLER","TO-PROBLEME","EXPLODIERT","DIELEKTRISCHER VERSAGER","THERMISCHER VERSAGER","SPAETRUECKZUENDUNG","KAPSELUNGS-UEBERSCHLAG","IS-PROBLEM","TR-PROBLEM"
          TestResult = "N"
        Case else
          TestResult = "-"
      end select
      
      CHT(ShotsRead,thisGroup&"/"&PropertyHeading(1)) = Shotnr
      CHT(ShotsRead,thisGroup&"/"&PropertyHeading(Propertycount)) = TestResult
      
      ShotsRead = ShotsRead + 1         'one more shot read in
      Groupdel(StartPage+1)             'remove shot
      
      'if there are as much shots as columns, write datapage
      if Columncount = (ShotsRead-1) then
        dim k
        
        Call WriteToPage(StartPage,Maxpage,thisGroup)
        StartPage = Startpage + 1
        
        'if it wasn't the last page, reset the variables for next datapage
        if i < Shotcount then
          thisGroup = ChannelGroup&StartPage
          
          'create Channel group and Datachannels for the page
          call GroupCreate(thisGroup, StartPage, false)
          For j = 1 to PropertyCount
            call chnalloc(PropertyHeading(j),ColumnCount,1,DataTypeString,"Text",StartPage,0)
          next
        End if
        
        ShotsRead = 1
      end if
    next
    
    'if there are some shots left, write them to datapage
    if Shotsread > 1 then
      call WriteToPage(StartPage,Maxpage,thisGroup)
      StartPage = Startpage + 1
    end if
    
  End Sub
  '-------------------------------------------
  
  
  '-------------------------------------------
  sub WriteToPage(Page,MaxPage,Group)
    dim i
    
    'write to page
    Call PicFileAppend(template)
    Call GraphObjOpen("Text2")
      TxtTxt           = DataPageHeading
    Call GraphObjClose("Text2")
    Call GraphObjOpen("2DTable1")
      for i = 1 to PropertyCount
        D2TabChnName(i)  = Group&"/"& PropertyHeading(i)
      next
    Call GraphObjClose("2DTable1")
    Call GraphObjOpen("Text5")
      TxtTxt           = Page&" / "&MaxPage
    Call GraphObjClose("Text5")
    call Picupdate(0)
    Call GraphSheetNGet(Page)
    Call GraphSheetRename(GraphSheetName,"Page " & Page)
  end sub
  '-------------------------------------------
  
  
  '-------------------------------------------
  'returns the number of pages needed to display
  'all shots
  Function GetNumberOfPages()
    dim shots
    dim pages
    
    shots = Shotcount
    Pages = 0
    
    if shotcount > 0 then
      do
        shots = shots - ColumnCount
        pages = pages + 1
      loop while(shots > 0)
    
      GetNumberOfPages = pages
    
    else
      GetNumberOfPages = 0
    end if
  end Function
  '-------------------------------------------
end class
'-------------------------------------------------------------------------------



'-------------------------------------------------------------------------------
'purpose: holds all information and procedures needed to create a test record of this
'channel
Class Shot
	Dim RBAPath           'Full RBA Filepath
	Dim RBDPath           'Full RBD Filepath
	Dim XChannels()       'X Channel Names, String
	Dim YChannels()       'Y Channel Names, String
  dim ChannelCount      'number of channels
  Dim ErrorOccured      'true if error occured, String
	'-------------------------------------------
  
  
  '-------------------------------------------
	'initialize object
	Sub InitShot(RBAShotPath)
		RBAPath = RBAShotPath
    RBDPath = GetRBD(RBAPath)
    ErrorOccured = "false"
    
    Redim XChannels(0)
    Redim YChannels(0)
    
    'get all channels
    'call GetChannels()
	End Sub
  '-------------------------------------------
  
  
  '-------------------------------------------
  function GetRBD(rbaFile)
    dim rbdFile
    rbdFile = left(rbaFile,InStr(rbaFile, "\report"))
	  rbdFile = rbdFile & mid(rbaFile,InStr(rbaFile, "\report")+8)  'cut out \report
	  rbdFile = left(rbdFile,len(rbdFile)-4) & ".rbd"
    GetRBD = rbdFile
  end function
  '-------------------------------------------
  
  
  '-------------------------------------------
	'get all channels in this shot
	Sub GetChannels()
    Dim i
    Dim j
    
	  data.root.clear()
	  Call Datafileload(RBDPath)

	  if 0 = (Data.Root.ActiveChannelGroup.Channels.Count Mod 2) then
      redim XChannels(Data.Root.ActiveChannelGroup.Channels.Count/2 -1)
      Redim YChannels(Data.Root.ActiveChannelGroup.Channels.Count/2 -1)
      j = 0
      ChannelCount= Data.Root.ActiveChannelGroup.Channels.Count/2 -1
      
      'Add Channels to Variables
		  For i = 1 To Data.Root.ActiveChannelGroup.Channels.Count
		    XChannels(j)= Data.Root.ActiveChannelGroup.Channels(i).Properties.Item("Name").Value
        i = i + 1
        YChannels(j)= Data.Root.ActiveChannelGroup.Channels(i).Properties.Item("Name").Value
        j = j + 1
		  Next
    Else
      Redim XChannels(-1)
      Redim YChannels(-1)
      ErrorOccured = "Couldn't read channels"
    End if
	End Sub
  '-------------------------------------------
  
  
  '-------------------------------------------
  sub GetRBA()
    Call Datafileload(RBAPath)
  End sub
  '-------------------------------------------
  
  
  '-------------------------------------------
  Sub LoadShotEquidistant()
    Dim i
    Dim iXYChNo
    Dim iGroupindex()
    dim group
    
    call GetChannels()
    Redim iXYChNo(ChannelCount,1)
    Redim iGroupindex(ChannelCount)
    
    group = ChnGroup(CNo(Data.Root.ActiveChannelGroup.Name&"/"&XChannels(0)))
    
    for i = 0 to ChannelCount
      iGroupindex(i) = group
      iXYChNo(i,0)  = CNo( Data.Root.ActiveChannelGroup.Name&"/"&xChannels(i))
      iXYChNo(i,1)  = CNo( Data.Root.ActiveChannelGroup.Name&"/"&yChannels(i))
    next
  
    Call equix(iGroupindex,iXYChNo,1,0.0001)
    
    'remove raw group
    Data.Root.ChannelGroups.Remove(iGroupindex(0))
  End sub
  '-------------------------------------------
End Class
'-------------------------------------------------------------------------------



'-------------------------------------------------------------------------------
'summary: Fill Selection into the global class oGUIParam. This cannot be
'         done in the .sud file itself because it doesn't recognize
'         the shot class for an unknown reason.
sub FillSelection(selection)
  dim sel(): redim sel(oGUIParam.SelectionCount)
  dim i
  
  'create new shot object for every selection
  for i = 0 to oGUIParam.SelectionCount
    set sel(i) = new Shot
    
    sel(i).InitShot(selection(i))
  next

  oGUIParam.SelectedShots = Sel
end sub
'-------------------------------------------------------------------------------



'-------------------------------------------------------------------------------
'summary:
sub LoadSingleShot(rbaPath)
  Dim SingleShot
  Set singleShot = New Shot
  
	Call singleShot.InitShot(rbaPath)
  Call singleShot.LoadShotEquidistant() 
End sub
'-------------------------------------------------------------------------------



'-------------------------------------------------------------------------------
'purpose:
class TestRecordPDF
    dim exportPath
    'Dim DataObjectCount
    Dim DataPageObject
    Dim front
    Dim ShotCount
    Dim Serial
    
    'Dim firstShot
    Dim firstshotnr
    Dim lastShotnr
    Dim Testshift()
    Dim Testshiftcount
    
    Dim PictureObjects()
    Dim Picturecount
    '-------------------------------------------
    
    
    '-------------------------------------------
    sub init(Serialnr)
        PictureCount = -1
        redim PictureObjects(0)
        redim Testshift(0)
        Testshiftcount = -1
        
        Serial = Serialnr
        ShotCount = 0
        firstshotnr = ""
        lastShotnr = ""
        
        set DataPageObject = new DataPage:call DataPageObject.initVariables(DataTemplate,Serial)
        set front = new FrontPage:front.initVariables(FrontTemplate)
    end sub
    '-------------------------------------------
    
    
    '-------------------------------------------
    sub SetFilepath(Path)
        exportpath = path
    End sub
    '-------------------------------------------
    
    
    '-------------------------------------------
    Sub setDatatable(Heading,PropName,Count)
        dim i
        
        for i = 0 to count
            call DataPageObject.AddProperty(Heading(i),PropName(i))
        next
    End sub
    '-------------------------------------------
    
    
    '-------------------------------------------
    function getFilePathUser()
        dim returnval: returnval = true
        dim defaultname
        
        defaultname =  Serial&"_"&firstShotnr&"-"&lastshotnr&".pdf"
        If (FileNameGet("ANY", "FileWrite", , "PDF File (*.pdf),"&DefaultName&",All Files (*.*),*.*") = "IDOk") Then
            exportPath =FileDlgName
        else
            returnval = false
        end if
        
        getFilePathUser = returnval
    End Function
    '-------------------------------------------
    
    
    '-------------------------------------------
    Sub AddPicture(Heading,PicturePath)
        PictureCount = Picturecount + 1
        redim preserve PictureObjects(picturecount)
        set PictureObjects(PictureCount) = new PicturePDF
        call PictureObjects(PictureCount).init(pictemplate)
        call PictureObjects(PictureCount).SetPicture(Heading,PicturePath)
    End Sub
    '-------------------------------------------
    
    
    '-------------------------------------------
    Sub AddShot(RBAPath)
        Dim Shotnr
        Shotnr = RBAPath
        Shotnr = mid(Shotnr,instr(Shotnr,"\data\")+6)
	      Shotnr = mid(Shotnr,instr(Shotnr,"\")+1)
	      Shotnr = mid(Shotnr,2)
	      Shotnr = mid(Shotnr,instr(Shotnr,"-")+1)
        Shotnr = left(Shotnr,4)
        
        Testshiftcount = testshiftcount + 1
        Redim Preserve Testshift(testshiftcount)
        Testshift(Testshiftcount) = RBAPath
        DataPageObject.AddShot(RBAPath)
        
        if Testshiftcount = 0 then
            firstShotnr = shotnr
        End if
        
        lastShotnr = shotnr
        SerialNumber = Serial&"_"&firstShotnr&"-"&lastShotnr&".pdf"
    End sub
    '-------------------------------------------
    
    
    '-------------------------------------------
    Sub WritePDF()
        Dim BoxResponse
        Dim currentPage
        Dim MaxPages
        Dim PicturePath
        Dim i

        'set Datapage Heading
        DataPageObject.DataPageHeading = SerialNumber
        'Serial&"/"&firstshotnr&"-"&lastshotnr
        
        'calculate maximum number of Pages
        currentPage = 1
        maxPages = 1
        
        'calculate maximum number of pages
        maxPages = maxPages + Picturecount+1
        maxPages = maxPages + DataPageObject.GetNumberOfPages
        
        'write frontpage
        call front.WriteFrontPage(currentPage,MaxPages,Testshift(0))
        currentPage = currentPage + front.GetNumberOfPages
         
        'write Datapages
        call DataPageObject.WriteDataPages(currentPage,MaxPages)
          
        'write Picturepage
        for i = 0 to Picturecount
            call PictureObjects(i).WritePage(currentPage,MaxPages)
        next
        
        'save pdf and reset
        Call PicPDFExport(exportPath,0)
        Call GraphDeleteAll()
        Data.Root.Clear()
    end sub
    '-------------------------------------------
End Class
'-------------------------------------------------------------------------------



'-------------------------------------------------------------------------------
class PicturePDF
    Dim templatePath
    Dim Heading
    Dim PicturePath
    '-------------------------------------------
    
    
    '-------------------------------------------
    Sub init(templPath)
        templatePath = templPath
        
        Heading = ""
        PicturePath = ""
    End Sub
    '-------------------------------------------
    
    
    '-------------------------------------------
    Sub SetPicture(Pageheading, picPath)
        Heading = Pageheading
        PicturePath= picPath
    End Sub
    '-------------------------------------------
    
    
    '-------------------------------------------
    Sub WritePage(currentPage,MaxPages)
        If PicturePath <> "" Then
            Call PicFileAppend(templatePath)
            Call GraphObjOpen("Metafile1")
                MtaFileName      = PicturePath
            Call GraphObjClose("Metafile1")
            Call GraphObjOpen("Text3")
                TxtTxt           = CurrentPage&" / "&MaxPages
            Call GraphObjClose("Text3")
            Call GraphObjOpen("Text2")
                TxtTxt           = Heading
            Call GraphObjClose("Text2")
            Call PicUpdate(0)
            Call GraphSheetNGet(CurrentPage)
            Call GraphSheetRename(GraphSheetName,"Page " & CurrentPage)
        end if
            
        currentPage = currentPage + 1
    End Sub
    '-------------------------------------------
    
End Class
'-------------------------------------------------------------------------------



'-------------------------------------------------------------------------------
class Standard
    Dim StandardName
    Dim Table()
    Dim TableCount
    
    sub init(DataTable,DataTableCount,DataStandardName)
      dim i
      StandardName = DatastandardName
      TableCount = DataTableCount
      redim Table(Tablecount)
      
      for i = 0 to TableCount
        Table(i) = DataTAble(i)
      Next
    End sub
End Class
'-------------------------------------------------------------------------------








'debug
'-------------------------------------------------------------------------------
sub DebugRecord()
  dim selection()
  dim count
  
  data.Root.Clear()
  Call GraphDeleteAll()
  Globaldim("rootFolder")
		rootFolder = "C:\DIAdem\abbext\TestRecord"
  Globaldim("functionLibrary")
		functionLibrary = "TestRecordFunctions.vbs"
  Globaldim("FrontTemplate")
		FrontTemplate = rootFolder&"\FrontREPORT.TDR"
  Globaldim("DataTemplate")
		DataTemplate = rootFolder&"\ContentREPORT.TDR"  
  Globaldim("PicTemplate")
		PicTemplate = rootFolder&"\PictureREPORT.TDR"
  Globaldim("oGUIParam")
    set oGUIParam = new Reportparameter
    set oGUIParam.oDataFinder  = Navigator.ConnectDataFinder(Navigator.Display.CurrDataProvider.Name)
    oGUIParam.oDataFinder.Results.MaxCount = 1000
    dim sel(0)
    oGuiParam.SelectedShots = sel
    oGuiParam.SelectionCount = count
  Globaldim("Version")
    Version = "V 1.2.2 BETA"
  
  Dim PdfPath
  Dim PDfCount:PDFCount = 1
  redim Pdfpath(PDFCount)
  Dim PdfHeading
  Redim PdfHeading(PDFCount)
  pdfPath(0) = "C:\Documents and Settings\CHJOSCH2\Desktop\imgres.jpg"
  pdfPath(1) = "C:\Documents and Settings\CHJOSCH2\Desktop\imgres.jpg"
  PdfHeading(0) = "Name"
  PdfHeading(1) = "Grleep"
  
  Dim pres
  pres = "\\CH-W-PTHX001\acqui\acqui.l\report\data\"
  
  count = 15
  redim selection(count)
  Selection(0) = pres& "hdiv\l2636-0062.rba"
  Selection(1) = pres& "hdiv\l2636-0063.rba"
  Selection(2) = pres& "hdiv\l2636-0064.rba"
  Selection(3) = pres& "hdiv\l2636-0065.rba"
  Selection(4) = pres& "hdiv\l2636-0066.rba"
  Selection(5) = pres& "hdiv\l2636-0067.rba"
  Selection(6) = pres& "hdiv\l2636-0068.rba"
  Selection(7) = pres& "hdiv\l2636-0069.rba"
  Selection(8) = pres& "hdiv\l2636-0070.rba"
  Selection(9) = pres& "hdiv\l2636-0071.rba"
  Selection(10) = pres& "hdiv\l2636-0072.rba"
  Selection(11) = pres& "hdiv\l2636-0073.rba"
  Selection(12) = pres& "hdiv\l2636-0074.rba"
  Selection(13) = pres& "hdiv\l2636-0075.rba"
  Selection(14) = pres& "hdiv\l2636-0084.rba"
  Selection(15) = pres& "hdiv\l2636-0086.rba"
  
  call FillSelection(selection)
  call ReadConfig(rootFolder&"\TableStandards.txt")
 ' call SearchForTestShifts()

  Dim Table():redim table(2)
  table(0) = "CLIENT"
  table(1) = "MEDIUM"
  Table(2) = "CIRCUIT_CODE"
  Call CreateRecord("2636",Selection,count,table,2,PdfPath,PdfHeading,PDFCount)
end sub
'-------------------------------------------------------------------------------



'-------------------------------------------------------------------------------
class ReportParameter
  Dim oDataFinder               'The DataFinder used (use a Server Edition)
  Dim SelectedShots             'shots array which were selected
  Dim SelectionCount            'Number of Shots
  
  Dim StandardTable             'Standard Datatables for Tabledialog
  Dim StandardCount
  Dim TestShiftStart
  dim TestShiftEnd
  Dim TestShift
  Dim TestShiftCount
  
  Dim TestShiftProperties
  Dim Propertiescount
  
  Dim PDFPath                   'viariables for SetPicture Dialog
  Dim PDFHeading
  Dim PDFCount
end class
'-------------------------------------------------------------------------------



'-------------------------------------------------------------------------------
Sub ReadConfig(cfgfile)
  Dim fso
  Dim txt
  Dim line
  Dim CurrentStandard:CurrentStandard = ""
  Dim CFGStandard
  Dim CFGParameter
  Dim table()
  Dim Tablecount:Tablecount = -1
  Dim Standards()
  Dim StandardCount
  dim debug, i
  
  'init standards
  StandardCount = 0
  redim standards(StandardCount)
  set standards(standardcount) = new standard
  
  set fso = CreateObject("Scripting.FileSystemObject")
  set txt =  fso.OpenTextFile(cfgFile, 1, 0)
  
  'read file
  Do while txt.AtEndOfStream = false
    line = txt.ReadLine
    
    'skip empty and comment lines
    If "#"=left(line,1) Or " " = left(line,1) Then
    
    Elseif line <> "" Then
    	CFGStandard = Left(line,instr(line,",")-1)
    	line = mid(line,instr(line,",")+1)
    	CFGParameter = line
        
        if CurrentStandard = CFGStandard or  CurrentStandard = "" then
            Tablecount = tablecount + 1
            redim preserve table(tablecount)
            Table(tablecount) = CFGParameter
            
            CurrentStandard= CFGStandard
        Else
            Call Standards(Standardcount).init(Table,TableCount,CurrentStandard)
            Standardcount = Standardcount +1
            Redim preserve Standards(Standardcount)
            set Standards(Standardcount) = new Standard
            
            Tablecount = 0
            redim preserve table(tablecount)
            Table(tablecount) = CFGParameter
            CurrentStandard= CFGStandard
        End if
    End If
  loop
  
  Call Standards(Standardcount).init(Table,TableCount,CurrentStandard)
  Standardcount = Standardcount +1
  Redim preserve Standards(Standardcount)
  set Standards(Standardcount) = new Standard
  
  oGuiparam.StandardTable = Standards
  oGuiparam.StandardCount = Standardcount-1

  'add to reportparameters
  txt.Close
  Set fso = Nothing
End Sub

dim tabledialog:tabledialog = "SelectChannels"
'call debugRecord()
'end debug
