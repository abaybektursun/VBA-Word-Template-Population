Public Sub printer(ByRef dataArray() As String, ByRef fieldsArray() As String, Document As String, printer_name As String)
Dim i As Integer
    'Add a template
    Dim wdApp      As Word.Application
    Dim data_index As Integer
        
    Set wdApp = GetObject(, "Word.Application")
    
    wdApp.Documents.add Template:=Document
    'do not open word windows!
    wdApp.ScreenUpdating = False
    wdApp.Visible = False
    'Turn off DisplayAlerts
    wdApp.DisplayAlerts = wdAlertsNone


    
    'Replacements
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    
    data_index = -1
    Dim element As Variant
    For Each element In fieldsArray
        data_index = data_index + 1
        
        With Selection.Find
            .Text = CStr(element)
            .Replacement.Text = dataArray(data_index)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
    Next element
        
    
    'Save as new file
    'newFileName = ""
    'wdApp.ActiveDocument.SaveAs newFileName
    
    'PRINT
    Dim strPrinter As String
    
    wdApp.ActivePrinter = printer_name
    ' Print, Background print must be turned off to prevent messages about margins
    wdApp.PrintOut Background:=False
    
    With wdApp
        'Loop Through open documents
        Do Until .Documents.Count = 0
            'Close no save
            .Documents(1).Close SaveChanges:=wdDoNotSaveChanges
        Loop
    End With

End Sub

Sub VBA_printer()

    Dim fileName      As String
    Dim fileNum       As Integer
    Dim fileBuf       As String
    Dim dataArray()   As String
    Dim fieldsArray() As String
    Dim firstRecord   As Boolean
    Dim tempDocName   As String
    Dim docPath       As String
    Dim dataPath      As String
    Dim printer_name  As String
    
    ' ! Specify path to the template document
    docPath     = ""
    ' ! Specify path to the data
    dataPath     = ""
    ' ! Specify data file name
    fileName = ""
    ' ! Specify template document Name
    tempDocName  = ""
    ' ! Specify printer
    printer_name  = ""

    ' ! These are the Strings that will be replaced in the document
    ' fieldsArray needs to be in order coresponding to the dataArray
    fieldsArray = Split("", ",")
    ' Example:
    'fieldsArray = Split("<<PROVIDER_NBR>>,<<IRS_NBR>>,<<NPI_NBR>>,<<APP_YYYYMMDD>>,<<PREV_YYYYMMDD>>,<<YMD_RECRED>>,<<Title>>,<<FirstName>>,<<LASTNAME>>,<<Letter_LastName>>,<<DOB>>,<<DEGREE>>,<<DEA_NUMBER>>,<<LICENSE_NBR>>,<<OFFICE_NBR>>,<<CERTIFICATION>>,<<HOSP_PRIV>>,<<LISTING>>,<<PHO_AFFILIATION>>,<<PADDRESS>>,<<PCITY>>,<<PSTATE>>,<<PZIP>>,<<PPHONE_NBR>>,<<PFAX_NBR>>,<<PEMAIL>>,<<PMANAGER>>,<<PNAME>>,<<BADDRESS>>,<<BCITY>>,<<BSTATE>>,<<BZIP>>,<<BPHONE_NBR>>", ",")
    
    ' Does the file exist?
    If Len(Dir$(fileName)) = 0 Then
        Exit Sub
    End If

    fileNum = FreeFile()
    Open fileName For Input As fileNum
    
    ' We are assuming that the first row in the data file is column names, so we ignore it
    firstRecord = True
    Do While Not EOF(fileNum)
        Line Input #fileNum, fileBuf
        ' ! Watch out for the order of the elemnts, need to corespond to fieldsArray
        dataArray = Split(fileBuf, "~")
        If Not firstRecord Then
            printer dataArray:=dataArray, fieldsArray:=fieldsArray, Document:=docPath & tempDocName, printer_name:=printer_name
        End If
        firstRecord = False
    Loop
End Sub

