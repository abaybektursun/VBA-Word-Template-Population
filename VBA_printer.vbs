Public Sub Replace_print(ByRef myarray() As String)
Dim i As Integer
		'Add a template
        Dim wdApp As Word.Application
            
        Set wdApp = GetObject(, "Word.Application")
		
        'Template here:
        wdApp.Documents.add Template:="YOUR TEMPLATE"
		
        'Do not open word windows
        wdApp.ScreenUpdating = False
        wdApp.Visible = False
        wdApp.DisplayAlerts = wdAlertsNone
        
		'Replacements
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        
        With Selection.Find
            .Text = "[String You Want To Replace]"
            .Replacement.Text = myarray(0)
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
    
        'You can save the file if you want:
		
        'Dim dateTimeNow As String, newFileName As String
        'dateTimeNow = Format(Now(), "yyyy_MM_dd_hh_mm_ss")
        'newFileName = "DOC" & dateTimeNow & i & ".docx"
        'wdApp.ActiveDocument.SaveAs newFileName
        
        'Set The Printer
        Dim strPrinter As String
        wdApp.ActivePrinter = "[YOUR PRINTER]"
		
        'Print the document
        wdApp.ActiveDocument.PrintOut
		'Keep the project derectory clean be closing the project
		With wdApp
            'Loop Through open documents
            Do Until .Documents.Count = 0
                'Close no save
                .Documents(1).Close SaveChanges:=wdDoNotSaveChanges
            Loop
        End With
End Sub


Sub Main()

    Dim FileName As String
    Dim FileNum As Integer
    Dim sBuf As String
    Dim myarray() As String

    FileName = ""

    ' does the file exist?  simpleminded test:
    If Len(Dir$(FileName)) = 0 Then
        Exit Sub
    End If

    FileNum = FreeFile()
    Open FileName For Input As FileNum

    Do While Not EOF(FileNum)
        Line Input #FileNum, sBuf
        ' now you have the next line of the file in sBuf
        ' do something useful:
        myarray = Split(sBuf, ",")
        Debug.Print ""
        Debug.Print "/---------------------------------------------------"
        Debug.Print myarray(0), myarray(1), myarray(2), myarray(3), myarray(4)
        Debug.Print "---------------------------------------------------\"
        Debug.Print ""
        replace_print myarray:=myarray
    Loop

    ' close the file
    Close FileNum

End Sub
