Public Sub replace_print(ByRef myarray() As String)
Dim i As Integer
		'Add a template
        Dim wdApp As Word.Application
            
        Set wdApp = GetObject(, "Word.Application")
        'Template here:
        wdApp.Documents.add Template:=""
        'do not open word windows!
        wdApp.ScreenUpdating = False
        
		'Replacements
    
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        
        With Selection.Find
            .Text = "String You Want To Replace"
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
    
        'Save as new file

        'Dim dateTimeNow As String, newFileName As String
        'dateTimeNow = Format(Now(), "yyyy_MM_dd_hh_mm_ss")
        'newFileName = "G:\Information Technology\Abay\Hendrix\card printer\Python\docs\st\" & dateTimeNow & i & ".docx"
        'wdApp.ActiveDocument.SaveAs newFileName
        
        'PRINT
        Dim strPrinter As String
        wdApp.ActivePrinter = "[YOUR PRINTER]"
        ' The original line to print the document
        wdApp.ActiveDocument.PrintOut
End Sub


Sub ReadAsciiFile()

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
        Debug.Print myarray(0), myarray(1), myarray(2), myarray(3), myarray(4), myarray(5), myarray(6), myarray(7), myarray(8), myarray(9), myarray(10); myarray(11)
        Debug.Print "---------------------------------------------------\"
        Debug.Print ""
        Main myarray:=myarray
    Loop

    ' close the file
    Close FileNum

End Sub
