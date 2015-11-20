Public Sub Main(ByRef myarray() As String)
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
            .Text = "<<NAME_1>>"
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
        
        With Selection.Find
            .Text = "<<NAME_2>>"
            .Replacement.Text = myarray(1)
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
        
        With Selection.Find
            .Text = "<<ADDR_1>>"
            .Replacement.Text = myarray(2)
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
        
        With Selection.Find
            .Text = "<<ADDR_2>>"
            .Replacement.Text = myarray(3)
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
            
        With Selection.Find
            .Text = "<<ZIP>>"
            .Replacement.Text = myarray(4)
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
        
        
         With Selection.Find
            .Text = "<<STATE>>"
            .Replacement.Text = myarray(5)
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
        
        
        With Selection.Find
            .Text = "<<CITY>>"
            .Replacement.Text = myarray(6)
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
        
        With Selection.Find
            .Text = "<<GROUPNAME1>>"
            .Replacement.Text = myarray(7)
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
        
         With Selection.Find
            .Text = "<<MEMBERn>>"
            .Replacement.Text = myarray(8)
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
        
         With Selection.Find
            .Text = "<<DIVNAME>>"
            .Replacement.Text = myarray(9)
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
        
         With Selection.Find
            .Text = "<<MS2>>"
            .Replacement.Text = myarray(10)
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
        
         With Selection.Find
            .Text = "<<MS3>>"
            .Replacement.Text = myarray(11)
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
        
         With Selection.Find
            .Text = "<<MS4>>"
            .Replacement.Text = myarray(12)
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
        
         With Selection.Find
            .Text = "<<MS5>>"
            .Replacement.Text = myarray(13)
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
        
         With Selection.Find
            .Text = "<<MS6>>"
            .Replacement.Text = myarray(14)
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
        
        With Selection.Find
            .Text = "<<FN2>>"
            .Replacement.Text = myarray(15)
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
        
         With Selection.Find
            .Text = "<<FN3>>"
            .Replacement.Text = myarray(16)
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
        
         With Selection.Find
            .Text = "<<FN4>>"
            .Replacement.Text = myarray(17)
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
        
         With Selection.Find
            .Text = "<<FN5>>"
            .Replacement.Text = myarray(18)
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
        
        
         With Selection.Find
            .Text = "<<FN6>>"
            .Replacement.Text = myarray(19)
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
        
        
        With Selection.Find
            .Text = "<<MS7>>"
            .Replacement.Text = myarray(20)
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
        
        With Selection.Find
            .Text = "<<FN7>>"
            .Replacement.Text = myarray(21)
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
        
        
         With Selection.Find
            .Text = "<<GROUPn>>"
            .Replacement.Text = myarray(22)
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
        
        
        With Selection.Find
            .Text = "<<COPAY>>"
            .Replacement.Text = myarray(23)
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
        
    
    
        With Selection.Find
            .Text = "<<MS8>>"
            .Replacement.Text = myarray(24)
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
        
        With Selection.Find
            .Text = "<<FN8>>"
            .Replacement.Text = myarray(25)
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
        
        
           With Selection.Find
            .Text = "<<MS9>>"
            .Replacement.Text = myarray(26)
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
        
        With Selection.Find
            .Text = "<<FN9>>"
            .Replacement.Text = myarray(27)
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
        wdApp.ActivePrinter = "\\stlctxprn02p\LRO6055C01"
        ' The original line to print the document
        wdApp.ActiveDocument.PrintOut
End Sub




Sub ReadAsciiFile()

    Dim sFileName As String
    Dim iFileNum As Integer
    Dim sBuf As String
    Dim myarray() As String

    sFileName = ""

    ' does the file exist?  simpleminded test:
    If Len(Dir$(sFileName)) = 0 Then
        Exit Sub
    End If

    iFileNum = FreeFile()
    Open sFileName For Input As iFileNum

    Do While Not EOF(iFileNum)
        Line Input #iFileNum, sBuf
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
    Close iFileNum

End Sub
