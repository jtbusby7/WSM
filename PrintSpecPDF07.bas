Sub MAIN()
'
' PrintSpecPDF07.MAIN Macro
' Macro created 12/22/2008 by Brian Mehlferber
'
'===Recreate Project Header
WordBasic.FileSummaryInfo Update:=1     'Update document statistics
Dim dlg As Object: Set dlg = WordBasic.DialogRecord.FileSummaryInfo(False)  'Create a dialog record
WordBasic.CurValues.FileSummaryInfo dlg 'Fill record with values
'===Combine path of current spec file with contents of projname file
ProjHeaderFile$ = ActiveDocument.Path & Application.PathSeparator & "projname.doc"
'
'===Open header and delete current header
    WordBasic.ViewHeader
    Selection.WholeStory
    Selection.Delete Unit:=wdCharacter, Count:=1
    ''WordBasic.StartOfDocument
    ''
    ''WordBasic.EndOfLine 1
    ''WordBasic.EditClear
    '
'===Insert file with header info
WordBasic.InsertFile name:=ProjHeaderFile$
'
'===Assure formatting of header
    Selection.WholeStory
    Selection.Style = ActiveDocument.Styles("JH")
    ''WordBasic.CharLeft 1
    ''WordBasic.StartOfLine 1
    ''WordBasic.Style "JH"
    '
'===Close header
WordBasic.CloseViewHeaderFooter
'
'===Insert Blank Page If Odd Number of Pages
    '
    'count number of pages in document
    Dim pageCount As Long
    pageCount = ActiveDocument.ComputeStatistics(Statistic:=wdStatisticPages)
    '
    'determine if page count is odd or even
    Dim EvenOdd As String
    If (pageCount Mod 2 = 0) Then
    EvenOdd = "Even"
    Else
    EvenOdd = "Odd"
    End If
    'if page count is odd insert blank page
    If EvenOdd = "Odd" Then
    Selection.EndKey Unit:=wdStory
    Selection.InsertBreak Type:=wdPageBreak
    Else
    End If
    '
'===Print spec file to PDF
    Dim sCurrentPrinter As String
    'save current printer name
    sCurrentPrinter = Application.ActivePrinter
    'set printer to Adode PDF & print the document
    Application.ActivePrinter = "Adobe PDF"
    'set printer to black & white
    Dim iColor As Long
    iColor = GetColorMode
    If iColor = 2 Then
       SetColorMode 1
    Else
    End If
    '
        Application.PrintOut FileName:="", Range:=wdPrintAllDocument, Item:= _
        wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
        ManualDuplexPrint:=False, Collate:=True, Background:=True, PrintToFile:= _
        False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
        PrintZoomPaperHeight:=0
    'set printer back to color if needed
    If iColor = 2 Then
    SetColorMode 2
    Else
    End If
    'change printer back to original printer
    Application.ActivePrinter = sCurrentPrinter

End Sub


