Option Explicit

' === API Types for 64-Bit Printer Control ===
Private Type PRINTER_DEFAULTS
    pDatatype As LongPtr
    pDevmode As LongPtr
    DesiredAccess As Long
End Type

Private Type PRINTER_INFO_2
    pServerName As LongPtr
    pPrinterName As LongPtr
    pShareName As LongPtr
    pPortName As LongPtr
    pDriverName As LongPtr
    pComment As LongPtr
    pLocation As LongPtr
    pDevmode As LongPtr
    pSepFile As LongPtr
    pPrintProcessor As LongPtr
    pDatatype As LongPtr
    pParameters As LongPtr
    pSecurityDescriptor As LongPtr
    Attributes As Long
    Priority As Long
    DefaultPriority As Long
    StartTime As Long
    UntilTime As Long
    Status As Long
    cJobs As Long
    AveragePPM As Long
End Type

Private Type DEVMODE
    dmDeviceName As String * 32
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * 32
    dmUnusedPadding As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    dmICMMethod As Long
    dmICMIntent As Long
    dmMediaType As Long
    dmDitherType As Long
    dmReserved1 As Long
    dmReserved2 As Long
End Type

' === Constants ===
Private Const DM_COLOR = &H800
Private Const DM_OUT_BUFFER = 2
Private Const DM_IN_BUFFER = 8
Private Const PRINTER_ACCESS_USE = &H8

' === 64-Bit API Declarations ===
Private Declare PtrSafe Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As LongPtr) As Long
Private Declare PtrSafe Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hwnd As LongPtr, ByVal hPrinter As LongPtr, ByVal pDeviceName As String, ByVal pDevModeOutput As LongPtr, ByVal pDevModeInput As LongPtr, ByVal fMode As Long) As Long
Private Declare PtrSafe Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As LongPtr, ByVal Level As Long, pPrinter As Byte, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Private Declare PtrSafe Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As LongPtr, pDefault As PRINTER_DEFAULTS) As Long
Private Declare PtrSafe Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As LongPtr, ByVal Level As Long, pPrinter As Byte, ByVal Command As Long) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal cbLength As Long)

' =========================================================
' MAIN MACRO
' =========================================================
Sub MAIN()
    Dim sCurrentPrinter As String
    Dim ProjHeaderFile As String
    Dim iColor As Long
    Dim tbl As Table
    
    On Error GoTo ErrorHandler
    
    ' 1. Refresh Document Properties
    ActiveDocument.BuiltInDocumentProperties(wdPropertyTitle) = ActiveDocument.BuiltInDocumentProperties(wdPropertyTitle)
    ProjHeaderFile = ActiveDocument.Path & Application.PathSeparator & "projname.doc"

    ' 2. "SEARCH AND DESTROY" - Fix for Office 365 (Build 2601)
    ' Physically remove hidden text so modern Word engines cannot render it
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Font.Hidden = True
        .Text = ""
        .Replacement.Text = ""
        .Execute Replace:=wdReplaceAll
    End With

    ' 3. Clean Header and Strip Hidden Fields
    With ActiveWindow.View
        .Type = wdPrintView
        .ShowHiddenText = False 
        .SeekView = wdSeekCurrentPageHeader
    End With

    Selection.WholeStory
    Selection.Delete Unit:=wdCharacter, Count:=1

    ' Apply style to empty header
    On Error Resume Next
    Selection.Style = ActiveDocument.Styles("JH")
    On Error GoTo ErrorHandler

    ' Insert header file
    Selection.InsertFile FileName:=ProjHeaderFile

    ' 4. "BOX KILLER" - Fix for AIA MasterSpec Tables
    ' Strips borders from MasterSpec instructional containers in the header
    Selection.WholeStory
    For Each tbl In Selection.Tables
        tbl.Borders.Enable = False
        tbl.Rows.LeftIndent = 0
    Next tbl
    
    ' Force paragraph indents and borders to zero
    With Selection.ParagraphFormat
        .Borders.Enable = False
        .LeftIndent = 0
    End With

    ActiveWindow.View.SeekView = wdSeekMainDocument

    ' 5. Handle Even Page Logic
    Options.PrintHiddenText = False
    ActiveDocument.Repaginate
    
    If (ActiveDocument.ComputeStatistics(wdStatisticPages) Mod 2 <> 0) Then
        Selection.EndKey Unit:=wdStory
        Selection.InsertBreak Type:=wdPageBreak
    End If

    ' 6. Printer Selection and Execution
    sCurrentPrinter = Application.ActivePrinter

    With Application.Dialogs(wdDialogFilePrint)
        If .Display = -1 Then
            iColor = GetColorMode()
            SetColorMode 1 ' Force Color
            
            .Execute
            
            ' Restore Color Mode
            SetColorMode iColor
        Else
            MsgBox "Printing cancelled.", vbInformation
        End If
    End With

    ' Restore original printer
    Application.ActivePrinter = sCurrentPrinter
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

' =========================================================
' PRINTER UTILITY FUNCTIONS
' =========================================================

Public Sub SetColorMode(iColorMode As Long)
    SetPrinterProperty DM_COLOR, iColorMode
End Sub

Public Function GetColorMode() As Long
    GetColorMode = GetPrinterProperty(DM_COLOR)
End Function

Private Function SetPrinterProperty(ByVal iPropertyType As Long, ByVal iPropertyValue As Long) As Boolean
    Dim hPrinter As LongPtr, pd As PRINTER_DEFAULTS, pinfo As PRINTER_INFO_2, dm As DEVMODE
    Dim sPrinterName As String, yDevModeData() As Byte, yPInfoMemory() As Byte
    Dim iBytesNeeded As Long, iRet As Long, iJunk As Long, iCount As Long
      
    On Error GoTo cleanup
    sPrinterName = Trim$(Left$(ActivePrinter, InStr(ActivePrinter, " on ")))
    If (sPrinterName = "") Then sPrinterName = ActivePrinter

    pd.DesiredAccess = PRINTER_ACCESS_USE
    iRet = OpenPrinter(sPrinterName, hPrinter, pd)
    If (iRet = 0) Or (hPrinter = 0) Then Exit Function

    iRet = DocumentProperties(0, hPrinter, sPrinterName, 0, 0, 0)
    ReDim yDevModeData(0 To iRet + 100) As Byte
    iRet = DocumentProperties(0, hPrinter, sPrinterName, VarPtr(yDevModeData(0)), 0, DM_OUT_BUFFER)
    
    Call CopyMemory(dm, yDevModeData(0), Len(dm))
    dm.dmColor = iPropertyValue
    Call CopyMemory(yDevModeData(0), dm, Len(dm))
    
    iRet = DocumentProperties(0, hPrinter, sPrinterName, VarPtr(yDevModeData(0)), VarPtr(yDevModeData(0)), DM_IN_BUFFER Or DM_OUT_BUFFER)

    Call GetPrinter(hPrinter, 2, 0, 0, iBytesNeeded)
    ReDim yPInfoMemory(0 To iBytesNeeded + 100) As Byte
    iRet = GetPrinter(hPrinter, 2, yPInfoMemory(0), iBytesNeeded, iJunk)
    
    Call CopyMemory(pinfo, yPInfoMemory(0), Len(pinfo))
    pinfo.pDevmode = VarPtr(yDevModeData(0))
    pinfo.pSecurityDescriptor = 0
    Call CopyMemory(yPInfoMemory(0), pinfo, Len(pinfo))

    iRet = SetPrinter(hPrinter, 2, yPInfoMemory(0), 0)
    SetPrinterProperty = CBool(iRet)

cleanup:
    If (hPrinter <> 0) Then Call ClosePrinter(hPrinter)
    For iCount = 1 To 20: DoEvents: Next iCount
End Function

Private Function GetPrinterProperty(ByVal iPropertyType As Long) As Long
    Dim hPrinter As LongPtr, pd As PRINTER_DEFAULTS, dm As DEVMODE, sPrinterName As String
    Dim yDevModeData() As Byte, iRet As Long
      
    On Error GoTo cleanup
    sPrinterName = Trim$(Left$(ActivePrinter, InStr(ActivePrinter, " on ")))
    If (sPrinterName = "") Then sPrinterName = ActivePrinter
      
    pd.DesiredAccess = PRINTER_ACCESS_USE
    iRet = OpenPrinter(sPrinterName, hPrinter, pd)
    If (iRet = 0) Or (hPrinter = 0) Then Exit Function

    iRet = DocumentProperties(0, hPrinter, sPrinterName, 0, 0, 0)
    ReDim yDevModeData(0 To iRet + 100) As Byte
    iRet = DocumentProperties(0, hPrinter, sPrinterName, VarPtr(yDevModeData(0)), 0, DM_OUT_BUFFER)

    Call CopyMemory(dm, yDevModeData(0), Len(dm))
    Select Case iPropertyType
        Case DM_COLOR: GetPrinterProperty = dm.dmColor
    End Select
      
cleanup:
    If (hPrinter <> 0) Then Call ClosePrinter(hPrinter)
End Function
