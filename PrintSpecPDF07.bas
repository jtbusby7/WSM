Attribute VB_Name = "PrintSpectPDF0764bit"
Option Explicit
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
   pDevmode As LongPtr               ' Pointer to DEVMODE
   pSepFile As LongPtr
   pPrintProcessor As LongPtr
   pDatatype As LongPtr
   pParameters As LongPtr
   pSecurityDescriptor As LongPtr    ' Pointer to SECURITY_DESCRIPTOR
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
Private Const DM_ORIENTATION = &H1
Private Const DM_PAPERSIZE = &H2
Private Const DM_PAPERLENGTH = &H4
Private Const DM_PAPERWIDTH = &H8
Private Const DM_DEFAULTSOURCE = &H200
Private Const DM_PRINTQUALITY = &H400
Private Const DM_COLOR = &H800
Private Const DM_DUPLEX = &H1000
Private Const DM_IN_BUFFER = 8
Private Const DM_OUT_BUFFER = 2
Private Const PRINTER_ACCESS_USE = &H8
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const PRINTER_NORMAL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_USE)
Private Const PRINTER_ENUM_CONNECTIONS = &H4
Private Const PRINTER_ENUM_LOCAL = &H2
Private Declare PtrSafe Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As LongPtr) As Long
Private Declare PtrSafe Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hwnd As LongPtr, ByVal hPrinter As LongPtr, ByVal pDeviceName As String, ByVal pDevModeOutput As LongPtr, ByVal pDevModeInput As LongPtr, ByVal fMode As Long) As Long
Private Declare PtrSafe Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As LongPtr, ByVal Level As Long, pPrinter As Byte, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Private Declare PtrSafe Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As LongPtr, pDefault As PRINTER_DEFAULTS) As Long
Private Declare PtrSafe Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As LongPtr, ByVal Level As Long, pPrinter As Byte, ByVal Command As Long) As Long
Private Declare PtrSafe Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" (ByVal flags As Long, ByVal name As String, ByVal Level As Long, pPrinterEnum As LongPtr, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare PtrSafe Function PtrToStr Lib "kernel32" Alias "lstrcpyA" (ByVal RetVal As String, ByVal Ptr As LongPtr) As Long
Private Declare PtrSafe Function StrLen Lib "kernel32" Alias "lstrlenA" (ByVal Ptr As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal cbLength As Long)
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, lpOutput As Any, ByVal dev As Long) As Long
Sub MAIN()

    ' PrintSpecPDF07.MAIN Macro
    ' Macro created 12/22/2008 by Brian Mehlferber
    ' Macro updated to 64-bit on 8/26/2024 by todd.busby@gmail.com
    ' Macro updated to use Bluebeam PDF on 8/27/2024 by todd.busby@gmail.com

    '=== Recreate Project Header

    ' Update document statistics
    ActiveDocument.BuiltInDocumentProperties(wdPropertyTitle) = ActiveDocument.BuiltInDocumentProperties(wdPropertyTitle)

    Dim dlg As Object
    Set dlg = ActiveDocument.BuiltInDocumentProperties(wdPropertyTitle)

    ' Combine path of current spec file with contents of projname file
    Dim ProjHeaderFile As String
    ProjHeaderFile = ActiveDocument.Path & Application.PathSeparator & "projname.doc"

    '=== Open header and delete current header
    With ActiveWindow.View
        .Type = wdPrintView
        .ShowHiddenText = True
        .SeekView = wdSeekCurrentPageHeader
    End With

    Selection.WholeStory
    Selection.Delete Unit:=wdCharacter, Count:=1

    '=== Insert file with header info
    Selection.InsertFile FileName:=ProjHeaderFile

    '=== Assure formatting of header
    Selection.WholeStory

    On Error GoTo StyleError
    Selection.Style = ActiveDocument.Styles("JH")
    On Error GoTo 0 ' Reset error handling

    '=== Close header
    ActiveWindow.View.SeekView = wdSeekMainDocument

    '=== Insert Blank Page If Odd Number of Pages
    Dim pageCount As Long
    'pageCount = ActiveDocument.ComputeStatistics(Statistic:=wdStatisticPages)
    pageCount = ActiveDocument.ActiveWindow.Panes(1).Pages.Count

    ' Determine if page count is odd or even
    Dim EvenOdd As String
    If (pageCount Mod 2 = 0) Then
        EvenOdd = "Even"
    Else
        EvenOdd = "Odd"
    End If

    ' If page count is odd insert blank page
    If EvenOdd = "Odd" Then
        Selection.EndKey Unit:=wdStory
        Selection.InsertBreak Type:=wdPageBreak
        'Selection.InsertBreak Type:=wdSectionBreakNextPage
    End If

    '=== Print spec file to PDF
    Dim sCurrentPrinter As String

    ' Save current printer name
    sCurrentPrinter = Application.ActivePrinter

    ' Set printer to Adobe PDF & print the document
    On Error GoTo PrinterError
    Application.ActivePrinter = "Bluebeam PDF"
    On Error GoTo 0 ' Reset error handling

    Dim iColor As Long
    iColor = GetColorMode()
    SetColorMode 1

    ' Print the document
    Application.PrintOut FileName:="", Range:=wdPrintAllDocument, Item:=wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, ManualDuplexPrint:=False, Collate:=True, Background:=True, PrintToFile:=False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, PrintZoomPaperHeight:=0

    SetColorMode iColor

    ' Change printer back to original printer
    Application.ActivePrinter = sCurrentPrinter

    Exit Sub

StyleError:
    MsgBox "Error setting style 'JH': " & Err.Description, vbExclamation
    Resume Next

PrinterError:
    MsgBox "Error setting printer 'Bluebeam PDF': " & Err.Description, vbExclamation
    Application.ActivePrinter = sCurrentPrinter ' Restore original printer
    Resume Next

End Sub

Public Sub SetColorMode(iColorMode As Long)
   SetPrinterProperty DM_COLOR, iColorMode
End Sub
Public Function GetColorMode() As Long
  GetColorMode = GetPrinterProperty(DM_COLOR)
End Function
Public Sub SetDuplex(iDuplex As Long)
   SetPrinterProperty DM_DUPLEX, iDuplex
End Sub

Public Function GetDuplex() As Long
   GetDuplex = GetPrinterProperty(DM_DUPLEX)
End Function
Public Sub SetPrintQuality(iQuality As Long)
   SetPrinterProperty DM_PRINTQUALITY, iQuality
End Sub
Public Function GetPrintQuality() As Long
   GetPrintQuality = GetPrinterProperty(DM_PRINTQUALITY)
End Function
Private Function SetPrinterProperty(ByVal iPropertyType As Long, ByVal iPropertyValue As Long) As Boolean
    'Code adapted from Microsoft KB article Q230743
    Dim hPrinter As LongPtr          ' Handle for the current printer
    Dim pd As PRINTER_DEFAULTS
    Dim pinfo As PRINTER_INFO_2
    Dim dm As DEVMODE
    Dim sPrinterName As String
    Dim yDevModeData() As Byte        ' Byte array to hold contents of DEVMODE structure
    Dim yPInfoMemory() As Byte        ' Byte array to hold contents of PRINTER_INFO_2 structure
    Dim iBytesNeeded As Long
    Dim iRet As Long
    Dim iJunk As Long
    Dim iCount As Long
      
    On Error GoTo cleanup
    ' Get the name of the current printer
    sPrinterName = Trim$(Left$(ActivePrinter, InStr(ActivePrinter, " on ")))

   If (sPrinterName = "") Then
      sPrinterName = ActivePrinter
   End If

    pd.DesiredAccess = PRINTER_ACCESS_USE
    iRet = OpenPrinter(sPrinterName, hPrinter, pd)
    If (iRet = 0) Or (hPrinter = 0) Then
       ' Can't access current printer. Bail out doing nothing
       Exit Function
    End If

    ' Get the size of the DEVMODE structure to be loaded
    iRet = DocumentProperties(0, hPrinter, sPrinterName, 0, 0, 0)
    If (iRet < 0) Then
       ' Can't access printer properties.
       GoTo cleanup
    End If

    ' Make sure the byte array is large enough
    ' Some printer drivers lie about the size of the DEVMODE structure they
    ' return, so an extra 100 bytes is provided just in case!
    ReDim yDevModeData(0 To iRet + 100) As Byte
      
    ' Load the byte array
    iRet = DocumentProperties(0, hPrinter, sPrinterName, _
                VarPtr(yDevModeData(0)), 0, DM_OUT_BUFFER)
    If (iRet < 0) Then
       GoTo cleanup
    End If
    ' Copy the byte array into a structure so it can be manipulated
    Call CopyMemory(dm, yDevModeData(0), Len(dm))
    If (dm.dmFields And iPropertyType) = 0 Then
       ' Wanted property not available. Bail out.
       GoTo cleanup
    End If

    Select Case iPropertyType
    Case DM_ORIENTATION
       dm.dmOrientation = iPropertyValue
    Case DM_PAPERSIZE
       dm.dmPaperSize = iPropertyValue
    Case DM_PAPERLENGTH
       dm.dmPaperLength = iPropertyValue
    Case DM_PAPERWIDTH
       dm.dmPaperWidth = iPropertyValue
    Case DM_DEFAULTSOURCE
       dm.dmDefaultSource = iPropertyValue
    Case DM_PRINTQUALITY
       dm.dmPrintQuality = iPropertyValue
    Case DM_COLOR
       dm.dmColor = iPropertyValue
    Case DM_DUPLEX
       dm.dmDuplex = iPropertyValue
    End Select
      
    ' Load the structure back into the byte array
    Call CopyMemory(yDevModeData(0), dm, Len(dm))

    ' Tell the printer about the new property
    iRet = DocumentProperties(0, hPrinter, sPrinterName, _
          VarPtr(yDevModeData(0)), VarPtr(yDevModeData(0)), _
          DM_IN_BUFFER Or DM_OUT_BUFFER)
    If (iRet < 0) Then
       GoTo cleanup
    End If

    ' The code above *ought* to be sufficient to set the property
    ' correctly. Unfortunately some brands of Postscript printer don't
    ' seem to respond correctly. The following code is used to make
    ' sure they also respond correctly.
    Call GetPrinter(hPrinter, 2, 0, 0, iBytesNeeded)
    If (iBytesNeeded = 0) Then
       ' Couldn't access shared printer settings
       GoTo cleanup
    End If
      
    ' Set byte array large enough for PRINTER_INFO_2 structure
    ReDim yPInfoMemory(0 To iBytesNeeded + 100) As Byte

    ' Load the PRINTER_INFO_2 structure into byte array
    iRet = GetPrinter(hPrinter, 2, yPInfoMemory(0), iBytesNeeded, iJunk)
    If (iRet = 0) Then
       ' Couldn't access shared printer settings
       GoTo cleanup
    End If

    ' Copy byte array into the structured type
    Call CopyMemory(pinfo, yPInfoMemory(0), Len(pinfo))

    ' Load the DEVMODE structure with byte array containing
    ' the new property value
    pinfo.pDevmode = VarPtr(yDevModeData(0))
      
    ' Set security descriptor to null
    pinfo.pSecurityDescriptor = 0
     
    ' Copy the PRINTER_INFO_2 structure back into byte array
    Call CopyMemory(yPInfoMemory(0), pinfo, Len(pinfo))

    ' Send the new details to the printer
    iRet = SetPrinter(hPrinter, 2, yPInfoMemory(0), 0)

    ' Indicate whether it all worked or not!
    SetPrinterProperty = CBool(iRet)

cleanup:
   ' Release the printer handle
   If (hPrinter <> 0) Then Call ClosePrinter(hPrinter)
      
   ' Flush the message queue. If you don't do this,
   ' you can get page fault errors when you try to
   ' print a document immediately after setting a printer property.
   Dim iCounts As Long
   For iCounts = 1 To 20
      DoEvents
   Next iCounts
End Function
Private Function GetPrinterProperty(ByVal iPropertyType As Long) As Long
  ' Code adapted from Microsoft KB article Q230743
  Dim hPrinter As LongPtr
  Dim pd As PRINTER_DEFAULTS
  Dim dm As DEVMODE
  Dim sPrinterName As String
  Dim yDevModeData() As Byte
  Dim iRet As Long
      
  On Error GoTo cleanup
      
  ' Get the name of the current printer
  sPrinterName = Trim$(Left$(ActivePrinter, _
        InStr(ActivePrinter, " on ")))

   If (sPrinterName = "") Then
      sPrinterName = ActivePrinter
   End If
      
  pd.DesiredAccess = PRINTER_ACCESS_USE
      
  ' Get the printer handle
  iRet = OpenPrinter(sPrinterName, hPrinter, pd)
  If (iRet = 0) Or (hPrinter = 0) Then
     ' Couldn't access the printer
      Exit Function
  End If

  ' Find out how many bytes needed for the printer properties
  iRet = DocumentProperties(0, hPrinter, sPrinterName, 0, 0, 0)
  If (iRet < 0) Then
     ' Couldn't access printer properties
      GoTo cleanup
  End If

  ' Make sure the byte array is large enough, including the
  ' 100 bytes extra in case the printer driver is lying.
  ReDim yDevModeData(0 To iRet + 100) As Byte
      
  ' Load the printer properties into the byte array
  iRet = DocumentProperties(0, hPrinter, sPrinterName, _
              VarPtr(yDevModeData(0)), 0, DM_OUT_BUFFER)
  If (iRet < 0) Then
     ' Couldn't access printer properties
     GoTo cleanup
  End If

  ' Copy the byte array to the DEVMODE structure
  Call CopyMemory(dm, yDevModeData(0), Len(dm))

  If Not (dm.dmFields And iPropertyType) = 0 Then
     ' Requested property not available on this printer.
     GoTo cleanup
  End If

  ' Get the value of the requested property
  Select Case iPropertyType
  Case DM_ORIENTATION
     GetPrinterProperty = dm.dmOrientation
  Case DM_PAPERSIZE
     GetPrinterProperty = dm.dmPaperSize
  Case DM_PAPERLENGTH
     GetPrinterProperty = dm.dmPaperLength
  Case DM_PAPERWIDTH
     GetPrinterProperty = dm.dmPaperWidth
  Case DM_DEFAULTSOURCE
     GetPrinterProperty = dm.dmDefaultSource
  Case DM_PRINTQUALITY
     GetPrinterProperty = dm.dmPrintQuality
  Case DM_COLOR
     GetPrinterProperty = dm.dmColor
  Case DM_DUPLEX
     GetPrinterProperty = dm.dmDuplex
  End Select
      
cleanup:
   ' Release the printer handle
   If (hPrinter <> 0) Then Call ClosePrinter(hPrinter)
End Function
