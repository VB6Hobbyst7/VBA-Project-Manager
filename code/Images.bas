Attribute VB_Name = "Images"
'Similar to jaslake,https://www.excelforum.com/excel-programming-vba-macros/1202015-print-userform-to-pdf-and-then-attach-it-to-an-email.html
#If VBA7 Then
    Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, _
                                                  ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As LongPtr)
#Else
    Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, _
                                          ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
#End If
Private Const VK_SNAPSHOT = 44
Private Const VK_LMENU = 164
Private Const KEYEVENTF_KEYUP = 2
Private Const KEYEVENTF_EXTENDEDKEY = 1

Sub UserformToPDF(wb As Workbook, path As String)
    Application.VBE.MainWindow.Visible = True
    Do While Application.VBE.MainWindow.Visible = False
        DoEvents
    Loop
    CloseVBEwindows
    Dim vbcomp As VBComponent
    For Each vbcomp In wb.VBProject.VBComponents
        If vbcomp.Type = vbext_ct_MSForm Then
            vbcomp.Activate
            DoEvents
            'Application.Wait (Now + TimeValue("0:00:3"))
            Call WindowToPDF(PathMaker(path, vbcomp.name, "pdf"))
        End If
    Next
End Sub

Sub ExportWorksheetsToPDF(wb As Workbook, expPath As String)
    wb.Activate
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Application.PrintCommunication = False
        With ws.PageSetup
            .FitToPagesWide = 1
            .FitToPagesTall = False
        End With
        Application.PrintCommunication = True
        If WorksheetFunction.CountA(ws.Cells) > 0 Then
            ws.ExportAsFixedFormat xlTypePDF, PathMaker(expPath, ws.name, "pdf"), , True
        End If
    Next ws
End Sub

Function WindowToPDF(pdf$, Optional Orientation As Integer = xlLandscape, _
                     Optional FitToPagesWide As Integer = 1) As Boolean
    Dim calc As Integer, ws As Worksheet
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        calc = .Calculation
        .Calculation = xlCalculationManual
    End With
    keybd_event VK_LMENU, 0, KEYEVENTF_EXTENDEDKEY, 0
    keybd_event VK_SNAPSHOT, 0, KEYEVENTF_EXTENDEDKEY, 0
    keybd_event VK_SNAPSHOT, 0, KEYEVENTF_EXTENDEDKEY + KEYEVENTF_KEYUP, 0
    keybd_event VK_LMENU, 0, KEYEVENTF_EXTENDEDKEY + KEYEVENTF_KEYUP, 0
    Set ws = Workbooks.Add(xlWBATWorksheet).Worksheets(1)
    Application.Wait (Now + TimeValue("0:00:1"))
    With ws
        .PasteSpecial Format:="Bitmap", Link:=False, DisplayAsIcon:=False
        .Range("A1").Select
        .PageSetup.Orientation = Orientation
        .PageSetup.FitToPagesWide = FitToPagesWide
        .PageSetup.Zoom = False
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdf, _
                             Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                             IgnorePrintAreas:=False, OpenAfterPublish:=False
        .Parent.Close False
    End With
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = calc
        .CutCopyMode = False
    End With
    WindowToPDF = Dir(pdf) <> ""
End Function

Sub DebugPrintWindows()
    Dim wd As Window
    Dim i As Long
    For i = 1 To Application.VBE.Windows.Count
        With Application.VBE.Windows(i)
            Debug.Print .Type & vbTab & .Caption
        End With
    Next
End Sub

Sub CloseVBEwindows()
    Dim i As Long
    For i = 1 To Application.VBE.Windows.Count
        Select Case Application.VBE.Windows(i).Type
        Case 2 To 7
            Application.VBE.Windows(i).Close
        End Select
    Next
End Sub

Function PathMaker(wbPath As String, fileName As String, fileExtention As String) As String
    If Right(wbPath, 1) <> "\" Then wbPath = wbPath & "\"
    PathMaker = wbPath & fileName & "." & fileExtention
    Do While InStr(1, PathMaker, "..") > 0
        PathMaker = Replace(PathMaker, "..", ".")
    Loop
End Function


