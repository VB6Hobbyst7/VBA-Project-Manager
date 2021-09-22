Attribute VB_Name = "mCodePrinter"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Author    Anastasiou Alex
' Project   CodePrinter
' Purpose   Export active project's code as PDF.Code blocks linked by shape. Keywords colored. Oddlines colored.
' Website   https://github.com/alexofrhodes
' Copyright MIT License 2021 Anastasiou Alex
'
' Required References
'   - Microsoft Visual Basic for Application Extensibility
'   - mscorlib.dll
'
' Revision History:
' #  yyyy-mm-dd  COMMENTS
' 1  2021-08-05  Initial Release
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public printFileName As String
Public Found1 As String
Public found2 As String
Dim rng As Range
Public cell As Range
Public s As Shape
Public counter As Long
Dim mafChrWid(32 To 127) As Double
Dim msFontName As String


Public Function PrintProject(wb As Workbook)
    Dim workbookName As String, ModuleName As String, procedure As String
    workbookName = wb.name           'ActiveProjName
    If ProtectedVBProject(wb) = True Or HasProject(wb) = False Then
        MsgBox "Project Empty or Protected"
        Exit Function
    End If
    
    'ThisWorkbook.Application.Visible = False
    ThisWorkbook.IsAddin = False
    
    ModuleName = ActiveComp.name
    Dim vbcomp As VBComponent

    ResetPrinter
    Dim tmpString As Variant
    Dim i As Long
    Dim procedures As Collection
    Set procedures = New Collection
    
    Dim ws As Worksheet
    Dim wsName As String

    
    'Table of contents
    procedures.Add "--- Table Of Contents ---" & vbNewLine & vbNewLine
    'document
    For Each vbcomp In wb.VBProject.VBComponents
        If vbcomp.Type = vbext_ct_Document Then
            For Each ws In wb.Worksheets
                If ws.CodeName = vbcomp.name Then wsName = ws.name
            Next ws
            If vbcomp.name <> "ThisWorkbook" Then
                procedures.Add "(" & ComponentTypeToString(vbcomp.Type) & ")" & " " & wsName & " - " & vbcomp.name
            Else
                procedures.Add "(" & ComponentTypeToString(vbcomp.Type) & ")" & " " & vbcomp.name
            End If
            wsName = ""
        End If
    Next vbcomp
    'class
    For Each vbcomp In wb.VBProject.VBComponents
        If vbcomp.Type = vbext_ct_ClassModule Then
            procedures.Add "(" & ComponentTypeToString(vbcomp.Type) & ")" & " " & vbcomp.name
        End If
    Next vbcomp
    'module
    For Each vbcomp In wb.VBProject.VBComponents
        If vbcomp.Type = vbext_ct_StdModule Then
            procedures.Add "(" & ComponentTypeToString(vbcomp.Type) & ")" & " " & vbcomp.name
        End If
    Next vbcomp
    'userform
    For Each vbcomp In wb.VBProject.VBComponents
        If vbcomp.Type = vbext_ct_MSForm Then
            procedures.Add "(" & ComponentTypeToString(vbcomp.Type) & ")" & " " & vbcomp.name
        End If
    Next vbcomp


    'Code of components
    
    'document
    For Each vbcomp In wb.VBProject.VBComponents
        If vbcomp.Type = vbext_ct_Document Then
            'get sheet name
            For Each ws In wb.Worksheets
                If ws.CodeName = vbcomp.name Then wsName = ws.name
            Next ws
            If vbcomp.name <> "ThisWorkbook" Then
                procedures.Add "--- " & wsName & " - " & vbcomp.name & " ---"
            Else
                procedures.Add "--- " & vbcomp.name & " ---"
            End If
            wsName = ""
            
            If vbcomp.CodeModule.CountOfLines > 0 Then
                tmpString = Split(GetCompText(vbcomp.CodeModule), vbNewLine)
                For i = LBound(tmpString) To UBound(tmpString)
                    procedures.Add " " & tmpString(i)
                Next i
            End If
        End If
    Next vbcomp
    'class
    For Each vbcomp In wb.VBProject.VBComponents
        If vbcomp.Type = vbext_ct_ClassModule Then
            procedures.Add "--- " & vbcomp.name & " ---"
            If vbcomp.CodeModule.CountOfLines > 0 Then
                tmpString = Split(GetCompText(vbcomp.CodeModule), vbNewLine)
                For i = LBound(tmpString) To UBound(tmpString)
                    procedures.Add " " & tmpString(i)
                Next i
            End If
        End If
    Next vbcomp
    'module
    For Each vbcomp In wb.VBProject.VBComponents
        If vbcomp.Type = vbext_ct_StdModule Then
            procedures.Add "--- " & vbcomp.name & " ---"
            If vbcomp.CodeModule.CountOfLines > 0 Then
                tmpString = Split(GetCompText(vbcomp.CodeModule), vbNewLine)
                For i = LBound(tmpString) To UBound(tmpString)
                    procedures.Add " " & tmpString(i)
                Next i
            End If
        End If
    Next vbcomp
    'userform
    For Each vbcomp In wb.VBProject.VBComponents
        If vbcomp.Type = vbext_ct_MSForm Then
            procedures.Add "--- " & vbcomp.name & " ---"
            If vbcomp.CodeModule.CountOfLines > 0 Then
                tmpString = Split(GetCompText(vbcomp.CodeModule), vbNewLine)
                For i = LBound(tmpString) To UBound(tmpString)
                    procedures.Add " " & tmpString(i)
                Next i
            End If
        End If
    Next vbcomp
    
    tmpString = CollectionToArray(procedures)
    ThisWorkbook.Sheets("PRINTER").Range("B1:B" & UBound(tmpString) + 1).Value = WorksheetFunction.transpose(tmpString)
    
    If CodePrinter = False Then GoTo ErrorHandler

    PrintPDF
    
ErrorHandler:
    ThisWorkbook.IsAddin = True
    'ThisWorkbook.Application.Visible = True

End Function

Function CodePrinter() As Boolean
    ThisWorkbook.Sheets("PRINTER").Cells.Font.name = "Consolas"
    RemoveBreaks
    BreakText
    NumberLinesPrinter
    ChgTxtColor
    GreenifyComments
    BoldPrinterComponents
    If findPairs = False Then
        CodePrinter = False
        Exit Function
    End If
    ShapesCompareLeft
    PrinterPageSetup
    ThisWorkbook.Sheets("PRINTER").Rows(1).EntireRow.Insert
    copyLogo
    PageBreaksInPrinter
    CodePrinter = True
End Function

Function findPairs() As Boolean

    Dim ShapeTypeNumber As Long
    ShapeTypeNumber = 29
    Dim CloseTXT As String
    Dim X As Variant
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PRINTER")
    Dim trimCell As String
    
    Dim counter As Long
    For Each cell In ThisWorkbook.Sheets("PRINTER").Range("B:B").SpecialCells(xlCellTypeConstants)
        
        trimCell = Trim(cell.Text)
        If IsBlockStart(trimCell) Then
            Select Case openPair(trimCell)
            Case Is = "Case", "Else"
                GoTo skip
            Case Is = "If", "#If"
                If Right(trimCell, 4) = "Then" Then 'Or Right(trimCell, 1) = "_" Then
                    'ok
                Else
                    GoTo skip
                End If
            Case Is = "skip"
                GoTo skip
            Case Else
                '
            End Select
            CloseTXT = closePair(trimCell)
            counter = Len(cell) - Len(trimCell)
            Found1 = cell.Address
            If FOUND2FOUND(ws, WorksheetFunction.Rept(" ", counter) & CloseTXT) = False Then
                GoTo skip
'                MsgBox "Code not properly indented." & vbNewLine & _
'                       "Error with closing pair of " & vbNewLine & cell.Text
'                findPairs = False
'                Exit Function
            End If
            found2 = ws.Range("B1:B" & ws.Cells(Rows.Count, 2).End(xlUp).Row) _
        .Find(WorksheetFunction.Rept(" ", counter) & CloseTXT & "*", after:=cell, lookat:=xlWhole).Address
            X = StrWidth(Application.WorksheetFunction.Rept("A", counter), "Consolas", 11)
            ws.Shapes.AddShape ShapeTypeNumber, ws.Range(Found1).Left + X - 10, ws.Range(Found1).Top + (cell.Height / 2), 5, Range(Found1, found2).Height - cell.Height
        End If
skip:
    Next cell
    findPairs = True
End Function

Function FOUND2FOUND(ws As Worksheet, str As String) As Boolean
    FOUND2FOUND = True
    Dim TMP As Range
    Set TMP = ws.Range("B1:B" & ws.Cells(Rows.Count, 2).End(xlUp).Row) _
        .Find(str & "*", after:=cell, lookat:=xlWhole)
    If TMP Is Nothing Then FOUND2FOUND = False
End Function

Function IsBlockStart(strLine As String) As Boolean
    Dim bOK As Boolean
    Dim nPos As Integer
    Dim strTemp As String
    nPos = InStr(1, strLine, " ") - 1
    If nPos < 0 Then nPos = Len(strLine)
    strTemp = Left$(strLine, nPos)
    Select Case strTemp
    Case "With", "For", "Do", "While", "Select", "Case", "Else", "Else:", "#Else", "#Else:", "Sub", "Function", "Property", "Enum", "Type"
        bOK = True
    Case "If", "#If", "ElseIf", "#ElseIf"
        bOK = (Len(strLine) = (InStr(1, strLine, " Then") + 4)) Or Right(strLine, 1) = "_"
    Case "Private", "Public", "Friend"
        nPos = InStr(1, strLine, " Static ")
        If nPos Then
            nPos = InStr(nPos + 7, strLine, " ")
        Else
            nPos = InStr(Len(strTemp) + 1, strLine, " ")
        End If
        Select Case Mid$(strLine, nPos + 1, InStr(nPos + 1, strLine, " ") - nPos - 1)
        Case "Sub", "Function", "Property", "Enum", "Type"
            bOK = True
        End Select
    End Select
    IsBlockStart = bOK
End Function

Function IsBlockEnd(strLine As String) As Boolean
    Dim bOK As Boolean
    Dim nPos As Integer
    Dim strTemp As String
    nPos = InStr(1, strLine, " ") - 1
    If nPos < 0 Then nPos = Len(strLine)
    strTemp = Left$(strLine, nPos)
    Select Case strTemp
    Case "Next", "Loop", "Wend", "End Select", "Case", "Else", "#Else", "Else:", "#Else:", "ElseIf", "#ElseIf", "End If", "#End If"
        bOK = True
    Case "End"
        bOK = (Len(strLine) > 3)
    End Select
    IsBlockEnd = bOK
End Function

Function openPair(strLine As String) As String
    Dim nPos As Integer
    Dim strTemp As String
    strTemp = Trim(strLine)
    
    nPos = InStr(1, strTemp, " ") - 1
    If nPos < 0 Then nPos = Len(strLine)
    strTemp = Left$(strLine, nPos)
    
    Select Case strTemp
    Case Is = "Private", "Public"
        strTemp = Trim(strLine)
        strTemp = Replace(strTemp, "Private ", "")
        strTemp = Replace(strTemp, "Public ", "")
        nPos = InStr(1, strTemp, " ") - 1
        If nPos < 0 Then nPos = Len(strTemp)
        strTemp = Left$(strTemp, nPos)
        If strTemp = "Function" Then
            openPair = "Function"
        ElseIf strTemp = "Sub" Then
            openPair = "Sub"
        Else
            GoTo skip
        End If
    Case Is = "With"
        openPair = "With"
    Case Is = "For"
        openPair = "For"
    Case Is = "Do"
        openPair = "Do"
    Case Is = "While"
        openPair = "While"
    Case Is = "Select"
        openPair = "Select"
    Case Is = "Case"
        openPair = "Case"
    Case Is = "Sub"
        openPair = "Sub"
    Case Is = "Function"
        openPair = "Function"
    Case Is = "Property"
        openPair = "Property"
    Case Is = "Enum"
        openPair = "Enum"
    Case Is = "Type"
        openPair = "Type"
    Case "If", "#If"
        openPair = "If"
    Case "ElseIf", "#ElseIf", "Else", "Else:", "#Else", "#Else:"
        openPair = "Else"
    Case Else
skip:
        openPair = "skip"
    End Select

End Function

Function closePair(strLine As String) As String
    Dim nPos As Integer
    Dim strTemp As String
    nPos = InStr(1, strLine, " ") - 1
    If nPos < 0 Then nPos = Len(strLine)
    strTemp = Left$(strLine, nPos)
    Select Case strTemp
    Case Is = "Private", "Public"
        strTemp = Trim(strLine)
        strTemp = Replace(strTemp, "Private ", "")
        strTemp = Replace(strTemp, "Public ", "")
        nPos = InStr(1, strTemp, " ") - 1
        If nPos < 0 Then nPos = Len(strTemp)
        strTemp = Left$(strTemp, nPos)
        If strTemp = "Function" Then
            closePair = "End Function"
        ElseIf strTemp = "Sub" Then
            closePair = "End Sub"
        Else
            '
        End If
    Case Is = "With"
        closePair = "End With"
    Case Is = "For"
        closePair = "Next"
    Case Is = "Do", "While"
        closePair = "Loop"
    Case Is = "Select"                           ', "Case"
        closePair = "End Select"
    Case Is = "Sub"
        closePair = "End Sub"
    Case Is = "Function"
        closePair = "End Function"
    Case Is = "Property"
        closePair = "End Property"
    Case Is = "Enum"
        closePair = "End Enum"
    Case Is = "Type"
        closePair = "End Type"
    Case "If", "#If", "ElseIf", "#ElseIf", "Else", "Else:", "#Else", "#Else:"
        closePair = "End If"
    Case Else
        '
    End Select
End Function

Sub PageBreaksInPrinter()
    ThisWorkbook.Sheets("PRINTER").ResetAllPageBreaks
    Dim rng As Range
    Set rng = Nothing
    Dim cell As Range
    With ThisWorkbook.Sheets("PRINTER")
        For Each cell In .Range("B1:B" & .Range("B" & .Rows.Count).End(xlUp).Row)
            If Left(Trim(cell.Value), 3) = "---" Then
                If rng Is Nothing Then
                    Set rng = cell
                Else
                    Set rng = Union(rng, cell)
                End If
            End If
        Next
        For Each cell In rng
            .HPageBreaks.Add Before:=.Rows(cell.Row)
            .Rows(cell.Row).PageBreak = xlPageBreakManual
        Next
    End With
End Sub

Sub FormatColourFormatters()
    With ThisWorkbook.Sheets("TXTColour")
        ProjectManager.LBLcolourCode.ForeColor = .Range("GeneralFontBackground").Value
        ProjectManager.LBLcolourKey.ForeColor = .Range("ColourKeywords").Value
        ProjectManager.LBLcolourOdd.BackColor = .Range("OddLine").Value
        ProjectManager.LBLcolourComment.ForeColor = .Range("ColourComments").Value
    End With
End Sub

Sub ColorPaletteDialog(rng As Range, lbl As MSForms.Label)
    If Application.Dialogs(xlDialogEditColor).Show(10, 0, 125, 125) = True Then
        'user pressed OK
        Lcolor = ActiveWorkbook.Colors(10)
        rng.Value = Lcolor
        rng.Offset(0, 1).Interior.Color = Lcolor
        lbl.ForeColor = Lcolor
    End If
    ActiveWorkbook.ResetColors
End Sub

Sub RemoveBreaks()
    'remove line break loop
    Dim cell As Range
    Dim rng As Range
    With ThisWorkbook.Sheets("PRINTER")
        Set rng = .Range("B1:B" & .Range("B" & Rows.Count).End(xlUp).Row)
    End With
    Dim coll As Collection
    Set coll = New Collection
    For Each cell In rng
        coll.Add CleanTrim(cell.Value)
    Next cell
    Dim Arr As Variant
    Arr = CollectionToArray(coll)
    rng.Value = WorksheetFunction.transpose(Arr)
End Sub

Function CleanTrim(ByVal s As String, Optional ConvertNonBreakingSpace As Boolean = True) As String
    'remove line break function
    Dim X As Long, CodesToClean As Variant
    CodesToClean = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, _
                         21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 127, 129, 141, 143, 144, 157)
    If ConvertNonBreakingSpace Then s = Replace(s, Chr(160), " ")
    For X = LBound(CodesToClean) To UBound(CodesToClean)
        If InStr(s, Chr(CodesToClean(X))) Then
            s = Replace(s, Chr(CodesToClean(X)), vbNullString)
        End If
    Next
    CleanTrim = s
    '  CleanTrim = WorksheetFunction.Trim(S)
End Function

Sub GreenifyComments()
    Dim cell As Range
    Dim sh As Worksheet
    Set ws = ThisWorkbook.Sheets("PRINTER")
    Set rng = Nothing
    For Each cell In ws.Range("B1:B" & ws.Range("B" & Rows.Count).End(xlUp).Row)
        If Left(Trim(cell.Value), 1) = "'" Or Left(Trim(cell.Value), 3) = "Rem" Then
            cell.Font.Color = ThisWorkbook.Sheets("TXTColour").Range("ColourComments").Value
        End If
    Next
End Sub

Sub SpaceProcsInPrinter()
    'add empty line between end of sub/fun and start of next
    Dim cell As Range
    Dim rng As Range
    With ThisWorkbook.Sheets("PRINTER")
        Set rng = .Range("B2:B" & .Range("B" & Rows.Count).End(xlUp).Row)
    End With
    With rng
        Set cell = .Find("*End Sub", LookIn:=xlValues)
        If Not cell Is Nothing Then
            firstAddress = cell.Address
            Do
                With cell.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .Color = vbBlack
                End With
                Set cell = .FindNext(cell)
            Loop While Not cell Is Nothing And cell.Address <> firstAddress
        End If
    End With
    With rng
        Set cell = .Find("*End Function", LookIn:=xlValues)
        If Not cell Is Nothing Then
            firstAddress = cell.Address
            Do
                With cell.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                Set cell = .FindNext(cell)
            Loop While Not cell Is Nothing And cell.Address <> firstAddress
        End If
    End With
End Sub

Sub NumberLinesPrinter()
    Dim lrow
    Dim cell As Range

    With ThisWorkbook.Sheets("PRINTER")
        lrow = .Range("B" & .Rows.Count).End(xlUp).Row
        For Each cell In .Range("B1:B" & lrow)
            If cell.Row Mod 2 = 0 Then
                Range(cell.Offset(0, -1), cell.Offset(0, 1)).Interior.Color = _
                                                                            ThisWorkbook.Sheets("TXTColour").Range("OddLine").Value
            End If
        Next cell
        '.Columns(1).HorizontalAlignment = xlLeft
    End With
End Sub

Sub BoldPrinterComponents()
    'format printer lines with component names
    Dim rng As Range
    Set rng = Nothing
    Dim cell As Range

    With ThisWorkbook.Sheets("PRINTER")
        For Each cell In .Range("B1:B" & .Range("B" & .Rows.Count).End(xlUp).Row)
            If Left(Trim(cell.Value), 3) = "---" Then
                If rng Is Nothing Then
                    Set rng = cell
                Else
                    Set rng = Union(rng, cell)
                End If
            End If
        Next
    End With
    If rng Is Nothing Then Exit Sub
    rng.Font.Size = 18
    rng.Font.Bold = True
    rng.Font.Color = vbBlack
End Sub

Sub copyLogo()
    ThisWorkbook.Sheets("SETTINGS").Shapes("LOGO").Copy
    ThisWorkbook.Sheets("PRINTER").Paste ThisWorkbook.Sheets("PRINTER").Range("B1")
    Dim shp As Shape
    Set shp = ThisWorkbook.Sheets("PRINTER").Shapes("LOGO")
    With ThisWorkbook.Sheets("PRINTER")
        shp.Left = .Range("B1").Left + ((.Range("B1").Width - shp.Width) / 2)
        shp.Top = .Range("B1").Top
        .Rows(1).RowHeight = shp.Height + 50
        .Range("A2:C2").Interior.ColorIndex = 0
        With .Range("B1")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlVAlignBottom
            .Value = vbNewLine & wb.name & vbNewLine & "www.github.com/alexofrhodes"
            .Characters.Font.Size = 18
            .Characters.Font.Bold = True
            .Characters.Font.Underline = False
            .Characters.Font.ColorIndex = 10
            .Characters.Font.name = "Comic Sans MS"
        End With
    End With
End Sub

Sub PrinterPageSetup()
    With ThisWorkbook.Sheets("PRINTER").PageSetup
        'narrow margins
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.75)
        'left footer filename
        Dim fileName As String
        fileName = printFileName
        .LeftFooter = fileName
        '.LeftFooter =  "&F"    'Filename?
        'center footer page of pages
        .CenterFooter = "Page &P of &N"
        'right footer date
        .RightFooter = "&D"
        'fit all columns in one page width
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
End Sub

Sub ShapesCompareLeft()
    'if code block connector lines spill to the next page,
    'we can easily follow the one we want if each line has it's own colour
    Dim rnd As Long
    Dim n As Variant
    Dim i As Long
    Dim s As Shape
    Dim sNames
    Set sNames = CreateObject("System.Collections.ArrayList")
    'rename lines to their .left position
    For Each s In ThisWorkbook.Sheets("PRINTER").Shapes
        If Left(s.name, 2) <> "logo" Then
            s.name = s.Left
            'create a unique array of names
            If Not sNames.Contains(s.name) Then
                sNames.Add s.name
            End If
        End If
    Next s
    'assign unique colour to lines by level (left)
    For Each n In sNames
        rnd = RandomRGB
        For Each s In ThisWorkbook.Sheets("PRINTER").Shapes
            If s.name = n Then
                With s.Line
                    .ForeColor.RGB = rnd
                    .Weight = 1.5
                End With
            End If
        Next s
    Next n
    Set sNames = Nothing
End Sub

Function RandomRGB()
    RandomRGB = RGB(Int(rnd() * 255), Int(rnd() * 255), Int(rnd() * 255))
End Function

Function StrWidth(s As String, sFontName As String, fFontSize As Double) As Double
    ' Returns the approximate width in points of a text string
    ' in a specified font name and font size
    ' Does not account for kerning
    Dim i As Long
    Dim j As Long
    If Len(sFontName) = 0 Then
        Exit Function
    End If
    If sFontName <> msFontName Then
        If Not InitChrWidths(sFontName) Then
            Exit Function
        End If
    End If
    For i = 1 To Len(s)
        j = Asc(Mid(s, i, 1))
        If j >= 32 Then
            StrWidth = StrWidth + fFontSize * mafChrWid(j)
        End If
    Next i
End Function

Function InitChrWidths(sFontName As String) As Boolean
    Dim i As Long
    Select Case sFontName
    Case "Consolas"
        For i = 32 To 127
            Select Case i
            Case 32 To 127
                mafChrWid(i) = 0.5634
            End Select
        Next i
        '    Case "Arial"
        '        For i = 32 To 127
        '            Select Case i
        '            Case 39, 106, 108
        '                mafChrWid(i) = 0.1902
        '            Case 105, 116
        '                mafChrWid(i) = 0.2526
        '            Case 32, 33, 44, 46, 47, 58, 59, 73, 91 To 93, 102, 124
        '                mafChrWid(i) = 0.3144
        '            Case 34, 40, 41, 45, 96, 114, 123, 125
        '                mafChrWid(i) = 0.3768
        '            Case 42, 94, 118, 120
        '                mafChrWid(i) = 0.4392
        '            Case 107, 115, 122
        '                mafChrWid(i) = 0.501
        '            Case 35, 36, 48 To 57, 63, 74, 76, 84, 90, 95, 97 To 101, 103, 104, 110 To 113, 117, 121
        '                mafChrWid(i) = 0.5634
        '            Case 43, 60 To 62, 70, 126
        '                mafChrWid(i) = 0.6252
        '            Case 38, 65, 66, 69, 72, 75, 78, 80, 82, 83, 85, 86, 88, 89, 119
        '                mafChrWid(i) = 0.6876
        '            Case 67, 68, 71, 79, 81
        '                mafChrWid(i) = 0.7494
        '            Case 77, 109, 127
        '                mafChrWid(i) = 0.8118
        '            Case 37
        '                mafChrWid(i) = 0.936
        '            Case 64, 87
        '                mafChrWid(i) = 1.0602
        '            End Select
        '        Next i
        '
        '    Case "Calibri"
        '        For i = 32 To 127
        '            Select Case i
        '            Case 32, 39, 44, 46, 73, 105, 106, 108
        '                mafChrWid(i) = 0.2526
        '            Case 40, 41, 45, 58, 59, 74, 91, 93, 96, 102, 123, 125
        '                mafChrWid(i) = 0.3144
        '            Case 33, 114, 116
        '                mafChrWid(i) = 0.3768
        '            Case 34, 47, 76, 92, 99, 115, 120, 122
        '                mafChrWid(i) = 0.4392
        '            Case 35, 42, 43, 60 To 63, 69, 70, 83, 84, 89, 90, 94, 95, 97, 101, 103, 107, 118, 121, 124, 126
        '                mafChrWid(i) = 0.501
        '            Case 36, 48 To 57, 66, 67, 75, 80, 82, 88, 98, 100, 104, 110 To 113, 117, 127
        '                mafChrWid(i) = 0.5634
        '            Case 65, 68, 86
        '                mafChrWid(i) = 0.6252
        '            Case 71, 72, 78, 79, 81, 85
        '                mafChrWid(i) = 0.6876
        '            Case 37, 38, 119
        '                mafChrWid(i) = 0.7494
        '            Case 109
        '                mafChrWid(i) = 0.8742
        '            Case 64, 77, 87
        '                mafChrWid(i) = 0.936
        '            End Select
        '        Next i
        '    Case "Tahoma"
        '        For i = 32 To 127
        '            Select Case i
        '            Case 39, 105, 108
        '                mafChrWid(i) = 0.2526
        '            Case 32, 44, 46, 102, 106
        '                mafChrWid(i) = 0.3144
        '            Case 33, 45, 58, 59, 73, 114, 116
        '                mafChrWid(i) = 0.3768
        '            Case 34, 40, 41, 47, 74, 91 To 93, 124
        '                mafChrWid(i) = 0.4392
        '            Case 63, 76, 99, 107, 115, 118, 120 To 123, 125
        '                mafChrWid(i) = 0.501
        '            Case 36, 42, 48 To 57, 70, 80, 83, 95 To 98, 100, 101, 103, 104, 110 To 113, 117
        '                mafChrWid(i) = 0.5634
        '            Case 66, 67, 69, 75, 84, 86, 88, 89, 90
        '                mafChrWid(i) = 0.6252
        '            Case 38, 65, 71, 72, 78, 82, 85
        '                mafChrWid(i) = 0.6876
        '            Case 35, 43, 60 To 62, 68, 79, 81, 94, 126
        '                mafChrWid(i) = 0.7494
        '            Case 77, 119
        '                mafChrWid(i) = 0.8118
        '            Case 109
        '                mafChrWid(i) = 0.8742
        '            Case 64, 87
        '                mafChrWid(i) = 0.936
        '            Case 37, 127
        '                mafChrWid(i) = 1.0602
        '            End Select
        '        Next i
        '    Case "Lucida Console"
        '        For i = 32 To 127
        '            Select Case i
        '            Case 32 To 127
        '                mafChrWid(i) = 0.6252
        '            End Select
        '        Next i
        '
        '    Case "Times New Roman"
        '        For i = 32 To 127
        '            Select Case i
        '            Case 39, 124
        '                mafChrWid(i) = 0.1902
        '            Case 32, 44, 46, 59
        '                mafChrWid(i) = 0.2526
        '            Case 33, 34, 47, 58, 73, 91 To 93, 105, 106, 108, 116
        '                mafChrWid(i) = 0.3144
        '            Case 40, 41, 45, 96, 102, 114
        '                mafChrWid(i) = 0.3768
        '            Case 63, 74, 97, 115, 118, 122
        '                mafChrWid(i) = 0.4392
        '            Case 94, 98 To 101, 103, 104, 107, 110, 112, 113, 117, 120, 121, 123, 125
        '                mafChrWid(i) = 0.501
        '            Case 35, 36, 42, 48 To 57, 70, 83, 84, 95, 111, 126
        '                mafChrWid(i) = 0.5634
        '            Case 43, 60 To 62, 69, 76, 80, 90
        '                mafChrWid(i) = 0.6252
        '            Case 65 To 67, 82, 86, 89, 119
        '                mafChrWid(i) = 0.6876
        '            Case 68, 71, 72, 75, 78, 79, 81, 85, 88
        '                mafChrWid(i) = 0.7494
        '            Case 38, 109, 127
        '                mafChrWid(i) = 0.8118
        '            Case 37
        '                mafChrWid(i) = 0.8742
        '            Case 64, 77
        '                mafChrWid(i) = 0.936
        '            Case 87
        '                mafChrWid(i) = 0.9984
        '            End Select
        '        Next i
    Case Else
        MsgBox "Font name """ & sFontName & """ not available!", vbCritical, "StrWidth"
        Exit Function
    End Select
    msFontName = sFontName
    InitChrWidths = True
End Function

Sub ChgTxtColor()
    OptOn
    With ThisWorkbook.Sheets("PRINTER").Cells.Font
        .Color = ThisWorkbook.Sheets("TXTColour").Range("GeneralFontBackground").Value
        .FontStyle = "Normal"
    End With
    '    Dim l As Long
    '    l = Timer()
    Dim NumChars As Long
    Dim StartChar As Long
    Dim Words
    Dim Word
    Dim WordLength As Long
    Dim rng As Range
    Set rng = Nothing
    Dim cell As Range
    Dim rngKeywords As Range
    Dim rngPrinter As Range
    Dim ArrListKeywords As New ArrayList
    Dim ArrListPrintRows As New ArrayList
    With ThisWorkbook.Sheets("TXTColour")
        Set rngKeywords = .Range("A1:A" & .Range("A" & Rows.Count).End(xlUp).Row)
    End With
    For Each cell In rngKeywords
        ArrListKeywords.Add (cell.Value)
    Next cell
    With ThisWorkbook.Sheets("PRINTER")
        Set rngPrinter = .Range("B1:B" & .Range("B" & Rows.Count).End(xlUp).Row)
    End With
    For Each cell In rngPrinter
        ArrListPrintRows.Add (cell.Value)
    Next cell
    Dim LoopPrinterRows
    Dim LoopKeywords
    Dim c As Range
    Dim firstAddress As String
    Dim counter As Long
    For Each LoopKeywords In ArrListKeywords
        With rngPrinter
            Set c = .Find(LoopKeywords, LookIn:=xlValues, lookat:=xlPart, MatchCase:=True)
            If Not c Is Nothing Then
                If InStrExact(1, c.Text, CStr(LoopKeywords), True) > 0 Then
                    firstAddress = c.Address
                    Do
                        StartChar = InStrExact(1, c.Value, CStr(LoopKeywords))
                        WordLength = Len(LoopKeywords)
                        Do Until StartChar >= Len(c.Value) Or StartChar = 0
                            With c.Characters(start:=StartChar, Length:=WordLength).Font
                                .FontStyle = "Bold"
                                .Color = ThisWorkbook.Sheets("TXTColour").Range("ColourKeywords").Value
                            End With
                            StartChar = InStrExact(StartChar + WordLength, c.Value, CStr(LoopKeywords), True)
                        Loop

                        Set c = .FindNext(c)
                    Loop While Not c Is Nothing And c.Address <> firstAddress
                End If
            End If
        End With
        counter = 0
    Next LoopKeywords
    '    Debug.Print Timer() - l
    OptOff
End Sub

Function InStrExact(start As Long, SourceText As String, WordToFind As String, _
                    Optional CaseSensitive As Boolean = False, _
                    Optional AllowAccentedCharacters As Boolean = False) As Long
    Dim X As Long, Str1 As String, Str2 As String, Pattern As String
    Const UpperAccentsOnly As String = "ÇÉÑ"
    Const UpperAndLowerAccents As String = "ÇÉÑçéñ"
    If CaseSensitive Then
        Str1 = SourceText
        Str2 = WordToFind
        Pattern = "[!A-Za-z0-9]"
        If AllowAccentedCharacters Then Pattern = Replace(Pattern, "!", "!" & UpperAndLowerAccents)
    Else
        Str1 = UCase(SourceText)
        Str2 = UCase(WordToFind)
        Pattern = "[!A-Z0-9]"
        If AllowAccentedCharacters Then Pattern = Replace(Pattern, "!", "!" & UpperAccentsOnly)
    End If
    For X = start To Len(Str1) - Len(Str2) + 1
        If Mid(" " & Str1 & " ", X, Len(Str2) + 2) Like Pattern & Str2 & Pattern _
                                                   And Not Mid(Str1, X) Like Str2 & "'[" & Mid(Pattern, 3) & "*" Then
            InStrExact = X
            Exit Function
        End If
    Next
End Function

Sub ResetPrinter(Optional keepText As Boolean = False)
    '    OptOn
    With ThisWorkbook.Sheets("PRINTER")
        .ResetAllPageBreaks
        If keepText = False Then
            .[A:C].Clear
        Else
            .[A:C].ClearFormats
            .Cells.Font.ColorIndex = vbBlack
            .Cells.Font.Bold = False
        End If
        .Columns("A:A").ColumnWidth = 3          '3
        .Columns("C:C").ColumnWidth = 1
        For Each s In ThisWorkbook.Sheets("PRINTER").Shapes
            'If Left(s.name, 2) <> "cp" Then
            s.Delete
            'End If
        Next
        .Cells.Font.name = "Consolas"
        If .PageSetup.Orientation = xlPortrait Then
            .Columns("B:B").ColumnWidth = 90
        Else
            .Columns("B:B").ColumnWidth = 120
        End If
        .Cells.WrapText = False
        .Cells.UseStandardHeight = True
        '        .Cells.UseStandardWidth = True
    End With
    '    Application.ScreenUpdating = True
End Sub

Sub BreakText()
    'Coded by Anastasiou Alex
    'Version 1
    '20/1/2021
    '    Dim l As Long
    '    l = Timer()
    'to get things right, use a monospace font like Consolas
    Dim cell      As Range
    Dim tmpstr    As String
    Dim Splitter  As Integer
    Dim counter   As Integer
    Dim Limit     As Integer
    'how many characters fit your cell width (find manually)
    If ThisWorkbook.Sheets("PRINTER").PageSetup.Orientation = xlPortrait Then
        Limit = 75
    Else
        Limit = 100                              '80
    End If
    'For which range to run
    Dim rng As Range
    With ThisWorkbook.Sheets("PRINTER")
        Set rng = .Range("B1:B" & .Range("B" & .Rows.Count).End(xlUp).Row)
    End With
    Dim coll As Collection
    Set coll = New Collection
    On Error Resume Next
    For Each cell In rng
        tmpstr = cell.Text
        'remove unnecessary spaces (not trimming)
        If Right(cell.Offset(-1, 0), 1) = "_" Then
            counter = Len(cell.Offset(-1, 0)) - Len(Trim(cell.Offset(-1, 0)))
            tmpstr = Application.WorksheetFunction.Rept(" ", counter) & Trim(cell.Text)
            cell.Value = tmpstr
        End If
        'create collection
        'if len of cell text <= limit then take as is
REPEATME:
        If Len(tmpstr) > Limit Then
            counter = Len(tmpstr) - Len(Trim(tmpstr))
            'if comment
            'BreakText and add first part to collection. Repeat
            If Left(Trim(tmpstr), 1) = "'" Or Left(Trim(tmpstr), 3) = "Rem" Then
                Splitter = Len(cell) / 2
                coll.Add Left(tmpstr, Splitter)  '& " _"
                tmpstr = Application.WorksheetFunction.Rept(" ", counter) & _
                                                                          "'" & Trim(Mid(tmpstr, Splitter + 1))
                GoTo REPEATME
                'if not comment
            Else
                'find which symbol is closest to the limit and before it
                Splitter = InStrRev(tmpstr, WhichFirst(tmpstr, ".`,`/`-` `)", "`", Limit), Limit)
                coll.Add Left(tmpstr, Splitter) & " _"
                tmpstr = Application.WorksheetFunction.Rept(" ", counter) & _
                                                                          Trim(Mid(tmpstr, Splitter + 1))
                GoTo REPEATME
            End If
        Else
            coll.Add (tmpstr)
        End If
    Next cell
    'replace sheet printer cells with broken text from collection
    Dim Arr
    Arr = CollectionToArray(coll)
    With ThisWorkbook.Sheets("PRINTER")
        .Cells.Clear
        .Range("B1:B" & UBound(Arr) + 1).Value = WorksheetFunction.transpose(Arr)
        .Cells.Font.name = "Consolas"
    End With
    '    Debug.Print Timer() - l
End Sub

Sub testWhichFirst()
    'test sub
    If ActiveCell = vbNullString Then
        Exit Sub
    End If
    WhichFirst ActiveCell, ".`,`/`-`_` `)", "`", Len(ActiveCell)
End Sub

Function WhichFirst(st As String, items As String, delim As String, AfterPosition As Integer)
    'Coded by Anastasiou Alex
    'Version 1
    '20/1/2021
    '
    'PARAMETERS
    'st : which string to parse
    'items : which characters are we looking for
    'delim : delimeter to split passed items
    'AfterPosition :
    Dim i As Long
    Dim varr As Variant
    varr = Split(items, delim)
lp:

    On Error Resume Next
    'WhichFirst set to last varr item so it will be looped again?
    WhichFirst = varr(UBound(varr))
    For i = LBound(varr) To UBound(varr)
        'Debug.Print varr(i) & InStrRev(st, varr(i), AfterPosition)
        'find the item closest to the limit
        If InStrRev(st, varr(i), AfterPosition) > InStrRev(st, WhichFirst, AfterPosition) Then
            WhichFirst = varr(i)
        End If
    Next i
    '    Debug.Print "Limit", AfterPosition & vbNewLine & _
    "Closest Item", WhichFirst & vbNewLine & _
    "Found At", InStrRev(st, WhichFirst, AfterPosition)
End Function

Function HasProject(wb As Workbook) As Boolean
    Dim WbProjComp As Object
    On Error Resume Next
    Set WbProjComp = wb.VBProject.VBComponents
    If Not WbProjComp Is Nothing Then HasProject = True
End Function

Function ProtectedVBProject(ByVal wb As Workbook) As Boolean
    ' returns TRUE if the VB project in the active document is protected
    If wb.VBProject.Protection = 1 Then
        ProtectedVBProject = True
    Else
        ProtectedVBProject = False
    End If
End Function

Sub OptOn()
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ' Note: this is a sheet-level setting.
    ActiveSheet.DisplayPageBreaks = False
End Sub

Sub OptOff()
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    ' Note: this is a sheet-level setting.
    ActiveSheet.DisplayPageBreaks = False
End Sub

Public Function ActiveProjName() As String
    'name of active project in vbeditor
    ActiveProjName = Mid(Application.VBE.ActiveVBProject.fileName, InStrRev(Application.VBE.ActiveVBProject.fileName, "\") + 1)
End Function

Function ActiveComp() As VBComponent
    'name of component where mouse is inside of
    Set ActiveComp = Application.VBE.SelectedVBComponent
End Function


Function CollectionToArray(c As Collection) As Variant
    Dim a() As Variant: ReDim a(0 To c.Count - 1)
    Dim i As Long
    For i = 1 To c.Count
        a(i - 1) = c.Item(i)
    Next
    CollectionToArray = a
End Function


Sub PrintPDF()

    Dim filePath As String
    filePath = Environ("USERprofile") & "\Documents\" & "vbaCodeArchive\CodePrinter\"
    Dim fileName As String
    fileName = Left(wb.name, InStr(1, wb.name, ".") - 1)
    Dim saveLocation As String
    saveLocation = filePath                      '& fileName & "\"
    If Dir(saveLocation, vbDirectory) = "" Then
        FoldersCreate saveLocation
    End If
    filePath = saveLocation & fileName
    ThisWorkbook.Sheets("PRINTER").ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=filePath
        
    'FollowLink saveLocation
    'FollowLink filePath & ".pdf"
End Sub


Public Sub delay(seconds As Long)
    Dim endTime As Date
    endTime = DateAdd("s", seconds, Now())
    Do While Now() < endTime
        DoEvents
    Loop
End Sub


