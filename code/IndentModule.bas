Attribute VB_Name = "IndentModule"

Public Function IndentWorkbook(fromWorkbook As Workbook)
    Dim nIndent As Integer
    Dim nLine As Long
    Dim strNewLine As String
    Dim vbcomp As VBComponent
    
    Dim proceduresCollection As Collection
    Dim procedure As Variant
    
    For Each vbcomp In fromWorkbook.VBProject.VBComponents
        Set proceduresCollection = ProcList(vbcomp.CodeModule)
        For Each procedure In proceduresCollection
            IndentProcedure vbcomp, CStr(procedure)
        Next procedure
    Next vbcomp
End Function
Sub IndentProcedure(vbcomp As VBComponent, procName As String)
    Dim nIndent As Integer
    Dim nLine As Long
    Dim strNewLine As String
    Dim startLine As Long
    startLine = ProcedureStartLine(vbcomp.CodeModule, procName)
    Dim endLine As Long
    endLine = ProcedureEndLine(vbcomp.CodeModule, procName)
    For nLine = startLine To endLine
        ' Get next line.
        strNewLine = vbcomp.CodeModule.Lines(nLine, 1)
        ' Remove leading space.
        strNewLine = LTrim$(strNewLine)
        If IsBlockEnd(strNewLine) Then nIndent = nIndent - 1
        If nIndent < 0 Then nIndent = 0
        ' Put back new line.
        vbcomp.CodeModule.ReplaceLine nLine, Space$(nIndent * 4) & strNewLine
        If IsBlockStart(strNewLine) Then nIndent = nIndent + 1
    Next nLine
End Sub
Function ProcedureStartLine(codeMod As CodeModule, ProcedureName As String) As Long
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim StartAt As Long
    Dim EndAt As Long
    Dim CountOf As Long
    StartAt = codeMod.ProcStartLine(ProcedureName, ProcKind)
    EndAt = codeMod.ProcStartLine(ProcedureName, ProcKind) + codeMod.ProcCountLines(ProcedureName, ProcKind) - 1
    CountOf = codeMod.ProcCountLines(ProcedureName, ProcKind)
    ProcedureStartLine = StartAt
End Function

Function ProcedureEndLine(codeMod As CodeModule, procName As String) As Long
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim StartAt As Long
    Dim EndAt As Long
    Dim CountOf As Long
    StartAt = codeMod.ProcStartLine(procName, ProcKind)
    EndAt = codeMod.ProcStartLine(procName, ProcKind) + codeMod.ProcCountLines(procName, ProcKind) - 1
    CountOf = codeMod.ProcCountLines(procName, ProcKind)
    ProcedureEndLine = EndAt
End Function

