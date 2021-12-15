Function ExportProject(wb As Workbook, Optional ExportSheets As Boolean, Optional ExportForms As Boolean, Optional PrintCode As Boolean)
    Dim workbookCleanName   As String: workbookCleanName = Left(wb.name, InStrRev(wb.name, ".") - 1)
    Dim workbookExtension   As String: workbookExtension = Right(wb.name, Len(wb.name) - InStr(1, wb.name, "."))
    Dim mainPath            As String: mainPath = Environ("USERprofile") & "\Documents\" & "vbaCodeArchive\Code Library\"
    Dim exportPath          As String: exportPath = mainPath & workbookCleanName & "\"
    exportPath = exportPath & Format(Now, "YYMMDD HHNNSS") & "\"
    FoldersCreate mainPath: FoldersCreate exportPath 'create folders
    
    If PrintCode = True Then
        printFileName = wb.name
        IndentWorkbook wb
        PrintProject wb
    End If
    
    'On Error Resume Next: Kill exportPath & "*.*": On Error GoTo 0 'empty folder if previous export exists
    'wb.SaveCopyAs exportPath & wb.name           'export workbook backup
    'If wb.name <> ThisWorkbook.name And (ExportSheets = True Or ExportForms = True) Then
        Dim EXT As String: EXT = Right(wb.name, Len(wb.name) - InStr(1, wb.name, "."))
        If EXT = "xlam" Or EXT = "xla" Then wb.IsAddin = False
        If ExportSheets = True Then ExportWorksheetsToPDF wb, exportPath 'Export Worksheets To Image
        If wb.name <> ThisWorkbook.name Then If ExportForms = True Then UserformToPDF wb, exportPath 'export Userform To PDF
        If EXT = "xlam" Or EXT = "xla" Then wb.IsAddin = True
    'End If
    
    Dim procColl As Collection, procedure As Variant, vbcomp As VBComponent, Extension As String
    For Each vbcomp In wb.VBProject.VBComponents
        Select Case vbcomp.Type
        Case vbext_ct_ClassModule, vbext_ct_Document:   Extension = ".cls"
        Case vbext_ct_MSForm:                           Extension = ".frm"
        Case vbext_ct_StdModule:                        Extension = ".bas"
        Case Else:                                      Extension = ".txt"
        End Select
        TxtAppend exportPath & "#UnifiedProject.txt", GetCompText(vbcomp.CodeModule) 'add comp's text to unified project's txt
        If vbcomp.Type = vbext_ct_Document Then
            If vbcomp.name = "ThisWorkbook" Then
                vbcomp.Export exportPath & "DocClass " & vbcomp.name & Extension
            Else
                vbcomp.Export exportPath & "DocClass " & GetSheetByCodeName(workbookOfVbcomponent(vbcomp), vbcomp.name).name & Extension
            End If
        Else
            vbcomp.Export exportPath & vbcomp.name & Extension 'export component
        End If
        Set procColl = ProcList(vbcomp.CodeModule)
        For Each procedure In procColl           'export component's procedures as txt
            TxtAppend exportPath & procedure & ".txt", GetProcText(vbcomp.CodeModule, CStr(procedure))
        Next procedure
    Next
    MsgBox "Export complete"
    'FollowLink exportPath                        'open export folder
End Function

