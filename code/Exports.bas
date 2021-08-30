Attribute VB_Name = "Exports"
Function ExportProject(wb As Workbook, Optional ExportSheets As Boolean, Optional ExportForms As Boolean)
    Dim workbookCleanName   As String: workbookCleanName = Left(wb.name, InStrRev(wb.name, ".") - 1)
    Dim workbookExtension   As String: workbookExtension = Right(wb.name, Len(wb.name) - InStr(1, wb.name, "."))
    Dim mainPath            As String: mainPath = Environ("USERprofile") & "\Documents\" & "vbaCodeArchive\Code Library\"
    Dim exportPath          As String: exportPath = mainPath & workbookCleanName & "\"
    FoldersCreate mainPath: FoldersCreate exportPath 'create folders
    On Error Resume Next: Kill exportPath & "*.*": On Error GoTo 0 'empty folder if previous export exists
    wb.SaveCopyAs exportPath & wb.name           'export workbook backup
    'If wb.name <> ThisWorkbook.name And (ExportSheets = True Or ExportForms = True) Then
        Dim EXT As String: EXT = Right(wb.name, Len(wb.name) - InStr(1, wb.name, "."))
        If EXT = "xlam" Or EXT = "xla" Then wb.IsAddin = False
        If ExportSheets = True Then ExportWorksheetsToPDF wb, exportPath 'Export Worksheets To Image
        If wb.name <> ThisWorkbook.name Then If ExportForms = True Then UserformToPDF wb, exportPath 'export Userform To PDF
        If EXT = "xlam" Or EXT = "xla" Then wb.IsAddin = True
    'End If
    Dim procColl As Collection, procedure As Variant, vbComp As VBComponent, Extension As String
    For Each vbComp In wb.VBProject.VBComponents
        Select Case vbComp.Type
        Case vbext_ct_ClassModule, vbext_ct_Document:   Extension = ".cls"
        Case vbext_ct_MSForm:                           Extension = ".frm"
        Case vbext_ct_StdModule:                        Extension = ".bas"
        Case Else:                                      Extension = ".txt"
        End Select
        TxtAppend exportPath & "#UnifiedProject.txt", GetCompText(vbComp.CodeModule) 'add comp's text to unified project's txt
        If vbComp.Type = vbext_ct_Document Then
            If vbComp.name = "ThisWorkbook" Then
                vbComp.Export exportPath & "DocClass " & vbComp.name & Extension
            Else
                vbComp.Export exportPath & "DocClass " & getSheetByCodeName(workbookOfVbcomponent(vbComp), vbComp.name).name & Extension
            End If
        Else
            vbComp.Export exportPath & vbComp.name & Extension 'export component
        End If
        Set procColl = ProcList(vbComp.CodeModule)
        For Each procedure In procColl           'export component's procedures as txt
            TxtAppend exportPath & procedure & ".txt", GetProcText(vbComp.CodeModule, CStr(procedure))
        Next procedure
    Next
    FollowLink exportPath                        'open export folder
End Function

