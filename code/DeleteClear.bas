Attribute VB_Name = "DeleteClear"
Sub DeleteComponent(vbcomp As VBComponent)
Application.DisplayAlerts = False
    If vbcomp.Type = vbext_ct_Document Then
        If vbcomp.name = "ThisWorkbook" Then
            vbcomp.CodeModule.DeleteLines 1, vbcomp.CodeModule.CountOfLines
        Else
            If workbookOfVbcomponent(vbcomp).Sheets.Count > 1 Then
                GetSheetByCodeName(workbookOfVbcomponent(vbcomp), vbcomp.name).Delete
            Else
                If RemoveComps.oDeleteSheets.Value = True Then
                    Dim ws As Worksheet
                    Set ws = workbookOfVbcomponent(vbcomp).Sheets.Add
                    ws.name = "All other sheets were deleted"
                    GetSheetByCodeName(workbookOfVbcomponent(vbcomp), vbcomp.name).Delete
                End If
            End If
            
        End If
    Else
        workbookOfVbcomponent(vbcomp).VBProject.VBComponents.Remove vbcomp
    End If
Application.DisplayAlerts = True
End Sub

Sub ClearComponent(vbcomp As VBComponent)
    vbcomp.CodeModule.DeleteLines 1, vbcomp.CodeModule.CountOfLines
End Sub

