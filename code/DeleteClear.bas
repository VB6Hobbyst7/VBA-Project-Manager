Attribute VB_Name = "DeleteClear"
Sub DeleteComponent(vbComp As VBComponent)
    If vbComp.Type = vbext_ct_Document Then
        If vbComp.name = "ThisWorkbook" Then
            vbComp.CodeModule.DeleteLines 1, vbComp.CodeModule.CountOfLines
        Else
            If workbookOfVbcomponent(vbComp).Sheets.Count > 1 Then
                getSheetByCodeName(workbookOfVbcomponent(vbComp), vbComp.name).Delete
            End If
        End If
    Else
        workbookOfVbcomponent(vbComp).VBProject.VBComponents.Remove vbComp
    End If
End Sub

Sub ClearComponent(vbComp As VBComponent)
    vbComp.CodeModule.DeleteLines 1, vbComp.CodeModule.CountOfLines
End Sub
