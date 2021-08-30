VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Comps 
   Caption         =   "UserForm1"
   ClientHeight    =   6072
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3588
   OleObjectBlob   =   "Comps.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Comps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Remover_Click()
    If LComponents.ListCount = 0 Then Exit Sub
    Dim i As Long
    For i = 0 To LComponents.ListCount - 1
        If LComponents.Selected(i) Then
            If oCode.Value = True Then
                ClearComponent wb.VBProject.VBComponents(LComponents.List(i, 1))
            ElseIf oComps.Value = True Then
                DeleteComponent wb.VBProject.VBComponents(LComponents.List(i, 1))
            End If
        End If
    Next i
    addCompsList
End Sub

Private Sub UserForm_Initialize()
    addCompsList
    Me.Caption = "Comps of " & wb.name
End Sub

