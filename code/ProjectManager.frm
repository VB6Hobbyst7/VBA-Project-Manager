VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProjectManager 
   Caption         =   "github.com/AlexOfRhodes"
   ClientHeight    =   3372
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4656
   OleObjectBlob   =   "ProjectManager.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProjectManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chPrintCode_Click()
    ThisWorkbook.Sheets("SETTINGS").Range("PrintCode").Value = chPrintCode.Value
End Sub





Private Sub goToFolder_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
FollowLink Environ("USERprofile") & "\Documents\" & "vbaCodeArchive\"
End Sub


Private Sub iExport_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
oExport.Value = True
optionsBlank
iExport.BorderStyle = fmBorderStyleSingle
ExportOptionsShow
End Sub
Private Sub iImport_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
oImport.Value = True
optionsBlank
iImport.BorderStyle = fmBorderStyleSingle
ExportOptionsHide
End Sub
Private Sub iRename_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
oRename.Value = True
optionsBlank
iRename.BorderStyle = fmBorderStyleSingle
ExportOptionsHide
End Sub
Private Sub iRemove_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
oDelete.Value = True
optionsBlank
iRemove.BorderStyle = fmBorderStyleSingle
End Sub
Private Sub iRefresh_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
oRefresh.Value = True
optionsBlank
iRefresh.BorderStyle = fmBorderStyleSingle
ExportOptionsHide
End Sub
Sub optionsBlank()
iExport.BorderStyle = fmBorderStyleNone
iImport.BorderStyle = fmBorderStyleNone
iRename.BorderStyle = fmBorderStyleNone
iRefresh.BorderStyle = fmBorderStyleNone
iRemove.BorderStyle = fmBorderStyleNone
End Sub
Sub ExportOptionsHide()
chExportSheets.Visible = False
chExportForms.Visible = False
chPrintCode.Visible = False
chExportCode.Visible = False
iSettings.Visible = False
End Sub
Sub ExportOptionsShow()
chExportSheets.Visible = True
chExportForms.Visible = True
chPrintCode.Visible = True
chExportCode.Visible = True
iSettings.Visible = True
End Sub



Private Sub UserForm_Initialize()
    chExportSheets.Value = ThisWorkbook.Sheets("SETTINGS").Range("ExportSheets").Value
    chExportForms.Value = ThisWorkbook.Sheets("SETTINGS").Range("ExportForms").Value
    chPrintCode.Value = ThisWorkbook.Sheets("SETTINGS").Range("PrintCode").Value
    
    FormatColourFormatters
End Sub

Private Sub chExportSheets_Click()
    ThisWorkbook.Sheets("SETTINGS").Range("ExportSheets").Value = chExportSheets.Value
End Sub

Private Sub chExportForms_Click()
    ThisWorkbook.Sheets("SETTINGS").Range("ExportForms").Value = chExportForms.Value
End Sub

Private Sub ActiveFile_Click()
    Set wb = ActiveWorkbook
    SelectAction
End Sub

Private Sub iSettings_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

If LBLcolourCode.Visible = True Then
LBLcolourCode.Visible = False
LBLcolourKey.Visible = False
LBLcolourComment.Visible = False
LBLcolourOdd.Visible = False
Else
LBLcolourCode.Visible = True
LBLcolourKey.Visible = True
LBLcolourComment.Visible = True
LBLcolourOdd.Visible = True
End If
End Sub

Sub SelectAction()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("SETTINGS")
    Select Case True
    Case oExport.Value = True
        Me.Hide
        ExportProject wb, ws.Range("ExportSheets"), ws.Range("ExportForms"), ws.Range("PrintCode")
        Me.Show
    Case oImport.Value = True
        ImportComponents wb
    Case oRefresh.Value = True
        RefreshComponents wb
    Case oDelete.Value = True
        RemoveComps.Show
    Case oRename.Value = True
        RenameComps.Show
    Case Else
        '
    End Select
End Sub

Private Sub cInfo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    uDEV.Show
End Sub

Private Sub Image1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ExportSettings.Show
End Sub

Private Sub SelectFile_Click()
    Dim fPath As String
    fPath = PickExcelFile
    If fPath = "" Then Exit Sub
    Set wb = Workbooks.Open(fileName:=fPath, UpdateLinks:=0, ReadOnly:=False)
    SelectAction
    Set wb = Nothing
End Sub


Private Sub LBLcolourCode_Click()
    ColorPaletteDialog ThisWorkbook.Sheets("TXTColour").Range("GeneralFontBackground"), LBLcolourCode
End Sub
Private Sub LBLcolourComment_Click()
    ColorPaletteDialog ThisWorkbook.Sheets("TXTColour").Range("ColourComments"), LBLcolourComment
End Sub
Private Sub LBLcolourKey_Click()
    ColorPaletteDialog ThisWorkbook.Sheets("TXTColour").Range("ColourKeywords"), LBLcolourKey
End Sub
Private Sub LBLcolourOdd_Click()
    ColorPaletteDialog ThisWorkbook.Sheets("TXTColour").Range("OddLine"), LBLcolourOdd
End Sub


