VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uDEV 
   Caption         =   "vbaCodeArchive ~ Anastasiou Alex"
   ClientHeight    =   3084
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4104
   OleObjectBlob   =   "uDEV.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uDEV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Feedback_Click()
    If OutlookCheck = True Then
        MailDev
    Else
        Clipboard ("anastasioualex@gmail.com")
        MsgBox ("Outlook not found" & Chr(10) & _
                "DEV email address copied to clipboard")
    End If
End Sub

Private Sub Image4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If OutlookCheck = True Then
        MailDev
    Else
        Clipboard ("anastasioualex@gmail.com")
        msgPOP ("Outlook not found" & Chr(10) & _
                "DEV email address copied to clipboard")
    End If
End Sub

Private Sub Image1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    FollowLink ("https://www.facebook.com/VBA-Code-Archive-110295994460212")
End Sub

Private Sub Image2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    FollowLink ("https://bit.ly/2QT4wFe")
End Sub

Private Sub Image3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    FollowLink ("http://paypal.me/alexofrhodes")
End Sub

Private Sub Label3_Click()
    FollowLink ("http://paypal.me/alexofrhodes")
End Sub

Private Sub Image5_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    FollowLink ("https://github.com/alexofrhodes")
End Sub


Sub FollowLink(folderPath As String)
    Dim oShell As Object
    Dim Wnd As Object
    Set oShell = CreateObject("Shell.Application")
    For Each Wnd In oShell.Windows
        If Wnd.name = "File Explorer" Then
            If Wnd.Document.Folder.Self.path & "\" = folderPath Then Exit Sub
        End If
    Next Wnd
    Application.ThisWorkbook.FollowHyperlink Address:=folderPath, NewWindow:=True
End Sub


Private Function Clipboard(Optional StoreText As String) As String
    'PURPOSE: Read/Write to Clipboard
    'Source: ExcelHero.com (Daniel Ferry)

    Dim X As Variant

    'Store as variant for 64-bit VBA support
    X = StoreText

    'Create HTMLFile Object
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            Select Case True
            Case Len(StoreText)
                'Write to the clipboard
                .SetData "text", X
            Case Else
                'Read from the clipboard (no variable passed through)
                Clipboard = .GetData("text")
            End Select
        End With
    End With

End Function
Sub MailDev()
    'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
    'Working in Office 2000-2016
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    '    strbody = "Hi there" & vbNewLine & vbNewLine & _
    "This is line 1" & vbNewLine & _
    "This is line 2" & vbNewLine & _
    "This is line 3" & vbNewLine & _
    "This is line 4"
    On Error Resume Next
    With OutMail
        .To = "anastasioualex@gmail.com"
        .CC = vbNullString
        .BCC = vbNullString
        .Subject = "DEV REQUEST OR FEEDBACK FOR -CODE ARCHIVE-"
        .body = strbody
        'You can add a file like this
        '.Attachments.Add ("C:\test.txt")
        '.Send
        .Display
    End With
    On Error GoTo 0
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

Function OutlookCheck() As Boolean
    'is outlook installed?
    Dim xOLApp As Object
    '    On Error GoTo L1
    Set xOLApp = CreateObject("Outlook.Application")
    If Not xOLApp Is Nothing Then
        OutlookCheck = True
        'MsgBox "Outlook " & xOLApp.Version & " installed", vbExclamation
        Set xOLApp = Nothing
        Exit Function
    End If
    OutlookCheck = False
    'L1: MsgBox "Outlook not installed", vbExclamation, "Kutools for Outlook"
End Function

