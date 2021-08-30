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
        msgPOP ("Outlook not found" & Chr(10) & _
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

Private Sub Label2_Click()
    FollowLink ("https://www.facebook.com/VBA-Code-Archive-110295994460212")
End Sub

Private Sub Image2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    FollowLink ("https://bit.ly/2QT4wFe")
End Sub

Private Sub Label1_Click()
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

Private Sub Label5_Click()
    FollowLink ("https://github.com/alexofrhodes")
End Sub


