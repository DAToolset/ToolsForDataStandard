VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbout 
   Caption         =   "About & License"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7260
   OleObjectBlob   =   "frmAbout.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub lblBlogLink_Click()
    ThisWorkbook.FollowHyperlink Address:=lblBlogLink.Caption
End Sub

Private Sub lblGithubLink_Click()
    ThisWorkbook.FollowHyperlink Address:=lblGithubLink.Caption
End Sub

Private Sub txtEmailAddress_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisWorkbook.FollowHyperlink Address:="mailto:" + txtEmailAddress.Text
End Sub

Private Sub UserForm_Activate()
    lblVersion.Caption = GetVersionString
End Sub
