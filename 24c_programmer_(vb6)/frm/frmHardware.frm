VERSION 5.00
Begin VB.Form frmHardware 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8190
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHardware.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmHardware.frx":000C
   ScaleHeight     =   4875
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgVer1 
      Height          =   4905
      Left            =   0
      Picture         =   "frmHardware.frx":597E
      Top             =   0
      Width           =   8190
   End
End
Attribute VB_Name = "frmHardware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()

  imgVer1.Visible = Not imgVer1.Visible

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

'  If KeyAscii = 27 Then
    Unload Me
    Set frmHardware = Nothing
'  End If

End Sub

Private Sub imgVer1_Click()

  imgVer1.Visible = Not imgVer1.Visible

End Sub

