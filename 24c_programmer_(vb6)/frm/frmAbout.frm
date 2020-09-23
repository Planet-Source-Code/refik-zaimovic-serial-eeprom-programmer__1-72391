VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4800
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
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin i2c_prog.GGCommand cmdExit 
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   3060
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Exit"
      UseHoverProperties=   0   'False
      FontColorOnHover=   0
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AlwaysShowBorder=   -1  'True
      UseGradientFill =   0   'False
      GradientFromColor=   -2147483633
      GradientToColor =   -2147483633
   End
   Begin VB.Label Label2 
      Enabled         =   0   'False
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   3120
      Width           =   1395
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   120
      X2              =   4680
      Y1              =   2840
      Y2              =   2840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   2
      X1              =   120
      X2              =   4680
      Y1              =   2820
      Y2              =   2820
   End
   Begin VB.Label Label1 
      Caption         =   "e-mail: rz@edaboard.com                                   rele@cg.yu"
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   2280
      Width           =   2835
   End
   Begin VB.Label Label1 
      Caption         =   "Author: Refik Zaimovic"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   1980
      Width           =   4155
   End
   Begin VB.Label Label1 
      Caption         =   "This program allow reading and writing 24cxx serial eeproms. Special interface is required."
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   4155
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   120
      X2              =   4680
      Y1              =   980
      Y2              =   980
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   120
      X2              =   4680
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "i2c eeprom programmer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   3060
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmAbout.frx":000C
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()

  Unload Me
  Set frmAbout = Nothing

End Sub

Private Sub Form_Click()

  Unload Me
  Set frmAbout = Nothing

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

  If KeyAscii = 27 Then
    Unload Me
    Set frmAbout = Nothing
  End If
  
End Sub

Private Sub Form_Load()

  Label2.Caption = "Ver. " & App.Major & "." & App.Minor & "." & App.Revision

End Sub

