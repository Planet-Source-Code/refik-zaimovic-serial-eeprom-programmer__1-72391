VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin i2c_prog.GGCommand cmdCancel 
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   2520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Cancel"
      UseHoverProperties=   0   'False
      FontColorOnHover=   0
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      GradientInverseOnPress=   0   'False
   End
   Begin i2c_prog.GGCommand cmdOk 
      Height          =   375
      Left            =   3660
      TabIndex        =   13
      Top             =   2520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Ok"
      UseHoverProperties=   0   'False
      FontColorOnHover=   0
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      GradientInverseOnPress=   0   'False
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1440
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      BaudRate        =   2400
   End
   Begin VB.Frame frmCom 
      Height          =   1035
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   4875
      Begin VB.CheckBox chkVerDuring 
         Caption         =   "Verify during programming"
         Height          =   255
         Left            =   2100
         TabIndex        =   11
         Top             =   660
         Width           =   2295
      End
      Begin VB.CheckBox chkVerAfter 
         Caption         =   "Verify after programming"
         Height          =   255
         Left            =   2100
         TabIndex        =   10
         Top             =   300
         Width           =   2175
      End
      Begin VB.ComboBox cboComPort 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   1755
      End
      Begin VB.Label lblInfo 
         Caption         =   "Select Com port:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame frmDelay 
      Caption         =   "Delay settings"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4875
      Begin i2c_prog.ctlEBSlider sldI2CDelay 
         Height          =   255
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         SliderWidth     =   150
         BorderStyle     =   2
      End
      Begin i2c_prog.ctlEBSlider sldStartDelay 
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Top             =   780
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         Max             =   500
         SliderWidth     =   150
         BorderStyle     =   2
      End
      Begin VB.Label lblInfo 
         Caption         =   "Start Delay"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   9
         Top             =   780
         Width           =   855
      End
      Begin VB.Label lblInfo 
         Caption         =   "I2C Delay"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblStartDelay 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3960
         TabIndex        =   4
         Top             =   780
         Width           =   795
      End
      Begin VB.Label lblI2CDelay 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3960
         TabIndex        =   3
         Top             =   360
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

  Unload Me
  Set frmSettings = Nothing
  
End Sub

Private Sub cmdOk_Click()

  '\\ set globals variables
  i2cDelay = sldI2CDelay.Value
  iStartDelay = sldStartDelay.Value
  iComPort = CInt(Right$(cboComPort.Text, 1))
  iVerAfter = chkVerAfter.Value
  iVerDuring = chkVerDuring.Value

  '\\ save settings to registry
  SaveSetting App.Title, "Settings", "i2cDelay", i2cDelay       '\\
  SaveSetting App.Title, "Settings", "StartDelay", iStartDelay  '\\
  SaveSetting App.Title, "Settings", "VerAfter", iVerAfter       '\\
  SaveSetting App.Title, "Settings", "VerDuring", iVerDuring     '\\
  SaveSetting App.Title, "Settings", "ComPort", iComPort       '\\

  Unload Me
  Set frmSettings = Nothing
  
End Sub

Private Sub Form_Load()

  CheckAvailableComPort

  sldI2CDelay.Value = i2cDelay
  sldStartDelay.Value = iStartDelay
  
  lblI2CDelay.Caption = i2cDelay & " ms"
  lblStartDelay.Caption = iStartDelay & " ms"
  
  cboComPort.Text = "Com" & iComPort
  chkVerAfter.Value = iVerAfter
  chkVerDuring.Value = iVerDuring
  
End Sub

Private Sub sldI2CDelay_Changed()

  lblI2CDelay.Caption = sldI2CDelay.Value & " ms"
  
End Sub

Private Sub sldStartDelay_Changed()

  lblStartDelay.Caption = sldStartDelay.Value & " ms"

End Sub

Private Sub CheckAvailableComPort()

  Dim i As Integer
  Dim bError As Boolean
  
  For i = 0 To 15
    bError = False
    MSComm1.CommPort = i + 1
    On Error GoTo ErrPort
    MSComm1.PortOpen = True
    If bError = False Then
      MSComm1.PortOpen = False
      cboComPort.AddItem "Com" & i + 1, (cboComPort.ListCount)
    End If
    
ErrPort:
    bError = True
    Resume Next
  Next i
  
End Sub
