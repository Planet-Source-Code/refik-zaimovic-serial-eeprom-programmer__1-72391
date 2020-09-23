VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "I2C Programmer"
   ClientHeight    =   6390
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   10695
   FillColor       =   &H8000000D&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin i2c_prog.GGCommand cmdToolBar 
      Height          =   375
      Index           =   6
      Left            =   3660
      TabIndex        =   8
      TabStop         =   0   'False
      Tag             =   "Blank check"
      Top             =   60
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      Caption         =   ""
      UseHoverProperties=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMain.frx":08CA
      PictureAlignment=   4
      UseGradientFill =   0   'False
      GradientFromColor=   -2147483633
      GradientToColor =   -2147483633
      GradientInverseOnPress=   0   'False
   End
   Begin i2c_prog.GGCommand cmdToolBar 
      Height          =   375
      Index           =   5
      Left            =   3120
      TabIndex        =   7
      TabStop         =   0   'False
      Tag             =   "Write eeprom"
      Top             =   60
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      Caption         =   ""
      UseHoverProperties=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMain.frx":0E64
      PictureAlignment=   4
      UseGradientFill =   0   'False
      GradientFromColor=   -2147483633
      GradientToColor =   -2147483633
      GradientInverseOnPress=   0   'False
   End
   Begin i2c_prog.GGCommand cmdToolBar 
      Height          =   375
      Index           =   4
      Left            =   2580
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "Read eeprom"
      Top             =   60
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      Caption         =   ""
      UseHoverProperties=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMain.frx":13FE
      PictureAlignment=   4
      UseGradientFill =   0   'False
      GradientFromColor=   -2147483633
      GradientToColor =   -2147483633
      GradientInverseOnPress=   0   'False
   End
   Begin i2c_prog.GGCommand cmdToolBar 
      Height          =   375
      Index           =   3
      Left            =   1860
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "Settings"
      Top             =   60
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      Caption         =   ""
      UseHoverProperties=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMain.frx":1998
      PictureAlignment=   4
      UseGradientFill =   0   'False
      GradientFromColor=   -2147483633
      GradientToColor =   -2147483633
      GradientInverseOnPress=   0   'False
   End
   Begin i2c_prog.GGCommand cmdToolBar 
      Height          =   375
      Index           =   2
      Left            =   1140
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "Save File"
      Top             =   60
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      Caption         =   ""
      UseHoverProperties=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMain.frx":1AF2
      PictureAlignment=   4
      UseGradientFill =   0   'False
      GradientFromColor=   -2147483633
      GradientToColor =   -2147483633
      GradientInverseOnPress=   0   'False
   End
   Begin i2c_prog.GGCommand cmdToolBar 
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   "Reload File"
      Top             =   60
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      Caption         =   ""
      UseHoverProperties=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMain.frx":208C
      PictureAlignment=   4
      UseGradientFill =   0   'False
      GradientFromColor=   -2147483633
      GradientToColor =   -2147483633
      GradientInverseOnPress=   0   'False
   End
   Begin i2c_prog.GGCommand cmdToolBar 
      Height          =   375
      Index           =   0
      Left            =   60
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "Open File"
      Top             =   60
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      Caption         =   ""
      UseHoverProperties=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMain.frx":21E6
      PictureAlignment=   4
      UseGradientFill =   0   'False
      GradientFromColor=   -2147483633
      GradientToColor =   -2147483633
      GradientInverseOnPress=   0   'False
   End
   Begin i2c_prog.GGCommand cmdToolBar 
      Height          =   375
      Index           =   7
      Left            =   4200
      TabIndex        =   9
      TabStop         =   0   'False
      Tag             =   "Verify data"
      Top             =   60
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      Caption         =   ""
      UseHoverProperties=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMain.frx":2340
      PictureAlignment=   4
      UseGradientFill =   0   'False
      GradientFromColor=   -2147483633
      GradientToColor =   -2147483633
      GradientInverseOnPress=   0   'False
   End
   Begin i2c_prog.GGCommand cmdToolBar 
      Height          =   375
      Index           =   9
      Left            =   10080
      TabIndex        =   11
      TabStop         =   0   'False
      Tag             =   "Exit program"
      Top             =   60
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      Caption         =   ""
      UseHoverProperties=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMain.frx":28DA
      PictureAlignment=   4
      UseGradientFill =   0   'False
      GradientFromColor=   -2147483633
      GradientToColor =   -2147483633
      GradientInverseOnPress=   0   'False
   End
   Begin i2c_prog.GGCommand cmdToolBar 
      Height          =   375
      Index           =   8
      Left            =   8040
      TabIndex        =   10
      TabStop         =   0   'False
      Tag             =   "Device info"
      Top             =   60
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      Caption         =   ""
      UseHoverProperties=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMain.frx":2E74
      PictureAlignment=   4
      UseGradientFill =   0   'False
      GradientFromColor=   -2147483633
      GradientToColor =   -2147483633
      GradientInverseOnPress=   0   'False
   End
   Begin VB.ComboBox cboMemory 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmMain.frx":340E
      Left            =   6360
      List            =   "frmMain.frx":3430
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select memory"
      Top             =   120
      Width           =   1635
   End
   Begin i2c_prog.ucStatusbar Sbar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      Top             =   6135
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   450
   End
   Begin VB.Frame fMain 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   60
      TabIndex        =   12
      Top             =   480
      Width           =   10575
      Begin VB.ListBox lstBuffer 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         ItemData        =   "frmMain.frx":347D
         Left            =   180
         List            =   "frmMain.frx":347F
         TabIndex        =   15
         Tag             =   "Double click to edit data"
         Top             =   480
         Width           =   10215
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5235
         Index           =   0
         Left            =   1020
         ScaleHeight     =   5235
         ScaleWidth      =   75
         TabIndex        =   14
         Top             =   180
         Width           =   75
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5235
         Index           =   1
         Left            =   7500
         ScaleHeight     =   5235
         ScaleWidth      =   75
         TabIndex        =   13
         Top             =   180
         Width           =   75
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00808080&
         Caption         =   "  Adress:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00808080&
         Caption         =   "   00    01    02    03    04    05    06    07   08    09    0A    0B    0C    0D    0E    0F"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   17
         Top             =   240
         Width           =   6435
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "ASCII"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   7560
         TabIndex        =   16
         Top             =   240
         Width           =   2835
      End
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   60
      X2              =   10620
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   60
      X2              =   10620
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   5
      X1              =   4820
      X2              =   4820
      Y1              =   60
      Y2              =   420
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   4
      X1              =   4800
      X2              =   4800
      Y1              =   60
      Y2              =   420
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   2475
      X2              =   2475
      Y1              =   60
      Y2              =   420
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   2
      X1              =   2460
      X2              =   2460
      Y1              =   60
      Y2              =   420
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   1740
      X2              =   1740
      Y1              =   60
      Y2              =   420
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   1755
      X2              =   1755
      Y1              =   60
      Y2              =   420
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Select memory:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   5220
      TabIndex        =   0
      Top             =   180
      Width           =   1125
   End
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu mOpen 
         Caption         =   "Open File ..."
      End
      Begin VB.Menu mSave 
         Caption         =   "Save File As ..."
      End
      Begin VB.Menu crtica0 
         Caption         =   "-"
      End
      Begin VB.Menu mReload 
         Caption         =   "Reload File"
      End
      Begin VB.Menu crtica3 
         Caption         =   "-"
      End
      Begin VB.Menu mPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu crtica1 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mFillBuffer 
         Caption         =   "Fill buffer with random data"
      End
      Begin VB.Menu mClearBuffer 
         Caption         =   "Clear Buffer"
      End
   End
   Begin VB.Menu mSettings 
      Caption         =   "&Settings"
      Begin VB.Menu mDevice 
         Caption         =   "Device"
         Begin VB.Menu mI2C 
            Caption         =   "I2C eeprom"
            Begin VB.Menu m24cxx 
               Caption         =   "24C01"
               Index           =   0
            End
            Begin VB.Menu m24cxx 
               Caption         =   "24c02"
               Index           =   1
            End
            Begin VB.Menu m24cxx 
               Caption         =   "24c04"
               Index           =   2
            End
            Begin VB.Menu m24cxx 
               Caption         =   "24c08"
               Index           =   3
            End
            Begin VB.Menu m24cxx 
               Caption         =   "24c16"
               Index           =   4
            End
            Begin VB.Menu m24cxx 
               Caption         =   "24c32"
               Index           =   5
            End
            Begin VB.Menu m24cxx 
               Caption         =   "24c64"
               Index           =   6
            End
            Begin VB.Menu m24cxx 
               Caption         =   "24c128"
               Index           =   7
            End
            Begin VB.Menu m24cxx 
               Caption         =   "24c256"
               Index           =   8
            End
            Begin VB.Menu m24cxx 
               Caption         =   "24c512"
               Index           =   9
            End
         End
      End
      Begin VB.Menu mOptions 
         Caption         =   "Options"
      End
   End
   Begin VB.Menu mCommand 
      Caption         =   "&Action"
      Begin VB.Menu mOperation 
         Caption         =   "Read Eeprom"
         Index           =   0
      End
      Begin VB.Menu mOperation 
         Caption         =   "Program eeprom"
         Index           =   1
      End
      Begin VB.Menu mOperation 
         Caption         =   "Blank Check"
         Index           =   2
      End
      Begin VB.Menu mOperation 
         Caption         =   "Verify Data"
         Index           =   3
      End
   End
   Begin VB.Menu mHlp 
      Caption         =   "&Help"
      Begin VB.Menu mHelp 
         Caption         =   "Hardware"
      End
      Begin VB.Menu mAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mouse As Single     '\\ tracking mouse position

Private Sub Fill_FF(Index As Long)

  Dim i As Long
  Dim adr As String
  Dim cont As String
  Dim m As Long
  
  lstBuffer.Clear
  cont = vbTab & "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF" & vbTab & vbTab & "ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ"
  m = 0
  
  For i = 0 To (2 ^ (Index + 3)) - 1
    adr = Right$("000" & Hex(i * 16), 4) & ":"
    lstBuffer.AddItem adr & cont, m
    m = m + 1
  Next i

End Sub

Private Sub cboMemory_Click()

  DeviceSelection (cboMemory.ListIndex)
  
End Sub

Private Sub cmdToolBar_Click(Index As Integer)

  Select Case Index
  
    Case 0: mOpen_Click                       '\\ open
    
    Case 1: OpenFile (Sbar.PanelText(2))      '\\ reload
    
    Case 2: SaveFile                          '\\ save
    
    Case 3: frmSettings.Show vbModal          '\\ settings
    
    Case 4, 5, 6, 7: EepromAction = Index - 4 '\\ read, program, blankcheck
                     frmProgram.Show vbModal  '\\ and verify data
     
    Case 8: frmDeviceInfo.Show vbModal        '\\ device info
    
    Case 9: Unload Me: Set frmMain = Nothing  '\\ exit
      
  End Select
  
End Sub

Private Sub cmdToolBar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  Sbar.PanelText(1) = cmdToolBar(Index).Tag
  
End Sub

Private Sub Form_Load()

  Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
  Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
  Me.Width = 10785
  Me.Height = 7050
  cboMemory.ListIndex = GetSetting(App.Title, "Settings", "Memory", 0)
    
  i2cDelay = GetSetting(App.Title, "Settings", "i2cDelay", 20)        '\\
  iStartDelay = GetSetting(App.Title, "Settings", "StartDelay", 200)  '\\
  iVerAfter = GetSetting(App.Title, "Settings", "VerAfter", 0)        '\\
  iVerDuring = GetSetting(App.Title, "Settings", "VerDuring", 0)      '\\
  iComPort = GetSetting(App.Title, "Settings", "ComPort", 1)          '\\
  
  With Sbar
    .Initialize
    .Font.Bold = False
    .Font.Name = "Tahoma"
    .Font.Size = "8"
    .AddPanel sbNormal, , 200, sbNoAutosize
    .AddPanel sbNormal, , 200, sbSpring
    .AddPanel sbNormal, 10, 50, sbNoAutosize
    .PanelText(3) = "Com " & GetSetting(App.Title, "Settings", "ComPort", 1)
  End With
  
  bDirty = False
  Fill_FF (cboMemory.ListIndex)
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Sbar.PanelText(1) = vbNullString
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

  If Me.WindowState <> vbMinimized Then
    SaveSetting App.Title, "Settings", "MainLeft", Me.Left
    SaveSetting App.Title, "Settings", "MainTop", Me.Top
    SaveSetting App.Title, "Settings", "Memory", cboMemory.ListIndex
  End If

End Sub

Private Sub lstBuffer_DblClick()

  Dim Index As Integer
  Dim iPos As Integer
  Dim iAsciiPos As Integer
  
  If mouse > 885 And mouse < 6645 Then    '\\ here are the data
  
    sLine = lstBuffer.Text
    Index = lstBuffer.ListIndex     '\\ index
    
    iPos = 0
    iAsciiPos = 0
    
    If mouse > 885 And mouse < 1245 Then iPos = 7:   iAsciiPos = 56
    If mouse > 1245 And mouse < 1605 Then iPos = 10: iAsciiPos = 57
    If mouse > 1605 And mouse < 1965 Then iPos = 13: iAsciiPos = 58
    If mouse > 1965 And mouse < 2325 Then iPos = 16: iAsciiPos = 59
    If mouse > 2325 And mouse < 2685 Then iPos = 19: iAsciiPos = 60
    If mouse > 2685 And mouse < 3045 Then iPos = 22: iAsciiPos = 61
    If mouse > 3045 And mouse < 3405 Then iPos = 25: iAsciiPos = 62
    If mouse > 3405 And mouse < 3765 Then iPos = 28: iAsciiPos = 63
    If mouse > 3765 And mouse < 4125 Then iPos = 31: iAsciiPos = 64
    If mouse > 4125 And mouse < 4485 Then iPos = 34: iAsciiPos = 65
    If mouse > 4485 And mouse < 4845 Then iPos = 37: iAsciiPos = 66
    If mouse > 4845 And mouse < 5205 Then iPos = 40: iAsciiPos = 67
    If mouse > 5205 And mouse < 5565 Then iPos = 43: iAsciiPos = 68
    If mouse > 5565 And mouse < 5925 Then iPos = 46: iAsciiPos = 69
    If mouse > 5925 And mouse < 6285 Then iPos = 49: iAsciiPos = 70
    If mouse > 6285 And mouse < 6645 Then iPos = 52: iAsciiPos = 71
    
    If iPos <> 0 Then
    
      bByte = CByte("&H" & Mid$(lstBuffer.Text, iPos, 2))
      frmEdit.Show vbModal
      
      If bIsChanged Then
      
        '\\ hex replacment \\
        '\\ first part of string is up to the position we clicked. \\
        '\\ then we extract selected character and replace it with now one. \\
        '\\ Replace function cut everything before changed part \\
        '\\ and return string from changed part up to the end, including changes. \\
        '\\ therefore we need Left$(sLine, iPos - 1) on the begining. \\
        sLine = Left$(sLine, iPos - 1) & Replace(sLine, Mid$(sLine, iPos + 1, 1), sHex, iPos + 1, 1, vbTextCompare)
        
        '\\ h00 and hA0 produce 'empty' spaces in ascii and may couse errors when \\
        '\\ we try to edit data. also h09 and h20 can couse formating errors. \\
        '\\ to prevent this we put h2E or ascii(.) instead \\
        If sHex = "00" Or sHex = "A0" Or sHex = "09" Or sHex = "20" Then sHex = "2E"
        
        '\\ ascii replacment \\
        '\\ same like hex just move forward to the ascii part of the string \\
        '\\ string up to selected ascii character - Left$(sLine, iAsciiPos - 1) \\
        '\\ this is character we must change - Mid$(sLine, iAsciiPos, 1) \\
        sLine = Left$(sLine, iAsciiPos - 1) & Replace(sLine, Mid$(sLine, iAsciiPos, 1), Chr(Fix("&H" & sHex)), iAsciiPos, 1, vbTextCompare)
        
        lstBuffer.RemoveItem Index
        lstBuffer.AddItem sLine, Index
          
      End If
    End If
  End If
    
End Sub

Private Sub lstBuffer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
  mouse = X
  Sbar.PanelText(1) = IIf((mouse > 885 And mouse < 6645), lstBuffer.Tag, vbNullString)
  
End Sub

Private Sub m24cxx_Click(Index As Integer)

  cboMemory.ListIndex = Index
  DeviceSelection (Index)
  
End Sub

Private Sub DeviceSelection(Index As Integer)

  Fill_FF (Index)
  Sbar.PanelText(2) = vbNullString
  
  '\\ set menu bar
  Dim i As Integer

  For i = 0 To 9
    m24cxx(i).Checked = False
  Next i
  
  m24cxx(Index).Checked = True
  
  iMemory = Index

End Sub

Private Sub mAbout_Click()

  frmAbout.Show vbModal
  
End Sub

Private Sub mClearBuffer_Click()

  Fill_FF (cboMemory.ListIndex)
  Sbar.PanelText(2) = vbNullString
  
End Sub

Private Sub mExit_Click()
  
  Unload Me
  Set frmMain = Nothing

End Sub

Private Sub mFillBuffer_Click()

  Dim adr       As String    '\\ address string
  Dim r         As Long      '\\ row
  Dim c         As Integer   '\\ column
  Dim Index     As Long      '\\ list index
  Dim a         As Byte      '\\ one byte of data
  Dim sAscTemp  As String    '\\ temporary for data (ascii)
  Dim sHexTemp  As String    '\\ temporary for data (hex)
  
  lstBuffer.Clear
  lstBuffer.Visible = False
  
  For r = 0 To (2 ^ (cboMemory.ListIndex + 3)) - 1
  
    adr = Right$("000" & Hex(r * 16), 4) & ":"
    
    For c = 0 To 15
      a = Int((255 * Rnd) + 1) '\\ Generate random value between 1 and 255.
      
      '\\ generate random number in given range
      'Int((upperbound - lowerbound + 1) * Rnd + lowerbound)

      '\\ prepare hex data \\
      sHexTemp = sHexTemp & Right$("0" & Hex(a), 2) & " "
      
      '\\ convert h00, hA0, h09 and h20 to ascii(.) - see lstBuffer_DblClick routine \\
      If a = 160 Or a = 0 Or a = 9 Or a = 32 Then a = 46
      
      '\\ prepare ascii data \\
      sAscTemp = sAscTemp & Chr$(a)
    Next c
    
    lstBuffer.AddItem adr & vbTab & RTrim(sHexTemp) & vbTab & vbTab & sAscTemp, Index
    
    sAscTemp = vbNullString
    sHexTemp = vbNullString
    
    Index = Index + 1
      
  Next r
  
  lstBuffer.Visible = True
    
End Sub

Private Sub mHelp_Click()

  frmHardware.Show vbModal
  
End Sub

Private Sub mOpen_Click()
  Dim sFileName As String
  
  If bDirty Then  '\\ file is changed
    If MsgBox("File is changed. Do you want to save file first?", vbYesNo Or vbInformation Or vbDefaultButton1, "Confirm") = vbNo Then
      VBGetOpenFileName sFileName, , False, False, False, False, ".bin (*.bin)|*.bin", , App.Path, "Open file", "bin"
      OpenFile (sFileName)
    Else
      SaveFile
    End If
  Else
    VBGetOpenFileName sFileName, , False, False, False, False, ".bin (*.bin)|*.bin", , App.Path, "Open file", "bin"
    OpenFile (sFileName)
  End If
  
End Sub

Private Sub OpenFile(sFileName As String)
  
  Dim ff        As Integer
  Dim adr       As String    '\\ address string
  Dim r         As Long      '\\ row
  Dim c         As Integer   '\\ column
  Dim Index     As Long      '\\ list index
  Dim a         As Byte      '\\ one byte of data
  Dim sAscTemp  As String    '\\ temporary for data (ascii)
  Dim sHexTemp  As String    '\\ temporary for data (hex)
    
  If sFileName <> vbNullString Then
  
    ff = FreeFile
    Index = 0
    lstBuffer.Clear
    lstBuffer.Visible = False

    Open sFileName For Binary As #ff
    
    For r = 0 To (2 ^ (cboMemory.ListIndex + 3)) - 1
    
      adr = Right$("000" & Hex(r * 16), 4) & ":"
      
      For c = 0 To 15
        Get #ff, , a
        
        '\\ prepare hex data \\
        sHexTemp = sHexTemp & Right$("0" & Hex(a), 2) & " "
        
        '\\ convert h00, hA0, h09 and h20 to ascii(.) - see lstBuffer_DblClick routine \\
        If a = 160 Or a = 0 Or a = 9 Or a = 32 Then a = 46
        
        '\\ prepare ascii data \\
        sAscTemp = sAscTemp & Chr$(a)
      Next c
      
      lstBuffer.AddItem adr & vbTab & RTrim(sHexTemp) & vbTab & vbTab & sAscTemp, Index
      
      sAscTemp = vbNullString
      sHexTemp = vbNullString
      
      Index = Index + 1
        
    Next r
    
    Close #ff
    lstBuffer.Visible = True
    Sbar.PanelText(2) = sFileName
    bDirty = False
    
  End If
    
End Sub

Private Sub mOperation_Click(Index As Integer)

  EepromAction = Index
  
  frmProgram.Show vbModal

End Sub

Private Sub mOptions_Click()

  frmSettings.Show vbModal
  
End Sub

Private Sub mPrint_Click()

  PrintFile
  
End Sub

Private Sub mReload_Click()

  OpenFile (Sbar.PanelText(2))
  
End Sub

Private Sub mSave_Click()

  SaveFile
  
End Sub

Private Sub PrintFile()
  Dim i As Long
  
  If VBPrintDlg(5, eprPageNumbers, False, , , , , True, False, , , True) = True Then
  
    With Printer
      .Font = "Lucida Console"
      .FontBold = False
      .FontItalic = False
      .FontName = "Lucida Console"
      .FontSize = 8
      .FontStrikethru = False
      .FontUnderline = False
      .Orientation = 1
      .Copies = 1
    End With
    
    For i = 0 To lstBuffer.ListCount - 1
      Printer.Print lstBuffer.List(i)
    Next i
    
    Printer.EndDoc
  
  End If
  
End Sub

Private Sub SaveFile()

  Dim ff As Integer
  Dim sFileName As String
  Dim i As Integer
  Dim m As Integer
  
  VBGetSaveFileName sFileName, , True, ".bin (*.bin)|*.bin", , App.Path, "Save file", "bin"

  If sFileName <> vbNullString Then

    ff = FreeFile
    Open sFileName For Binary As #ff
    
    For i = 0 To lstBuffer.ListCount - 1
      For m = 1 To 16
        '\\ one line in lstBuffer looks like: '0000:' then one vbTab then hex data each
        '\\ separated with " ". we can get each hex value like this:
        Put #ff, , Chr$("&h" & Mid$(lstBuffer.List(i), (3 * m) + 4, 2))
      Next m
    Next i
    
    Close #ff
    Sbar.PanelText(2) = sFileName
    bDirty = False

  End If

End Sub
