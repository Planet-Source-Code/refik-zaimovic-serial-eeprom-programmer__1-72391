VERSION 5.00
Begin VB.Form frmEdit 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3315
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   3315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin i2c_prog.GGCommand cmdExit 
      Height          =   375
      Left            =   180
      TabIndex        =   9
      Top             =   2280
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
         Charset         =   238
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
      Left            =   1920
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Ok"
      UseHoverProperties=   0   'False
      FontColorOnHover=   0
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
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
   Begin VB.Frame frmInfo 
      Caption         =   "Data"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   3075
      Begin VB.TextBox txtBin 
         Height          =   345
         Left            =   1740
         MaxLength       =   8
         TabIndex        =   4
         Top             =   1440
         Width           =   1155
      End
      Begin VB.TextBox txtAsc 
         Height          =   345
         Left            =   1740
         MaxLength       =   1
         TabIndex        =   3
         Top             =   600
         Width           =   1155
      End
      Begin VB.TextBox txtHex 
         Height          =   345
         Left            =   180
         MaxLength       =   2
         TabIndex        =   2
         Top             =   1440
         Width           =   1155
      End
      Begin VB.TextBox txtDec 
         Height          =   345
         Left            =   180
         MaxLength       =   3
         TabIndex        =   1
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label lblInfo 
         Caption         =   "Binary"
         Height          =   195
         Index           =   3
         Left            =   1800
         TabIndex        =   8
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         Caption         =   "Hexadecimal"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         Caption         =   "Ascii"
         Height          =   195
         Index           =   1
         Left            =   1800
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         Caption         =   "Decimal"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    
  bIsChanged = False
  Unload Me
  Set frmEdit = Nothing

End Sub

Private Sub cmdOk_Click()
    
  bDirty = True
  sHex = Right$("0" & UCase(txtHex.Text), 2)
  Unload Me
  Set frmEdit = Nothing
    
End Sub

Private Sub Form_Load()

  frmEdit.Caption = "Edit data" ' at address: " & Left$(sLine, 4) & "h"
  bIsChanged = False                                '\\ reset var
  txtDec.Text = CStr(bByte)                         '\\ decimal
  txtHex.Text = Right$("0" & CStr(Hex(bByte)), 2)   '\\ hex
  txtAsc.Text = Chr$(bByte)                         '\\ ascii
  txtBin.Text = Bin_Conv(bByte)                     '\\ bin
  
  bByte = CByte(txtDec.Text)
  
End Sub

Private Function Bin_Conv(iDec As Integer) As String
    
  Bin_Conv = Trim(CStr(iDec Mod 2))
  iDec = iDec \ 2
  
  Do While iDec <> 0
    Bin_Conv = Trim(CStr(iDec Mod 2)) & Bin_Conv
    iDec = iDec \ 2
  Loop
  
  Bin_Conv = Right$("0000000" & Bin_Conv, 8)
    
End Function

Private Sub txtAsc_KeyUp(KeyCode As Integer, Shift As Integer)

  Dim i As Integer
  
  i = Asc(txtAsc.Text)
  txtDec.Text = i
  txtHex.Text = Hex(i)
  txtBin.Text = Bin_Conv(i)
  bIsChanged = True
    
End Sub

Private Sub txtBin_KeyPress(KeyAscii As Integer)

  If KeyAscii <> 48 And KeyAscii <> 49 Then KeyAscii = 0
    
End Sub

Private Sub txtBin_KeyUp(KeyCode As Integer, Shift As Integer)

  Dim i As Integer
  
  i = Bin_To_Dec(txtBin.Text)
  txtDec.Text = i
  txtHex.Text = Hex$(i)
  txtAsc.Text = Chr$(i)
  bIsChanged = True
    
End Sub

Private Function Bin_To_Dec(sBin As String) As Integer

  Dim c As Integer
  Dim i As Integer
  Dim m As Integer
  
  i = Len(txtBin.Text)
  m = 0
  For c = i To 1 Step -1
    Bin_To_Dec = Bin_To_Dec + (CInt(Mid$(txtBin.Text, c, 1))) * 2 ^ m
    m = m + 1
  Next c
    
End Function

Private Sub txtDec_GotFocus()

  txtDec.SelStart = 0
  txtDec.SelLength = Len(txtDec.Text)
    
End Sub

Private Sub txtAsc_GotFocus()

  txtAsc.SelStart = 0
  txtAsc.SelLength = Len(txtAsc.Text)
    
End Sub

Private Sub txtBin_GotFocus()

  txtBin.SelStart = 0
  txtBin.SelLength = Len(txtBin.Text)
    
End Sub

Private Sub txtDec_KeyPress(KeyAscii As Integer)

  If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
    
End Sub

Private Sub txtDec_KeyUp(KeyCode As Integer, Shift As Integer)

  Dim i As Integer
  
  If CInt(txtDec.Text) <= 255 Then
    i = CInt(txtDec.Text)
    txtHex.Text = Hex$(i)
    txtAsc.Text = Chr$(i)
    txtBin.Text = Bin_Conv(i)
    bIsChanged = True
  End If
    
End Sub

Private Sub txtHex_GotFocus()

  txtHex.SelStart = 0
  txtHex.SelLength = Len(txtHex.Text)
    
End Sub

Private Sub txtHex_KeyPress(KeyAscii As Integer)

  If InStr("1234567890abcdefABCDEF", Chr(KeyAscii)) = 0 Then KeyAscii = 0
  
End Sub

Private Sub txtHex_KeyUp(KeyCode As Integer, Shift As Integer)

  Dim i As Integer
  
  i = Fix("&H" & txtHex.Text)     'val,int
  txtDec.Text = i
  txtAsc.Text = Chr$(i)
  txtBin.Text = Bin_Conv(i)
  bIsChanged = True
    
End Sub
