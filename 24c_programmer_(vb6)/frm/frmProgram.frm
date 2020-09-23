VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmProgram 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Status monitor"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProgram.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin i2c_prog.GGCommand cmdAccept 
      Height          =   375
      Left            =   2820
      TabIndex        =   5
      Top             =   1980
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Accept"
      UseHoverProperties=   0   'False
      FontColorOnHover=   0
      ForeColor       =   -2147483631
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
      GradientInverseOnPress=   0   'False
      Enabled         =   0   'False
   End
   Begin i2c_prog.GGCommand cmdStop 
      Height          =   375
      Left            =   180
      TabIndex        =   4
      Top             =   1980
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Stop"
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
      GradientInverseOnPress=   0   'False
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2280
      Top             =   2040
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   1680
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      BaudRate        =   2400
   End
   Begin VB.CheckBox chkClose 
      Caption         =   "Close this dialog when process finished"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   2460
      Width           =   4215
   End
   Begin i2c_prog.progressBar PBar 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   450
      txtpercentage   =   -1  'True
      Border          =   -1  'True
      percentage      =   0
      look            =   1
      Speed           =   70
      TimerStatus     =   0   'False
   End
   Begin VB.TextBox txtInfo 
      Height          =   975
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   420
      Width           =   4095
   End
   Begin VB.Label lblInfo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0     '\\ array base = 0

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const IDChip        As Integer = &HA0   '\\ 10100000 - physical address of 24Cxx

Private bStop       As Boolean  '\\ presed stop
Private IsBlank     As Boolean  '\\ to check if device is blank
Private i           As Long     '\\ counter
Private pg          As Single   '\\ progress bar and index
Private k           As Long     '\\ counter
Private aData()     As Byte     '\\ data array
Private sAdr        As String   '\\ adress for buffer
Private sHexTemp    As String   '\\ temporary for hex data
Private sAscTemp    As String   '\\ temporary for ascii data
Private aTempVer()  As Byte     '\\ temporary if verify is requested

Private Sub chkClose_Click()

  SaveSetting App.Title, "Settings", "AutoClose", chkClose.Value
  
End Sub

Private Sub cmdAccept_Click()

  Unload Me
  
End Sub

Private Sub Form_Load()

  chkClose = GetSetting(App.Title, "Settings", "AutoClose", 0)
  bStop = False
  txtInfo.Text = " -= Com " & iComPort & " selected =- " & vbCrLf

  Select Case EepromAction
  
    Case Blank_Check
      lblInfo.Caption = "Blank check - " & frmMain.cboMemory.Text
      txtInfo.Text = txtInfo.Text & "Blank check" & "   - > " _
                                    & ((2 ^ iMemory) * 128) & " byte" & vbCrLf
    Case Read_Memory
      lblInfo.Caption = "Read memory - " & frmMain.cboMemory.Text
      txtInfo.Text = txtInfo.Text & "Reading Data" & "   - > " _
                                  & ((2 ^ iMemory) * 128) & " byte" & vbCrLf

    Case Verify_Data
      lblInfo.Caption = "Verify data - " & frmMain.cboMemory.Text
      txtInfo.Text = txtInfo.Text & "Verify data" & "   - > " _
                                  & ((2 ^ iMemory) * 128) & " byte" & vbCrLf
    
    Case Write_Memory
      lblInfo.Caption = "Write memory - " & frmMain.cboMemory.Text
      txtInfo.Text = txtInfo.Text & "I2C bus delay" & "   - > " _
                                  & i2cDelay & " ms" & vbCrLf
      txtInfo.Text = txtInfo.Text & "Start delay" & "   - > " _
                                  & iStartDelay & " ms" & vbCrLf
      txtInfo.Text = txtInfo.Text & "Memory size" & "   - > " _
                                  & ((2 ^ iMemory) * 128) & " byte" & vbCrLf

  End Select
  
End Sub

Private Sub cmdStop_Click()

  bStop = True
  
End Sub

Private Sub Form_Resize()

  txtInfo.SelStart = Len(txtInfo.Text)
  txtInfo.SelLength = 0
  
  If EepromAction = Write_Memory Then
    Write_i2c
  Else
    Read_i2c
  End If
  
  Close_Port
  cmdAccept.Enabled = True
  cmdAccept.ForeColor = &H80000012
  If chkClose.Value = 1 Then Timer1.Enabled = True
   
End Sub

Private Sub Form_Unload(Cancel As Integer)

  Set frmProgram = Nothing
  
End Sub

Private Sub Timer1_Timer()

  Timer1.Enabled = False
  Unload Me

End Sub

Private Sub Send_8_I2c(adr As Integer)

  Dim n As Single
  
  n = 128                                 '\\ 8-bit value

  Do                                      '\\
    MSComm.RTSEnable = False              '\\ SCL=0
    MSComm.DTREnable = IIf((adr And n) = n, True, False)  '\\ State of bit adress[i]
    MSComm.RTSEnable = True               '\\ SCL=1
    n = n / 2                             '\\ Value of the bit n-1
  Loop Until n = 0.5                      '\\
  
    MSComm.RTSEnable = False               '\\ SCL=0
    MSComm.DTREnable = True                '\\ SDA=1
    MSComm.RTSEnable = True                '\\ SCL=1

End Sub

Private Function Receive_8_I2c() As Integer  '\\ Read 8 bits on bus I2C

  Dim n As Single

  Receive_8_I2c = 0
  n = 128                               '\\ 8 bit Value
  MSComm.RTSEnable = False              '\\ SCL=0
    
  Do                                    '\\ To pass to the following bit   (7..0)
    MSComm.RTSEnable = True             '\\ SCL=1
    
    If MSComm.DSRHolding = True Then
      Receive_8_I2c = Receive_8_I2c + n '\\ DSR=SDA(7..0)
    End If
    
    MSComm.RTSEnable = False            '\\ SCL=0
    n = n / 2                           '\\ Value of the bit n-1
  Loop Until n = 0.5                    '\\ If 8 bits not all received, start again
    
End Function

Private Function Acknowlege_I2C(i2c_adr As Long) As Boolean

  Acknowlege_I2C = False

  '\\ Test answer memory, Pressed Cancel
  If MSComm.DSRHolding = True Or bStop = True Then
    Acknowlege_I2C = True   '\\ Memory not respond
    Close_Port
    
    '\\ send msg
    txtInfo.Text = txtInfo.Text & "Reading aborted by user or memory not responding!"
    
    '\\ move cursor to the end of msg
    txtInfo.SelStart = Len(txtInfo.Text)
    txtInfo.SelLength = 0
    
    '\\ Post error massage
    MsgBox "Error at address: " & Right("000" & Hex(i2c_adr), 4) & "h"
'    PBar.percentage = 100
  End If
  
End Function

Private Sub Open_Port()

  If MSComm.PortOpen Then MSComm.PortOpen = False
  
  MSComm.CommPort = iComPort
  
  MSComm.PortOpen = True
  MSComm.RTSEnable = True           '\\ SCL=1   RTS=SCL
  MSComm.DTREnable = True           '\\ SDA=1   DTR=SDA

End Sub

Private Sub Init_I2C()

  MSComm.RTSEnable = False         '\\ SCL=0
  MSComm.DTREnable = False         '\\ SDA=0
    
End Sub

Private Sub Start_I2C()

  MSComm.DTREnable = True          '\\ SDA=1
  MSComm.RTSEnable = True          '\\ SCL=1
  MSComm.DTREnable = False         '\\ SDA=0
  MSComm.RTSEnable = False         '\\ SCL=0
    
End Sub

Private Sub Stop_I2C()

  MSComm.RTSEnable = False         '\\ SCL=0
  MSComm.DTREnable = False         '\\ SDA=0
  MSComm.RTSEnable = True          '\\ SCL=1
  MSComm.DTREnable = True          '\\ SDA=1
    
End Sub

Private Sub Close_Port()

  If MSComm.PortOpen Then MSComm.PortOpen = False

End Sub

Private Sub Prepare_Data()  '\\ prepare array to compare with
  
  ReDim aTempVer(((2 ^ iMemory) * 128) - 1) '\\ redim array with memory size
  
  For i = 0 To ((2 ^ iMemory) * 8) - 1  '\\ move line by line
    For k = 1 To 16                     '\\ get all characters in line
      aTempVer((k + (i * 16)) - 1) = Int("&h" & Mid$(frmMain.lstBuffer.List(i), (3 * k) + 4, 2))
    Next k
  Next i
    
End Sub

Private Sub Fill_Buffer() '\\ fill buffer

  pg = 0                                                  '\\ index for lstBuffer
  frmMain.lstBuffer.Clear                                 '\\ clear buffer
  
  For i = 0 To (UBound(aData) / 16) - 1                   '\\ loop (memory size/16) times
  
    sAdr = Right$("0000" & CStr(Hex(i * 16)), 4) & ":"    '\\ address for data
    
    For k = i * 16 To (i * 16) + 15                       '\\ first line of data
    
      sHexTemp = sHexTemp & _
                 Right$("00" & Hex(aData(k)), 2) & " "    '\\ one line of hex data
      
      '\\ replace some characters in ascii part (see 'lstBuffer_DblClick' routine)
      sAscTemp = sAscTemp & IIf(aData(k) = 160 Or aData(k) = 0 Or aData(k) = 9 _
      Or aData(k) = 32, Chr$(46), Chr$(aData(k)))         '\\ one line of ascii data
         
    Next k
    
    frmMain.lstBuffer.AddItem sAdr & vbTab & RTrim$(sHexTemp) & vbTab & vbTab & sAscTemp, pg

    sAscTemp = vbNullString
    sHexTemp = vbNullString

    pg = pg + 1   '\\ increase index for lstBuffer

  Next i
  
End Sub

Private Sub Read_i2c()  '\\ read, verify and blank check

  If EepromAction = Verify_Data Then Prepare_Data
  
  ReDim aData(((2 ^ iMemory) * 128) - 1)    '\\ redim array with memory size
  
  Open_Port
  DoEvents
  Sleep (iStartDelay)                       '\\ without this can not write first 8 bytes

  If iMemory > 4 Then                                         '\\ Memory 24c32 - 24c128
  
    Init_I2C
    Start_I2C
    Send_8_I2c (IDChip And &HFE)  ':If Acknowlege_I2C(i) = True Then Exit Sub
    Send_8_I2c (0)                ':If Acknowlege_I2C(i) = True Then Exit Sub
    Send_8_I2c (0)                ':If Acknowlege_I2C(i) = True Then Exit Sub
    Init_I2C
    
    For i = 0 To ((2 ^ iMemory) * 128) - 1
      DoEvents
      pg = pg + 100 / ((2 ^ iMemory) * 128): PBar.percentage = Int(pg) '\\ calc:set percentage
      Start_I2C
      Send_8_I2c (IDChip Or 1)
      If Acknowlege_I2C(i) = True Then Exit Sub               '\\ test ack and send msg
      aData(i) = Receive_8_I2c
      
      If EepromAction = Verify_Data Then
        If aData(i) <> aTempVer(i) Then
          MsgBox "Verify failed at address: " & Right("000" & Hex(i), 4) & "h"
          txtInfo.Text = txtInfo.Text & "Verify memory   - > Failed" & vbCrLf
          txtInfo.SelStart = Len(txtInfo.Text)
          txtInfo.SelLength = 0
          PBar.percentage = 100
          Exit Sub
        End If
      End If
      
      If EepromAction = Blank_Check Then
        If aData(i) <> 255 Then
          MsgBox "Blank check failed at address: " & Right("000" & Hex(i), 4) & "h"
          txtInfo.Text = txtInfo.Text & "Blank check   - > Failed" & vbCrLf
          PBar.percentage = 100
          Exit Sub
        End If
      End If
    Next i
  
  Else                                                        '\\ Memory 24c01 - 24c16
    
    For i = 0 To (((2 ^ iMemory) * 128) / 8) - 1
      Init_I2C
      Start_I2C
      Send_8_I2c ((IDChip And &HFE) Or (Int(i / 32) * 2)) ':If Acknowlege_I2C((i * 8) + k) = True Then Exit Sub
      Send_8_I2c (i * 8)                                  ':If Acknowlege_I2C((i * 8) + k) = True Then Exit Sub
      Init_I2C

      For k = 0 To 7                                          '\\ send 8 bit package
        DoEvents
        pg = pg + (100 / ((2 ^ iMemory) * 128)): PBar.percentage = Int(pg) '\\ calc:set percentage
    
        Start_I2C
        Send_8_I2c (IDChip Or 1)                              '\\ chip ID
        If Acknowlege_I2C((i * 8) + k) = True Then Exit Sub   '\\ test ack and send msg
        aData((i * 8) + k) = Receive_8_I2c
        
        If EepromAction = Verify_Data Then
          If aData((i * 8) + k) <> aTempVer((i * 8) + k) Then
            MsgBox "Verify failed at address: " & Right("000" & Hex((i * 8) + k), 4) & "h"
            txtInfo.Text = txtInfo.Text & "Verify memory   - > Failed" & vbCrLf
            txtInfo.SelStart = Len(txtInfo.Text)
            txtInfo.SelLength = 0
            PBar.percentage = 100
            Exit Sub
          End If
        End If
        
      If EepromAction = Blank_Check Then
        If aData((i * 8) + k) <> 255 Then
          MsgBox "Blank check failed at address: " & Right("000" & Hex((i * 8) + k), 4) & "h"
          txtInfo.Text = txtInfo.Text & "Blank check   - > Failed" & vbCrLf
          PBar.percentage = 100
          Exit Sub
        End If
      End If
      Next k
    Next i
        
  End If
  
  Stop_I2C

  Select Case EepromAction
  
    Case Read_Memory
      Fill_Buffer
      txtInfo.Text = txtInfo.Text & "Read memory   - > OK" & vbCrLf
      
    Case Verify_Data
      txtInfo.Text = txtInfo.Text & "Verify memory   - > OK" & vbCrLf
      txtInfo.SelStart = Len(txtInfo.Text)
      txtInfo.SelLength = 0

    Case Blank_Check
      txtInfo.Text = txtInfo.Text & "Device is blank" & vbCrLf
      
  End Select
    
  Erase aData(): Erase aTempVer() '\\ destroy array

End Sub

Private Sub Write_i2c()

  Prepare_Data

  Open_Port
  DoEvents
  Sleep (iStartDelay)                 '\\ without this can not write first 8 bytes

  If iMemory > 4 Then                 '\\ memory 24c32....24c128

    Init_I2C
    For i = 0 To ((2 ^ iMemory) * 128) - 1

      DoEvents
      pg = pg + 100 / ((2 ^ iMemory) * 128): PBar.percentage = Int(pg)  'pg = pg + (100 / (iMaxLenght + 3))
      Start_I2C
      
      Send_8_I2c (IDChip And &HFE)              '\\ chip ID (mode W)
      If Acknowlege_I2C(i) = True Then Exit Sub '\\ test ack and send msg
      
      Send_8_I2c (Int(i / 256))                 '\\ eeprom address page
      If Acknowlege_I2C(i) = True Then Exit Sub '\\ test ack and send msg
      
      Send_8_I2c (i)                            '\\ address
      If Acknowlege_I2C(i) = True Then Exit Sub '\\ test ack and send msg
      
      Send_8_I2c (aTempVer(i))                  '\\ Value
      If Acknowlege_I2C(i) = True Then Exit Sub '\\ test ack and send msg
      
      Stop_I2C
      Sleep (i2cDelay)

    Next i

  Else                                            '\\ memory 24c01....24c16

    For i = 0 To (((2 ^ iMemory) * 128) / 8) - 1

      Init_I2C
      Start_I2C
      Send_8_I2c ((IDChip And &HFE) Or (Int(i / 32)) * 2)
      If Acknowlege_I2C(i) = True Then Exit Sub   '\\ test ack and send msg
      Send_8_I2c (i * 8)
      If Acknowlege_I2C(i) = True Then Exit Sub   '\\ test ack and send msg

      For k = 0 To 7                              '\\ 8 bit data package
        Send_8_I2c (aTempVer((i * 8) + k))
        If Acknowlege_I2C(i) = True Then Exit Sub '\\ test ack and send msg
        DoEvents
        pg = pg + (100 / ((2 ^ iMemory) * 128)): PBar.percentage = Int(pg) 'pg = pg + (100 / (iMaxLenght + 3))
      Next k

      Stop_I2C
      Sleep (i2cDelay)

    Next i

  End If

  Erase aTempVer(): Erase aData()   '\\ clear arrays
  
  txtInfo.Text = txtInfo.Text & "Write memory   - > OK" & vbCrLf
  txtInfo.SelStart = Len(txtInfo.Text)
  txtInfo.SelLength = 0

  If iVerAfter = 1 Then   '\\ is verify after checked
    Sleep (300)           '\\ set little pause (no need realy)
    pg = 0                '\\ reset variable for progress bar
    PBar.percentage = 0
    EepromAction = Verify_Data  '\\ set action
    txtInfo.Text = txtInfo.Text & "Verifying data" & "   - > " & ((2 ^ iMemory) * 128) & " byte" & vbCrLf
    txtInfo.SelStart = Len(txtInfo.Text)
    txtInfo.SelLength = 0
    Read_i2c              '\\ read and verify eeprom
  End If

End Sub
