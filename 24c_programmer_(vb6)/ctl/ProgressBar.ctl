VERSION 5.00
Begin VB.UserControl progressBar 
   BackStyle       =   0  'Transparent
   ClientHeight    =   225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4155
   ControlContainer=   -1  'True
   FillColor       =   &H8000000D&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   225
   ScaleWidth      =   4155
   Begin VB.Timer tmr_Speed 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2340
      Top             =   -60
   End
   Begin VB.PictureBox pict_bar_d 
      AutoRedraw      =   -1  'True
      DrawMode        =   8  'Xor Pen
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   4095
      TabIndex        =   0
      Top             =   0
      Width           =   4155
      Begin VB.PictureBox pict_bar 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   375
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   1335
         TabIndex        =   1
         Top             =   0
         Width           =   1335
      End
   End
End
Attribute VB_Name = "progressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Const SliderColor = &H8000000D        '2147483661
    Const inverzSliderColor = &H80FFFF    '8454143
    
    Enum look_bar
       decorative
        continued
        discontinued
    End Enum
    
    Private m_enulook As look_bar
    Private m_bytpercentage As Byte
    Private X, m_intNovaSpeed As Byte
    Private m_bolBorder, m_bolPrikazipercentage As Boolean
    
    Private m_Color As Long
    
    Public Event Click()
    Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Property Let look(n_look As look_bar)
  m_enulook = n_look
  pict_bar.Visible = True
  Select Case m_enulook
                  Case continued
                                    tmr_Speed.Enabled = False
                                    pict_bar.Move (0)
                                    ucrtajVrijednost_kontinuirana
                  Case decorative
                                    tmr_Speed.Enabled = tmr_Speed.Enabled
                                    UserControl_Resize
                                    kvadrat
                  Case discontinued
                                    tmr_Speed.Enabled = False
                                    pict_bar.Move (0)
                                    ucrtajVrijednost_diskontinuirana
   End Select
                                    
End Property

Property Get look() As look_bar
    look = m_enulook
End Property

Property Let Border(n_Border As Boolean)
    m_bolBorder = n_Border
    If m_bolBorder = True Then
                    pict_bar_d.BorderStyle = 1
                           Else
                    pict_bar_d.BorderStyle = 0
                           End If
End Property

Property Get Border() As Boolean
   Border = m_bolBorder
End Property

Property Let percentage(n_percentage As Byte)

If n_percentage > 100 Then Err.Raise (380)

m_bytpercentage = n_percentage
If m_enulook = continued Then
                        ucrtajVrijednost_kontinuirana
                              Else
                        ucrtajVrijednost_diskontinuirana
                               End If
End Property

Property Get percentage() As Byte
    percentage = m_bytpercentage
End Property

Property Let txtpercentage(n_txtpercentage As Boolean)
    m_bolPrikazipercentage = n_txtpercentage
 If m_enulook = continued Then ucrtajVrijednost_kontinuirana
End Property

Property Get txtpercentage() As Boolean
    txtpercentage = m_bolPrikazipercentage
End Property

Property Let Speed(n_Speed As Byte)
On Error GoTo greske

m_intNovaSpeed = n_Speed
Exit Property
greske:
    MsgBox Err.Description
End Property

Property Get Speed() As Byte
    Speed = m_intNovaSpeed
End Property

Property Let TimerStatus(n_TimerStatus As Boolean)
    tmr_Speed.Enabled = n_TimerStatus
End Property

Property Get TimerStatus() As Boolean
    TimerStatus = tmr_Speed.Enabled
End Property

Private Sub tmr_Speed_Timer()
  pict_bar.Move (X)
  kvadrat
  X = X + m_intNovaSpeed
  If X >= pict_bar_d.Width Then X = -pict_bar.Width
End Sub

Private Sub ucrtajVrijednost_kontinuirana()
    Dim percent As Single
    
    pict_bar_d.Cls
    pict_bar.Visible = False
    percent = pict_bar_d.Width * m_bytpercentage / 100

If m_bolPrikazipercentage = True Then
'THESE COMMANDES ARE COPIED FORM PROGRAM "ProgBar4"
'www.freevbcode.com/ShowCode.Asp?ID=317
    With pict_bar_d
       .CurrentX = (.ScaleWidth - .TextWidth(percentage & " %")) / 2
       .CurrentY = (.ScaleHeight - .TextHeight(percentage & " %")) / 2
        pict_bar_d.Print percentage & " %"
    End With
End If
pict_bar_d.Line (0, 0)-(percent, pict_bar.Height), inverzSliderColor, BF
End Sub

Private Sub ucrtajVrijednost_diskontinuirana()
    Dim percent As Single
    Dim brojKvadrata As Byte
    
    pict_bar.Cls
    pict_bar_d.Cls
    percent = pict_bar_d.Width * m_bytpercentage / 100
    brojKvadrata = CInt(percent / pict_bar.Height)
    pict_bar.Width = brojKvadrata * pict_bar.Height
    kvadrat
End Sub

Private Sub kvadrat()
    Dim pocetak As Integer
    
Do Until pocetak >= UserControl.pict_bar.Width
UserControl.pict_bar.Line (pocetak, 0)-(pocetak, UserControl.Height)
pocetak = pocetak + UserControl.pict_bar_d.Height
Loop
End Sub

Private Sub UserControl_Initialize()
    m_intNovaSpeed = 70
    m_enulook = discontinued
    m_bytpercentage = 55
    m_bolBorder = True
    pict_bar.DrawWidth = 2
    m_bolPrikazipercentage = True
End Sub

Private Sub UserControl_Paint()
'kada je dio svojstava PictureBox-a promjenit, pa sljedeæe
'dvije naredbe treba staviti i u svaku proceduru u kojoj se nešto iscrtava
 pict_bar.Cls
 pict_bar_d.Cls
Select Case m_enulook
                Case decorative
                            UserControl_Resize
                            kvadrat
                Case continued
                            ucrtajVrijednost_kontinuirana
                Case discontinued
                            ucrtajVrijednost_diskontinuirana
 End Select
End Sub

Private Sub UserControl_Resize()
    pict_bar_d.Width = UserControl.Width - 50
    pict_bar_d.Height = UserControl.Height
    pict_bar.Height = UserControl.Height
    pict_bar.Width = pict_bar_d.Height * 4
End Sub

Private Sub UserControl_Terminate()
    X = Empty
    m_intNovaSpeed = Empty
    m_bytpercentage = Empty
    m_enulook = Empty
End Sub

Private Sub pict_bar_d_Click()
 RaiseEvent Click
End Sub

Private Sub pict_bar_Click()
 RaiseEvent Click
End Sub

Private Sub pict_bar_d_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub pict_bar_d_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub pict_bar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub pict_bar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
m_bolPrikazipercentage = PropBag.ReadProperty("txtpercentage", True)
m_bolBorder = PropBag.ReadProperty("Border", True)
m_bytpercentage = PropBag.ReadProperty("percentage", 80)
m_enulook = PropBag.ReadProperty("look", 0)
m_intNovaSpeed = PropBag.ReadProperty("Speed", 70)
tmr_Speed.Enabled = PropBag.ReadProperty("TimerStatus", TimerStatus)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "txtpercentage", txtpercentage
PropBag.WriteProperty "Border", Border
PropBag.WriteProperty "percentage", percentage
PropBag.WriteProperty "look", look
PropBag.WriteProperty "Speed", Speed
PropBag.WriteProperty "TimerStatus", tmr_Speed.Enabled
End Sub






