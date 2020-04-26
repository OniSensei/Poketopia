VERSION 5.00
Object = "{50347EDF-F3EF-4392-AFDD-71AE67A3A978}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LVbutton.ocx"
Begin VB.Form frmIntro 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Welcome to Poketopia!"
   ClientHeight    =   150
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmIntro.frx":0000
   ScaleHeight     =   150
   ScaleWidth      =   885
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrpos 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5400
      Top             =   1200
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   3120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      Caption         =   ">"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   14737632
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgOak 
      Height          =   2295
      Left            =   2040
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4048
      Effects         =   "frmIntro.frx":63342
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   3000
      Width           =   4335
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl eveimg 
      Height          =   1215
      Left            =   2040
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2143
      Attr            =   1536
      Effects         =   "frmIntro.frx":6335A
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl StageImg 
      Height          =   855
      Left            =   840
      Top             =   2880
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1508
      Effects         =   "frmIntro.frx":63372
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl pokeback 
      Height          =   4095
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   7223
      Effects         =   "frmIntro.frx":6338A
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
 
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
 
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
 
Private Const WM_SYSCOMMAND = &H112
Private Const G = (-20)
Private Const LWA_COLORKEY = &H1        'to trans'
Private Const LWA_ALPHA = &H2           'to semi trans'
Private Const WS_EX_LAYERED = &H80000

Private HappySecs As Long
Dim CurrentMsg As Long
Dim OldLeft As Long
Dim pos As Long

Private Sub AlphaImgCtl9_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub cmdClose_Click(index As Integer)
Unload Me
End Sub

Private Sub Form_Activate()
    Me.BackColor = vbBlack
    If Options.FormTransparency = YES Then
    trans 0
    End If
End Sub

Private Sub trans(Level As Integer)
    Dim Msg As Long
    Msg = GetWindowLong(Me.hwnd, G)
    Msg = Msg Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, G, Msg
    SetLayeredWindowAttributes Me.hwnd, vbBlack, Level, LWA_ALPHA
End Sub




Private Sub Command1_Click()
Me.Caption = GetPlayerPosition
End Sub

Private Sub Form_Load()
CurrentMsg = 1

GoranPlay (App.Path & "\oak.wav")
tmrpos.Enabled = True
pos = 0
frmMainGame.Visible = True
frmChat.Visible = False
InIntro = True
DrawOak = True
frmMainGame.picDialog.Visible = True
OldLeft = 120
frmMainGame.picDialog.Left = 15
frmMainGame.lvButtons_H10.Visible = False


CanMoveNow = False
End Sub

Private Sub Form_Unload(Cancel As Integer)

frmMainGame.Visible = True
frmChat.Visible = True
StopPlay
End Sub

Private Sub lvButtons_H1_Click()
PlayClick
frmMainGame.Visible = True
frmChat.Visible = True
StopPlay
Unload frmIntro
PlayMapMusic MapMusic
End Sub

Sub SetText(ByVal text As String)
Label1.Caption = text
End Sub

Private Sub tmrpos_Timer()
Dim i As Long
i = GetPlayerPosition
'Me.Caption = i

'1
If i >= 1 And i <= 4 And pos <> 1 Then

frmMainGame.DisplayDialogText "Hello there , and welcome to the world of Pokemon!"
pos = 1
End If
'2
If i >= 5 And i <= 7 And pos <> 2 Then
frmMainGame.DisplayDialogText "My name is professor Oak."
pos = 2
End If
'3
If i >= 8 And i <= 10 And pos <> 3 Then
frmMainGame.DisplayDialogText "I'm often referd to as the prof. of Pokemon."
pos = 3
End If
'4
If i >= 11 And i <= 13 And pos <> 4 Then
frmMainGame.DisplayDialogText "They are some very interesting creatures,"
pos = 4
End If
'5
If i >= 14 And i <= 15 And pos <> 5 Then
frmMainGame.DisplayDialogText "that inhabit this world."
pos = 5
End If
'6
If i >= 16 And i <= 20 And pos <> 6 Then
frmMainGame.DisplayDialogText "And as you maybe aware they are known the world over quite simply as Pokemon."
pos = 6
End If
'7
If i >= 21 Then
pokeback.Visible = True

End If

If i >= 21 And i <= 27 And pos <> 7 Then
frmMainGame.DisplayDialogText "For some people pokemon are pets, while other people enjoy using them for battle,"
pos = 7
End If
'8
If i >= 28 And i <= 31 And pos <> 8 Then
frmMainGame.DisplayDialogText "we find ourselves coexisting various styles."
pos = 8

End If
If i > 31 Then
'frmMainGame.txtDialog.Visible = False
pokeback.Visible = False
End If
'9
If i >= 34 Then
DrawEevee = True
pokeback.Visible = False
End If
'10
If i >= 37 And i <= 40 And pos <> 9 Then
'frmMainGame.txtDialog.Visible = True
frmMainGame.DisplayDialogText "Are you interested in learning about this pokemon world?"
pokeback.Visible = False
pos = 9
End If
'11
If i >= 41 And pos <> 10 Then
frmMainGame.DisplayDialogText "I'll be your guide for a story of dreams and adventures."
pokeback.Visible = False
pos = 10
End If

If i = 45 And pos <> 11 Then

Label1.Caption = "Press > to continue. Good luck!"
StopPlay
pokeback.Visible = False
tmrpos.Enabled = False
'PlayMapMusic MapMusic
InIntro = False
DrawOak = False
DrawEevee = False
frmMainGame.picDialog.Visible = False
frmMainGame.picDialog.Left = OldLeft
frmMainGame.lvButtons_H10.Visible = True
frmMainGame.DisplayDialogText "Welcome to PEO! Enjoy,have fun and catch some pokemon!"
frmMainGame.picDialog.Visible = True
pos = 11
CanMoveNow = True
Unload Me
End If
End Sub


