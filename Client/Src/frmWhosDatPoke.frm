VERSION 5.00
Object = "{50347EDF-F3EF-4392-AFDD-71AE67A3A978}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LVbutton.ocx"
Begin VB.Form frmWhosDatPoke 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Whos dat pokemon?"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   120
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   240
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Text            =   "Name of pokemon"
      Top             =   2880
      Width           =   3375
   End
   Begin lvButton.lvButtons_H lvButtons_H3 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   3240
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      Caption         =   "Ok"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgPoke 
      Height          =   1335
      Index           =   0
      Left            =   1800
      Top             =   600
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   2355
      Image           =   "frmWhosDatPoke.frx":0000
      Settings        =   100
      Effects         =   "frmWhosDatPoke.frx":142BB
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgPoke 
      Height          =   1335
      Index           =   4
      Left            =   1440
      Top             =   480
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   2355
      Image           =   "frmWhosDatPoke.frx":142D3
      Angle           =   -224
      Settings        =   167772260
      Effects         =   "frmWhosDatPoke.frx":2858E
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgPoke 
      Height          =   1335
      Index           =   3
      Left            =   1440
      Top             =   840
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   2355
      Image           =   "frmWhosDatPoke.frx":285A6
      Angle           =   68
      Settings        =   167772260
      Effects         =   "frmWhosDatPoke.frx":3C861
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgPoke 
      Height          =   1335
      Index           =   2
      Left            =   2160
      Top             =   960
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   2355
      Image           =   "frmWhosDatPoke.frx":3C879
      Angle           =   -205
      Settings        =   167772260
      Effects         =   "frmWhosDatPoke.frx":50B34
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgPoke 
      Height          =   1335
      Index           =   1
      Left            =   2280
      Top             =   480
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   2355
      Image           =   "frmWhosDatPoke.frx":50B4C
      Angle           =   127
      Settings        =   167772260
      Effects         =   "frmWhosDatPoke.frx":64E07
   End
End
Attribute VB_Name = "frmWhosDatPoke"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'function to make transparent'
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
  ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const G = (-20)
Private Const LWA_COLORKEY = &H1        'to trans'
Private Const LWA_ALPHA = &H2           'to semi trans'
Private Const WS_EX_LAYERED = &H80000
Private Sub Form_Activate()
    Me.BackColor = vbBlack
   If Options.FormTransparency = YES Then
    trans 215
    End If
End Sub

Private Sub trans(Level As Integer)
    Dim Msg As Long
    Msg = GetWindowLong(Me.hwnd, G)
    Msg = Msg Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, G, Msg
    SetLayeredWindowAttributes Me.hwnd, vbBlack, Level, LWA_ALPHA
End Sub

Private Sub Form_Load()
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Private Sub lvButtons_H3_Click()
SendRequest 0, 0, Trim$(Text1.text), "WHOSDATPOKE"
Unload Me
End Sub

Private Sub Timer1_Timer()
imgPoke(0).Rotation = imgPoke(0).Rotation + 1
If imgPoke(0).Rotation = 360 Then
imgPoke(0).Rotation = -360
End If
End Sub

Private Sub Timer2_Timer()
If imgPoke(0).LightnessPct > 0 Then
imgPoke(0).LightnessPct = imgPoke(0).LightnessPct - 1
End If
End Sub

Sub LoadPoke(ByVal num As Long)

Dim i As Long
For i = 0 To 4
imgPoke(i).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\" & num & ".gif")
Next
imgPoke(0).LightnessPct = 100
imgPoke(0).Rotation = 0
End Sub
