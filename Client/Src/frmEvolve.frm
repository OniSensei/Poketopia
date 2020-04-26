VERSION 5.00
Object = "{50347EDF-F3EF-4392-AFDD-71AE67A3A978}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LVbutton.ocx"
Begin VB.Form frmEvolve 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Evolution"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3360
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      Caption         =   "Evolve"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cBhover         =   8421504
      LockHover       =   1
      cGradient       =   0
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   3360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Close"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16777215
      cFHover         =   16777215
      cBhover         =   8421504
      LockHover       =   1
      cGradient       =   0
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      cBack           =   255
   End
   Begin VB.Label lblnum 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TO"
      BeginProperty Font 
         Name            =   "Eurostar"
         Size            =   15.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgNewPoke 
      Height          =   2775
      Left            =   3000
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4895
      Effects         =   "frmEvolve.frx":0000
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgOldPoke 
      Height          =   2775
      Left            =   120
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4895
      Effects         =   "frmEvolve.frx":0018
   End
End
Attribute VB_Name = "frmEvolve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'function to make transparent'
Option Explicit
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

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
Dim pokeslot As Long
Dim evNewPoke As Long
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

Private Sub lvButtons_H1_Click()
Call SendRequest(pokeslot, evNewPoke, "", "PEV")
Unload Me
End Sub

Private Sub lvButtons_H2_Click()
Unload Me
End Sub
Public Sub LoadEvolution(ByVal slot As Long, ByVal newpoke As Long)
If PokemonInstance(slot).isShiny = YES Then
imgOldPoke.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\Shiny\" & PokemonInstance(slot).PokemonNumber & ".gif")
imgNewPoke.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\Shiny\" & newpoke & ".gif")
imgOldPoke.Animate (lvicAniCmdStart)
imgNewPoke.Animate (lvicAniCmdStart)
Else
imgOldPoke.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\" & PokemonInstance(slot).PokemonNumber & ".gif")
imgNewPoke.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\" & newpoke & ".gif")
imgOldPoke.Animate (lvicAniCmdStart)
imgNewPoke.Animate (lvicAniCmdStart)
End If
pokeslot = slot
evNewPoke = newpoke
End Sub
