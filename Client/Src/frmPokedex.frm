VERSION 5.00
Object = "{50347EDF-F3EF-4392-AFDD-71AE67A3A978}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmPokedex 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pokedex"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   405
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   548
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstMoves 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Eurostar"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4440
      Left            =   5040
      TabIndex        =   9
      Top             =   960
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   6075
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Moves:"
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
      Left            =   4920
      TabIndex        =   10
      Top             =   480
      Width           =   3015
   End
   Begin VB.Image imgType1 
      Height          =   375
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Image imgType2 
      Height          =   375
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1215
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgPoke 
      Height          =   1815
      Left            =   2400
      Top             =   3240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   3201
      Attr            =   1536
      Effects         =   "frmPokedex.frx":0000
   End
   Begin VB.Label lblSPEED 
      BackStyle       =   0  'Transparent
      Caption         =   "SPEED"
      BeginProperty Font 
         Name            =   "Eurostar"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Label lblSPDEF 
      BackStyle       =   0  'Transparent
      Caption         =   "SP.DEF"
      BeginProperty Font 
         Name            =   "Eurostar"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label lblSPATK 
      BackStyle       =   0  'Transparent
      Caption         =   "SP:ATK"
      BeginProperty Font 
         Name            =   "Eurostar"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label lblDEF 
      BackStyle       =   0  'Transparent
      Caption         =   "DEF"
      BeginProperty Font 
         Name            =   "Eurostar"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label lblATK 
      BackStyle       =   0  'Transparent
      Caption         =   "ATK"
      BeginProperty Font 
         Name            =   "Eurostar"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label lblHP 
      BackStyle       =   0  'Transparent
      Caption         =   "HP"
      BeginProperty Font 
         Name            =   "Eurostar"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Base stats"
      BeginProperty Font 
         Name            =   "Eurostar"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label lblPOKE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Bulbasaur"
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
      Left            =   2040
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
End
Attribute VB_Name = "frmPokedex"
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
loadAllPokes
loadpoke (1)
End Sub

Private Sub lblnum_Click()

End Sub

Private Sub ShinyImage_Click()

End Sub

Sub loadpoke(ByVal poke As Long)
lblPOKE.Caption = Trim$(Pokemon(poke).Name)
lblHP.Caption = "HP:" & Pokemon(poke).MaxHp
lblATK.Caption = "ATK:" & Pokemon(poke).ATK
lblDEF.Caption = "DEF:" & Pokemon(poke).DEF
lblSPATK.Caption = "SP.ATK:" & Pokemon(poke).SPATK
lblSPDEF.Caption = "SP.DEF:" & Pokemon(poke).SPDEF
lblSPEED.Caption = "SPEED:" & Pokemon(poke).SPD
imgPoke.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\" & poke & ".gif")
imgType1.Picture = LoadPicture(App.Path & "\Data Files\graphics\types\" & Pokemon(poke).Type & ".bmp")
imgType2.Picture = LoadPicture(App.Path & "\Data Files\graphics\types\" & Pokemon(poke).Type2 & ".bmp")
LoadPokeMoves (poke)
End Sub

Sub LoadPokeMoves(ByVal slot As Long)
lstMoves.Clear
Dim i As Long
Dim pokenum As Long
pokenum = slot
For i = 1 To 30
If Pokemon(pokenum).moves(i) > 0 Then
lstMoves.AddItem (Trim$(PokemonMove(Pokemon(pokenum).moves(i)).Name) & " - Lv." & Pokemon(pokenum).movesLV(i))
End If
Next
End Sub
Sub loadAllPokes()
List1.Clear
Dim i As Long
For i = 1 To MAX_POKEMONS
List1.AddItem (i & ": " & Trim$(Pokemon(i).Name))
Next
End Sub

Private Sub List1_Click()
loadpoke (List1.ListIndex + 1)
End Sub
