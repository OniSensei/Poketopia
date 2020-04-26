VERSION 5.00
Begin VB.Form frmCharacter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Character"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCharacter.frx":0000
   ScaleHeight     =   4425
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgPokemon 
      Height          =   975
      Index           =   5
      Left            =   3600
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Image imgPokemon 
      Height          =   975
      Index           =   4
      Left            =   2040
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Image imgPokemon 
      Height          =   975
      Index           =   3
      Left            =   5160
      Top             =   840
      Width           =   1455
   End
   Begin VB.Image imgPokemon 
      Height          =   975
      Index           =   2
      Left            =   3600
      Top             =   840
      Width           =   1455
   End
   Begin VB.Image imgPokemon 
      Height          =   975
      Index           =   6
      Left            =   5160
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Image imgPokemon 
      Height          =   975
      Index           =   1
      Left            =   2040
      Top             =   840
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1200
      Left            =   360
      Picture         =   "frmCharacter.frx":62E02
      Top             =   1320
      Width           =   1200
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "frmCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Sub loadPlayerData(ByVal Index As Long)
lblName = Player(Index).Name
Dim pokenum() As Long
Dim i As Long
Dim a As Long


For a = 1 To 6
If PokemonInstance(a).PokemonNumber <= 0 Then
Set imgPokemon(a).Picture = Nothing
Else
imgPokemon(a).Picture = LoadPicture(App.Path & "\Data Files\graphics\pokemonsprites\" & PokemonInstance(a).PokemonNumber & ".gif")
End If
Next

End Sub

