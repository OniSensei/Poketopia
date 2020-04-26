VERSION 5.00
Begin VB.Form frmPlayer 
   Caption         =   "Edit Player"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Save"
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   5280
      Width           =   4455
   End
   Begin VB.ListBox lstPlayers 
      Height          =   4545
      Left            =   4680
      TabIndex        =   13
      Top             =   600
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      Caption         =   "Storage"
      Height          =   2415
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   4335
      Begin VB.CommandButton rmvst 
         Caption         =   "Remove"
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   975
      End
      Begin VB.ListBox storagee 
         Height          =   1035
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pokemon"
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   4335
      Begin VB.CommandButton Command4 
         Caption         =   "Make shiny"
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Change nature"
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add TP"
         Height          =   315
         Index           =   1
         Left            =   1200
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton rmv 
         Caption         =   "Remove"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   975
      End
      Begin VB.ListBox lstPokemon 
         Height          =   1035
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblnature 
         Caption         =   "Nature num 1"
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblTP 
         Caption         =   "TP 0"
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Player name"
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "lbl"
      Height          =   615
      Left            =   4680
      TabIndex        =   15
      Top             =   5280
      Width           =   2175
   End
End
Attribute VB_Name = "frmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim playerLoad As PlayerRec
Dim loadedPlayer As String
Private Sub cmdLoad_Click()
EditorLoadPlayer (Text1.Text & ".bin")
End Sub

Sub EditorLoadPlayer(ByVal Name As String)
On Error Resume Next
    Dim FileName As String
    Dim F As Long
    Call EditorClearPlayer
    FileName = App.path & "\data\accounts\" & Trim(Name)
    F = FreeFile
    Open FileName For Binary As #F
    Get #F, , playerLoad
    Close #F
    LoadThings
    loadedPlayer = Name
End Sub

Sub EditorClearPlayer()
On Error Resume Next
    Dim i As Long
    
    Call ZeroMemory(ByVal VarPtr(playerLoad), LenB(playerLoad))
   playerLoad.Login = vbNullString
    playerLoad.Password = vbNullString
  playerLoad.Name = vbNullString
  playerLoad.Class = 1
End Sub

Sub LoadThings()
Dim xyz As Long
lstPokemon.Clear
For xyz = 1 To 6
If playerLoad.PokemonInstance(xyz).PokemonNumber > 0 Then
lstPokemon.AddItem (Trim$(Pokemon(playerLoad.PokemonInstance(xyz).PokemonNumber).Name) & " lvl." & playerLoad.PokemonInstance(xyz).level)
Else
lstPokemon.AddItem ("Empty")
End If

Next

storagee.Clear
For xyz = 1 To 250
If playerLoad.StoragePokemonInstance(xyz).PokemonNumber > 0 Then
storagee.AddItem (Trim$(Pokemon(playerLoad.StoragePokemonInstance(xyz).PokemonNumber).Name) & " lvl." & playerLoad.StoragePokemonInstance(xyz).level)
Else
storagee.AddItem ("Empty")
End If
Next

End Sub


Private Sub ListFilesa(strPath As String, Optional Extention As String)
Dim file As String
If Right$(strPath, 1) <> "\" Then strPath = strPath & "\"
If Trim$(Extention) = "" Then
Extention = "*.*"
ElseIf Left$(Extention, 2) <> "*." Then
Extention = "*." & Extention
End If
file = Dir$(strPath)
Do While Len(file)
lstPlayers.AddItem file
file = Dir$
Loop
End Sub





Private Sub Command1_Click(Index As Integer)

End Sub

Private Sub Command5_Click()
On Error Resume Next
    Dim FileName As String
    Dim F As Long

    FileName = App.path & "\data\accounts\" & loadedPlayer
    
    F = FreeFile
    
    Open FileName For Binary As #F
    Put #F, , playerLoad
    Close #F
End Sub

Private Sub Form_Load()
Call ListFilesa(App.path & "\Data\accounts\", ".bin")
End Sub

Private Sub lstPlayers_Click()
EditorLoadPlayer (lstPlayers.Text)
End Sub


Private Sub rmv_Click()
Dim newPoke1 As PokemonInstanceRec
playerLoad.PokemonInstance(lstPokemon.ListIndex + 1) = newPoke1
LoadThings
End Sub

Private Sub rmvst_Click()
Dim newPoke2 As PokemonInstanceRec
playerLoad.StoragePokemonInstance(storagee.ListIndex + 1) = newPoke2
LoadThings
End Sub
