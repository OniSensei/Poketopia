VERSION 5.00
Begin VB.Form frmEditor_Pokemon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pokemon Editor"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12600
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_Pokemon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   12600
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtStone 
      Height          =   285
      Left            =   8400
      TabIndex        =   47
      Text            =   "None"
      Top             =   3240
      Width           =   4095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Moves"
      Height          =   2415
      Left            =   8280
      TabIndex        =   41
      Top             =   240
      Width           =   4095
      Begin VB.HScrollBar scrlMove 
         Height          =   255
         Left            =   240
         Max             =   500
         TabIndex        =   45
         Top             =   1680
         Width           =   3615
      End
      Begin VB.HScrollBar scrlLevel 
         Height          =   255
         Left            =   240
         Max             =   100
         Min             =   1
         TabIndex        =   43
         Top             =   840
         Value           =   1
         Width           =   3615
      End
      Begin VB.Label lblMove 
         Caption         =   "Move: None."
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Label lblMoveLevel 
         Caption         =   "Level: 1"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   480
         Width           =   3495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pokemon Properties"
      Height          =   7215
      Left            =   3360
      TabIndex        =   6
      Top             =   120
      Width           =   4695
      Begin VB.HScrollBar scrlExp 
         Height          =   255
         Left            =   1080
         Max             =   999
         TabIndex        =   39
         Top             =   6840
         Width           =   3255
      End
      Begin VB.HScrollBar scrlCatchRate 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   37
         Top             =   6480
         Width           =   4215
      End
      Begin VB.HScrollBar scrlHappiness 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   35
         Top             =   5880
         Width           =   4215
      End
      Begin VB.HScrollBar scrlFemalePerc 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   33
         Top             =   5280
         Width           =   4215
      End
      Begin VB.HScrollBar scrlRareness 
         Height          =   255
         Left            =   2280
         Max             =   255
         TabIndex        =   31
         Top             =   4680
         Width           =   2055
      End
      Begin VB.HScrollBar scrlSpDef 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   30
         Top             =   4680
         Width           =   2055
      End
      Begin VB.HScrollBar scrlSpAtk 
         Height          =   255
         Left            =   2280
         Max             =   255
         TabIndex        =   28
         Top             =   4080
         Width           =   2055
      End
      Begin VB.HScrollBar scrlSpd 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   26
         Top             =   4080
         Width           =   2055
      End
      Begin VB.HScrollBar scrlDef 
         Height          =   255
         Left            =   2280
         Max             =   255
         TabIndex        =   24
         Top             =   3480
         Width           =   2055
      End
      Begin VB.HScrollBar scrlAtk 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   22
         Top             =   3480
         Width           =   2055
      End
      Begin VB.HScrollBar scrlEvolveLvl 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   19
         Top             =   2880
         Width           =   4215
      End
      Begin VB.HScrollBar scrlEvolvePoke 
         Height          =   255
         Left            =   120
         Max             =   649
         TabIndex        =   17
         Top             =   2280
         Width           =   4215
      End
      Begin VB.ComboBox cmbType2 
         Height          =   315
         ItemData        =   "frmEditor_Pokemon.frx":020A
         Left            =   1080
         List            =   "frmEditor_Pokemon.frx":0247
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1680
         Width           =   3255
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         ItemData        =   "frmEditor_Pokemon.frx":02D2
         Left            =   1080
         List            =   "frmEditor_Pokemon.frx":030F
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1320
         Width           =   3255
      End
      Begin VB.HScrollBar scrlPP 
         Height          =   255
         Left            =   1080
         Max             =   999
         TabIndex        =   12
         Top             =   960
         Width           =   3255
      End
      Begin VB.HScrollBar scrlHP 
         Height          =   255
         Left            =   1080
         Max             =   999
         TabIndex        =   10
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label lblExp 
         Caption         =   "EXP: 9999"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   6840
         Width           =   1455
      End
      Begin VB.Label lblCatchRate 
         Caption         =   "Catch Rate: 255"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   6240
         Width           =   3975
      End
      Begin VB.Label lblHappiness 
         Caption         =   "Happiness: 255"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   5640
         Width           =   3975
      End
      Begin VB.Label lblFemalePerc 
         Caption         =   "Percent Female: 100%"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   5040
         Width           =   3975
      End
      Begin VB.Label lblRareness 
         Caption         =   "Rareness: 0"
         Height          =   255
         Left            =   2280
         TabIndex        =   32
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label lblSpDef 
         Caption         =   "SpDef: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label lblSpAtk 
         Caption         =   "SpAtk: 0"
         Height          =   255
         Left            =   2280
         TabIndex        =   27
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label lblSpd 
         Caption         =   "Spd: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label lblDef 
         Caption         =   "Def: 0"
         Height          =   255
         Left            =   2280
         TabIndex        =   23
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label lblAtk 
         Caption         =   "Atk: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label lblEvolveLvl 
         Caption         =   "Evolves at level: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2640
         Width           =   3975
      End
      Begin VB.Label lblEvolvePoke 
         Caption         =   "Evolves into: None"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   3975
      End
      Begin VB.Label Label2 
         Caption         =   "Type2:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblPP 
         Caption         =   "PP: 999"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblHP 
         Caption         =   "HP: 999"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   570
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Height          =   375
      Left            =   8160
      TabIndex        =   5
      Top             =   6240
      Width           =   2895
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   9600
      TabIndex        =   4
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   11040
      TabIndex        =   3
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   8160
      TabIndex        =   2
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Pokemon List"
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   6690
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Evolution stone:"
      Height          =   255
      Left            =   8400
      TabIndex        =   46
      Top             =   2880
      Width           =   3015
   End
   Begin VB.Image Image2 
      Height          =   1695
      Left            =   10320
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   8160
      Top             =   4440
      Width           =   2055
   End
End
Attribute VB_Name = "frmEditor_Pokemon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbType2_Click()
    Pokemon(EditorIndex).Type2 = cmbType2.ListIndex
End Sub

Private Sub cmdSave_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
    Call PokemonEditorOk
End If
End Sub

Private Sub cmdCancel_Click()
    Call PokemonEditorCancel
End Sub

Private Sub cmbType_Click()
    Pokemon(EditorIndex).Type = cmbType.ListIndex
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

ClearPokemon EditorIndex

tmpIndex = lstIndex.ListIndex
lstIndex.RemoveItem EditorIndex - 1
lstIndex.AddItem EditorIndex & ": " & Pokemon(EditorIndex).Name, EditorIndex - 1
lstIndex.ListIndex = tmpIndex

PokemonEditorInit
End Sub

Private Sub lstIndex_Click()
    PokemonEditorInit
End Sub

Private Sub scrlAtk_Change()
    lblAtk.Caption = "Atk: " & scrlAtk.Value
    Pokemon(EditorIndex).ATK = scrlAtk.Value
End Sub

Private Sub scrlDef_Change()
    lblDef.Caption = "Def: " & scrlDef.Value
    Pokemon(EditorIndex).DEF = scrlDef.Value
End Sub

Private Sub scrlExp_Change()
    lblEXP.Caption = "EXP: " & scrlExp.Value
    Pokemon(EditorIndex).BaseEXP = scrlExp.Value
End Sub

Private Sub scrlHappiness_Change()
    lblHappiness.Caption = "Happiness: " & scrlHappiness.Value
    Pokemon(EditorIndex).Happiness = scrlHappiness.Value
End Sub

Private Sub scrlCatchRate_Change()
    lblCatchRate.Caption = "Catch Rate: " & scrlCatchRate.Value
    Pokemon(EditorIndex).CatchRate = scrlCatchRate.Value
End Sub







Private Sub scrlLevel_Change()
lblMoveLevel.Caption = "Level: " & scrlLevel.Value
scrlMove.Value = Pokemon(EditorIndex).moves(scrlLevel.Value)
End Sub

Private Sub scrlMove_Change()
On Error Resume Next
If scrlMove.Value <= 0 Then
lblMove.Caption = "Move: None."
Else
lblMove.Caption = "Move: " & scrlMove.Value & " - " & PokemonMove(scrlMove.Value).Name
End If

Pokemon(EditorIndex).moves(scrlLevel.Value) = scrlMove.Value
End Sub

Private Sub scrlRareness_Change()
    lblRareness.Caption = "Rareness: " & scrlRareness.Value
    Pokemon(EditorIndex).Rareness = scrlRareness.Value
End Sub

Private Sub scrlFemalePerc_Change()
    lblFemalePerc.Caption = "Percent Female: " & scrlFemalePerc.Value & "%"
    Pokemon(EditorIndex).PercentFemale = scrlFemalePerc.Value
End Sub

Private Sub scrlSpd_Change()
    lblSpd.Caption = "Spd: " & scrlSpd.Value
    Pokemon(EditorIndex).SPD = scrlSpd.Value
End Sub

Private Sub scrlSpAtk_Change()
    lblSpAtk.Caption = "SpAtk: " & scrlSpAtk.Value
    Pokemon(EditorIndex).SPATK = scrlSpAtk.Value
End Sub

Private Sub scrlSpDef_Change()
    lblSpDef.Caption = "SpDef: " & scrlSpDef.Value
    Pokemon(EditorIndex).SPDEF = scrlSpDef.Value
End Sub

Private Sub scrlEvolveLvl_Change()
    lblEvolveLvl.Caption = "Evolves at level: " & scrlEvolveLvl.Value
    Pokemon(EditorIndex).Evolution = scrlEvolveLvl.Value
End Sub

Private Sub scrlEvolvePoke_Change()
    If scrlEvolvePoke.Value = 0 Then
        lblEvolvePoke.Caption = "Evolves into: None"
    Else
        lblEvolvePoke.Caption = "Evolves into: " & Trim$(Pokemon(scrlEvolvePoke.Value).Name)
    End If
    Pokemon(EditorIndex).EvolvesTo = scrlEvolvePoke.Value
End Sub

Private Sub scrlHP_Change()
    lblHp.Caption = "HP: " & scrlHP.Value
    Pokemon(EditorIndex).MaxHp = scrlHP.Value
End Sub

Private Sub scrlPP_Change()
    lblPP.Caption = "PP: " & scrlPP.Value
    Pokemon(EditorIndex).MaxPP = scrlPP.Value
End Sub

Private Sub txtName_Change()
Dim tmpIndex As Long
    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Pokemon(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Pokemon(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
End Sub

Private Sub txtStone_Change()
Pokemon(EditorIndex).Stone = Trim$(txtStone.text)
End Sub
