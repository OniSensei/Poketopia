VERSION 5.00
Begin VB.Form frmSimulator 
   Caption         =   "Spawn Simulator"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3120
      TabIndex        =   4
      Text            =   "100"
      Top             =   600
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Simulate"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Times"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Results:"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmSimulator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function SpawnChance(ByVal OneOf As Long) As Boolean
On Error Resume Next
'n = Int(Rnd * OneOf) + 1
Dim x As Long
Dim y As Long
x = Rand(1, OneOf)
y = Rand(1, OneOf)
'If n = 1 then
If x = y Then
SpawnChance = True
Else
SpawnChance = False
End If
End Function
Private Sub Command1_Click()
Dim x(1 To MAX_MAP_POKEMONS) As Integer
Dim i As Integer
For i = 1 To MAX_MAP_POKEMONS
x(i) = 0
Next
For a = 1 To Val(Text1.text)
For i = 1 To MAX_MAP_POKEMONS
If SpawnChance(map.Pokemon(i).Chance) = True Then
x(i) = x(i) + 1
Exit For
End If
Next
Next
List1.Clear

For i = 1 To MAX_MAP_POKEMONS
If map.Pokemon(i).PokemonNumber > 0 Then
List1.AddItem (Trim$(Pokemon(map.Pokemon(i).PokemonNumber).Name) & " - " & x(i))
End If
Next
End Sub
