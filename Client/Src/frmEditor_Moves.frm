VERSION 5.00
Begin VB.Form frmEditor_Moves 
   Caption         =   "Moves"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmbCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Information"
      Height          =   7575
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   5055
      Begin VB.TextBox txtEffectdescription 
         Height          =   1215
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   20
         Text            =   "frmEditor_Moves.frx":0000
         Top             =   4560
         Width           =   4455
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   240
         Max             =   256
         TabIndex        =   19
         Top             =   4200
         Width           =   3255
      End
      Begin VB.ComboBox cmbCategory 
         Height          =   315
         ItemData        =   "frmEditor_Moves.frx":000C
         Left            =   1080
         List            =   "frmEditor_Moves.frx":0019
         TabIndex        =   18
         Text            =   "Combo1"
         Top             =   5880
         Width           =   3015
      End
      Begin VB.TextBox txtDescription 
         Height          =   765
         Left            =   1440
         TabIndex        =   16
         Text            =   "Description"
         Top             =   6600
         Width           =   2655
      End
      Begin VB.HScrollBar scrlAccuracy 
         Height          =   255
         Left            =   240
         Max             =   100
         TabIndex        =   11
         Top             =   3480
         Width           =   2775
      End
      Begin VB.HScrollBar scrlPower 
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2640
         Width           =   2775
      End
      Begin VB.HScrollBar scrlPP 
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   2775
      End
      Begin VB.ComboBox cmbType1 
         Height          =   315
         ItemData        =   "frmEditor_Moves.frx":0038
         Left            =   840
         List            =   "frmEditor_Moves.frx":0075
         TabIndex        =   5
         Text            =   "None"
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txtMoveName 
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Text            =   "Move"
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "Category:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   5880
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Description:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   6600
         Width           =   975
      End
      Begin VB.Label lblEffect 
         Caption         =   "Effect:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label lblAccuracy 
         Caption         =   "Accuracy:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3120
         Width           =   3735
      End
      Begin VB.Label lblPower 
         Caption         =   "Power:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2280
         Width           =   3735
      End
      Begin VB.Label lblPP 
         Caption         =   "PP:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "Type:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.ListBox lstIndex 
      Height          =   5520
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "frmEditor_Moves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmbCancel_Click()
Call MovesEditorCancel
End Sub

Private Sub cmbCategory_Click()
Select Case cmbCategory.ListIndex
Case 0
'Physical Damage
PokemonMove(EditorIndex).Category = "Physical Damage"
Case 1
'Special Damage
PokemonMove(EditorIndex).Category = "Special Damage"
Case 2
'Status
PokemonMove(EditorIndex).Category = "Status"
End Select

End Sub



Private Sub cmbEffect_Click()
PokemonMove(EditorIndex).Effect = cmbEffect.ListIndex
End Sub

Private Sub cmbType1_Click()
PokemonMove(EditorIndex).Type = cmbType1.ListIndex
End Sub

Private Sub cmdSave_Click()
Call MovesEditorOk
End Sub



Private Sub HScroll1_Change()
lblEffect.Caption = "Effect: " & HScroll1.Value
PokemonMove(EditorIndex).Effect = HScroll1.Value
Call ReadText(App.Path & "\Data Files\database\Moves\" & HScroll1.Value & ".txt", txtEffectdescription)
End Sub

Private Sub lstIndex_Click()
MovesEditorInit
End Sub

Private Sub scrlAccuracy_Change()
lblAccuracy.Caption = "Accuracy: " & scrlAccuracy.Value
PokemonMove(EditorIndex).Accuracy = scrlAccuracy.Value
End Sub


Private Sub scrlPower_Change()
lblPower.Caption = "Power: " & scrlPower.Value
PokemonMove(EditorIndex).Power = scrlPower.Value
End Sub

Private Sub scrlPP_Change()
lblPP.Caption = "PP: " & scrlPP.Value
PokemonMove(EditorIndex).PP = scrlPP.Value
End Sub

Private Sub txtDescription_Change()
PokemonMove(EditorIndex).Description = txtDescription.text
End Sub

Private Sub txtMoveName_Change()
PokemonMove(EditorIndex).Name = txtMoveName.text
End Sub
