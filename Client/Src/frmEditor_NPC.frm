VERSION 5.00
Begin VB.Form frmEditor_NPC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Npc Editor"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12090
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_NPC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   504
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   806
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtP3 
      Height          =   270
      Left            =   8520
      TabIndex        =   55
      Text            =   "0"
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtP2 
      Height          =   270
      Left            =   8520
      TabIndex        =   53
      Text            =   "0"
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtP1 
      Height          =   270
      Left            =   8520
      TabIndex        =   51
      Text            =   "0"
      Top             =   600
      Width           =   2175
   End
   Begin VB.CheckBox chkCanMove 
      Caption         =   "Can NPC Move?"
      Height          =   255
      Left            =   8520
      TabIndex        =   49
      Top             =   2640
      Width           =   4455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   8640
      TabIndex        =   46
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   10200
      TabIndex        =   45
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   8640
      TabIndex        =   44
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "NPC Properties"
      Height          =   6975
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   5055
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   2640
         TabIndex        =   48
         Top             =   2880
         Width           =   2055
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
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
         Height          =   480
         Left            =   4440
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   35
         Top             =   960
         Width           =   480
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   34
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txtName 
         Height          =   270
         Left            =   960
         TabIndex        =   33
         Top             =   240
         Width           =   3975
      End
      Begin VB.ComboBox cmbBehaviour 
         Height          =   300
         ItemData        =   "frmEditor_NPC.frx":020A
         Left            =   1320
         List            =   "frmEditor_NPC.frx":021D
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1680
         Width           =   3615
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   31
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox txtAttackSay 
         Height          =   255
         Left            =   960
         TabIndex        =   30
         Top             =   600
         Width           =   3975
      End
      Begin VB.Frame fraDrop 
         Caption         =   "Drop"
         Height          =   2175
         Left            =   120
         TabIndex        =   20
         Top             =   4680
         Width           =   4815
         Begin VB.TextBox txtSpawnSecs 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2880
            TabIndex        =   24
            Text            =   "0"
            Top             =   600
            Width           =   1815
         End
         Begin VB.HScrollBar scrlValue 
            Height          =   255
            Left            =   1200
            Max             =   255
            TabIndex        =   23
            Top             =   1680
            Width           =   3495
         End
         Begin VB.HScrollBar scrlNum 
            Height          =   255
            Left            =   1200
            Max             =   255
            TabIndex        =   22
            Top             =   1320
            Width           =   3495
         End
         Begin VB.TextBox txtChance 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2880
            TabIndex        =   21
            Text            =   "0"
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Spawn Rate (in seconds)"
            Height          =   180
            Left            =   120
            TabIndex        =   29
            Top             =   600
            UseMnemonic     =   0   'False
            Width           =   1845
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            Caption         =   "Value: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   28
            Top             =   1680
            UseMnemonic     =   0   'False
            Width           =   645
         End
         Begin VB.Label lblItemName 
            AutoSize        =   -1  'True
            Caption         =   "Item: None"
            Height          =   180
            Left            =   120
            TabIndex        =   27
            Top             =   960
            Width           =   855
         End
         Begin VB.Label lblNum 
            AutoSize        =   -1  'True
            Caption         =   "Num: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   26
            Top             =   1320
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Chance 1 out of"
            Height          =   180
            Left            =   120
            TabIndex        =   25
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   1185
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Stats"
         Height          =   1455
         Left            =   120
         TabIndex        =   7
         Top             =   3120
         Width           =   4815
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   1
            Left            =   120
            Max             =   255
            TabIndex        =   13
            Top             =   240
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   2
            Left            =   1680
            Max             =   255
            TabIndex        =   12
            Top             =   240
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   3
            Left            =   3240
            Max             =   255
            TabIndex        =   11
            Top             =   240
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   4
            Left            =   120
            Max             =   255
            TabIndex        =   10
            Top             =   840
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   5
            Left            =   1680
            Max             =   255
            TabIndex        =   9
            Top             =   840
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   6
            Left            =   3240
            Max             =   255
            TabIndex        =   8
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Str: 0"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   19
            Top             =   480
            Width           =   435
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "End: 0"
            Height          =   180
            Index           =   2
            Left            =   1680
            TabIndex        =   18
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Vit: 0"
            Height          =   180
            Index           =   3
            Left            =   3240
            TabIndex        =   17
            Top             =   480
            Width           =   435
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Will: 0"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   16
            Top             =   1080
            Width           =   480
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Int: 0"
            Height          =   180
            Index           =   5
            Left            =   1680
            TabIndex        =   15
            Top             =   1080
            Width           =   435
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Spr: 0"
            Height          =   180
            Index           =   6
            Left            =   3240
            TabIndex        =   14
            Top             =   1080
            Width           =   465
         End
      End
      Begin VB.ComboBox cmbFaction 
         Height          =   300
         ItemData        =   "frmEditor_NPC.frx":0266
         Left            =   1320
         List            =   "frmEditor_NPC.frx":0273
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2040
         Width           =   3615
      End
      Begin VB.TextBox txtHP 
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox txtEXP 
         Height          =   285
         Left            =   3120
         TabIndex        =   4
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label lblAnimation 
         Caption         =   "Anim: None"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   43
         Top             =   960
         Width           =   660
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   42
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Behaviour:"
         Height          =   180
         Left            =   120
         TabIndex        =   41
         Top             =   1680
         UseMnemonic     =   0   'False
         Width           =   810
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         Caption         =   "Range: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   40
         Top             =   1320
         UseMnemonic     =   0   'False
         Width           =   675
      End
      Begin VB.Label lblSay 
         AutoSize        =   -1  'True
         Caption         =   "Say:"
         Height          =   180
         Left            =   120
         TabIndex        =   39
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   345
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Exp:"
         Height          =   180
         Left            =   2640
         TabIndex        =   38
         Top             =   2520
         Width           =   345
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Health:"
         Height          =   180
         Left            =   120
         TabIndex        =   37
         Top             =   2520
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Faction:"
         Height          =   180
         Left            =   120
         TabIndex        =   36
         Top             =   2040
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "NPC List"
      Height          =   6975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   6540
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Height          =   375
      Left            =   8640
      TabIndex        =   0
      Top             =   6720
      Width           =   2895
   End
   Begin VB.Label Label6 
      Caption         =   "Outfit"
      Height          =   255
      Left            =   8520
      TabIndex        =   54
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Jacket"
      Height          =   255
      Left            =   8520
      TabIndex        =   52
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Head"
      Height          =   255
      Left            =   8520
      TabIndex        =   50
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmEditor_NPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkCanMove_Click()
NPC(EditorIndex).CanMove = chkCanMove.Value
End Sub

Private Sub cmbBehaviour_Click()
    NPC(EditorIndex).Behaviour = cmbBehaviour.ListIndex
End Sub

Private Sub cmbFaction_Click()
    NPC(EditorIndex).faction = cmbFaction.ListIndex
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

ClearNPC EditorIndex

tmpIndex = lstIndex.ListIndex
lstIndex.RemoveItem EditorIndex - 1
lstIndex.AddItem EditorIndex & ": " & NPC(EditorIndex).Name, EditorIndex - 1
lstIndex.ListIndex = tmpIndex

NpcEditorInit
End Sub

Private Sub Form_Load()
    scrlSprite.Max = NumCharacters
    scrlAnimation.Max = MAX_ANIMATIONS
End Sub

Private Sub cmdSave_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
    Call NpcEditorOk
End If
End Sub

Private Sub cmdCancel_Click()
    Call NpcEditorCancel
End Sub

Private Sub Label7_Click()

End Sub

Private Sub lstIndex_Click()
    NpcEditorInit
End Sub

Private Sub scrlAnimation_Change()
Dim sString As String
    If scrlAnimation.Value = 0 Then sString = "None" Else sString = Trim$(Animation(scrlAnimation.Value).Name)
    lblAnimation.Caption = "Anim: " & sString
    NPC(EditorIndex).Animation = scrlAnimation.Value
End Sub

Private Sub scrlSprite_Change()
    lblSprite.Caption = "Sprite: " & scrlSprite.Value
    Call EditorNpc_BltSprite
    NPC(EditorIndex).Sprite = scrlSprite.Value
End Sub

Private Sub scrlRange_Change()
    lblRange.Caption = "Range: " & scrlRange.Value
    NPC(EditorIndex).Range = scrlRange.Value
End Sub

Private Sub scrlNum_Change()
    lblNum.Caption = "Num: " & scrlNum.Value

    If scrlNum.Value > 0 Then
        lblItemName.Caption = "Item: " & Trim$(Item(scrlNum.Value).Name)
    End If
    
    NPC(EditorIndex).DropItem = scrlNum.Value
End Sub

Private Sub scrlStat_Change(Index As Integer)
    Dim prefix As String
    Select Case Index
        Case 1 ' str
            prefix = "Str: "
        Case 2 ' end
            prefix = "End: "
        Case 3 ' vit
            prefix = "Vit: "
        Case 4 ' will
            prefix = "Will: "
        Case 5 ' int
            prefix = "Int: "
        Case 6 ' spr
            prefix = "Spr: "
    End Select
    lblStat(Index).Caption = prefix & scrlStat(Index).Value
    NPC(EditorIndex).Stat(Index) = scrlStat(Index).Value
End Sub

Private Sub scrlValue_Change()
    lblValue.Caption = "Value: " & scrlValue.Value
    NPC(EditorIndex).DropItemValue = scrlValue.Value
End Sub

Private Sub txtAttackSay_Change()
    NPC(EditorIndex).AttackSay = txtAttackSay.text
End Sub

Private Sub txtChance_Change()
    NPC(EditorIndex).DropChance = txtChance.text
End Sub

Private Sub txtEXP_Change()
    NPC(EditorIndex).EXP = txtEXP.text
End Sub

Private Sub txtHP_Change()
    If IsNumeric(txtHP.text) Then NPC(EditorIndex).HP = txtHP.text
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long
    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    NPC(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & NPC(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
End Sub

Private Sub txtP1_Change()
NPC(EditorIndex).Paperdoll1 = Val(txtP1.text)
End Sub

Private Sub txtP2_Change()
NPC(EditorIndex).Paperdoll2 = Val(txtP2.text)
End Sub

Private Sub txtP3_Change()
NPC(EditorIndex).Paperdoll3 = Val(txtP3.text)
End Sub

Private Sub txtSpawnSecs_Change()
    NPC(EditorIndex).SpawnSecs = txtSpawnSecs.text
End Sub
