VERSION 5.00
Begin VB.Form frmEditor_Resource 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resource Editor"
   ClientHeight    =   375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1440
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
   Icon            =   "frmEditor_Resource.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   96
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5160
      TabIndex        =   28
      Top             =   7320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Nothing"
      Height          =   375
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3480
      TabIndex        =   26
      Top             =   7320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resource Properties"
      Height          =   7095
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   5055
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   120
         Max             =   6000
         TabIndex        =   29
         Top             =   6720
         Width           =   4815
      End
      Begin VB.HScrollBar scrlExhaustedPic 
         Height          =   255
         Left            =   2640
         TabIndex        =   24
         Top             =   1920
         Width           =   2295
      End
      Begin VB.PictureBox picExhaustedPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
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
         Height          =   1680
         Left            =   2640
         ScaleHeight     =   112
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   152
         TabIndex        =   23
         Top             =   2280
         Width           =   2280
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   3975
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Resource.frx":020A
         Left            =   960
         List            =   "frmEditor_Resource.frx":021A
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1320
         Width           =   3975
      End
      Begin VB.HScrollBar scrlNormalPic 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   2295
      End
      Begin VB.HScrollBar scrlReward 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   4320
         Width           =   4815
      End
      Begin VB.HScrollBar scrlTool 
         Height          =   255
         Left            =   120
         Max             =   3
         TabIndex        =   9
         Top             =   4920
         Width           =   4815
      End
      Begin VB.HScrollBar scrlHealth 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   8
         Top             =   5520
         Width           =   4815
      End
      Begin VB.PictureBox picNormalPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
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
         Height          =   1680
         Left            =   120
         ScaleHeight     =   112
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   152
         TabIndex        =   7
         Top             =   2280
         Width           =   2280
      End
      Begin VB.HScrollBar scrlRespawn 
         Height          =   255
         Left            =   120
         Max             =   6000
         TabIndex        =   6
         Top             =   6120
         Width           =   4815
      End
      Begin VB.TextBox txtMessage 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox txtMessage2 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         Caption         =   "Animation: None"
         Height          =   180
         Left            =   120
         TabIndex        =   30
         Top             =   6480
         Width           =   1260
      End
      Begin VB.Label lblExhaustedPic 
         AutoSize        =   -1  'True
         Caption         =   "Exhausted Image: 0"
         Height          =   180
         Left            =   2640
         TabIndex        =   25
         Top             =   1680
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   180
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   435
      End
      Begin VB.Label lblNormalPic 
         AutoSize        =   -1  'True
         Caption         =   "Normal Image: 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   1470
      End
      Begin VB.Label lblReward 
         AutoSize        =   -1  'True
         Caption         =   "Item Reward: None"
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   4080
         Width           =   1440
      End
      Begin VB.Label lblTool 
         AutoSize        =   -1  'True
         Caption         =   "Tool Required: None"
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   4680
         Width           =   1530
      End
      Begin VB.Label lblHealth 
         AutoSize        =   -1  'True
         Caption         =   "Health: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   5280
         Width           =   705
      End
      Begin VB.Label lblRespawn 
         AutoSize        =   -1  'True
         Caption         =   "Respawn Time (Seconds): 0"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   5880
         Width           =   2100
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Success:"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empty:"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   7320
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Frame Frame3 
      Caption         =   "Resource List"
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   6540
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Resource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbType_Click()
    Resource(EditorIndex).ResourceType = cmbType.ListIndex
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

ClearResource EditorIndex

tmpIndex = lstIndex.ListIndex
lstIndex.RemoveItem EditorIndex - 1
lstIndex.AddItem EditorIndex & ": " & Resource(EditorIndex).Name, EditorIndex - 1
lstIndex.ListIndex = tmpIndex

ResourceEditorInit
End Sub

Private Sub cmdSave_Click()
    Call ResourceEditorOk
End Sub

Private Sub Form_Load()
    scrlReward.Max = MAX_ITEMS
End Sub

Private Sub cmdCancel_Click()
    Call ResourceEditorCancel
End Sub

Private Sub Label5_Click()

End Sub

Private Sub lstIndex_Click()
    ResourceEditorInit
End Sub

Private Sub scrlAnimation_Change()
Dim sString As String
    If scrlAnimation.Value = 0 Then sString = "None" Else sString = Trim$(Animation(scrlAnimation.Value).Name)
    lblAnim.Caption = "Animation: " & sString
    Resource(EditorIndex).Animation = scrlAnimation.Value
End Sub

Private Sub scrlExhaustedPic_Change()
    lblExhaustedPic.Caption = "Exhausted Image: " & scrlExhaustedPic.Value
    EditorResource_BltSprite
    Resource(EditorIndex).ExhaustedImage = scrlExhaustedPic.Value
End Sub

Private Sub scrlHealth_Change()
    lblHealth.Caption = "Health: " & scrlHealth.Value
    Resource(EditorIndex).Health = scrlHealth.Value
End Sub

Private Sub scrlNormalPic_Change()
    lblNormalPic.Caption = "Normal Image: " & scrlNormalPic.Value
    EditorResource_BltSprite
    Resource(EditorIndex).ResourceImage = scrlNormalPic.Value
End Sub

Private Sub scrlRespawn_Change()
    lblRespawn.Caption = "Respawn Time (Seconds): " & scrlRespawn.Value
    Resource(EditorIndex).RespawnTime = scrlRespawn.Value
End Sub

Private Sub scrlReward_Change()

    If scrlReward.Value > 0 Then
        lblReward.Caption = "Item Reward: " & Trim$(Item(scrlReward.Value).Name)
    Else
        lblReward.Caption = "Item Reward: None"
    End If
    
    Resource(EditorIndex).ItemReward = scrlReward.Value

End Sub

Private Sub scrlTool_Change()
    Dim Name As String

    Select Case scrlTool.Value
        Case 0
            Name = "None"
        Case 1
            Name = "Hatchet"
        Case 2
            Name = "Rod"
        Case 3
            Name = "Pickaxe"
    End Select

    lblTool.Caption = "Tool Required: " & Name
    
    Resource(EditorIndex).ToolRequired = scrlTool.Value
End Sub

Private Sub txtMessage_Change()
    Resource(EditorIndex).SuccessMessage = Trim$(txtMessage.text)
End Sub

Private Sub txtMessage2_Change()
    Resource(EditorIndex).EmptyMessage = Trim$(txtMessage2.text)
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Resource(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Resource(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
End Sub
