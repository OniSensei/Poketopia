VERSION 5.00
Begin VB.Form frmEditor_Map 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Editor"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14670
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
   Icon            =   "frmEditor_Map.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   609
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   978
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command5 
      Caption         =   "Clear list"
      Height          =   255
      Left            =   1800
      TabIndex        =   97
      Top             =   8760
      Width           =   975
   End
   Begin VB.TextBox txtTilesetName 
      Height          =   270
      Left            =   2880
      TabIndex        =   96
      Text            =   "TilesetName"
      Top             =   8760
      Width           =   4335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add this tileset"
      Height          =   255
      Left            =   120
      TabIndex        =   95
      Top             =   8760
      Width           =   1455
   End
   Begin VB.ListBox TileList 
      Height          =   960
      Left            =   120
      TabIndex        =   94
      Top             =   7800
      Width           =   7095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Use squares"
      Height          =   375
      Left            =   2040
      TabIndex        =   87
      Top             =   6960
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.PictureBox picAttributes 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   7440
      ScaleHeight     =   7215
      ScaleWidth      =   7095
      TabIndex        =   29
      Top             =   240
      Visible         =   0   'False
      Width           =   7095
      Begin VB.Frame frmCS 
         Caption         =   "Cs"
         Height          =   975
         Left            =   1800
         TabIndex        =   84
         Top             =   960
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton Command2 
            Caption         =   "Accept"
            Height          =   255
            Left            =   600
            TabIndex        =   86
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox txtcs 
            Height          =   270
            Left            =   120
            TabIndex        =   85
            Text            =   "Text1"
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame fraNpcSpawn 
         Caption         =   "Npc Spawn"
         Height          =   2655
         Left            =   1800
         TabIndex        =   36
         Top             =   2040
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ListBox lstNpc 
            Height          =   780
            Left            =   240
            TabIndex        =   40
            Top             =   360
            Width           =   2895
         End
         Begin VB.HScrollBar scrlNpcDir 
            Height          =   255
            Left            =   240
            Max             =   3
            TabIndex        =   38
            Top             =   1560
            Width           =   2895
         End
         Begin VB.CommandButton cmdNpcSpawn 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   37
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblNpcDir 
            Caption         =   "Direction: Up"
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   1320
            Width           =   2535
         End
      End
      Begin VB.Frame fraResource 
         Caption         =   "Object"
         Height          =   1695
         Left            =   1800
         TabIndex        =   30
         Top             =   2520
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdResourceOk 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   33
            Top             =   1080
            Width           =   1455
         End
         Begin VB.HScrollBar scrlResource 
            Height          =   255
            Left            =   240
            Max             =   100
            Min             =   1
            TabIndex        =   32
            Top             =   600
            Value           =   1
            Width           =   2895
         End
         Begin VB.Label lblResource 
            Caption         =   "Object:"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame fraMapWarp 
         Caption         =   "Map Warp"
         Height          =   2775
         Left            =   1800
         TabIndex        =   59
         Top             =   1920
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdMapWarp 
            Caption         =   "Accept"
            Height          =   375
            Left            =   1080
            TabIndex        =   66
            Top             =   2160
            Width           =   1215
         End
         Begin VB.HScrollBar scrlMapWarpY 
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   1680
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMapWarpX 
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   1080
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMapWarp 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   61
            Top             =   480
            Value           =   1
            Width           =   3135
         End
         Begin VB.Label lblMapWarpY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   1440
            Width           =   3135
         End
         Begin VB.Label lblMapWarpX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   840
            Width           =   3135
         End
         Begin VB.Label lblMapWarp 
            Caption         =   "Map: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame fragymblock 
         Caption         =   "Gym:"
         Height          =   1695
         Left            =   1800
         TabIndex        =   77
         Top             =   1920
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ComboBox cmbDirection 
            Height          =   300
            ItemData        =   "frmEditor_Map.frx":020A
            Left            =   120
            List            =   "frmEditor_Map.frx":0217
            TabIndex        =   81
            Text            =   "Direction"
            Top             =   720
            Width           =   3135
         End
         Begin VB.HScrollBar scrlgymnum 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   79
            Top             =   240
            Value           =   1
            Width           =   3135
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Accept"
            Height          =   255
            Left            =   600
            TabIndex        =   78
            Top             =   1080
            Width           =   2175
         End
      End
      Begin VB.Frame fraShop 
         Caption         =   "Shop"
         Height          =   1335
         Left            =   1920
         TabIndex        =   67
         Top             =   2640
         Visible         =   0   'False
         Width           =   3135
         Begin VB.CommandButton cmdShop 
            Caption         =   "Accept"
            Height          =   375
            Left            =   960
            TabIndex        =   69
            Top             =   720
            Width           =   1215
         End
         Begin VB.ComboBox cmbShop 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame fraKeyOpen 
         Caption         =   "Key Open"
         Height          =   2055
         Left            =   1800
         TabIndex        =   53
         Top             =   2400
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdKeyOpen 
            Caption         =   "Accept"
            Height          =   375
            Left            =   1080
            TabIndex        =   58
            Top             =   1440
            Width           =   1215
         End
         Begin VB.HScrollBar scrlKeyY 
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   1080
            Width           =   3015
         End
         Begin VB.HScrollBar scrlKeyX 
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label lblKeyY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   840
            Width           =   3015
         End
         Begin VB.Label lblKeyX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame fraMapKey 
         Caption         =   "Map Key"
         Height          =   1815
         Left            =   1800
         TabIndex        =   47
         Top             =   2520
         Visible         =   0   'False
         Width           =   3375
         Begin VB.PictureBox picMapKey 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   2760
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   52
            Top             =   600
            Width           =   480
         End
         Begin VB.CommandButton cmdMapKey 
            Caption         =   "Accept"
            Height          =   375
            Left            =   1080
            TabIndex        =   51
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CheckBox chkMapKey 
            Caption         =   "Take key away upon use."
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   960
            Value           =   1  'Checked
            Width           =   2535
         End
         Begin VB.HScrollBar scrlMapKey 
            Height          =   255
            Left            =   120
            Max             =   5
            Min             =   1
            TabIndex        =   49
            Top             =   600
            Value           =   1
            Width           =   2535
         End
         Begin VB.Label lblMapKey 
            Caption         =   "Item: None"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame fraMapItem 
         Caption         =   "Map Item"
         Height          =   1815
         Left            =   1800
         TabIndex        =   41
         Top             =   2520
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdMapItem 
            Caption         =   "Accept"
            Height          =   375
            Left            =   1200
            TabIndex        =   46
            Top             =   1200
            Width           =   1215
         End
         Begin VB.HScrollBar scrlMapItemValue 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   45
            Top             =   840
            Value           =   1
            Width           =   2535
         End
         Begin VB.HScrollBar scrlMapItem 
            Height          =   255
            Left            =   120
            Max             =   10
            Min             =   1
            TabIndex        =   44
            Top             =   600
            Value           =   1
            Width           =   2535
         End
         Begin VB.PictureBox picMapItem 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   2760
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   43
            Top             =   600
            Width           =   480
         End
         Begin VB.Label lblMapItem 
            Caption         =   "Item: None x0"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   3135
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton cmdProperties 
      Caption         =   "Properties"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Type"
      Height          =   1095
      Left            =   5760
      TabIndex        =   25
      Top             =   5760
      Width           =   1455
      Begin VB.OptionButton optAttribs 
         Alignment       =   1  'Right Justify
         Caption         =   "Attributes"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton optLayers 
         Alignment       =   1  'Right Justify
         Caption         =   "Layers"
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.HScrollBar scrlPictureX 
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   5400
      Width           =   5295
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5280
      Left            =   120
      ScaleHeight     =   352
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   352
      TabIndex        =   14
      Top             =   120
      Width           =   5280
      Begin VB.PictureBox picBackSelect 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   0
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   15
         Top             =   0
         Width           =   960
         Begin VB.Shape shpLoc 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Height          =   480
            Left            =   0
            Top             =   0
            Width           =   480
         End
         Begin VB.Shape shpSelected 
            BorderColor     =   &H000000FF&
            BorderWidth     =   2
            Height          =   480
            Left            =   0
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.VScrollBar scrlPictureY 
      Height          =   5295
      Left            =   5400
      Max             =   255
      TabIndex        =   13
      Top             =   120
      Width           =   255
   End
   Begin VB.Frame fraTileSet 
      Caption         =   "Tileset: 0"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   6000
      Width           =   5535
      Begin VB.HScrollBar scrlTileSet 
         Height          =   255
         Left            =   120
         Max             =   10
         Min             =   1
         TabIndex        =   1
         Top             =   360
         Value           =   1
         Width           =   5295
      End
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Frame fraLayers 
      Caption         =   "Layers"
      Height          =   5535
      Left            =   5760
      TabIndex        =   16
      Top             =   120
      Width           =   1455
      Begin VB.CheckBox chkVisFringe2 
         Caption         =   "Fringe2"
         Height          =   180
         Left            =   240
         TabIndex        =   92
         Top             =   2880
         Width           =   975
      End
      Begin VB.CheckBox chkVisFringe 
         Caption         =   "Fringe"
         Height          =   180
         Left            =   240
         TabIndex        =   91
         Top             =   2640
         Width           =   975
      End
      Begin VB.CheckBox chkVisMask2 
         Caption         =   "Mask2"
         Height          =   180
         Left            =   240
         TabIndex        =   90
         Top             =   2400
         Width           =   975
      End
      Begin VB.CheckBox chkVisMask 
         Caption         =   "Mask"
         Height          =   180
         Left            =   240
         TabIndex        =   89
         Top             =   2160
         Width           =   975
      End
      Begin VB.CheckBox chkVisGround 
         Caption         =   "Ground"
         Height          =   180
         Left            =   240
         TabIndex        =   88
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdFill 
         Caption         =   "Fill"
         Height          =   390
         Left            =   120
         TabIndex        =   20
         Top             =   5040
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Fringe"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Mask"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Ground"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Fringe2"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Mask2"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Visible:"
         Height          =   255
         Left            =   120
         TabIndex        =   93
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.Frame fraAttribs 
      Caption         =   "Attributes"
      Height          =   5535
      Left            =   5760
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
      Begin VB.OptionButton optCS 
         Caption         =   "Custom Scr."
         Height          =   180
         Left            =   120
         TabIndex        =   83
         Top             =   4080
         Width           =   1215
      End
      Begin VB.OptionButton optgymblock 
         Caption         =   "Gym Block"
         Height          =   180
         Left            =   120
         TabIndex        =   80
         Top             =   3840
         Width           =   1215
      End
      Begin VB.OptionButton optBank 
         Caption         =   "Bank"
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   3600
         Width           =   1095
      End
      Begin VB.OptionButton optStorage 
         Caption         =   "Storage"
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   3360
         Width           =   1095
      End
      Begin VB.OptionButton optSpawn 
         Caption         =   "Spawn"
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   3120
         Width           =   975
      End
      Begin VB.OptionButton optHeal 
         Caption         =   "Heal"
         Height          =   270
         Left            =   120
         TabIndex        =   73
         Top             =   2880
         Width           =   1215
      End
      Begin VB.OptionButton optBattle 
         Caption         =   "Battle"
         Height          =   270
         Left            =   120
         TabIndex        =   72
         Top             =   2640
         Width           =   1215
      End
      Begin VB.OptionButton optShop 
         Caption         =   "Shop"
         Height          =   270
         Left            =   120
         TabIndex        =   70
         Top             =   2400
         Width           =   1215
      End
      Begin VB.OptionButton optNpcSpawn 
         Caption         =   "Npc Spawn"
         Height          =   270
         Left            =   120
         TabIndex        =   35
         Top             =   2160
         Width           =   1215
      End
      Begin VB.OptionButton optDoor 
         Caption         =   "Door"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton optResource 
         Caption         =   "Resource"
         Height          =   240
         Left            =   120
         TabIndex        =   28
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton optKeyOpen 
         Caption         =   "Key Open"
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optBlocked 
         Caption         =   "Blocked"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optWarp 
         Caption         =   "Warp"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear2 
         Caption         =   "Clear"
         Height          =   390
         Left            =   120
         TabIndex        =   6
         Top             =   5040
         Width           =   1215
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item"
         Height          =   270
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optNpcAvoid 
         Caption         =   "Npc Avoid"
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optKey 
         Caption         =   "Key"
         Height          =   270
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.Label lblPosition 
      Caption         =   "0,0"
      Height          =   255
      Left            =   120
      TabIndex        =   82
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Drag mouse to select multiple tiles"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   71
      Top             =   5760
      Width           =   5535
   End
End
Attribute VB_Name = "frmEditor_Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkVisFringe_Click()
If chkVisFringe.Value = 0 Then
FringeUnvisible = True
Else
FringeUnvisible = False
End If
End Sub

Private Sub chkVisFringe2_Click()
If chkVisFringe2.Value = 0 Then
Fringe2Unvisible = True
Else
Fringe2Unvisible = False
End If
End Sub

Private Sub chkVisGround_Click()
If chkVisGround.Value = 0 Then
GroundUnvisible = True
Else
GroundUnvisible = False
End If
End Sub

Private Sub chkVisMask_Click()
If chkVisMask.Value = 0 Then
MaskUnvisible = True
Else
MaskUnvisible = False
End If
End Sub

Private Sub chkVisMask2_Click()
If chkVisMask2.Value = 0 Then
Mask2Unvisible = True
Else
Mask2Unvisible = False
End If
End Sub

Private Sub cmdKeyOpen_Click()
    KeyOpenEditorX = scrlKeyX.Value
    KeyOpenEditorY = scrlKeyY.Value
    picAttributes.Visible = False
    fraKeyOpen.Visible = False
End Sub

Private Sub cmdMapItem_Click()
    ItemEditorNum = scrlMapItem.Value
    ItemEditorValue = scrlMapItemValue.Value
    picAttributes.Visible = False
    fraMapItem.Visible = False
End Sub

Private Sub cmdMapKey_Click()
    KeyEditorNum = scrlMapKey.Value
    KeyEditorTake = chkMapKey.Value
    picAttributes.Visible = False
    fraMapKey.Visible = False
End Sub

Private Sub cmdMapWarp_Click()
    EditorWarpMap = scrlMapWarp.Value
    EditorWarpX = scrlMapWarpX.Value
    EditorWarpY = scrlMapWarpY.Value
    picAttributes.Visible = False
    fraMapWarp.Visible = False
End Sub

Private Sub cmdNpcSpawn_Click()
    SpawnNpcNum = lstNpc.ListIndex + 1
    SpawnNpcDir = scrlNpcDir.Value
    picAttributes.Visible = False
    fraNpcSpawn.Visible = False
End Sub

Private Sub cmdResourceOk_Click()
    ResourceEditorNum = scrlResource.Value
    picAttributes.Visible = False
    fraResource.Visible = False
End Sub

Private Sub cmdShop_Click()
    EditorShop = cmbShop.ListIndex
    picAttributes.Visible = False
    fraShop.Visible = False
End Sub

Private Sub Command1_Click()
EditorGymBlockNum = scrlgymnum.Value
Select Case cmbDirection.ListIndex
Case 0
EditorGymBlockDir = GYMBLOCK_DIRECTION_UPDOWN
Case 1
EditorGymBlockDir = GYMBLOCK_DIRECTION_LEFT
Case 2
EditorGymBlockDir = GYMBLOCK_DIRECTION_RIGHT
End Select
    picAttributes.Visible = False
    fragymblock.Visible = False
End Sub

Private Sub Command2_Click()
EditorGymBlockNum = Val(txtcs.text)

    picAttributes.Visible = False
    frmCS.Visible = False
End Sub

Private Sub Command3_Click()
Dim i As Long
Dim mytiles As Long
    mytiles = Val(GetVar(App.Path & "\myt.ini", "DATA", "TilesetNum"))
    Dim str As String
    Dim str2 As String
    str2 = scrlTileSet.Value
    str = mytiles + 1
    Dim str3 As String
    str3 = mytiles + 1
   Call PutVar(App.Path & "\myt.ini", "DATA", "Tile" & str, txtTilesetName.text)
   Call PutVar(App.Path & "\myt.ini", "DATA", "TileNum" & str, str2)
   Call PutVar(App.Path & "\myt.ini", "DATA", "TilesetNum", str3)
    mytiles = Val(GetVar(App.Path & "\myt.ini", "DATA", "TilesetNum"))
    If mytiles > 0 Then
    TileList.Clear
    For i = 1 To mytiles
    str = i
    TileList.AddItem (GetVar(App.Path & "\myt.ini", "DATA", "Tile" & str))
    Next
    End If
End Sub

Private Sub Command4_Click()
PutVar App.Path & "\usablet.ini", "DATA", Trim$(scrlTileSet.Value), "YES"

End Sub

Private Sub Command5_Click()
WriteText App.Path & "\myt.ini", "[DATA]"
TileList.Clear

End Sub

Private Sub Form_Load()
Dim i As Long
    ' move the entire attributes box on screen
    picAttributes.Left = 8
    picAttributes.Top = 8
    Me.Width = 7425
    If GroundUnvisible Then
    chkVisGround.Value = 0
    Else
    chkVisGround.Value = 1
    End If
    If MaskUnvisible Then
    chkVisMask.Value = 0
    Else
    chkVisMask.Value = 1
    End If
    If Mask2Unvisible Then
    chkVisMask2.Value = 0
    Else
    chkVisMask2.Value = 1
    End If
    If FringeUnvisible Then
    chkVisFringe.Value = 0
    Else
    chkVisFringe.Value = 1
    End If
    If Fringe2Unvisible Then
    chkVisFringe2.Value = 0
    Else
    chkVisFringe2.Value = 1
    End If
    Dim mytiles As Long
    mytiles = Val(GetVar(App.Path & "\myt.ini", "DATA", "TilesetNum"))
    If mytiles > 0 Then
    TileList.Clear
    Dim str As String
    For i = 1 To mytiles
    str = i
    TileList.AddItem (GetVar(App.Path & "\myt.ini", "DATA", "Tile" & str))
    Next
    End If
    
End Sub

Private Sub optCS_Click()
ClearAttributeDialogue
    picAttributes.Visible = True
    frmCS.Visible = True
    
    
    txtcs.text = "1"
End Sub

Private Sub optDoor_Click()
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraMapWarp.Visible = True
    
    scrlMapWarp.Max = MAX_MAPS
    scrlMapWarp.Value = 1
    scrlMapWarpX.Max = MAX_BYTE
    scrlMapWarpY.Max = MAX_BYTE
    scrlMapWarpX.Value = 0
    scrlMapWarpY.Value = 0
End Sub

Private Sub optgymblock_Click()
ClearAttributeDialogue
    picAttributes.Visible = True
    fragymblock.Visible = True
    
    scrlgymnum.Max = MAX_GYMS
    scrlgymnum.Value = 1
End Sub

Private Sub optLayers_Click()

    If optLayers.Value Then
        fraLayers.Visible = True
        fraAttribs.Visible = False
    End If

End Sub

Private Sub optAttribs_Click()

    If optAttribs.Value Then
        fraLayers.Visible = False
        fraAttribs.Visible = True
    End If

End Sub

Private Sub optNpcSpawn_Click()
Dim n As Long

For n = 1 To MAX_MAP_NPCS
    If map.NPC(n) > 0 Then
        lstNpc.AddItem n & ": " & NPC(map.NPC(n)).Name
    Else
        lstNpc.AddItem n & ": No Npc"
    End If
Next n

scrlNpcDir.Value = 0
lstNpc.ListIndex = 0

ClearAttributeDialogue
picAttributes.Visible = True
fraNpcSpawn.Visible = True
End Sub

Private Sub optResource_Click()
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraResource.Visible = True
End Sub

Private Sub optShop_Click()
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraShop.Visible = True
End Sub


Private Sub picBackSelect_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MapEditorChooseTile(Button, x, y)
End Sub
 
Private Sub picBackSelect_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    shpLoc.Top = (y \ PIC_Y) * PIC_Y
    shpLoc.Left = (x \ PIC_X) * PIC_X
    shpLoc.Visible = True
    Call MapEditorDrag(Button, x, y)
End Sub

Private Sub cmdSend_Click()
    Call MapEditorSend
End Sub

Private Sub cmdCancel_Click()
    Call MapEditorCancel
End Sub

Private Sub cmdProperties_Click()
    frmMapProperties.Show vbModal
End Sub

Private Sub optWarp_Click()
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraMapWarp.Visible = True
    
    scrlMapWarp.Max = MAX_MAPS
    scrlMapWarp.Value = 1
    scrlMapWarpX.Max = MAX_BYTE
    scrlMapWarpY.Max = MAX_BYTE
    scrlMapWarpX.Value = 0
    scrlMapWarpY.Value = 0
End Sub

Private Sub optItem_Click()
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraMapItem.Visible = True

    scrlMapItem.Max = MAX_ITEMS
    scrlMapItem.Value = 1
    lblMapItem.Caption = Trim$(Item(scrlMapItem.Value).Name) & " x" & scrlMapItemValue.Value
    EditorMap_BltMapItem
End Sub

Private Sub optKey_Click()
'    frmMapKey.Show vbModal
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraMapKey.Visible = True
    
    scrlMapKey.Max = MAX_ITEMS
    scrlMapKey.Value = 1
    chkMapKey.Value = 1
    EditorMap_BltKey
    lblMapKey.Caption = "Item: " & Trim$(Item(scrlMapKey.Value).Name)
End Sub

Private Sub optKeyOpen_Click()
    ClearAttributeDialogue
    fraKeyOpen.Visible = True
    picAttributes.Visible = True
    
    scrlKeyX.Max = map.MaxX
    scrlKeyY.Max = map.MaxY
    scrlKeyX.Value = 0
    scrlKeyY.Value = 0
End Sub

Private Sub cmdFill_Click()
    MapEditorFillLayer
End Sub

Private Sub cmdClear_Click()
    Call MapEditorClearLayer
End Sub

Private Sub cmdClear2_Click()
    Call MapEditorClearAttribs
End Sub

Private Sub scrlgymnum_Change()
fragymblock.Caption = "Gym:" & scrlgymnum.Value
End Sub

Private Sub scrlKeyX_Change()
lblKeyX.Caption = "X: " & scrlKeyX.Value
End Sub

Private Sub scrlKeyX_Scroll()
scrlKeyX_Change
End Sub

Private Sub scrlKeyY_Change()
lblKeyY.Caption = "Y: " & scrlKeyY.Value
End Sub

Private Sub scrlKeyY_Scroll()
scrlKeyY_Change
End Sub

Private Sub scrlMapItem_Change()
If Item(scrlMapItem.Value).Type = ITEM_TYPE_CURRENCY Then
    scrlMapItemValue.Enabled = True
Else
    scrlMapItemValue.Value = 1
    scrlMapItemValue.Enabled = False
End If
    
EditorMap_BltMapItem
lblMapItem.Caption = Trim$(Item(scrlMapItem.Value).Name) & " x" & scrlMapItemValue.Value
End Sub

Private Sub scrlMapItem_Scroll()
    scrlMapItem_Change
End Sub

Private Sub scrlMapItemValue_Change()
lblMapItem.Caption = Trim$(Item(scrlMapItem.Value).Name) & " x" & scrlMapItemValue.Value
End Sub

Private Sub scrlMapItemValue_Scroll()
scrlMapItemValue_Change
End Sub

Private Sub scrlMapKey_Change()
lblMapKey.Caption = "Item: " & Trim$(Item(scrlMapKey.Value).Name)
End Sub

Private Sub scrlMapKey_Scroll()
scrlMapKey_Change
End Sub

Private Sub scrlMapWarp_Change()
lblMapWarp.Caption = "Map: " & scrlMapWarp.Value
End Sub

Private Sub scrlMapWarp_Scroll()
scrlMapWarp_Change
End Sub

Private Sub scrlMapWarpX_Change()
lblMapWarpX.Caption = "X: " & scrlMapWarpX.Value
End Sub

Private Sub scrlMapWarpX_Scroll()
scrlMapWarpX_Change
End Sub

Private Sub scrlMapWarpY_Change()
lblMapWarpY.Caption = "Y: " & scrlMapWarpY.Value
End Sub

Private Sub scrlMapWarpY_Scroll()
scrlMapWarpY_Change
End Sub

Private Sub scrlNpcDir_Change()
Select Case scrlNpcDir.Value
    Case DIR_DOWN
        lblNpcDir = "Direction: Down"
    Case DIR_UP
        lblNpcDir = "Direction: Up"
    Case DIR_LEFT
        lblNpcDir = "Direction: Left"
    Case DIR_RIGHT
        lblNpcDir = "Direction: Right"
End Select
End Sub

Private Sub scrlNpcDir_Scroll()
    scrlNpcDir_Change
End Sub

Private Sub scrlResource_Change()
    lblResource.Caption = "Resource: " & Resource(scrlResource.Value).Name
End Sub

Private Sub scrlResource_Scroll()
    scrlResource_Change
End Sub

Private Sub scrlPictureX_Change()
    Call MapEditorTileScroll
End Sub

Private Sub scrlPictureY_Change()
    Call MapEditorTileScroll
End Sub

Private Sub scrlPictureX_Scroll()
    scrlPictureY_Change
End Sub

Private Sub scrlPictureY_Scroll()
    scrlPictureY_Change
End Sub

Private Sub scrlTileSet_Change()
    map.tileset = scrlTileSet.Value
    fraTileSet.Caption = "Tileset: " & scrlTileSet.Value

    Call EditorMap_BltTileset
    
    frmEditor_Map.scrlPictureY.Max = (frmEditor_Map.picBackSelect.Height \ PIC_Y) - (frmEditor_Map.picBack.Height \ PIC_Y)
    frmEditor_Map.scrlPictureX.Max = (frmEditor_Map.picBackSelect.Width \ PIC_X) - (frmEditor_Map.picBack.Width \ PIC_X)
End Sub

Private Sub scrlTileSet_Scroll()
    
    
    scrlTileSet_Change
End Sub

Private Sub TileList_Click()
scrlTileSet.Value = Val(GetVar(App.Path & "\myt.ini", "DATA", "TileNum" & TileList.ListIndex + 1))
    map.tileset = Val(GetVar(App.Path & "\myt.ini", "DATA", "TileNum" & TileList.ListIndex + 1))
    fraTileSet.Caption = "Tileset: " & Val(GetVar(App.Path & "\myt.ini", "DATA", "TileNum" & TileList.ListIndex + 1))
    
    Call EditorMap_BltTileset
    
    frmEditor_Map.scrlPictureY.Max = (frmEditor_Map.picBackSelect.Height \ PIC_Y) - (frmEditor_Map.picBack.Height \ PIC_Y)
    frmEditor_Map.scrlPictureX.Max = (frmEditor_Map.picBackSelect.Width \ PIC_X) - (frmEditor_Map.picBack.Width \ PIC_X)
End Sub

Function isValidTileset(ByVal num As Long) As Boolean
If GetVar(App.Path & "\usablet.ini", "DATA", Trim$(num)) = "YES" Then
isValidTileset = False
Else
isValidTileset = True
End If

End Function
