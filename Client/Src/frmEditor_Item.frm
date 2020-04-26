VERSION 5.00
Begin VB.Form frmEditor_Item 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14280
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
   Icon            =   "frmEditor_Item.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   952
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraPotionData 
      Caption         =   "Potion"
      Height          =   1215
      Left            =   3360
      TabIndex        =   72
      Top             =   3360
      Width           =   4575
      Begin VB.HScrollBar scrlPotionHP 
         Height          =   255
         Left            =   360
         Max             =   10000
         TabIndex        =   73
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label lblPotionHP 
         Caption         =   "+HP : 0"
         Height          =   255
         Left            =   360
         TabIndex        =   74
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Paperdoll"
      Height          =   3015
      Left            =   9720
      TabIndex        =   71
      Top             =   1440
      Width           =   4575
      Begin VB.PictureBox picPaperdoll 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2560
         Left            =   120
         ScaleHeight     =   256
         ScaleMode       =   0  'User
         ScaleWidth      =   256
         TabIndex        =   75
         Top             =   240
         Width           =   2560
      End
   End
   Begin VB.Frame frmPokeball 
      Caption         =   "Pokeball"
      Height          =   1215
      Left            =   9720
      TabIndex        =   68
      Top             =   120
      Width           =   4575
      Begin VB.HScrollBar scrlCatchRate 
         Height          =   255
         Left            =   120
         Max             =   500
         TabIndex        =   69
         Top             =   720
         Value           =   1
         Width           =   4335
      End
      Begin VB.Label lvlrate 
         Caption         =   "Catch Rate : 1"
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Info"
      Height          =   1695
      Left            =   3360
      TabIndex        =   25
      Top             =   120
      Width           =   6255
      Begin VB.HScrollBar scrlRarity 
         Height          =   255
         Left            =   4200
         Max             =   5
         TabIndex        =   33
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox cmbBind 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":020A
         Left            =   4200
         List            =   "frmEditor_Item.frx":0217
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   600
         Width           =   1935
      End
      Begin VB.HScrollBar scrlPrice 
         Height          =   255
         LargeChange     =   100
         Left            =   4200
         Max             =   30000
         TabIndex        =   31
         Top             =   240
         Width           =   1935
      End
      Begin VB.HScrollBar scrlAnim 
         Height          =   255
         Left            =   5040
         Max             =   5
         TabIndex        =   30
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":0240
         Left            =   120
         List            =   "frmEditor_Item.frx":0283
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtName 
         Height          =   255
         Left            =   720
         TabIndex        =   28
         Top             =   240
         Width           =   2055
      End
      Begin VB.HScrollBar scrlPic 
         Height          =   255
         Left            =   840
         Max             =   255
         TabIndex        =   27
         Top             =   600
         Width           =   1335
      End
      Begin VB.PictureBox picItem 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2280
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   26
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblRarity 
         AutoSize        =   -1  'True
         Caption         =   "Rarity: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   39
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Bind Type:"
         Height          =   180
         Left            =   2880
         TabIndex        =   38
         Top             =   600
         Width           =   810
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         Caption         =   "Price: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   37
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         Caption         =   "Anim: None"
         Height          =   180
         Left            =   2880
         TabIndex        =   36
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   35
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
         Caption         =   "Pic: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   34
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Requirements"
      Height          =   1455
      Left            =   3360
      TabIndex        =   6
      Top             =   1920
      Width           =   6255
      Begin VB.HScrollBar scrlLevelReq 
         Height          =   255
         LargeChange     =   10
         Left            =   1560
         Max             =   99
         TabIndex        =   15
         Top             =   1080
         Width           =   4455
      End
      Begin VB.HScrollBar scrlAccessReq 
         Height          =   255
         Left            =   1560
         Max             =   5
         TabIndex        =   14
         Top             =   720
         Width           =   4455
      End
      Begin VB.ComboBox cmbClassReq 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":0358
         Left            =   1200
         List            =   "frmEditor_Item.frx":035A
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   360
         Width           =   4815
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   840
         Max             =   255
         TabIndex        =   12
         Top             =   1440
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   3000
         Max             =   255
         TabIndex        =   11
         Top             =   1440
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   5160
         Max             =   255
         TabIndex        =   10
         Top             =   1440
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   840
         Max             =   255
         TabIndex        =   9
         Top             =   1800
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   3000
         Max             =   255
         TabIndex        =   8
         Top             =   1800
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   6
         LargeChange     =   10
         Left            =   5160
         Max             =   255
         TabIndex        =   7
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblLevelReq 
         AutoSize        =   -1  'True
         Caption         =   "Level req: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label lblAccessReq 
         AutoSize        =   -1  'True
         Caption         =   "Access Req: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Class Req:"
         Height          =   180
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   825
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Str: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "End: 0"
         Height          =   180
         Index           =   2
         Left            =   2280
         TabIndex        =   20
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Vit: 0"
         Height          =   180
         Index           =   3
         Left            =   4440
         TabIndex        =   19
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Will: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         UseMnemonic     =   0   'False
         Width           =   480
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Int: 0"
         Height          =   180
         Index           =   5
         Left            =   2280
         TabIndex        =   17
         Top             =   1800
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Spr: 0"
         Height          =   180
         Index           =   6
         Left            =   4440
         TabIndex        =   16
         Top             =   1800
         UseMnemonic     =   0   'False
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   6480
      Width           =   2895
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Item List"
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   6900
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraSpell 
      Caption         =   "Spell Data"
      Height          =   1215
      Left            =   14160
      TabIndex        =   62
      Top             =   3840
      Visible         =   0   'False
      Width           =   3735
      Begin VB.HScrollBar scrlSpell 
         Height          =   255
         Left            =   1080
         Max             =   255
         Min             =   1
         TabIndex        =   63
         Top             =   720
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label lblSpellName 
         AutoSize        =   -1  'True
         Caption         =   "Name: None"
         Height          =   180
         Left            =   240
         TabIndex        =   65
         Top             =   360
         Width           =   930
      End
      Begin VB.Label lblSpell 
         AutoSize        =   -1  'True
         Caption         =   "Num: 0"
         Height          =   180
         Left            =   240
         TabIndex        =   64
         Top             =   720
         Width           =   555
      End
   End
   Begin VB.Frame fraEquipment 
      Caption         =   "Equipment Data"
      Height          =   1095
      Left            =   3360
      TabIndex        =   40
      Top             =   3360
      Visible         =   0   'False
      Width           =   6255
      Begin VB.HScrollBar scrlPaperdoll 
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   600
         Width           =   1455
      End
      Begin VB.HScrollBar scrlSpeed 
         Height          =   255
         LargeChange     =   100
         Left            =   1920
         Max             =   3000
         Min             =   100
         SmallChange     =   100
         TabIndex        =   49
         Top             =   1200
         Value           =   100
         Width           =   4095
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   6
         LargeChange     =   10
         Left            =   5160
         Max             =   255
         TabIndex        =   48
         Top             =   1920
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   3000
         Max             =   255
         TabIndex        =   47
         Top             =   1920
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   46
         Top             =   1920
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   5160
         Max             =   255
         TabIndex        =   45
         Top             =   1560
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   3000
         Max             =   255
         TabIndex        =   44
         Top             =   1560
         Width           =   855
      End
      Begin VB.HScrollBar scrlDamage 
         Height          =   255
         LargeChange     =   10
         Left            =   1440
         Max             =   255
         TabIndex        =   43
         Top             =   1440
         Width           =   4575
      End
      Begin VB.ComboBox cmbTool 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":035C
         Left            =   360
         List            =   "frmEditor_Item.frx":036C
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   2280
         Width           =   4695
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   41
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblPaperdoll 
         AutoSize        =   -1  'True
         Caption         =   "Paperdoll: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   66
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lblSpeed 
         AutoSize        =   -1  'True
         Caption         =   "Speed: 0.1 sec"
         Height          =   180
         Left            =   120
         TabIndex        =   58
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   1140
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Spr: 0"
         Height          =   180
         Index           =   6
         Left            =   4320
         TabIndex        =   57
         Top             =   1920
         UseMnemonic     =   0   'False
         Width           =   615
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Int: 0"
         Height          =   180
         Index           =   5
         Left            =   2160
         TabIndex        =   56
         Top             =   1920
         UseMnemonic     =   0   'False
         Width           =   585
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Will: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   55
         Top             =   1920
         UseMnemonic     =   0   'False
         Width           =   630
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Vit: 0"
         Height          =   180
         Index           =   3
         Left            =   4320
         TabIndex        =   54
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   585
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ End: 0"
         Height          =   180
         Index           =   2
         Left            =   2160
         TabIndex        =   53
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   645
      End
      Begin VB.Label lblDamage 
         AutoSize        =   -1  'True
         Caption         =   "Damage: 0"
         Height          =   180
         Left            =   0
         TabIndex        =   52
         Top             =   2040
         UseMnemonic     =   0   'False
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Object Tool:"
         Height          =   180
         Left            =   120
         TabIndex        =   51
         Top             =   2760
         Width           =   945
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Str: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   50
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   585
      End
   End
   Begin VB.Frame fraVitals 
      Caption         =   "Vitals Data"
      Height          =   1215
      Left            =   3360
      TabIndex        =   59
      Top             =   4680
      Visible         =   0   'False
      Width           =   4575
      Begin VB.HScrollBar scrlVitalMod 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   60
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label lblVitalMod 
         AutoSize        =   -1  'True
         Caption         =   "Vital Mod: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   61
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmEditor_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LastIndex As Long

Private Sub cmbBind_Click()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).BindType = cmbBind.ListIndex
End Sub

Private Sub cmbClassReq_Click()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).ClassReq = cmbClassReq.ListIndex
End Sub

Private Sub cmbTool_Click()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).data3 = cmbTool.ListIndex
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

ClearItem EditorIndex

tmpIndex = lstIndex.ListIndex
lstIndex.RemoveItem EditorIndex - 1
lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).Name, EditorIndex - 1
lstIndex.ListIndex = tmpIndex

ItemEditorInit
End Sub

Private Sub Form_Load()
    scrlPic.Max = NumItems
    scrlAnim.Max = MAX_ANIMATIONS
    scrlPaperdoll.Max = NumPaperdolls
End Sub

Private Sub cmdSave_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
    Call ItemEditorOk
    End If
End Sub

Private Sub cmdCancel_Click()
    Call ItemEditorCancel
End Sub

Private Sub cmbType_Click()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    If (cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        fraEquipment.Visible = True
        'scrlDamage_Change
    Else
    If (cmbType.ListIndex >= ITEM_TYPE_MASK) And (cmbType.ListIndex <= ITEM_TYPE_OUTFIT) Then
        fraEquipment.Visible = True
        'scrlDamage_Change
    Else
        fraEquipment.Visible = False
    End If
    End If
    
    

    If (cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        fraVitals.Visible = True
        'scrlVitalMod_Change
    Else
        fraVitals.Visible = False
    End If

    If (cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        fraSpell.Visible = True
    Else
        fraSpell.Visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_POKEBALL) Then
        frmPokeball.Visible = True
    Else
        frmPokeball.Visible = False
    End If
    If (cmbType.ListIndex = ITEM_TYPE_POKEPOTION) Then
        fraPotionData.Visible = True
        scrlPotionHP.Value = Item(EditorIndex).AddHP
        lblPotionHP.Caption = "+HP : " & Item(EditorIndex).AddHP
    Else
        fraPotionData.Visible = False
    End If
    
    Item(EditorIndex).Type = cmbType.ListIndex
 
End Sub

Private Sub lblStr_Click()

End Sub

Private Sub lstIndex_Click()
    ItemEditorInit
End Sub


Private Sub scrlAccessReq_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblAccessReq.Caption = "Access Req: " & scrlAccessReq.Value
    Item(EditorIndex).AccessReq = scrlAccessReq.Value
End Sub

Private Sub scrlAnim_Change()
Dim sString As String
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    If scrlAnim.Value = 0 Then
        sString = "None"
    Else
        sString = Trim$(Animation(scrlAnim.Value).Name)
    End If
    lblAnim.Caption = "Anim: " & sString
    Item(EditorIndex).Animation = scrlAnim.Value
End Sub

Private Sub scrlCatchRate_Change()
 If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lvlrate.Caption = "Catch Rate : " & scrlCatchRate.Value
    Item(EditorIndex).CatchRate = scrlCatchRate.Value
End Sub

Private Sub scrlDamage_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblDamage.Caption = "Damage: " & scrlDamage.Value
    Item(EditorIndex).data2 = scrlDamage.Value
End Sub

Private Sub scrlLevelReq_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblLevelReq.Caption = "Level req: " & scrlLevelReq
    Item(EditorIndex).LevelReq = scrlLevelReq.Value
End Sub

Private Sub scrlPaperdoll_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPaperdoll.Caption = "Paperdoll: " & scrlPaperdoll.Value
    Item(EditorIndex).Paperdoll = scrlPaperdoll.Value
    Call EditorItem_BltPaperdoll
End Sub

Private Sub scrlPic_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPic.Caption = "Pic: " & scrlPic.Value
    Item(EditorIndex).pic = scrlPic.Value
    Call EditorItem_BltItem
End Sub

Private Sub scrlPotionHP_Change()
lblPotionHP.Caption = "+HP : " & scrlPotionHP.Value
Item(EditorIndex).AddHP = scrlPotionHP.Value
End Sub

Private Sub scrlPrice_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPrice.Caption = "Price: " & scrlPrice.Value
    Item(EditorIndex).Price = scrlPrice.Value
End Sub

Private Sub scrlRarity_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblRarity.Caption = "Rarity: " & scrlRarity.Value
    Item(EditorIndex).Rarity = scrlRarity.Value
End Sub

Private Sub scrlSpeed_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblSPEED.Caption = "Speed: " & scrlSpeed.Value / 1000 & " sec"
    Item(EditorIndex).speed = scrlSpeed.Value
End Sub

Private Sub scrlStatBonus_Change(Index As Integer)
    Dim text As String
    
    Select Case Index
        Case 1 ' str
            text = "+ Str: "
        Case 2 ' end
            text = "+ End: "
        Case 3 ' vit
            text = "+ Vit: "
        Case 4 ' int
            text = "+ Int: "
        Case 5 ' will
            text = "+ Will: "
        Case 6 ' spr
            text = "+ Spr: "
    End Select
            
    lblStatBonus(Index).Caption = text & scrlStatBonus(Index).Value
    Item(EditorIndex).Add_Stat(Index) = scrlStatBonus(Index).Value
End Sub

Private Sub scrlStatReq_Change(Index As Integer)
    Dim text As String
    
    Select Case Index
        Case 1 ' str
            text = "Str: "
        Case 2 ' end
            text = "End: "
        Case 3 ' vit
            text = "Vit: "
        Case 4 ' int
            text = "Int: "
        Case 5 ' will
            text = "Will: "
        Case 6 ' spr
            text = "Spr: "
    End Select
    
    lblStatReq(Index).Caption = text & scrlStatReq(Index).Value
    Item(EditorIndex).Stat_Req(Index) = scrlStatReq(Index).Value
End Sub

Private Sub scrlVitalMod_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblVitalMod.Caption = "Vital Mod: " & scrlVitalMod.Value
    Item(EditorIndex).data1 = scrlVitalMod.Value
End Sub

Private Sub scrlSpell_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    If Len(Trim$(Spell(scrlSpell.Value).Name)) > 0 Then
        lblSpellName.Caption = "Name: " & Trim$(Spell(scrlSpell.Value).Name)
    Else
        lblSpellName.Caption = "Name: None"
    End If
    
    lblSpell.Caption = "Spell: " & scrlSpell.Value
    
    Item(EditorIndex).data1 = scrlSpell.Value
End Sub



Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Item(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
End Sub
