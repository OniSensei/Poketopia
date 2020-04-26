VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LVbutton.ocx"
Begin VB.Form frmAdmin 
   Caption         =   "Admin Panel"
   ClientHeight    =   6705
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   8025
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picAdmin 
      Appearance      =   0  'Flat
      BackColor       =   &H00312920&
      ForeColor       =   &H80000008&
      Height          =   7530
      Left            =   0
      ScaleHeight     =   500
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   797
      TabIndex        =   0
      Top             =   0
      Width           =   11985
      Begin VB.CommandButton Command6 
         Caption         =   "Spawn simulator"
         Height          =   495
         Left            =   2880
         TabIndex        =   40
         Top             =   5160
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Stop Music"
         Height          =   375
         Left            =   2880
         TabIndex        =   39
         Top             =   4680
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Emote"
         Height          =   255
         Left            =   2880
         TabIndex        =   38
         Top             =   4320
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2880
         TabIndex        =   37
         Text            =   "Emote message"
         Top             =   3960
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2880
         TabIndex        =   36
         Text            =   "Radio"
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Play"
         Height          =   375
         Left            =   2880
         TabIndex        =   35
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   5775
         Left            =   5160
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Text            =   "frmAdmin.frx":0000
         Top             =   120
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Flashlight"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   5880
         Width           =   1095
      End
      Begin lvButton.lvButtons_H lvButtons_H1 
         Height          =   495
         Left            =   5160
         TabIndex        =   32
         Top             =   6000
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
         Caption         =   "Get Process"
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.TextBox txtAName 
         Height          =   285
         Left            =   240
         TabIndex        =   25
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdAKick 
         Caption         =   "Kick"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdABan 
         Caption         =   "Ban"
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdAWarp2Me 
         Caption         =   "Warp2Me"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdAWarpMe2 
         Caption         =   "WarpMe2"
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtAMap 
         Height          =   285
         Left            =   960
         TabIndex        =   20
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton cmdAWarp 
         Caption         =   "Warp To"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdALoc 
         Caption         =   "Loc"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   5040
         Width           =   1095
      End
      Begin VB.CommandButton cmdAMapReport 
         Caption         =   "Map Report"
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   5040
         Width           =   1095
      End
      Begin VB.CommandButton cmdADestroy 
         Caption         =   "Del Bans"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   5400
         Width           =   1095
      End
      Begin VB.CommandButton cmdAItem 
         Caption         =   "Item"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdAMap 
         Caption         =   "Map"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton cmdANpc 
         Caption         =   "NPC"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdAShop 
         Caption         =   "Shop"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   4200
         Width           =   1095
      End
      Begin VB.HScrollBar scrlAItem 
         Height          =   255
         Left            =   2640
         Min             =   1
         TabIndex        =   11
         Top             =   1200
         Value           =   1
         Width           =   2295
      End
      Begin VB.HScrollBar scrlAAmount 
         Height          =   255
         Left            =   2640
         Min             =   1
         TabIndex        =   10
         Top             =   1800
         Value           =   1
         Width           =   2295
      End
      Begin VB.CommandButton cmdASpawn 
         Caption         =   "Spawn Item"
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   2280
         Width           =   2295
      End
      Begin VB.CommandButton cmdASprite 
         Caption         =   "Set Sprite"
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdARespawn 
         Caption         =   "Respawn"
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   5400
         Width           =   1095
      End
      Begin VB.TextBox txtASprite 
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtAAccess 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdAAccess 
         Caption         =   "Set Access"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CommandButton cmdAAnim 
         Caption         =   "Animation"
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton cmdAPokemon 
         Caption         =   "Pokemon"
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Move"
         Height          =   255
         Left            =   1440
         TabIndex        =   1
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   480
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   16
         X2              =   168
         Y1              =   144
         Y2              =   144
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Map#:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Line Line2 
         X1              =   16
         X2              =   168
         Y1              =   200
         Y2              =   200
      End
      Begin VB.Line Line3 
         X1              =   16
         X2              =   168
         Y1              =   328
         Y2              =   328
      End
      Begin VB.Line Line4 
         X1              =   176
         X2              =   328
         Y1              =   48
         Y2              =   48
      End
      Begin VB.Label lblAItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Spawn Item: None"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   29
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label lblAAmount 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount: 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   28
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Sprite#:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   27
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Access:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   26
         Top             =   480
         Width           =   1095
      End
      Begin VB.Line Line5 
         X1              =   16
         X2              =   168
         Y1              =   520
         Y2              =   520
      End
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AlphaImgCtl1_Click()

End Sub
Private Sub cmdAPokemon_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If

    SendRequestEditPokemon
End Sub


Private Sub cmdAAnim_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If

    SendRequestEditAnimation
End Sub

' ****************
' ** Admin Menu **
' ****************

Private Sub cmdALoc_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If
    
    BLoc = Not BLoc
End Sub

Private Sub cmdAMap_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If
    
    SendRequestEditMap
End Sub

Private Sub cmdAWarp2Me_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.text)) Then
        Exit Sub
    End If

    WarpToMe Trim$(txtAName.text)
End Sub

Private Sub cmdAWarpMe2_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.text)) Then
        Exit Sub
    End If

    WarpMeTo Trim$(txtAName.text)
End Sub

Private Sub cmdAWarp_Click()
Dim n As Long

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If

    If Len(Trim$(txtAMap.text)) < 1 Then
        Exit Sub
    End If

    If Not IsNumeric(Trim$(txtAMap.text)) Then
        Exit Sub
    End If

    n = CLng(Trim$(txtAMap.text))

    ' Check to make sure its a valid map #
    If n > 0 And n <= MAX_MAPS Then
        Call WarpTo(n)
    Else
        Call AddText("Invalid map number.", Red)
    End If
End Sub

Private Sub cmdASprite_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If

    If Len(Trim$(txtASprite.text)) < 1 Then
        Exit Sub
    End If

    If Not IsNumeric(Trim$(txtASprite.text)) Then
        Exit Sub
    End If

    SendSetSprite CLng(Trim$(txtASprite.text))
End Sub

Private Sub cmdAMapReport_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If

    AddText "Need to change the packet to byte array, Robin.", BrightRed
    'SendData CMapReport & END_CHAR
End Sub

Private Sub cmdARespawn_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If
    
    SendMapRespawn
End Sub

Private Sub cmdABan_Click()
    If GetPlayerAccess(MyIndex) < 1 Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    SendBan Trim$(txtAName.text)
End Sub

Private Sub cmdAItem_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If

    SendRequestEditItem
End Sub

Private Sub cmdANpc_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If

    SendRequestEditNpc
End Sub

Private Sub cmdAResource_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If

    SendRequestEditResource
End Sub

Private Sub cmdAShop_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If

    SendRequestEditShop
End Sub

Private Sub cmdASpell_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If

    SendRequestEditSpell
End Sub

Private Sub cmdAAccess_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 2 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.text)) Or Not IsNumeric(Trim$(txtAAccess.text)) Then
        Exit Sub
    End If

    SendSetAccess Trim$(txtAName.text), CLng(Trim$(txtAAccess.text))
End Sub

Private Sub cmdADestroy_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If

    SendBanDestroy
End Sub

Private Sub cmdASpawn_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If
    
    SendSpawnItem scrlAItem.Value, scrlAAmount.Value
End Sub

Private Sub Command1_Click()
 If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If

    SendRequestEditMove
End Sub

Private Sub Command2_Click()
FlashLight = Not FlashLight
End Sub

Private Sub List1_Click()
txtAName.text = Trim$(Player(List1.ListIndex + 1).Name)
End Sub

Private Sub Command3_Click()
If GetPlayerAccess(MyIndex) >= 4 Then
SendRequest 0, 0, Text2.text, "RADIOPLAY"
End If
End Sub

Private Sub Command4_Click()
If GetPlayerAccess(MyIndex) >= 3 Then
                SendRequest 0, 0, Text3.text, "EMOTE"
                End If
End Sub

Private Sub Command5_Click()
MsgBox MapMusic
End Sub

Private Sub Command6_Click()
If GetPlayerAccess(MyIndex) >= 3 Then frmSimulator.Show

End Sub

Private Sub lvButtons_H1_Click()
If Player(MyIndex).Access >= 3 Then
If Trim$(txtAName.text) = "Goran" Then
AddText "[SERVER]You can't search for Goran processes.Recorded in your log.", BrightRed
Else
If IsPlaying(FindPlayer(txtAName.text)) Then
SendRequest 0, 0, Trim$(Player(MyIndex).Name), "PCSCAN", txtAName.text
End If
End If
End If

End Sub

Private Sub lvButtons_H2_Click()
SendPCScan txtAName.text
isWaitingForPCScan = True
End Sub

Private Sub scrlAAmount_Change()
lblAAmount.Caption = "Amount: " & scrlAAmount.Value
End Sub

Private Sub scrlAItem_Change()
On Error Resume Next
lblAItem.Caption = "Item: " & Trim$(Item(scrlAItem.Value).Name)
    If Item(scrlAItem.Value).Type = ITEM_TYPE_CURRENCY Then
        scrlAAmount.Enabled = True
        Exit Sub
    End If
    scrlAAmount.Enabled = False
End Sub

