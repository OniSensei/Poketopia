VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCN.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmServer 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Poketopia [ Revival ] "
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   10
      TabHeight       =   503
      BackColor       =   -2147483644
      TabCaption(0)   =   "Console"
      TabPicture(0)   =   "frmServer.frx":1708A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtText"
      Tab(0).Control(1)=   "txtChat"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Players"
      TabPicture(1)   =   "frmServer.frx":170A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "lvwInfo"
      Tab(1).Control(2)=   "HScroll1"
      Tab(1).Control(3)=   "Command7"
      Tab(1).Control(4)=   "Frame5"
      Tab(1).Control(5)=   "Frame6"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Control "
      TabPicture(2)   =   "frmServer.frx":170C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraDatabase"
      Tab(2).Control(1)=   "fraServer"
      Tab(2).Control(2)=   "Frame4"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Event"
      TabPicture(3)   =   "frmServer.frx":170DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).Control(1)=   "tmrglobalspawn"
      Tab(3).Control(2)=   "fraItem2All"
      Tab(3).Control(3)=   "Frame2"
      Tab(3).Control(4)=   "Frame3"
      Tab(3).Control(5)=   "Command2"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Database"
      TabPicture(4)   =   "frmServer.frx":170FA
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Command4"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Command5"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Command6"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Command10"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      Begin VB.CommandButton Command2 
         Caption         =   "35 EXP"
         Height          =   255
         Left            =   -72000
         TabIndex        =   65
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Frame Frame6 
         Caption         =   "Pokemon"
         Height          =   1695
         Left            =   -70920
         TabIndex        =   58
         Top             =   1920
         Width           =   1455
         Begin VB.HScrollBar scrlPokemon 
            Height          =   255
            Left            =   120
            Max             =   721
            Min             =   1
            TabIndex        =   62
            Top             =   480
            Value           =   1
            Width           =   975
         End
         Begin VB.TextBox txtlvl 
            Height          =   285
            Left            =   120
            TabIndex        =   61
            Text            =   "Level"
            Top             =   840
            Width           =   495
         End
         Begin VB.CheckBox chkPokeShiny 
            Caption         =   "Shiny?"
            Height          =   495
            Left            =   120
            TabIndex        =   60
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtNat 
            Height          =   285
            Left            =   600
            TabIndex        =   59
            Text            =   "Nature"
            Top             =   840
            Width           =   495
         End
         Begin VB.Label lblPokemon 
            Caption         =   "Pokemon: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Dialog"
         Height          =   1095
         Left            =   -70920
         TabIndex        =   55
         Top             =   720
         Width           =   1455
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   120
            TabIndex        =   57
            Text            =   "Image Number"
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   120
            TabIndex        =   56
            Text            =   "Text Here"
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Crew"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   52
         Top             =   3480
         Width           =   1455
         Begin VB.CommandButton Command11 
            Caption         =   "Delete"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   120
            TabIndex        =   53
            Text            =   "Name"
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Type effects"
         Height          =   375
         Left            =   240
         TabIndex        =   51
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Caption         =   "Whos dat pokemon"
         Height          =   1695
         Left            =   -72120
         TabIndex        =   45
         Top             =   2640
         Width           =   2535
         Begin VB.TextBox txtWhosNum 
            Height          =   285
            Left            =   1800
            TabIndex        =   50
            Text            =   "Num"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtWhosVal 
            Height          =   285
            Left            =   1200
            TabIndex        =   49
            Text            =   "Val"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtWhosItem 
            Height          =   285
            Left            =   240
            TabIndex        =   48
            Text            =   "Item"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtWhosAnswer 
            Height          =   285
            Left            =   240
            TabIndex        =   47
            Text            =   "Answer"
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Start"
            Height          =   255
            Left            =   480
            TabIndex        =   46
            Top             =   1320
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Item2One"
         Height          =   2175
         Left            =   -74880
         TabIndex        =   37
         Top             =   3000
         Width           =   2535
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   120
            TabIndex        =   43
            Text            =   "Name"
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   120
            TabIndex        =   40
            Text            =   "1"
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   120
            TabIndex        =   39
            Text            =   "1"
            Top             =   1680
            Width           =   1455
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Give"
            Height          =   255
            Left            =   1680
            TabIndex        =   38
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Player"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Item"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Value"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   1440
            Width           =   975
         End
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Edit Player"
         Height          =   300
         Left            =   -70920
         TabIndex        =   36
         Top             =   4320
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Save Moves"
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Move data"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1035
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Load Pokemon"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   555
         Width           =   2175
      End
      Begin VB.Frame fraItem2All 
         Caption         =   "Item2All"
         Height          =   2055
         Left            =   -72120
         TabIndex        =   27
         Top             =   480
         Width           =   2535
         Begin VB.CommandButton Command3 
            Caption         =   "Give"
            Height          =   255
            Left            =   480
            TabIndex        =   31
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   240
            TabIndex        =   30
            Text            =   "1"
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   240
            TabIndex        =   28
            Text            =   "1"
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Value"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   720
            Width           =   975
         End
      End
      Begin VB.Timer tmrglobalspawn 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   -74760
         Top             =   4275
      End
      Begin VB.Frame Frame1 
         Caption         =   "GS"
         Height          =   2655
         Left            =   -74880
         TabIndex        =   19
         Top             =   360
         Width           =   2655
         Begin VB.CheckBox chkisShiny 
            Caption         =   "Is Shiny?"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   2160
            Width           =   2295
         End
         Begin VB.HScrollBar HScroll3 
            Height          =   255
            Left            =   120
            Max             =   1000
            Min             =   1
            TabIndex        =   23
            Top             =   1200
            Value           =   1
            Width           =   2295
         End
         Begin VB.HScrollBar HScroll2 
            Height          =   255
            Left            =   120
            Max             =   721
            Min             =   1
            TabIndex        =   21
            Top             =   480
            Value           =   1
            Width           =   2295
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Global Spawn"
            Height          =   375
            Left            =   1200
            TabIndex        =   20
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label lblGSlvl 
            Caption         =   "Lvl.1"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   960
            Width           =   2295
         End
         Begin VB.Label lblGSname 
            Caption         =   "Bulbasaur"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   -70920
         Max             =   20
         Min             =   1
         TabIndex        =   17
         Top             =   3960
         Value           =   1
         Width           =   1335
      End
      Begin VB.Frame fraServer 
         Caption         =   "Server"
         Height          =   2775
         Left            =   -71880
         TabIndex        =   1
         Top             =   675
         Width           =   1455
         Begin VB.CheckBox Check1 
            Caption         =   "Admin only"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   1440
            Width           =   1215
         End
         Begin VB.CheckBox chkServerLog 
            Caption         =   "Server Log"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "Exit"
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdShutDown 
            Caption         =   "Shut Down"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame fraDatabase 
         Caption         =   "Reload"
         Height          =   2775
         Left            =   -74880
         TabIndex        =   8
         Top             =   675
         Width           =   2895
         Begin VB.CommandButton cmdReloadAnimations 
            Caption         =   "Animations"
            Height          =   375
            Left            =   1440
            TabIndex        =   16
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadResources 
            Caption         =   "Resources"
            Height          =   375
            Left            =   1440
            TabIndex        =   15
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadItems 
            Caption         =   "Items"
            Height          =   375
            Left            =   1440
            TabIndex        =   14
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadNPCs 
            Caption         =   "Npcs"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   2160
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadShops 
            Caption         =   "Shops"
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton CmdReloadSpells 
            Caption         =   "Spells"
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadMaps 
            Caption         =   "Maps"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadClasses 
            Caption         =   "Classes"
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox txtChat 
         Height          =   375
         Left            =   -74880
         TabIndex        =   3
         Top             =   4320
         Width           =   5895
      End
      Begin VB.TextBox txtText 
         Height          =   3735
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   510
         Width           =   5895
      End
      Begin MSComctlLib.ListView lvwInfo 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   4
         Top             =   795
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   6800
         View            =   3
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Index"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "IP Address"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Account"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Character"
            Object.Width           =   2999
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Gym:1"
         Height          =   255
         Left            =   -70920
         TabIndex        =   18
         Top             =   3720
         Width           =   975
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Poketopia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   22.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   64
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label lblstatus 
      BackColor       =   &H80000004&
      Caption         =   "Loading..."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   120
      Width           =   4815
   End
   Begin VB.Menu mnuKick 
      Caption         =   "&Kick"
      Visible         =   0   'False
      Begin VB.Menu mnuKickPlayer 
         Caption         =   "Kick"
      End
      Begin VB.Menu mnuDisconnectPlayer 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuBanPlayer 
         Caption         =   "Ban"
      End
      Begin VB.Menu mnuAdminPlayer 
         Caption         =   "Make Admin"
      End
      Begin VB.Menu mnuRemoveAdmin 
         Caption         =   "Remove Admin"
      End
      Begin VB.Menu mnuGivePoke 
         Caption         =   "Give Pokémon"
      End
      Begin VB.Menu mnuTakePoke 
         Caption         =   "Take Pokémon"
      End
      Begin VB.Menu mnuSpawn 
         Caption         =   "Spawn"
      End
      Begin VB.Menu mnuSetGym 
         Caption         =   "Set Gym"
      End
      Begin VB.Menu mnuTakeGym 
         Caption         =   "Take Gym"
      End
      Begin VB.Menu mnuDialog 
         Caption         =   "Dialog"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'function to make transparent'

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const G = (-20)
Private Const LWA_COLORKEY = &H1        'to trans'
Private Const LWA_ALPHA = &H2           'to semi trans'
Private Const WS_EX_LAYERED = &H80000
Public GlobalSeconds As Long

Private Sub Command10_Click()
Dim i As Long
Dim n As Long
Dim m As Long
Dim tM As Double
Dim x As Double
Dim y As Double
For i = 1 To TYPE_FAIRY
For n = 1 To TYPE_FAIRY
For m = 0 To TYPE_FAIRY
x = GetTypeEffect(i, n)
y = GetTypeEffect(i, m)
tM = x * y
PutVar App.Path & "\TypeEffects.ini", TypeToText(i), TypeToText(n) & "|" & TypeToText(m), Trim$(tM)
Next
Next
Next
MsgBox "DONE"
End Sub

Private Sub Command11_Click()
DeleteCrew Trim$(Text9.Text)
End Sub

Private Sub Command12_Click()

End Sub

Private Sub Command5_Click()
Dim i As Long
Dim str As String
For i = 1 To MAX_MOVES
str = i
Call PutVar(App.Path & "\Data\MoveNums.ini", "DATA", Trim$(PokemonMove(i).Name), str)
Next
MsgBox ("Moves saved")

End Sub

Private Sub Command6_Click()
Dim str As String
Dim i As Long
For i = 1 To MAX_MOVES
str = i
Call PutVar(App.Path & "\MoveNums.ini", "DATA", str, Trim$(PokemonMove(i).Name))
Next
MsgBox ("DONE")



End Sub

Private Sub Command7_Click()
frmPlayer.Show
End Sub

Private Sub Command8_Click()
If FindPlayer(Trim$(Text8.Text)) > 0 Then
GiveItem FindPlayer(Trim$(Text8.Text)), Val(Trim$(Text7.Text)), Val(Trim$(Text6.Text))
End If
End Sub

Sub ConvertMapTile(ByVal mapnum As Long, ByVal x As Long, ByVal y As Long)
Dim i As Long
Dim tempy As Long
 With map(mapnum).Tile(x, y)
       For i = MapLayer.GROUND To MapLayer.Fringe2
        If .Layer(i).Tileset > 0 Then
        If .Layer(i).y < 85 Then
        .Layer(i).Tileset = GetOldTilesetNewNum(.Layer(i).Tileset, 1)
        End If
        If .Layer(i).y > 84 And .Layer(i).y < 170 Then
         tempy = .Layer(i).y
           .Layer(i).y = tempy - 85
           .Layer(i).Tileset = GetOldTilesetNewNum(.Layer(i).Tileset, 2)
          
        End If
        If .Layer(i).y > 169 And .Layer(i).y <= 255 Then
        tempy = .Layer(i).y
           .Layer(i).y = tempy - 170
           .Layer(i).Tileset = GetOldTilesetNewNum(.Layer(i).Tileset, 3)
        End If
        End If
       Next
 End With
End Sub
Function GetOldTilesetNewNum(ByVal old As Long, ByVal newT As Long)
Dim oldString As String
Dim newString As String
oldString = old
newString = newT
GetOldTilesetNewNum = Val(GetVar(App.Path & "\Data\Maps.ini", oldString, newString))
End Function


Private Sub Command9_Click()
WhosDatPokemon = txtWhosAnswer.Text
WhosRewardItem = Val(txtWhosItem.Text)
WhosRewardItemVal = Val(txtWhosVal.Text)
isWhosOn = True
SendWhos 1, Val(txtWhosNum.Text)

End Sub

Private Sub trans(level As Integer)
    Dim msg As Long
    msg = GetWindowLong(Me.hWnd, G)
    msg = msg Or WS_EX_LAYERED
    SetWindowLong Me.hWnd, G, msg
    SetLayeredWindowAttributes Me.hWnd, vbBlack, level, LWA_ALPHA
End Sub



Private Sub Check1_Click()
Dim i As Long
If Check1.value = 1 Then
AdminOnly = True
Else
AdminOnly = False
End If
For i = 1 To MAX_PLAYERS
If IsPlaying(i) Then
If player(i).Access >= 1 Then
Else
CloseSocket i
End If
End If
Next
End Sub

Private Sub Command1_Click()
GlobalSeconds = 6
GlobalMsg "[Server]Global Spawn pokemon will appear in 5 sec.!", Yellow
tmrglobalspawn.Enabled = True
End Sub

Private Sub Command2_Click()
EXP35 = Not EXP35
'MsgBox ("Name:" & Trim$(PokemonMove(Val(Text1.Text)).Name) & "," & "Description:" & Trim$(PokemonMove(Val(Text1.Text)).Description) & "," & "PP:" & PokemonMove(Val(Text1.Text)).pp)
End Sub

Private Sub Command3_Click()
Dim i As Long
For i = 1 To MAX_PLAYERS
If IsPlaying(i) Then
  GiveItem i, Val(Text4.Text), Val(Text5.Text)
End If
Next
End Sub

Private Sub Command4_Click()
Dim i As Long
Dim a As Long
Dim str As String
For i = 1 To MAX_POKEMONS
Pokemon(i).Name = GetVar(App.Path & "\Data\Pokemon Data\" & i & ".ini", "DATA", "Name")
'type
Pokemon(i).Type = TextToType(GetVar(App.Path & "\Data\Pokemon Data\" & i & ".ini", "DATA", "Element"))
Pokemon(i).Type2 = TextToType(GetVar(App.Path & "\Data\Pokemon Data\" & i & ".ini", "DATA", "Element2"))
'stats
Pokemon(i).MaxHp = Val(GetVar(App.Path & "\Data\Pokemon Data\" & i & ".ini", "DATA", "HP"))
Pokemon(i).atk = Val(GetVar(App.Path & "\Data\Pokemon Data\" & i & ".ini", "DATA", "Attack"))
Pokemon(i).def = Val(GetVar(App.Path & "\Data\Pokemon Data\" & i & ".ini", "DATA", "Defense"))
Pokemon(i).spd = Val(GetVar(App.Path & "\Data\Pokemon Data\" & i & ".ini", "DATA", "Speed"))
Pokemon(i).spatk = Val(GetVar(App.Path & "\Data\Pokemon Data\" & i & ".ini", "DATA", "Sp.Atk"))
Pokemon(i).spdef = Val(GetVar(App.Path & "\Data\Pokemon Data\" & i & ".ini", "DATA", "Sp.Def"))
Pokemon(i).Stone = GetVar(App.Path & "\Data\Pokemon Data\" & i & ".ini", "DATA", "Stone")
Pokemon(i).Evolution = Val(GetVar(App.Path & "\Data\Pokemon Data\" & i & ".ini", "DATA", "Evolve"))
Pokemon(i).EvolvesTo = NameToNum(GetVar(App.Path & "\Data\Pokemon Data\" & i & ".ini", "DATA", "Next"))
'LOAD MOVES
For a = 1 To 30
Dim stra As String
stra = a
If Trim$(GetVar(App.Path & "\Data\Pokemon Data\" & i & ".ini", "DATA", "Learn" & stra)) <> "" Then
Pokemon(i).moves(a) = GetMoveID(Trim$(GetVar(App.Path & "\Data\Pokemon Data\" & i & ".ini", "DATA", "Learn" & stra)))
Pokemon(i).movesLV(a) = Val(GetVar(App.Path & "\Data\Pokemon Data\" & i & ".ini", "DATA", "Learn" & stra & "LV"))
End If
Next

str = i
Pokemon(i).BaseEXP = Val(GetVar(App.Path & "\Data\Pokemon Data\EXPData.ini", "DATA", str))
Call SendUpdatePokemonToAll(i)
Call SavePokemon(i)
Next
MsgBox ("Pokes are loaded")
End Sub

Private Sub HScroll1_Change()
Label1.Caption = "Gym:" & HScroll1.value
End Sub

Private Sub HScroll2_Change()
lblGSname.Caption = Trim$(Pokemon(HScroll2.value).Name)

End Sub

Private Sub HScroll3_Change()
lblGSlvl.Caption = "Lvl." & HScroll3.value
End Sub

Private Sub scrlPokemon_Change()
    lblPokemon.Caption = "Pokemon: " & scrlPokemon.value
End Sub

' ********************
' ** Winsock object **
' ********************
Private Sub Socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Call AcceptConnection(Index, requestID)
End Sub

Private Sub Socket_Accept(Index As Integer, SocketId As Integer)
    Call AcceptConnection(Index, SocketId)
End Sub

Private Sub Socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)

    If IsConnected(Index) Then
        Call IncomingData(Index, bytesTotal)
    End If

End Sub

Private Sub Socket_Close(Index As Integer)
    Call CloseSocket(Index)
End Sub

' ********************
Private Sub chkServerLog_Click()

    ' if its not 0, then its true
    If Not chkServerLog.value Then
        ServerLog = True
    End If

End Sub

Private Sub cmdExit_Click()
    Call DestroyServer
End Sub

Private Sub cmdReloadClasses_Click()
frmPlayer.Show
End Sub

Private Sub cmdReloadItems_Click()
Dim i As Long
    Call LoadItems
    Call TextAdd("All items reloaded.")
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            SendItems i
        End If
    Next
End Sub

Private Sub cmdReloadMaps_Click()
Dim i As Long
    Call LoadMaps
    Call TextAdd("All maps reloaded.")
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            PlayerWarp i, GetPlayerMap(i), GetPlayerX(i), GetPlayerY(i)
        End If
    Next
End Sub

Private Sub cmdReloadNPCs_Click()
Dim i As Long
    Call LoadNpcs
    Call TextAdd("All npcs reloaded.")
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            SendNpcs i
        End If
    Next
End Sub

Private Sub cmdReloadShops_Click()
Dim i As Long
    Call LoadShops
    Call TextAdd("All shops reloaded.")
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            SendShops i
        End If
    Next
End Sub

Private Sub cmdReloadSpells_Click()
Dim i As Long
    Call LoadSpells
    Call TextAdd("All spells reloaded.")
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            SendSpells i
        End If
    Next
End Sub

Private Sub cmdReloadResources_Click()
Dim i As Long
    Call LoadResources
    Call TextAdd("All Resources reloaded.")
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            SendResources i
        End If
    Next
End Sub

Private Sub cmdReloadAnimations_Click()
Dim i As Long
    Call LoadAnimations
    Call TextAdd("All Animations reloaded.")
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            SendAnimations i
        End If
    Next
End Sub

Private Sub cmdShutDown_Click()
    If isShuttingDown Then
        isShuttingDown = False
        cmdShutDown.Caption = "Shutdown"
        GlobalMsg "Shutdown canceled.", BrightBlue
    Else
        isShuttingDown = True
        cmdShutDown.Caption = "Cancel"
    End If
End Sub

Private Sub Form_Load()
    Call UsersOnline_Start
End Sub

Private Sub Form_Resize()

    If frmServer.WindowState = vbMinimized Then
        frmServer.Hide
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    Call DestroyServer
End Sub

Private Sub lvwInfo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    'When a ColumnHeader object is clicked, the ListView control is sorted by the subitems of that column.
    'Set the SortKey to the Index of the ColumnHeader - 1
    'Set Sorted to True to sort the list.
    If lvwInfo.SortOrder = lvwAscending Then
        lvwInfo.SortOrder = lvwDescending
    Else
        lvwInfo.SortOrder = lvwAscending
    End If

    lvwInfo.SortKey = ColumnHeader.Index - 1
    lvwInfo.Sorted = True
End Sub

Private Sub Text11_Change()

End Sub

Private Sub Text1_Change()

End Sub

Private Sub tmrglobalspawn_Timer()
GlobalSeconds = GlobalSeconds - 1
GlobalMsg GlobalSeconds & " seconds until Global Spawn!", Yellow
If GlobalSeconds = 0 Then
Dim i As Long
Call GlobalMsg("[GLOBAL SPAWN] Global spawn pokemon appeared: " & Trim$(Pokemon(HScroll2.value).Name) & "!", BrightRed)
For i = 1 To MAX_PLAYERS
If IsPlaying(i) Then
If TempPlayer(i).PokemonBattle.PokemonNumber = 0 Then
Call CustomPoke(i, HScroll2.value, HScroll3.value, chkisShiny.value)
End If
End If
Next
GlobalSeconds = 6
tmrglobalspawn.Enabled = False
End If
End Sub

Private Sub txtText_GotFocus()
    txtChat.SetFocus
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If LenB(Trim$(txtChat.Text)) > 0 Then
            Call GlobalMsg(txtChat.Text, White)
            Call TextAdd("Server: " & txtChat.Text)
            txtChat.Text = vbNullString
        End If

        KeyAscii = 0
    End If

End Sub

Sub UsersOnline_Start()
    Dim i As Integer

    For i = 1 To MAX_PLAYERS
        frmServer.lvwInfo.ListItems.Add (i)

        If i < 10 Then
            frmServer.lvwInfo.ListItems(i).Text = "00" & i
        ElseIf i < 100 Then
            frmServer.lvwInfo.ListItems(i).Text = "0" & i
        Else
            frmServer.lvwInfo.ListItems(i).Text = i
        End If

        frmServer.lvwInfo.ListItems(i).SubItems(1) = vbNullString
        frmServer.lvwInfo.ListItems(i).SubItems(2) = vbNullString
        frmServer.lvwInfo.ListItems(i).SubItems(3) = vbNullString
    Next

End Sub

Private Sub lvwInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        PopupMenu mnuKick
    End If

End Sub

Private Sub mnuKickPlayer_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        Call AlertMsg(FindPlayer(Name), "You have been kicked by the server owner!")
    End If

End Sub

Private Sub mnuDialog_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        SendDialog FindPlayer(Name), Text3.Text, Val(Text2.Text)
    End If

End Sub

Private Sub mnuSpawn_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        SpawnPlayer FindPlayer(Name)
    End If

End Sub

Private Sub mnuSetGym_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        player(FindPlayer(Name)).Bedages(HScroll1.value) = GYM_DEFEATED
    End If

End Sub

Private Sub mnuTakeGym_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        player(FindPlayer(Name)).Bedages(HScroll1.value) = GYM_UNDEFEATED
    End If

End Sub

Sub mnuDisconnectPlayer_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        CloseSocket (FindPlayer(Name))
    End If

End Sub

Sub mnuBanPlayer_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        Call ServerBanIndex(FindPlayer(Name))
    End If

End Sub

Sub mnuAdminPlayer_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        Call SetPlayerAccess(FindPlayer(Name), 4)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "You have been granted administrator access.", BrightCyan)
    End If

End Sub

Sub mnuRemoveAdmin_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        Call SetPlayerAccess(FindPlayer(Name), 0)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "You have had your administrator access revoked.", BrightRed)
    End If

End Sub

Private Sub mnuGivePoke_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)
    
    If Not Name = "Not Playing" Then
        GivePokemon FindPlayer(Name), scrlPokemon.value, Val(txtlvl.Text), chkPokeShiny.value, 1, Val(txtNat.Text)
    End If
End Sub

Private Sub mnuTakePoke_Click()
Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)
    
    If Not Name = "Not Playing" Then
        TakePokemon FindPlayer(Name), scrlPokemon.value
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lmsg As Long
    lmsg = x / Screen.TwipsPerPixelX

    Select Case lmsg
        Case WM_LBUTTONDBLCLK
            frmServer.WindowState = vbNormal
            frmServer.Show
            txtText.SelStart = Len(txtText.Text)
    End Select
End Sub

Public Function TextToType(ByVal Text As String) As Byte
Select Case Text
Case "None"
TextToType = TYPE_NONE
Case "Normal"
TextToType = TYPE_NORMAL
Case "Bug"
TextToType = TYPE_BUG
Case "Dark"
TextToType = TYPE_DARK
Case "Dragon"
TextToType = TYPE_DRAGON
Case "Electric"
TextToType = TYPE_ELECTRIC
Case "Fairy"
TextToType = TYPE_FAIRY
Case "Fighting"
TextToType = TYPE_FIGHTING
Case "Fire"
TextToType = TYPE_FIRE
Case "Flying"
TextToType = TYPE_FLYING
Case "Ghost"
TextToType = TYPE_GHOST
Case "Grass"
TextToType = TYPE_GRASS
Case "Ground"
TextToType = TYPE_GROUND
Case "Ice"
TextToType = TYPE_ICE
Case "Poison"
TextToType = TYPE_POISON
Case "Psychic"
TextToType = TYPE_PSYCHIC
Case "Rock"
TextToType = TYPE_ROCK
Case "Steel"
TextToType = TYPE_STEEL
Case "Water"
TextToType = TYPE_WATER
End Select
End Function


Public Function NumToName(ByVal Num As Long) As String
If Num < 1 Or Num > MAX_POKEMONS Then Exit Function
Dim str As String
str = Num
NumToName = GetVar(App.Path & "\Data\Pokemon Data\Nums_Names.ini", "DATA", str)
End Function

Public Function NameToNum(ByVal Name As String) As Long
If Name = vbNullString Or Name = "" Then Exit Function
NameToNum = Val(GetVar(App.Path & "\Data\Pokemon Data\Names_Nums.ini", "DATA", Name))
End Function
