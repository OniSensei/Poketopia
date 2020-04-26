VERSION 5.00
Begin VB.Form frmSetNpc 
   Caption         =   "NPC"
   ClientHeight    =   2670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   2670
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Move Down"
      Height          =   495
      Left            =   4440
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Move Right"
      Height          =   495
      Left            =   4920
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Move Left"
      Height          =   495
      Left            =   3840
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Move Up"
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Edit NPCS"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   3495
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   135
      Left            =   240
      Max             =   3
      TabIndex        =   2
      Top             =   1320
      Value           =   1
      Width           =   3375
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   240
      Max             =   255
      TabIndex        =   0
      Top             =   600
      Value           =   1
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Direction : Down"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "None"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmSetNpc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nx As Long
Dim ny As Long
Dim fs As Long
Dim firsttime As Boolean
Dim isnpc As Boolean

Private Sub Command1_Click()
Dim i As Long
If HScroll1.Value > 0 Then
map.NPC(fs) = HScroll1.Value
map.Tile(nx, ny).Type = TILE_TYPE_NPCSPAWN
map.Tile(nx, ny).data1 = fs
map.Tile(nx, ny).data2 = HScroll2.Value
map.Tile(nx, ny).data3 = 0
Call SendMap
Unload Me
Exit Sub
Else
map.NPC(fs) = 0
map.Tile(nx, ny).Type = 0
map.Tile(nx, ny).data1 = 0
map.Tile(nx, ny).data2 = 0
map.Tile(nx, ny).data3 = 0
SendEditMapNpc fs, ""
Call SendMap
Unload Me
Exit Sub

End If

End Sub

Private Sub Command2_Click()
If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If

    SendRequestEditNpc
End Sub

Private Sub Command3_Click()
If map.NPC(fs) > 0 Then
If ny - 1 >= 0 Then
If map.Tile(nx, ny - 1).Type = 0 Or map.Tile(nx, ny - 1).Type = TILE_TYPE_NPCSPAWN Then
   map.Tile(nx, ny - 1).Type = TILE_TYPE_NPCSPAWN
   map.Tile(nx, ny - 1).data1 = fs
   map.Tile(nx, ny - 1).data2 = HScroll2.Value
   map.Tile(nx, ny - 1).data3 = 0
   map.Tile(nx, ny).Type = 0
   map.Tile(nx, ny).data1 = 0
   map.Tile(nx, ny).data2 = 0
   map.Tile(nx, ny).data3 = 0
   nx = nx
   ny = ny - 1
End If
End If
End If
Call SendMap
End Sub

Private Sub Command4_Click()
If map.NPC(fs) > 0 Then
If nx + 1 >= 0 Then
If map.Tile(nx - 1, ny).Type = 0 Or map.Tile(nx - 1, ny).Type = TILE_TYPE_NPCSPAWN Then
   map.Tile(nx - 1, ny).Type = TILE_TYPE_NPCSPAWN
   map.Tile(nx - 1, ny).data1 = fs
   map.Tile(nx - 1, ny).data2 = HScroll2.Value
   map.Tile(nx - 1, ny).data3 = 0
   map.Tile(nx, ny).Type = 0
   map.Tile(nx, ny).data1 = 0
   map.Tile(nx, ny).data2 = 0
   map.Tile(nx, ny).data3 = 0
   nx = nx - 1
   ny = ny
   End If
End If
End If
Call SendMap
End Sub

Private Sub Command5_Click()
If map.NPC(fs) > 0 Then
If nx + 1 <= map.MaxX Then
If map.Tile(nx + 1, ny).Type = 0 Or map.Tile(nx + 1, ny).Type = TILE_TYPE_NPCSPAWN Then
   map.Tile(nx + 1, ny).Type = TILE_TYPE_NPCSPAWN
   map.Tile(nx + 1, ny).data1 = fs
   map.Tile(nx + 1, ny).data2 = HScroll2.Value
   map.Tile(nx + 1, ny).data3 = 0
   map.Tile(nx, ny).Type = 0
   map.Tile(nx, ny).data1 = 0
   map.Tile(nx, ny).data2 = 0
   map.Tile(nx, ny).data3 = 0
   nx = nx + 1
   ny = ny
End If
End If
End If
Call SendMap
End Sub

Private Sub Command6_Click()
If map.NPC(fs) > 0 Then
If ny + 1 <= map.MaxY Then
If map.Tile(nx, ny + 1).Type = 0 Or map.Tile(nx, ny + 1).Type = TILE_TYPE_NPCSPAWN Then
   map.Tile(nx, ny + 1).Type = TILE_TYPE_NPCSPAWN
   map.Tile(nx, ny + 1).data1 = fs
   map.Tile(nx, ny + 1).data2 = HScroll2.Value
   map.Tile(nx, ny + 1).data3 = 0
   map.Tile(nx, ny).Type = 0
   map.Tile(nx, ny).data1 = 0
   map.Tile(nx, ny).data2 = 0
   map.Tile(nx, ny).data3 = 0
   nx = nx
   ny = ny + 1
End If
End If
End If
Call SendMap
End Sub

Private Sub Form_Load()
Label1.Caption = NPC(1).Name
HScroll1.Value = 1
End Sub

 Sub LoadPos(ByVal X As Long, ByVal Y As Long)
 nx = X
 ny = Y
 Dim i As Long
 For i = 1 To MAX_MAP_NPCS
 If MapNpc(i).X = nx And MapNpc(i).Y = ny Then
 HScroll1.Value = MapNpc(i).num
 Label1.Caption = Trim$(NPC(MapNpc(i).num).Name)
 HScroll2.Value = MapNpc(i).Dir
 fs = i
 isnpc = True
 Exit For
 End If
 Next
 End Sub

Private Sub HScroll1_Change()
If HScroll1.Value = 0 Then
Label1.Caption = "None"
Else
Label1.Caption = NPC(HScroll1).Name
End If

Dim i As Long
Dim freeslot As Long
For i = 1 To 30
If map.NPC(i) = 0 And isnpc = False Then
freeslot = i
Exit For
End If
Next
If freeslot = 0 And isnpc = False Then
MsgBox "There is no free slot on map!"
Exit Sub
End If
If isnpc = False Then
fs = freeslot
End If
If firsttime = False Then
 For i = 1 To MAX_MAP_NPCS
 If MapNpc(i).X = nx And MapNpc(i).Y = ny Then
 HScroll1.Value = MapNpc(i).num
 'Label1.Caption = Trim$(NPC(MapNpc(i).num).Name)
 HScroll2.Value = MapNpc(i).Dir
 fs = i
 Exit For
 End If
 Next
 firsttime = True
End If

End Sub

Private Sub HScroll2_Change()
Select Case HScroll2.Value
Case DIR_DOWN
Label2.Caption = "Direction: Down"
Case DIR_UP
Label2.Caption = "Direction: Up"
Case DIR_RIGHT
Label2.Caption = "Direction: Right"
Case DIR_LEFT
Label2.Caption = "Direction: Left"
End Select

End Sub

