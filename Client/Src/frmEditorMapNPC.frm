VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmEditorMapNPC 
   Caption         =   "NPC"
   ClientHeight    =   4815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   5475
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   4320
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6376
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmEditorMapNPC.frx":0000
   End
   Begin VB.Label Label2 
      Caption         =   "Script:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmEditorMapNPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentNpc As Long

Sub load(ByVal cnpc As Long)
CurrentNpc = cnpc
End Sub

Private Sub Command1_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
SendEditMapNpc CurrentNpc, RichTextBox1.text
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

