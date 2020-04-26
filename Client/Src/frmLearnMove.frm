VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LVbutton.ocx"
Begin VB.Form frmLearnMove 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Learn move"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H btnMove 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Move1"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cBhover         =   14737632
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin lvButton.lvButtons_H btnMove 
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Move1"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cBhover         =   14737632
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin lvButton.lvButtons_H btnMove 
      Height          =   375
      Index           =   3
      Left            =   2760
      TabIndex        =   2
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Move1"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cBhover         =   14737632
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin lvButton.lvButtons_H btnMove 
      Height          =   375
      Index           =   4
      Left            =   4080
      TabIndex        =   3
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Move1"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cBhover         =   14737632
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin lvButton.lvButtons_H btnDONT 
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   960
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      Caption         =   "Don't learn"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cBhover         =   14737632
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Bulbasaur want to learn Razor Leaf.Choose a move to replace."
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmLearnMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
  ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Dim pSlot As Long
Dim pMove As Long
Dim pMoveSlot As Long
Dim exitFormNow As Boolean
Private Sub btnDONT_Click()
exitFormNow = True
frmMainGame.Enabled = True
Unload Me
End Sub

Private Sub btnMove_Click(Index As Integer)
Call SendLearnMove(pSlot, Index, pMove)
exitFormNow = True
frmMainGame.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If exitFormNow = False Then
Dim aa As String
aa = MsgBox("Are you sure you want to exit?", vbYesNo)
If aa = vbYes Then
Else
Cancel = True
End If
End If
End Sub

Public Sub LoadMoveAndPoke(ByVal slot As Long, ByVal move As Long)
Dim i As Long
For i = 1 To 4
If PokemonInstance(slot).moves(i).number > 0 Then
btnMove(i).Caption = Trim$(PokemonMove(PokemonInstance(slot).moves(i).number).Name)
Else
btnMove(i).Caption = "None."
End If
Next
Label1.Caption = Trim$(Pokemon(PokemonInstance(slot).PokemonNumber).Name) & " wants to learn " & Trim$(PokemonMove(move).Name) & ".Choose move to replace."
pSlot = slot
pMove = move
End Sub


