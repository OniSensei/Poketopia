VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LVbutton.ocx"
Begin VB.Form frmBank 
   BackColor       =   &H000000FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bank"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBank.frx":0000
   ScaleHeight     =   2175
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H cmdMove 
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Caption         =   "Withdraw 500"
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
      cBhover         =   16777215
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin lvButton.lvButtons_H cmdDep 
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Caption         =   "Deposit 500"
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
      cBhover         =   16777215
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      Caption         =   "Close"
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
      cBhover         =   16777215
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin VB.Label lblSPC 
      BackStyle       =   0  'Transparent
      Caption         =   "Stored PC:0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label lblCPC 
      BackStyle       =   0  'Transparent
      Caption         =   "PokeCoins:0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub




Private Sub cmdClose_Click()
isInBank = False
Unload Me
End Sub

Private Sub cmdDep_Click()
SendDepositPC

End Sub

Private Sub cmdMove_Click()
SendWithdrawPC
End Sub


Private Sub Form_Load()
isInBank = True
LoadBank
End Sub

Sub LoadBank()
lblSPC.Caption = "Stored PC:" & Player(MyIndex).StoredPC
Dim i As Long
Dim pcslot As Long
For i = 1 To MAX_INV
If GetPlayerInvItemNum(Index, i) = 1 Then
pcslot = i
Exit For
End If
Next
If pcslot = 0 Then
lblCPC.Caption = "PokeCoins:0"
Else
lblCPC.Caption = "PokeCoins:" & GetPlayerInvItemValue(MyIndex, pcslot)
End If
End Sub

