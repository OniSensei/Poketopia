VERSION 5.00
Object = "{50347EDF-F3EF-4392-AFDD-71AE67A3A978}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LVbutton.ocx"
Begin VB.Form frmStorage 
   BackColor       =   &H00312920&
   BorderStyle     =   0  'None
   Caption         =   "Storage"
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8145
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00554C42&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   240
      ScaleHeight     =   3705
      ScaleWidth      =   4785
      TabIndex        =   12
      Top             =   600
      Width           =   4815
      Begin lvButton.lvButtons_H lvButtons_H4 
         Height          =   375
         Left            =   3360
         TabIndex        =   23
         Top             =   2880
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Withdraw"
         CapAlign        =   2
         BackStyle       =   5
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
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H lvButtons_H5 
         Height          =   375
         Left            =   3360
         TabIndex        =   24
         Top             =   3240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Release"
         CapAlign        =   2
         BackStyle       =   5
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
         cBack           =   16777215
      End
      Begin VB.Label lblStat 
         BackStyle       =   0  'Transparent
         Caption         =   "HP: 0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   22
         Top             =   1560
         Width           =   4095
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         Caption         =   "Name: None."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   21
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblLevel 
         BackStyle       =   0  'Transparent
         Caption         =   "Level: 0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   20
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label lblNature 
         BackStyle       =   0  'Transparent
         Caption         =   "Nature: None."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   19
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label lblStat 
         BackStyle       =   0  'Transparent
         Caption         =   "Atk: 0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   1920
         Width           =   4095
      End
      Begin VB.Label lblStat 
         BackStyle       =   0  'Transparent
         Caption         =   "Def: 0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   2280
         Width           =   4095
      End
      Begin VB.Label lblStat 
         BackStyle       =   0  'Transparent
         Caption         =   "Speed: 0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   2640
         Width           =   4095
      End
      Begin VB.Label lblStat 
         BackStyle       =   0  'Transparent
         Caption         =   "Sp.Atk: 0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   3000
         Width           =   4095
      End
      Begin VB.Label lblStat 
         BackStyle       =   0  'Transparent
         Caption         =   "Sp.Def: 0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   14
         Top             =   3360
         Width           =   4095
      End
      Begin VB.Label lblStat 
         BackStyle       =   0  'Transparent
         Caption         =   "HP: 0"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   5
         Left            =   -600
         TabIndex        =   13
         Top             =   2400
         Width           =   4095
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl imgPicture 
         Height          =   1575
         Left            =   240
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   2778
         Attr            =   1536
         Effects         =   "frmStorage.frx":0000
      End
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   4440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Deposit"
      CapAlign        =   2
      BackStyle       =   5
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
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   4440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Close"
      CapAlign        =   2
      BackStyle       =   5
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
      cBack           =   16777215
   End
   Begin VB.PictureBox picDeposit 
      BackColor       =   &H00554C42&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   5160
      ScaleHeight     =   4935
      ScaleWidth      =   2895
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   2895
      Begin lvButton.lvButtons_H CmdDp 
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   5
         Top             =   3720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Caption         =   "Deposit"
         CapAlign        =   2
         BackStyle       =   5
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
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H CmdDp 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Caption         =   "Deposit"
         CapAlign        =   2
         BackStyle       =   5
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
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H CmdDp 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Caption         =   "Deposit"
         CapAlign        =   2
         BackStyle       =   5
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
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H CmdDp 
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   1920
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Caption         =   "Deposit"
         CapAlign        =   2
         BackStyle       =   5
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
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H CmdDp 
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   9
         Top             =   2520
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Caption         =   "Deposit"
         CapAlign        =   2
         BackStyle       =   5
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
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H CmdDp 
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   10
         Top             =   3120
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Caption         =   "Deposit"
         CapAlign        =   2
         BackStyle       =   5
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
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H lvButtons_H6 
         Height          =   375
         Left            =   720
         TabIndex        =   11
         Top             =   4320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "Close"
         CapAlign        =   2
         BackStyle       =   5
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
         cBack           =   16777215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Deposit:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.ListBox lstStorage 
      Appearance      =   0  'Flat
      BackColor       =   &H00554C42&
      ForeColor       =   &H00FFFFFF&
      Height          =   4905
      Left            =   5160
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
End
Attribute VB_Name = "frmStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'function to make transparent'

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
  ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const G = (-20)
Private Const LWA_COLORKEY = &H1        'to trans'
Private Const LWA_ALPHA = &H2           'to semi trans'
Private Const WS_EX_LAYERED = &H80000

Private xm As Integer
Private ym As Integer

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  xm = X
  ym = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    Me.Left = Me.Left + X - xm
    Me.Top = Me.Top + Y - ym
  End If
End Sub

Private Sub cmdClose_Click()
picDeposit.Visible = False
End Sub





Private Sub cmdWithdraw_Click(Index As Integer)

End Sub

Private Sub cmdRemove_Click(Index As Integer)
If MsgBox("Are you sure you want to remove this pokemon?", vbYesNo, "Remove Pokemon") = vbYes Then
Call SendRemoveStoragePokemon(storagenum)
End If
End Sub

Private Sub Command1_Click()
isInStorage = False
Unload Me
End Sub

Private Sub Command2_Click()
Dim i As Long
Dim Name As String
For i = 1 To 6
If PokemonInstance(i).PokemonNumber <= 0 Then
Name = "Empty."
Else
Name = Pokemon(PokemonInstance(i).PokemonNumber).Name
End If
CmdDp(i).Caption = Trim$(Name)
Next
picDeposit.Visible = True
End Sub

Private Sub Command3_Click()
If StorageInstance(storagenum).PokemonNumber >= 1 Then
SendWithdrawPokemon (storagenum)
End If
End Sub

Private Sub CmdDp_Click(Index As Integer)
SendDepositPokemon Index
picDeposit.Visible = False
End Sub

Private Sub Form_Load()
isInStorage = True
initStorage
storagenum = 1
LoadPokemon (storagenum)
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Private Sub lstStorage_Click()
storagenum = lstStorage.ListIndex + 1
LoadPokemon (storagenum)
End Sub

Sub LoadPokemon(ByVal pn As Long)
Dim pokenum As Integer
pokenum = StorageInstance(pn).PokemonNumber
If pokenum > 0 Then
If StorageInstance(pn).isShiny = YES Then
imgPicture.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\Shiny\" & pokenum & ".gif")
Else
imgPicture.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\" & pokenum & ".gif")
End If

lblName.Caption = "Name: " & Pokemon(pokenum).Name
lblLevel.Caption = "Level: " & StorageInstance(pn).Level
If StorageInstance(pn).nature = 0 Then
lblNature.Caption = "Nature: None."
Else
lblNature.Caption = "Nature: " & Trim$(nature(StorageInstance(pn).nature).Name)
End If
lblStat(0).Caption = "Atk: " & StorageInstance(pn).ATK
lblStat(1).Caption = "Def: " & StorageInstance(pn).DEF
lblStat(2).Caption = "Speed: " & StorageInstance(pn).SPD
lblStat(3).Caption = "Sp.Atk: " & StorageInstance(pn).SPATK
lblStat(4).Caption = "Sp.Def: " & StorageInstance(pn).SPDEF
lblStat(6).Caption = "HP: " & StorageInstance(pn).MaxHp
Else
imgPicture.Picture = Nothing
lblName.Caption = "Name: None. "
lblLevel.Caption = "Level: 0"
lblNature.Caption = "Nature: None."
lblStat(0).Caption = "Atk: 0"
lblStat(1).Caption = "Def: 0"
lblStat(2).Caption = "Speed: 0"
lblStat(3).Caption = "Sp.Atk: 0"
lblStat(4).Caption = "Sp.Def: 0"
lblStat(6).Caption = "HP: 0"
End If


End Sub

Private Sub lvButton1_Click()
If StorageInstance(storagenum).PokemonNumber >= 1 Then
SendWithdrawPokemon (storagenum)
End If
End Sub

Private Sub lvButton3_Click()
If MsgBox("Are you sure you want to remove this pokemon?", vbYesNo, "Remove Pokemon") = vbYes Then
Call SendRemoveStoragePokemon(storagenum)
End If
End Sub

Private Sub lvButton4_Click()
Dim i As Long
Dim Name As String
For i = 1 To 6
If PokemonInstance(i).PokemonNumber <= 0 Then
Name = "Empty."
Else
Name = Pokemon(PokemonInstance(i).PokemonNumber).Name
End If
CmdDp(i).Caption = Trim$(Name)
Next
picDeposit.Visible = True
End Sub

Private Sub lvButton5_Click()
isInStorage = False
Unload Me
End Sub

Private Sub lvButton2_Click()

End Sub

Private Sub lvButton6_Click()
picDeposit.Visible = False
End Sub

Private Sub lvButton7_Click()

End Sub

Private Sub lvButtons_H1_Click()
isInStorage = False
Unload Me
End Sub

Private Sub lvButtons_H2_Click()
Dim i As Long
Dim Name As String
For i = 1 To 6
If PokemonInstance(i).PokemonNumber <= 0 Then
Name = "Empty."
Else
Name = Pokemon(PokemonInstance(i).PokemonNumber).Name
End If
CmdDp(i).Caption = Trim$(Name)
Next
picDeposit.Visible = True
End Sub

Private Sub lvButtons_H3_Click()

End Sub

Private Sub lvButtons_H4_Click()
If StorageInstance(storagenum).PokemonNumber >= 1 Then
SendWithdrawPokemon (storagenum)
End If
End Sub

Private Sub lvButtons_H5_Click()
If MsgBox("Are you sure you want to remove this pokemon?", vbYesNo, "Remove Pokemon") = vbYes Then
Call SendRemoveStoragePokemon(storagenum)
End If
End Sub

Private Sub lvButtons_H6_Click()
picDeposit.Visible = False
End Sub

