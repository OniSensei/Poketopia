VERSION 5.00
Object = "{50347EDF-F3EF-4392-AFDD-71AE67A3A978}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LVbutton.ocx"
Begin VB.Form frmBag 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bag"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDesc 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   6960
      TabIndex        =   4
      Top             =   3600
      Width           =   3495
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   2160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "Remove Item"
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
      cBack           =   14737632
   End
   Begin VB.ListBox lstItems 
      Appearance      =   0  'Flat
      Height          =   4710
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   3960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "Unequip items"
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
   Begin lvButton.lvButtons_H lvButtons_H3 
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   4560
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      Caption         =   "Close"
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
   Begin LaVolpeAlphaImg.AlphaImgCtl imgicon 
      Height          =   1695
      Left            =   3480
      Top             =   360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   2990
      Effects         =   "frmBag.frx":0000
   End
End
Attribute VB_Name = "frmBag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'function to make transparent'

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const G = (-20)
Private Const LWA_COLORKEY = &H1        'to trans'
Private Const LWA_ALPHA = &H2           'to semi trans'
Private Const WS_EX_LAYERED = &H80000

Private Sub Form_Activate()
    Me.BackColor = vbBlack
   If Options.FormTransparency = YES Then
    trans 215
    End If
End Sub

Private Sub trans(Level As Integer)
    Dim Msg As Long
    Msg = GetWindowLong(Me.hwnd, G)
    Msg = Msg Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, G, Msg
    SetLayeredWindowAttributes Me.hwnd, vbBlack, Level, LWA_ALPHA
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Dim invnum As Long
invnum = lstItems.ListIndex + 1
Call SendDropItem(invnum, 1)
LoadInv
End If

End Sub

Private Sub Command2_Click()
Dim i As Long
Dim eslot As Long
For i = 1 To 4
eslot = GetPlayerEquipment(MyIndex, i)
If eslot > 0 Then
SendUnequip i
End If
Next
LoadInv
End Sub

Private Sub Command3_Click()
CanMoveNow = True
Unload Me

End Sub

Private Sub AlphaImgCtl1_Click()

End Sub

Private Sub Form_Load()
LoadInv
Dim invnum As Long
invnum = 1
If FileExist("Data Files\itemicons\" & GetPlayerInvItemNum(MyIndex, invnum) & ".png") Then
imgicon.Picture = LoadPictureGDIplus(App.Path & "\Data Files\itemicons\" & GetPlayerInvItemNum(MyIndex, invnum) & ".png")
Else
If FileExist("Data Files\itemicons\" & GetPlayerInvItemNum(MyIndex, invnum) & ".gif") Then
imgicon.Picture = LoadPictureGDIplus(App.Path & "\Data Files\itemicons\" & GetPlayerInvItemNum(MyIndex, invnum) & ".gif")
imgicon.Animate (lvicAniCmdStart)
End If
End If
End Sub

Private Sub lstItems_Click()
Dim invnum As Long
invnum = lstItems.ListIndex + 1
If FileExist("Data Files\itemicons\" & GetPlayerInvItemNum(MyIndex, invnum) & ".png") Then
imgicon.Picture = LoadPictureGDIplus(App.Path & "\Data Files\itemicons\" & GetPlayerInvItemNum(MyIndex, invnum) & ".png")
Exit Sub
Else
If FileExist("Data Files\itemicons\" & GetPlayerInvItemNum(MyIndex, invnum) & ".gif") Then
imgicon.Picture = LoadPictureGDIplus(App.Path & "\Data Files\itemicons\" & GetPlayerInvItemNum(MyIndex, invnum) & ".gif")
imgicon.Animate (lvicAniCmdStart)
Exit Sub
End If
End If
imgicon.Picture = Nothing
End Sub

Private Sub lstItems_DblClick()
Dim invnum As Long
invnum = lstItems.ListIndex + 1
Call SendUseItem(invnum)
LoadInv



End Sub

Public Sub LoadInv()
Dim i As Long
Dim itemnum As Long
Dim itemvalue As Long
lstItems.Clear
For i = 1 To MAX_INV
itemnum = GetPlayerInvItemNum(MyIndex, i)
itemvalue = GetPlayerInvItemValue(MyIndex, i)
If itemvalue = 0 Then itemvalue = 1
If itemnum <= MAX_ITEMS Then
If itemnum = 0 Then
lstItems.AddItem ("Empty")
Else
lstItems.AddItem (Item(itemnum).Name & " x" & itemvalue)
End If

End If
Next
End Sub

Private Sub lstItems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Dim invnum As Long
invnum = lstItems.ListIndex + 1
Call SendDropItem(invnum, 1)
LoadInv
End If
End Sub

Private Sub lvButtons_H1_Click()
Dim invnum As Long
invnum = lstItems.ListIndex + 1
Call SendDropItem(invnum, 1)
LoadInv
End Sub

Private Sub lvButtons_H2_Click()
Dim i As Long
Dim eslot As Long
For i = 1 To 4
eslot = GetPlayerEquipment(MyIndex, i)
If eslot > 0 Then
SendUnequip i
End If
Next
LoadInv
End Sub

Private Sub lvButtons_H3_Click()
CanMoveNow = True
Unload Me

End Sub
