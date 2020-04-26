VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LVbutton.ocx"
Begin VB.Form frmShop 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Shop"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5505
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   5760
      Width           =   5415
      _ExtentX        =   9551
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
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Price: 0 PokeCoins"
      Top             =   5280
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sell"
      Height          =   255
      Left            =   6960
      TabIndex        =   4
      Top             =   5040
      Width           =   1935
   End
   Begin VB.ListBox lstMyItems 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   4710
      Left            =   2760
      TabIndex        =   3
      Top             =   0
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Cost: 0 PokeCoins"
      Top             =   5280
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buy"
      Height          =   255
      Left            =   6840
      TabIndex        =   1
      Top             =   4080
      Width           =   1935
   End
   Begin VB.ListBox lstItems 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   4710
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   4800
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      Caption         =   "Sell"
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
      Left            =   120
      TabIndex        =   8
      Top             =   4800
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      Caption         =   "Buy"
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
End
Attribute VB_Name = "frmShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Shop(InShop).TradeItem(lstItems.ListIndex + 1).Item > 0 And Shop(InShop).TradeItem(lstItems.ListIndex + 1).Item <= MAX_ITEMS Then
BuyItem lstItems.ListIndex + 1
End If

End Sub

Private Sub Command2_Click()
Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CCloseShop
    Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    frmShop.Visible = False
    InShop = 0
    ShopAction = 0
End Sub

Private Sub Command3_Click()
If GetPlayerInvItemNum(MyIndex, lstMyItems.ListIndex + 1) > 0 Then
SellItem lstMyItems.ListIndex + 1
End If

End Sub

Private Sub Form_Load()
loadItems
loadMyItems
End Sub

Sub loadItems()
Dim i As Long
lstItems.Clear
For i = 1 To MAX_TRADES
If Shop(InShop).TradeItem(i).Item > 0 Then
lstItems.AddItem (Trim$(Item(Shop(InShop).TradeItem(i).Item).Name))
Else
lstItems.AddItem ("Empty")
End If




Next
End Sub

Sub loadMyItems()
Dim i As Long
Dim itemnum As Long
Dim itemvalue As Long
lstMyItems.Clear
For i = 1 To MAX_INV
itemnum = GetPlayerInvItemNum(MyIndex, i)
itemvalue = GetPlayerInvItemValue(MyIndex, i)
If itemvalue = 0 Then itemvalue = 1
If itemnum <= MAX_ITEMS Then
If itemnum = 0 Then
lstMyItems.AddItem ("Empty")
Else
lstMyItems.AddItem (Item(itemnum).Name & " x" & itemvalue)
End If

End If
Next
End Sub

Private Sub lstItems_Click()
If Shop(InShop).TradeItem(lstItems.ListIndex + 1).Item > 0 And Shop(InShop).TradeItem(lstItems.ListIndex + 1).Item <= MAX_ITEMS Then
Text1.text = "Cost:" & Shop(InShop).TradeItem(lstItems.ListIndex + 1).CostValue & " " & Trim$(Item(Shop(InShop).TradeItem(lstItems.ListIndex + 1).CostItem).Name)
Else
Text1.text = "Cost:0 PokeCoins"
End If

End Sub

Private Sub lstMyItems_Click()
If GetPlayerInvItemNum(MyIndex, lstMyItems.ListIndex + 1) > 0 Then
Text2.text = "Price:" & Item(GetPlayerInvItemNum(MyIndex, lstMyItems.ListIndex + 1)).Price & "PokeCoins"
Else
Text2.text = "Price:0 PokeCoins"
End If
End Sub

Private Sub lvButtons_H1_Click()
Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CCloseShop
    Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    frmShop.Visible = False
    InShop = 0
    ShopAction = 0
End Sub

Private Sub lvButtons_H2_Click()
On Error Resume Next
If lstMyItems.text <> "" Then
If GetPlayerInvItemNum(MyIndex, lstMyItems.ListIndex + 1) > 0 Then
SellItem lstMyItems.ListIndex + 1
End If
End If
End Sub

Private Sub lvButtons_H3_Click()
If Shop(InShop).TradeItem(lstItems.ListIndex + 1).Item > 0 And Shop(InShop).TradeItem(lstItems.ListIndex + 1).Item <= MAX_ITEMS Then
BuyItem lstItems.ListIndex + 1
End If
End Sub
