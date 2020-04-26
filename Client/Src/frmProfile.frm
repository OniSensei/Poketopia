VERSION 5.00
Object = "{50347EDF-F3EF-4392-AFDD-71AE67A3A978}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LVbutton.ocx"
Begin VB.Form frmProfile 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Player profile"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Text            =   "Image Link"
      Top             =   1440
      Width           =   1815
   End
   Begin lvButton.lvButtons_H lvButtons_H9 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   960
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Caption         =   "Set profile picture"
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
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   2640
      Width           =   2055
      _ExtentX        =   3625
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
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgProfile 
      Height          =   1335
      Left            =   360
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2355
      Attr            =   513
      Effects         =   "frmProfile.frx":0000
   End
   Begin VB.Label lblPOKE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Profile"
      BeginProperty Font 
         Name            =   "Eurostar"
         Size            =   15.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmProfile"
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
    Me.backColor = vbBlack
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


Private Sub lvButtons_H1_Click()
Unload Me
End Sub

Private Sub lvButtons_H9_Click()
    If InStr(Text1.text, "http") Then
        imgProfile.Picture = LoadPictureGDIplus(Text1.text)

        SendRequest 0, 0, Trim$(Text1.text), "PPIC"
    Else
        MsgBox "You must enter a URL.", vbExclamation
    End If
End Sub
