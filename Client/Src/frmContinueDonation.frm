VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LVbutton.ocx"
Begin VB.Form frmContinueDonation 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Continue donation"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picPaypal 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   120
      ScaleHeight     =   4935
      ScaleWidth      =   5055
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox txtPalMail 
         Height          =   285
         Left            =   3000
         TabIndex        =   11
         Text            =   "Paypal Email"
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txtPalID 
         Height          =   285
         Left            =   3000
         TabIndex        =   10
         Text            =   "Payment transaction ID"
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txtPalName 
         Height          =   285
         Left            =   3000
         TabIndex        =   8
         Text            =   "Paypal name"
         Top             =   1560
         Width           =   1935
      End
      Begin lvButton.lvButtons_H btnDONT 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         Caption         =   "5$ - 500 DC"
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
      Begin lvButton.lvButtons_H lvButtons_H1 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         Caption         =   "10$ - 1100 DC"
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
      Begin lvButton.lvButtons_H lvButtons_H2 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   2880
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         Caption         =   "50$ - 5500 DC"
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
      Begin lvButton.lvButtons_H lvButtons_H3 
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   2400
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         Caption         =   "25$ - 2750 DC"
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
      Begin lvButton.lvButtons_H lvButtons_H4 
         Height          =   375
         Left            =   3000
         TabIndex        =   12
         Top             =   2640
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "Continue"
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
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Make sure you filled every box before you continue!"
         BeginProperty Font 
            Name            =   "Eurostar"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   3000
         TabIndex        =   13
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "After payment is completed fill information below."
         BeginProperty Font 
            Name            =   "Eurostar"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   2760
         TabIndex        =   9
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "You will be redirected to checkout.Once the payment is done copy transaction ID from payment details."
         BeginProperty Font 
            Name            =   "Eurostar"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   360
         TabIndex        =   7
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Choose option"
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
         Left            =   360
         TabIndex        =   6
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PayPal"
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
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.PictureBox picKarma 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   120
      ScaleHeight     =   4935
      ScaleWidth      =   5055
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox txtKarmaKey 
         Height          =   285
         Left            =   480
         TabIndex        =   20
         Text            =   "Karma Koin key"
         Top             =   1200
         Width           =   3975
      End
      Begin lvButton.lvButtons_H lvButtons_H6 
         Height          =   375
         Left            =   1560
         TabIndex        =   21
         Top             =   1920
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "Continue"
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
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Karma Koin"
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
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Do not use fake or used karma koin keys!"
         BeginProperty Font 
            Name            =   "Eurostar"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   600
         TabIndex        =   22
         Top             =   2640
         Width           =   3975
      End
   End
   Begin VB.PictureBox picPaysafe 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   120
      ScaleHeight     =   4935
      ScaleWidth      =   5055
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox txtPaysafe 
         Height          =   285
         Left            =   480
         TabIndex        =   15
         Text            =   "Paysafe key"
         Top             =   1200
         Width           =   3975
      End
      Begin lvButton.lvButtons_H lvButtons_H5 
         Height          =   375
         Left            =   1560
         TabIndex        =   17
         Top             =   1920
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "Continue"
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
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Do not use fake or used paysafe keys!"
         BeginProperty Font 
            Name            =   "Eurostar"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   600
         TabIndex        =   18
         Top             =   2640
         Width           =   3975
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Paysafe"
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
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmContinueDonation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
  ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
 

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const G = (-20)
Private Const LWA_COLORKEY = &H1        'to trans'
Private Const LWA_ALPHA = &H2           'to semi trans'
Private Const WS_EX_LAYERED = &H80000

Private Sub btnDONT_Click()
CreateObject("Wscript.Shell").Run "https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=YK6XUWHDFDM3C"
End Sub

Private Sub Form_Activate()
    Me.BackColor = vbBlack
    If Options.FormTransparency = YES Then
    trans 215
    End If
End Sub

Private Sub trans(Level As Integer)
    Dim Msg As Long
    Msg = GetWindowLong(Me.hWnd, G)
    Msg = Msg Or WS_EX_LAYERED
    SetWindowLong Me.hWnd, G, Msg
    SetLayeredWindowAttributes Me.hWnd, vbBlack, Level, LWA_ALPHA
End Sub
Private Sub Form_Load()
Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Private Sub lvButtons_H1_Click()
CreateObject("Wscript.Shell").Run "https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=TP75G5SPRETP4"
End Sub

Private Sub lvButtons_H2_Click()
CreateObject("Wscript.Shell").Run "https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=P3XJ6JWQJSQVL"
End Sub

Private Sub lvButtons_H3_Click()
CreateObject("Wscript.Shell").Run "https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=FDN36UDYHLVRL"

End Sub

Private Sub lvButtons_H4_Click()
SendDonate "PAYPAL", txtPalID.text, txtPalName.text, txtPalMail.text
MsgBox ("Your donation has been set.It is now in pending status.Contact our staff if you have any problems!")
Unload Me
End Sub

Private Sub lvButtons_H5_Click()
SendDonate "PAYSAFE", txtPaysafe.text, "", ""
MsgBox ("Your donation has been set.It is now in pending status.Contact our staff if you have any problems!")
Unload Me
End Sub

Private Sub lvButtons_H6_Click()
SendDonate "KARMA KOIN", txtKarmaKey.text, "", ""
MsgBox ("Your donation has been set.It is now in pending status.Contact our staff if you have any problems!")
Unload Me
End Sub
