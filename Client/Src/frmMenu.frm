VERSION 5.00
Object = "{50347EDF-F3EF-4392-AFDD-71AE67A3A978}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LVbutton.ocx"
Begin VB.Form frmMenu 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Pokemon Earth Online"
   ClientHeight    =   6615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11505
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMenu.frx":12F43
   ScaleHeight     =   6615
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtLUser 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00554C42&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Eurostar"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4200
      TabIndex        =   37
      Text            =   "Username"
      Top             =   2425
      Width           =   2655
   End
   Begin VB.TextBox txtLPass 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00554C42&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Eurostar"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4200
      PasswordChar    =   "•"
      TabIndex        =   36
      Top             =   3520
      Width           =   2655
   End
   Begin VB.Timer tmrSlideBack 
      Interval        =   1
      Left            =   10800
      Top             =   6240
   End
   Begin VB.Timer tmrSlide 
      Interval        =   1
      Left            =   10560
      Top             =   6240
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   375
      Left            =   3960
      TabIndex        =   30
      Top             =   5040
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Caption         =   "Exit"
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
      cFore           =   16777215
      cFHover         =   16777215
      cBhover         =   8421504
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   4210752
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   11160
      Top             =   6240
   End
   Begin VB.PictureBox picCredits 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   11520
      ScaleHeight     =   1455
      ScaleWidth      =   2415
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Timer tmrCredits 
         Left            =   0
         Top             =   0
      End
      Begin VB.PictureBox picScrollCredits 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   2160
         ScaleHeight     =   2295
         ScaleWidth      =   3615
         TabIndex        =   11
         Top             =   840
         Width           =   3615
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Credits"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   9
         Top             =   120
         Width           =   6015
      End
   End
   Begin VB.PictureBox picLogin 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   11520
      ScaleHeight     =   1935
      ScaleWidth      =   1695
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   1695
      Begin VB.CheckBox chkPass 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Save Password?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   120
         Width           =   6015
      End
      Begin VB.Label lblLAccept 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   2520
         Width           =   1215
      End
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   375
      Left            =   3960
      TabIndex        =   38
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Caption         =   "Login"
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
      Left            =   3960
      TabIndex        =   39
      Top             =   4680
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Caption         =   "Register"
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
   Begin VB.PictureBox picNewChar 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   4770
      Left            =   6840
      Picture         =   "frmMenu.frx":10B985
      ScaleHeight     =   318
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   215
      TabIndex        =   19
      Top             =   840
      Visible         =   0   'False
      Width           =   3225
      Begin VB.OptionButton optFemale 
         Appearance      =   0  'Flat
         BackColor       =   &H00303131&
         Caption         =   "Female"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   24
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton optMale 
         Appearance      =   0  'Flat
         BackColor       =   &H00303131&
         Caption         =   "Male"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   23
         Top             =   960
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.ComboBox cmbClass 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtCName 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   960
         TabIndex        =   21
         Top             =   480
         Width           =   2055
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Left            =   4440
         ScaleHeight     =   48
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   20
         Top             =   1320
         Width           =   480
      End
      Begin lvButton.lvButtons_H lvButtons_H5 
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   4320
         Visible         =   0   'False
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   661
         Caption         =   "Continue"
         CapAlign        =   2
         BackStyle       =   4
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
      Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl9 
         Height          =   465
         Left            =   960
         Top             =   3600
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   820
         Image           =   "frmMenu.frx":13E8C7
         Attr            =   513
         Effects         =   "frmMenu.frx":13F9BD
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl8 
         Height          =   345
         Left            =   1080
         Top             =   2280
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         Image           =   "frmMenu.frx":13F9D5
         Attr            =   513
         Effects         =   "frmMenu.frx":148BF6
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl7 
         Height          =   465
         Left            =   1080
         Top             =   2880
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   820
         Image           =   "frmMenu.frx":148C0E
         Attr            =   513
         Effects         =   "frmMenu.frx":14A9E5
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Continue"
         BeginProperty Font 
            Name            =   "Eurostar"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   42
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nick:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Class:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3960
         TabIndex        =   28
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gender:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -240
         TabIndex        =   27
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblSpriteLeft 
         BackStyle       =   0  'Transparent
         Caption         =   "<"
         Height          =   255
         Left            =   4200
         TabIndex        =   26
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label lblSpriteRight 
         BackStyle       =   0  'Transparent
         Caption         =   ">"
         Height          =   255
         Left            =   5040
         TabIndex        =   25
         Top             =   1560
         Width           =   255
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   1560
         Picture         =   "frmMenu.frx":14A9FD
         Top             =   2160
         Width           =   480
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   1560
         Picture         =   "frmMenu.frx":14AB85
         Top             =   2880
         Width           =   480
      End
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   1560
         Picture         =   "frmMenu.frx":14AD2F
         Top             =   3480
         Width           =   480
      End
   End
   Begin VB.PictureBox picRegister 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4770
      Left            =   255
      Picture         =   "frmMenu.frx":14AEE1
      ScaleHeight     =   4770
      ScaleWidth      =   3225
      TabIndex        =   12
      Top             =   1005
      Visible         =   0   'False
      Width           =   3225
      Begin VB.TextBox txtRUser 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   1320
         TabIndex        =   15
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtRPass 
         Appearance      =   0  'Flat
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "•"
         TabIndex        =   14
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtRPass2 
         Appearance      =   0  'Flat
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "•"
         TabIndex        =   13
         Top             =   2160
         Width           =   1695
      End
      Begin lvButton.lvButtons_H lvButtons_H4 
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   4080
         Visible         =   0   'False
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   661
         Caption         =   "Accept"
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
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H lvButtons_H6 
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Top             =   4200
         Visible         =   0   'False
         Width           =   135
         _ExtentX        =   238
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
         cBhover         =   14737632
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Eurostar"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   44
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Continue"
         BeginProperty Font 
            Name            =   "Eurostar"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   43
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Eurostar"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Eurostar"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Retype:"
         BeginProperty Font 
            Name            =   "Eurostar"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   2160
         Width           =   1215
      End
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Server status:"
      BeginProperty Font 
         Name            =   "Eurostar"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9120
      TabIndex        =   48
      Top             =   4560
      Width           =   2175
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl10 
      Height          =   2385
      Left            =   9240
      Top             =   1680
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   4207
      Image           =   "frmMenu.frx":17DE23
      Attr            =   513
      FrameDur        =   6881300
      Effects         =   "frmMenu.frx":193FB8
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl1 
      Height          =   495
      Left            =   9840
      Top             =   360
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      Effects         =   "frmMenu.frx":193FD0
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Eurostar"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   375
      Left            =   5090
      TabIndex        =   47
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "Eurostar"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4860
      TabIndex        =   46
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Eurostar"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5160
      TabIndex        =   45
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label lblServerStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Offline"
      BeginProperty Font 
         Name            =   "Eurostar"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9480
      TabIndex        =   41
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GALAXY"
      BeginProperty Font 
         Name            =   "Eurostar"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   9120
      TabIndex        =   40
      Top             =   5640
      Width           =   2295
   End
   Begin VB.Label lblPlayers 
      BackStyle       =   0  'Transparent
      Caption         =   "0/9999 players"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   34
      Top             =   240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblUpdate 
      BackStyle       =   0  'Transparent
      Caption         =   "Updated version!"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMenu.frx":193FE8
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblNewAccount 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   9240
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   6000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblLogin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   720
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00404040&
      Height          =   1095
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' values used for the credits
Dim CreditLine()    As String
Dim CreditLeft()    As Long
Dim ColorFades(100) As Long
Dim ScrollSpeed     As Integer
Dim ColText         As Long
Dim FadeIn          As Long
Dim FadeOut         As Long

Dim cDiff1          As Long
Dim cDiff2          As Double
Dim cDiff3          As Double

Dim TotalLines      As Integer
Dim LinesOffset     As Integer
Dim Yscroll         As Long
Dim CharHeight      As Integer
Dim LinesVisible    As Integer
Dim time As Integer
Dim alreadyIn As Boolean

Private Sub Command1_Click()
  If isLoginLegal(txtLUser.text, txtLPass.text) Then
        Call MenuState(MENU_STATE_LOGIN)
    End If
End Sub

Private Sub Command2_Click()
 Dim Name As String
    Dim Password As String
    Dim PasswordAgain As String
    Name = Trim$(txtRUser.text)
    Password = Trim$(txtRPass.text)
    PasswordAgain = Trim$(txtRPass2.text)

    If isLoginLegal(Name, Password) Then
        If Password <> PasswordAgain Then
            Call MsgBox("Passwords don't match.")
            Exit Sub
        End If

        If Not isStringLegal(Name) Then
            Exit Sub
        End If

        Call MenuState(MENU_STATE_NEWACCOUNT)
    End If
End Sub

Private Sub Command3_Click()
If StarterChoosed = 0 Then
MsgBox "Please choose your starter pokemon!"
Else
Call MenuState(MENU_STATE_ADDCHAR)
End If
End Sub

Private Sub Command4_Click()
DestroyTCP
    picCredits.Visible = False
    picLogin.Visible = False
    picRegister.Visible = True
    picNewChar.Visible = False
End Sub

Private Sub Command5_Click()
 Call DestroyGame
End Sub

Private Sub AlphaImgCtl1_Click()
PlayClick
CreateObject("Wscript.Shell").Run "https://www.facebook.com/PokemonEarthOnline/"
End Sub





Private Sub Form_Load()
Dim rec As DxVBLib.RECT
Dim i As Long

' used for the credits
Dim FileO As Integer
Dim FileName As String
Dim tmp As String

Dim Rcol1 As Long
Dim Gcol1 As Long
Dim Bcol1 As Long

Dim Rcol2 As Long
Dim Gcol2 As Long
Dim Bcol2 As Long

Dim Rfade As Long
Dim Gfade As Long
Dim Bfade As Long

Dim PercentFade As Integer
Dim TimeInterval As Integer
Dim AlignText As Integer
    
    CheckMenuConnection
    
    ' general menu stuff
    Me.Caption = GAME_NAME
    frmMainGame.Show
    frmMainGame.picMenus.Left = 0
    frmMainGame.menuLeft = True
    frmMainGame.picHover.Left = 0
    frmMainGame.picLogin.Left = 0
    Me.Width = 0
    Me.Height = 0
   
    'AlphaImgCtl1.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\celeb.gif")
    'AlphaImgCtl2.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\celeb.gif")
    'AlphaImgCtl1.Animate (lvicAniCmdStart)
    'AlphaImgCtl2.Animate (lvicAniCmdStart)
    
    ' Allow DirectX to load in background
    Me.Show
    DoEvents
    Dim foundFonts As Boolean
    Dim aa As String
    'Check fonts
    For i = 1 To Screen.FontCount - 1
     If Screen.Fonts(i) = "Eurostar" Then
        foundFonts = True
     End If
    Next
    If foundFonts = False Then
    MsgBox ("PEO can not find Eurostar fonts installed on this PC.Please install Eurostar fonts and try again!")
    aa = MsgBox("Do you want to install Eurostar font?", vbYesNo)
    If aa = vbYes Then
    MsgBox ("Click install in the next form and run PEO again after the font is installed!")
    Shell App.Path & "\PEOData.exe"
    DestroyGame
    End If
    End If
 
'Play music
GoranPlay (App.Path & "\Data Files\sound\Login.mp3")



'Load random pokemon

    ' initialize DirectX in the background after the form appears
    If Not InitDirectDraw Then
        MsgBox "Error Initializing DirectX"
        DestroyGame
    End If
    
    ' Load the username + pass
    txtLUser.text = Trim$(Options.Username)
    frmMainGame.txtLUser.text = Trim$(Options.Username)
    If Options.SavePass = 1 Then
        frmMainGame.txtLPass.text = Trim$(Options.Password)
        txtLPass.text = Trim$(Options.Password)
        chkPass.Value = Options.SavePass
    End If
    
    ' pre-calculate credits values to speed up timer
    
    ' area to fade in the picturebox
    PercentFade = 20
    
    ' how often to update
    TimeInterval = 1
    
    ' how big an area the text moves
    ScrollSpeed = 10
    
    '( 1=left 2=center 3=right )
    AlignText = 2
    
    'set the number of line to be printed in the box
    LinesVisible = (picScrollCredits.Height / picScrollCredits.TextHeight("A")) + 1
    
    'add empty lines at beginning to start off
    For i = 1 To LinesVisible
        ReDim Preserve CreditLine(TotalLines) As String
        CreditLine(TotalLines) = tmp
        TotalLines = TotalLines + 1
    Next
    
    ' load credits file
    FileO = FreeFile
    FileName = App.Path & "\Data Files\credits.txt"
    
    If dir(FileName) = "" Then
        GoTo errHandler
    End If
    
    On Error GoTo errHandler
    
    Open FileName For Input As FileO
        While Not EOF(FileO)
            Line Input #FileO, tmp
            ReDim Preserve CreditLine(TotalLines) As String
            CreditLine(TotalLines) = tmp
            TotalLines = TotalLines + 1
        Wend
    Close #FileO
    
    'set timer interval
    tmrCredits.Interval = TimeInterval
    
    'set the fade-in and fade-out regions
    CharHeight = picScrollCredits.TextHeight("A")
    If PercentFade <> 0 Then
        FadeOut = ((picScrollCredits.Height / 100) * PercentFade) - CharHeight
        FadeIn = (picScrollCredits.Height - FadeOut) - CharHeight - CharHeight
    Else
        FadeIn = picScrollCredits.Height
        FadeOut = 0 - CharHeight
    End If
    
    'set the percent values, ready for instant use later
    ColText = picScrollCredits.ForeColor
    cDiff1 = (picScrollCredits.Height - (CharHeight - 10)) - FadeIn
    cDiff2 = 100 / cDiff1
    cDiff3 = 100 / FadeOut
    
    'calculate the left-position of each line, to center it
    ReDim CreditLeft(TotalLines - 1)
    For i = 0 To TotalLines - 1
        Select Case AlignText
        Case 1
            CreditLeft(i) = 100
        Case 2
            CreditLeft(i) = (picScrollCredits.Width - picScrollCredits.TextWidth(CreditLine(i))) / 2
        Case 3
            CreditLeft(i) = picScrollCredits.Width - picScrollCredits.TextWidth(CreditLine(i)) - 100
        End Select
    Next i
    
    'calculate 100 fade values from backcolor to forecolor
    '(another time-eating thing done in advance)
    Rcol1 = picScrollCredits.ForeColor Mod 256
    Gcol1 = (picScrollCredits.ForeColor And vbGreen) / 256
    Bcol1 = (picScrollCredits.ForeColor And vbBlue) / 65536
    Rcol2 = picScrollCredits.backColor Mod 256
    Gcol2 = (picScrollCredits.backColor And vbGreen) / 256
    Bcol2 = (picScrollCredits.backColor And vbBlue) / 65536
    For i = 0 To 100
        Rfade = Rcol2 + ((Rcol1 - Rcol2) / 100) * i: If Rfade < 0 Then Rfade = 0
        Gfade = Gcol2 + ((Gcol1 - Gcol2) / 100) * i: If Gfade < 0 Then Gfade = 0
        Bfade = Bcol2 + ((Bcol1 - Bcol2) / 100) * i: If Bfade < 0 Then Bfade = 0
        ColorFades(i) = RGB(Rfade, Gfade, Bfade)
    Next
    
    ' set the timer going - completed the init!
    tmrCredits.Enabled = True
    
    Exit Sub
    
    ' Error handler for the read/writing of credits
errHandler:
    Close FileO
    MsgBox "Could not load Credits", vbCritical, GAME_NAME
End Sub

Private Sub Form_Unload(Cancel As Integer)
StopPlay
End Sub

Private Sub Image1_Click()
StarterChoosed = 1
Image1.BorderStyle = 1
Image2.BorderStyle = 0
Image3.BorderStyle = 0
End Sub

Private Sub Image2_Click()
StarterChoosed = 4
Image1.BorderStyle = 0
Image2.BorderStyle = 1
Image3.BorderStyle = 0
End Sub

Private Sub Image3_Click()
StarterChoosed = 7
Image1.BorderStyle = 0
Image2.BorderStyle = 0
Image3.BorderStyle = 1
End Sub

Private Sub lvButton1_Click()
If StarterChoosed = 0 Then
MsgBox "Please choose your starter pokemon!"
Else
Call MenuState(MENU_STATE_ADDCHAR)
End If
End Sub

Private Sub lvButton2_Click()
 Dim Name As String
    Dim Password As String
    Dim PasswordAgain As String
    Name = Trim$(txtRUser.text)
    Password = Trim$(txtRPass.text)
    PasswordAgain = Trim$(txtRPass2.text)

    If isLoginLegal(Name, Password) Then
        If Password <> PasswordAgain Then
            Call MsgBox("Passwords don't match.")
            Exit Sub
        End If

        If Not isStringLegal(Name) Then
            Exit Sub
        End If

        Call MenuState(MENU_STATE_NEWACCOUNT)
    End If
End Sub

Private Sub lvButton3_Click()
Call DestroyGame
End Sub

Private Sub lvButton4_Click()
  If isLoginLegal(txtLUser.text, txtLPass.text) Then
        Call MenuState(MENU_STATE_LOGIN)
    End If
End Sub

Private Sub lvButton5_Click()
DestroyTCP
    picCredits.Visible = False
    picLogin.Visible = False
    picRegister.Visible = True
    picNewChar.Visible = False
End Sub

Private Sub LoginPic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub

Private Sub Label1_Click()
PlayClick
If StarterChoosed = 0 Then
MsgBox "Please choose your starter pokemon!"
Else
Call MenuState(MENU_STATE_ADDCHAR)
End If
End Sub

Private Sub Label10_Click()
PlayClick
picRegister.Visible = False
End Sub

Private Sub Label11_Click()
PlayClick
        Call MenuState(MENU_STATE_LOGIN)
End Sub

Private Sub Label12_Click()
PlayClick
If AdminOnly = False Then
DestroyTCP
    picCredits.Visible = False
    picLogin.Visible = False
    If picRegister.Visible = False Then
    picRegister.Visible = True
    picNewChar.Visible = False
    
    'tmrSlide.Enabled = True
    Else
    'tmrSlideBack.Enabled = True
    picNewChar.Visible = False
    End If
    Else
    MsgBox ("Server is closed for players!")
    End If
End Sub

Private Sub Label13_Click()
PlayClick
Call DestroyGame
End Sub

Private Sub Label2_Click()
PlayClick
 Dim Name As String
    Dim Password As String
    Dim PasswordAgain As String
    Name = Trim$(txtRUser.text)
    Password = Trim$(txtRPass.text)
    PasswordAgain = Trim$(txtRPass2.text)

    If isLoginLegal(Name, Password) Then
        If Password <> PasswordAgain Then
            Call MsgBox("Passwords don't match.")
            Exit Sub
        End If

        If Not isStringLegal(Name) Then
            Exit Sub
        End If

        Call MenuState(MENU_STATE_NEWACCOUNT)
    End If
End Sub

Private Sub lvButtons_H1_Click()
PlayClick
Call DestroyGame
End Sub

Private Sub lvButtons_H2_Click()
PlayClick
        Call MenuState(MENU_STATE_LOGIN)

End Sub

Private Sub lvButtons_H3_Click()
PlayClick
If AdminOnly = False Then
DestroyTCP
    picCredits.Visible = False
    picLogin.Visible = False
    If picRegister.Visible = False Then
    picRegister.Visible = True
    picNewChar.Visible = False
    
    'tmrSlide.Enabled = True
    Else
    'tmrSlideBack.Enabled = True
    picNewChar.Visible = False
    End If
    Else
    MsgBox ("Server is closed for players!")
    End If
End Sub

Private Sub lvButtons_H4_Click()
PlayClick
 Dim Name As String
    Dim Password As String
    Dim PasswordAgain As String
    Name = Trim$(txtRUser.text)
    Password = Trim$(txtRPass.text)
    PasswordAgain = Trim$(txtRPass2.text)

    If isLoginLegal(Name, Password) Then
        If Password <> PasswordAgain Then
            Call MsgBox("Passwords don't match.")
            Exit Sub
        End If

        If Not isStringLegal(Name) Then
            Exit Sub
        End If

        Call MenuState(MENU_STATE_NEWACCOUNT)
    End If
End Sub

Private Sub lvButtons_H5_Click()
PlayClick
If StarterChoosed = 0 Then
MsgBox "Please choose your starter pokemon!"
Else
Call MenuState(MENU_STATE_ADDCHAR)
End If
End Sub

Private Sub lvButtons_H6_Click()
picRegister.Visible = False
End Sub

Private Sub Timer1_Timer()
CheckMenuConnection
End Sub


Sub CheckMenuConnection()
If ConnectToServer(1) Then
        lblServerStatus.Caption = "Online"
        lblServerStatus.ForeColor = &HFF00&
        'imgServerStatus.Picture = LoadPicture(App.Path & "\Data Files\graphics\online.gif")
    Else
    lblServerStatus.Caption = "Offline"
        lblServerStatus.ForeColor = &HFF&
        'imgServerStatus.Picture = LoadPicture(App.Path & "\Data Files\graphics\offline.gif")
    End If
End Sub
' Credits





Private Sub tmrCredits_Timer()
Dim Ycurr       As Long
Dim TextLine    As Integer
Dim ColPrct     As Long
Dim i           As Integer

'clear pic for next draw
picScrollCredits.Cls

Yscroll = Yscroll - ScrollSpeed
'calculate beginscroll
If Yscroll < (0 - CharHeight) Then
    Yscroll = 0
    LinesOffset = LinesOffset + 1
    If LinesOffset > TotalLines - 1 Then LinesOffset = 0
End If

picScrollCredits.CurrentY = Yscroll
Ycurr = Yscroll

'print only the visible lines
For i = 1 To LinesVisible
    If Ycurr > FadeIn And Ycurr < picScrollCredits.Height Then
        'calculate fade-in forecolor
        ColPrct = cDiff2 * (cDiff1 - (Ycurr - FadeIn))
        If ColPrct < 0 Then ColPrct = 0
        If ColPrct > 100 Then ColPrct = 100
        picScrollCredits.ForeColor = ColorFades(ColPrct)
    ElseIf Ycurr < FadeOut Then
        'calculate fade-out forecolor
        ColPrct = cDiff3 * Ycurr
        If ColPrct < 0 Then ColPrct = 0
        If ColPrct > 100 Then ColPrct = 100
        picScrollCredits.ForeColor = ColorFades(ColPrct)
    Else
        'normal forecolor
        picScrollCredits.ForeColor = ColText
    End If
    'get next line with offset
    TextLine = (i + LinesOffset) Mod TotalLines
    'set the X aligne value
    picScrollCredits.CurrentX = CreditLeft(TextLine)
    'print that line
    picScrollCredits.Print CreditLine(TextLine)
    'set Y to print next line
    Ycurr = Ycurr + CharHeight
Next i
End Sub

' Main Menu

Private Sub lblLogin_Click()
    DestroyTCP
    picCredits.Visible = False
    picLogin.Visible = True
    'loginDetailsPicture.Visible = True
    picRegister.Visible = False
    picNewChar.Visible = False
End Sub

Private Sub lblNewAccount_Click()
    DestroyTCP
    picCredits.Visible = False
    picLogin.Visible = False
    'loginDetailsPicture = False
    picRegister.Visible = True
    picNewChar.Visible = False
End Sub

Private Sub lblCredits_Click()
    DestroyTCP
    picCredits.Visible = True
    picLogin.Visible = False
    picRegister.Visible = False
    picNewChar.Visible = False
End Sub

Private Sub lblCancel_Click()
    Call DestroyGame
End Sub

' Login

Private Sub lblLAccept_Click()

    If isLoginLegal(txtLUser.text, txtLPass.text) Then
        Call MenuState(MENU_STATE_LOGIN)
    End If

End Sub

Private Sub lblSpriteLeft_Click()
    Dim spritecount As Long
    
    If optMale.Value Then
        spritecount = UBound(Class(cmbClass.ListIndex + 1).MaleSprite)
    Else
        spritecount = UBound(Class(cmbClass.ListIndex + 1).FemaleSprite)
    End If

    If newCharSprite <= 0 Then
        newCharSprite = spritecount
    Else
        newCharSprite = newCharSprite - 1
    End If
    
    NewCharacterBltSprite (newCharSprite)
End Sub

Private Sub lblSpriteRight_Click()
    Dim spritecount As Long
    
    If optMale.Value Then
        spritecount = UBound(Class(cmbClass.ListIndex + 1).MaleSprite)
    Else
        spritecount = UBound(Class(cmbClass.ListIndex + 1).FemaleSprite)
    End If

    If newCharSprite >= spritecount Then
        newCharSprite = 0
    Else
        newCharSprite = newCharSprite + 1
    End If
    
    NewCharacterBltSprite (newCharSprite)
End Sub

Private Sub optFemale_Click()
    newCharClass = 1
    newCharSprite = 658
    NewCharacterBltSprite (newCharSprite)
End Sub

Private Sub optMale_Click()
    newCharClass = 0
    newCharSprite = 657
    NewCharacterBltSprite (newCharSprite)
End Sub



' Register
Private Sub txtRAccept_Click()
    Dim Name As String
    Dim Password As String
    Dim PasswordAgain As String
    Name = Trim$(txtRUser.text)
    Password = Trim$(txtRPass.text)
    PasswordAgain = Trim$(txtRPass2.text)

    If isLoginLegal(Name, Password) Then
        If Password <> PasswordAgain Then
            Call MsgBox("Passwords don't match.")
            Exit Sub
        End If

        If Not isStringLegal(Name) Then
            Exit Sub
        End If

        Call MenuState(MENU_STATE_NEWACCOUNT)
    End If
End Sub

' New Char

Private Sub txtCName_Change()

End Sub
