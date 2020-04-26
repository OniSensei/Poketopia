VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{50347EDF-F3EF-4392-AFDD-71AE67A3A978}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LVbutton.ocx"
Begin VB.Form frmBattle 
   BackColor       =   &H00312920&
   BorderStyle     =   0  'None
   Caption         =   "Battle"
   ClientHeight    =   6855
   ClientLeft      =   8250
   ClientTop       =   4605
   ClientWidth     =   12675
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBattleInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00312920&
      FillColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   0
      ScaleHeight     =   6345
      ScaleWidth      =   9345
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   9375
      Begin lvButton.lvButtons_H lvButtons_H2 
         Height          =   375
         Left            =   3720
         TabIndex        =   13
         Top             =   4920
         Width           =   2415
         _ExtentX        =   4260
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
         cFore           =   16777215
         cFHover         =   4210752
         cBhover         =   12632319
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   8421631
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00554C42&
         BorderStyle     =   0  'None
         Height          =   6375
         Left            =   3600
         ScaleHeight     =   6375
         ScaleWidth      =   2655
         TabIndex        =   23
         Top             =   0
         Width           =   2655
         Begin LaVolpeAlphaImg.AlphaImgCtl imgPokeball 
            Height          =   2055
            Left            =   240
            Top             =   600
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   3625
            Effects         =   "frmBattle.frx":0000
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Pokemon:"
            BeginProperty Font 
               Name            =   "Arial"
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
            TabIndex        =   31
            Top             =   3120
            Width           =   1335
         End
         Begin VB.Label lblInfoPoke 
            BackStyle       =   0  'Transparent
            Caption         =   "PokemonName"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1200
            TabIndex        =   30
            Top             =   3120
            Width           =   2055
         End
         Begin VB.Label label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Chance:"
            BeginProperty Font 
               Name            =   "Arial"
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
            TabIndex        =   29
            Top             =   3600
            Width           =   1335
         End
         Begin VB.Label lblInfoChance 
            BackStyle       =   0  'Transparent
            Caption         =   "1 of 10000"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   960
            TabIndex        =   28
            Top             =   3600
            Width           =   1335
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "PokeCoins Won:"
            BeginProperty Font 
               Name            =   "Arial"
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
            TabIndex        =   27
            Top             =   4440
            Width           =   1575
         End
         Begin VB.Label lblInfoPC 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1680
            TabIndex        =   26
            Top             =   4440
            Width           =   1335
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Exp Gained:"
            BeginProperty Font 
               Name            =   "Arial"
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
            TabIndex        =   25
            Top             =   4080
            Width           =   1335
         End
         Begin VB.Label lblInfoExp 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFC0&
            Height          =   255
            Left            =   1320
            TabIndex        =   24
            Top             =   4080
            Width           =   1335
         End
      End
   End
   Begin VB.PictureBox picSwitch 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   17760
      ScaleHeight     =   2505
      ScaleWidth      =   3705
      TabIndex        =   12
      Top             =   4680
      Visible         =   0   'False
      Width           =   3735
      Begin lvButton.lvButtons_H lvButtons_H1 
         Height          =   375
         Left            =   0
         TabIndex        =   18
         Top             =   2040
         Width           =   3735
         _ExtentX        =   6588
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
         cFore           =   16777215
         cFHover         =   4210752
         cBhover         =   8421631
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   255
      End
   End
   Begin VB.TextBox txtBattleLog 
      Appearance      =   0  'Flat
      Height          =   4560
      Left            =   13680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   600
      Width           =   3315
   End
   Begin RichTextLib.RichTextBox txtBtlLog 
      Height          =   6360
      Left            =   9360
      TabIndex        =   10
      Top             =   480
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   11218
      _Version        =   393217
      BackColor       =   12632319
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmBattle.frx":0018
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdMove 
      Height          =   495
      Index           =   4
      Left            =   6360
      TabIndex        =   14
      Top             =   6240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "Move1"
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
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H cmdMove 
      Height          =   495
      Index           =   1
      Left            =   1680
      TabIndex        =   15
      Top             =   6240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "Move1"
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
   Begin lvButton.lvButtons_H cmdMove 
      Height          =   495
      Index           =   2
      Left            =   3240
      TabIndex        =   16
      Top             =   6240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "Move1"
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
   Begin lvButton.lvButtons_H cmdMove 
      Height          =   495
      Index           =   3
      Left            =   4800
      TabIndex        =   17
      Top             =   6240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "Move1"
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
   Begin lvButton.lvButtons_H lvButtons_H3 
      Height          =   495
      Left            =   8100
      TabIndex        =   19
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "Bag"
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00312920&
      BorderStyle     =   0  'None
      Height          =   5745
      Left            =   0
      ScaleHeight     =   5745
      ScaleWidth      =   9375
      TabIndex        =   0
      Top             =   480
      Width           =   9375
      Begin lvButton.lvButtons_H lvButtons_H4 
         Height          =   495
         Left            =   8100
         TabIndex        =   21
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "Run"
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
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   5400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "Stuck"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6
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
         Height          =   255
         Left            =   8160
         TabIndex        =   33
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         Caption         =   "Auto close"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6
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
      Begin VB.ListBox listBag 
         Appearance      =   0  'Flat
         BackColor       =   &H00554C42&
         ForeColor       =   &H00FFFFFF&
         Height          =   4125
         Left            =   3360
         TabIndex        =   20
         Top             =   720
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Shape shapeHpMine 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   90
         Left            =   1200
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00554C42&
         BackStyle       =   1  'Opaque
         Height          =   90
         Left            =   1200
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Shape shapeHPEnemy 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   90
         Left            =   5520
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00554C42&
         BackStyle       =   1  'Opaque
         Height          =   90
         Left            =   5520
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label lblEXP 
         Alignment       =   2  'Center
         BackColor       =   &H00554C42&
         Caption         =   "0/0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   22
         Top             =   5520
         Width           =   3255
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl EnemyImg 
         Height          =   720
         Index           =   0
         Left            =   8325
         Top             =   3000
         Visible         =   0   'False
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Frame           =   0
         Attr            =   1536
         Effects         =   "frmBattle.frx":0094
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl vsimage 
         Height          =   375
         Left            =   8480
         Top             =   2760
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Effects         =   "frmBattle.frx":00AC
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl maleimg 
         Height          =   1335
         Left            =   8160
         Top             =   4440
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   2355
         Effects         =   "frmBattle.frx":00C4
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl imgSwitch 
         Height          =   720
         Index           =   1
         Left            =   120
         Top             =   360
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Frame           =   0
         Attr            =   1536
         Effects         =   "frmBattle.frx":00DC
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl imgSwitch 
         Height          =   720
         Index           =   2
         Left            =   120
         Top             =   1200
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Frame           =   0
         Attr            =   1536
         Effects         =   "frmBattle.frx":00F4
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl imgSwitch 
         Height          =   720
         Index           =   3
         Left            =   120
         Top             =   2040
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Frame           =   0
         Attr            =   1536
         Effects         =   "frmBattle.frx":010C
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl imgSwitch 
         Height          =   720
         Index           =   6
         Left            =   120
         Top             =   4560
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Frame           =   0
         Attr            =   1536
         Effects         =   "frmBattle.frx":0124
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl imgSwitch 
         Height          =   720
         Index           =   5
         Left            =   120
         Top             =   3720
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Frame           =   0
         Attr            =   1536
         Effects         =   "frmBattle.frx":013C
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl imgSwitch 
         Height          =   720
         Index           =   4
         Left            =   120
         Top             =   2880
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Frame           =   0
         Attr            =   1536
         Effects         =   "frmBattle.frx":0154
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl leftStage 
         Height          =   5775
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   10186
         Frame           =   12
         BackColor       =   5590082
         Effects         =   "frmBattle.frx":016C
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl picMyPoke 
         Height          =   1815
         Left            =   1800
         Top             =   3120
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   3201
         Image           =   "frmBattle.frx":0184
         Attr            =   1539
         Effects         =   "frmBattle.frx":7D65
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl picEnemyPoke 
         Height          =   1815
         Left            =   5880
         Top             =   2160
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   3201
         Image           =   "frmBattle.frx":7D7D
         Attr            =   1539
         Effects         =   "frmBattle.frx":DB33
      End
      Begin VB.Label lblRound 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   8400
         TabIndex        =   9
         Top             =   4080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblEnemyLevel 
         BackStyle       =   0  'Transparent
         Caption         =   "Lvl. 5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5520
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lvlMyLevel 
         BackStyle       =   0  'Transparent
         Caption         =   "Lvl. 5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblMyName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ivysaur"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label lblEnemyName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ivysaur"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5400
         TabIndex        =   4
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label lblEnemyHP 
         BackStyle       =   0  'Transparent
         Caption         =   "100/100"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6720
         TabIndex        =   3
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblMyHP 
         BackStyle       =   0  'Transparent
         Caption         =   "100/100"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2400
         TabIndex        =   1
         Top             =   2040
         Width           =   735
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl EnemyBack 
         Height          =   855
         Left            =   5400
         Top             =   840
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1508
         Frame           =   12
         BackColor       =   5590082
         Effects         =   "frmBattle.frx":DB4B
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl MyBack 
         Height          =   855
         Left            =   1080
         Top             =   1920
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1508
         Frame           =   12
         BackColor       =   5590082
         Effects         =   "frmBattle.frx":DB63
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl rightStage 
         Height          =   6000
         Left            =   8040
         Top             =   -120
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   10583
         Frame           =   12
         BackColor       =   5590082
         Effects         =   "frmBattle.frx":DB7B
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl1 
         Height          =   1080
         Left            =   5520
         Top             =   2760
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   1905
         Image           =   "frmBattle.frx":DB93
         Attr            =   515
         Effects         =   "frmBattle.frx":DF3A
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl2 
         Height          =   1080
         Left            =   1560
         Top             =   3840
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   1905
         Image           =   "frmBattle.frx":DF52
         Attr            =   515
         Effects         =   "frmBattle.frx":E2F9
      End
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl3 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   21000
      _ExtentX        =   37042
      _ExtentY        =   873
      Image           =   "frmBattle.frx":E311
      Effects         =   "frmBattle.frx":EF8C
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   1935
      Left            =   2400
      Top             =   1080
      Width           =   2175
   End
End
Attribute VB_Name = "frmBattle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim originalwidth As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
 
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
 
Private Const WM_SYSCOMMAND = &H112



Private Sub Command1_Click()
SendBattleCommand 2, BattlePokemon, 1 'Bag wont do any demage thats why we send move 1
End Sub

Private Sub Command2_Click()
SendBattleCommand 4, BattlePokemon, 1 'Run away
End Sub

Private Sub Command3_Click()
SendBattleCommand 3, BattlePokemon, 1 'Switch pokemon
End Sub

Private Sub AlphaImgCtl3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

  ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
   

End Sub

Private Sub cmdMove_Click(Index As Integer)
PlayClick
Dim i As Long
Dim n As Long
For i = 1 To 4
If PokemonInstance(BattlePokemon).moves(i).number = 0 Then n = n + 1
Next
If cmdMove(Index).Caption = "Struggle (Infinite)" Then
If PokemonInstance(BattlePokemon).HP > 0 Then
SendBattleCommand 4, BattlePokemon, Index
BlockBattle
End If
Else
If Not n >= 4 Then
If PokemonInstance(BattlePokemon).HP > 0 Then
If PokemonInstance(BattlePokemon).moves(Index).number > 0 Then
If PokemonInstance(BattlePokemon).moves(Index).pp > 0 Then
SendBattleCommand 1, BattlePokemon, Index
BlockBattle
End If
End If
End If
End If
End If

End Sub

Private Sub Form_Load()
On Error Resume Next
originalwidth = 169
'MyBack.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\stage.png")
'EnemyBack.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\stage.png")
'Call MyBack.SetFixedSizeAspect(MyBack.width / 15, MyBack.height / 15, True)
'Call EnemyBack.SetFixedSizeAspect(EnemyBack.width / 15, EnemyBack.height / 15, True)
'leftStage.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\stage.png")
'rightStage.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\stage.png")
'Call leftStage.SetFixedSizeAspect(leftStage.width / 15, leftStage.height / 15, True)
'Call rightStage.SetFixedSizeAspect(rightStage.width / 15, rightStage.height / 15, True)
maleimg.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\male.png")
vsimage.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\vs.png")
Call vsimage.SetFixedSizeAspect(vsimage.Width / 15, vsimage.Height / 15, True)
EnemyImg(0).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokeicons\" & enemyPokemon.PokemonNumber & ".png")
lblInfoPoke.Caption = lblEnemyName.Caption
picBattleInfo.Visible = False
End Sub


Sub loadGUI(ByVal mypoke As Long, ByVal enemypoke As Long, ByVal myhp As Long, ByVal mymaxhp As Long, ByVal enemyhp As Long, ByVal enemymaxhp As Long, ByVal isEnemyShiny As Long, ByVal AmIShiny As Long)
If mypoke > 0 Then
If AmIShiny = YES Then
picMyPoke.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\Shiny\Back\" & mypoke & ".gif")
Else
picMyPoke.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\Back\" & mypoke & ".gif")
End If
lblEXP.Caption = Val(PokemonInstance(BattlePokemon).EXP) & "/" & Val(PokemonInstance(BattlePokemon).expNeeded)
If PokemonInstance(BattlePokemon).HP <= 0 Then
picMyPoke.Inverted = True
Else
picMyPoke.Inverted = False
End If
End If
If enemypoke > 0 Then
If isEnemyShiny = YES Then
picEnemyPoke.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\Shiny\" & enemypoke & ".gif")
Else
picEnemyPoke.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\" & enemypoke & ".gif")
End If

If enemyPokemon.HP <= 0 Then
picEnemyPoke.Inverted = True
Else
picEnemyPoke.Inverted = False
End If
End If

End Sub

Private Sub imgSwitch_Click(Index As Integer)
PlayClick
If PokemonInstance(Index).PokemonNumber > 0 Then
If PokemonInstance(Index).HP > 0 Then
If Not Index = BattlePokemon Then
PlayClick
SendBattleCommand 2, Index, 1
BlockBattle
picSwitch.Visible = False
End If
End If
End If

End Sub

Private Sub lvButton1_Click()
UpdateBattle
picSwitch.Visible = True
End Sub

Private Sub lvButton3_Click()

End Sub

Private Sub lvButton4_Click()
picSwitch.Visible = False
End Sub

Private Sub lvButton2_Click()

End Sub

Private Sub lvButton5_Click()

End Sub

Private Sub listBag_DblClick()
Dim invnum As Long
invnum = listBag.ListIndex + 1
SendBattleCommand 3, BattlePokemon, invnum
BlockBattle
listBag.Visible = False
End Sub

Private Sub lvButtons_H1_Click()
PlayClick
picSwitch.Visible = False
End Sub

Private Sub lvButtons_H2_Click()


picBattleInfo.Visible = False
frmMainGame.Enabled = True
frmMainGame.SetFocus
txtBtlLog.text = vbNullString
StopPlay
PlayMapMusic MapMusic
PlayClick
Unload Me
End Sub

Private Sub lvButtons_H3_Click()
If PokemonInstance(BattlePokemon).HP > 0 Then
listBag.Clear
UpdateBattle
PlayClick
listBag.Visible = Not listBag.Visible
Dim i As Long
For i = 1 To MAX_INV
If GetPlayerInvItemNum(MyIndex, i) > 0 Then
listBag.AddItem Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name & " x" & GetPlayerInvItemValue(MyIndex, i))
Else
listBag.AddItem ("Empty")
End If
Next
End If
End Sub

Private Sub lvButtons_H4_Click()
'Try to run
SendBattleCommand 5, BattlePokemon, 1
BlockBattle
End Sub

Private Sub Timer1_Timer()
If Player(MyIndex).inBattle = False And picBattleInfo.Visible = False Then
Unload Me
End If
End Sub

Private Sub lvButtons_H5_Click()
If Player(MyIndex).inBattle = False Then Unload Me
End Sub

Private Sub lvButtons_H6_Click()
AutoCloseBattle = Not AutoCloseBattle
End Sub

