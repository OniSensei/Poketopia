VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCN.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{50347EDF-F3EF-4392-AFDD-71AE67A3A978}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LVbutton.ocx"
Object = "{8B280688-BAE6-4FF7-8D90-A8B9EB792A13}#5.0#0"; "RichtextEditor.ocx"
Begin VB.Form frmMainGame 
   BackColor       =   &H00D9B870&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9615
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   22530
   ControlBox      =   0   'False
   FillColor       =   &H00D9B870&
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMainGame.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainGame.frx":3AFA
   ScaleHeight     =   641
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1502
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock Socket 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picDialog 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   2310
      Left            =   7680
      Picture         =   "frmMainGame.frx":4E0C4
      ScaleHeight     =   154
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   497
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   7455
      Begin lvButton.lvButtons_H lvButtons_H10 
         Height          =   360
         Left            =   1560
         TabIndex        =   2
         Top             =   1800
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   635
         Caption         =   "Next"
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
         cFHover         =   4210752
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdClanYES 
         Height          =   375
         Left            =   6600
         TabIndex        =   219
         Top             =   1800
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         Caption         =   "Yes"
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
         cFHover         =   4210752
         cBhover         =   14737632
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdClanNO 
         Height          =   375
         Left            =   5760
         TabIndex        =   220
         Top             =   1800
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         Caption         =   "No"
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
         cFHover         =   4210752
         cBhover         =   14737632
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl imgDialogPic 
         Height          =   2175
         Left            =   60
         Top             =   60
         Visible         =   0   'False
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   3836
         Effects         =   "frmMainGame.frx":86290
      End
      Begin VB.Label txtDialog 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Dialog"
         BeginProperty Font 
            Name            =   "Eurostar"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1335
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   6615
      End
   End
   Begin VB.PictureBox picMenus 
      BackColor       =   &H00312920&
      BorderStyle     =   0  'None
      Height          =   9615
      Left            =   10560
      Picture         =   "frmMainGame.frx":862A8
      ScaleHeight     =   9615
      ScaleWidth      =   11970
      TabIndex        =   3
      Top             =   0
      Width           =   11970
      Begin VB.PictureBox picLogin 
         BackColor       =   &H00D9B870&
         BorderStyle     =   0  'None
         FillColor       =   &H00D9B870&
         Height          =   9645
         Left            =   1080
         Picture         =   "frmMainGame.frx":A7C60
         ScaleHeight     =   643
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1576
         TabIndex        =   39
         Top             =   0
         Width           =   23640
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00312920&
            BorderStyle     =   0  'None
            Height          =   2010
            Left            =   3720
            Picture         =   "frmMainGame.frx":F222A
            ScaleHeight     =   2010
            ScaleWidth      =   4185
            TabIndex        =   343
            Top             =   3360
            Visible         =   0   'False
            Width           =   4185
            Begin VB.TextBox txtRUser 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00373737&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Left            =   1440
               TabIndex        =   346
               Top             =   600
               Width           =   2535
            End
            Begin VB.TextBox txtRPass 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00373737&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               IMEMode         =   3  'DISABLE
               Left            =   1440
               PasswordChar    =   "•"
               TabIndex        =   345
               Top             =   960
               Width           =   2535
            End
            Begin VB.TextBox txtRPass2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00373737&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               IMEMode         =   3  'DISABLE
               Left            =   1440
               PasswordChar    =   "•"
               TabIndex        =   344
               Top             =   1320
               Width           =   2535
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl13 
               Height          =   255
               Index           =   7
               Left            =   240
               Top             =   1680
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   450
               Image           =   "frmMainGame.frx":FA53E
               Attr            =   514
               Effects         =   "frmMainGame.frx":25F7D4
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl13 
               Height          =   255
               Index           =   6
               Left            =   3240
               Top             =   1680
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   450
               Image           =   "frmMainGame.frx":25F7EC
               Attr            =   514
               Effects         =   "frmMainGame.frx":3C4A82
            End
         End
         Begin VB.PictureBox loginDetailsPicture 
            BackColor       =   &H00312920&
            BorderStyle     =   0  'None
            Height          =   1650
            Left            =   3720
            Picture         =   "frmMainGame.frx":3C4A9A
            ScaleHeight     =   1650
            ScaleWidth      =   4185
            TabIndex        =   340
            Top             =   3720
            Visible         =   0   'False
            Width           =   4185
            Begin VB.TextBox txtLUser 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00373737&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   247
               Left            =   1440
               TabIndex        =   342
               Top             =   600
               Width           =   2535
            End
            Begin VB.TextBox txtLPass 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00373737&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               IMEMode         =   3  'DISABLE
               Left            =   1440
               PasswordChar    =   "•"
               TabIndex        =   341
               Top             =   960
               Width           =   2535
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl13 
               Height          =   255
               Index           =   3
               Left            =   120
               Top             =   1320
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   450
               Image           =   "frmMainGame.frx":3CBAE9
               Attr            =   514
               Effects         =   "frmMainGame.frx":530D7F
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl13 
               Height          =   255
               Index           =   2
               Left            =   2400
               Top             =   1320
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   450
               Image           =   "frmMainGame.frx":530D97
               Attr            =   514
               Effects         =   "frmMainGame.frx":69602D
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl13 
               Height          =   255
               Index           =   0
               Left            =   3240
               Top             =   1320
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   450
               Image           =   "frmMainGame.frx":696045
               Attr            =   514
               Effects         =   "frmMainGame.frx":7FB2DB
            End
         End
         Begin VB.PictureBox picNewChar 
            BackColor       =   &H00312920&
            BorderStyle     =   0  'None
            ForeColor       =   &H00808080&
            Height          =   6210
            Left            =   2640
            Picture         =   "frmMainGame.frx":7FB2F3
            ScaleHeight     =   414
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   423
            TabIndex        =   42
            Top             =   2880
            Visible         =   0   'False
            Width           =   6345
            Begin VB.PictureBox picSprite 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   720
               Left            =   6720
               ScaleHeight     =   48
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   47
               Top             =   1320
               Width           =   480
            End
            Begin VB.TextBox txtCName 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00373737&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Left            =   3840
               TabIndex        =   46
               Top             =   600
               Width           =   2295
            End
            Begin VB.ComboBox cmbClass 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   6360
               Style           =   2  'Dropdown List
               TabIndex        =   45
               Top             =   840
               Visible         =   0   'False
               Width           =   2175
            End
            Begin VB.OptionButton optMale 
               Appearance      =   0  'Flat
               BackColor       =   &H00373737&
               Caption         =   "Male"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   4290
               MaskColor       =   &H00373737&
               TabIndex        =   44
               Top             =   990
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton optFemale 
               Appearance      =   0  'Flat
               BackColor       =   &H00373737&
               Caption         =   "Female"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   4290
               MaskColor       =   &H00373737&
               TabIndex        =   43
               Top             =   1260
               Width           =   1095
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl13 
               Height          =   255
               Index           =   4
               Left            =   5280
               Top             =   5760
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   450
               Image           =   "frmMainGame.frx":807F76
               Attr            =   514
               Effects         =   "frmMainGame.frx":96D20C
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl Starter 
               Height          =   135
               Index           =   22
               Left            =   5880
               Top             =   2040
               Width           =   120
               _ExtentX        =   212
               _ExtentY        =   238
               Image           =   "frmMainGame.frx":96D224
               Attr            =   513
               Effects         =   "frmMainGame.frx":AD24BA
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl Starter 
               Height          =   135
               Index           =   21
               Left            =   3960
               Top             =   2040
               Width           =   120
               _ExtentX        =   212
               _ExtentY        =   238
               Image           =   "frmMainGame.frx":AD24D2
               Attr            =   513
               Effects         =   "frmMainGame.frx":C37768
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl Starter 
               Height          =   960
               Index           =   20
               Left            =   4560
               Top             =   1920
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   1693
               Image           =   "frmMainGame.frx":C37780
               Attr            =   513
               Effects         =   "frmMainGame.frx":D9CA16
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl Starter 
               Height          =   135
               Index           =   19
               Left            =   3960
               Top             =   1680
               Width           =   120
               _ExtentX        =   212
               _ExtentY        =   238
               Image           =   "frmMainGame.frx":D9CA2E
               Attr            =   513
               Effects         =   "frmMainGame.frx":F01CC4
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl Starter 
               Height          =   135
               Index           =   18
               Left            =   5880
               Top             =   1680
               Width           =   120
               _ExtentX        =   212
               _ExtentY        =   238
               Image           =   "frmMainGame.frx":F01CDC
               Attr            =   513
               Effects         =   "frmMainGame.frx":1066F72
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl Starter 
               Height          =   135
               Index           =   17
               Left            =   5880
               Top             =   2400
               Width           =   120
               _ExtentX        =   212
               _ExtentY        =   238
               Image           =   "frmMainGame.frx":1066F8A
               Attr            =   513
               Effects         =   "frmMainGame.frx":11CC220
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl Starter 
               Height          =   135
               Index           =   16
               Left            =   3960
               Top             =   2400
               Width           =   120
               _ExtentX        =   212
               _ExtentY        =   238
               Image           =   "frmMainGame.frx":11CC238
               Attr            =   513
               Effects         =   "frmMainGame.frx":13314CE
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl Starter 
               Height          =   960
               Index           =   15
               Left            =   4560
               Top             =   1920
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   1693
               Image           =   "frmMainGame.frx":13314E6
               Attr            =   513
               Effects         =   "frmMainGame.frx":149677C
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl Starter 
               Height          =   450
               Index           =   14
               Left            =   2055
               Top             =   4050
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   794
               Image           =   "frmMainGame.frx":1496794
               Attr            =   513
               Effects         =   "frmMainGame.frx":15FBA2A
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl Starter 
               Height          =   450
               Index           =   13
               Left            =   1155
               Top             =   4050
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   794
               Image           =   "frmMainGame.frx":15FBA42
               Attr            =   513
               Effects         =   "frmMainGame.frx":1760CD8
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl Starter 
               Height          =   450
               Index           =   12
               Left            =   255
               Top             =   4050
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   794
               Image           =   "frmMainGame.frx":1760CF0
               Attr            =   513
               Effects         =   "frmMainGame.frx":18C5F86
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl Starter 
               Height          =   450
               Index           =   11
               Left            =   2055
               Top             =   3300
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   794
               Image           =   "frmMainGame.frx":18C5F9E
               Attr            =   513
               Effects         =   "frmMainGame.frx":1A2B234
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl Starter 
               Height          =   450
               Index           =   10
               Left            =   1155
               Top             =   3300
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   794
               Image           =   "frmMainGame.frx":1A2B24C
               Attr            =   513
               Effects         =   "frmMainGame.frx":1B904E2
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl Starter 
               Height          =   450
               Index           =   9
               Left            =   255
               Top             =   3300
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   794
               Image           =   "frmMainGame.frx":1B904FA
               Attr            =   513
               Effects         =   "frmMainGame.frx":1CF5790
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl Starter 
               Height          =   450
               Index           =   8
               Left            =   2055
               Top             =   2550
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   794
               Image           =   "frmMainGame.frx":1CF57A8
               Attr            =   513
               Effects         =   "frmMainGame.frx":1E5AA3E
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl Starter 
               Height          =   450
               Index           =   7
               Left            =   1155
               Top             =   2550
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   794
               Image           =   "frmMainGame.frx":1E5AA56
               Attr            =   513
               Effects         =   "frmMainGame.frx":1FBFCEC
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl Starter 
               Height          =   450
               Index           =   6
               Left            =   255
               Top             =   2550
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   794
               Image           =   "frmMainGame.frx":1FBFD04
               Attr            =   513
               Effects         =   "frmMainGame.frx":2124F9A
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl Starter 
               Height          =   450
               Index           =   5
               Left            =   2055
               Top             =   1800
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   794
               Image           =   "frmMainGame.frx":2124FB2
               Attr            =   513
               Effects         =   "frmMainGame.frx":228A248
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl Starter 
               Height          =   450
               Index           =   4
               Left            =   1155
               Top             =   1800
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   794
               Image           =   "frmMainGame.frx":228A260
               Attr            =   513
               Effects         =   "frmMainGame.frx":23EF4F6
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl Starter 
               Height          =   450
               Index           =   3
               Left            =   255
               Top             =   1800
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   794
               Image           =   "frmMainGame.frx":23EF50E
               Attr            =   513
               Effects         =   "frmMainGame.frx":25547A4
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl Starter 
               Height          =   450
               Index           =   2
               Left            =   2055
               Top             =   1050
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   794
               Image           =   "frmMainGame.frx":25547BC
               Attr            =   513
               Effects         =   "frmMainGame.frx":26B9A52
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl Starter 
               Height          =   450
               Index           =   1
               Left            =   1155
               Top             =   1050
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   794
               Image           =   "frmMainGame.frx":26B9A6A
               Attr            =   513
               Effects         =   "frmMainGame.frx":281ED00
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl Starter 
               Height          =   450
               Index           =   0
               Left            =   255
               Top             =   1050
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   794
               Image           =   "frmMainGame.frx":281ED18
               Attr            =   513
               Effects         =   "frmMainGame.frx":2983FAE
            End
            Begin VB.Label lblSpriteRight 
               BackStyle       =   0  'Transparent
               Caption         =   ">"
               Height          =   255
               Left            =   7320
               TabIndex        =   50
               Top             =   1560
               Width           =   255
            End
            Begin VB.Label lblSpriteLeft 
               BackStyle       =   0  'Transparent
               Caption         =   "<"
               Height          =   255
               Left            =   6480
               TabIndex        =   49
               Top             =   1560
               Width           =   255
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
               Left            =   6240
               TabIndex        =   48
               Top             =   600
               Visible         =   0   'False
               Width           =   735
            End
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl13 
            Height          =   1185
            Index           =   1
            Left            =   3360
            Top             =   960
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   2090
            Image           =   "frmMainGame.frx":2983FC6
            Attr            =   513
            Effects         =   "frmMainGame.frx":2AE925C
         End
         Begin VB.Label UselessLabel 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   48
            Left            =   6480
            TabIndex        =   338
            Top             =   3000
            Width           =   2415
         End
         Begin VB.Label UselessLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Poketopia: [ Revival ]"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   47
            Left            =   120
            TabIndex        =   337
            Top             =   9360
            Width           =   2415
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl2 
            Height          =   3660
            Left            =   8640
            Top             =   7080
            Visible         =   0   'False
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   6456
            Image           =   "frmMainGame.frx":2AE9274
            Attr            =   514
            Effects         =   "frmMainGame.frx":2AF910F
         End
         Begin VB.Label lblSGInfo 
            Alignment       =   2  'Center
            BackColor       =   &H00373737&
            Caption         =   "Info"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1815
            Left            =   4200
            TabIndex        =   41
            Top             =   5520
            Width           =   3135
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl5 
            Height          =   3240
            Left            =   -840
            Top             =   0
            Width           =   5100
            _ExtentX        =   8996
            _ExtentY        =   5715
            Image           =   "frmMainGame.frx":2AF9127
            Attr            =   1538
            Effects         =   "frmMainGame.frx":2B23F2D
         End
      End
      Begin VB.Timer tmrDialog 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   600
         Top             =   8520
      End
      Begin VB.PictureBox picHover 
         BackColor       =   &H00D9B870&
         BorderStyle     =   0  'None
         Height          =   9615
         Left            =   10800
         ScaleHeight     =   9615
         ScaleWidth      =   2295
         TabIndex        =   40
         Top             =   0
         Width           =   2295
      End
      Begin lvButton.lvButtons_H btnAdminPanel 
         Height          =   375
         Left            =   10080
         TabIndex        =   125
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Admin Panel"
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
      Begin VB.Timer tmrmenu 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   600
         Top             =   8160
      End
      Begin VB.PictureBox picMore 
         BackColor       =   &H00312920&
         BorderStyle     =   0  'None
         Height          =   4095
         Left            =   1080
         Picture         =   "frmMainGame.frx":2B23F45
         ScaleHeight     =   273
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   73
         TabIndex        =   6
         Top             =   0
         Width           =   1095
         Begin VB.Image Image8 
            Height          =   705
            Left            =   240
            Top             =   3240
            Width           =   690
         End
         Begin VB.Image Image7 
            Height          =   705
            Left            =   255
            Top             =   2475
            Width           =   690
         End
         Begin VB.Image Image2 
            Height          =   600
            Left            =   240
            Top             =   165
            Width           =   690
         End
         Begin VB.Image Image3 
            Height          =   600
            Left            =   240
            Top             =   960
            Width           =   690
         End
         Begin VB.Image Image4 
            Height          =   705
            Left            =   240
            Top             =   1680
            Width           =   690
         End
      End
      Begin VB.PictureBox picTravel 
         BackColor       =   &H00312920&
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   4200
         ScaleHeight     =   3615
         ScaleWidth      =   3015
         TabIndex        =   172
         Top             =   4800
         Visible         =   0   'False
         Width           =   3015
         Begin lvButton.lvButtons_H btnTravel 
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   173
            Top             =   480
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   661
            Caption         =   "Pallet Town- 100 PC"
            CapAlign        =   2
            BackStyle       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            cBack           =   16777215
         End
         Begin lvButton.lvButtons_H btnTravel 
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   174
            Top             =   960
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   661
            Caption         =   "Viridian City - 100 PC"
            CapAlign        =   2
            BackStyle       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            cBack           =   16777215
         End
         Begin lvButton.lvButtons_H lvButtons_H35 
            Height          =   375
            Left            =   240
            TabIndex        =   175
            Top             =   2880
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   661
            Caption         =   "Close"
            CapAlign        =   2
            BackStyle       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            cBack           =   16777215
         End
         Begin lvButton.lvButtons_H btnTravel 
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   229
            Top             =   1440
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   661
            Caption         =   "Pewter City  - 300 PC"
            CapAlign        =   2
            BackStyle       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            cBack           =   16777215
         End
         Begin lvButton.lvButtons_H btnTravel 
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   336
            Top             =   1920
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   661
            Caption         =   "Cerulean City- 1000 PC"
            CapAlign        =   2
            BackStyle       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            cBack           =   16777215
         End
         Begin VB.Label UselessLabel 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Travel to"
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
            Index           =   0
            Left            =   -600
            TabIndex        =   176
            Top             =   0
            Width           =   4335
         End
      End
      Begin VB.PictureBox picEgg 
         BackColor       =   &H00312920&
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   3120
         ScaleHeight     =   2295
         ScaleWidth      =   5535
         TabIndex        =   312
         Top             =   6240
         Visible         =   0   'False
         Width           =   5535
         Begin lvButton.lvButtons_H btnHatch 
            Height          =   375
            Left            =   480
            TabIndex        =   318
            Top             =   1680
            Visible         =   0   'False
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   661
            Caption         =   "Hatch"
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
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   16777215
         End
         Begin VB.Label EggInfo 
            Alignment       =   2  'Center
            BackColor       =   &H00554C42&
            Caption         =   "100000"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   11.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   3480
            TabIndex        =   317
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label UselessLabel 
            Alignment       =   2  'Center
            BackColor       =   &H00554C42&
            Caption         =   "EXP left:"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   11.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   44
            Left            =   1920
            TabIndex        =   316
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label EggInfo 
            Alignment       =   2  'Center
            BackColor       =   &H00554C42&
            Caption         =   "50000"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   11.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   3480
            TabIndex        =   315
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label UselessLabel 
            Alignment       =   2  'Center
            BackColor       =   &H00554C42&
            Caption         =   "Steps left:"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   11.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   42
            Left            =   1920
            TabIndex        =   314
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label UselessLabel 
            Alignment       =   2  'Center
            BackColor       =   &H00554C42&
            Caption         =   "Egg"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   11.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   41
            Left            =   1920
            TabIndex        =   313
            Top             =   120
            Width           =   3135
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl3 
            Height          =   1200
            Left            =   600
            Top             =   120
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   2117
            Image           =   "frmMainGame.frx":2B3478D
            Effects         =   "frmMainGame.frx":2B365E4
         End
      End
      Begin VB.PictureBox picPokedex 
         BackColor       =   &H00312920&
         BorderStyle     =   0  'None
         Height          =   9255
         Left            =   1200
         ScaleHeight     =   9255
         ScaleWidth      =   9255
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   9255
         Begin VB.ListBox PokedexList1 
            Appearance      =   0  'Flat
            BackColor       =   &H00312920&
            ForeColor       =   &H00FFFFFF&
            Height          =   9000
            Left            =   1080
            TabIndex        =   20
            Top             =   0
            Width           =   1935
         End
         Begin VB.ListBox lstPokedexMoves 
            Appearance      =   0  'Flat
            BackColor       =   &H00312920&
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
            Height          =   8220
            Left            =   6120
            TabIndex        =   19
            Top             =   960
            Width           =   2895
         End
         Begin VB.Label lblPokedexPOKE 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Bulbasaur"
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
            Left            =   3000
            TabIndex        =   29
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label UselessLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Base stats"
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
            Height          =   255
            Index           =   19
            Left            =   2040
            TabIndex        =   28
            Top             =   1560
            Width           =   3015
         End
         Begin VB.Label lblPokedexHP 
            BackStyle       =   0  'Transparent
            Caption         =   "HP"
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
            Height          =   255
            Left            =   2040
            TabIndex        =   27
            Top             =   1800
            Width           =   3015
         End
         Begin VB.Label lblPokedexATK 
            BackStyle       =   0  'Transparent
            Caption         =   "ATK"
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
            Height          =   255
            Left            =   1800
            TabIndex        =   26
            Top             =   2040
            Width           =   3015
         End
         Begin VB.Label lblPokedexDEF 
            BackStyle       =   0  'Transparent
            Caption         =   "DEF"
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
            Height          =   255
            Left            =   2520
            TabIndex        =   25
            Top             =   2520
            Width           =   3015
         End
         Begin VB.Label lblPokedexSPATK 
            BackStyle       =   0  'Transparent
            Caption         =   "SP:ATK"
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
            Height          =   255
            Left            =   2040
            TabIndex        =   24
            Top             =   2520
            Width           =   3015
         End
         Begin VB.Label lblPokedexSPDEF 
            BackStyle       =   0  'Transparent
            Caption         =   "SP.DEF"
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
            Height          =   255
            Left            =   2040
            TabIndex        =   23
            Top             =   2760
            Width           =   3015
         End
         Begin VB.Label lblPokedexSPEED 
            BackStyle       =   0  'Transparent
            Caption         =   "SPEED"
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
            Height          =   255
            Left            =   2040
            TabIndex        =   22
            Top             =   3000
            Width           =   3015
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl imgPokedexPoke 
            Height          =   1815
            Left            =   3000
            Top             =   3600
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   3201
            Attr            =   1536
            Effects         =   "frmMainGame.frx":2B365FC
         End
         Begin VB.Image imgPokedexType2 
            Height          =   375
            Left            =   4200
            Stretch         =   -1  'True
            Top             =   5640
            Width           =   1215
         End
         Begin VB.Image imgPokedexType1 
            Height          =   375
            Left            =   2880
            Stretch         =   -1  'True
            Top             =   5640
            Width           =   1215
         End
         Begin VB.Label UselessLabel 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Moves:"
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
            Index           =   20
            Left            =   6000
            TabIndex        =   21
            Top             =   480
            Width           =   3015
         End
      End
      Begin VB.PictureBox picCrew 
         BackColor       =   &H00312920&
         BorderStyle     =   0  'None
         Height          =   9255
         Left            =   2160
         ScaleHeight     =   9255
         ScaleWidth      =   6735
         TabIndex        =   202
         Top             =   120
         Visible         =   0   'False
         Width           =   6735
         Begin RichTextLib.RichTextBox txtClanNews 
            Height          =   2655
            Left            =   480
            TabIndex        =   222
            Top             =   6480
            Visible         =   0   'False
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   4683
            _Version        =   393217
            BackColor       =   0
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            ScrollBars      =   3
            Appearance      =   0
            TextRTF         =   $"frmMainGame.frx":2B36614
         End
         Begin VB.ListBox lstClanMembers 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H00FFFFFF&
            Height          =   5685
            Left            =   3600
            TabIndex        =   207
            Top             =   360
            Width           =   2775
         End
         Begin lvButton.lvButtons_H ClanButton 
            Height          =   375
            Index           =   3
            Left            =   480
            TabIndex        =   208
            Top             =   4440
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            Caption         =   "Kick Player"
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
         Begin lvButton.lvButtons_H ClanButton 
            Height          =   375
            Index           =   2
            Left            =   480
            TabIndex        =   209
            Top             =   3960
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            Caption         =   "Delete clan"
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
         Begin lvButton.lvButtons_H ClanButton 
            Height          =   375
            Index           =   1
            Left            =   480
            TabIndex        =   210
            Top             =   3480
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            Caption         =   "Set clan picture"
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
         Begin lvButton.lvButtons_H ClanButton 
            Height          =   375
            Index           =   4
            Left            =   480
            TabIndex        =   211
            Top             =   5640
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            Caption         =   "Leave clan"
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
         Begin vb6projectRichtextEditor.ucRTBEditor txtClanNewsEdit 
            Height          =   3015
            Left            =   480
            TabIndex        =   221
            Top             =   6120
            Visible         =   0   'False
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   5318
            TextRTF         =   $"frmMainGame.frx":2B36699
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin lvButton.lvButtons_H ClanButton 
            Height          =   375
            Index           =   5
            Left            =   480
            TabIndex        =   223
            Top             =   4920
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            Caption         =   "Edit news"
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
         Begin LaVolpeAlphaImg.AlphaImgCtl imgClan 
            Height          =   1455
            Left            =   840
            Top             =   1440
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   2566
            Image           =   "frmMainGame.frx":2B36728
            Attr            =   514
            Effects         =   "frmMainGame.frx":2B41A88
         End
         Begin VB.Label lblClanName 
            Alignment       =   2  'Center
            BackColor       =   &H00554C42&
            Caption         =   "karps"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   240
            TabIndex        =   204
            Top             =   1080
            Width           =   2775
         End
         Begin VB.Label UselessLabel 
            Alignment       =   2  'Center
            BackColor       =   &H00554C42&
            Caption         =   "Clan"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   31
            Left            =   240
            TabIndex        =   203
            Top             =   600
            Width           =   2775
         End
      End
      Begin VB.PictureBox picPokemons 
         BackColor       =   &H00312920&
         BorderStyle     =   0  'None
         Height          =   11175
         Left            =   1080
         ScaleHeight     =   11175
         ScaleWidth      =   9855
         TabIndex        =   51
         Top             =   -2160
         Visible         =   0   'False
         Width           =   9855
         Begin VB.PictureBox picInformation 
            BackColor       =   &H00312920&
            BorderStyle     =   0  'None
            Height          =   7335
            Left            =   2280
            ScaleHeight     =   7335
            ScaleWidth      =   10455
            TabIndex        =   52
            Top             =   2400
            Width           =   10455
            Begin VB.PictureBox RosterpicMoves 
               BackColor       =   &H00554C42&
               BorderStyle     =   0  'None
               Height          =   4215
               Left            =   -240
               ScaleHeight     =   4215
               ScaleWidth      =   6135
               TabIndex        =   53
               Top             =   1680
               Visible         =   0   'False
               Width           =   6135
               Begin VB.ListBox RosterlstMoves 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0C0C0&
                  BeginProperty Font 
                     Name            =   "Eurostar"
                     Size            =   9.75
                     Charset         =   238
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   3180
                  Left            =   3120
                  TabIndex        =   54
                  Top             =   600
                  Width           =   2895
               End
               Begin lvButton.lvButtons_H lvButtons_H15 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   55
                  Top             =   3600
                  Width           =   2895
                  _ExtentX        =   5106
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
                  cBhover         =   12632256
                  LockHover       =   3
                  cGradient       =   0
                  Mode            =   0
                  Value           =   0   'False
                  cBack           =   16777215
               End
               Begin VB.Label RosterlblMove 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "Fire Spin"
                  BeginProperty Font 
                     Name            =   "Eurostar"
                     Size            =   15.75
                     Charset         =   238
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   615
                  Index           =   1
                  Left            =   0
                  TabIndex        =   60
                  Top             =   600
                  Width           =   3135
               End
               Begin VB.Label RosterlblMove 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "001"
                  BeginProperty Font 
                     Name            =   "Eurostar"
                     Size            =   15.75
                     Charset         =   238
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   615
                  Index           =   2
                  Left            =   0
                  TabIndex        =   59
                  Top             =   1320
                  Width           =   3135
               End
               Begin VB.Label RosterlblMove 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "001"
                  BeginProperty Font 
                     Name            =   "Eurostar"
                     Size            =   15.75
                     Charset         =   238
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   615
                  Index           =   3
                  Left            =   0
                  TabIndex        =   58
                  Top             =   2040
                  Width           =   3135
               End
               Begin VB.Label RosterlblMove 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "001"
                  BeginProperty Font 
                     Name            =   "Eurostar"
                     Size            =   15.75
                     Charset         =   238
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   615
                  Index           =   4
                  Left            =   0
                  TabIndex        =   57
                  Top             =   2760
                  Width           =   3135
               End
               Begin VB.Label UselessLabel 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "Double click to learn move."
                  ForeColor       =   &H00404040&
                  Height          =   255
                  Index           =   4
                  Left            =   3120
                  TabIndex        =   56
                  Top             =   240
                  Width           =   2895
               End
            End
            Begin VB.PictureBox RosterpicItems 
               BackColor       =   &H00554C42&
               BorderStyle     =   0  'None
               Height          =   4215
               Left            =   0
               ScaleHeight     =   4215
               ScaleWidth      =   6135
               TabIndex        =   115
               Top             =   0
               Visible         =   0   'False
               Width           =   6135
               Begin VB.ListBox RosterlstItems 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0C0C0&
                  BeginProperty Font 
                     Name            =   "Eurostar"
                     Size            =   9.75
                     Charset         =   238
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   2340
                  Left            =   120
                  TabIndex        =   116
                  Top             =   1080
                  Width           =   5895
               End
               Begin lvButton.lvButtons_H lvButtons_H2 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   117
                  Top             =   3600
                  Width           =   5895
                  _ExtentX        =   10398
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
                  cBhover         =   12632256
                  LockHover       =   3
                  cGradient       =   0
                  Mode            =   0
                  Value           =   0   'False
                  cBack           =   16777215
               End
               Begin VB.Label UselessLabel 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "   Use items"
                  BeginProperty Font 
                     Name            =   "Eurostar"
                     Size            =   15.75
                     Charset         =   238
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   615
                  Index           =   2
                  Left            =   1080
                  TabIndex        =   119
                  Top             =   120
                  Width           =   6135
               End
               Begin VB.Label UselessLabel 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "Double click to use"
                  ForeColor       =   &H00404040&
                  Height          =   255
                  Index           =   3
                  Left            =   1680
                  TabIndex        =   118
                  Top             =   720
                  Width           =   2895
               End
            End
            Begin VB.CommandButton Command2 
               Caption         =   "+"
               Height          =   255
               Left            =   3600
               TabIndex        =   74
               Top             =   1920
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.CommandButton Command3 
               Caption         =   "+"
               Height          =   255
               Left            =   3600
               TabIndex        =   73
               Top             =   2400
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.CommandButton Command4 
               Caption         =   "+"
               Height          =   255
               Left            =   3600
               TabIndex        =   72
               Top             =   2880
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.CommandButton Command5 
               Caption         =   "+"
               Height          =   255
               Left            =   3600
               TabIndex        =   71
               Top             =   3360
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.CommandButton Command6 
               Caption         =   "+"
               Height          =   255
               Left            =   3600
               TabIndex        =   70
               Top             =   3840
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.CommandButton Command7 
               Caption         =   "+"
               Height          =   255
               Left            =   3600
               TabIndex        =   69
               Top             =   4320
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.Timer Timer1 
               Interval        =   1000
               Left            =   9240
               Top             =   4200
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H00554C42&
               BorderStyle     =   0  'None
               Height          =   1695
               Left            =   0
               ScaleHeight     =   1695
               ScaleWidth      =   1575
               TabIndex        =   61
               Top             =   5160
               Width           =   1575
               Begin VB.Label lblNatureSpd 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Speed +1"
                  BeginProperty Font 
                     Name            =   "Eurostar"
                     Size            =   8.25
                     Charset         =   238
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   0
                  TabIndex        =   68
                  Top             =   1200
                  Width           =   1215
               End
               Begin VB.Label lblNatureSpDef 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Sp.Def +1"
                  BeginProperty Font 
                     Name            =   "Eurostar"
                     Size            =   8.25
                     Charset         =   238
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   0
                  TabIndex        =   67
                  Top             =   960
                  Width           =   1215
               End
               Begin VB.Label lblNatureSpAtk 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Sp.Atk +1"
                  BeginProperty Font 
                     Name            =   "Eurostar"
                     Size            =   8.25
                     Charset         =   238
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   0
                  TabIndex        =   66
                  Top             =   720
                  Width           =   1215
               End
               Begin VB.Label lblNatureDef 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Def +1"
                  BeginProperty Font 
                     Name            =   "Eurostar"
                     Size            =   8.25
                     Charset         =   238
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   0
                  TabIndex        =   65
                  Top             =   480
                  Width           =   1215
               End
               Begin VB.Label lblNatureAtk 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Atk +1"
                  BeginProperty Font 
                     Name            =   "Eurostar"
                     Size            =   8.25
                     Charset         =   238
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   0
                  TabIndex        =   64
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.Label UselessLabel 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nature boost"
                  BeginProperty Font 
                     Name            =   "Eurostar"
                     Size            =   8.25
                     Charset         =   238
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   375
                  Index           =   16
                  Left            =   120
                  TabIndex        =   63
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.Label lblNatureHP 
                  BackStyle       =   0  'Transparent
                  Caption         =   "HP +1"
                  BeginProperty Font 
                     Name            =   "Eurostar"
                     Size            =   8.25
                     Charset         =   238
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   0
                  TabIndex        =   62
                  Top             =   1440
                  Width           =   1215
               End
            End
            Begin lvButton.lvButtons_H lvButtons_H16 
               Height          =   375
               Left            =   5280
               TabIndex        =   75
               Top             =   5160
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               Caption         =   "Leader"
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
               cBhover         =   12632256
               LockHover       =   3
               cGradient       =   0
               Mode            =   0
               Value           =   0   'False
               cBack           =   16777215
            End
            Begin lvButton.lvButtons_H PokeButton 
               Height          =   375
               Index           =   1
               Left            =   1680
               TabIndex        =   76
               Top             =   5160
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               Caption         =   "Set as leader"
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
               cBhover         =   12632256
               LockHover       =   3
               cGradient       =   0
               Mode            =   0
               Value           =   0   'False
               cBack           =   16777215
            End
            Begin lvButton.lvButtons_H PokeButton 
               Height          =   375
               Index           =   6
               Left            =   3360
               TabIndex        =   77
               Top             =   6120
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               Caption         =   "Set as leader"
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
               cBhover         =   12632256
               LockHover       =   3
               cGradient       =   0
               Mode            =   0
               Value           =   0   'False
               cBack           =   16777215
            End
            Begin lvButton.lvButtons_H PokeButton 
               Height          =   375
               Index           =   2
               Left            =   1680
               TabIndex        =   78
               Top             =   5640
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               Caption         =   "Set as leader"
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
               cBhover         =   12632256
               LockHover       =   3
               cGradient       =   0
               Mode            =   0
               Value           =   0   'False
               cBack           =   16777215
            End
            Begin lvButton.lvButtons_H PokeButton 
               Height          =   375
               Index           =   3
               Left            =   1680
               TabIndex        =   79
               Top             =   6120
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               Caption         =   "Set as leader"
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
               cBhover         =   12632256
               LockHover       =   3
               cGradient       =   0
               Mode            =   0
               Value           =   0   'False
               cBack           =   16777215
            End
            Begin lvButton.lvButtons_H PokeButton 
               Height          =   375
               Index           =   4
               Left            =   3360
               TabIndex        =   80
               Top             =   5160
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               Caption         =   "Set as leader"
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
               cBhover         =   12632256
               LockHover       =   3
               cGradient       =   0
               Mode            =   0
               Value           =   0   'False
               cBack           =   16777215
            End
            Begin lvButton.lvButtons_H PokeButton 
               Height          =   375
               Index           =   5
               Left            =   3360
               TabIndex        =   81
               Top             =   5640
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               Caption         =   "Set as leader"
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
               cBhover         =   12632256
               LockHover       =   3
               cGradient       =   0
               Mode            =   0
               Value           =   0   'False
               cBack           =   16777215
            End
            Begin lvButton.lvButtons_H lvButtons_H17 
               Height          =   375
               Left            =   5280
               TabIndex        =   82
               Top             =   6120
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               Caption         =   "Moves"
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
               cBhover         =   12632256
               LockHover       =   3
               cGradient       =   0
               Mode            =   0
               Value           =   0   'False
               cBack           =   16777215
            End
            Begin lvButton.lvButtons_H lvButtons_H18 
               Height          =   375
               Left            =   5280
               TabIndex        =   83
               Top             =   6600
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               Caption         =   "Evolve"
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
               cBhover         =   12632064
               LockHover       =   3
               cGradient       =   0
               Mode            =   0
               Value           =   0   'False
               cBack           =   16776960
            End
            Begin lvButton.lvButtons_H lvButtons_H19 
               Height          =   375
               Left            =   5280
               TabIndex        =   84
               Top             =   5640
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               Caption         =   "Items"
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
               cBhover         =   12632256
               LockHover       =   3
               cGradient       =   0
               Mode            =   0
               Value           =   0   'False
               cBack           =   16777215
            End
            Begin lvButton.lvButtons_H lvButtons_H20 
               Height          =   300
               Left            =   3600
               TabIndex        =   85
               Top             =   1920
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   529
               Caption         =   "+"
               CapAlign        =   2
               BackStyle       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Eurostar"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFore           =   0
               cFHover         =   0
               cBhover         =   12632064
               LockHover       =   3
               cGradient       =   0
               Mode            =   0
               Value           =   0   'False
               cBack           =   16776960
            End
            Begin lvButton.lvButtons_H lvButtons_H21 
               Height          =   300
               Left            =   3600
               TabIndex        =   86
               Top             =   2400
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   529
               Caption         =   "+"
               CapAlign        =   2
               BackStyle       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Eurostar"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFore           =   0
               cFHover         =   0
               cBhover         =   12632064
               LockHover       =   3
               cGradient       =   0
               Mode            =   0
               Value           =   0   'False
               cBack           =   16776960
            End
            Begin lvButton.lvButtons_H lvButtons_H22 
               Height          =   300
               Left            =   3600
               TabIndex        =   87
               Top             =   2880
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   529
               Caption         =   "+"
               CapAlign        =   2
               BackStyle       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Eurostar"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFore           =   0
               cFHover         =   0
               cBhover         =   12632064
               LockHover       =   3
               cGradient       =   0
               Mode            =   0
               Value           =   0   'False
               cBack           =   16776960
            End
            Begin lvButton.lvButtons_H lvButtons_H23 
               Height          =   300
               Left            =   3600
               TabIndex        =   88
               Top             =   3360
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   529
               Caption         =   "+"
               CapAlign        =   2
               BackStyle       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Eurostar"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFore           =   0
               cFHover         =   0
               cBhover         =   12632064
               LockHover       =   3
               cGradient       =   0
               Mode            =   0
               Value           =   0   'False
               cBack           =   16776960
            End
            Begin lvButton.lvButtons_H lvButtons_H24 
               Height          =   300
               Left            =   3600
               TabIndex        =   89
               Top             =   3840
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   529
               Caption         =   "+"
               CapAlign        =   2
               BackStyle       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Eurostar"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFore           =   0
               cFHover         =   0
               cBhover         =   12632064
               LockHover       =   3
               cGradient       =   0
               Mode            =   0
               Value           =   0   'False
               cBack           =   16776960
            End
            Begin lvButton.lvButtons_H lvButtons_H25 
               Height          =   300
               Left            =   3600
               TabIndex        =   90
               Top             =   4320
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   529
               Caption         =   "+"
               CapAlign        =   2
               BackStyle       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Eurostar"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFore           =   0
               cFHover         =   0
               cBhover         =   12632064
               LockHover       =   3
               cGradient       =   0
               Mode            =   0
               Value           =   0   'False
               cBack           =   16776960
            End
            Begin VB.Label UselessLabel 
               Alignment       =   2  'Center
               BackColor       =   &H00554C42&
               Caption         =   "Holding item:"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   32
               Left            =   4560
               TabIndex        =   226
               Top             =   960
               Width           =   2175
            End
            Begin VB.Label lblPokeItem 
               Alignment       =   2  'Center
               BackColor       =   &H00554C42&
               Caption         =   "None"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   11.25
                  Charset         =   238
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   4560
               TabIndex        =   225
               Top             =   1200
               Width           =   2175
            End
            Begin VB.Label Rosterlblnum 
               Alignment       =   2  'Center
               BackColor       =   &H00554C42&
               Caption         =   "001"
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
               Left            =   4800
               TabIndex        =   187
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label UselessLabel 
               BackColor       =   &H00554C42&
               Caption         =   "Name:"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   5
               Left            =   0
               TabIndex        =   114
               Top             =   120
               Width           =   615
            End
            Begin VB.Label RosterlblName 
               BackColor       =   &H00312920&
               Caption         =   "Bulbasaur"
               ForeColor       =   &H0000FFFF&
               Height          =   255
               Left            =   720
               TabIndex        =   113
               Top             =   120
               Width           =   2775
            End
            Begin VB.Label UselessLabel 
               BackColor       =   &H00554C42&
               Caption         =   "Level:"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   6
               Left            =   0
               TabIndex        =   112
               Top             =   480
               Width           =   615
            End
            Begin VB.Label UselessLabel 
               BackColor       =   &H00554C42&
               Caption         =   "Nature:"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   7
               Left            =   0
               TabIndex        =   111
               Top             =   840
               Width           =   615
            End
            Begin VB.Label UselessLabel 
               BackColor       =   &H00554C42&
               Caption         =   "ATK:"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   10
               Left            =   0
               TabIndex        =   110
               Top             =   2400
               Width           =   615
            End
            Begin VB.Label UselessLabel 
               BackColor       =   &H00554C42&
               Caption         =   "DEF:"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   11
               Left            =   0
               TabIndex        =   109
               Top             =   2880
               Width           =   615
            End
            Begin VB.Label UselessLabel 
               BackColor       =   &H00554C42&
               Caption         =   "SPATK:"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   12
               Left            =   0
               TabIndex        =   108
               Top             =   3360
               Width           =   615
            End
            Begin VB.Label UselessLabel 
               BackColor       =   &H00554C42&
               Caption         =   "SPDEF:"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   13
               Left            =   0
               TabIndex        =   107
               Top             =   3840
               Width           =   615
            End
            Begin VB.Label UselessLabel 
               BackColor       =   &H00554C42&
               Caption         =   "SPEED:"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   14
               Left            =   0
               TabIndex        =   106
               Top             =   4320
               Width           =   615
            End
            Begin VB.Label UselessLabel 
               BackColor       =   &H00554C42&
               Caption         =   "HP:"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   9
               Left            =   0
               TabIndex        =   105
               Top             =   1920
               Width           =   615
            End
            Begin VB.Label RosterlblLevel 
               BackColor       =   &H00312920&
               Caption         =   "Bulbasaur"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   720
               TabIndex        =   104
               Top             =   480
               Width           =   2775
            End
            Begin VB.Label RosterlblNature 
               BackColor       =   &H00312920&
               Caption         =   "Bulbasaur"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   720
               TabIndex        =   103
               Top             =   840
               Width           =   2775
            End
            Begin VB.Label RosterlblHp 
               BackColor       =   &H00312920&
               Caption         =   "Bulbasaur"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   720
               TabIndex        =   102
               Top             =   1200
               Width           =   2775
            End
            Begin VB.Label RosterlblMaxHp 
               BackColor       =   &H00312920&
               Caption         =   "Bulbasaur"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   720
               TabIndex        =   101
               Top             =   1920
               Width           =   2775
            End
            Begin VB.Label RosterlblAtk 
               BackColor       =   &H00312920&
               Caption         =   "Bulbasaur"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   720
               TabIndex        =   100
               Top             =   2400
               Width           =   2775
            End
            Begin VB.Label RosterlblDef 
               BackColor       =   &H00312920&
               Caption         =   "Bulbasaur"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   720
               TabIndex        =   99
               Top             =   2880
               Width           =   2775
            End
            Begin VB.Label RosterlblSpAtk 
               BackColor       =   &H00312920&
               Caption         =   "Bulbasaur"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   720
               TabIndex        =   98
               Top             =   3360
               Width           =   2775
            End
            Begin VB.Label RosterlblSpDef 
               BackColor       =   &H00312920&
               Caption         =   "Bulbasaur"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   720
               TabIndex        =   97
               Top             =   3840
               Width           =   2775
            End
            Begin VB.Label RosterlblSpeed 
               BackColor       =   &H00312920&
               Caption         =   "Bulbasaur"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   720
               TabIndex        =   96
               Top             =   4320
               Width           =   2775
            End
            Begin VB.Label UselessLabel 
               BackColor       =   &H00554C42&
               Caption         =   "Hp:"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   8
               Left            =   0
               TabIndex        =   95
               Top             =   1200
               Width           =   615
            End
            Begin VB.Label UselessLabel 
               BackColor       =   &H0080FFFF&
               Caption         =   "TP:"
               Height          =   255
               Index           =   15
               Left            =   0
               TabIndex        =   94
               Top             =   4800
               Width           =   615
            End
            Begin VB.Label RosterlblTP 
               BackColor       =   &H00312920&
               Caption         =   "Bulbasaur"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   720
               TabIndex        =   93
               Top             =   4800
               Width           =   2775
            End
            Begin VB.Image RosterimgType1 
               Height          =   375
               Left            =   4440
               Stretch         =   -1  'True
               Top             =   3960
               Width           =   1215
            End
            Begin VB.Image RosterimgType2 
               Height          =   375
               Left            =   5640
               Stretch         =   -1  'True
               Top             =   3960
               Width           =   1215
            End
            Begin VB.Label UselessLabel 
               BackColor       =   &H00404040&
               Caption         =   "EXP:"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   17
               Left            =   4560
               TabIndex        =   92
               Top             =   1680
               Width           =   615
            End
            Begin VB.Label lblEXP 
               BackColor       =   &H00554C42&
               Caption         =   "0/0"
               ForeColor       =   &H00FFFF00&
               Height          =   255
               Left            =   5160
               TabIndex        =   91
               Top             =   1680
               Width           =   1575
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl RosterShinyImage 
               Height          =   1815
               Left            =   4560
               Top             =   2040
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   3201
               Attr            =   1536
               Effects         =   "frmMainGame.frx":2B41AA0
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl RosterimgPokemon 
               Height          =   1815
               Left            =   4560
               Top             =   2040
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   3201
               Attr            =   1536
               Effects         =   "frmMainGame.frx":2B41AB8
            End
         End
      End
      Begin VB.PictureBox picBag 
         Appearance      =   0  'Flat
         BackColor       =   &H00312920&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6615
         Left            =   2640
         ScaleHeight     =   6615
         ScaleWidth      =   8415
         TabIndex        =   120
         Top             =   1560
         Visible         =   0   'False
         Width           =   8415
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   34
            Left            =   3720
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   298
            Top             =   5400
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   34
               Left            =   0
               TabIndex        =   299
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   34
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41AD0
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   33
            Left            =   2880
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   296
            Top             =   5400
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   33
               Left            =   0
               TabIndex        =   297
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   33
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41AE8
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   32
            Left            =   2040
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   294
            Top             =   5400
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   32
               Left            =   0
               TabIndex        =   295
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   32
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41B00
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   31
            Left            =   1200
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   292
            Top             =   5400
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   31
               Left            =   0
               TabIndex        =   293
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   31
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41B18
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   30
            Left            =   360
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   290
            Top             =   5400
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   30
               Left            =   0
               TabIndex        =   291
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   30
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41B30
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   29
            Left            =   3720
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   288
            Top             =   4560
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   29
               Left            =   0
               TabIndex        =   289
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   29
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41B48
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   28
            Left            =   2880
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   286
            Top             =   4560
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   28
               Left            =   0
               TabIndex        =   287
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   28
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41B60
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   27
            Left            =   2040
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   284
            Top             =   4560
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   27
               Left            =   0
               TabIndex        =   285
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   27
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41B78
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   26
            Left            =   1200
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   282
            Top             =   4560
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   26
               Left            =   0
               TabIndex        =   283
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   26
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41B90
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   25
            Left            =   360
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   280
            Top             =   4560
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   25
               Left            =   0
               TabIndex        =   281
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   25
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41BA8
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   24
            Left            =   3720
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   278
            Top             =   3720
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   24
               Left            =   0
               TabIndex        =   279
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   24
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41BC0
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   23
            Left            =   2880
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   276
            Top             =   3720
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   23
               Left            =   0
               TabIndex        =   277
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   23
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41BD8
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   22
            Left            =   2040
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   274
            Top             =   3720
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   22
               Left            =   0
               TabIndex        =   275
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   22
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41BF0
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   21
            Left            =   1200
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   272
            Top             =   3720
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   21
               Left            =   0
               TabIndex        =   273
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   21
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41C08
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   20
            Left            =   360
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   270
            Top             =   3720
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   20
               Left            =   0
               TabIndex        =   271
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   20
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41C20
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   19
            Left            =   3720
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   268
            Top             =   2880
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   19
               Left            =   0
               TabIndex        =   269
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   19
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41C38
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   18
            Left            =   2880
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   266
            Top             =   2880
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   18
               Left            =   0
               TabIndex        =   267
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   18
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41C50
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   17
            Left            =   2040
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   264
            Top             =   2880
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   17
               Left            =   0
               TabIndex        =   265
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   17
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41C68
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   16
            Left            =   1200
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   262
            Top             =   2880
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   16
               Left            =   0
               TabIndex        =   263
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   16
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41C80
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   15
            Left            =   360
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   260
            Top             =   2880
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   15
               Left            =   0
               TabIndex        =   261
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   15
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41C98
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   14
            Left            =   3720
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   258
            Top             =   2040
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   14
               Left            =   0
               TabIndex        =   259
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   14
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41CB0
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   13
            Left            =   2880
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   256
            Top             =   2040
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   13
               Left            =   0
               TabIndex        =   257
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   13
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41CC8
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   12
            Left            =   2040
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   254
            Top             =   2040
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   12
               Left            =   0
               TabIndex        =   255
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   12
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41CE0
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   11
            Left            =   1200
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   252
            Top             =   2040
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   11
               Left            =   0
               TabIndex        =   253
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   11
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41CF8
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   10
            Left            =   360
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   250
            Top             =   2040
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   10
               Left            =   0
               TabIndex        =   251
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   10
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41D10
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   9
            Left            =   3720
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   248
            Top             =   1200
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   9
               Left            =   0
               TabIndex        =   249
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   9
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41D28
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   8
            Left            =   2880
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   246
            Top             =   1200
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   8
               Left            =   0
               TabIndex        =   247
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   8
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41D40
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   7
            Left            =   2040
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   244
            Top             =   1200
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   7
               Left            =   0
               TabIndex        =   245
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   7
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Attr            =   515
               Effects         =   "frmMainGame.frx":2B41D58
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   6
            Left            =   1200
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   242
            Top             =   1200
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   6
               Left            =   0
               TabIndex        =   243
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   6
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41D70
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   5
            Left            =   360
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   240
            Top             =   1200
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   5
               Left            =   0
               TabIndex        =   241
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   5
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41D88
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   4
            Left            =   3720
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   238
            Top             =   360
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   4
               Left            =   0
               TabIndex        =   239
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   4
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41DA0
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   3
            Left            =   2880
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   236
            Top             =   360
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   3
               Left            =   0
               TabIndex        =   237
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   3
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41DB8
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   2
            Left            =   2040
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   234
            Top             =   360
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   2
               Left            =   0
               TabIndex        =   235
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   2
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41DD0
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   1
            Left            =   1200
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   232
            Top             =   360
            Width           =   720
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   1
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41DE8
            End
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   165
               Index           =   1
               Left            =   0
               TabIndex        =   233
               Top             =   480
               Width           =   645
            End
         End
         Begin VB.PictureBox picItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   0
            Left            =   360
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   230
            Top             =   360
            Width           =   720
            Begin VB.Label lblItemVal 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "x456"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   160
               Index           =   0
               Left            =   0
               TabIndex        =   231
               Top             =   480
               Width           =   645
            End
            Begin LaVolpeAlphaImg.AlphaImgCtl imgItemIcon 
               Height          =   480
               Index           =   0
               Left            =   105
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Effects         =   "frmMainGame.frx":2B41E00
            End
         End
         Begin VB.ListBox BaglstItems 
            Appearance      =   0  'Flat
            BackColor       =   &H00312920&
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   360
            TabIndex        =   122
            Top             =   6600
            Visible         =   0   'False
            Width           =   2895
         End
         Begin lvButton.lvButtons_H lvButtons_H26 
            Height          =   375
            Left            =   4920
            TabIndex        =   121
            Top             =   3840
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            Caption         =   "Remove Item"
            CapAlign        =   2
            BackStyle       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            cBack           =   16777215
         End
         Begin lvButton.lvButtons_H lvButtons_H27 
            Height          =   375
            Left            =   4920
            TabIndex        =   123
            Top             =   5280
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            Caption         =   "Unequip items"
            CapAlign        =   2
            BackStyle       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            cBack           =   16777215
         End
         Begin lvButton.lvButtons_H lvButtons_H28 
            Height          =   375
            Left            =   4920
            TabIndex        =   124
            Top             =   5760
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            Caption         =   "Close"
            CapAlign        =   2
            BackStyle       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            cBack           =   16777215
         End
         Begin lvButton.lvButtons_H lvButtons_H7 
            Height          =   375
            Left            =   4920
            TabIndex        =   302
            Top             =   3360
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            Caption         =   "Use Item"
            CapAlign        =   2
            BackStyle       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            cBack           =   16777215
         End
         Begin VB.Label lblItemInfo 
            Alignment       =   2  'Center
            BackColor       =   &H00554C42&
            Caption         =   "x1000"
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
            Height          =   375
            Index           =   1
            Left            =   4920
            TabIndex        =   301
            Top             =   960
            Width           =   3015
         End
         Begin VB.Label lblItemInfo 
            Alignment       =   2  'Center
            BackColor       =   &H00554C42&
            Caption         =   "ItemName"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   0
            Left            =   4920
            TabIndex        =   300
            Top             =   600
            Width           =   3015
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl Bagimgicon 
            Height          =   960
            Left            =   5910
            Top             =   1800
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   1693
            Attr            =   515
            Effects         =   "frmMainGame.frx":2B41E18
         End
      End
      Begin VB.PictureBox picBank 
         BackColor       =   &H00312920&
         BorderStyle     =   0  'None
         Height          =   4455
         Left            =   4920
         ScaleHeight     =   4455
         ScaleWidth      =   4935
         TabIndex        =   131
         Top             =   2520
         Visible         =   0   'False
         Width           =   4935
         Begin lvButton.lvButtons_H cmdMove 
            Height          =   375
            Left            =   480
            TabIndex        =   132
            Top             =   1680
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   661
            Caption         =   "Withdraw 500"
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
         Begin lvButton.lvButtons_H cmdDep 
            Height          =   375
            Left            =   480
            TabIndex        =   133
            Top             =   2160
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   661
            Caption         =   "Deposit 500"
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
         Begin lvButton.lvButtons_H cmdClose 
            Height          =   375
            Left            =   480
            TabIndex        =   134
            Top             =   2760
            Width           =   4095
            _ExtentX        =   7223
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
            LockHover       =   3
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   16777215
         End
         Begin VB.Label lblCPC 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "PokeCoins:0"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   136
            Top             =   240
            Width           =   4815
         End
         Begin VB.Label lblSPC 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Stored PC:0"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   135
            Top             =   840
            Width           =   4695
         End
      End
      Begin VB.PictureBox picShop 
         BackColor       =   &H00312920&
         BorderStyle     =   0  'None
         Height          =   6735
         Left            =   4200
         ScaleHeight     =   6735
         ScaleWidth      =   5175
         TabIndex        =   177
         Top             =   1080
         Visible         =   0   'False
         Width           =   5175
         Begin VB.CommandButton Command8 
            Caption         =   "Buy"
            Height          =   255
            Left            =   6840
            TabIndex        =   183
            Top             =   4080
            Width           =   1935
         End
         Begin VB.TextBox ShoptxtCost 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   182
            Text            =   "Cost: 0 PokeCoins"
            Top             =   3360
            Width           =   2175
         End
         Begin VB.ListBox ShoplstMyItems 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H00FFFFFF&
            Height          =   4710
            Left            =   120
            TabIndex        =   181
            Top             =   0
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Sell"
            Height          =   255
            Left            =   6960
            TabIndex        =   180
            Top             =   5040
            Width           =   1935
         End
         Begin VB.TextBox ShoptxtPrice 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   179
            Text            =   "Price: 0 PokeCoins"
            Top             =   3360
            Visible         =   0   'False
            Width           =   2175
         End
         Begin lvButton.lvButtons_H lvButtons_H36 
            Height          =   375
            Left            =   600
            TabIndex        =   178
            Top             =   6000
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   661
            Caption         =   "Close"
            CapAlign        =   2
            BackStyle       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            cBack           =   16777215
         End
         Begin lvButton.lvButtons_H lvButtons_H37 
            Height          =   375
            Left            =   2880
            TabIndex        =   185
            Top             =   3720
            Visible         =   0   'False
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            Caption         =   "Sell"
            CapAlign        =   2
            BackStyle       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            cBack           =   16777215
         End
         Begin lvButton.lvButtons_H lvButtons_H38 
            Height          =   375
            Left            =   2880
            TabIndex        =   186
            Top             =   3720
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            Caption         =   "Buy"
            CapAlign        =   2
            BackStyle       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            cBack           =   16777215
         End
         Begin lvButton.lvButtons_H lvButtons_H5 
            Height          =   375
            Left            =   2760
            TabIndex        =   334
            Top             =   0
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            Caption         =   "Buy items"
            CapAlign        =   2
            BackStyle       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            cBack           =   16777215
         End
         Begin lvButton.lvButtons_H lvButtons_H8 
            Height          =   375
            Left            =   2760
            TabIndex        =   335
            Top             =   480
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            Caption         =   "Sell items"
            CapAlign        =   2
            BackStyle       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            cBack           =   16777215
         End
         Begin VB.ListBox ShoplstItems 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            ForeColor       =   &H00FFFFFF&
            Height          =   4710
            Left            =   120
            TabIndex        =   184
            Top             =   0
            Width           =   2535
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl imgShopItem 
            Height          =   1575
            Left            =   3120
            Top             =   1560
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   2778
            Attr            =   515
            Effects         =   "frmMainGame.frx":2B41E30
         End
      End
      Begin VB.PictureBox picMenuOptions 
         BackColor       =   &H00312920&
         BorderStyle     =   0  'None
         Height          =   4695
         Left            =   3720
         ScaleHeight     =   4695
         ScaleWidth      =   5055
         TabIndex        =   30
         Top             =   2280
         Visible         =   0   'False
         Width           =   5055
         Begin VB.PictureBox pnlOptions 
            BackColor       =   &H00312920&
            BorderStyle     =   0  'None
            Height          =   3255
            Left            =   960
            ScaleHeight     =   3255
            ScaleWidth      =   3255
            TabIndex        =   31
            Top             =   720
            Width           =   3255
            Begin VB.CheckBox optCheck 
               BackColor       =   &H00312920&
               Caption         =   "Show nearby maps"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   8.25
                  Charset         =   238
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   200
               Top             =   1920
               Width           =   1935
            End
            Begin VB.CheckBox optCheck 
               BackColor       =   &H00312920&
               Caption         =   "Repeat Map Music"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   8.25
                  Charset         =   238
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   36
               Top             =   480
               Width           =   1935
            End
            Begin VB.CheckBox optCheck 
               BackColor       =   &H00312920&
               Caption         =   "Play Audio"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   8.25
                  Charset         =   238
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   35
               Top             =   120
               Width           =   1935
            End
            Begin VB.CheckBox optCheck 
               BackColor       =   &H00312920&
               Caption         =   "Camera Follow Player"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   8.25
                  Charset         =   238
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   34
               Top             =   840
               Width           =   1935
            End
            Begin VB.CheckBox optCheck 
               BackColor       =   &H00312920&
               Caption         =   "Form transparency"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   8.25
                  Charset         =   238
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   33
               Top             =   1200
               Width           =   1935
            End
            Begin VB.CheckBox optCheck 
               BackColor       =   &H00312920&
               Caption         =   "Play Radio"
               BeginProperty Font 
                  Name            =   "Eurostar"
                  Size            =   8.25
                  Charset         =   238
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   32
               Top             =   1560
               Width           =   1935
            End
            Begin lvButton.lvButtons_H lvButtons_H9 
               Height          =   375
               Left            =   480
               TabIndex        =   37
               Top             =   2520
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   661
               Caption         =   "Save"
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
               cGradient       =   0
               Mode            =   0
               Value           =   0   'False
               cBack           =   16777215
            End
         End
         Begin VB.Label UselessLabel 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Options"
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
            Index           =   18
            Left            =   960
            TabIndex        =   38
            Top             =   120
            Width           =   3015
         End
      End
      Begin VB.PictureBox picTPRemove 
         BackColor       =   &H00312920&
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   4080
         ScaleHeight     =   2295
         ScaleWidth      =   5175
         TabIndex        =   188
         Top             =   3240
         Visible         =   0   'False
         Width           =   5175
         Begin lvButton.lvButtons_H btnStat 
            Height          =   375
            Index           =   0
            Left            =   600
            TabIndex        =   190
            Top             =   720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Caption         =   "HP"
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
         Begin lvButton.lvButtons_H btnStat 
            Height          =   375
            Index           =   1
            Left            =   1920
            TabIndex        =   191
            Top             =   720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Caption         =   "ATK"
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
         Begin lvButton.lvButtons_H btnStat 
            Height          =   375
            Index           =   2
            Left            =   3240
            TabIndex        =   192
            Top             =   720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Caption         =   "DEF"
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
         Begin lvButton.lvButtons_H btnStat 
            Height          =   375
            Index           =   3
            Left            =   600
            TabIndex        =   193
            Top             =   1320
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Caption         =   "SP.ATK"
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
         Begin lvButton.lvButtons_H btnStat 
            Height          =   375
            Index           =   4
            Left            =   1920
            TabIndex        =   194
            Top             =   1320
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Caption         =   "SP.DEF"
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
         Begin lvButton.lvButtons_H btnStat 
            Height          =   375
            Index           =   5
            Left            =   3240
            TabIndex        =   195
            Top             =   1320
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Caption         =   "SPEED"
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
         Begin VB.Label UselessLabel 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Choose a stat to remove"
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
            Index           =   23
            Left            =   -120
            TabIndex        =   189
            Top             =   0
            Width           =   5295
         End
      End
      Begin VB.PictureBox picTrade 
         BackColor       =   &H00312920&
         BorderStyle     =   0  'None
         Height          =   5160
         Left            =   3840
         ScaleHeight     =   344
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   385
         TabIndex        =   148
         Top             =   2040
         Visible         =   0   'False
         Width           =   5775
         Begin VB.PictureBox picTradeValue 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   4935
            Left            =   120
            Picture         =   "frmMainGame.frx":2B41E48
            ScaleHeight     =   4905
            ScaleWidth      =   5505
            TabIndex        =   149
            Top             =   120
            Visible         =   0   'False
            Width           =   5535
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               TabIndex        =   151
               Text            =   "0"
               Top             =   1800
               Width           =   5175
            End
            Begin lvButton.lvButtons_H lvButtons_H14 
               Height          =   375
               Left            =   2040
               TabIndex        =   150
               Top             =   2880
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               Caption         =   "Ok"
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
            Begin VB.Label lblVal 
               BackStyle       =   0  'Transparent
               Caption         =   "How many PokeCoins do you want to trade?"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   120
               TabIndex        =   152
               Top             =   1440
               Width           =   5175
            End
         End
         Begin VB.TextBox txtPoke 
            BackColor       =   &H00554C42&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   157
            Top             =   2280
            Width           =   1935
         End
         Begin VB.TextBox txtItem 
            BackColor       =   &H00554C42&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   156
            Top             =   4680
            Width           =   1935
         End
         Begin VB.ComboBox cmbItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            Height          =   315
            Left            =   360
            TabIndex        =   154
            Top             =   2280
            Width           =   1935
         End
         Begin VB.ComboBox cmbPoke 
            Appearance      =   0  'Flat
            BackColor       =   &H00554C42&
            Height          =   315
            Left            =   360
            TabIndex        =   153
            Top             =   1560
            Width           =   1935
         End
         Begin lvButton.lvButtons_H lvButtons_H12 
            Height          =   375
            Left            =   240
            TabIndex        =   155
            Top             =   3960
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
            Caption         =   "Accept"
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
            cBhover         =   14737632
            LockHover       =   1
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   16777215
         End
         Begin lvButton.lvButtons_H lvButtons_H13 
            Height          =   375
            Left            =   240
            TabIndex        =   158
            Top             =   4440
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
            Caption         =   "Cancel"
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
            cBhover         =   14737632
            LockHover       =   1
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   16777215
         End
         Begin VB.Label UselessLabel 
            Alignment       =   2  'Center
            BackColor       =   &H00554C42&
            Caption         =   "Me"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   27
            Left            =   360
            TabIndex        =   171
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label lblTradeName 
            Alignment       =   2  'Center
            BackColor       =   &H00554C42&
            Caption         =   "Goran"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   3240
            TabIndex        =   170
            Top             =   600
            Width           =   1935
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl picTradePoke 
            Height          =   1215
            Left            =   3240
            Top             =   960
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   2143
            Effects         =   "frmMainGame.frx":2BB718A
         End
         Begin VB.Label Label29 
            Alignment       =   2  'Center
            BackColor       =   &H00554C42&
            Caption         =   "Item:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3240
            TabIndex        =   169
            Top             =   4440
            Width           =   1935
         End
         Begin VB.Label UselessLabel 
            Alignment       =   2  'Center
            BackColor       =   &H00554C42&
            Caption         =   "Item:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   29
            Left            =   360
            TabIndex        =   168
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label UselessLabel 
            Alignment       =   2  'Center
            BackColor       =   &H00554C42&
            Caption         =   "Pokemon:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   28
            Left            =   360
            TabIndex        =   167
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label lvlTradeVal 
            BackColor       =   &H00FFFFFF&
            Caption         =   "0"
            Height          =   255
            Left            =   360
            TabIndex        =   166
            Top             =   2760
            Width           =   1935
         End
         Begin VB.Label lblTradeNature 
            BackColor       =   &H00554C42&
            Caption         =   "Nature: None"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3240
            TabIndex        =   165
            Top             =   2640
            Width           =   1935
         End
         Begin VB.Label lblTradeHp 
            BackColor       =   &H00554C42&
            Caption         =   "Hp: 0"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3240
            TabIndex        =   164
            Top             =   2880
            Width           =   1935
         End
         Begin VB.Label lblTradeAtk 
            BackColor       =   &H00554C42&
            Caption         =   "Atk: 0"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3240
            TabIndex        =   163
            Top             =   3120
            Width           =   1935
         End
         Begin VB.Label lblTradeDef 
            BackColor       =   &H00554C42&
            Caption         =   "Def: 0"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3240
            TabIndex        =   162
            Top             =   3360
            Width           =   1935
         End
         Begin VB.Label lblTradeSpAtk 
            BackColor       =   &H00554C42&
            Caption         =   "Sp.Atk: 0"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3240
            TabIndex        =   161
            Top             =   3600
            Width           =   1935
         End
         Begin VB.Label lblTradeSpDef 
            BackColor       =   &H00554C42&
            Caption         =   "Sp.Def: 0"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3240
            TabIndex        =   160
            Top             =   3840
            Width           =   1935
         End
         Begin VB.Label lblTradeSpeed 
            BackColor       =   &H00554C42&
            Caption         =   "Speed: 0"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3240
            TabIndex        =   159
            Top             =   4080
            Width           =   1935
         End
      End
      Begin VB.PictureBox picEvolve 
         BackColor       =   &H00312920&
         BorderStyle     =   0  'None
         Height          =   4455
         Left            =   3720
         ScaleHeight     =   4455
         ScaleWidth      =   6735
         TabIndex        =   137
         Top             =   1800
         Visible         =   0   'False
         Width           =   6735
         Begin lvButton.lvButtons_H lvButtons_H31 
            Height          =   375
            Left            =   360
            TabIndex        =   138
            Top             =   3600
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   661
            Caption         =   "Evolve"
            CapAlign        =   2
            BackStyle       =   5
            Shape           =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFore           =   0
            cFHover         =   0
            cBhover         =   14737632
            LockHover       =   1
            cGradient       =   0
            CapStyle        =   2
            Mode            =   0
            Value           =   0   'False
            cBack           =   16777215
         End
         Begin lvButton.lvButtons_H lvButtons_H32 
            Height          =   375
            Left            =   4560
            TabIndex        =   139
            Top             =   3600
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Caption         =   "Close"
            CapAlign        =   2
            BackStyle       =   5
            Shape           =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFore           =   16777215
            cFHover         =   16777215
            cBhover         =   8421631
            LockHover       =   1
            cGradient       =   0
            CapStyle        =   2
            Mode            =   0
            Value           =   0   'False
            cBack           =   255
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl imgOldPoke 
            Height          =   2775
            Left            =   480
            Top             =   240
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   4895
            Effects         =   "frmMainGame.frx":2BB71A2
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl imgNewPoke 
            Height          =   2775
            Left            =   3360
            Top             =   240
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   4895
            Effects         =   "frmMainGame.frx":2BB71BA
         End
         Begin VB.Label UselessLabel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "TO"
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
            Index           =   30
            Left            =   2760
            TabIndex        =   140
            Top             =   1440
            Width           =   1215
         End
      End
      Begin VB.PictureBox picLearnMove 
         BackColor       =   &H00312920&
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   4080
         ScaleHeight     =   3135
         ScaleWidth      =   5655
         TabIndex        =   141
         Top             =   2880
         Visible         =   0   'False
         Width           =   5655
         Begin lvButton.lvButtons_H btnMove 
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   142
            Top             =   1440
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
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
         Begin lvButton.lvButtons_H btnMove 
            Height          =   375
            Index           =   2
            Left            =   1560
            TabIndex        =   143
            Top             =   1440
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
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
         Begin lvButton.lvButtons_H btnMove 
            Height          =   375
            Index           =   3
            Left            =   2880
            TabIndex        =   144
            Top             =   1440
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
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
         Begin lvButton.lvButtons_H btnMove 
            Height          =   375
            Index           =   4
            Left            =   4200
            TabIndex        =   145
            Top             =   1440
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
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
         Begin lvButton.lvButtons_H btnDONT 
            Height          =   375
            Left            =   1560
            TabIndex        =   146
            Top             =   1920
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   661
            Caption         =   "Don't learn"
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
         Begin VB.Label lblLearnMove 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Bulbasaur wants to learn Razor Leaf.Choose a move to replace."
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   147
            Top             =   1080
            Width           =   5535
         End
      End
      Begin VB.PictureBox picProfile 
         BackColor       =   &H00312920&
         BorderStyle     =   0  'None
         Height          =   7095
         Left            =   2400
         ScaleHeight     =   7095
         ScaleWidth      =   5895
         TabIndex        =   126
         Top             =   720
         Visible         =   0   'False
         Width           =   5895
         Begin VB.TextBox txtProfileImg 
            BackColor       =   &H00554C42&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   480
            TabIndex        =   127
            Text            =   "Image Link"
            Top             =   840
            Width           =   1575
         End
         Begin lvButton.lvButtons_H lvButtons_H29 
            Height          =   615
            Left            =   480
            TabIndex        =   128
            Top             =   1200
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   1085
            Caption         =   "Set profile picture"
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
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H lvButtons_H30 
            Height          =   375
            Left            =   840
            TabIndex        =   129
            Top             =   6120
            Width           =   4335
            _ExtentX        =   7646
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
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   16777215
         End
         Begin VB.Label lblPlayerInfo 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Hours:3 , Minutes: 18"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Index           =   4
            Left            =   3120
            TabIndex        =   311
            Top             =   5160
            Width           =   2295
         End
         Begin VB.Label lblPlayerInfo 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "400 Days"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   3480
            TabIndex        =   310
            Top             =   4800
            Width           =   1455
         End
         Begin VB.Label lblPlayerInfo 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Bronze 2"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   3480
            TabIndex        =   309
            Top             =   4440
            Width           =   1455
         End
         Begin VB.Label lblPlayerInfo 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1000"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   3480
            TabIndex        =   308
            Top             =   4080
            Width           =   1455
         End
         Begin VB.Label UselessLabel 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Play time:"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   40
            Left            =   2280
            TabIndex        =   307
            Top             =   5160
            Width           =   1215
         End
         Begin VB.Label UselessLabel 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Membership:"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   39
            Left            =   2280
            TabIndex        =   306
            Top             =   4800
            Width           =   1215
         End
         Begin VB.Label UselessLabel 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Division:"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   38
            Left            =   2280
            TabIndex        =   305
            Top             =   4440
            Width           =   1215
         End
         Begin VB.Label UselessLabel 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Ranked Points:"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   37
            Left            =   2280
            TabIndex        =   304
            Top             =   4080
            Width           =   1215
         End
         Begin VB.Label UselessLabel 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Player information:"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   11.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   36
            Left            =   2400
            TabIndex        =   303
            Top             =   3720
            Width           =   2895
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl imgProfileDivision 
            Height          =   1695
            Left            =   480
            Top             =   3720
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   2990
            Frame           =   12
            BackColor       =   5590082
            Attr            =   515
            Effects         =   "frmMainGame.frx":2BB71D2
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl imgClothes 
            Height          =   735
            Index           =   1
            Left            =   3720
            Top             =   2040
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   1296
            Trans           =   33554432
            Frame           =   12
            BackColor       =   5590082
            Effects         =   "frmMainGame.frx":2BB71EA
         End
         Begin VB.Label UselessLabel 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "(Double click to remove)"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   34
            Left            =   2640
            TabIndex        =   228
            Top             =   840
            Width           =   3135
         End
         Begin VB.Label UselessLabel 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Clothes"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   11.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   33
            Left            =   2640
            TabIndex        =   227
            Top             =   600
            Width           =   3135
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl imgClothes 
            Height          =   735
            Index           =   5
            Left            =   2640
            Top             =   1680
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   1296
            Trans           =   33554432
            Frame           =   12
            BackColor       =   5590082
            Effects         =   "frmMainGame.frx":2BB7202
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl imgClothes 
            Height          =   735
            Index           =   4
            Left            =   4560
            Top             =   1200
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   1296
            Trans           =   33554432
            Frame           =   12
            BackColor       =   5590082
            Effects         =   "frmMainGame.frx":2BB721A
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl imgClothes 
            Height          =   735
            Index           =   3
            Left            =   3720
            Top             =   1200
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   1296
            Trans           =   33554432
            Frame           =   12
            BackColor       =   5590082
            Effects         =   "frmMainGame.frx":2BB7232
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl imgClothes 
            Height          =   735
            Index           =   2
            Left            =   4560
            Top             =   2040
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   1296
            Trans           =   33554432
            Frame           =   12
            BackColor       =   5590082
            Effects         =   "frmMainGame.frx":2BB724A
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl imgClothes 
            Height          =   735
            Index           =   0
            Left            =   4185
            Top             =   2880
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   1296
            Trans           =   33554432
            Frame           =   12
            BackColor       =   5590082
            Effects         =   "frmMainGame.frx":2BB7262
         End
         Begin VB.Label UselessLabel 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Player"
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
            Index           =   1
            Left            =   240
            TabIndex        =   130
            Top             =   0
            Width           =   5295
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl imgProfile 
            Height          =   1335
            Left            =   480
            Top             =   1920
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   2355
            Frame           =   12
            BackColor       =   5590082
            Attr            =   513
            Effects         =   "frmMainGame.frx":2BB727A
         End
      End
      Begin VB.PictureBox picTrainerCard 
         BackColor       =   &H00312920&
         BorderStyle     =   0  'None
         Height          =   9375
         Left            =   2280
         Picture         =   "frmMainGame.frx":2BB7292
         ScaleHeight     =   625
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   625
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   9375
         Begin VB.PictureBox picPlayerJournal 
            BackColor       =   &H00312920&
            BorderStyle     =   0  'None
            Height          =   3615
            Left            =   1800
            ScaleHeight     =   3615
            ScaleWidth      =   5055
            TabIndex        =   213
            Top             =   2760
            Visible         =   0   'False
            Width           =   5055
            Begin RichTextLib.RichTextBox txtJournal 
               Height          =   2655
               Left            =   0
               TabIndex        =   216
               Top             =   360
               Visible         =   0   'False
               Width           =   5055
               _ExtentX        =   8916
               _ExtentY        =   4683
               _Version        =   393217
               BackColor       =   0
               Enabled         =   -1  'True
               ReadOnly        =   -1  'True
               ScrollBars      =   3
               Appearance      =   0
               TextRTF         =   $"frmMainGame.frx":2C0A598
            End
            Begin vb6projectRichtextEditor.ucRTBEditor txtJournalEdit 
               Height          =   3015
               Left            =   0
               TabIndex        =   214
               Top             =   0
               Visible         =   0   'False
               Width           =   5055
               _ExtentX        =   8916
               _ExtentY        =   5318
               TextRTF         =   $"frmMainGame.frx":2C0A61D
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin lvButton.lvButtons_H cmdSaveJournal 
               Height          =   375
               Left            =   3720
               TabIndex        =   215
               Top             =   3120
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               Caption         =   "Edit"
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
               cFHover         =   4210752
               cBhover         =   14737632
               LockHover       =   3
               cGradient       =   0
               Mode            =   0
               Value           =   0   'False
               cBack           =   16777215
            End
            Begin lvButton.lvButtons_H lvButtons_H4 
               Height          =   375
               Left            =   2520
               TabIndex        =   217
               Top             =   3120
               Width           =   1095
               _ExtentX        =   1931
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
               cFHover         =   4210752
               cBhover         =   14737632
               LockHover       =   3
               cGradient       =   0
               Mode            =   0
               Value           =   0   'False
               cBack           =   16777215
            End
         End
         Begin lvButton.lvButtons_H lvButtons_H6 
            Height          =   375
            Left            =   5520
            TabIndex        =   8
            Top             =   7320
            Width           =   975
            _ExtentX        =   1720
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
            cFHover         =   4210752
            cBhover         =   14737632
            LockHover       =   3
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   16777215
         End
         Begin lvButton.lvButtons_H lvButtons_H11 
            Height          =   375
            Left            =   4320
            TabIndex        =   9
            Top             =   6840
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            Caption         =   "Trade"
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
            cFHover         =   4210752
            cBhover         =   14737632
            LockHover       =   3
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   16777215
         End
         Begin lvButton.lvButtons_H lvButtons_H1 
            Height          =   375
            Left            =   4320
            TabIndex        =   196
            Top             =   7320
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            Caption         =   "Battle"
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
            cFHover         =   4210752
            cBhover         =   14737632
            LockHover       =   3
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   16777215
         End
         Begin lvButton.lvButtons_H cmdJournal 
            Height          =   375
            Left            =   5520
            TabIndex        =   212
            Top             =   6840
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            Caption         =   "Journal"
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
            cFHover         =   4210752
            cBhover         =   14737632
            LockHover       =   3
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   16777215
         End
         Begin lvButton.lvButtons_H cmdClanInvite 
            Height          =   375
            Left            =   2280
            TabIndex        =   218
            Top             =   7080
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            Caption         =   "Invite to clan"
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
            cFHover         =   4210752
            cBhover         =   14737632
            LockHover       =   3
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   16777215
         End
         Begin VB.Label UselessLabel 
            Alignment       =   2  'Center
            BackColor       =   &H00554C42&
            Caption         =   "Points:"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   26
            Left            =   1800
            TabIndex        =   198
            Top             =   8280
            Width           =   975
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl imgPlayerClan 
            Height          =   975
            Left            =   1800
            Top             =   6840
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   1720
            Image           =   "frmMainGame.frx":2C0A6AC
            Frame           =   12
            BackColor       =   5590082
            Attr            =   515
            Effects         =   "frmMainGame.frx":2C1B5D0
         End
         Begin VB.Label lblPlayerCrew 
            Alignment       =   2  'Center
            BackColor       =   &H00554C42&
            Caption         =   "Crew"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1800
            TabIndex        =   201
            Top             =   6600
            Width           =   2295
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl imgRank 
            Height          =   855
            Left            =   1800
            Top             =   8520
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   1508
            Image           =   "frmMainGame.frx":2C1B5E8
            Frame           =   12
            BackColor       =   5590082
            Attr            =   515
            Effects         =   "frmMainGame.frx":2C2C50C
         End
         Begin VB.Label lblRankPoints 
            Alignment       =   2  'Center
            BackColor       =   &H00554C42&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1800
            TabIndex        =   199
            Top             =   8280
            Width           =   4695
         End
         Begin VB.Label UselessLabel 
            Alignment       =   2  'Center
            BackColor       =   &H00554C42&
            Caption         =   "Ranked"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   24
            Left            =   1800
            TabIndex        =   197
            Top             =   8040
            Width           =   4695
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl imgProfilePic 
            Height          =   1215
            Left            =   2280
            Top             =   1080
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   2143
            Frame           =   12
            BackColor       =   5590082
            Attr            =   515
            Effects         =   "frmMainGame.frx":2C2C524
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl GameMaster 
            Height          =   615
            Left            =   3840
            Top             =   4440
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   1085
            Attr            =   514
            Effects         =   "frmMainGame.frx":2C2C53C
         End
         Begin VB.Label lblName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "PlayerName"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   -960
            TabIndex        =   17
            Top             =   600
            Width           =   10575
         End
         Begin VB.Label lblCharPowerLvl 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Power Lvl:0"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   3360
            TabIndex        =   16
            Top             =   2520
            Width           =   1575
         End
         Begin VB.Label lblCharPkmnLvl 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Lvl:100"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   135
            Index           =   5
            Left            =   3480
            TabIndex        =   15
            Top             =   4800
            Width           =   495
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl imgBadge 
            Height          =   480
            Index           =   1
            Left            =   4680
            Top             =   1200
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
            Frame           =   12
            BackColor       =   5590082
            Effects         =   "frmMainGame.frx":2C2C554
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl imgBadge 
            Height          =   480
            Index           =   2
            Left            =   5400
            Top             =   1200
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
            Frame           =   12
            BackColor       =   5590082
            Effects         =   "frmMainGame.frx":2C2C56C
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl imgBadge 
            Height          =   480
            Index           =   3
            Left            =   6120
            Top             =   1200
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
            Frame           =   12
            BackColor       =   5590082
            Effects         =   "frmMainGame.frx":2C2C584
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl imgBadge 
            Height          =   480
            Index           =   4
            Left            =   4680
            Top             =   1920
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
            Frame           =   12
            BackColor       =   5590082
            Effects         =   "frmMainGame.frx":2C2C59C
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl imgBadge 
            Height          =   480
            Index           =   5
            Left            =   5400
            Top             =   1920
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
            Frame           =   12
            BackColor       =   5590082
            Effects         =   "frmMainGame.frx":2C2C5B4
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl imgBadge 
            Height          =   480
            Index           =   6
            Left            =   6120
            Top             =   1920
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
            Frame           =   12
            BackColor       =   5590082
            Effects         =   "frmMainGame.frx":2C2C5CC
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl imgCharPokemon 
            Height          =   1440
            Index           =   2
            Left            =   3480
            Top             =   3120
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   2540
            Frame           =   12
            BackColor       =   5590082
            Attr            =   1539
            Effects         =   "frmMainGame.frx":2C2C5E4
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl imgCharPokemon 
            Height          =   1440
            Index           =   3
            Left            =   5040
            Top             =   3120
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   2540
            Frame           =   12
            BackColor       =   5590082
            Attr            =   1539
            Effects         =   "frmMainGame.frx":2C2C5FC
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl imgCharPokemon 
            Height          =   1440
            Index           =   6
            Left            =   5040
            Top             =   4920
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   2540
            Frame           =   12
            BackColor       =   5590082
            Attr            =   1539
            Effects         =   "frmMainGame.frx":2C2C614
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl imgCharPokemon 
            Height          =   1440
            Index           =   1
            Left            =   1920
            Top             =   3120
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   2540
            Frame           =   12
            BackColor       =   5590082
            Attr            =   1539
            Effects         =   "frmMainGame.frx":2C2C62C
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl imgCharPokemon 
            Height          =   1440
            Index           =   4
            Left            =   1920
            Top             =   4920
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   2540
            Frame           =   12
            BackColor       =   5590082
            Attr            =   1539
            Effects         =   "frmMainGame.frx":2C2C644
         End
         Begin VB.Label lblCharPkmnLvl 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Lvl:100"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   135
            Index           =   1
            Left            =   1920
            TabIndex        =   14
            Top             =   2880
            Width           =   495
         End
         Begin VB.Label lblCharPkmnLvl 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Lvl:100"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   135
            Index           =   6
            Left            =   5040
            TabIndex        =   13
            Top             =   4800
            Width           =   495
         End
         Begin VB.Label lblCharPkmnLvl 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Lvl:100"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   135
            Index           =   3
            Left            =   5040
            TabIndex        =   12
            Top             =   2880
            Width           =   495
         End
         Begin VB.Label lblCharPkmnLvl 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Lvl:100"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   135
            Index           =   4
            Left            =   1920
            TabIndex        =   11
            Top             =   4800
            Width           =   495
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl imgCharPokemon 
            Height          =   1425
            Index           =   5
            Left            =   3480
            Top             =   4920
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   2514
            Frame           =   12
            BackColor       =   5590082
            Attr            =   1539
            Effects         =   "frmMainGame.frx":2C2C65C
         End
         Begin VB.Label lblCharPkmnLvl 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Lvl:100"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   135
            Index           =   2
            Left            =   3480
            TabIndex        =   10
            Top             =   2880
            Width           =   495
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl9 
            Height          =   1215
            Left            =   2160
            Top             =   1080
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   2143
            Image           =   "frmMainGame.frx":2C2C674
            Attr            =   513
            Effects         =   "frmMainGame.frx":2C2D51F
         End
         Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl7 
            Height          =   1215
            Left            =   2160
            Top             =   1080
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   2143
            Image           =   "frmMainGame.frx":2C2D537
            Attr            =   513
            Effects         =   "frmMainGame.frx":2C2E789
         End
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl4 
         Height          =   5055
         Left            =   3720
         Top             =   4560
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   8916
         Effects         =   "frmMainGame.frx":2C2E7A1
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl imgButton 
         Height          =   600
         Index           =   4
         Left            =   0
         Top             =   3300
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1058
         Effects         =   "frmMainGame.frx":2C2E7B9
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl imgSwitch 
         Height          =   720
         Index           =   4
         Left            =   170
         Top             =   6480
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Frame           =   0
         Attr            =   1536
         Effects         =   "frmMainGame.frx":2C2E7D1
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl imgSwitch 
         Height          =   720
         Index           =   5
         Left            =   170
         Top             =   7320
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Frame           =   0
         Attr            =   1536
         Effects         =   "frmMainGame.frx":2C2E7E9
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl imgSwitch 
         Height          =   720
         Index           =   6
         Left            =   170
         Top             =   8160
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Frame           =   0
         Attr            =   1536
         Effects         =   "frmMainGame.frx":2C2E801
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl imgSwitch 
         Height          =   720
         Index           =   3
         Left            =   170
         Top             =   5640
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Frame           =   0
         Attr            =   1536
         Effects         =   "frmMainGame.frx":2C2E819
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl imgSwitch 
         Height          =   720
         Index           =   2
         Left            =   170
         Top             =   4800
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Frame           =   0
         Attr            =   1536
         Effects         =   "frmMainGame.frx":2C2E831
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl imgSwitch 
         Height          =   720
         Index           =   1
         Left            =   170
         Top             =   3960
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Frame           =   0
         Attr            =   1536
         Effects         =   "frmMainGame.frx":2C2E849
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl6 
         Height          =   540
         Left            =   0
         Top             =   9000
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   953
         Effects         =   "frmMainGame.frx":2C2E861
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl imgButton 
         Height          =   600
         Index           =   3
         Left            =   0
         Top             =   2520
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1058
         Effects         =   "frmMainGame.frx":2C2E879
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl imgButton 
         Height          =   660
         Index           =   2
         Left            =   0
         Top             =   1680
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1164
         Effects         =   "frmMainGame.frx":2C2E891
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl imgButton 
         Height          =   660
         Index           =   1
         Left            =   0
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1164
         Effects         =   "frmMainGame.frx":2C2E8A9
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl imgButton 
         Height          =   615
         Index           =   0
         Left            =   0
         Top             =   60
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
         Effects         =   "frmMainGame.frx":2C2E8C1
      End
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H00312920&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9600
      Left            =   0
      MouseIcon       =   "frmMainGame.frx":2C2E8D9
      MousePointer    =   99  'Custom
      ScaleHeight     =   640
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   704
      TabIndex        =   205
      Top             =   0
      Width           =   10560
      Begin VB.PictureBox picBattleCommands 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   2310
         Left            =   270
         Picture         =   "frmMainGame.frx":2C2EBE3
         ScaleHeight     =   152
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   495
         TabIndex        =   319
         Top             =   5640
         Visible         =   0   'False
         Width           =   7455
         Begin lvButton.lvButtons_H cmdPokeMove 
            Height          =   495
            Index           =   4
            Left            =   5550
            TabIndex        =   320
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
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
         Begin lvButton.lvButtons_H cmdPokeMove 
            Height          =   495
            Index           =   1
            Left            =   150
            TabIndex        =   321
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            Caption         =   "Move1"
            CapAlign        =   2
            BackStyle       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Eurostar"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
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
         Begin lvButton.lvButtons_H cmdPokeMove 
            Height          =   495
            Index           =   2
            Left            =   1950
            TabIndex        =   322
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
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
         Begin lvButton.lvButtons_H cmdPokeMove 
            Height          =   495
            Index           =   3
            Left            =   3750
            TabIndex        =   323
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
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
         Begin lvButton.lvButtons_H cmdBag 
            Height          =   495
            Left            =   150
            TabIndex        =   324
            Top             =   1200
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
            ImgSize         =   24
            cBack           =   16777215
         End
         Begin lvButton.lvButtons_H cmdRun 
            Height          =   495
            Left            =   1440
            TabIndex        =   325
            Top             =   1200
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
            LockHover       =   3
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   16777215
         End
         Begin lvButton.lvButtons_H cmdAutoClose 
            Height          =   255
            Left            =   6240
            TabIndex        =   326
            Top             =   1920
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
         Begin lvButton.lvButtons_H btnCloseBattle 
            Height          =   495
            Left            =   120
            TabIndex        =   329
            Top             =   960
            Visible         =   0   'False
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   873
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
         Begin VB.Label UselessLabel 
            Alignment       =   2  'Center
            BackColor       =   &H00554C42&
            Caption         =   "Exp"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   45
            Left            =   2880
            TabIndex        =   333
            Top             =   1200
            Width           =   4335
         End
         Begin VB.Label lblBattleEXP 
            Alignment       =   2  'Center
            BackColor       =   &H00554C42&
            Caption         =   "0/0"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2880
            TabIndex        =   332
            Top             =   1440
            Width           =   4335
         End
         Begin VB.Label UselessLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "To change pokemon click on its icon in the menu."
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   43
            Left            =   120
            TabIndex        =   331
            Top             =   2040
            Width           =   4455
         End
         Begin VB.Label UselessLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Choose what to do:"
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   35
            Left            =   120
            TabIndex        =   330
            Top             =   120
            Width           =   3135
         End
      End
      Begin VB.ListBox listBag 
         Appearance      =   0  'Flat
         BackColor       =   &H00312920&
         ForeColor       =   &H00FFFFFF&
         Height          =   4125
         Left            =   2580
         TabIndex        =   327
         Top             =   840
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.PictureBox picPokemonImage 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   0
         Left            =   10560
         Picture         =   "frmMainGame.frx":2C66DAF
         ScaleHeight     =   975
         ScaleWidth      =   1335
         TabIndex        =   206
         Top             =   4920
         Width           =   1335
         Begin VB.Image imgPokemon 
            Height          =   960
            Index           =   0
            Left            =   360
            Top             =   120
            Visible         =   0   'False
            Width           =   1320
         End
      End
      Begin RichTextLib.RichTextBox txtBtlLog 
         Height          =   6615
         Left            =   7725
         TabIndex        =   328
         Top             =   1335
         Visible         =   0   'False
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   11668
         _Version        =   393217
         BackColor       =   3221792
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmMainGame.frx":2C6B1FD
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
      Begin LaVolpeAlphaImg.AlphaImgCtl btnImgMoves 
         Height          =   780
         Index           =   3
         Left            =   360
         Top             =   4800
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   1376
         Image           =   "frmMainGame.frx":2C6B279
         Attr            =   513
         Effects         =   "frmMainGame.frx":2DD050F
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl btnImgMoves 
         Height          =   780
         Index           =   2
         Left            =   1800
         Top             =   4800
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   1376
         Image           =   "frmMainGame.frx":2DD0527
         Attr            =   513
         Effects         =   "frmMainGame.frx":2F357BD
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl btnImgMoves 
         Height          =   780
         Index           =   1
         Left            =   3240
         Top             =   4800
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   1376
         Image           =   "frmMainGame.frx":2F357D5
         Attr            =   513
         Effects         =   "frmMainGame.frx":309AA6B
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl btnImgMoves 
         Height          =   780
         Index           =   0
         Left            =   4680
         Top             =   4800
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   1376
         Image           =   "frmMainGame.frx":309AA83
         Attr            =   513
         Effects         =   "frmMainGame.frx":31FFD19
      End
   End
   Begin VB.PictureBox picNews 
      Appearance      =   0  'Flat
      BackColor       =   &H00312920&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9615
      Left            =   0
      ScaleHeight     =   9615
      ScaleWidth      =   10575
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   10575
      Begin RichTextLib.RichTextBox txtGameNews 
         Height          =   8535
         Left            =   240
         TabIndex        =   224
         Top             =   360
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   15055
         _Version        =   393217
         BackColor       =   3221792
         Appearance      =   0
         TextRTF         =   $"frmMainGame.frx":31FFD31
      End
      Begin lvButton.lvButtons_H lvButtons_H3 
         Height          =   375
         Left            =   9000
         TabIndex        =   5
         Top             =   9000
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
   End
   Begin VB.Label UselessLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Poketopia [Revival]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   50
      Left            =   0
      TabIndex        =   339
      Top             =   0
      Width           =   13575
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgPlayerClothe 
      Height          =   615
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      Frame           =   12
      BackColor       =   5590082
      Effects         =   "frmMainGame.frx":31FFDB6
   End
End
Attribute VB_Name = "frmMainGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ************
' ** Events **
' ************
Private MoveForm As Boolean
Private MouseX As Long
Private MouseY As Long
Private PresentX As Long
Private PresentY As Long
Public isClosed As Boolean
Public menuLeft As Boolean
Private evPokesloT As Long
Private evNewPoke As Long
Dim pSlot As Long
Dim pMove As Long
Dim pMoveSlot As Long
Dim selectedItemSlot As Long

Dim DialogText As String
Dim DialogPosition As Long
Dim DialogLenght As Long

'CONSTS
'MENU


'

Private Sub btnAdminPnl_Click()
 If Player(MyIndex).Access > 1 Then
   frmAdmin.Visible = Not frmAdmin.Visible
End If
End Sub

Private Sub AlphaImgCtl13_Click(Index As Integer)
loginDetailsPicture.Visible = False
Picture1.Visible = False
Select Case Index
Case 0
frmMenu.txtLUser.text = txtLUser.text
frmMenu.txtLPass.text = txtLPass.text
PlayClick
Call MenuState(MENU_STATE_LOGIN)
Case 2
PlayClick
If AdminOnly = False Then
DestroyTCP
    
    If Picture1.Visible = False Then
    Picture1.Visible = True
    picNewChar.Visible = False

    Else
    picNewChar.Visible = False
    End If
    Else
    MsgBox ("Server is closed for players!")
    End If
Case 3
PlayClick
Call DestroyGame
Case 4
PlayClick
If StarterChoosed = 0 Then
MsgBox "Please choose your starter pokemon!"
Else
Call MenuState(MENU_STATE_ADDCHAR)
End If
Case 7
PlayClick
Picture1.Visible = False
loginDetailsPicture.Visible = True
Case 6

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
End Select
End Sub

Private Sub AlphaImgCtl2_Click()
PlaySound ("130Cry.wav")

End Sub

Private Sub AlphaImgCtl3_Click()
PlayClick
SendRosterRequest
End Sub

Private Sub AlphaImgCtl4_Click()
PlayClick
 SendRequestPlayerData
OpenMenu (MENU_PROFILE)
  ' picTrainerCard.Visible = True
   'loadPlayerData (MyIndex)
End Sub

Private Sub AlphaImgCtl6_Click()
If inBattle Then Exit Sub
PlayClick
Dim Buffer As clsBuffer
    Dim i As Long
    
   
    InGame = False
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CQuit
    Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Call DestroyTCP
    
    ' destroy the animations loaded
    For i = 1 To MAX_BYTE
        ClearAnimInstance (i)
    Next
    
    ' destroy temp values
     DragInvSlotNum = 0
     InvX = 0
     InvY = 0
     EqX = 0
     EqY = 0
     SpellX = 0
     SpellY = 0
     LastItemDesc = 0
     MyIndex = 0
     InventoryItemSelected = 0
     SpellBuffer = 0
     SpellBufferTimer = 0
     tmpDropItem = 0
    
    frmChat.txtChat.text = vbNullString
    picTrade.Visible = False
    Call DestroyGame
End Sub

Private Sub AlphaImgCtl7_Click()
Dim ChatText As String
Dim Name As String
 ChatText = Trim$(MyText)
Dim i As Long, n As Long
Dim Buffer As clsBuffer
Dim Command() As String
    If LenB(ChatText) = 0 Then Exit Sub
    MyText = LCase$(ChatText)
        ' Broadcast message
        If Left$(ChatText, 1) = "'" Then
            ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)

            If Len(ChatText) > 0 Then
                Call BroadcastMsg(ChatText)
            End If

            MyText = vbNullString
            frmChat.txtMyChat.text = vbNullString
            Exit Sub
        End If

        ' Emote message
        If Left$(ChatText, 1) = "-" Then
            MyText = Mid$(ChatText, 2, Len(ChatText) - 1)

            If Len(ChatText) > 0 Then
                Call EmoteMsg(ChatText)
            End If

            MyText = vbNullString
            frmChat.txtMyChat.text = vbNullString
            Exit Sub
        End If

        ' Player message
        If Left$(ChatText, 1) = "!" Then
            ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)
            Name = vbNullString

            ' Get the desired player from the user text
            For i = 1 To Len(ChatText)

                If Mid$(ChatText, i, 1) <> Space(1) Then
                    Name = Name & Mid$(ChatText, i, 1)
                Else
                    Exit For
                End If

            Next

            ChatText = Mid$(ChatText, i, Len(ChatText) - 1)

            ' Make sure they are actually sending something
            If Len(ChatText) - i > 0 Then
                MyText = Mid$(ChatText, i + 1, Len(ChatText) - i)
                ' Send the message to the player
                Call PlayerMsg(ChatText, Name)
            Else
                Call AddText("Usage: !playername (message)", AlertColor)
            End If

            MyText = vbNullString
            frmChat.txtMyChat.text = vbNullString
            Exit Sub
        End If

        If Left$(MyText, 1) = "/" Then
            Command = Split(MyText, Space(1))

            Select Case Command(0)
            Case "/intro"
            frmMainGame.Visible = False
            frmChat.Visible = False
            frmIntro.Show
            
            
              Case "/moodhappy"
              SendMood 0
              AddText "Mood set to happy!", BrightBlue
              Case "/moodsad"
              AddText "Mood set to sad!", BrightBlue
                SendMood 1
                Case "/help"
                    Call AddText("Social Commands:", HelpColor)
                    Call AddText("'msghere = Broadcast Message", HelpColor)
                    Call AddText("-msghere = Emote Message", HelpColor)
                    Call AddText("!namehere msghere = Player Message", HelpColor)
                    Call AddText("Available Commands: /help, /info, /who, /fps, /stats, /trade, /party, /join, /leave, /resetui", HelpColor)
                Case "/info"

                    ' Checks to make sure we have more than one string in the array
                    If UBound(Command) < 1 Then
                        AddText "Usage: /info (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /info (name)", AlertColor
                        GoTo Continue
                    End If

                    Set Buffer = New clsBuffer
                    Buffer.WriteLong CPlayerInfoRequest
                    Buffer.WriteLong TCP_CODE
                    Buffer.WriteString Command(1)
                    SendData Buffer.ToArray()
                    Set Buffer = Nothing
                    ' Whos Online
                Case "/who"
                    SendWhosOnline
                    ' Checking fps
                Case "/fps"
                    BFPS = Not BFPS
                    ' toggle fps lock
                Case "/fpslock"
                    FPS_Lock = Not FPS_Lock
                    ' Request stats
                Case "/stats"
                    Set Buffer = New clsBuffer
                    Buffer.WriteLong CGetStats
                    Buffer.WriteLong TCP_CODE
                    SendData Buffer.ToArray()
                    Set Buffer = Nothing
                Case "/party"

                    ' Make sure they are actually sending something
                    If UBound(Command) < 1 Then
                        AddText "Usage: /party (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /party (name)", AlertColor
                        GoTo Continue
                    End If

                    Call SendPartyRequest(Command(1))
                    ' Join party
                Case "/join"
                    SendJoinParty
                    ' Leave party
                Case "/leave"
                    SendLeaveParty
                    
                Case "/options"
              
                    ' // Monitor Admin Commands //
                    ' Admin Help
                Case "/admin"

                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    Call AddText("Social Commands:", HelpColor)
                    Call AddText("""msghere = Global Admin Message", HelpColor)
                    Call AddText("=msghere = Private Admin Message", HelpColor)
                    Call AddText("Available Commands: /admin, /loc, /mapeditor, /warpmeto, /warptome, /warpto, /setsprite, /mapreport, /kick, /ban, /edititem, /respawn, /editnpc, /motd, /editshop, /editspell, /debug", HelpColor)
                    ' Kicking a player
                    
                    
                 Case "/mapmusic"
                 If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
                  GoTo Continue
                  End If
                    
                    If UBound(Command) < 1 Then
                    GoTo Continue
                    End If
                    
                    SetMapMusic Command(1)
                Case "/kick"

                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    If UBound(Command) < 1 Then
                        AddText "Usage: /kick (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /kick (name)", AlertColor
                        GoTo Continue
                    End If

                    SendKick Command(1)
                    ' // Mapper Admin Commands //
                    ' Location
                Case "/loc"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    BLoc = Not BLoc
                    ' Map Editor
                    
               
                Case "/mapeditor"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    SendRequestEditMap
                    ' Warping to a player
                Case "/warpmeto"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo Continue
                    End If

                    WarpMeTo Command(1)
                    ' Warping a player to you
                Case "/warptome"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo Continue
                    End If

                    WarpToMe Command(1)
                    ' Warping to a map
                Case "/warpto"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo Continue
                    End If

                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo Continue
                    End If

                    n = CLng(Command(1))

                    ' Check to make sure its a valid map #
                    If n > 0 And n <= MAX_MAPS Then
                        Call WarpTo(n)
                    Else
                        Call AddText("Invalid map number.", Red)
                    End If

                    ' Setting sprite
                Case "/setsprite"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    If UBound(Command) < 1 Then
                        AddText "Usage: /setsprite (sprite #)", AlertColor
                        GoTo Continue
                    End If

                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /setsprite (sprite #)", AlertColor
                        GoTo Continue
                    End If

                    SendSetSprite CLng(Command(1))
                    ' Map report
                Case "/mapreport"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    SendData CMapReport & END_CHAR
                    ' Respawn request
                Case "/respawn"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    SendMapRespawn
                    ' MOTD change
                Case "/motd"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    If UBound(Command) < 1 Then
                        AddText "Usage: /motd (new motd)", AlertColor
                        GoTo Continue
                    End If

                    SendMOTDChange Right$(ChatText, Len(ChatText) - 5)
                    ' Check the ban list
                Case "/banlist"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    SendBanList
                    ' Banning a player
                Case "/ban"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    If UBound(Command) < 1 Then
                        AddText "Usage: /ban (name)", AlertColor
                        GoTo Continue
                    End If

                    SendBan Command(1)
                    ' // Developer Admin Commands //
                    ' Editing item request
                Case "/edititem"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    SendRequestEditItem
                Case "/editpokemon"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    SendRequestEditPokemon
                ' Editing animation request
                Case "/editanimation"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    SendRequestEditAnimation
                    ' Editing npc request
                Case "/editnpc"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    SendRequestEditNpc
                Case "/editresource"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    SendRequestEditResource
                    ' Editing shop request
                    
                   
                    
                Case "/editshop"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    SendRequestEditShop
                    ' Editing spell request
                Case "/editspell"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    SendRequestEditSpell
                    ' // Creator Admin Commands //
                    ' Giving another player access
                Case "/setaccess"

                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    If UBound(Command) < 2 Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Or Not IsNumeric(Command(2)) Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo Continue
                    End If

                    SendSetAccess Command(1), CLng(Command(2))
                    ' Ban destroy
                Case "/destroybanlist"

                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    SendBanDestroy
                    ' Packet debug mode
                Case "/debug"

                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    DEBUG_MODE = (Not DEBUG_MODE)
                Case Else
                    AddText "Not a valid command!", HelpColor
            End Select

            'continue label where we go instead of exiting the sub
Continue:
            MyText = vbNullString
            frmChat.txtMyChat.text = vbNullString
            Exit Sub
        End If

        ' Say message
        If Len(ChatText) > 0 Then
            Call SayMsg(ChatText)
        End If

        MyText = vbNullString
        frmChat.txtMyChat.text = vbNullString
       
    
End Sub

Private Sub AlphaImgCtl8_Click()
PlayClick

If menuLeft = True Then
If picTrade.Visible = True Or picShop.Visible = True Then Exit Sub
picTrainerCard.Visible = False
picPokedex.Visible = False
picMenuOptions.Visible = False
picPokemons.Visible = False
picBag.Visible = False
picProfile.Visible = False
picBank.Visible = False
picEvolve.Visible = False
picLearnMove.Visible = False
picTrade.Visible = False
picTravel.Visible = False
picShop.Visible = False
End If
tmrmenu.Enabled = True
End Sub

Private Sub BaglstItems_Click()
Dim invnum As Long
invnum = BaglstItems.ListIndex + 1
If FileExist("Data Files\itemicons\" & GetPlayerInvItemNum(MyIndex, invnum) & ".png") Then
Bagimgicon.Picture = LoadPictureGDIplus(App.Path & "\Data Files\itemicons\" & GetPlayerInvItemNum(MyIndex, invnum) & ".png")
Exit Sub
Else
If FileExist("Data Files\itemicons\" & GetPlayerInvItemNum(MyIndex, invnum) & ".gif") Then
Bagimgicon.Picture = LoadPictureGDIplus(App.Path & "\Data Files\itemicons\" & GetPlayerInvItemNum(MyIndex, invnum) & ".gif")
Bagimgicon.Animate (lvicAniCmdStart)
Exit Sub
Else
Bagimgicon.Picture = LoadPictureGDIplus(App.Path & "\Data Files\itemicons\0.png")
Exit Sub
End If
End If
Bagimgicon.Picture = Nothing
End Sub

Private Sub BaglstItems_DblClick()
Dim invnum As Long
invnum = BaglstItems.ListIndex + 1
Call SendUseItem(invnum)
BagLoadInv
End Sub

Private Sub btnAdminPanel_Click()
 If Player(MyIndex).Access >= 1 Then
   frmAdmin.Visible = Not frmAdmin.Visible
End If
End Sub

Private Sub btnBag_Click()
PlayClick
frmBag.Show
End Sub



Private Sub cmdAAnim_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If

    SendRequestEditAnimation
End Sub

Private Sub cmdAPokemon_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If

    SendRequestEditPokemon
End Sub



Private Sub cmdLevel_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If

    SendRequestLevelUp
End Sub

Private Sub btnCloseBattle_Click()
CanMoveNow = True
inBattle = False
picBattleCommands.Visible = False
txtBtlLog.text = vbNullString
txtBtlLog.Visible = False
StopPlay
PlayMapMusic MapMusic
PlayClick
End Sub

Private Sub btnHatch_Click()
SendRequest 0, 0, "", "HATCHEGG"
'close
picEgg.Visible = False
menuLeft = True
tmrmenu.Enabled = True
End Sub

Private Sub btnImgMoves_Click(Index As Integer)

End Sub

Private Sub btnStat_Click(Index As Integer)
Select Case Index
Case 0
Call SendRequest(TPRemoveSlot, STAT_HP, 0, "TPREMOVE")
Case 1
Call SendRequest(TPRemoveSlot, STAT_ATK, 0, "TPREMOVE")
Case 2
Call SendRequest(TPRemoveSlot, STAT_DEF, 0, "TPREMOVE")
Case 3
Call SendRequest(TPRemoveSlot, STAT_SPATK, 0, "TPREMOVE")
Case 4
Call SendRequest(TPRemoveSlot, STAT_SPDEF, 0, "TPREMOVE")
Case 5
Call SendRequest(TPRemoveSlot, STAT_SPEED, 0, "TPREMOVE")
End Select
PlayClick
picTPRemove.Visible = False
End Sub



Private Sub btnTravel_Click(Index As Integer)
Select Case Index

Case 0
PlayClick
SendRequest 1, 0, "", "TRAVEL"
picTravel.Visible = False
menuLeft = True
tmrmenu.Enabled = True

Case 1
PlayClick
SendRequest 2, 0, "", "TRAVEL"
picTravel.Visible = False
menuLeft = True
tmrmenu.Enabled = True

Case 2
PlayClick
SendRequest 3, 0, "", "TRAVEL"
picTravel.Visible = False
menuLeft = True
tmrmenu.Enabled = True
Case 3
PlayClick
SendRequest 4, 0, "", "TRAVEL"
picTravel.Visible = False
menuLeft = True
tmrmenu.Enabled = True


End Select
End Sub

Private Sub ClanButton_Click(Index As Integer)
Select Case Index
Case 1
Dim str As String
str = InputBox("Picture link", "Clan Picture")
SendRequest 0, 0, str, "CREWPICTURE"
Case 2
  If MsgBox("Are you sure you wish to delete this clan?", vbYesNo, GAME_NAME) = vbYes Then
  SendRequest 0, 0, str, "CREWDELETE"
  End If
    picCrew.Visible = False
  menuLeft = True
  tmrmenu.Enabled = True
  
 Case 3
  If MsgBox("Are you sure you wish to kick this person?", vbYesNo, GAME_NAME) = vbYes Then
  SendRequest lstClanMembers.ListIndex, 0, str, "CREWKICK"
  End If
    picCrew.Visible = False
  menuLeft = True
  tmrmenu.Enabled = True
  
  Case 4
  If MsgBox("Are you sure you wish to leave the clan?", vbYesNo, GAME_NAME) = vbYes Then
  SendRequest 0, 0, str, "CREWLEAVE"
  End If
  picCrew.Visible = False
  menuLeft = True
  tmrmenu.Enabled = True
  
  Case 5
  
  If txtClanNewsEdit.Visible = False Then
  txtClanNewsEdit.Visible = True
  txtClanNews.Visible = False
  ClanButton(5).Caption = "Save"
  Else
  txtClanNewsEdit.Visible = False
  txtClanNews.Visible = True
  txtClanNews.TextRTF = txtClanNewsEdit.TextRTF
  ClanButton(5).Caption = "Edit news"
  SendRequest 0, 0, txtClanNewsEdit.TextRTF, "CLANNEWS"
  End If
  
End Select

End Sub

Private Sub cmdAutoClose_Click()
If inBattle = False Then Exit Sub
AutoCloseBattle = Not AutoCloseBattle
End Sub

Private Sub cmdBag_Click()
If inBattle = False Then Exit Sub
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

Private Sub cmdClanInvite_Click()
SendRequest 0, 0, Trim$(lblName.Caption), "CLANINVITE"
End Sub

Private Sub cmdClanNO_Click()
SendRequest 0, 0, "NO", "CLANRESPOND"
picDialog.Visible = False
cmdClanNO.Visible = False
cmdClanYES.Visible = False
lvButtons_H10.Visible = True
End Sub

Private Sub cmdClanYES_Click()
SendRequest 0, 0, "YES", "CLANRESPOND"
picDialog.Visible = False
cmdClanNO.Visible = False
cmdClanYES.Visible = False
lvButtons_H10.Visible = True
End Sub

Private Sub cmdJournal_Click()
SendRequest 0, 0, Trim$(lblName.Caption), "JOURNAL"
End Sub

Private Sub cmdPokeMove_Click(Index As Integer)
If inBattle = False Then Exit Sub
PlayClick
Dim i As Long
Dim n As Long
For i = 1 To 4
If PokemonInstance(BattlePokemon).moves(i).number = 0 Then n = n + 1
Next
If cmdPokeMove(Index).Caption = "Struggle (Infinite)" Then
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

Private Sub cmdRun_Click()
If inBattle = False Then Exit Sub
'Try to run
SendBattleCommand 5, BattlePokemon, 1
BlockBattle
End Sub

Private Sub cmdSaveJournal_Click()
PlayClick
If txtJournalEdit.Visible = True Then
SendRequest 0, 0, txtJournalEdit.TextRTF, "SAVEJOURNAL"
txtJournalEdit.Visible = False
txtJournal.Visible = True
txtJournal.TextRTF = txtJournalEdit.TextRTF
cmdSaveJournal.Caption = "Edit"
Else
txtJournalEdit.Visible = True
txtJournal.Visible = False
cmdSaveJournal.Caption = "Save"
End If
End Sub

Private Sub Command1_Click()
 If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If
    SendRequestEditMove
End Sub

Private Sub btnDONT_Click()
picLearnMove.Visible = False
menuLeft = True
tmrmenu.Enabled = True
PlayClick
End Sub

Private Sub btnMove_Click(Index As Integer)
Call SendLearnMove(pSlot, Index, pMove)
picLearnMove.Visible = False
PlayClick
tmrmenu.Enabled = True
End Sub

Private Sub cmdClose_Click()
isInBank = False
picBank.Visible = False
menuLeft = True
tmrmenu.Enabled = True
PlayClick
End Sub

Private Sub cmdDep_Click()
PlayClick
SendDepositPC
End Sub

Private Sub cmdMove_Click()
PlayClick
SendWithdrawPC
End Sub

Private Sub Command2_Click()
MsgBox (Options.music)
Options.music = 1
PlaySound ("Heal.wav")
End Sub

Private Sub Command3_Click()
frmBattle.Show

End Sub

Private Sub btnOut_Click()
PlayClick
Dim Buffer As clsBuffer
    Dim i As Long
    
    isLogging = True
    InGame = False
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CQuit
    Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Call DestroyTCP
    
    ' destroy the animations loaded
    For i = 1 To MAX_BYTE
        ClearAnimInstance (i)
    Next
    
    ' destroy temp values
     DragInvSlotNum = 0
     InvX = 0
     InvY = 0
     EqX = 0
     EqY = 0
     SpellX = 0
     SpellY = 0
     LastItemDesc = 0
     MyIndex = 0
     InventoryItemSelected = 0
     SpellBuffer = 0
     SpellBufferTimer = 0
     tmpDropItem = 0
    
    frmChat.txtChat.text = vbNullString
    picTrade.Visible = False
End Sub




Private Sub cmbItem_Click()
On Error Resume Next
If cmbItem.ListIndex > 0 Then
If Item(GetPlayerInvItemNum(MyIndex, cmbItem.ListIndex)).Type = ITEM_TYPE_CURRENCY Then
picTradeValue.Visible = True
lblVal.Caption = "How many " & Trim$(Item(GetPlayerInvItemNum(MyIndex, cmbItem.ListIndex)).Name) & " do you want to trade?"
Else
SendRequest cmbItem.ListIndex, cmbPoke.ListIndex, "", "TRADEUPDATE"
lvlTradeVal.Caption = "1"
End If
End If
End Sub

Private Sub cmbPoke_Click()
If cmbItem.ListIndex > 0 Then
If Item(GetPlayerInvItemNum(MyIndex, cmbItem.ListIndex)).Type = ITEM_TYPE_CURRENCY Then
SendRequest cmbItem.ListIndex, cmbPoke.ListIndex, Text1.text, "TRADEUPDATE"
lvlTradeVal.Caption = Text1.text
Else
SendRequest cmbItem.ListIndex, cmbPoke.ListIndex, "", "TRADEUPDATE"
lvlTradeVal.Caption = "1"
End If
Else
SendRequest cmbItem.ListIndex, cmbPoke.ListIndex, "", "TRADEUPDATE"
lvlTradeVal.Caption = "0"
End If


End Sub


Private Sub Form_Load()
 Dim tmp As String, contents As String
    If FileExists(App.Path & "\status.txt") = True Then
        Kill App.Path & "\status.txt"
        URLDownloadToFile 0, "http://default.redirectme.net/status.php", App.Path & "\status.txt", 0, 0
        Open App.Path & "\status.txt" For Input As #1
        While EOF(1) = 0
            Line Input #1, tmp
            contents = contents + tmp
        Wend
        Close #1
    Else
       URLDownloadToFile 0, "http://default.redirectme.net/status.php", App.Path & "\status.txt", 0, 0
        Open App.Path & "\status.txt" For Input As #2
        While EOF(1) = 0
            Line Input #2, tmp
            contents = contents + tmp
        Wend
        Close #2
    End If
    
    Call SetStatus("Server Status: " & contents)
    
    spriteIndex = 1
    spriteGender = "m"
    hairIndex = 1
    hairColor = 1
    'picAdmin.Left = 10
    
   ' picLogin.Picture = Nothing
    picLogin.ZOrder 0
    loginDetailsPicture.ZOrder 0
    Set Me.Picture = Nothing
    isInBank = False
    isInStorage = False
    
    AlphaImgCtl13(1).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\logo.png")
    AlphaImgCtl13(0).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\btns\enter.png")
    AlphaImgCtl13(2).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\btns\register.png")
    AlphaImgCtl13(3).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\btns\exit.png")
    AlphaImgCtl13(4).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\btns\continue.png")
    AlphaImgCtl13(7).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\btns\stop.png")
    AlphaImgCtl13(6).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\btns\continue.png")
    
    ' starters
    frmMainGame.GameMaster.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\GM.png")
    frmMainGame.Starter(0).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokeicons\1.png")
    frmMainGame.Starter(1).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokeicons\4.png")
    frmMainGame.Starter(2).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokeicons\7.png")
    frmMainGame.Starter(3).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokeicons\152.png")
    frmMainGame.Starter(4).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokeicons\155.png")
    frmMainGame.Starter(5).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokeicons\158.png")
    frmMainGame.Starter(6).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokeicons\252.png")
    frmMainGame.Starter(7).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokeicons\255.png")
    frmMainGame.Starter(8).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokeicons\258.png")
    frmMainGame.Starter(9).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokeicons\387.png")
    frmMainGame.Starter(10).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokeicons\390.png")
    frmMainGame.Starter(11).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokeicons\393.png")
    frmMainGame.Starter(12).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokeicons\495.png")
    frmMainGame.Starter(13).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokeicons\498.png")
    frmMainGame.Starter(14).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokeicons\501.png")
    
    ' sprite edits
    frmMainGame.Starter(15).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\trainers\m1.png")
    frmMainGame.Starter(16).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\btns\left.png")
    frmMainGame.Starter(17).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\btns\right.png")
    ' paperdoll hair
    frmMainGame.Starter(20).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\trainers\hair\m\1\1.png")
    frmMainGame.Starter(19).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\btns\left.png")
    frmMainGame.Starter(18).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\btns\right.png")
    ' hair color
    frmMainGame.Starter(21).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\btns\left.png")
    frmMainGame.Starter(22).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\btns\right.png")
    
    ' battle hud
    frmMainGame.btnImgMoves(0).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\btns\move.png")
    frmMainGame.btnImgMoves(1).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\btns\move.png")
    frmMainGame.btnImgMoves(2).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\btns\move.png")
    frmMainGame.btnImgMoves(3).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\btns\move.png")
    
       
    Me.Width = 11600
    'txtMyChat.BackColor = RGB(49, 49, 49)
    
    loginDetailsPicture.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    InGame = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then
        MoveForm = True
        MouseX = (PixelsToTwips(x, 0))
        MouseY = (PixelsToTwips(y, 1))
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveForm = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If MoveForm Then
        PresentX = Me.Left - MouseX + (PixelsToTwips(x, 0))
        PresentY = Me.Top - MouseY + (PixelsToTwips(y, 1))
        Me.move PresentX, PresentY
    End If
End Sub

Private Sub Image2_Click()

SendRequest 0, 0, "", "OPENNEWS"
PlayClick
End Sub

Private Sub Image3_Click()
OpenMenu (MENU_POKEDEX)
PlayClick
End Sub

Private Sub Image4_Click()
PlayClick
OpenMenu (MENU_OPTIONS)
End Sub

Private Sub Image7_Click()
SendRequest 0, 0, 0, "EGG"
PlayClick
End Sub

Private Sub Image8_Click()
PlayClick
SendRequest 0, 0, 0, "MARKET"
End Sub

Private Sub imgButton_Click(Index As Integer)
If inBattle Then Exit Sub
Select Case Index
Case 0

PlayClick
SendRequest 0, 0, "", "BAGUPDATE"
OpenMenu (MENU_BAG)
Case 1
PlayClick
SendRosterRequest

Case 2
PlayClick
SendRequestPlayerData
SendRequest 0, 0, 0, "PROFILE"


Case 3
PlayClick
SendRequest 0, 0, "", "CREW"
Case 4


PlayClick

If menuLeft = True Then
If picTrade.Visible = True Or picShop.Visible = True Or picTPRemove.Visible = True Then Exit Sub
picTrainerCard.Visible = False
picPokedex.Visible = False
picMenuOptions.Visible = False
picPokemons.Visible = False
picBag.Visible = False
picProfile.Visible = False
picBank.Visible = False
picEvolve.Visible = False
picLearnMove.Visible = False
picTrade.Visible = False
picTravel.Visible = False
picShop.Visible = False
picTPRemove.Visible = False
picCrew.Visible = False
picPlayerJournal.Visible = False

picEgg.Visible = False
End If
tmrmenu.Enabled = True


End Select
End Sub

Private Sub imgClothes_DblClick(Index As Integer)
Dim l As Long
l = Index + 1
SendUnequip l
End Sub

Private Sub imgItemIcon_Click(Index As Integer)
LoadInvItem Index + 1
PlayClick
End Sub

Private Sub imgSwitch_Click(Index As Integer)
If inBattle = True Then
If PokemonInstance(Index).PokemonNumber > 0 Then
If PokemonInstance(Index).HP > 0 Then
If Not Index = BattlePokemon Then
PlayClick
SendBattleCommand 2, Index, 1
BlockBattle
PlayClick
End If
End If
End If
Else
If PokemonInstance(Index).PokemonNumber > 0 Then
SendSetAsLeader Index
PlayClick
End If
End If
End Sub

Private Sub Label15_Click()

End Sub

Private Sub Label27_Click()
'picBattleInfo.Visible = True
End Sub

Private Sub Label32_Click()
picTrainerCard.Visible = False
End Sub

Private Sub Label54_Click()
'picBattleInfo.Visible = False
End Sub

Private Sub lblCancelDrop_Click()
  
End Sub

Private Sub lblLeaveShop_Click()

End Sub

Private Sub lblOkDrop_Click()
  
End Sub

Private Sub lblShopBuy_Click()
    If ShopAction = 1 Then Exit Sub
    ShopAction = 1 ' buying an item
    AddText "Click on the item in the shop you wish to buy.", White
End Sub

Private Sub lblShopSell_Click()
    If ShopAction = 2 Then Exit Sub
    ShopAction = 2 ' selling an item
    AddText "Double-click on the item in your inventory you wish to sell.", White
End Sub

Private Sub lblTrainStat_Click(Index As Integer)
    If GetPlayerPOINTS(MyIndex) = 0 Then Exit Sub
    SendTrainStat Index
End Sub

Private Sub lvButton1_Click()
SendRosterRequest
End Sub

Private Sub lvButton2_Click()
frmBag.Show
End Sub

Private Sub lvButton3_Click()
 SendRequestPlayerData
   picTrainerCard.Visible = True
   loadPlayerData (MyIndex)
   If Player(MyIndex).Access > 0 Then
GameMaster.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\GM.png")
Else
GameMaster.Picture = Nothing
End If
End Sub

Private Sub lvButton4_Click()
Dim Buffer As clsBuffer
    Dim i As Long
    
    isLogging = True
    InGame = False
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CQuit
    Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Call DestroyTCP
    
    ' destroy the animations loaded
    For i = 1 To MAX_BYTE
        ClearAnimInstance (i)
    Next
    
    ' destroy temp values
     DragInvSlotNum = 0
     InvX = 0
     InvY = 0
     EqX = 0
     EqY = 0
     SpellX = 0
     SpellY = 0
     LastItemDesc = 0
     MyIndex = 0
     InventoryItemSelected = 0
     SpellBuffer = 0
     SpellBufferTimer = 0
     tmpDropItem = 0
    
    frmChat.txtChat.text = vbNullString
End Sub

Private Sub lvButton8_Click()
picTrainerCard.Visible = False
End Sub

Private Sub lvButton5_Click()

End Sub

Private Sub lvButton6_Click()

End Sub

Private Sub lvButton7_Click()

End Sub

Private Sub Label10_Click()
PlayClick
Picture1.Visible = False
End Sub

Private Sub Label12_Click()
PlayClick
If AdminOnly = False Then
DestroyTCP
    
    If Picture1.Visible = False Then
    Picture1.Visible = True
    picNewChar.Visible = False

    Else
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

Private Sub Label5_Click()
PlayClick
If StarterChoosed = 0 Then
MsgBox "Please choose your starter pokemon!"
Else
Call MenuState(MENU_STATE_ADDCHAR)
End If
End Sub

Private Sub Label6_Click()
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

Private Sub lblItemsCaption_Click(Index As Integer)

End Sub

Private Sub lblItemVal_Click(Index As Integer)
LoadInvItem Index + 1
PlayClick
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
SendRequest 0, 0, lblName.Caption, "BATTLE"
tmrmenu.Enabled = True
End Sub

Private Sub List1_Click()

End Sub

Private Sub Label11_Click()
frmMenu.txtLUser.text = txtLUser.text
frmMenu.txtLPass.text = txtLPass.text
PlayClick
Call MenuState(MENU_STATE_LOGIN)
End Sub

Private Sub lvButtons_H10_Click()
PlayClick
Dim i As Long
If CurrentDialog = Dialogs Then
If CurrentDialog > 0 Then
If IsDialogTrigger(CurrentDialog) Then
SendRequest 0, 0, "", "DIALOGTRIGGER"
End If
End If
tmrDialog.Enabled = False
DialogText = ""
DialogPosition = 0
DialogLenght = 0
picDialog.Visible = False
CanMoveNow = True
CurrentDialog = 0
Dialogs = 0
For i = 1 To 100
Dialog(i) = ""
DialogImage(i) = 0
IsDialogTrigger(i) = False
Next
Else
If IsDialogTrigger(CurrentDialog) Then
SendRequest 0, 0, "", "DIALOGTRIGGER"
End If
CurrentDialog = CurrentDialog + 1
DisplayDialogText Trim$(Dialog(CurrentDialog))

'txtDialog.Caption = Trim$(Dialog(CurrentDialog))
If DialogImage(CurrentDialog) > 0 Then
picDialog.Left = 15
'frmMainGame.imgDialogPic.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\faces\" & DialogImage(CurrentDialog) & ".png")
Else
picDialog.Left = 120
frmMainGame.imgDialogPic.Picture = Nothing
End If

End If
End Sub

Private Sub lvButtons_H11_Click()
SendRequest 0, 0, lblName.Caption, "TRADE"
PlayClick
picTrainerCard.Visible = False
End Sub

Private Sub lvButtons_H12_Click()
SendRequest cmbItem.ListIndex, cmbPoke.ListIndex, Text1.text, "TRADELOCK"
End Sub

Private Sub lvButtons_H13_Click()
SendRequest 0, 0, "", "TRADESTOP"
End Sub

Private Sub lvButtons_H14_Click()
If Val(Text1.text) = 0 Then
MsgBox ("You can't trade 0 " & Trim$(Item(GetPlayerInvItemNum(MyIndex, cmbItem.ListIndex)).Name))
Else
If Val(Text1.text) > GetPlayerInvItemValue(MyIndex, cmbItem.ListIndex) Then
MsgBox ("You dont have enough " & Trim$(Item(GetPlayerInvItemNum(MyIndex, cmbItem.ListIndex)).Name))
Else
'SEND
lvlTradeVal.Caption = Text1.text
SendRequest cmbItem.ListIndex, cmbPoke.ListIndex, Text1.text, "TRADEUPDATE"
picTradeValue.Visible = False
End If
End If
End Sub

Private Sub lvButtons_H15_Click()
RosterpicMoves.Visible = False

End Sub

Private Sub lvButtons_H16_Click()
SendSetAsLeader selectedpoke
selectedpoke = 1
RosterLoadPokemon (selectedpoke)
End Sub

Private Sub lvButtons_H17_Click()
RosterLoadPokeMoves (selectedpoke)
RosterpicMoves.Visible = True
RosterpicItems.Visible = False
End Sub

Private Sub lvButtons_H18_Click()
Call SendRequest(selectedpoke, 1, "", "TRYEVOLVE")
End Sub

Private Sub lvButtons_H19_Click()
RosterLoadInv
RosterpicItems.Visible = True
RosterpicMoves.Visible = False
End Sub

Private Sub lvButtons_H2_Click()
RosterpicItems.Visible = False
End Sub

Private Sub lvButtons_H20_Click()
SendAddTP STAT_HP, selectedpoke
End Sub

Private Sub lvButtons_H21_Click()
SendAddTP STAT_ATK, selectedpoke
End Sub

Private Sub lvButtons_H22_Click()
SendAddTP STAT_DEF, selectedpoke
End Sub

Private Sub lvButtons_H23_Click()
SendAddTP STAT_SPATK, selectedpoke
End Sub

Private Sub lvButtons_H24_Click()
SendAddTP STAT_SPDEF, selectedpoke
End Sub

Private Sub lvButtons_H25_Click()
SendAddTP STAT_SPEED, selectedpoke
End Sub

Private Sub lvButtons_H26_Click()
PlayClick
Dim invnum As Long
invnum = BaglstItems.ListIndex + 1
Call SendDropItem(invnum, 1)
BagLoadInv
End Sub

Private Sub lvButtons_H27_Click()
PlayClick
Dim i As Long
Dim eslot As Long
For i = 1 To 6
eslot = GetPlayerEquipment(MyIndex, i)
If eslot > 0 Then
SendUnequip i
End If
Next
BagLoadInv
End Sub

Private Sub lvButtons_H28_Click()
picBag.Visible = False
menuLeft = True
tmrmenu.Enabled = True
PlayClick
End Sub

Private Sub lvButtons_H29_Click()
If InStr(txtProfileImg.text, "http") Then
        imgProfile.Picture = LoadPictureGDIplus(txtProfileImg.text)

        SendRequest 0, 0, Trim$(txtProfileImg.text), "PPIC"
    Else
        MsgBox "You must enter a URL.", vbExclamation
    End If
End Sub

Private Sub lvButtons_H3_Click()
picNews.Visible = False
End Sub



Private Sub lvButtons_H30_Click()
picProfile.Visible = False
menuLeft = True
tmrmenu.Enabled = True
PlayClick

End Sub

Private Sub lvButtons_H31_Click()
Call SendRequest(evPokesloT, evNewPoke, "", "PEV")
picEvolve.Visible = False
PlayClick
End Sub

Private Sub lvButtons_H32_Click()
picEvolve.Visible = False
menuLeft = True
tmrmenu.Enabled = True
PlayClick
End Sub

Private Sub lvButtons_H33_Click()
PlayClick
SendRequest 1, 0, "", "TRAVEL"
picTravel.Visible = False
menuLeft = True
tmrmenu.Enabled = True
End Sub

Private Sub lvButtons_H34_Click()
PlayClick
SendRequest 2, 0, "", "TRAVEL"
picTravel.Visible = False
menuLeft = True
tmrmenu.Enabled = True
End Sub

Private Sub lvButtons_H35_Click()
picTravel.Visible = False
menuLeft = True
tmrmenu.Enabled = True
PlayClick
End Sub

Private Sub lvButtons_H36_Click()
Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CCloseShop
    Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    picShop.Visible = False
    menuLeft = True
    tmrmenu.Enabled = True
    PlayClick
    InShop = 0
    ShopAction = 0
End Sub

Private Sub lvButtons_H37_Click()
On Error Resume Next
If ShoplstMyItems.text <> "" Then
If GetPlayerInvItemNum(MyIndex, ShoplstMyItems.ListIndex + 1) > 0 Then
SellItem ShoplstMyItems.ListIndex + 1
End If
End If
PlayClick
End Sub

Private Sub lvButtons_H38_Click()
If Shop(InShop).TradeItem(ShoplstItems.ListIndex + 1).Item > 0 And Shop(InShop).TradeItem(ShoplstItems.ListIndex + 1).Item <= MAX_ITEMS Then
BuyItem ShoplstItems.ListIndex + 1
End If
PlayClick
End Sub

Private Sub lvButtons_H4_Click()
picPlayerJournal.Visible = False
End Sub

Private Sub lvButtons_H5_Click()
lvButtons_H38.Visible = True
ShoptxtCost.Visible = True
ShoplstItems.Visible = True
lvButtons_H37.Visible = False
ShoptxtPrice.Visible = False
ShoplstMyItems.Visible = False
End Sub

Private Sub lvButtons_H6_Click()
PlayClick
picTrainerCard.Visible = False
tmrmenu.Enabled = True
End Sub



Private Sub lvButtons_H7_Click()
PlayClick
Dim invnum As Long
invnum = selectedItemSlot
Call SendUseItem(invnum)
BagLoadInv
End Sub

Private Sub lvButtons_H8_Click()
lvButtons_H38.Visible = False
ShoptxtCost.Visible = False
ShoplstItems.Visible = False
lvButtons_H37.Visible = True
ShoptxtPrice.Visible = True
ShoplstMyItems.Visible = True
End Sub

Private Sub lvButtons_H9_Click()
If optCheck(0).Value = 1 Then
Options.PlayMusic = YES
PlayMapMusic MapMusic
Else
Options.PlayMusic = NO
StopPlay
End If
If optCheck(1).Value = 1 Then
Options.repeatmusic = YES
Else
Options.repeatmusic = NO
End If
If optCheck(2).Value = 1 Then
Options.CameraFollowPlayer = YES
Else
Options.CameraFollowPlayer = NO
End If
If optCheck(3).Value = 1 Then
Options.FormTransparency = YES
Else
Options.FormTransparency = NO
End If
If optCheck(4).Value = 1 Then
Options.PlayRadio = YES
Else
Options.PlayRadio = NO
End If

If optCheck(5).Value = 1 Then
Options.NearbyMaps = YES
Else
Options.NearbyMaps = NO
End If

SaveOptions
PlayClick
End Sub

Private Sub picClose_Click()

End Sub

Private Sub optFemale_Click()
    spriteGender = "f"
    frmMainGame.Starter(15).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\trainers\" & spriteGender & spriteIndex & ".png")
    frmMainGame.Starter(20).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\trainers\hair\" & spriteGender & "\" & hairIndex & "\" & hairColor & ".png")
    newCharClass = 1
    newCharSprite = 901
    NewCharacterBltSprite (newCharSprite)
End Sub

Private Sub optMale_Click()
    spriteGender = "m"
    frmMainGame.Starter(15).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\trainers\" & spriteGender & spriteIndex & ".png")
    frmMainGame.Starter(20).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\trainers\hair\" & spriteGender & "\" & hairIndex & "\" & hairColor & ".png")
    newCharClass = 0
    newCharSprite = 900
    NewCharacterBltSprite (newCharSprite)
End Sub

Private Sub picItem_Click(Index As Integer)
LoadInvItem Index + 1
PlayClick
End Sub



Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If inBattle Then Exit Sub
Dim i As Long
    If InMapEditor Then
        Call MapEditorMouseDown(Button)
    Else
        Call PlayerSearch(CurX, CurY)
    End If
    If Button = 2 Then
        If Not InMapEditor Then
            If ShiftDown Then
                If Player(MyIndex).Access >= 3 Then
                    For i = 1 To 30
                        If MapNpc(i).x = CurX And MapNpc(i).y = CurY Then
                            frmEditorMapNPC.Show
                            frmEditorMapNPC.load (i)
                            SendRequest i, 0, "", "NPC"
                        Else
                    If InMapEditor = False Then
                        frmSetNpc.Show
                        frmSetNpc.LoadPos CurX, CurY
                    End If
                End If
                Next
            End If
        Else
            WarpAdmin
        End If
        If ControlDown Then
            If Player(MyIndex).Access >= 3 Then
                If InMapEditor = False Then
                    frmSetNpc.Show
                    frmSetNpc.LoadPos CurX, CurY
                End If
            End If
        End If
    End If
    
    End If
    Call SetFocusOnChat
   ChatFocus = False
End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    CurX = TileView.Left + ((x + Camera.Left) \ PIC_X)
    CurY = TileView.Top + ((y + Camera.Top) \ PIC_Y)

    If InMapEditor Then
        frmEditor_Map.shpLoc.Visible = False

        If Button = vbLeftButton Or Button = vbRightButton Then
            Call MapEditorMouseDown(Button)
        End If
    End If
    
   

End Sub

Private Function IsShopItem(ByVal x As Single, ByVal y As Single) As Long
    Dim tempRec As RECT
    Dim i As Long
    IsShopItem = 0

    For i = 1 To MAX_TRADES

        If Shop(InShop).TradeItem(i).Item > 0 And Shop(InShop).TradeItem(i).Item <= MAX_ITEMS Then
            With tempRec
                .Top = ShopTop
                .Bottom = .Top + PIC_Y
                .Left = ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
                .Right = .Left + PIC_X
            End With

            If x >= tempRec.Left And x <= tempRec.Right Then
                If y >= tempRec.Top And y <= tempRec.Bottom Then
                    IsShopItem = i
                    Exit Function
                End If
            End If
        End If
    Next
End Function



Private Sub picShopItems_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim shopItem As Long
    shopItem = IsShopItem(x, y)
    
    If shopItem > 0 Then
        Select Case ShopAction
            Case 0 ' no action, give cost
                With Shop(InShop).TradeItem(shopItem)
                    AddText "You can buy this item for " & .CostValue & " " & Trim$(Item(.CostItem).Name) & ".", White
                End With
            Case 1 ' buy item
                ' buy item code
                BuyItem shopItem
        End Select
    End If
End Sub



Private Sub picSpellDesc_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  
End Sub

Private Sub picSpells_DblClick()
    Dim spellnum As Long
    
    spellnum = IsPlayerSpell(SpellX, SpellY)

    If spellnum <> 0 Then
        Call CastSpell(spellnum)
        Exit Sub
    End If
End Sub

Private Sub picSpells_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim spellnum As Long
    
    If Button = 2 Then ' right click
        spellnum = IsPlayerSpell(SpellX, SpellY)

        If spellnum <> 0 Then
            Call ForgetSpell(spellnum)
            Exit Sub
        End If
    End If
End Sub









Private Sub Picture3_Click()

End Sub

Private Sub PokeButton_Click(Index As Integer)
If PokemonInstance(Index).PokemonNumber > 0 Then
selectedpoke = Index
RosterLoadPokemon (selectedpoke)
End If
End Sub

Private Sub PokedexList1_Click()
PokedexLoadpoke (PokedexList1.ListIndex + 1)
PlayClick
End Sub

Private Sub RichTextBox1_Change()

End Sub

Private Sub RosterlstItems_DblClick()
SendRequest selectedpoke, RosterlstItems.ListIndex + 1, "", "USEITEMONPOKEMON"
End Sub

Private Sub RosterlstMoves_DblClick()
SendRequest selectedpoke, RosterlstMoves.ListIndex + 1, "", "TRYLM"
End Sub

Private Sub ShoplstItems_Click()
If Shop(InShop).TradeItem(ShoplstItems.ListIndex + 1).Item > 0 And Shop(InShop).TradeItem(ShoplstItems.ListIndex + 1).Item <= MAX_ITEMS Then
ShoptxtCost.text = "Cost:" & Shop(InShop).TradeItem(ShoplstItems.ListIndex + 1).CostValue & " " & Trim$(Item(Shop(InShop).TradeItem(ShoplstItems.ListIndex + 1).CostItem).Name)
Else
ShoptxtCost.text = "Cost:0 PokeCoins"
End If
If FileExist("Data Files\itemicons\" & Shop(InShop).TradeItem(ShoplstItems.ListIndex + 1).Item & ".png") Then
imgShopItem.Picture = LoadPictureGDIplus(App.Path & "\Data Files\itemicons\" & Shop(InShop).TradeItem(ShoplstItems.ListIndex + 1).Item & ".png")
Exit Sub
Else
If FileExist("Data Files\itemicons\" & Shop(InShop).TradeItem(ShoplstItems.ListIndex + 1).Item & ".gif") Then
imgShopItem.Picture = LoadPictureGDIplus(App.Path & "\Data Files\itemicons\" & Shop(InShop).TradeItem(ShoplstItems.ListIndex + 1).Item & ".gif")
imgShopItem.Animate (lvicAniCmdStart)
Exit Sub
Else
imgShopItem.Picture = LoadPictureGDIplus(App.Path & "\Data Files\itemicons\0.png")
Exit Sub
End If
End If
imgShopItem.Picture = Nothing
End Sub

Private Sub ShoplstMyItems_Click()
If GetPlayerInvItemNum(MyIndex, ShoplstMyItems.ListIndex + 1) > 0 Then
ShoptxtPrice.text = "Price:" & Item(GetPlayerInvItemNum(MyIndex, ShoplstMyItems.ListIndex + 1)).Price & "PokeCoins"
Else
ShoptxtPrice.text = "Price:0 PokeCoins"
End If
If FileExist("Data Files\itemicons\" & GetPlayerInvItemNum(MyIndex, ShoplstMyItems.ListIndex + 1) & ".png") Then
imgShopItem.Picture = LoadPictureGDIplus(App.Path & "\Data Files\itemicons\" & GetPlayerInvItemNum(MyIndex, ShoplstMyItems.ListIndex + 1) & ".png")
Exit Sub
Else
If FileExist("Data Files\itemicons\" & GetPlayerInvItemNum(MyIndex, ShoplstMyItems.ListIndex + 1) & ".gif") Then
imgShopItem.Picture = LoadPictureGDIplus(App.Path & "\Data Files\itemicons\" & GetPlayerInvItemNum(MyIndex, ShoplstMyItems.ListIndex + 1) & ".gif")
imgShopItem.Animate (lvicAniCmdStart)
Exit Sub
Else
imgShopItem.Picture = LoadPictureGDIplus(App.Path & "\Data Files\itemicons\0.png")
Exit Sub
End If
End If
imgShopItem.Picture = Nothing
End Sub

' Winsock event
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)

    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call HandleKeypresses(KeyAscii)

    ' prevents textbox on error ding sound
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then
        KeyAscii = 0
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyF1
            If Player(MyIndex).Access > 1 Then
                frmAdmin.Visible = Not frmAdmin.Visible
            End If
      Case vbKeyT
            'If isChatVisible = True Then
            'If txtMyChat.text = "" Then
            'txtMyChat.Visible = False
            'txtMyChat.Enabled = False
            'txtChat.Visible = False
            'txtChat.Enabled = False
            'picChat.Visible = False
            'isChatVisible = False
            'End If
            'Else
            'If txtMyChat.text = "" Then
            'txtMyChat.Visible = True
            'txtMyChat.Enabled = True
            'txtChat.Visible = True
            'txtChat.Enabled = True
            'picChat.Visible = True
            'isChatVisible = True
            'End If
            'End If
            Case vbKeyB
            SendRequest 0, 0, "", "BIKE"
    End Select

End Sub

Private Sub Starter_Click(Index As Integer)
Select Case Index
Case 0
StarterChoosed = 1
Starter(0).border = True
Starter(1).border = False
Starter(2).border = False
Starter(3).border = False
Starter(4).border = False
Starter(5).border = False
Starter(6).border = False
Starter(7).border = False
Starter(8).border = False
Starter(9).border = False
Starter(10).border = False
Starter(11).border = False
Starter(12).border = False
Starter(13).border = False
Starter(14).border = False
Case 1
StarterChoosed = 4
Starter(0).border = False
Starter(1).border = True
Starter(2).border = False
Starter(3).border = False
Starter(4).border = False
Starter(5).border = False
Starter(6).border = False
Starter(7).border = False
Starter(8).border = False
Starter(9).border = False
Starter(10).border = False
Starter(11).border = False
Starter(12).border = False
Starter(13).border = False
Starter(14).border = False
Case 2
StarterChoosed = 7
Starter(0).border = False
Starter(1).border = False
Starter(2).border = True
Starter(3).border = False
Starter(4).border = False
Starter(5).border = False
Starter(6).border = False
Starter(7).border = False
Starter(8).border = False
Starter(9).border = False
Starter(10).border = False
Starter(11).border = False
Starter(12).border = False
Starter(13).border = False
Starter(14).border = False
Case 3
StarterChoosed = 152
Starter(0).border = False
Starter(1).border = False
Starter(2).border = False
Starter(3).border = True
Starter(4).border = False
Starter(5).border = False
Starter(6).border = False
Starter(7).border = False
Starter(8).border = False
Starter(9).border = False
Starter(10).border = False
Starter(11).border = False
Starter(12).border = False
Starter(13).border = False
Starter(14).border = False
Case 4
StarterChoosed = 155
Starter(0).border = False
Starter(1).border = False
Starter(2).border = False
Starter(3).border = False
Starter(4).border = True
Starter(5).border = False
Starter(6).border = False
Starter(7).border = False
Starter(8).border = False
Starter(9).border = False
Starter(10).border = False
Starter(11).border = False
Starter(12).border = False
Starter(13).border = False
Starter(14).border = False
Case 5
StarterChoosed = 158
Starter(0).border = False
Starter(1).border = False
Starter(2).border = False
Starter(3).border = False
Starter(4).border = False
Starter(5).border = True
Starter(6).border = False
Starter(7).border = False
Starter(8).border = False
Starter(9).border = False
Starter(10).border = False
Starter(11).border = False
Starter(12).border = False
Starter(13).border = False
Starter(14).border = False
Case 6
StarterChoosed = 252
Starter(0).border = False
Starter(1).border = False
Starter(2).border = False
Starter(3).border = False
Starter(4).border = False
Starter(5).border = False
Starter(6).border = True
Starter(7).border = False
Starter(8).border = False
Starter(9).border = False
Starter(10).border = False
Starter(11).border = False
Starter(12).border = False
Starter(13).border = False
Starter(14).border = False
Case 7
StarterChoosed = 255
Starter(0).border = False
Starter(1).border = False
Starter(2).border = False
Starter(3).border = False
Starter(4).border = False
Starter(5).border = False
Starter(6).border = False
Starter(7).border = True
Starter(8).border = False
Starter(9).border = False
Starter(10).border = False
Starter(11).border = False
Starter(12).border = False
Starter(13).border = False
Starter(14).border = False
Case 8
StarterChoosed = 258
Starter(0).border = False
Starter(1).border = False
Starter(2).border = False
Starter(3).border = False
Starter(4).border = False
Starter(5).border = False
Starter(6).border = False
Starter(7).border = False
Starter(8).border = True
Starter(9).border = False
Starter(10).border = False
Starter(11).border = False
Starter(12).border = False
Starter(13).border = False
Starter(14).border = False
Case 9
StarterChoosed = 387
Starter(0).border = False
Starter(1).border = False
Starter(2).border = False
Starter(3).border = False
Starter(4).border = False
Starter(5).border = False
Starter(6).border = False
Starter(7).border = False
Starter(8).border = False
Starter(9).border = True
Starter(10).border = False
Starter(11).border = False
Case 10
StarterChoosed = 390
Starter(0).border = False
Starter(1).border = False
Starter(2).border = False
Starter(3).border = False
Starter(4).border = False
Starter(5).border = False
Starter(6).border = False
Starter(7).border = False
Starter(8).border = False
Starter(9).border = False
Starter(10).border = True
Starter(11).border = False
Starter(12).border = False
Starter(13).border = False
Starter(14).border = False
Case 11
StarterChoosed = 393
Starter(0).border = False
Starter(1).border = False
Starter(2).border = False
Starter(3).border = False
Starter(4).border = False
Starter(5).border = False
Starter(6).border = False
Starter(7).border = False
Starter(8).border = False
Starter(9).border = False
Starter(10).border = False
Starter(11).border = True
Starter(12).border = False
Starter(13).border = False
Starter(14).border = False
Case 12
StarterChoosed = 495
Starter(0).border = False
Starter(1).border = False
Starter(2).border = False
Starter(3).border = False
Starter(4).border = False
Starter(5).border = False
Starter(6).border = False
Starter(7).border = False
Starter(8).border = False
Starter(9).border = False
Starter(10).border = False
Starter(11).border = False
Starter(12).border = True
Starter(13).border = False
Starter(14).border = False
Case 13
StarterChoosed = 498
Starter(0).border = False
Starter(1).border = False
Starter(2).border = False
Starter(3).border = False
Starter(4).border = False
Starter(5).border = False
Starter(6).border = False
Starter(7).border = False
Starter(8).border = False
Starter(9).border = False
Starter(10).border = False
Starter(11).border = False
Starter(12).border = False
Starter(13).border = True
Starter(14).border = False
Case 14
StarterChoosed = 501
Starter(0).border = False
Starter(1).border = False
Starter(2).border = False
Starter(3).border = False
Starter(4).border = False
Starter(5).border = False
Starter(6).border = False
Starter(7).border = False
Starter(8).border = False
Starter(9).border = False
Starter(10).border = False
Starter(11).border = False
Starter(12).border = False
Starter(13).border = False
Starter(14).border = True

' sprite builder
Case 16
If spriteIndex = 1 Then
    spriteIndex = 4
    frmMainGame.Starter(15).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\trainers\" & spriteGender & spriteIndex & ".png")
Else
    spriteIndex = spriteIndex - 1
    frmMainGame.Starter(15).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\trainers\" & spriteGender & spriteIndex & ".png")
End If
Case 17
If spriteIndex = 4 Then
    spriteIndex = 1
    frmMainGame.Starter(15).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\trainers\" & spriteGender & spriteIndex & ".png")
Else
    spriteIndex = spriteIndex + 1
    frmMainGame.Starter(15).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\trainers\" & spriteGender & spriteIndex & ".png")
End If
Case 18
If hairIndex = 4 Then
    hairIndex = 1
    frmMainGame.Starter(20).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\trainers\hair\" & spriteGender & "\" & hairIndex & "\" & hairColor & ".png")
Else
    hairIndex = hairIndex + 1
    frmMainGame.Starter(20).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\trainers\hair\" & spriteGender & "\" & hairIndex & "\" & hairColor & ".png")
End If
Case 19
If hairIndex = 1 Then
    hairIndex = 4
    frmMainGame.Starter(20).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\trainers\hair\" & spriteGender & "\" & hairIndex & "\" & hairColor & ".png")
Else
    hairIndex = hairIndex - 1
    frmMainGame.Starter(20).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\trainers\hair\" & spriteGender & "\" & hairIndex & "\" & hairColor & ".png")
End If
Case 21
If hairColor = 11 Then
    hairColor = 1
    frmMainGame.Starter(20).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\trainers\hair\" & spriteGender & "\" & hairIndex & "\" & hairColor & ".png")
Else
    hairColor = hairColor + 1
    frmMainGame.Starter(20).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\trainers\hair\" & spriteGender & "\" & hairIndex & "\" & hairColor & ".png")
End If
Case 22
If hairColor = 11 Then
    hairColor = 1
    frmMainGame.Starter(20).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\trainers\hair\" & spriteGender & "\" & hairIndex & "\" & hairColor & ".png")
Else
    hairColor = hairColor + 1
    frmMainGame.Starter(20).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\trainers\hair\" & spriteGender & "\" & hairIndex & "\" & hairColor & ".png")
End If
End Select
End Sub

Private Sub tmrDialog_Timer()
If DialogPosition > DialogLenght Then
tmrDialog.Enabled = False
DialogText = ""
DialogPosition = 0
DialogLenght = 0
Exit Sub
End If
txtDialog.Caption = txtDialog.Caption & Mid(DialogText, DialogPosition, 1)
DialogPosition = DialogPosition + 1
End Sub

Private Sub tmrmenu_Timer()
If menuLeft Then
picMenus.Left = picMenus.Left + 70
If picMenus.Left >= 704 Then
 picMenus.Left = 704
menuLeft = False
tmrmenu.Enabled = False
End If
Else
picMenus.Left = picMenus.Left - 70
If picMenus.Left <= 0 Then
picMenus.Left = 0
menuLeft = True
tmrmenu.Enabled = False
End If
End If
End Sub

Private Sub txtChat_GotFocus()
    SetFocusOnChat
End Sub

' ***************
' ** Inventory **
' ***************
Private Sub lblUseItem_Click()
    Call UseItem
End Sub

Private Sub picInventory_DblClick()
    Dim invnum As Long
    Dim Value As Long
    Dim multiplier As Double
    
    DragInvSlotNum = 0
    invnum = IsInvItem(InvX, InvY)

    If invnum <> 0 Then
    
        ' are we in a shop?
        If InShop > 0 Then
            Select Case ShopAction
                Case 0 ' nothing, give value
                    multiplier = Shop(InShop).BuyRate / 100
                    Value = Item(GetPlayerInvItemNum(MyIndex, invnum)).Price * multiplier
                    If Value > 0 Then
                        AddText "You can sell this item for " & Value & " gold.", White
                    Else
                        AddText "The shop does not want this item.", BrightRed
                    End If
                Case 2 ' 2 = sell
                    SellItem invnum
            End Select
            
            Exit Sub
        End If
        
        ' use item if not in shop
        If Item(GetPlayerInvItemNum(MyIndex, invnum)).Type = ITEM_TYPE_NONE Then Exit Sub
        Call SendUseItem(invnum)
        Exit Sub
    End If

End Sub

Private Function IsEqItem(ByVal x As Single, ByVal y As Single) As Long
    Dim tempRec As RECT
    Dim i As Long
    IsEqItem = 0

    For i = 1 To Equipment.Equipment_Count - 1

        If GetPlayerEquipment(MyIndex, i) > 0 And GetPlayerEquipment(MyIndex, i) <= MAX_ITEMS Then

            With tempRec
                .Top = EqTop
                .Bottom = .Top + PIC_Y
                .Left = EqLeft + ((EqOffsetX + 32) * (((i - 1) Mod EqColumns)))
                .Right = .Left + PIC_X
            End With

            If x >= tempRec.Left And x <= tempRec.Right Then
                If y >= tempRec.Top And y <= tempRec.Bottom Then
                    IsEqItem = i
                    Exit Function
                End If
            End If
        End If

    Next

End Function

Private Function IsInvItem(ByVal x As Single, ByVal y As Single) As Long
    Dim tempRec As RECT
    Dim i As Long
    IsInvItem = 0

    For i = 1 To MAX_INV

        If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then

            With tempRec
                .Top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .Top + PIC_Y
                .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With

            If x >= tempRec.Left And x <= tempRec.Right Then
                If y >= tempRec.Top And y <= tempRec.Bottom Then
                    IsInvItem = i
                    Exit Function
                End If
            End If
        End If

    Next

End Function

Private Function IsPlayerSpell(ByVal x As Single, ByVal y As Single) As Long
    Dim tempRec As RECT
    Dim i As Long
    
    IsPlayerSpell = 0

    For i = 1 To MAX_PLAYER_SPELLS

        If PlayerSpells(i) > 0 And PlayerSpells(i) <= MAX_PLAYER_SPELLS Then

            With tempRec
                .Top = SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                .Bottom = .Top + PIC_Y
                .Left = SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                .Right = .Left + PIC_X
            End With

            If x >= tempRec.Left And x <= tempRec.Right Then
                If y >= tempRec.Top And y <= tempRec.Bottom Then
                    IsPlayerSpell = i
                    Exit Function
                End If
            End If
        End If

    Next

End Function







' *****************
' ** Char window **
' *****************


' *****************
' ** GUI Buttons **
' *****************

Private Sub lblInventory_Click()
    'picInventory.Visible = True
    ''picCharacter.Visible = False
    'picSpells.Visible = False
    frmBag.Show
    
End Sub




Sub loadPlayerData(ByVal Index As Long)
lblName.Caption = Player(Index).Name
Dim a As Long
Dim plvl As Integer
plvl = 0

For a = 1 To 6
If PokemonInstance(a).PokemonNumber <= 0 Then
Set imgCharPokemon(a).Picture = Nothing
lblCharPkmnLvl(a).Caption = "Lvl:0"
Else
If PokemonInstance(a).isShiny = YES Then
imgCharPokemon(a).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\Shiny\" & PokemonInstance(a).PokemonNumber & ".gif")
Else
imgCharPokemon(a).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\" & PokemonInstance(a).PokemonNumber & ".gif")
End If

lblCharPkmnLvl(a).Caption = "Lvl:" & PokemonInstance(a).Level
plvl = plvl + PokemonInstance(a).Level
End If
Next

lblCharPowerLvl.Caption = "Power Lvl:" & plvl

If Player(Index).Access > 0 Then
frmMainGame.GameMaster.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\GM.png")
Call frmMainGame.GameMaster.SetFixedSizeAspect(frmMainGame.GameMaster.Width / 15, frmMainGame.GameMaster.Height / 15, True)
Else
frmMainGame.GameMaster.Picture = Nothing
End If

frmMainGame.imgProfilePic.Picture = LoadPictureGDIplus("http://orig06.deviantart.net/4b80/f/2012/276/2/1/nate_icon_by_pheonixmaster1-d5go0io.png")


If Index = MyIndex Then

lvButtons_H11.Visible = False
lvButtons_H1.Visible = False
Else

End If

End Sub

Private Sub lblPokemonOpen_Click()
SendRosterRequest
    
End Sub

Private Sub lblLogout_Click()
    Dim Buffer As clsBuffer
    Dim i As Long
    
    isLogging = True
    InGame = False
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CQuit
    Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Call DestroyTCP
    
    ' destroy the animations loaded
    For i = 1 To MAX_BYTE
        ClearAnimInstance (i)
    Next
    
    ' destroy temp values
     DragInvSlotNum = 0
     InvX = 0
     InvY = 0
     EqX = 0
     EqY = 0
     SpellX = 0
     SpellY = 0
     LastItemDesc = 0
     MyIndex = 0
     InventoryItemSelected = 0
     SpellBuffer = 0
     SpellBufferTimer = 0
     tmpDropItem = 0
    
    frmChat.txtChat.text = vbNullString
End Sub

' ****************
' ** Admin Menu **
' ****************

Private Sub cmdALoc_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If
    
    BLoc = Not BLoc
End Sub

Private Sub cmdAMap_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If
    
    SendRequestEditMap
End Sub








Private Sub cmdAMapReport_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If

    AddText "Need to change the packet to byte array, Robin.", BrightRed
    'SendData CMapReport & END_CHAR
End Sub

Private Sub cmdARespawn_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If
    
    SendMapRespawn
End Sub



Private Sub cmdAItem_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If

    SendRequestEditItem
End Sub

Private Sub cmdANpc_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If

    SendRequestEditNpc
End Sub

Private Sub cmdAResource_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If

    SendRequestEditResource
End Sub

Private Sub cmdAShop_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If

    SendRequestEditShop
End Sub

Private Sub cmdASpell_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If

    SendRequestEditSpell
End Sub



Private Sub cmdADestroy_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If

    SendBanDestroy
End Sub


Private Sub txtMyChat_Click()
ChatFocus = True
End Sub


Sub loadNews(ByVal rtf As String)
txtGameNews.TextRTF = rtf


End Sub

Public Sub OpenMenu(ByVal menuCode As Byte)
If menuLeft Then
If picTrade.Visible = True Or picShop.Visible = True Or picTPRemove.Visible = True Then Exit Sub
End If
Select Case menuCode
Case MENU_TRAINERCARD
picTrainerCard.Visible = True
picPokedex.Visible = False
picMenuOptions.Visible = False
picPokemons.Visible = False
picBag.Visible = False
picProfile.Visible = False
picBank.Visible = False
picEvolve.Visible = False
picLearnMove.Visible = False
picTrade.Visible = False
picTravel.Visible = False
picShop.Visible = False
picTPRemove.Visible = False
picCrew.Visible = False
picPlayerJournal.Visible = False

picEgg.Visible = False

Case MENU_POKEDEX
PokedexloadAllPokes
PokedexLoadpoke (1)
picTrainerCard.Visible = False
picPokedex.Visible = True
picMenuOptions.Visible = False
picPokemons.Visible = False
picBag.Visible = False
picProfile.Visible = False
picBank.Visible = False
picEvolve.Visible = False
picLearnMove.Visible = False
picTrade.Visible = False
picTPRemove.Visible = False
picTravel.Visible = False
picShop.Visible = False
picCrew.Visible = False
picEgg.Visible = False

Case MENU_OPTIONS
optCheck(0).Value = Options.PlayMusic
optCheck(1).Value = Options.repeatmusic
optCheck(2).Value = Options.CameraFollowPlayer
optCheck(3).Value = Options.FormTransparency
optCheck(4).Value = Options.PlayRadio
optCheck(5).Value = Options.NearbyMaps
picTrainerCard.Visible = False
picPokedex.Visible = False
picMenuOptions.Visible = True
picPokemons.Visible = False
picBag.Visible = False
picProfile.Visible = False
picBank.Visible = False
picEvolve.Visible = False
picLearnMove.Visible = False
picTrade.Visible = False
picTravel.Visible = False
picShop.Visible = False
picTPRemove.Visible = False
picCrew.Visible = False
picEgg.Visible = False


Case MENU_ROSTER
picTrainerCard.Visible = False
picPokedex.Visible = False
picMenuOptions.Visible = False
picPokemons.Visible = True
picBag.Visible = False
picProfile.Visible = False
picBank.Visible = False
picEvolve.Visible = False
picLearnMove.Visible = False
picTrade.Visible = False
picTravel.Visible = False
picShop.Visible = False
picTPRemove.Visible = False
picCrew.Visible = False
picEgg.Visible = False
RosterloadImages
selectedpoke = 1
RosterLoadPokemon (selectedpoke)
RosterlblTP.Caption = Val(PokemonInstance(1).TP)
RosterloadImages
RosterLoadPokemon (selectedpoke)
RosterLoadInv

Case MENU_BAG
picTrainerCard.Visible = False
picPokedex.Visible = False
picMenuOptions.Visible = False
picPokemons.Visible = False
picBag.Visible = True
picProfile.Visible = False
picBank.Visible = False
picEvolve.Visible = False
picLearnMove.Visible = False
picTrade.Visible = False
picTravel.Visible = False
picShop.Visible = False
picTPRemove.Visible = False
picCrew.Visible = False
picEgg.Visible = False
BagLoadInv
LoadInvItem 1


Case MENU_PROFILE
picTrainerCard.Visible = False
picPokedex.Visible = False
picMenuOptions.Visible = False
picPokemons.Visible = False
picBag.Visible = False
picProfile.Visible = True
picBank.Visible = False
picEvolve.Visible = False
picLearnMove.Visible = False
picTrade.Visible = False
picTravel.Visible = False
picShop.Visible = False
picTPRemove.Visible = False
picCrew.Visible = False
picEgg.Visible = False
loadClothes

Case MENU_BANK
picTrainerCard.Visible = False
picPokedex.Visible = False
picMenuOptions.Visible = False
picPokemons.Visible = False
picBag.Visible = False
picProfile.Visible = False
picBank.Visible = True
picEvolve.Visible = False
picLearnMove.Visible = False
picTrade.Visible = False
picTravel.Visible = False
picShop.Visible = False
picTPRemove.Visible = False
picCrew.Visible = False
picEgg.Visible = False
isInBank = True
LoadBank

Case MENU_EVOLVE
picTrainerCard.Visible = False
picPokedex.Visible = False
picMenuOptions.Visible = False
picPokemons.Visible = False
picBag.Visible = False
picProfile.Visible = False
picBank.Visible = False
picEvolve.Visible = True
picLearnMove.Visible = False
picTrade.Visible = False
picTravel.Visible = False
picShop.Visible = False
picTPRemove.Visible = False
picCrew.Visible = False
picEgg.Visible = False

Case MENU_LEARNMOVE
picTrainerCard.Visible = False
picPokedex.Visible = False
picMenuOptions.Visible = False
picPokemons.Visible = False
picBag.Visible = False
picProfile.Visible = False
picBank.Visible = False
picEvolve.Visible = False
picLearnMove.Visible = True
picTrade.Visible = False
picTravel.Visible = False
picShop.Visible = False
picTPRemove.Visible = False
picCrew.Visible = False
picEgg.Visible = False

Case MENU_TRADE
picTrainerCard.Visible = False
picPokedex.Visible = False
picMenuOptions.Visible = False
picPokemons.Visible = False
picBag.Visible = False
picProfile.Visible = False
picBank.Visible = False
picEvolve.Visible = False
picLearnMove.Visible = False
picTrade.Visible = True
picTravel.Visible = False
picShop.Visible = False
picTPRemove.Visible = False
picCrew.Visible = False
picEgg.Visible = False

Case MENU_TRAVEL
picTrainerCard.Visible = False
picPokedex.Visible = False
picMenuOptions.Visible = False
picPokemons.Visible = False
picBag.Visible = False
picProfile.Visible = False
picBank.Visible = False
picEvolve.Visible = False
picLearnMove.Visible = False
picTrade.Visible = False
picTravel.Visible = True
picShop.Visible = False
picTPRemove.Visible = False
picCrew.Visible = False
picEgg.Visible = False

Case MENU_SHOP
picTrainerCard.Visible = False
picPokedex.Visible = False
picMenuOptions.Visible = False
picPokemons.Visible = False
picBag.Visible = False
picProfile.Visible = False
picBank.Visible = False
picEvolve.Visible = False
picLearnMove.Visible = False
picTrade.Visible = False
picTravel.Visible = False
picShop.Visible = True
picTPRemove.Visible = False
picCrew.Visible = False
picEgg.Visible = False
 ShoploadItems
 ShoploadMyItems
 
 Case MENU_TPREMOVE
 picTrainerCard.Visible = False
picPokedex.Visible = False
picMenuOptions.Visible = False
picPokemons.Visible = False
picBag.Visible = False
picProfile.Visible = False
picBank.Visible = False
picEvolve.Visible = False
picLearnMove.Visible = False
picTrade.Visible = False
picTravel.Visible = False
picShop.Visible = False
picTPRemove.Visible = True
picCrew.Visible = False
picEgg.Visible = False


Case MENU_CREW
picTrainerCard.Visible = False
picPokedex.Visible = False
picMenuOptions.Visible = False
picPokemons.Visible = False
picBag.Visible = False
picProfile.Visible = False
picBank.Visible = False
picEvolve.Visible = False
picLearnMove.Visible = False
picTrade.Visible = False
picTravel.Visible = False
picShop.Visible = False
picTPRemove.Visible = False
picCrew.Visible = True
picEgg.Visible = False

Case MENU_EGG
picTrainerCard.Visible = False
picPokedex.Visible = False
picMenuOptions.Visible = False
picPokemons.Visible = False
picBag.Visible = False
picProfile.Visible = False
picBank.Visible = False
picEvolve.Visible = False
picLearnMove.Visible = False
picTrade.Visible = False
picTravel.Visible = False
picShop.Visible = False
picTPRemove.Visible = False
picCrew.Visible = False
picEgg.Visible = True

End Select
menuLeft = False
tmrmenu.Enabled = True

End Sub


'POKEDEX

Sub PokedexLoadpoke(ByVal poke As Long)
lblPokedexPOKE.Caption = Trim$(Pokemon(poke).Name)
lblPokedexHP.Caption = "HP:" & Pokemon(poke).MaxHp
lblPokedexATK.Caption = "ATK:" & Pokemon(poke).ATK
lblPokedexDEF.Caption = "DEF:" & Pokemon(poke).DEF
lblPokedexSPATK.Caption = "SP.ATK:" & Pokemon(poke).SPATK
lblPokedexSPDEF.Caption = "SP.DEF:" & Pokemon(poke).SPDEF
lblPokedexSPEED.Caption = "SPEED:" & Pokemon(poke).SPD
imgPokedexPoke.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\" & poke & ".gif")
imgPokedexType1.Picture = LoadPicture(App.Path & "\Data Files\graphics\types\" & Pokemon(poke).Type & ".bmp")
imgPokedexType2.Picture = LoadPicture(App.Path & "\Data Files\graphics\types\" & Pokemon(poke).Type2 & ".bmp")
PokedexLoadPokeMoves (poke)
End Sub

Sub PokedexLoadPokeMoves(ByVal slot As Long)
lstPokedexMoves.Clear
Dim i As Long
Dim pokenum As Long
pokenum = slot
For i = 1 To 30
If Pokemon(pokenum).moves(i) > 0 Then
lstPokedexMoves.AddItem (Trim$(PokemonMove(Pokemon(pokenum).moves(i)).Name) & " - Lv." & Pokemon(pokenum).movesLV(i))
End If
Next
End Sub
Sub PokedexloadAllPokes()
PokedexList1.Clear
Dim i As Long
For i = 1 To MAX_POKEMONS
PokedexList1.AddItem (i & ": " & Trim$(Pokemon(i).Name))
Next
End Sub


'ROSTER
Sub RosterloadImages()
Dim i As Long
Dim pn As Long
For i = 1 To 6
If PokemonInstance(i).PokemonNumber > 0 Then
pn = PokemonInstance(i).PokemonNumber
PokeButton(i).Caption = Trim$(Pokemon(pn).Name)
Else
PokeButton(i).Caption = ""
End If
Next
End Sub

Sub RosterLoadPokemon(ByVal num As Long)
If PokemonInstance(num).PokemonNumber > 0 Then
RosterLoadPokeMoves (num)
RosterLoadNatureBoost (num)
Dim str As String
str = PokemonInstance(num).PokemonNumber
If Len(str) = 1 Then
str = "00" & str
End If
If Len(str) = 2 Then
str = "0" & str
End If
If Len(str) = 3 Then
End If
Rosterlblnum.Caption = str
If PokemonInstance(num).isShiny = YES Then
RosterimgPokemon.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\Shiny\" & PokemonInstance(num).PokemonNumber & ".gif")
RosterShinyImage.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\Shiny.gif")

RosterShinyImage.Animate (lvicAniCmdStart)
Else
RosterimgPokemon.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\" & PokemonInstance(num).PokemonNumber & ".gif")
RosterShinyImage.Picture = Nothing
End If

RosterlblTP.Caption = Val(PokemonInstance(num).TP)
RosterlblName.Caption = Trim$(Pokemon(PokemonInstance(num).PokemonNumber).Name)
If PokemonInstance(num).nature < 1 Or PokemonInstance(num).nature > MAX_NATURES Then
RosterlblNature.Caption = "None."
Else
RosterlblNature.Caption = Trim$(nature(PokemonInstance(num).nature).Name)
End If
RosterimgType1.Picture = LoadPicture(App.Path & "\Data Files\graphics\types\" & Pokemon(PokemonInstance(num).PokemonNumber).Type & ".bmp")
RosterimgType2.Picture = LoadPicture(App.Path & "\Data Files\graphics\types\" & Pokemon(PokemonInstance(num).PokemonNumber).Type2 & ".bmp")
RosterlblLevel.Caption = Val(PokemonInstance(num).Level)
RosterlblHp.Caption = Val(PokemonInstance(num).HP) & "/" & Val(PokemonInstance(num).MaxHp)
RosterlblMaxHp.Caption = Val(PokemonInstance(num).MaxHp)
RosterlblAtk.Caption = Val(PokemonInstance(num).ATK)
RosterlblDef.Caption = Val(PokemonInstance(num).DEF)
RosterlblSpAtk.Caption = Val(PokemonInstance(num).SPATK)
RosterlblSpDef.Caption = Val(PokemonInstance(num).SPDEF)
RosterlblSpeed.Caption = Val(PokemonInstance(num).SPD)
lblEXP.Caption = Val(PokemonInstance(num).EXP) & "/" & Val(PokemonInstance(num).expNeeded)
If PokemonInstance(num).HoldingItem > 0 Then
lblPokeItem.Caption = Trim$(Item(PokemonInstance(num).HoldingItem).Name)
Else
lblPokeItem.Caption = "None"
End If
Else
End If
RosterLoadInv
End Sub

Sub RosterLoadPokeMoves(ByVal slot As Long)
RosterlstMoves.Clear
Dim i As Long
Dim pokenum As Long
pokenum = PokemonInstance(slot).PokemonNumber
For i = 1 To 4
If PokemonInstance(slot).moves(i).number > 0 Then
RosterlblMove(i).Caption = Trim$(PokemonMove(PokemonInstance(slot).moves(i).number).Name) & vbNewLine & "PP " & PokemonInstance(slot).moves(i).pp
Else
RosterlblMove(i).Caption = "None." & vbNewLine & "PP 0"
End If
Next
For i = 1 To 30
If Pokemon(pokenum).moves(i) > 0 Then
RosterlstMoves.AddItem (Trim$(PokemonMove(Pokemon(pokenum).moves(i)).Name) & " - Lv." & Pokemon(pokenum).movesLV(i))
End If
Next
End Sub

Sub RosterLoadNatureBoost(ByVal slot As Long)
Dim natNum As Long
natNum = PokemonInstance(slot).nature
lblNatureAtk.Caption = "Atk +" & nature(natNum).AddAtk
lblNatureDef.Caption = "Def +" & nature(natNum).AddDef
lblNatureSpAtk.Caption = "Sp.Atk +" & nature(natNum).AddSpAtk
lblNatureSpDef.Caption = "Sp.Def +" & nature(natNum).AddSpDef
lblNatureSpd.Caption = "Speed +" & nature(natNum).AddSpd
lblNatureHP.Caption = "HP +" & nature(natNum).AddHP
End Sub
Public Sub RosterLoadInv()
Dim i As Long
Dim itemnum As Long
Dim itemvalue As Long
RosterlstItems.Clear
For i = 1 To MAX_INV
itemnum = GetPlayerInvItemNum(MyIndex, i)
itemvalue = GetPlayerInvItemValue(MyIndex, i)
If itemvalue = 0 Then itemvalue = 1
If itemnum <= MAX_ITEMS Then
If itemnum = 0 Then
RosterlstItems.AddItem ("Empty")
Else
RosterlstItems.AddItem (Item(itemnum).Name & " x" & itemvalue)
End If
End If
Next
End Sub
'BAG
Public Sub BagLoadInv()
Dim i As Long
Dim itemnum As Long
Dim itemvalue As Long
'BaglstItems.Clear

For i = 0 To 34
imgItemIcon(i).Aspect = lvicScaleDownOnly
Next
For i = 1 To MAX_INV
itemnum = GetPlayerInvItemNum(MyIndex, i)
itemvalue = GetPlayerInvItemValue(MyIndex, i)
If itemvalue = 0 Then itemvalue = 1
If itemnum <= MAX_ITEMS Then
If itemnum = 0 Then
'BaglstItems.AddItem ("Empty")
imgItemIcon(i - 1).Picture = Nothing
lblItemVal(i - 1).Caption = ""
Else
'BaglstItems.AddItem (Item(itemnum).Name & " x" & itemvalue)
lblItemVal(i - 1).Caption = "x" & itemvalue

If FileExist("Data Files\itemicons\" & itemnum & ".png") Then
imgItemIcon(i - 1).Picture = LoadPictureGDIplus(App.Path & "\Data Files\itemicons\" & itemnum & ".png")
Else
If FileExist("Data Files\itemicons\" & itemnum & ".gif") Then
imgItemIcon(i - 1).Picture = LoadPictureGDIplus(App.Path & "\Data Files\itemicons\" & itemnum & ".gif")
imgItemIcon(i - 1).Animate (lvicAniCmdStart)
Else
imgItemIcon(i - 1).Picture = LoadPictureGDIplus(App.Path & "\Data Files\itemicons\0.png")
End If
End If
End If
End If
Next
End Sub
'BANK
Sub LoadBank()
lblSPC.Caption = "Stored PC:" & Player(MyIndex).StoredPC
Dim i As Long
Dim pcslot As Long
For i = 1 To MAX_INV
If GetPlayerInvItemNum(MyIndex, i) = 1 Then
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
'EVOLUTION

Public Sub LoadEvolution(ByVal slot As Long, ByVal newpoke As Long)
If PokemonInstance(slot).isShiny = YES Then
imgOldPoke.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\Shiny\" & PokemonInstance(slot).PokemonNumber & ".gif")
imgNewPoke.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\Shiny\" & newpoke & ".gif")
imgOldPoke.Animate (lvicAniCmdStart)
imgNewPoke.Animate (lvicAniCmdStart)
Else
imgOldPoke.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\" & PokemonInstance(slot).PokemonNumber & ".gif")
imgNewPoke.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\" & newpoke & ".gif")
imgOldPoke.Animate (lvicAniCmdStart)
imgNewPoke.Animate (lvicAniCmdStart)
End If
evPokesloT = slot
evNewPoke = newpoke
End Sub

'LEARN MOVE
Public Sub LoadMoveAndPoke(ByVal slot As Long, ByVal move As Long)
Dim i As Long
For i = 1 To 4
If PokemonInstance(slot).moves(i).number > 0 Then
btnMove(i).Caption = Trim$(PokemonMove(PokemonInstance(slot).moves(i).number).Name)
Else
btnMove(i).Caption = "None."
End If
Next
lblLearnMove.Caption = Trim$(Pokemon(PokemonInstance(slot).PokemonNumber).Name) & " wants to learn " & Trim$(PokemonMove(move).Name) & ".Choose move to replace."
pSlot = slot
pMove = move
End Sub
'SHOP
Sub ShoploadItems()
Dim i As Long
ShoplstItems.Clear
For i = 1 To MAX_TRADES
If Shop(InShop).TradeItem(i).Item > 0 Then
ShoplstItems.AddItem (Trim$(Item(Shop(InShop).TradeItem(i).Item).Name))
Else
ShoplstItems.AddItem ("Empty")
End If




Next
End Sub

Sub ShoploadMyItems()
Dim i As Long
Dim itemnum As Long
Dim itemvalue As Long
ShoplstMyItems.Clear
For i = 1 To MAX_INV
itemnum = GetPlayerInvItemNum(MyIndex, i)
itemvalue = GetPlayerInvItemValue(MyIndex, i)
If itemvalue = 0 Then itemvalue = 1
If itemnum <= MAX_ITEMS Then
If itemnum = 0 Then
ShoplstMyItems.AddItem ("Empty")
Else
ShoplstMyItems.AddItem (Item(itemnum).Name & " x" & itemvalue)
End If

End If
Next
End Sub

Sub loadClothes()
Dim i As Long
For i = 1 To Equipment.Equipment_Count - 1
If GetPlayerEquipment(MyIndex, i) > 0 Then
If Item(GetPlayerEquipment(MyIndex, i)).Paperdoll > 0 Then
imgClothes(i - 1).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\paperdolls\" & Item(GetPlayerEquipment(MyIndex, i)).Paperdoll & ".bmp")
CropImage imgClothes(i - 1), 0, 16, 64, 64 ' cropimage - theimg, x, y, width, height
imgClothes(i - 1).Width = 735
imgClothes(i - 1).Height = 735
'imgClothes(i - 1).Aspect = lvicScaleDownOnly
imgClothes(i - 1).TransparentColorMode = lvicTransparentTopLeft
Else
imgClothes(i - 1).Picture = LoadPictureGDIplus(App.Path & "\Data Files\itemicons\0.png")
End If
Else
imgClothes(i - 1).Picture = LoadPictureGDIplus(App.Path & "\Data Files\itemicons\0.png")
End If
Next
End Sub


Sub LoadInvItem(ByVal itemIndex As Long)
Dim invnum As Long
Dim i As Long
invnum = itemIndex
selectedItemSlot = itemIndex
Dim indexOfPic As Long
indexOfPic = itemIndex - 1
For i = 0 To 34
If i = indexOfPic Then
picItem(i).backColor = &H808080
Else
picItem(i).backColor = &H554C42
End If
Next



If GetPlayerInvItemNum(MyIndex, invnum) > 0 Then
lblItemInfo(0).Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, invnum)).Name)
lblItemInfo(1).Caption = "x" & GetPlayerInvItemValue(MyIndex, invnum)
Else
lblItemInfo(0).Caption = "None"
lblItemInfo(1).Caption = "x0"
Bagimgicon.Picture = Nothing
Exit Sub
End If


If FileExist("Data Files\itemicons\" & GetPlayerInvItemNum(MyIndex, invnum) & ".png") Then
Bagimgicon.Picture = LoadPictureGDIplus(App.Path & "\Data Files\itemicons\" & GetPlayerInvItemNum(MyIndex, invnum) & ".png")
Exit Sub
Else
If FileExist("Data Files\itemicons\" & GetPlayerInvItemNum(MyIndex, invnum) & ".gif") Then
Bagimgicon.Picture = LoadPictureGDIplus(App.Path & "\Data Files\itemicons\" & GetPlayerInvItemNum(MyIndex, invnum) & ".gif")
Bagimgicon.Animate (lvicAniCmdStart)
Exit Sub
Else
Bagimgicon.Picture = LoadPictureGDIplus(App.Path & "\Data Files\itemicons\0.png")
Exit Sub
End If
End If
Bagimgicon.Picture = Nothing
End Sub


Sub DisplayDialogText(ByVal Txt As String)
txtDialog.Caption = ""
DialogText = Txt
DialogPosition = 1
DialogLenght = Len(Txt)
tmrDialog.Enabled = True
End Sub

