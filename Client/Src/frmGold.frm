VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LVbutton.ocx"
Begin VB.Form frmGold 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Donate"
   ClientHeight    =   2310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmGold.frx":0000
   ScaleHeight     =   2310
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   375
      Left            =   6960
      TabIndex        =   1
      Top             =   1800
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Caption         =   "X"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777088
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGold.frx":381CA
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
   Begin VB.Image Image2 
      Height          =   1500
      Left            =   360
      Picture         =   "frmGold.frx":38308
      Top             =   360
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   5400
      Picture         =   "frmGold.frx":3F87A
      Top             =   960
      Width           =   1350
   End
End
Attribute VB_Name = "frmGold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lvButtons_H1_Click()
Unload Me
End Sub
