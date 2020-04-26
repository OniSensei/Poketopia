VERSION 5.00
Object = "{50347EDF-F3EF-4392-AFDD-71AE67A3A978}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LVbutton.ocx"
Begin VB.Form frmChoose 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Choose Starter!"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   3555
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   2520
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
      Caption         =   "Choose!"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      cBack           =   14737632
   End
   Begin VB.Image imgType1 
      Height          =   375
      Left            =   480
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Image imgType2 
      Height          =   375
      Left            =   1680
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1215
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgPoke 
      Height          =   1335
      Left            =   960
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2355
      Effects         =   "frmChoose.frx":0000
   End
End
Attribute VB_Name = "frmChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
