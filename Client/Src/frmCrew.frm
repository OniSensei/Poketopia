VERSION 5.00
Object = "{50347EDF-F3EF-4392-AFDD-71AE67A3A978}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LVbutton.ocx"
Begin VB.Form frmCrew 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Crew"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCrew.frx":0000
   ScaleHeight     =   377
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   536
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   1980
      Left            =   5040
      TabIndex        =   8
      Top             =   3480
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   2565
      Left            =   5040
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   4080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Kick Player"
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
      cFHover         =   16777215
      cBhover         =   4210752
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   4560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Make admin"
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
      cFHover         =   16777215
      cBhover         =   4210752
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin lvButton.lvButtons_H lvButtons_H3 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4560
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      Caption         =   "Delete crew"
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
      cFHover         =   16777215
      cBhover         =   4210752
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin lvButton.lvButtons_H lvButtons_H4 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      Caption         =   "Set crew picture"
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
      cFHover         =   16777215
      cBhover         =   4210752
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Join requests"
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
      Left            =   4080
      TabIndex        =   7
      Top             =   3120
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Crew Members"
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
      Left            =   4080
      TabIndex        =   6
      Top             =   0
      Width           =   4575
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgClan 
      Height          =   2025
      Left            =   1080
      Top             =   1560
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   3572
      Image           =   "frmCrew.frx":9C814
      Attr            =   514
      Effects         =   "frmCrew.frx":9D94F
   End
   Begin VB.Label lblnum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CrewName"
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
      Top             =   960
      Width           =   4575
   End
End
Attribute VB_Name = "frmCrew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AlphaImgCtl1_Click()

End Sub
