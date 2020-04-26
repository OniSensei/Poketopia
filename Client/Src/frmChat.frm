VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LVbutton.ocx"
Begin VB.Form frmChat 
   BackColor       =   &H00312920&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11505
   Icon            =   "frmChat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChat.frx":000C
   ScaleHeight     =   2160
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSendTo 
      Appearance      =   0  'Flat
      BackColor       =   &H00554C42&
      BorderStyle     =   0  'None
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
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   0
      Width           =   1410
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00312920&
      Caption         =   "Drag"
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
      Left            =   0
      Picture         =   "frmChat.frx":AFCD0
      TabIndex        =   4
      Top             =   0
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   1560
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "Global"
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
      cFore           =   0
      cFHover         =   0
      cBhover         =   14737632
      LockHover       =   1
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777215
   End
   Begin VB.TextBox txtMyChat 
      Appearance      =   0  'Flat
      BackColor       =   &H00554C42&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   8850
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "Map"
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
      cFore           =   0
      cFHover         =   0
      cBhover         =   14737632
      LockHover       =   1
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H lvButtons_H3 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "Commands"
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
      cFore           =   0
      cFHover         =   0
      cBhover         =   14737632
      LockHover       =   1
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777215
   End
   Begin VB.PictureBox picCommands 
      Appearance      =   0  'Flat
      BackColor       =   &H00312920&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   2280
      ScaleHeight     =   1665
      ScaleWidth      =   10065
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   10095
      Begin lvButton.lvButtons_H lvButtons_H4 
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
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
         cFore           =   0
         cFHover         =   0
         cBhover         =   14737632
         LockHover       =   1
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H btnCommand 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   120
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "Make Clan"
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
         cFore           =   0
         cFHover         =   0
         cBhover         =   14737632
         LockHover       =   1
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H btnCommand 
         Height          =   375
         Index           =   1
         Left            =   2280
         TabIndex        =   10
         Top             =   120
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "Intro"
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
         cFore           =   0
         cFHover         =   0
         cBhover         =   14737632
         LockHover       =   1
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H btnCommand 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "Bike (B)"
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
         cFore           =   0
         cFHover         =   0
         cBhover         =   14737632
         LockHover       =   1
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   1665
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   2937
      _Version        =   393217
      BackColor       =   3221792
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmChat.frx":15F994
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Eurostar"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCommand_Click(Index As Integer)
Select Case Index
Case 0
txtMyChat.text = "/makeclan NAME"
Case 1
txtMyChat.text = "/intro"
Case 2
SendRequest 0, 0, "", "BIKE"
End Select
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
Timer1.Enabled = False
Else
Timer1.Enabled = True
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If UnloadMode = vbFormControlMenu Or UnloadMode = 1 Then
        'the X has been clicked or the user has pressed Alt+F4
        Cancel = True
    End If
End Sub

Private Sub Form_Resize()
If Me.Width <= 5205 Then Me.Width = 5205
Me.Height = 2640
txtMyChat.Width = Me.Width - txtMyChat.Left - 60
txtChat.Width = Me.Width - txtChat.Left - 60
End Sub

Private Sub lvButtons_H1_Click()
txtMyChat.text = "'"
MyText = txtMyChat
End Sub

Private Sub lvButtons_H2_Click()
txtMyChat.text = ""
MyText = txtMyChat
End Sub

Private Sub lvButtons_H3_Click()
picCommands.Visible = True
End Sub

Private Sub lvButtons_H4_Click()
picCommands.Visible = False
End Sub

Private Sub Timer1_Timer()
Me.Left = frmMainGame.Left
Me.Top = frmMainGame.Top + frmMainGame.Height
End Sub

Private Sub Timer2_Timer()

End Sub

Private Sub txtMyChat_Change()
 MyText = txtMyChat
End Sub

Private Sub txtMyChat_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Call HandleKeypresses(KeyAscii)
KeyAscii = 0
End If
End Sub

Private Sub txtSendTo_Change()
TextSendTo = txtSendTo.text
End Sub
