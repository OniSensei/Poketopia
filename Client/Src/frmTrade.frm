VERSION 5.00
Object = "{50347EDF-F3EF-4392-AFDD-71AE67A3A978}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmTrade 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Trade"
   ClientHeight    =   5550
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl2 
      Height          =   1815
      Left            =   3960
      Top             =   960
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   3201
      Effects         =   "frmTrade.frx":0000
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl1 
      Height          =   1815
      Left            =   360
      Top             =   840
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   3201
      Effects         =   "frmTrade.frx":0018
   End
End
Attribute VB_Name = "frmTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
