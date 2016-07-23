VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbout 
   OleObjectBlob   =   "frmAbout.frx":0000
   Caption         =   "About Excel2LaTeX"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   9
End
Attribute VB_Name = "frmAbout"
Attribute VB_Base = "0{9C7DC01B-B6DB-4D30-9443-DA5E449911AF}{C224D3FA-BA57-4098-AD07-72A8681ED9A2}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub cmdClose_Click()
  Hide
End Sub


Private Sub UserForm_Initialize()
  TextBox1.SelStart = 0
  TextBox1.SelLength = 0
End Sub
