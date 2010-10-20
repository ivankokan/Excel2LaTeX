VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbout 
   Caption         =   "About Excel2LaTeX"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   OleObjectBlob   =   "frmAbout.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Private Sub cmdClose_Click()
  Hide
End Sub


Private Sub UserForm_Initialize()
  TextBox1.SelStart = 0
  TextBox1.SelLength = 0
End Sub
