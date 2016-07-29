VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbout 
   OleObjectBlob   =   "frmAbout.frx":0000
   Caption         =   "About Excel2LaTeX"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   11
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
  Label2.Caption = Left$(Label2.Caption, Len(Label2.Caption) - 5) & "3.4.0"
  TextBox1.Text = "The development repository and the bug tracker for this package are hosted at" & vbCrLf & "        https://github.com/krlmlr/Excel2LaTeX" & vbCrLf & vbCrLf & _
                  "This work is distributed under the LaTeX Project Public License, version 1.3 or later, available at" & vbCrLf & "        http://www.latex-project.org/lppl.txt" & vbCrLf & vbCrLf & _
                  "Chelsea Hughes currently maintains this project (comprising Excel2LaTeX.xla and README.md) and will receive error reports at the project GitHub page (see above)."
  TextBox1.SelStart = 0
  TextBox1.SelLength = 0
End Sub
