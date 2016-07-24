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
  Label2.Caption = Left$(Label2.Caption, Len(Label2.Caption) - 5) & "3.4.0"
  TextBox1.Text = "The development repository and the bug tracker for this package are hosted at" & vbCrLf & "        https://github.com/krlmlr/Excel2LaTeX" & vbCrLf & vbCrLf & _
                  "This work may be distributed and/or modified under the conditions of the LaTeX Project Public License, either version 1.3 of this license or (at your option) any later version.  The latest version of this license is at" & vbCrLf & "        http://www.latex-project.org/lppl.txt" & vbCrLf & "and version 1.3 or later is part of all distributions of LaTeX version 2005/12/01 or later." & vbCrLf & vbCrLf & _
                  "This work has the LPPL maintenance status `maintained'." & vbCrLf & "The current maintainer of this work is Kirill Müller." & vbCrLf & "This work consists of the file Excel2LaTeX.xla."
  TextBox1.SelStart = 0
  TextBox1.SelLength = 0
End Sub
