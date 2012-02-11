
'  Copyright 2011 Prof K.Sridharan
'  This file is part of SAP2
'
'    SAP2 is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    SAP2 is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'   You should have received a copy of the GNU General Public License
'    along with SAP2.  If not, see <http://www.gnu.org/licenses/>.


VERSION 5.00
Begin VB.Form frmCI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data for CI Pipe"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1875
   ScaleWidth      =   5745
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin VB.ComboBox cmbClass 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4200
      TabIndex        =   0
      Text            =   "LA"
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Pressure Class of the Pipe "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "frmCI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOK_Click()
 Dim PWTHICK As Single
 Dim TEMPV As Single
 CODECI(NP) = cmbClass.Text
 If CODECI(NP) = "LA" Then
  TEMPV = 1#
 ElseIf CODECI(NP) = "A" Then
  TEMPV = 1.1
 ElseIf CODECI(NP) = "B" Then
  TEMPV = 1.2
 ElseIf CODECI(NP) = "C" Then
  TEMPV = 1.3
 ElseIf CODECI(NP) = "D" Then
  TEMPV = 1.4
 End If
 If PTYPE = "TYPEA" Then
  PDIA(1) = DIA
 End If
 PWTHICK = TEMPV * (10 / 12) * (7 + 0.02 * PDIA(NP))
 If PWTHICK <= 0 Then
  MsgBox "Improper Data, Please Check !!"
  Exit Sub
 End If
 WV(NP) = 1440 / (Sqr(1 + (2.12 / 100) * (PDIA(NP) / PWTHICK)))
 WVA = WV(NP)
 
 If PTYPE = "TYPEA" Then
   frmTranMainA.Enabled = True
 Else
   frmGridWave.Enabled = True
 End If
 MDIForm1.mnuExec.Item(10).Enabled = True
 MDIForm1.tbrMain.Buttons(6).Enabled = True
 Unload Me
 End Sub

Private Sub Command1_Click()
If PTYPE = "TYPEA" Then
   frmTranMainA.Enabled = True
   frmTranMainA.cmbPWV.Text = ""
 Else
   frmGridWave.Enabled = True
   frmGridWave.Combo1.Text = ""
 End If
 MDIForm1.mnuExec.Item(10).Enabled = True
 MDIForm1.tbrMain.Buttons(6).Enabled = True
 frmGridWave.Combo1.Text = ""
 Unload Me
End Sub

Private Sub Form_Load()
Left = 20
Top = 30
cmbClass.AddItem "LA"
cmbClass.AddItem "A"
cmbClass.AddItem "B"
cmbClass.AddItem "C"
cmbClass.AddItem "D"
cmbClass.Text = CODECI(NP)
End Sub

