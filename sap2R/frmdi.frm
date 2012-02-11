
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
Begin VB.Form frmDI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data for DI Pipe"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3465
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
      TabIndex        =   8
      Top             =   2760
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   4935
      Begin VB.TextBox txtTL 
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
         Left            =   3960
         TabIndex        =   2
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Thickness of the Lining (mm)"
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
         Left            =   0
         TabIndex        =   7
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.ComboBox cmbCML 
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
      TabIndex        =   1
      Text            =   "NO"
      Top             =   1080
      Width           =   855
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
      Text            =   "K8"
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
      TabIndex        =   3
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Cement Mortar Lining Provided ?"
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
      TabIndex        =   5
      Top             =   1200
      Width           =   3735
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
      TabIndex        =   4
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "frmDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmbCML_Click()
If cmbCML.Text = "YES" Then
 Frame1.Visible = True
Else
 Frame1.Visible = False
End If
End Sub
Private Sub cmdOK_Click()
 Dim PWTHICK As Single
 Dim TEMPV As Single
 CODECL(NP) = cmbClass.Text
 CODEDIL(NP) = cmbCML.Text
 If cmbCML.Text = "YES" Then
   THICKDIL(NP) = Val(txtTL.Text)
 Else
   THICKDIL(NP) = 0
 End If
 If CODECL(NP) = "K8" Then
  TEMPV = 8
 ElseIf CODECL(NP) = "K9" Then
  TEMPV = 9
 ElseIf CODECL(NP) = "K10" Then
  TEMPV = 10
 End If
 If PTYPE = "TYPEA" Then
  PDIA(1) = DIA
 End If
 PWTHICK = TEMPV * (0.5 + 0.001 * PDIA(NP)) + (1# / 12#) * THICKDIL(NP)
 If PWTHICK <= 0 Then
  MsgBox "Improper Data, Please Check !!"
  Exit Sub
 End If
 WV(NP) = 1440 / (Sqr(1 + (2.12 / 170) * (PDIA(NP) / PWTHICK)))
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

Private Sub Form_Activate()
If cmbCML.Text = "YES" Then
 Frame1.Visible = True
Else
 Frame1.Visible = False
End If
End Sub
Private Sub Form_Load()
Left = 20
Top = 30
cmbClass.AddItem "K8"
cmbClass.AddItem "K9"
cmbClass.AddItem "K10"
cmbCML.AddItem "YES"
cmbCML.AddItem "NO"
 cmbClass.Text = CODECL(NP)
 cmbCML.Text = CODEDIL(NP)
 txtTL.Text = THICKDIL(NP)
End Sub

