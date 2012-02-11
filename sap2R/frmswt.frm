
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
Begin VB.Form frmswt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Swing Check Vlave "
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
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
      Left            =   2760
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtSCV 
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
      Left            =   5160
      TabIndex        =   0
      Text            =   "0"
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Delay in Closure of Swing Check Valve (sec)"
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
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   4695
   End
End
Attribute VB_Name = "frmswt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  If Iflag_PB = 1 Then
    DLYPH(NP) = Val(txtSCV.Text)
    If PTYPE = "TYPEA" Then
     frmPumpA.Enabled = True
     MDIForm1.mnuExec.Item(10).Enabled = True
     MDIForm1.tbrMain.Buttons(6).Enabled = True
    Else
     frmGridPUMP.Enabled = True
    End If
  ElseIf Iflag_PB = 2 Then
    DLYBS(NP) = Val(txtSCV.Text)
    frmGridBOOST.Enabled = True
  End If
Unload Me
End Sub
Private Sub Form_Load()
Left = 20
Top = 30
  If Iflag_PB = 1 Then
    txtSCV.Text = DLYPH(NP)
  ElseIf Iflag_PB = 2 Then
    txtSCV.Text = DLYBS(NP)
  End If
End Sub
