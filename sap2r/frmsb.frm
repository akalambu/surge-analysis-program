
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
Begin VB.Form frmsb 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " NRV With Uniform Speed Closure - B "
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   " &OK"
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
      Left            =   1800
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtTC 
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
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Time of Closure (sec)"
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
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "frmsb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 If Iflag_PB = 1 Then
 TCLOSEP(NP) = Val(txtTC.Text)
 If PTYPE = "TYPEA" Then
  frmPumpA.Enabled = True
  MDIForm1.mnuExec.Item(10).Enabled = True
  MDIForm1.tbrMain.Buttons(6).Enabled = True
 Else
  frmGridPUMP.Enabled = True
 End If
 ElseIf Iflag_PB = 2 Then
 TCLOSEB(NP) = Val(txtTC.Text)
 frmGridBOOST.Enabled = True
 End If
Unload Me
End Sub
Private Sub Form_Load()
Left = 20
Top = 30
 If Iflag_PB = 1 Then
 txtTC.Text = TCLOSEP(NP)
 ElseIf Iflag_PB = 2 Then
 txtTC.Text = TCLOSEB(NP)
 End If
End Sub
