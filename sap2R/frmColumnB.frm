
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
Begin VB.Form frmColumnB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data for Column Separation"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbCS 
      Height          =   360
      Left            =   4920
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data for Location : 1"
      ForeColor       =   &H00FF0000&
      Height          =   3015
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   5895
      Begin VB.TextBox txtCh 
         Height          =   315
         Left            =   4200
         TabIndex        =   2
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "Previous"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtPN 
         Height          =   315
         Left            =   4200
         TabIndex        =   1
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtPIL 
         Height          =   315
         Left            =   4200
         TabIndex        =   3
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Chainage (m)"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label Label7 
         Caption         =   "Pipe Number "
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label3 
         Caption         =   "Pipe Invert Level (RL, m)"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   3855
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label label1 
      Caption         =   "Number of Locations of Column Separation"
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "frmColumnB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim jj As Integer
Private Sub cmdNext_Click()
Save
If jj = NLCS Then Exit Sub
jj = jj + 1
Load
End Sub
Private Sub cmdOK_Click()
Save
Me.Hide
If NLCS = 0 Then
frmAnalyB.cmbColSep.Text = "NO"
End If
frmAnalyB.Enabled = True
frmAnalyB.SetFocus
MDIForm1.mnuExec.Item(10).Enabled = True
MDIForm1.tbrMain.Buttons(6).Enabled = True
End Sub
Private Sub cmdPrev_Click()
Save
If jj = 1 Then Exit Sub
jj = jj - 1
Load
End Sub
Private Sub Form_Load()
cmbCS.AddItem "0"
cmbCS.AddItem "1"
cmbCS.AddItem "2"
Left = 20
Top = 30
jj = 1
If NLCS > 0 Then
 cmbCS.Text = NLCS
 End If
If NLCS > 1 Then
   cmdPrev.Enabled = True
   cmdNext.Enabled = True
 Else
   cmdPrev.Enabled = False
   cmdNext.Enabled = False
 End If
 If Not OpenFile = "" Then
 Load
 End If
 If cmbCS.Text = 0 Then
  Frame1.Visible = False
 Else
  Frame1.Visible = True
 End If
End Sub
Private Sub cmbCs_Click()
 If cmbCS.Text = 0 Then
  Frame1.Visible = False
 Else
  Frame1.Visible = True
 End If
 NLCS = Val(cmbCS.Text)
 If NLCS > 1 Then
   cmdPrev.Enabled = True
   cmdNext.Enabled = True
 Else
   cmdPrev.Enabled = False
   cmdNext.Enabled = False
 End If
 jj = 1
 Load
End Sub
Private Sub cmbcs_change()
cmbCs_Click
End Sub
Sub Load()
Frame1.Caption = "Data for Location :" & jj
txtPN.Text = CS_Data(jj, 0)
txtCh.Text = CS_Data(jj, 1)
txtPIL.Text = CS_Data(jj, 2)
End Sub
Sub Save()
CS_Data(jj, 0) = Val(txtPN.Text)
CS_Data(jj, 1) = Val(txtCh.Text)
CS_Data(jj, 2) = Val(txtPIL.Text)
End Sub
