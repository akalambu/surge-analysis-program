
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
Begin VB.Form frmSRVB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data for SRV"
   ClientHeight    =   6450
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
   ScaleHeight     =   6450
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNLSRV 
      Height          =   315
      Left            =   4920
      TabIndex        =   0
      Text            =   "1"
      Top             =   240
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data for Location : 1"
      ForeColor       =   &H00FF0000&
      Height          =   4935
      Left            =   120
      TabIndex        =   12
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
         Left            =   3000
         TabIndex        =   10
         Top             =   4440
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
         Left            =   1800
         TabIndex        =   9
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox txtPN 
         Height          =   315
         Left            =   4200
         TabIndex        =   1
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtNSRV 
         Height          =   315
         Left            =   4200
         TabIndex        =   3
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtSSRV 
         Height          =   315
         Left            =   4200
         TabIndex        =   4
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtPIL 
         Height          =   315
         Left            =   4200
         TabIndex        =   5
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txtLPPS 
         Height          =   315
         Left            =   4200
         TabIndex        =   6
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox txtHPPS 
         Height          =   315
         Left            =   4200
         TabIndex        =   7
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox txtCT 
         Height          =   315
         Left            =   4200
         TabIndex        =   8
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Chainage (m)"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label Label7 
         Caption         =   "Pipe Number at Valve Location"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label label1 
         Caption         =   "Number of Surge Relief Valves "
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "Size of  Surge Relief Valves (mm)"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   3615
      End
      Begin VB.Label Label3 
         Caption         =   "Pipe Invert Level at the Valve (RL, m)"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   2520
         Width           =   3855
      End
      Begin VB.Label Label4 
         Caption         =   "Low Pressure Pilot Setting (m)"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   3000
         Width           =   3855
      End
      Begin VB.Label Label5 
         Caption         =   "High Pressure Pilot Setting (m)"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   3480
         Width           =   3855
      End
      Begin VB.Label Label6 
         Caption         =   "Closure Time (sec)"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   3960
         Width           =   3735
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label label1 
      Caption         =   "Number of Locations of Surge Relief Valve"
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   20
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "frmSRVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim jj As Integer
Private Sub cmdNext_Click()
Save
If jj = NLSRV Then Exit Sub
jj = jj + 1
Load
End Sub
Private Sub cmdOK_Click()
Save
Me.Hide
frmProtB.Enabled = True
frmProtB.SetFocus
End Sub
Private Sub cmdPrev_Click()
Save
If jj = 1 Then Exit Sub
jj = jj - 1
Load
End Sub
Private Sub Form_Load()
Left = 20
Top = 30
jj = 1
cmdPrev.Enabled = False
cmdNext.Enabled = False
If Not OpenFile = "" Then
If NLSRV > 1 Then
 txtNLSRV.Text = NLSRV
End If
If NLSRV > 1 Then
   cmdPrev.Enabled = True
   cmdNext.Enabled = True
 Else
   cmdPrev.Enabled = False
   cmdNext.Enabled = False
 End If
Load
End If
End Sub
Private Sub txtNLSRV_Change()
NLSRV = Val(txtNLSRV.Text)
 If NLSRV > 1 Then
   cmdPrev.Enabled = True
   cmdNext.Enabled = True
 Else
   cmdPrev.Enabled = False
   cmdNext.Enabled = False
 End If
 jj = 1
 Load
End Sub
Sub Load()
Frame1.Caption = "Data for Location :" & jj
txtPN.Text = SRV_Data(jj, 0)
txtCH.Text = SRV_Data(jj, 1)
txtNSRV.Text = SRV_Data(jj, 2)
txtSSRV.Text = SRV_Data(jj, 3)
txtPIL.Text = SRV_Data(jj, 4)
txtLPPS.Text = SRV_Data(jj, 5)
txtHPPS.Text = SRV_Data(jj, 6)
txtCT.Text = SRV_Data(jj, 7)
If jj = NLSRV Then
 cmdOK.Caption = "OK"
End If
End Sub
Sub Save()
NLSRV = Val(txtNLSRV.Text)
SRV_Data(jj, 0) = Val(txtPN.Text)
SRV_Data(jj, 1) = Val(txtCH.Text)
SRV_Data(jj, 2) = Val(txtNSRV.Text)
SRV_Data(jj, 3) = Val(txtSSRV.Text)
SRV_Data(jj, 4) = Val(txtPIL.Text)
SRV_Data(jj, 5) = Val(txtLPPS.Text)
SRV_Data(jj, 6) = Val(txtHPPS.Text)
SRV_Data(jj, 7) = Val(txtCT.Text)
End Sub
