
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
Begin VB.Form frmSRVA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data for SRV"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
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
   ScaleHeight     =   5415
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCT 
      Height          =   315
      Left            =   4320
      TabIndex        =   5
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox txtHPPS 
      Height          =   315
      Left            =   4320
      TabIndex        =   4
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtLPPS 
      Height          =   315
      Left            =   4320
      TabIndex        =   3
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txtPIL 
      Height          =   315
      Left            =   4320
      TabIndex        =   2
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtSSRV 
      Height          =   315
      Left            =   4320
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtNSRV 
      Height          =   315
      Left            =   4320
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Closure Time (sec)"
      Height          =   615
      Left            =   240
      TabIndex        =   12
      Top             =   3720
      Width           =   3735
   End
   Begin VB.Label Label5 
      Caption         =   "High Pressure Pilot Setting (m)"
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   3000
      Width           =   3855
   End
   Begin VB.Label Label4 
      Caption         =   "Low Pressure Pilot Setting (m)"
      Height          =   615
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "Pipe Invert Level at the valve (RL, m)"
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "Size of  Surge Relief Valves (mm)"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Number of Surge Relief Valves "
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "frmSRVA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
SRV_Data(1, 2) = Val(txtNSRV.Text)
SRV_Data(1, 3) = Val(txtSSRV.Text)
SRV_Data(1, 4) = Val(txtPIL.Text)
SRV_Data(1, 5) = Val(txtLPPS.Text)
SRV_Data(1, 6) = Val(txtHPPS.Text)
SRV_Data(1, 7) = Val(txtCT.Text)
Me.Hide
frmProtA.Enabled = True
frmProtA.SetFocus
End Sub
Private Sub Form_Load()
Left = 20
Top = 30
If CODESV = "YES" And Not OpenFile = "" Then
 txtNSRV.Text = SRV_Data(1, 2)
 txtSSRV.Text = SRV_Data(1, 3)
 txtPIL.Text = SRV_Data(1, 4)
 txtLPPS.Text = SRV_Data(1, 5)
 txtHPPS.Text = SRV_Data(1, 6)
 txtCT.Text = SRV_Data(1, 7)
End If
End Sub
