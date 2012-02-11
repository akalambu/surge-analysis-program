
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
Begin VB.Form frmtsb 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NRV with Two Speed Closure - B"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   " Ok"
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
      Left            =   2040
      TabIndex        =   2
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtTC10 
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
      Left            =   3960
      TabIndex        =   1
      Text            =   "0"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txtTC90 
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
      Left            =   3960
      TabIndex        =   0
      Text            =   "0"
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   " Time of Next 10% Closure (sec)"
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
      TabIndex        =   4
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   " Time of First 90% Closure (sec)"
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
      TabIndex        =   3
      Top             =   480
      Width           =   3615
   End
End
Attribute VB_Name = "frmtsb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 If Iflag_PB = 1 Then
   TRAPIDP(NP) = Val(txtTC90.Text)
   TSLOWP(NP) = Val(txtTC10.Text)
   If PTYPE = "TYPEA" Then
     frmPumpA.Enabled = True
     MDIForm1.mnuExec.Item(10).Enabled = True
     MDIForm1.tbrMain.Buttons(6).Enabled = True
   Else
     frmGridPUMP.Enabled = True
   End If
 ElseIf Iflag_PB = 2 Then
   TRAPIDB(NP) = Val(txtTC90.Text)
   TSLOWB(NP) = Val(txtTC10.Text)
   frmGridBOOST.Enabled = True
 End If
Unload Me
End Sub
Private Sub Form_Load()
Left = 20
Top = 30
 If Iflag_PB = 1 Then
   txtTC90.Text = TRAPIDP(NP)
   txtTC10.Text = TSLOWP(NP)
 ElseIf Iflag_PB = 2 Then
   txtTC90.Text = TRAPIDB(NP)
   txtTC10.Text = TSLOWB(NP)
 End If
End Sub
