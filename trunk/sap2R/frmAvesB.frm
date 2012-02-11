
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
Begin VB.Form frmAvesB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Air Vessel Data"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtAVP 
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
      Left            =   4320
      TabIndex        =   0
      Text            =   " "
      Top             =   240
      Width           =   975
   End
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
      Left            =   2280
      TabIndex        =   8
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox txtCPS 
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
      Left            =   4320
      TabIndex        =   3
      Text            =   " "
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtSP 
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
      Left            =   4320
      TabIndex        =   2
      Text            =   " "
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtGE 
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
      Left            =   4320
      TabIndex        =   1
      Text            =   " "
      Top             =   960
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   " Type of Air Vessel"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      TabIndex        =   12
      Top             =   3600
      Width           =   4935
      Begin VB.TextBox txtOS 
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
         TabIndex        =   7
         Text            =   " "
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton optType2 
         Caption         =   " Type II"
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
         Left            =   3240
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optType1 
         Caption         =   " Type I"
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
         Left            =   600
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   " Orfice Size (mm)"
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
         Left            =   480
         TabIndex        =   13
         Top             =   960
         Width           =   2055
      End
   End
   Begin VB.CheckBox chkNRV 
      Caption         =   " NRV is Provided on the Transmission Main"
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
      Top             =   2880
      Width           =   5055
   End
   Begin VB.Label Label5 
      Caption         =   "The Pipe on Which Air Vessel is Located (Pipe Number)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   14
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   " Connecting Pipe Size (mm)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   " Size Parameter of the Air Vessel, KAV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Ground Elevation at Air Vessel Location (RL, m)"
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
      TabIndex        =   9
      Top             =   960
      Width           =   3375
   End
End
Attribute VB_Name = "frmAvesB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CODEAC = "YES"
GLAC = Val(txtGE.Text)
ACC = Val(txtSP.Text)
DCAC = Val(txtCPS.Text)
NPAC = Val(txtAVP.Text)
If chkNRV.Value = 1 Then
 ACNRV = "YES"
Else
 ACNRV = "NO"
End If
If optType1.Value = True Then
 KACTYP = 1
Else
 KACTYP = 2
End If
DORBY = Val(txtOS.Text)
Me.Hide
frmProtB.Enabled = True
frmProtB.SetFocus
End Sub
Private Sub Form_Load()
Left = 20
Top = 30
If CODEAC = "YES" And Not OpenFile = "" Then
 txtGE.Text = GLAC
 txtSP.Text = ACC
 txtCPS.Text = DCAC
 txtAVP.Text = NPAC
 If ACNRV = "YES" Then
  chkNRV.Value = 1
 Else
  chkNRV.Value = 0
 End If
 If KACTYP = 1 Then
  optType1.Value = True
 Else
  optType2.Value = True
 End If
  txtOS.Text = DORBY
Else
optType1.Value = True
End If
End Sub
Private Sub optType1_Click()
Label4.Caption = "Orifice Size (mm)"
End Sub
Private Sub optType2_Click()
Label4.Caption = "Bypass Size (mm)"
End Sub

