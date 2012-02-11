
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
Begin VB.Form frmVOB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Details of Valve with/without Bypass Outlet"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   5610
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
      Left            =   2280
      TabIndex        =   11
      Top             =   6000
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   120
      TabIndex        =   17
      Top             =   2640
      Width           =   5415
      Begin VB.TextBox txtDL 
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
         Left            =   4320
         TabIndex        =   7
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtDLBD 
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
         Left            =   4320
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtTCB 
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
         Left            =   4320
         TabIndex        =   10
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox txtDOP 
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
         Left            =   4320
         TabIndex        =   9
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtTO 
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
         Left            =   4320
         TabIndex        =   8
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtSB 
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
         Left            =   4320
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Delay in Bypass Valve Opening (sec)"
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
         Left            =   120
         TabIndex        =   23
         Top             =   1200
         Width           =   3975
      End
      Begin VB.Label Label10 
         Caption         =   "Bypass Outlet Discharge Level (RL, m) "
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
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label9 
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
         Left            =   120
         TabIndex        =   21
         Top             =   2640
         Width           =   3375
      End
      Begin VB.Label Label8 
         Caption         =   "Duration of Open Position (sec)"
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
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   3375
      End
      Begin VB.Label Label7 
         Caption         =   "Time of Opening (sec) "
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
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "Size of Bypass Outlet (mm)"
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
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   3975
      End
   End
   Begin VB.ComboBox cmbBOP 
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
      Left            =   4440
      TabIndex        =   4
      Text            =   "NO"
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtTCM 
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
      Left            =   4440
      TabIndex        =   3
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtDCM 
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
      Left            =   4440
      TabIndex        =   2
      Top             =   1000
      Width           =   855
   End
   Begin VB.TextBox txtCH 
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
      Left            =   4440
      TabIndex        =   1
      Top             =   520
      Width           =   855
   End
   Begin VB.TextBox txtPN 
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
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Bypass Outlet Provided ?"
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
      TabIndex        =   16
      Top             =   2040
      Width           =   3975
   End
   Begin VB.Label Label4 
      Caption         =   "Time of Closure of  Valve (sec)"
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
      TabIndex        =   15
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "Delay in Start of Closure of  Valve (sec)"
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
      TabIndex        =   14
      Top             =   960
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "Chainage ( m)"
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
      TabIndex        =   13
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Pipe Number at  Valve Location"
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
      TabIndex        =   12
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmVOB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmbBOP_Click()
If cmbBOP.Text = "NO" Then
 Frame1.Visible = False
ElseIf cmbBOP.Text = "YES" Then
  Frame1.Visible = True
End If
End Sub
Private Sub cmdOK_Click()
CODIV = "YES"
NPIV = Val(txtPN.Text)
CHIV = Val(txtCh.Text)
DLYIV = Val(txtDCM.Text)
TCIV = Val(txtTCM.Text)
CODBIV = cmbBOP.Text
If CODBIV = "YES" Then
 HDELB = Val(txtDLBD.Text)
 SZBIV = Val(txtSB.Text)
 DLYBIV = Val(txtDL.Text)
 TOBIV = Val(txtTO.Text)
 TOPGB = Val(txtDOP.Text)
 TCBIV = Val(txtTCB.Text)
End If
Me.Hide
frmProtB.Enabled = True
frmProtB.SetFocus
End Sub
Private Sub Form_Activate()
If Not cmbBOP.Text = "YES" Then
 Frame1.Visible = False
Else
  Frame1.Visible = True
End If
End Sub
Private Sub Form_Load()
Left = 20
Top = 30
cmbBOP.AddItem "YES"
cmbBOP.AddItem "NO"
If Not OpenFile = "" Then
txtPN.Text = NPIV
txtCh.Text = CHIV
txtDCM.Text = DLYIV
txtTCM.Text = TCIV
cmbBOP.Text = CODBIV
If CODBIV = "YES" Then
 txtDLBD.Text = HDELB
 txtSB.Text = SZBIV
 txtDL.Text = DLYBIV
 txtTO.Text = TOBIV
 txtDOP.Text = TOPGB
 txtTCB.Text = TCBIV
End If
End If
End Sub

