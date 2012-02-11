
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
Begin VB.Form frmHGL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Specific Details at Nodes"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   5985
   Begin VB.CommandButton cmdBack 
      Caption         =   "<< &Back"
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
      Left            =   1200
      TabIndex        =   9
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton cmdCont 
      Caption         =   "&Continue >>"
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
      TabIndex        =   10
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton cmdHGL 
      Caption         =   "Enter"
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
      Index           =   8
      Left            =   4440
      TabIndex        =   8
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdHGL 
      Caption         =   "Enter"
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
      Index           =   7
      Left            =   4440
      TabIndex        =   7
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdHGL 
      Caption         =   "Enter"
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
      Index           =   6
      Left            =   4440
      TabIndex        =   6
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdHGL 
      Caption         =   "Enter"
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
      Index           =   5
      Left            =   4440
      TabIndex        =   5
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdHGL 
      Caption         =   "Enter"
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
      Index           =   4
      Left            =   4440
      TabIndex        =   4
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdHGL 
      Caption         =   "Enter"
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
      Index           =   3
      Left            =   4440
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdHGL 
      Caption         =   "Enter"
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
      Index           =   2
      Left            =   4440
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdHGL 
      Caption         =   "Enter"
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
      Index           =   1
      Left            =   4440
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdHGL 
      Caption         =   "Enter"
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
      Index           =   0
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Booster Details"
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
      Index           =   8
      Left            =   240
      TabIndex        =   19
      Top             =   4920
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Pump Details "
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
      Index           =   7
      Left            =   240
      TabIndex        =   18
      Top             =   4320
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Data for Local Obstructions"
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
      Index           =   6
      Left            =   240
      TabIndex        =   17
      Top             =   3720
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Data for Condensor "
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
      Index           =   5
      Left            =   240
      TabIndex        =   16
      Top             =   3120
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Source Reservoir Data "
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
      Index           =   4
      Left            =   240
      TabIndex        =   15
      Top             =   2520
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Delivery Reservoir Data"
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
      Index           =   3
      Left            =   240
      TabIndex        =   14
      Top             =   1920
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "HGL at Dividing Junctions"
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
      Index           =   2
      Left            =   240
      TabIndex        =   13
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "HGL at Combining Junctions"
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
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "HGL at Ordinary Nodes"
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
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmHGL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdBack_Click()
Me.Hide
MDIForm1.mnuExec.Item(10).Enabled = False
MDIForm1.tbrMain.Buttons(6).Enabled = False
frmTranMainB.Show
End Sub
Private Sub cmdCont_Click()
Me.Hide
frmWaveTrip.Show
End Sub
Private Sub cmdHGL_Click(Index As Integer)
Select Case Index
 Case 0
  Iflag_But(3) = 1
  cmdHGL(0).Caption = "Change"
  frmGridORD.Show
 Case 1
  Iflag_But(4) = 1
  cmdHGL(1).Caption = "Change"
  frmGridCJN.Show
 Case 2
  Iflag_But(5) = 1
  cmdHGL(2).Caption = "Change"
  frmGridDJN.Show
 Case 3
  Iflag_But(6) = 1
  cmdHGL(3).Caption = "Change"
  frmGridRES.Show
 Case 4
  Iflag_But(7) = 1
  cmdHGL(4).Caption = "Change"
  frmGridSOV.Show
 Case 5
  Iflag_But(8) = 1
  cmdHGL(5).Caption = "Change"
  frmGridCDS.Show
 Case 6
  Iflag_But(9) = 1
  cmdHGL(6).Caption = "Change"
  frmGridOBS.Show
 Case 7
  Iflag_But(10) = 1
  cmdHGL(7).Caption = "Change"
  frmGridPUMP.Show
 Case 8
  Iflag_But(11) = 1
  cmdHGL(8).Caption = "Change"
  frmGridBOOST.Show
End Select
frmHGL.Enabled = False
MDIForm1.mnuExec.Item(10).Enabled = False
MDIForm1.tbrMain.Buttons(6).Enabled = False
End Sub
Private Sub Form_Activate()
If NORD = 0 Then
 cmdHGL(0).Enabled = False
Else
 cmdHGL(0).Enabled = True
End If
If NCJN = 0 Then
 cmdHGL(1).Enabled = False
Else
 cmdHGL(1).Enabled = True
End If
If NDJN = 0 Then
 cmdHGL(2).Enabled = False
Else
 cmdHGL(2).Enabled = True
End If
If NRES = 0 Then
 cmdHGL(3).Enabled = False
Else
 cmdHGL(3).Enabled = True
End If
If NSOU = 0 Then
 cmdHGL(4).Enabled = False
Else
 cmdHGL(4).Enabled = True
End If
If NCDS = 0 Then
 cmdHGL(5).Enabled = False
Else
 cmdHGL(5).Enabled = True
End If
If NOBS = 0 Then
 cmdHGL(6).Enabled = False
Else
 cmdHGL(6).Enabled = True
End If
If NPMP = 0 Then
 cmdHGL(7).Enabled = False
Else
 cmdHGL(7).Enabled = True
End If
If NBST = 0 Then
 cmdHGL(8).Enabled = False
Else
 cmdHGL(8).Enabled = True
End If
End Sub
Private Sub Form_Load()
Left = 20
Top = 30
End Sub
