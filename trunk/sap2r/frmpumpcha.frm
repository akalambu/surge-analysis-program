
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
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPumpChA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pump Characteristics Data"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4860
   ScaleWidth      =   7005
   Begin VB.CommandButton cmdCont 
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
      Left            =   2520
      TabIndex        =   5
      Top             =   4320
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmDial 
      Left            =   6480
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frmYES 
      Height          =   2175
      Left            =   360
      TabIndex        =   9
      Top             =   1920
      Width           =   6015
      Begin VB.CommandButton cmdGetIt 
         Caption         =   "Get From File"
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
         TabIndex        =   4
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton cmdEnter 
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
         Left            =   960
         TabIndex        =   3
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtShutOff 
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
         Left            =   3000
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Discharge, Head and Efficiency Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   11
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label Label3 
         Caption         =   "Shut-off Head of  Pump "
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
         TabIndex        =   10
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame frmNo 
      Height          =   735
      Left            =   360
      TabIndex        =   7
      Top             =   840
      Width           =   6015
      Begin VB.ComboBox cmbType 
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
         TabIndex        =   1
         Text            =   "RADIAL"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Type of Pump Characteristics "
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
         TabIndex        =   8
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.ComboBox cmbPumpCh 
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
      Left            =   5040
      TabIndex        =   0
      Text            =   "NO"
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Data for Pump Characteristics Available ?"
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
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   4335
   End
End
Attribute VB_Name = "frmPumpChA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OpFile As String
Private Sub cmbPumpCh_click()
 CODPCH = cmbPumpCh.Text
If cmbPumpCh.Text = "NO" Then
  frmNo.Visible = True
 frmYES.Visible = False
Else
 frmNo.Visible = False
 frmYES.Visible = True
End If
End Sub
Private Sub cmdCont_Click()
SHUOFF = Val(txtShutOff.Text)
Me.Hide
MDIForm1.mnuExec.Item(10).Enabled = True
MDIForm1.tbrMain.Buttons(6).Enabled = True
End Sub
Private Sub cmdEnter_Click()
NP = 1
frmGridPumpCH.Show
End Sub
Private Sub cmdgetIt_Click()
  Dim xxx As String
   cmDial.Filter = "pmp (*.pmp)|*.pmp"
   cmDial.FileName = ""
      cmDial.ShowOpen
      OpFile = cmDial.FileName
      xxx = Dir(OpFile)
      If xxx = "" Then
       MsgBox "File Not Found"
      ElseIf Not OpFile = "" Then
       ReadIt
      End If
End Sub
Private Sub Form_Load()
Left = 20
Top = 30
cmbPumpCh.AddItem "YES"
cmbPumpCh.AddItem "NO"
cmbType.AddItem "RADIAL"
cmbType.AddItem "MIXED"
cmbType.AddItem "AXIAL"
If Not CODPCH = "YES" Then
  cmbPumpCh.Text = "NO"
  frmNo.Visible = True
  If Not TYPCH = "" Then
    cmbType.Text = TYPCH
  Else
    cmbType.Text = "RADIAL"
  End If
  frmYES.Visible = False
Else
  cmbPumpCh.Text = "YES"
  frmNo.Visible = False
  frmYES.Visible = True
End If
End Sub
Private Sub ReadIt()
Dim check As Variant
NPUMPCH(1) = 0
Open OpFile For Input As #1
Do Until EOF(1)
 Input #1, check
 If Not IsNumeric(check) Then
  Close (1)
  Exit Sub
 End If
  NPUMPCH(1) = NPUMPCH(1) + 1
  DCP(1, NPUMPCH(1)) = check
  Input #1, HEADP(1, NPUMPCH(1)), EFFP(1, NPUMPCH(1))
 Loop
End Sub

