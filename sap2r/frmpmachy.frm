
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
Begin VB.Form frmPMachY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pump Characteristics and Other Data"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   7200
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
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
      Left            =   1680
      TabIndex        =   9
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Frame frmMach 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   6375
      Begin VB.TextBox txtRatedEff 
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
         Left            =   4920
         TabIndex        =   4
         Top             =   2400
         Width           =   975
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
         Left            =   4920
         TabIndex        =   5
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox txtRatedH 
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
         Left            =   4920
         TabIndex        =   3
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtRatedQ 
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
         Left            =   4920
         TabIndex        =   2
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtPumpGd2 
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
         Left            =   4920
         TabIndex        =   0
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtMotorGD2 
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
         Left            =   4920
         TabIndex        =   1
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox cmbRatchet 
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
         TabIndex        =   6
         Text            =   "NO"
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Rated Efficiency of the Pump (%)"
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
         TabIndex        =   20
         Top             =   2520
         Width           =   4455
      End
      Begin VB.Label Label3 
         Caption         =   "Shut-off Head of  Pump (m)"
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
         TabIndex        =   19
         Top             =   3000
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "Rated Head of the Pump (m)"
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
         TabIndex        =   18
         Top             =   2040
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "Rated Discharge of the Pump (cum/sec)"
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
         TabIndex        =   17
         Top             =   1560
         Width           =   4455
      End
      Begin VB.Label Label6 
         Caption         =   "GD-Square Value  of the Pump (kgf - sqm)"
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
         TabIndex        =   16
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label Label7 
         Caption         =   "GD-Square Value  of the Motor (kgf - sqm)"
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
         TabIndex        =   15
         Top             =   1080
         Width           =   4455
      End
      Begin VB.Label Label8 
         Caption         =   "Non-Reverse Rotation Ratchet Provided ?"
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
         TabIndex        =   14
         Top             =   3600
         Width           =   4455
      End
   End
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
      Left            =   3120
      TabIndex        =   10
      Top             =   6000
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmDial 
      Left            =   6480
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frmYES 
      Height          =   1455
      Left            =   240
      TabIndex        =   11
      Top             =   4440
      Width           =   6375
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
         Left            =   3240
         TabIndex        =   8
         Top             =   720
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
         Left            =   1440
         TabIndex        =   7
         Top             =   720
         Width           =   1455
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
         Left            =   1200
         TabIndex        =   12
         Top             =   240
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmPMachY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OpFile As String
 Dim HSHUT As Single
 Dim QPMP As Single
 Dim HPMP As Single
 Dim EFF As Single
 Dim FKNR As Single
 Dim NPCH As Long
 Dim QQ(1 To 100) As Single
 Dim HH(1 To 100) As Single
 Dim ETA(1 To 100) As Single
 Dim WH(1 To 89) As Single
 Dim WB(1 To 89) As Single


Private Sub cmdCont_Click()
GD2PMP(NP) = Val(txtPumpGd2.Text)
GD2PM(NP) = Val(txtMotorGD2.Text)
QRPMP(NP) = Val(txtRatedQ.Text)
HRPMP(NP) = Val(txtRatedH.Text)
EFFRP(NP) = Val(txtRatedEff.Text)
HSHPMP(NP) = Val(txtShutOff.Text)
CODNRR(NP) = cmbRatchet.Text
   NPCH = NPUMPCH(NP)
   QPMP = QRPMP(NP)
   HPMP = HRPMP(NP)
   EFF = EFFRP(NP)
   HSHUT = HSHPMP(NP)
 For i = 1 To NPCH
   QQ(i + 1) = DCP(NP, i)
   HH(i + 1) = HEADP(NP, i)
   ETA(i + 1) = EFFP(NP, i)
 Next
 NPCH = NPCH + 1
 
 Call PMPYES(NPCH, QPMP, HPMP, EFF, HSHUT, QQ(1), HH(1), ETA(1), WH(1), WB(1), FKNR)
 
 For i = 1 To 89
 WHP(NP, i) = WH(i)
 WBP(NP, i) = WB(i)
Next
FKNRR(NP) = FKNR
Unload Me
frmGridPUMP.Enabled = True
End Sub
Private Sub cmdEnter_Click()
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
Private Sub cmdCancel_Click()
frmGridPUMP.MSFlexGrid1.Text = ""
frmGridPUMP.Enabled = True
Me.Hide
End Sub
Private Sub Form_Load()
cmbRatchet.AddItem ("YES")
cmbRatchet.AddItem ("NO")
 If Not OpenFile = "" Then
     txtPumpGd2.Text = GD2PMP(NP)
     txtMotorGD2.Text = GD2PM(NP)
     txtRatedQ.Text = QRPMP(NP)
     txtRatedH.Text = HRPMP(NP)
     txtRatedEff.Text = EFFRP(NP)
     txtShutOff.Text = HSHPMP(NP)
  End If
Left = 20
Top = 30
End Sub
Private Sub ReadIt()
Dim check As Variant
NPUMPCH(NP) = 0
Open OpFile For Input As #1
Do Until EOF(1)
 Input #1, check
 If Not IsNumeric(check) Then
  Close (1)
  Exit Sub
 End If
  NPUMPCH(NP) = NPUMPCH(NP) + 1
  DCP(NP, NPUMPCH(NP)) = check
  Input #1, HEADP(NP, NPUMPCH(NP)), EFFP(NP, NPUMPCH(NP))
Loop
Close (1)
End Sub

