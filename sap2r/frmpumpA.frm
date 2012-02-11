
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
Begin VB.Form frmPumpA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pump Details"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6930
   ScaleWidth      =   7110
   Begin VB.Frame frmPMF 
      Height          =   735
      Left            =   240
      TabIndex        =   20
      Top             =   2160
      Width           =   4935
      Begin VB.ComboBox cmbMach 
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
         Left            =   3960
         TabIndex        =   21
         Text            =   "NO"
         Top             =   200
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Pumping Machinery Finalised ? "
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
         TabIndex        =   22
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.CommandButton cmdNRV 
      Caption         =   "Select"
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
      Left            =   4080
      TabIndex        =   2
      Top             =   1680
      Width           =   975
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
      Left            =   3480
      TabIndex        =   9
      Top             =   6240
      Width           =   1455
   End
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
      Left            =   1680
      TabIndex        =   8
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Frame frmMach 
      Caption         =   "Pump Details"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   240
      TabIndex        =   13
      Top             =   3000
      Visible         =   0   'False
      Width           =   6375
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
         Left            =   5160
         TabIndex        =   7
         Text            =   "NO"
         Top             =   2400
         Width           =   855
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
         Left            =   5040
         TabIndex        =   6
         Top             =   1920
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
         Left            =   5040
         TabIndex        =   5
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtPumpSp 
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
         Left            =   5040
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtPumpEff 
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
         Left            =   5040
         TabIndex        =   3
         Top             =   480
         Width           =   975
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
         TabIndex        =   18
         Top             =   2520
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
         TabIndex        =   17
         Top             =   2040
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
         Top             =   1560
         Width           =   4455
      End
      Begin VB.Label Label5 
         Caption         =   "Rated Speed of the Pump (rpm)"
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
         Width           =   3375
      End
      Begin VB.Label Label4 
         Caption         =   "Rated Pump Efficiency (%)"
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
         Top             =   600
         Width           =   3375
      End
   End
   Begin VB.Frame frmNoPumps 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox txtPumpNo 
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
         Left            =   3840
         TabIndex        =   1
         Text            =   "1"
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "           No of Working Pumps    "
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
         Left            =   -360
         TabIndex        =   12
         Top             =   200
         Width           =   3975
      End
   End
   Begin VB.ComboBox cmbPumps 
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
      ItemData        =   "frmpumpA.frx":0000
      Left            =   4200
      List            =   "frmpumpA.frx":0002
      TabIndex        =   0
      Text            =   "NO"
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Type of Pump House NRV  "
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
      Left            =   360
      TabIndex        =   19
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "No of Working Pumps Known ? "
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
      Left            =   360
      TabIndex        =   10
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frmPumpA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbMach_Click()
If Not cmbMach.Text = "NO" Then
 frmMach.Visible = True
Else
 frmMach.Visible = False
End If
End Sub
Private Sub cmbPumps_Click()
If Not cmbPumps.Text = "NO" Then
 frmNoPumps.Visible = True
 frmPMF.Visible = True
Else
 frmNoPumps.Visible = False
 frmPMF.Visible = False
End If
End Sub
Private Sub cmdBack_Click()
Me.Hide
frmTranMainA.Show
End Sub
Private Sub cmdCont_Click()
If cmbPumps.Text = "YES" Then
NPUMP = Val(txtPumpNo.Text)
Else
NPUMP = 1
End If
If cmbMach.Text = "YES" Then
 EFFA = Val(txtPumpEff.Text)
 ISPEED = Val(txtPumpSp.Text)
 GDSQP = Val(txtPumpGd2.Text)
 GDSQM = Val(txtMotorGD2.Text)
 CODNRRA = cmbRatchet.Text
Else
 EFFA = 85
 ISPEED = 1440
 bkw = (746# * (QR / NPUMP) * REFH * 1.2) / (75 * 0.85)
 GDSQP = (1# / 3#) * 540# * ((bkw / ISPEED) ^ 1.4)
 GDSQM = (2# / 3#) * 540# * ((bkw / ISPEED) ^ 1.4)
 CODNRRA = "NO"
End If

Me.Hide
frmAnalyA.Show
End Sub
Private Sub cmdNRV_Click()
 Iflag_PB = 1
 NP = 1
 frmPumpA.Enabled = False
 MDIForm1.mnuExec.Item(10).Enabled = False
 MDIForm1.tbrMain.Buttons(6).Enabled = False
 frmListP.Show
End Sub
Private Sub Form_Load()
Left = 20
Top = 30
cmbPumps.AddItem "YES"
cmbPumps.AddItem "NO"
cmbMach.AddItem "YES"
cmbMach.AddItem "NO"
cmbRatchet.AddItem "YES"
cmbRatchet.AddItem "NO"
If Not OpenFile = "" Then
If CODENP = "YES" Then
 cmbPumps.Text = "YES"
 frmNoPumps.Visible = True
 frmPMF.Visible = True
Else
  cmbPumps.Text = "NO"
 frmNoPumps.Visible = False
 frmPMF.Visible = False
End If
If CODEPM = "YES" Then
 cmbMach.Text = "YES"
 frmMach.Visible = True
Else
 cmbMach.Text = "NO"
 frmMach.Visible = False
End If
End If
End Sub
Private Sub txtPumpNo_Change()
If cmbPumps.Text = "YES" Then
NPUMP = Val(txtPumpNo.Text)
Else
NPUMP = 1
End If
End Sub
