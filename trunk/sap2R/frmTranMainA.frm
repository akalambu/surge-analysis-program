
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
Begin VB.Form frmTranMainA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Details of Transmission Main "
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5.49
   ScaleMode       =   5  'Inch
   ScaleWidth      =   5.104
   Begin VB.Frame frmPHP 
      Height          =   1695
      Left            =   120
      TabIndex        =   23
      Top             =   5520
      Visible         =   0   'False
      Width           =   6855
      Begin VB.ComboBox cmbPWVD 
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
         TabIndex        =   10
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtPumpDelDia 
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
         Left            =   5160
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Pump Delivery Pipe Diameter (mm)      "
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
         Left            =   480
         TabIndex        =   25
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label Label9 
         Caption         =   "Delivery Pipe Material"
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
         Index           =   0
         Left            =   480
         TabIndex        =   24
         Top             =   960
         Width           =   4575
      End
   End
   Begin VB.ComboBox cmbPumpPipe 
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
      Left            =   5520
      TabIndex        =   8
      Text            =   "NO"
      Top             =   5040
      Width           =   855
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
      Left            =   3600
      TabIndex        =   12
      Top             =   7320
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
      Left            =   1920
      TabIndex        =   11
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   6855
      Begin VB.TextBox txtDelLev 
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
         Left            =   5160
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtSumLev 
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
         Left            =   5160
         TabIndex        =   6
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtHead 
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
         Left            =   5160
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Delivery Level at the D/S Reservoir  (RL, m)  "
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
         TabIndex        =   21
         Top             =   1320
         Width           =   4575
      End
      Begin VB.Label Label6 
         Caption         =   "Water Level in the Sump  (RL, m)      "
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
         Left            =   480
         TabIndex        =   20
         Top             =   840
         Width           =   4095
      End
      Begin VB.Label Label5 
         Caption         =   "Pump Head  (m)                                "
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
         Left            =   480
         TabIndex        =   19
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Frame frameTran 
      Caption         =   "Transmission Main"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   6855
      Begin VB.ComboBox cmbPWV 
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
         TabIndex        =   4
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtChain 
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
         Left            =   5160
         TabIndex        =   3
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtLeng 
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
         Left            =   5160
         TabIndex        =   2
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtDia 
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
         Left            =   5160
         TabIndex        =   1
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtDisch 
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
         Left            =   5160
         TabIndex        =   0
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Transmission Main Pipe Material"
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
         Left            =   240
         TabIndex        =   26
         Top             =   2400
         Width           =   4455
      End
      Begin VB.Label Label4 
         Caption         =   "Starting Chainage  (m)                       "
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
         TabIndex        =   17
         Top             =   1920
         Width           =   3975
      End
      Begin VB.Label Label3 
         Caption         =   "Length of the Main (m)                         "
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
         Top             =   1440
         Width           =   3975
      End
      Begin VB.Label Label2 
         Caption         =   "Internal Diameter (mm)                        "
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
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "Design Discharge (cum/sec)              "
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
         Top             =   480
         Width           =   3855
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Pump House Piping  Finalised ? "
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
      TabIndex        =   22
      Top             =   5040
      Width           =   3615
   End
End
Attribute VB_Name = "frmTranMainA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbPumpPipe_Click()
If cmbPumpPipe.Text = "YES" Then
  frmPHP.Visible = True
Else
  frmPHP.Visible = False
End If
End Sub

Private Sub cmbPWV_Click()
 NP = 1
 DIA = Val(txtDia.Text)
 
 If cmbPWV.Text = "STEEL" Then
  frmTranMainA.Enabled = False
  MDIForm1.mnuExec.Item(10).Enabled = False
  MDIForm1.tbrMain.Buttons(6).Enabled = False
  frmSteel.Show
 ElseIf cmbPWV.Text = "DI" Then
  frmTranMainA.Enabled = False
  MDIForm1.mnuExec.Item(10).Enabled = False
  MDIForm1.tbrMain.Buttons(6).Enabled = False
  frmDI.Show
 ElseIf cmbPWV.Text = "CI" Then
  frmTranMainA.Enabled = False
  MDIForm1.mnuExec.Item(10).Enabled = False
  MDIForm1.tbrMain.Buttons(6).Enabled = False
  frmCI.Show
 ElseIf cmbPWV.Text = "BWSC" Then
  frmTranMainA.Enabled = False
  MDIForm1.mnuExec.Item(10).Enabled = False
  MDIForm1.tbrMain.Buttons(6).Enabled = False
  frmBWSC.Show
 ElseIf cmbPWV.Text = "PSC" Then
  frmTranMainA.Enabled = False
  MDIForm1.mnuExec.Item(10).Enabled = False
  MDIForm1.tbrMain.Buttons(6).Enabled = False
  frmPSC.Show
 ElseIf cmbPWV.Text = "AC" Then
  frmTranMainA.Enabled = False
  MDIForm1.mnuExec.Item(10).Enabled = False
  MDIForm1.tbrMain.Buttons(6).Enabled = False
  frmAC.Show
 Else
  frmTranMainA.Enabled = False
  MDIForm1.mnuExec.Item(10).Enabled = False
  MDIForm1.tbrMain.Buttons(6).Enabled = False
  frmWaveVel.Show
 End If
End Sub

Private Sub cmbPWVD_Click()
 NP = 2
 DIAP = Val(txtPumpDelDia.Text)
 
 If cmbPWVD.Text = "STEEL" Then
  frmTranMainA.Enabled = False
  MDIForm1.mnuExec.Item(10).Enabled = False
  MDIForm1.tbrMain.Buttons(6).Enabled = False
  frmSteelD.Show
 ElseIf cmbPWVD.Text = "DI" Then
  frmTranMainA.Enabled = False
  MDIForm1.mnuExec.Item(10).Enabled = False
  MDIForm1.tbrMain.Buttons(6).Enabled = False
  frmDID.Show
 ElseIf cmbPWVD.Text = "CI" Then
  frmTranMainA.Enabled = False
  MDIForm1.mnuExec.Item(10).Enabled = False
  MDIForm1.tbrMain.Buttons(6).Enabled = False
  frmCID.Show
 End If
End Sub

Private Sub cmdBack_Click()
Me.Hide
frmTitle.Show
End Sub
Private Sub cmdCont_Click()
QR = Val(txtDisch.Text)
REFH = Val(txtHead.Text)
Me.Hide
frmPumpA.Show
End Sub
Private Sub Form_Load()
Left = 20
Top = 30
cmbPumpPipe.AddItem "YES"
cmbPumpPipe.AddItem "NO"

cmbPWV.AddItem "STEEL"
cmbPWV.AddItem "DI"
cmbPWV.AddItem "CI"
cmbPWV.AddItem "BWSC"
cmbPWV.AddItem "PSC"
cmbPWV.AddItem "AC"
cmbPWV.AddItem "GRP"
cmbPWV.AddItem "PVC"
cmbPWV.AddItem "HDPE"
cmbPWVD.AddItem "STEEL"
cmbPWVD.AddItem "DI"
cmbPWVD.AddItem "CI"
If Not OpenFile = "" Then
If IDELP = "YES" Then
  cmbPumpPipe.Text = "YES"
  frmPHP.Visible = True
Else
  cmbPumpPipe.Text = "NO"
  frmPHP.Visible = False
End If
End If
End Sub

