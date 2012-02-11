
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
Begin VB.Form frmProtB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Protection Devices"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4665
   ScaleWidth      =   7500
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
      Left            =   3000
      TabIndex        =   10
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Frame frmProt 
      Caption         =   "Select the Protection Devices"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   360
      TabIndex        =   11
      Top             =   240
      Width           =   6975
      Begin VB.CheckBox chkVBO 
         Caption         =   "Valve with/without Bypass Outlet"
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
         Left            =   3840
         TabIndex        =   9
         Top             =   2760
         Width           =   3015
      End
      Begin VB.CheckBox chkSRV 
         Caption         =   "Surge Relief Valve"
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
         Left            =   3840
         TabIndex        =   8
         Top             =   2160
         Width           =   2775
      End
      Begin VB.CheckBox chkSP 
         Caption         =   "Stand Pipe"
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
         Left            =   3840
         TabIndex        =   7
         Top             =   1680
         Width           =   2655
      End
      Begin VB.CheckBox chkACV 
         Caption         =   "Air Cushion Valve "
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
         Left            =   3840
         TabIndex        =   6
         Top             =   1200
         Width           =   2895
      End
      Begin VB.CheckBox chkAirV 
         Caption         =   "Air Valve"
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
         Left            =   3840
         TabIndex        =   5
         Top             =   720
         Width           =   2775
      End
      Begin VB.CheckBox chkInrv 
         Caption         =   "Intermediate Non-return Valve "
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
         Top             =   2760
         Width           =   2775
      End
      Begin VB.CheckBox chkDualPl 
         Caption         =   "Dual Plate Check Valve"
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
         TabIndex        =   3
         Top             =   2160
         Width           =   2775
      End
      Begin VB.CheckBox chkZeroV 
         Caption         =   "Zero Velocity Valve"
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
         TabIndex        =   2
         Top             =   1680
         Width           =   2655
      End
      Begin VB.CheckBox chkOneWay 
         Caption         =   "One Way Surge Tank"
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
         TabIndex        =   1
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CheckBox chkAirVes 
         Caption         =   "Air Vessel"
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
         TabIndex        =   0
         Top             =   720
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmProtB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkACV_Click()
If chkACV.Value = 1 And ISEL = 1 Then
frmProtB.Enabled = False
frmGridACVB.Show
End If
End Sub
Private Sub chkAirV_Click()
If chkAirV.Value = 1 And ISEL = 1 Then
frmProtB.Enabled = False
frmGridAVB.Show
End If
End Sub
Private Sub chkAirVes_click()
If chkAirVes.Value = 1 And ISEL = 1 Then
frmProtB.Enabled = False
frmAvesB.Show
chkSRV.Enabled = False
Else
 chkSRV.Enabled = True
End If
End Sub
Private Sub chkDualPl_Click()
If chkDualPl.Value = 1 And ISEL = 1 Then
frmProtB.Enabled = False
frmGridDPCB.Show
chkZeroV.Enabled = False
Else
 chkZeroV.Enabled = True
End If
End Sub
Private Sub chkInrv_Click()
If chkInrv.Value = 1 And ISEL = 1 Then
frmProtB.Enabled = False
frmGridInB.Show
End If
End Sub
Private Sub chkOneWay_Click()
If chkOneWay.Value = 1 And ISEL = 1 Then
frmProtB.Enabled = False
frmgridOWSTB.Show
End If
End Sub
Private Sub chkSP_Click()
If chkSP.Value = 1 And ISEL = 1 Then
frmProtB.Enabled = False
frmGridSPB.Show
End If
End Sub
Private Sub chkSRV_Click()
If chkSRV.Value = 1 And ISEL = 1 Then
frmProtB.Enabled = False
frmSRVB.Show
chkAirVes.Enabled = False
Else
 chkAirVes.Enabled = True
End If
End Sub
Private Sub chkVBO_Click()
If chkVBO.Value = 1 And ISEL = 1 Then
 frmProtB.Enabled = False
 frmVOB.Show
End If
End Sub
Private Sub chkZeroV_Click()
If chkZeroV.Value = 1 And ISEL = 1 Then
frmProtB.Enabled = False
frmGridZVB.Show
chkDualPl.Enabled = False
Else
 chkDualPl.Enabled = True
End If
End Sub
Private Sub cmdOK_Click()
ISEL = 0
If chkAirVes.Value = 1 Then
 CODEAC = "YES"
Else
 CODEAC = "NO"
End If
If chkSRV.Value = 1 Then
 CODESV = "YES"
 Else
 CODESV = "NO"
End If
If chkOneWay.Value = 0 Then
 NOSTD = 0
End If
If chkZeroV.Value = 0 Then
 NZVD = 0
End If
If chkDualPl.Value = 0 Then
 NDPCVD = 0
End If
If chkInrv.Value = 0 Then
 NNRVD = 0
End If
If chkAirV.Value = 0 Then
 NAVD = 0
End If
If chkACV.Value = 0 Then
 NACVD = 0
End If
If chkSP.Value = 0 Then
 NSSD = 0
End If
If chkSRV.Value = 0 Then
 NLSRV = 0
End If
If chkVBO.Value = 1 Then
 CODIV = "YES"
Else
 CODIV = "NO"
End If
Me.Hide
If chkAirVes.Value = 0 And chkSRV.Value = 0 And chkOneWay.Value = 0 _
And chkZeroV.Value = 0 And chkDualPl.Value = 0 And chkInrv.Value = 0 _
And chkAirV.Value = 0 And chkACV.Value = 0 And chkSP.Value = 0 _
And chkVBO.Value = 0 Then
frmAnalyB.cmbAnProt.Text = "NO"
End If

frmAnalyB.Enabled = True
frmAnalyB.SetFocus
MDIForm1.mnuExec.Item(10).Enabled = True
MDIForm1.tbrMain.Buttons(6).Enabled = True
End Sub
Private Sub Form_Load()
Left = 20
Top = 30
End Sub

