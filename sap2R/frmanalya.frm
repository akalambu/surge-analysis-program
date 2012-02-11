
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
Begin VB.Form frmAnalyA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Type of Analysis"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   6555
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
      TabIndex        =   6
      Top             =   4440
      Width           =   1575
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
      Left            =   1200
      TabIndex        =   5
      Top             =   4440
      Width           =   1455
   End
   Begin VB.ComboBox cmbAnProt 
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
      Left            =   4560
      TabIndex        =   4
      Text            =   "NO"
      Top             =   3720
      Width           =   855
   End
   Begin VB.ComboBox cmbColSep 
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
      Left            =   4560
      TabIndex        =   3
      Text            =   "NO"
      Top             =   3120
      Width           =   855
   End
   Begin VB.Frame frmAnal 
      Caption         =   "Select the Type of Analysis "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   360
      TabIndex        =   7
      Top             =   360
      Width           =   4335
      Begin VB.OptionButton optAllPump 
         Caption         =   "All Pumps Failure"
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
         Top             =   1800
         Width           =   3015
      End
      Begin VB.OptionButton optSingPump 
         Caption         =   "Single Pump Failure"
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
         Width           =   3015
      End
      Begin VB.OptionButton optPowFail 
         Caption         =   "Power Failure"
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
         Top             =   600
         Width           =   3015
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Analysis with Protection      ?"
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
      TabIndex        =   9
      Top             =   3720
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Column Seperation Effect Considered ?"
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
      TabIndex        =   8
      Top             =   3120
      Width           =   4215
   End
End

Attribute VB_Exposed = False
Attribute VB_Name = "frmAnalyA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True

Private Sub cmbAnProt_Click()
If Not cmbAnProt.Text = "NO" Then
 frmAnalyA.Enabled = False
 MDIForm1.mnuExec.Item(10).Enabled = False
 MDIForm1.tbrMain.Buttons(6).Enabled = False
 frmProtA.Show
 ISEL = 1
Else
 ISEL = 0
End If
End Sub
Private Sub cmbColSep_Click()
If Not cmbColSep.Text = "NO" Then
 frmAnalyA.Enabled = False
 MDIForm1.mnuExec.Item(10).Enabled = False
MDIForm1.tbrMain.Buttons(6).Enabled = False
 frmColumnA.Show
End If
End Sub
Private Sub cmdBack_Click()
 Me.Hide
 frmPumpA.Show
End Sub
Private Sub cmdCont_Click()
 Me.Hide
 frmOutPutA.Show
End Sub
Private Sub Form_Activate()
 frmAnalyA.optSingPump.Enabled = True
 frmAnalyA.optAllPump.Enabled = True
If frmPumpA.txtPumpNo.Text = 1 Then
 frmAnalyA.optSingPump.Enabled = False
 frmAnalyA.optAllPump.Enabled = False
End If
If frmTranMainA.cmbPumpPipe.Text = "NO" Then
 frmAnalyA.optSingPump.Enabled = False
 frmAnalyA.optAllPump.Enabled = False
End If
End Sub
Private Sub Form_Load()
Left = 20
Top = 30
cmbColSep.AddItem "YES"
cmbColSep.AddItem "NO"
cmbAnProt.AddItem "YES"
cmbAnProt.AddItem "NO"
If Not OpenFile = "" Then
If CODECS = "YES" Then
 cmbColSep.Text = "YES"
Else
 cmbColSep.Text = "NO"
End If
If CODEPR = "YES" Then
 cmbAnProt.Text = "YES"
Else
 cmbAnProt.Text = "NO"
End If
End If
End Sub


