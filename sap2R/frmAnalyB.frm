
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
Begin VB.Form frmAnalyB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Type of Analysis"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3315
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
      Left            =   3120
      TabIndex        =   3
      Top             =   2400
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
      Left            =   1320
      TabIndex        =   2
      Top             =   2400
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
      Left            =   4800
      TabIndex        =   1
      Text            =   "NO"
      Top             =   1320
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
      Left            =   4800
      TabIndex        =   0
      Text            =   "NO"
      Top             =   360
      Width           =   855
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
      TabIndex        =   5
      Top             =   1440
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
      TabIndex        =   4
      Top             =   480
      Width           =   4215
   End
End
Attribute VB_Name = "frmAnalyB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbAnProt_Click()
If Not cmbAnProt.Text = "NO" Then
 frmAnalyB.Enabled = False
 MDIForm1.mnuExec.Item(10).Enabled = False
 MDIForm1.tbrMain.Buttons(6).Enabled = False
 frmProtB.Show
 ISEL = 1
Else
 ISEL = 0
End If
End Sub
Private Sub cmbColSep_Click()
If Not cmbColSep.Text = "NO" Then
 frmColumnB.Show
 frmAnalyB.Enabled = False
 MDIForm1.mnuExec.Item(10).Enabled = False
MDIForm1.tbrMain.Buttons(6).Enabled = False
End If
End Sub
Private Sub cmdBack_Click()
Me.Hide
frmWaveTrip.Show
End Sub
Private Sub cmdCont_Click()
Me.Hide
frmOutPutB.Show
End Sub
Private Sub Form_load()
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

