
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
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAlignB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alignment Data"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3810
   ScaleWidth      =   7005
   Begin VB.CommandButton Command2 
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
      Left            =   2880
      TabIndex        =   2
      Top             =   3240
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog cmDial 
      Left            =   5760
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frmYES 
      Height          =   2415
      Left            =   480
      TabIndex        =   3
      Top             =   480
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
         Left            =   3120
         TabIndex        =   1
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
         Left            =   1320
         TabIndex        =   0
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Alignment Data (Chainage and Elevation)"
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
         Left            =   960
         TabIndex        =   4
         Top             =   480
         Width           =   4815
      End
   End
End
Attribute VB_Name = "frmAlignB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim OpFile As String
Dim PN As Integer
Private Sub cmdEnter_Click()
 frmGridAlignB.Show
End Sub
Private Sub cmdgetIt_Click()
 Dim xxx As String
 cmDial.Filter = "aln (*.aln)|*.aln"
 cmDial.FileName = ""
 cmDial.ShowOpen
 OpFile = cmDial.FileName
 xxx = Dir(OpFile)
 If xxx = "" Then
  MsgBox "File Not Found"
 ElseIf Not OpFile = "" Then
  ReadItB
 End If
End Sub
Private Sub ReadItB()
Dim check As Variant
For i = 1 To NPMAX
 NALIGN(i) = 0
Next
Open OpFile For Input As #1
Do Until EOF(1)
  Input #1, check
  If Not IsNumeric(check) Then
   Close (1)
   Exit Sub
  End If
  PN = check
  NALIGN(PN) = NALIGN(PN) + 1
  Input #1, CHAIN(PN, NALIGN(PN)), GL(PN, NALIGN(PN))
Loop
Close (1)
End Sub
Private Sub Command2_Click()
Me.Hide
MDIForm1.mnuExec.Item(10).Enabled = True
MDIForm1.tbrMain.Buttons(6).Enabled = True
End Sub
Private Sub Form_Load()
Left = 20
Top = 30
End Sub
