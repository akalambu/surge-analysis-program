
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
Begin VB.Form frmTitle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Project Title"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3285
   ScaleWidth      =   7605
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
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox txtCase 
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
      Left            =   2880
      TabIndex        =   1
      Text            =   "Case"
      Top             =   1440
      Width           =   4335
   End
   Begin VB.TextBox txtName 
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
      Left            =   2880
      TabIndex        =   0
      Text            =   "Name"
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   "Analysis Case       :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Project Name        :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "frmTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCont_Click()
Me.Hide
If PTYPE = "TYPEA" Then
 frmTranMainA.Show
ElseIf PTYPE = "TYPEB" Then
 MDIForm1.mnuExec.Item(10).Enabled = False
 MDIForm1.tbrMain.Buttons(6).Enabled = False
 frmTranMainB.Show
End If
End Sub
Private Sub Form_Load()
Left = 20
Top = 30
MDIForm1.mnuViewItem.Item(20).Enabled = False
MDIForm1.mnuViewItem.Item(30).Enabled = False
MDIForm1.tbrMain.Buttons.Item(9).Enabled = False
MDIForm1.tbrMain.Buttons.Item(10).Enabled = False
End Sub
