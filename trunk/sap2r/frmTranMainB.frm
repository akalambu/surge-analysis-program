
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
Begin VB.Form frmTranMainB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Details of Transmission Main "
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4.219
   ScaleMode       =   5  'Inch
   ScaleWidth      =   4.167
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   5295
      Begin VB.TextBox txtNoNodes 
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
         Left            =   3360
         TabIndex        =   2
         Text            =   "0"
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdDN 
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
         Left            =   3360
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Number of Nodes "
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
         TabIndex        =   11
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label Label3 
         Caption         =   "Node Details  "
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
         TabIndex        =   10
         Top             =   1370
         Width           =   3615
      End
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
      Left            =   2640
      TabIndex        =   5
      Top             =   5160
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
      Left            =   1080
      TabIndex        =   4
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Frame frameTran 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton cmdDP 
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
         Left            =   3360
         TabIndex        =   1
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtNoPipe 
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
         Left            =   3360
         TabIndex        =   0
         Text            =   "0"
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Pipe Details  "
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
         TabIndex        =   8
         Top             =   1400
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "Number of Pipes  "
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
         TabIndex        =   7
         Top             =   720
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmTranMainB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
NNMAX = Val(txtNoNodes.Text)
NPMAX = Val(txtNoPipe.Text)
Iflag_Err = 0
CheckNode
If Iflag_Err = 0 Then
 CheckDis
End If
If Iflag_Err = 0 Then
CollectType
Me.Hide
frmTitle.Show
MDIForm1.mnuExec.Item(10).Enabled = True
MDIForm1.tbrMain.Buttons(6).Enabled = True
End If
End Sub
Private Sub cmdCont_Click()
NNMAX = Val(txtNoNodes.Text)
NPMAX = Val(txtNoPipe.Text)
Iflag_Err = 0
CheckNode
If Iflag_Err = 0 Then
 CheckDis
End If
If Iflag_Err = 0 Then
CollectType
Me.Hide
frmHGL.Show
MDIForm1.mnuExec.Item(10).Enabled = True
MDIForm1.tbrMain.Buttons(6).Enabled = True
End If
End Sub
Private Sub cmdDP_Click()
   Iflag_But(1) = 1
   cmdDP.Caption = "Change"
   NPMAX = Val(txtNoPipe.Text)
   If NPMAX > 0 Then
   frmGridPipe.Show
   frmTranMainB.Enabled = False
   End If
End Sub
Private Sub cmdDN_Click()
Iflag_But(2) = 1
cmdDN.Caption = "Change"
NNMAX = Val(txtNoNodes.Text)
If NNMAX > 0 Then
frmGridNode.Show
frmTranMainB.Enabled = False
End If
End Sub
Private Sub Form_Activate()
NP = Val(txtNoPipe.Text)
NN = Val(txtNoNodes.Text)
If txtNoPipe.Text = 0 Then
cmdDP.Enabled = False
Else
cmdDP.Enabled = True
End If
If txtNoNodes.Text = 0 Then
cmdDN.Enabled = False
Else
cmdDN.Enabled = True
End If
End Sub
Private Sub Form_Load()
Left = 20
Top = 30
End Sub
Private Sub txtNoPipe_Change()
If Val(txtNoPipe.Text) = 0 Then
cmdDP.Enabled = False
Else
cmdDP.Enabled = True
End If
End Sub
Private Sub txtNoNodes_Change()
If Val(txtNoNodes.Text) = 0 Then
cmdDN.Enabled = False
Else
cmdDN.Enabled = True
End If
End Sub
