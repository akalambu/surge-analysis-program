
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
Begin VB.Form frmValveType1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nature of Valve Operations"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2940
   ScaleWidth      =   6750
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
      Left            =   2760
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1095
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   6375
      Begin VB.TextBox txtTime 
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
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtDel 
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
         TabIndex        =   1
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Time of Closure (sec)"
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
         TabIndex        =   7
         Top             =   530
         Width           =   3135
      End
      Begin VB.Label Label3 
         Caption         =   "Delay in Closure (sec)"
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
         TabIndex        =   6
         Top             =   0
         Width           =   3135
      End
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   3360
      TabIndex        =   0
      Text            =   "Closed"
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Nature of Valve Operation"
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
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "frmValveType1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOK_Click()
RESVT1 = Combo1.Text
DLYDS = Val(txtDel.Text)
TOPDS = Val(txtTime.Text)
Me.Hide
End Sub
Private Sub Combo1_Click()
If Combo1.Text = "Closed" Then
 Frame2.Visible = False
 KODEDS = -1
End If
If Combo1.Text = "Full/Partially Open" Then
 KODEDS = 1
 Frame2.Visible = False
End If
If Combo1.Text = "Valve being Closed" Then
 KODEDS = 2
 Frame2.Visible = True
End If
End Sub
Private Sub Form_Load()
Left = 20
Top = 30
Combo1.AddItem "Closed"
Combo1.AddItem "Full/Partially Open"
Combo1.AddItem "Valve being Closed"
Frame2.Visible = False
KODEDS = -1
If Not OpenFile = "" Then
Combo1.Text = RESVT1
If Combo1.Text = "Closed" Then
 KODEDS = -1
End If
If Combo1.Text = "Full/Partially Open" Then
 KODEDS = 1
End If
If Combo1.Text = "Valve being Closed" Then
 KODEDS = 2
 Frame2.Visible = True
 txtDel.Text = DLYDS
 txtTime.Text = TOPDS
End If
End If
End Sub

