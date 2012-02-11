
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
Begin VB.Form frmWaveTrip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wave Velocity and Trip Code"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4275
   ScaleWidth      =   5730
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   4935
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
         Left            =   3960
         TabIndex        =   3
         Text            =   "NO"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Valve Closing at Delivery Reservoir to be Considered ?"
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
         Left            =   0
         TabIndex        =   10
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.CommandButton cmdTripB 
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
      Left            =   4080
      TabIndex        =   2
      Top             =   1320
      Width           =   975
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
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
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
      Left            =   2760
      TabIndex        =   5
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton cmdTripP 
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
      Left            =   4080
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdWave 
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
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Trip Code for Each Booster"
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
      TabIndex        =   8
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Trip Code for Each Pump"
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
      TabIndex        =   7
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Type of Pipe"
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
      TabIndex        =   6
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmWaveTrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdBack_Click()
Unload Me
frmHGL.Show
End Sub
Private Sub cmdTripP_Click()
Iflag_But(13) = 1
cmdTripP.Caption = "Change"
frmGridPTrip.Show
frmWaveTrip.Enabled = False
MDIForm1.mnuExec.Item(10).Enabled = False
MDIForm1.tbrMain.Buttons(6).Enabled = False
End Sub
Private Sub cmdTripB_Click()
Iflag_But(14) = 1
cmdTripB.Caption = "Change"
frmGridBTrip.Show
frmWaveTrip.Enabled = False
MDIForm1.mnuExec.Item(10).Enabled = False
MDIForm1.tbrMain.Buttons(6).Enabled = False
End Sub
Private Sub cmdWave_Click()
Iflag_But(12) = 1
cmdWave.Caption = "Change"
frmGridWave.Show
frmWaveTrip.Enabled = False
End Sub
Private Sub Combo1_Click()
DELV = Combo1.Text
If Combo1.Text = "YES" Then
  frmValveType2.Show
  frmWaveTrip.Enabled = False
  MDIForm1.mnuExec.Item(10).Enabled = False
  MDIForm1.tbrMain.Buttons(6).Enabled = False
ElseIf Combo1.Text = "NO" Then
  KODEDS = 0
  DLYDS = 0
  TOPDS = 0
End If
End Sub
Private Sub Command1_Click()
DELV = Combo1.Text
Unload Me
frmAnalyB.Show
End Sub

Private Sub Form_Load()
Left = 20
Top = 30
Combo1.AddItem "YES"
Combo1.AddItem "NO"
Combo1.Text = DELV
If NPMAX = 0 Then
cmdWave.Enabled = False
End If
If NPMP < 1 Then
cmdTripP.Enabled = False
End If
If NBST < 1 Then
cmdTripB.Enabled = False
End If
If NRES = 1 Then
Frame1.Visible = True
Else
Frame1.Visible = False
End If
 If Iflag_But(12) = 1 Then
 cmdWave.Caption = "Change"
 End If
 If Iflag_But(13) = 1 Then
 cmdTripP.Caption = "Change"
 End If
 If Iflag_But(14) = 1 Then
 cmdTripB.Caption = "Change"
 End If
End Sub
