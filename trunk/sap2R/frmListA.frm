

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
Begin VB.Form frmListA 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pump House NRV"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   2040
      Width           =   735
   End
   Begin VB.ListBox lstValve 
      Height          =   1500
      ItemData        =   "frmListA.frx":0000
      Left            =   240
      List            =   "frmListA.frx":0016
      TabIndex        =   0
      Top             =   240
      Width           =   9495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Select one from the list and click OK"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   3975
   End
End
Attribute VB_Name = "frmListA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdOK_Click()
Dim nLineclicked As Integer


With lstValve
    nLineclicked = .ListIndex
    If nLineclicked = -1 Then
     Exit Sub
    End If
    
Select Case nLineclicked
  Case 0
     frmswt.Show
     KODPHVA = 1
  Case 1
     frmsa.Show
     KODPHVA = 2
  Case 2
     frmsb.Show
     KODPHVA = 4
  Case 3
      frmta.Show
      KODPHVA = 3
  Case 4
      frmtsb.Show
      KODPHVA = 5
  Case 5
      KODPHVA = 6
End Select
    
End With

frmListA.Hide
End Sub


Private Sub Form_Load()
Left = 20
Top = 30
End Sub
