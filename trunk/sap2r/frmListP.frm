
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
Begin VB.Form frmListP 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pump House NRV"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
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
   ScaleHeight     =   2790
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   735
   End
   Begin VB.ListBox lstValve 
      Height          =   1500
      ItemData        =   "frmListP.frx":0000
      Left            =   240
      List            =   "frmListP.frx":0016
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Select one from the list and click OK"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   3975
   End
End
Attribute VB_Name = "frmListP"
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
     If Iflag_PB = 1 Then
     KODPHV(NP) = 1
     ElseIf Iflag_PB = 2 Then
     KODBSV(NP) = 1
     End If
  
  Case 1
     frmsa.Show
     If Iflag_PB = 1 Then
     KODPHV(NP) = 2
     ElseIf Iflag_PB = 2 Then
     KODBSV(NP) = 2
     End If
  
  Case 2
     frmta.Show
     If Iflag_PB = 1 Then
     KODPHV(NP) = 3
     ElseIf Iflag_PB = 2 Then
     KODBSV(NP) = 3
     End If
  
  Case 3
      frmsb.Show
      If Iflag_PB = 1 Then
     KODPHV(NP) = 4
     ElseIf Iflag_PB = 2 Then
     KODBSV(NP) = 4
     End If
  
  Case 4
      frmtsb.Show
      If Iflag_PB = 1 Then
     KODPHV(NP) = 5
     ElseIf Iflag_PB = 2 Then
     KODBSV(NP) = 5
     End If
  
  Case 5
      If Iflag_PB = 1 Then
     KODPHV(NP) = 6
     If PTYPE = "TYPEA" Then
       frmPumpA.Enabled = True
       MDIForm1.mnuExec.Item(10).Enabled = True
       MDIForm1.tbrMain.Buttons(6).Enabled = True
     Else
      frmGridPUMP.Enabled = True
     End If
     ElseIf Iflag_PB = 2 Then
     KODBSV(NP) = 6
     frmGridBOOST.Enabled = True
     End If
End Select
End With
Unload frmListP
If PTYPE = "TYPEB" Then
If Iflag_PB = 1 Then
frmGridPUMP.MSFlexGrid1.TextMatrix(NP, 7) = KODPHV(NP)
ElseIf Iflag_PB = 2 Then
frmGridBOOST.MSFlexGrid1.TextMatrix(NP, 7) = KODBSV(NP)
End If
End If
End Sub
Private Sub Form_Load()
Left = 20
Top = 30

Select Case KODPHV(NP)
   Case 1
    lstValve.ListIndex = 0
   Case 2
    lstValve.ListIndex = 1
   Case 3
    lstValve.ListIndex = 2
   Case 4
    lstValve.ListIndex = 3
   Case 5
    lstValve.ListIndex = 4
   Case 6
    lstValve.ListIndex = 5
End Select
End Sub
