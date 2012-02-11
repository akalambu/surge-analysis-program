
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
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPath 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data for Paths"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
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
   ScaleHeight     =   4830
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Data for Path  : 1"
      ForeColor       =   &H00FF0000&
      Height          =   3015
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   5895
      Begin VB.Frame Frame2 
         Height          =   1335
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   5535
         Begin VB.TextBox Text2 
            Height          =   330
            Left            =   480
            TabIndex        =   2
            Top             =   600
            Visible         =   0   'False
            Width           =   735
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   495
            Left            =   240
            TabIndex        =   11
            Top             =   600
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   873
            _Version        =   393216
            Rows            =   1
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   12632256
            BorderStyle     =   0
         End
         Begin VB.Label Label2 
            Caption         =   "Enter Pipe Numbers in the Path"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   4
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "Previous"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtNPipe 
         Height          =   315
         Left            =   3360
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Number of Pipes in the Path"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.TextBox txtNPath 
      Height          =   315
      Left            =   4920
      TabIndex        =   0
      Text            =   " "
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label label1 
      Caption         =   "Number of Paths for Plotting Pressure"
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "frmPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim jj As Integer
Dim ifl As Integer
Private Sub cmdNext_Click()
getit
Save
Path_Validate
If ifl = 0 Then Exit Sub
If jj = NPATH Then Exit Sub
jj = jj + 1
txtNPipe.Text = NPPATH(jj)
Load
End Sub

Private Sub cmdOK_Click()
getit
Save
Path_Validate
If ifl = 0 Then Exit Sub
NPATH = Val(txtNPath.Text)
Me.Hide
frmOutPutB.Enabled = True
MDIForm1.mnuExec.Item(10).Enabled = True
MDIForm1.tbrMain.Buttons(6).Enabled = True
End Sub

Private Sub cmdPrev_Click()
getit
Save
Path_Validate
If ifl = 0 Then Exit Sub
If jj = 1 Then Exit Sub
jj = jj - 1
Load
End Sub



Private Sub Form_Activate()
txtNPath.Text = NPATH
If NPATH > 0 Then
txtNPipe.Text = NPPATH(jj)
Frame1.Visible = True
End If
If NPPATH(jj) > 0 Then
  FlexGLoad
Else
 Frame2.Visible = False
End If
End Sub

Private Sub Form_Load()
jj = 1

Left = 20
Top = 30
Frame1.Visible = False
Frame2.Visible = False
End Sub


Private Sub txtNPath_Change()
NPATH = Val(txtNPath.Text)
If NPATH > 0 Then
   Frame1.Visible = True
   jj = 1
   Load
 Else
   Frame1.Visible = False
 End If
 End Sub
Sub Load()
Frame1.Caption = "Data for Path :" & jj
txtNPipe.Text = NPPATH(jj)
If NPPATH(jj) > 0 Then
  FlexGLoad
Else
 Frame2.Visible = False
End If
End Sub

Sub Save()
NPPATH(jj) = Val(txtNPipe.Text)
For ii = 1 To NPPATH(jj)
IPPATH(jj, ii) = MSFlexGrid1.TextMatrix(0, ii - 1)
Next
End Sub


Private Sub txtNPipe_change()
NPPATH(jj) = Val(txtNPipe.Text)
 If NPPATH(jj) > 0 Then
   Frame2.Visible = True
   FlexGLoad
 Else
   Frame2.Visible = False
 End If
End Sub





'==================================

Private Sub FlexGLoad()
   Dim iCount As Integer
   MSFlexGrid1.Cols = NPPATH(jj)  'Non-zero based
   MSFlexGrid1.FocusRect = flexFocusNone
   For iCount = 0 To NPPATH(jj) - 1 'Zero based
      MSFlexGrid1.ColWidth(iCount) = 500
   Next iCount
   
   For ii = 1 To NPPATH(jj)
       MSFlexGrid1.TextMatrix(0, ii - 1) = IPPATH(jj, ii)
   Next
End Sub



'Private Sub HighLightGridRow(iRow As Integer)
'   MSFlexGrid1.col = 2
'   MSFlexGrid1.row = iRow
'
'   MSFlexGrid1.ColSel = TOTALCOLUMNS - 1 'Zero Based
'   MSFlexGrid1.RowSel = iRow
'
'
'End Sub

Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
   MSHFlexGridEdit MSFlexGrid1, Text2, KeyAscii
End Sub

Sub MsFlexGrid1_DblClick()
   MSHFlexGridEdit MSFlexGrid1, Text2, 32 ' Simulate a space.
End Sub

Sub MSHFlexGridEdit(MSHFlexGrid As Control, _
Edt As Control, KeyAscii As Integer)
' Use the character that was typed.
   Select Case KeyAscii
   ' A space means edit the current text.
   Case 0 To 32
      Edt = MSHFlexGrid1
      Edt.SelStart = 1000
   ' Anything else means replace the current text.
   Case Else
      Edt = Chr(KeyAscii)
      Edt.SelStart = 1
      End Select
   ' Show Edt at the right place.
   Edt.Move MSHFlexGrid.Left + MSHFlexGrid.CellLeft, _
      MSHFlexGrid.Top + MSHFlexGrid.CellTop, _
      MSHFlexGrid.CellWidth - 8, _
      MSHFlexGrid.CellHeight - 6
      Edt.Visible = True
      ' And make it work.
   Edt.SetFocus
   End Sub

Sub text2_KeyDown(KeyCode As Integer, _
Shift As Integer)
   EditKeyCode MSFlexGrid1, Text2, KeyCode, Shift
End Sub

Sub EditKeyCode(MSHFlexGrid As Control, Edt As _
Control, KeyCode As Integer, Shift As Integer)
   ' Standard edit control processing.
   Select Case KeyCode

   Case 27   ' ESC: hide, return focus to MSHFlexGrid.
      Edt.Visible = False
      MSHFlexGrid1.SetFocus

   Case 13   ' ENTER return focus to MSHFlexGrid.
      MSFlexGrid1.SetFocus

   Case 37      ' Left
      MSFlexGrid1.SetFocus
      DoEvents
      If MSFlexGrid1.Col > MSFlexGrid1.FixedCols Then
         MSFlexGrid1.Col = MSFlexGrid1.Col - 1
      End If

   Case 39      ' Right
      MSFlexGrid1.SetFocus
      DoEvents
      If MSFlexGrid1.Col < MSFlexGrid1.Cols - 1 Then
         MSFlexGrid1.Col = MSFlexGrid1.Col + 1
      End If
   End Select
End Sub
Sub MSFlexGrid1_GotFocus()
   If Text2.Visible = False Then
   Exit Sub
   End If
   MSFlexGrid1 = Text2
   Text2.Visible = False
   End Sub

Sub MSFlexGrid1_LeaveCell()
   If Text2.Visible = False Then Exit Sub
   MSFlexGrid1 = Text2
   Text2.Visible = False
End Sub

Private Sub text2_KeyPress(KeyAscii As Integer)
' Delete returns to get rid of beep.
   If KeyAscii = Asc(vbCr) Then KeyAscii = 0
End Sub


Sub getit()
   If Text2.Visible = False Then Exit Sub
   MSFlexGrid1 = Text2
   Text2.Visible = False
End Sub





' Path Validation

  Sub Path_Validate()
  ifl = 1
  For ii = 1 To NPPATH(jj) - 1
   ifl = 0
   NP = IND2(IPPATH(jj, ii))
   For kk = 1 To NDSP(NP)
    If IPPATH(jj, ii + 1) = IDSP(NP, kk) Then
      ifl = 1
    End If
   Next
   If ifl = 0 Then
    MsgBox "The pipe " & IPPATH(jj, ii + 1) & " is not the next downstream pipe of " & NP
   Exit Sub
   End If
  Next
  End Sub
