
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
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmGridWave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Wave Velocity "
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmDPC 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   3975
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
         Left            =   2400
         TabIndex        =   3
         Top             =   960
         Visible         =   0   'False
         Width           =   1200
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3495
         Left            =   360
         TabIndex        =   0
         Top             =   360
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   6165
         _Version        =   393216
         Cols            =   3
         FixedCols       =   2
         RowHeightMin    =   410
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
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
      Left            =   1800
      TabIndex        =   1
      Top             =   4440
      Width           =   855
   End
End
Attribute VB_Name = "frmGridWave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const TOTALCOLUMNS = 3  'Zero based
Sub cmdOK_Click()
Dim ii As Integer
Dim jj As Integer
 
 For ii = 1 To NPMAX Step 1
        If IsEmpty(MSFlexGrid1.TextMatrix(ii, 2)) Then
         MsgBox "Data is not Complete, Please Check It !!"
         Exit Sub
        End If
 Next
 For ii = 1 To NPMAX
         PIPEMAT(IP(ii)) = MSFlexGrid1.TextMatrix(ii, 2)
 Next
 Me.Hide
 frmWaveTrip.Enabled = True
End Sub
Sub cmdOK_GotFocus()
   If Combo1.Visible = True Then
       Combo1.Visible = False
       MSFlexGrid1 = Combo1
   End If
End Sub
Private Sub Combo1_Click()
 NP = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
 If Combo1.Text = "STEEL" Then
  frmGridWave.Enabled = False
  MDIForm1.mnuExec.Item(10).Enabled = False
  MDIForm1.tbrMain.Buttons(6).Enabled = False
  frmSteel.Show
 ElseIf Combo1.Text = "DI" Then
  frmGridWave.Enabled = False
  MDIForm1.mnuExec.Item(10).Enabled = False
  MDIForm1.tbrMain.Buttons(6).Enabled = False
  frmDI.Show
 ElseIf Combo1.Text = "CI" Then
  frmGridWave.Enabled = False
  MDIForm1.mnuExec.Item(10).Enabled = False
  MDIForm1.tbrMain.Buttons(6).Enabled = False
  frmCI.Show
 ElseIf Combo1.Text = "BWSC" Then
  frmGridWave.Enabled = False
  MDIForm1.mnuExec.Item(10).Enabled = False
  MDIForm1.tbrMain.Buttons(6).Enabled = False
  frmBWSC.Show
 ElseIf Combo1.Text = "PSC" Then
  frmGridWave.Enabled = False
  MDIForm1.mnuExec.Item(10).Enabled = False
  MDIForm1.tbrMain.Buttons(6).Enabled = False
  frmPSC.Show
 ElseIf Combo1.Text = "AC" Then
  frmGridWave.Enabled = False
  MDIForm1.mnuExec.Item(10).Enabled = False
  MDIForm1.tbrMain.Buttons(6).Enabled = False
  frmAC.Show
 Else
  frmGridWave.Enabled = False
  MDIForm1.mnuExec.Item(10).Enabled = False
  MDIForm1.tbrMain.Buttons(6).Enabled = False
  frmWaveVel.Show
 End If
 MSFlexGrid1 = Combo1
 Combo1.Visible = False
End Sub
Private Sub Form_Activate()
MSFlexGrid1.Rows = NPMAX + 1
   For i = 1 To MSFlexGrid1.Rows - 1
       MSFlexGrid1.TextMatrix(i, 0) = i
       MSFlexGrid1.TextMatrix(i, 1) = IP(i)
   Next
End Sub
Private Sub Form_Load()
   Dim myArray As Variant
   Dim iCount As Integer

   Combo1.AddItem "STEEL"
   Combo1.AddItem "DI"
   Combo1.AddItem "CI"
   Combo1.AddItem "BWSC"
   Combo1.AddItem "PSC"
   Combo1.AddItem "AC"
   Combo1.AddItem "GRP"
   Combo1.AddItem "PVC"
   Combo1.AddItem "HDPE"
   
   Left = 20
   Top = 30
   If NPMAX > 0 Then
   MSFlexGrid1.Rows = NPMAX + 1
   Else
   MSFlexGrid1.FixedRows = 1
   End If
   myArray = Array("Sl. No.", "Pipe No.", "Material Type")
   MSFlexGrid1.Cols = TOTALCOLUMNS  'Non-zero based
   MSFlexGrid1.FixedCols = 2
   MSFlexGrid1.FocusRect = flexFocusNone
   
   
   'MSFlexGrid1.SelectionMode = flexSelectionByRow
   'add headings to grid
   MSFlexGrid1.Row = 0
   
   For iCount = 0 To 1 'Zero based
      MSFlexGrid1.ColWidth(iCount) = 800
      MSFlexGrid1.Col = iCount
      MSFlexGrid1.Text = myArray(iCount)
   Next iCount
  
   MSFlexGrid1.ColWidth(2) = 1200
   MSFlexGrid1.Col = 2
   MSFlexGrid1.Text = myArray(2)
  
  If NPMAX > 0 And Not OpenFile = "" Then
   For ii = 1 To MSFlexGrid1.Rows - 1
     MSFlexGrid1.TextMatrix(ii, 2) = PIPEMAT(IP(ii))
   Next
  End If
  'highlight 1st row
   HighLightGridRow (1)
End Sub
Private Sub HighLightGridRow(iRow As Integer)
   MSFlexGrid1.Col = 2
   MSFlexGrid1.Row = iRow
   MSFlexGrid1.ColSel = TOTALCOLUMNS - 1 'Zero Based
   MSFlexGrid1.RowSel = iRow
End Sub

Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
   MSHFlexGridEdit MSFlexGrid1, Combo1, KeyAscii
End Sub
Sub MsFlexGrid1_DblClick()
   MSHFlexGridEdit MSFlexGrid1, Combo1, 32 ' Simulate a space.
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
      MSHFlexGrid.Top + MSHFlexGrid.CellTop
   
   ' And make it work.
   Edt.Visible = True
   Edt.SetFocus
   End Sub
Sub EditKeyCode(MSHFlexGrid As Control, Edt As _
Control, KeyCode As Integer, Shift As Integer)

   ' Standard edit control processing.
   Select Case KeyCode

   Case 27   ' ESC: hide, return focus to MSHFlexGrid.
      Edt.Visible = False
      MSFlexGrid1.SetFocus

   Case 13   ' ENTER return focus to MSHFlexGrid.
      MSFlexGrid1.SetFocus

   Case 37      ' Left
      MSFlexGrid1.SetFocus
      DoEvents
      If MSFlexGrid1.Col > MSFlexGrid1.FixedCols Then
         MSFlexGrid1.Col = MSFlexGrid1.Col - 1
      End If

   Case 38      ' Up.
      MSFlexGrid1.SetFocus
      DoEvents
      If MSFlexGrid1.Row > MSFlexGrid1.FixedRows Then
         MSFlexGrid1.Row = MSFlexGrid1.Row - 1
      End If
      
   Case 39      ' Right
      MSFlexGrid1.SetFocus
      DoEvents
      If MSFlexGrid1.Col < MSFlexGrid1.Cols - 1 Then
         MSFlexGrid1.Col = MSFlexGrid1.Col + 1
      End If
   
   Case 40      ' Down.
      MSFlexGrid1.SetFocus
      DoEvents
      If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
         MSFlexGrid1.Row = MSFlexGrid1.Row + 1
      End If
   End Select
End Sub

Sub MSFlexGrid1_LeaveCell()
   If Combo1.Visible = False Then Exit Sub
   MSFlexGrid1 = Combo1
   Combo1.Visible = False
End Sub

