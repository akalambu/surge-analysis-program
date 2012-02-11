
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
Begin VB.Form frmGridAlignA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alignment Data"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
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
      Top             =   3360
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2655
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   4683
      _Version        =   393216
      RowHeightMin    =   400
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuAddDelete 
      Caption         =   "Add or Delete Row"
      Visible         =   0   'False
      Begin VB.Menu mnuAddRow 
         Caption         =   "Add Row"
      End
      Begin VB.Menu mnuDeleteRow 
         Caption         =   "Delete Row"
      End
   End
End
Attribute VB_Name = "frmGridAlignA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const TOTALCOLUMNS = 3  'Zero based
Private Sub cmdOK_Click()

   NALIGN(1) = 0
  
        
           
   For ii = 1 To (MSFlexGrid1.Rows - 1) Step 1
    If IsNumeric(MSFlexGrid1.TextMatrix(ii, 1)) And IsNumeric(MSFlexGrid1.TextMatrix(ii, 2)) Then
     NALIGN(1) = NALIGN(1) + 1
     CHAIN(1, NALIGN(1)) = MSFlexGrid1.TextMatrix(ii, 1)
     GL(1, NALIGN(1)) = MSFlexGrid1.TextMatrix(ii, 2)
    End If
   Next
   Unload Me
  'frmAlignB.SetFocus
End Sub

Sub cmdOK_GotFocus()
  If Text2.Visible = True Then
  MSFlexGrid1 = Text2
  Text2.Visible = False
  End If
End Sub


Private Sub Form_Load()
   Dim myArray As Variant
   Dim iCount As Integer
    Left = 20
    Top = 30
    myArray = Array("Sl. No.", "Chainage (m)", "Elevation (RL, m)")
   If NALIGN(1) > 0 Then
    MSFlexGrid1.Rows = NALIGN(1) + 1
   Else
    MSFlexGrid1.Rows = 2
   End If
   MSFlexGrid1.Cols = TOTALCOLUMNS  'Non-zero based
   MSFlexGrid1.FixedRows = 1
   MSFlexGrid1.FixedCols = 1
   MSFlexGrid1.FocusRect = flexFocusNone
   'MSFlexGrid1.SelectionMode = flexSelectionByRow
   'add headings to grid
   MSFlexGrid1.Row = 0
   For iCount = 0 To TOTALCOLUMNS - 1 'Zero based
      MSFlexGrid1.ColWidth(iCount) = 1100
      MSFlexGrid1.Col = iCount
      MSFlexGrid1.Text = myArray(iCount)
   Next iCount
   For i = 1 To MSFlexGrid1.Rows - 1
       MSFlexGrid1.TextMatrix(i, 0) = i
   Next
   If NALIGN(1) > 0 Then
     For j = 1 To NALIGN(1)
       MSFlexGrid1.TextMatrix(j, 1) = CHAIN(1, j)
       MSFlexGrid1.TextMatrix(j, 2) = GL(1, j)
     Next
   End If
   'highlight 1st row
   HighLightGridRow (1)
End Sub
Private Sub HighLightGridRow(iRow As Integer)
   MSFlexGrid1.Col = 1
   MSFlexGrid1.Row = iRow
   MSFlexGrid1.ColSel = TOTALCOLUMNS - 1 'Zero Based
   MSFlexGrid1.RowSel = iRow
End Sub
Private Sub mnuAddRow_Click()
AddRow
End Sub
Private Sub mnuDeleteRow_Click()
DeleteRow
End Sub
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
      MSHFlexGrid.CellHeight - 8
      Edt.Visible = True
      ' And make it work.
   Edt.SetFocus
   End Sub
Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
 PopupMenu mnuAddDelete
End If
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
      DoEvents
      If (MSFlexGrid1.Row + 1) = MSFlexGrid1.Rows Then
       MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
      End If
      MSFlexGrid1.Row = MSFlexGrid1.Row + 1
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0) = MSFlexGrid1.Row
      
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
      If (MSFlexGrid1.Row + 1) = MSFlexGrid1.Rows Then
       MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
      End If
      MSFlexGrid1.Row = MSFlexGrid1.Row + 1
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0) = MSFlexGrid1.Row
   
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
Sub AddRow()
 MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
 For ii = MSFlexGrid1.Rows - 1 To MSFlexGrid1.Row + 1 Step -1
   For jj = 1 To MSFlexGrid1.Cols - 1
   MSFlexGrid1.TextMatrix(ii, jj) = MSFlexGrid1.TextMatrix(ii - 1, jj)
   Next
 Next
 For jj = 1 To MSFlexGrid1.Cols - 1
   MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, jj) = " "
 Next
 For i = 1 To MSFlexGrid1.Rows - 1
       MSFlexGrid1.TextMatrix(i, 0) = i
 Next
End Sub

Sub DeleteRow()
 For ii = MSFlexGrid1.Row + 1 To MSFlexGrid1.Rows - 1 Step 1
   For jj = 1 To MSFlexGrid1.Cols - 1
   MSFlexGrid1.TextMatrix(ii - 1, jj) = MSFlexGrid1.TextMatrix(ii, jj)
   Next
 Next
 If MSFlexGrid1.Rows > 2 Then
 MSFlexGrid1.Rows = MSFlexGrid1.Rows - 1
 Else
 For jj = 1 To MSFlexGrid1.Cols - 1
  MSFlexGrid1.TextMatrix(1, jj) = ""
 Next
 End If
 For i = 1 To MSFlexGrid1.Rows - 1
    MSFlexGrid1.TextMatrix(i, 0) = i
 Next
End Sub

