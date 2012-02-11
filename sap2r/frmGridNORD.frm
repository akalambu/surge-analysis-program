
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
Begin VB.Form frmGridNORD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Information of Downstream pipes"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmINRVs 
      Caption         =   "Enter the D/S Pipe Numbers"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   3375
      Begin VB.TextBox Text2 
         Height          =   195
         Left            =   600
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1815
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   3201
         _Version        =   393216
         RowHeightMin    =   400
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
      Left            =   1560
      TabIndex        =   1
      Top             =   3120
      Width           =   855
   End
End
Attribute VB_Name = "frmGridNORD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'global declarations guff
Const TOTALCOLUMNS = 2  'Zero based
Dim gSort As Integer
Dim gColumn As Integer
Dim xx As Integer
Private Sub cmdOK_Click()
Dim ii As Integer
Dim jj As Integer
Dim flag As Integer

flag = 0
         For jj = 1 To NDSP(NoNode)
         If Not IsNumeric(MSFlexGrid1.TextMatrix(jj, 1)) Then
         MsgBox "Data is not Complete, Please Check It !!"
         flag = 1
         Exit For
         End If
         IDSP(NoNode, jj) = MSFlexGrid1.TextMatrix(jj, 1)
        Next
     
If flag = 0 Then
 Unload Me
End If
End Sub
Sub cmdOK_GotFocus()
   If Text2.Visible = True Then
       MSFlexGrid1 = Text2
      End If
    End Sub

Private Sub Form_Load()
   Dim iCount As Integer
   Dim myArray As Variant
'set array
    
    myArray = Array("Sl. No.", "D/S Pipes")

    Left = 20
    Top = 30
   
   'format MSFlexGrid
   'This format creates 1 fixed row for headings etc.
   ' One fixed column
   
   
   MSFlexGrid1.Rows = NDSP(NoNode) + 1
   
   
   MSFlexGrid1.Cols = TOTALCOLUMNS  'Non-zero based
   MSFlexGrid1.FixedRows = 1
   MSFlexGrid1.FixedCols = 1
   MSFlexGrid1.FocusRect = flexFocusNone
   'MSFlexGrid1.SelectionMode = flexSelectionByRow
  
  
   'add headings to grid
   MSFlexGrid1.row = 0
   For iCount = 0 To TOTALCOLUMNS - 1 'Zero based
      MSFlexGrid1.ColWidth(iCount) = 1100
      MSFlexGrid1.col = iCount
      MSFlexGrid1.Text = myArray(iCount)
   Next iCount
  
   For i = 1 To MSFlexGrid1.Rows - 1
       MSFlexGrid1.TextMatrix(i, 0) = i
   Next
   
   
   For j = 1 To MSFlexGrid1.Rows - 1
       If Not IDSP(NoNode, j) = 0 Then
       MSFlexGrid1.TextMatrix(j, 1) = IDSP(NoNode, j)
       End If
   Next
   'remove 1st blank row from MSFlexGrid
   'MSFlexGrid1.RemoveItem (1)
  
   'highlight 1st row
   HighLightGridRow (1)
  
  
End Sub
Private Sub HighLightGridRow(iRow As Integer)
  
   MSFlexGrid1.col = 1
   MSFlexGrid1.row = iRow
  
   MSFlexGrid1.ColSel = TOTALCOLUMNS - 1 'Zero Based
   MSFlexGrid1.RowSel = iRow
  

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
      If MSFlexGrid1.col > MSFlexGrid1.FixedCols Then
         MSFlexGrid1.col = MSFlexGrid1.col - 1
      End If

   Case 38      ' Up.
      MSFlexGrid1.SetFocus
      DoEvents
      If MSFlexGrid1.row > MSFlexGrid1.FixedRows Then
         MSFlexGrid1.row = MSFlexGrid1.row - 1
      End If
      
   Case 39      ' Right
      MSFlexGrid1.SetFocus
      DoEvents
      If MSFlexGrid1.col < MSFlexGrid1.Cols - 1 Then
         MSFlexGrid1.col = MSFlexGrid1.col + 1
      End If
   

   Case 40      ' Down.
      MSFlexGrid1.SetFocus
      DoEvents
      If MSFlexGrid1.row < MSFlexGrid1.Rows - 1 Then
         MSFlexGrid1.row = MSFlexGrid1.row + 1
      End If
   End Select
End Sub
Sub MSFlexGrid1_GotFocus()
   
   If Text2.Visible = False Then
   Exit Sub
   End If
   MSFlexGrid1 = Text2
   Text2.Visible = False
   'MSFlexGrid1.Sort = 1
   
   End Sub

Sub MSFlexGrid1_LeaveCell()
   If Text2.Visible = False Then Exit Sub
   MSFlexGrid1 = Text2
   Text2.Visible = False
   'MSFlexGrid1.Sort = 1
End Sub

Private Sub text2_KeyPress(KeyAscii As Integer)
' Delete returns to get rid of beep.
   If KeyAscii = Asc(vbCr) Then KeyAscii = 0
End Sub




