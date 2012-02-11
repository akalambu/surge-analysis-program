
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
Begin VB.Form frmGridAlign 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alignment Data"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   5700
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
      Left            =   2520
      TabIndex        =   1
      Top             =   3240
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2655
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   4683
      _Version        =   393216
      RowHeightMin    =   400
      WordWrap        =   -1  'True
      FormatString    =   ""
   End
End
Attribute VB_Name = "frmGridAlign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'global declarations guff
Const TOTALCOLUMNS = 3  'Zero based
Dim gSort As Integer
Dim gColumn As Integer
Dim xx As Integer


Private Sub Form_Activate()
If NALIGN > 0 Then
    MSFlexGrid1.Rows = NALIGN + 1
   Else
    MSFlexGrid1.Rows = 2
   End If
    
  If NALIGN > 0 Then
   For i = 1 To MSFlexGrid1.Rows - 1
     For j = 1 To MSFlexGrid1.Cols - 1
       MSFlexGrid1.TextMatrix(i, j) = AL_Data(i - 1, j - 1)
     Next
   Next
 End If
End Sub



Private Sub cmdOK_Click()

'MSFlexGrid1.Sort = 1
NALIGN = 0
   For tst = 1 To (MSFlexGrid1.Rows - 1) Step 1
     If IsNumeric(MSFlexGrid1.TextMatrix(tst, 1)) Then
        For tst1 = 1 To (MSFlexGrid1.Cols - 1) Step 1
        If Not IsNumeric(MSFlexGrid1.TextMatrix(tst, tst1)) Then
        MsgBox "You have entered discharge and missed one of the other data !!"
        Exit Sub
        End If
        AL_Data(NALIGN, tst1 - 1) = MSFlexGrid1.TextMatrix(tst, tst1)
        Next
    NALIGN = NALIGN + 1
    End If
   Next
  Me.Hide
  frmAlign.SetFocus
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

   'set array
    
    myArray = Array("Sl. No", "Chainage (m)", "Elevation (m)")
  
   'format MSFlexGrid
   'This format creates 1 fixed row for headings etc.
   ' One fixed column
   If NALIGN > 0 Then
    MSFlexGrid1.Rows = NALIGN + 1
   Else
    MSFlexGrid1.Rows = 2
   End If
    
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
   
   
   If NALIGN > 0 Then
   For i = 1 To MSFlexGrid1.Rows - 1
     For j = 1 To MSFlexGrid1.Cols - 1
       MSFlexGrid1.TextMatrix(i, j) = AL_Data(i - 1, j - 1)
     Next
   Next
   End If
     
  
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
      DoEvents
      If (MSFlexGrid1.row + 1) = MSFlexGrid1.Rows Then
       MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
      End If
      MSFlexGrid1.row = MSFlexGrid1.row + 1
      MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 0) = MSFlexGrid1.row
      


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
      If (MSFlexGrid1.row + 1) = MSFlexGrid1.Rows Then
       MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
      End If
      MSFlexGrid1.row = MSFlexGrid1.row + 1
      MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 0) = MSFlexGrid1.row
   
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
   'MSFlexGrid1.Sort = 1
End Sub

Private Sub text2_KeyPress(KeyAscii As Integer)
' Delete returns to get rid of beep.
   If KeyAscii = Asc(vbCr) Then KeyAscii = 0
End Sub




