
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
Begin VB.Form frmGridCDS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Condenser Details"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
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
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Frame frmINRVs 
      Caption         =   "Enter the Condenser Details"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   330
      TabIndex        =   2
      Top             =   240
      Width           =   5055
      Begin VB.TextBox Text2 
         Height          =   195
         Left            =   600
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2175
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   4770
         _ExtentX        =   8414
         _ExtentY        =   3836
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
      Left            =   2880
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
   End
End
Attribute VB_Name = "frmGridCDS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const TOTALCOLUMNS = 4  'Zero based
Private Sub cmdOK_Click()
Dim ii As Integer
Dim jj As Integer

For ii = 1 To NCDS
  For jj = 2 To 3
    If Not IsNumeric(MSFlexGrid1.TextMatrix(ii, jj)) Then
       MsgBox "Data is not Complete, Please Check It !!"
      Exit Sub
    End If
  Next
  Next

 For ii = 1 To NCDS
   HGL1CD(ii) = MSFlexGrid1.TextMatrix(ii, 2)
   HGL2CD(ii) = MSFlexGrid1.TextMatrix(ii, 3)
   QOCDS(ii) = PDC(IUSP(MSFlexGrid1.TextMatrix(ii, 1), 1))
   NPUSCD(ii) = IUSP(MSFlexGrid1.TextMatrix(ii, 1), 1)
   ICDS(MSFlexGrid1.TextMatrix(ii, 1)) = ii
 Next
 
Me.Hide
frmHGL.Enabled = True
MDIForm1.mnuExec.Item(10).Enabled = True
MDIForm1.tbrMain.Buttons(6).Enabled = True
End Sub
Sub cmdOK_GotFocus()
   If Text2.Visible = True Then
       MSFlexGrid1 = Text2
       Text2.Visible = False
      End If
End Sub

Private Sub Command1_Click()
Me.Hide
frmHGL.Enabled = True
MDIForm1.mnuExec.Item(10).Enabled = True
MDIForm1.tbrMain.Buttons(6).Enabled = True
End Sub

Private Sub Form_Activate()
MSFlexGrid1.Rows = NCDS + 1
ii = 0
   For i = 1 To NNMAX
        If (NTYPE(i) = 6) Then
         ii = ii + 1
         MSFlexGrid1.TextMatrix(ii, 0) = ii
         MSFlexGrid1.TextMatrix(ii, 1) = i
        End If
   Next
End Sub

Private Sub Form_Load()
   Dim iCount As Integer
   Dim myArray As Variant
   Left = 20
   Top = 30
   
  'set array
   myArray = Array("Sl. No.", "Node No.", "HGL U/S (RL, m)", "HGL D/S (RL, m)")

   MSFlexGrid1.Rows = NCDS + 1
   MSFlexGrid1.Cols = TOTALCOLUMNS  'Non-zero based
   MSFlexGrid1.FixedRows = 1
   MSFlexGrid1.FixedCols = 2
   MSFlexGrid1.FocusRect = flexFocusNone
   'add headings to grid
   MSFlexGrid1.Row = 0
   For iCount = 0 To TOTALCOLUMNS - 1 'Zero based
      MSFlexGrid1.ColWidth(iCount) = 1100
      MSFlexGrid1.Col = iCount
      MSFlexGrid1.Text = myArray(iCount)
   Next iCount
   ii = 0
   For i = 1 To NNMAX
        If (NTYPE(i) = 6) Then
         ii = ii + 1
         MSFlexGrid1.TextMatrix(ii, 0) = ii
         MSFlexGrid1.TextMatrix(ii, 1) = i
        End If
   Next
   
  If Not OpenFile = "" Then
   For ii = 1 To MSFlexGrid1.Rows - 1
       MSFlexGrid1.TextMatrix(ii, 2) = HGL1CD(ii)
       MSFlexGrid1.TextMatrix(ii, 3) = HGL2CD(ii)
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

