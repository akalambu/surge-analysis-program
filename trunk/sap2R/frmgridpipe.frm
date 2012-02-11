

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
Begin VB.Form frmGridPipe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pipe Data "
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   10320
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
      Left            =   3960
      TabIndex        =   4
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Frame frmDPC 
      Caption         =   "Enter the Pipe Data "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   9855
      Begin VB.TextBox Text2 
         Height          =   195
         Left            =   1080
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4695
         Left            =   360
         TabIndex        =   0
         Top             =   480
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   8281
         _Version        =   393216
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
      Left            =   5400
      TabIndex        =   1
      Top             =   5760
      Width           =   1215
   End
End
Attribute VB_Name = "frmGridPipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const TOTALCOLUMNS = 8  'Zero based

Private Sub cmdOK_Click()
Dim ii As Integer
Dim jj As Integer
 
 For ii = 1 To NPMAX Step 1
     For jj = 1 To 7 Step 1
        If Not IsNumeric(MSFlexGrid1.TextMatrix(ii, jj)) Then
         MsgBox "Data is not Complete, Please Check It !!"
         Exit Sub
        End If
     Next
 Next
 
 For ii = 1 To NPMAX
   IP(ii) = MSFlexGrid1.TextMatrix(ii, 1)
   IND1(IP(ii)) = MSFlexGrid1.TextMatrix(ii, 2)
   IND2(IP(ii)) = MSFlexGrid1.TextMatrix(ii, 3)
   PDC(IP(ii)) = MSFlexGrid1.TextMatrix(ii, 4)
   PDIA(IP(ii)) = MSFlexGrid1.TextMatrix(ii, 5)
   PLEN(IP(ii)) = MSFlexGrid1.TextMatrix(ii, 6)
   CHST(IP(ii)) = MSFlexGrid1.TextMatrix(ii, 7)
 Next
 Me.Hide
 frmTranMainB.Enabled = True
 frmTranMainB.SetFocus
 frmTranMainB.cmdCont.Enabled = True
End Sub
Sub cmdOK_GotFocus()
   If Text2.Visible = True Then
       Text2.Visible = False
       MSFlexGrid1 = Text2
   End If
End Sub

Private Sub Command1_Click()
MSFlexGrid1.Rows = NP + 1
NPMAX = NP
frmTranMainB.txtNoPipe.Text = NP
 Me.Hide
 frmTranMainB.Enabled = True
 frmTranMainB.SetFocus
 frmTranMainB.cmdCont.Enabled = True
End Sub

Private Sub Form_Activate()
MSFlexGrid1.Rows = NPMAX + 1
   For i = 1 To MSFlexGrid1.Rows - 1
       MSFlexGrid1.TextMatrix(i, 0) = i
   Next
End Sub
Private Sub Form_Load()
   Dim myArray As Variant
   Dim iCount As Integer
   Left = 20
   Top = 30


   MSFlexGrid1.Rows = NPMAX + 1
   myArray = Array("Sl. No.", "Pipe No.", "U/S Node No.", "D/S Node No.", "Discharge (cum/sec)", "Diameter   (mm)", "Length       (m)", "Starting Chainage (m)")
   
   MSFlexGrid1.Cols = TOTALCOLUMNS  'Non-zero based
   MSFlexGrid1.FixedRows = 1
   MSFlexGrid1.FixedCols = 1
   MSFlexGrid1.FocusRect = flexFocusNone
   'add headings to grid
   MSFlexGrid1.Row = 0
   For iCount = 0 To TOTALCOLUMNS - 1 'Zero based
      MSFlexGrid1.ColWidth(iCount) = 1100
      MSFlexGrid1.Col = iCount
      MSFlexGrid1.Text = myArray(iCount)
   Next iCount
  
   If NPMAX > 0 And Not OpenFile = "" Then
   For ii = 1 To MSFlexGrid1.Rows - 1
     MSFlexGrid1.TextMatrix(ii, 1) = IP(ii)
     MSFlexGrid1.TextMatrix(ii, 2) = IND1(IP(ii))
     MSFlexGrid1.TextMatrix(ii, 3) = IND2(IP(ii))
     MSFlexGrid1.TextMatrix(ii, 4) = PDC(IP(ii))
     MSFlexGrid1.TextMatrix(ii, 5) = PDIA(IP(ii))
     MSFlexGrid1.TextMatrix(ii, 6) = PLEN(IP(ii))
     MSFlexGrid1.TextMatrix(ii, 7) = CHST(IP(ii))
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
