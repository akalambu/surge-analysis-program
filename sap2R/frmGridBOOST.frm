
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
Begin VB.Form frmGridBOOST 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Booster Details"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   9750
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
      Left            =   3840
      TabIndex        =   6
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Frame frmINRVs 
      Caption         =   "Enter the Booster Details"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   9015
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
         ItemData        =   "frmGridBOOST.frx":0000
         Left            =   7440
         List            =   "frmGridBOOST.frx":0002
         TabIndex        =   4
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
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
         Top             =   480
         Width           =   8730
         _ExtentX        =   15399
         _ExtentY        =   3201
         _Version        =   393216
         Cols            =   9
         FixedCols       =   2
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
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "If Discharge, Head or Speed in a  partcular row is changed, select the ""Machinery Finalised ?"" column again."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   1560
         TabIndex        =   5
         Top             =   2400
         Width           =   5655
         WordWrap        =   -1  'True
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
      Left            =   5160
      TabIndex        =   1
      Top             =   3360
      Width           =   1095
   End
End
Attribute VB_Name = "frmGridBOOST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const TOTALCOLUMNS = 9  'Zero based
Private Sub cmdOK_Click()
Dim ii As Integer
Dim jj As Integer

  For ii = 1 To NBST
  For jj = 2 To 7
    If Not IsNumeric(MSFlexGrid1.TextMatrix(ii, jj)) Then
       MsgBox "Data is not Complete, Please Check It !!"
      Exit Sub
    End If
  Next
  Next

 For ii = 1 To NBST
   NBOOST(ii) = MSFlexGrid1.TextMatrix(ii, 2)
   BSTDC(ii) = MSFlexGrid1.TextMatrix(ii, 3)
   BSTH(ii) = MSFlexGrid1.TextMatrix(ii, 4)
   BSTSP(ii) = MSFlexGrid1.TextMatrix(ii, 5)
   HGLSUC(ii) = MSFlexGrid1.TextMatrix(ii, 6)
   KODBSV(ii) = MSFlexGrid1.TextMatrix(ii, 7)
   CODBST(ii) = MSFlexGrid1.TextMatrix(ii, 8)
   IBST(MSFlexGrid1.TextMatrix(ii, 1)) = ii
   NPDSBS(ii) = IDSP(MSFlexGrid1.TextMatrix(ii, 1), 1)
   If Abs(NBOOST(ii) * BSTDC(ii) - PDC(IDSP(MSFlexGrid1.TextMatrix(ii, 1), 1))) > 0.002 Then
    MsgBox "Mismatch between individual booster discharge and total discharge for operating booster : " & ii
    Exit Sub
   End If
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
   ElseIf Combo1.Visible = True Then
       MSFlexGrid1 = Combo1
       Combo1.Visible = False
   End If
End Sub
Private Sub Combo1_Click()
     NP = MSFlexGrid1.Row
      MSFlexGrid1 = Combo1
      Combo1.Visible = False
     If Combo1.Text = "YES" Then
       frmGridBOOST.Enabled = False
       frmBMachY.Show
      ElseIf Combo1.Text = "NO" Then
       BSTSP(NP) = MSFlexGrid1.TextMatrix(NP, 5)
       BSTDC(NP) = MSFlexGrid1.TextMatrix(NP, 3)
       BSTH(NP) = MSFlexGrid1.TextMatrix(NP, 4)
       frmGridBOOST.Enabled = False
       frmBMachN.Show
     End If
End Sub


Private Sub Command1_Click()
Me.Hide
frmHGL.Enabled = True
MDIForm1.mnuExec.Item(10).Enabled = True
MDIForm1.tbrMain.Buttons(6).Enabled = True
End Sub

Private Sub Form_Activate()
MSFlexGrid1.Rows = NBST + 1
ii = 0
   For i = 1 To NNMAX
        If (NTYPE(i) = 9) Then
         ii = ii + 1
         MSFlexGrid1.TextMatrix(ii, 0) = ii
         MSFlexGrid1.TextMatrix(ii, 1) = i
       End If
   Next
End Sub


Private Sub Form_Load()
   Dim iCount As Integer
   Dim myArray As Variant
   Combo1.AddItem "YES"
   Combo1.AddItem "NO"
   Combo1.Visible = False
   Left = 20
   Top = 30
   
   'set array
    myArray = Array("Sl. No.", "Node No.", "No. of Pumps", "Discharge (cum/sec)", "Head (m)", "Speed (rpm)", "Suction Side HGL (RL, m)", "Type of NRV", "Machinery Finalised ?")

   MSFlexGrid1.Rows = NBST + 1
   MSFlexGrid1.Cols = TOTALCOLUMNS  'Non-zero based
   MSFlexGrid1.FixedRows = 1
   MSFlexGrid1.FixedCols = 2
   MSFlexGrid1.FocusRect = flexFocusNone
   'add headings to grid
   MSFlexGrid1.Row = 0
   MSFlexGrid1.ColWidth(0) = 600
   MSFlexGrid1.ColWidth(1) = 600
   MSFlexGrid1.TextMatrix(0, 0) = myArray(0)
   MSFlexGrid1.TextMatrix(0, 1) = myArray(1)
   
   For iCount = 2 To TOTALCOLUMNS - 1 'Zero based
      MSFlexGrid1.ColWidth(iCount) = 1000
      MSFlexGrid1.Col = iCount
      MSFlexGrid1.Text = myArray(iCount)
   Next iCount
   
   ii = 0
   For i = 1 To NNMAX
        If (NTYPE(i) = 9) Then
         ii = ii + 1
         MSFlexGrid1.TextMatrix(ii, 0) = ii
         MSFlexGrid1.TextMatrix(ii, 1) = i
       End If
   Next
   
  If Not OpenFile = "" Then
   For ii = 1 To MSFlexGrid1.Rows - 1
     MSFlexGrid1.TextMatrix(ii, 2) = NBOOST(ii)
     MSFlexGrid1.TextMatrix(ii, 3) = BSTDC(ii)
     MSFlexGrid1.TextMatrix(ii, 4) = BSTH(ii)
     MSFlexGrid1.TextMatrix(ii, 5) = BSTSP(ii)
     MSFlexGrid1.TextMatrix(ii, 6) = HGLSUC(ii)
     MSFlexGrid1.TextMatrix(ii, 7) = KODBSV(ii)
     MSFlexGrid1.TextMatrix(ii, 8) = CODBST(ii)
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
   If MSFlexGrid1.Col = 8 Then
     MSHFlexGridEdit MSFlexGrid1, Combo1, KeyAscii
     Combo1.Visible = True
   Else
     MSHFlexGridEdit MSFlexGrid1, Text2, KeyAscii
     Text2.Visible = True
   End If
End Sub
Sub MsFlexGrid1_DblClick()
   If MSFlexGrid1.Col = 8 Then
     MSHFlexGridEdit MSFlexGrid1, Combo1, 32
     Combo1.Visible = True
   Else
     MSHFlexGridEdit MSFlexGrid1, Text2, 32
     Text2.Visible = True
   End If
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
   
   If Not Edt = Combo1 Then
   Edt.Move MSHFlexGrid.Left + MSHFlexGrid.CellLeft, _
      MSHFlexGrid.Top + MSHFlexGrid.CellTop, _
      MSHFlexGrid.CellWidth - 8, _
      MSHFlexGrid.CellHeight - 8
   Else
   Edt.Move MSHFlexGrid.Left + MSHFlexGrid.CellLeft, _
      MSHFlexGrid.Top + MSHFlexGrid.CellTop
   End If
      
      Edt.Visible = True
      ' And make it work.
   Edt.SetFocus
   End Sub
Private Sub Text2_Gotfocus()
If Text2.Visible = False Then
   Exit Sub
   End If
If MSFlexGrid1.Col = 7 Then
  Text2.Visible = False
  NP = MSFlexGrid1.Row
  Iflag_PB = 2
  frmListP.Show
  frmGridBOOST.Enabled = False
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
   If Text2.Visible = False And Combo1.Visible = False Then
   Exit Sub
   ElseIf Text2.Visible = True Then
     MSFlexGrid1 = Text2
     Text2.Visible = False
   ElseIf Combo1.Visible = True Then
     MSFlexGrid1 = Combo1
     Combo1.Visible = False
   End If
End Sub
Sub MSFlexGrid1_LeaveCell()
   If Text2.Visible = False And Combo1.Visible = False Then
   Exit Sub
   ElseIf Text2.Visible = True Then
     MSFlexGrid1 = Text2
     Text2.Visible = False
   ElseIf Combo1.Visible = True Then
     MSFlexGrid1 = Combo1
     Combo1.Visible = False
   End If
End Sub
Private Sub text2_KeyPress(KeyAscii As Integer)
' Delete returns to get rid of beep.
   If KeyAscii = Asc(vbCr) Then KeyAscii = 0
End Sub

