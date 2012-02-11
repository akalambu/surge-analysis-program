
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
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRes 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Execute and View Results"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1950
   ScaleWidth      =   5700
   Begin MSComctlLib.ProgressBar pgrbar 
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Program Executing -----"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   3855
   End
   Begin VB.CommandButton cmdFile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please Wait ...."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "frmRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyXL As Object   ' Variable to hold reference to Microsoft Excel.
Dim iplot As Integer
Dim ExcelWasNotRunning As Boolean ' Flag for final release.

Private Declare Function FindWindow Lib "user32" Alias _
"FindWindowA" (ByVal lpClassName As String, _
               ByVal lpWindowName As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias _
"SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
               ByVal wParam As Long, _
               ByVal lParam As Long) As Long


Dim xlApp As Object   ' Declare variable to hold the reference.
   ' the application, then release the reference.

Sub GetExcel()
   
   
' Test to see if there is a copy of Microsoft Excel already running.
   
   On Error Resume Next   ' Defer error trapping.
' Getobject function called without the first argument returns a
' reference to an instance of the application. If the application isn't
' running, an error occurs.
   
   Set MyXL = GetObject(, "Excel.Application")
   If Err.Number <> 0 Then ExcelWasNotRunning = True
   Err.Clear   ' Clear Err object in case error occurred.

' Check for Microsoft Excel. If Microsoft Excel is running,
' enter it into the Running Object table.
   
   'DetectExcel

' Set the object variable to reference the file you want to see.
   Set MyXL = GetObject("c:\iisc\sap2.xls")


' Show Microsoft Excel through its Application property. Then
' show the actual window containing the file using the Windows
' collection of the MyXL object reference.
   
   MyXL.Application.Visible = True
   MyXL.Parent.windows(1).Visible = True
   
'    Do manipulations of your  file here.
   ' ...
' If this copy of Microsoft Excel was not running when you
' started, close it using the Application property's Quit method.
' Note that when you try to quit Microsoft Excel, the
' title bar blinks and a message is displayed asking if you
' want to save any loaded files.
   'If ExcelWasNotRunning = True Then
      'MyXL.Application.Quit
   'End If

   'Set MyXL = Nothing   ' Release reference to the
                        ' application and spreadsheet.
End Sub

Sub DetectExcel()
' Procedure dectects a running Excel and registers it.
   Const WM_USER = 1024
   Dim hWnd As Long
' If Excel is running this API call returns its handle.
   hWnd = FindWindow("XLMAIN", 0)
   If hWnd = 0 Then   ' 0 means Excel not running.
      Exit Sub
   Else
   ' Excel is running so use the SendMessage API
   ' function to enter it in the Running Object Table.
      SendMessage hWnd, WM_USER + 18, 0, 0
   End If
End Sub

Private Sub cmdFile_Click()
X = Shell("c:\windows\command\edit.com " & "c:\iisc\sap2.res", vbMaximizedFocus)
'frmTextE.Show
End Sub

' Declare necessary API routines:
Sub execute()
Fort_Data
X = Shell("c:\iisc\sapf2r.exe")
DoEvents
pgrbar.Visible = True
MDIForm1.mnuViewItem.Item(20).Enabled = True
MDIForm1.mnuViewItem.Item(30).Enabled = True
MDIForm1.tbrMain.Buttons.Item(9).Enabled = True
MDIForm1.tbrMain.Buttons.Item(10).Enabled = True
End Sub

Private Sub command2_lostfocus()
Dim check As String
For i = 1 To 1000
 pgrbar.Value = i / 10
Next
 pgrbar.Visible = False
 frmRes.Hide
 pgrbar.Value = 0
frmRes.MousePointer = 0
End Sub

Private Sub Form_Load()
Left = 20
Top = 30
pgrbar.Visible = False
End Sub

