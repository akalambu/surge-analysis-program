
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
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "SAP"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmnDialogOpen 
      Left            =   4200
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   3240
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0000
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0454
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0770
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0BC4
            Key             =   "SaveAs"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1018
            Key             =   "Results"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":146C
            Key             =   "Graph"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":18C0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1D14
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2168
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":25BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2A10
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2E68
            Key             =   "Exec"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New Project"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open Project"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Close"
            Object.ToolTipText     =   "Close Project"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save Project "
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SaveAs"
            Object.ToolTipText     =   "Save Project As .."
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Run"
            Object.ToolTipText     =   "Run"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Data"
            Object.ToolTipText     =   "View Data"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Result"
            Object.ToolTipText     =   "View Results"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Graph"
            Object.ToolTipText     =   "View Graph"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "New  Project"
         Index           =   10
         Begin VB.Menu mnuFileItemn 
            Caption         =   "Project Type A"
            Index           =   1
            Shortcut        =   ^A
         End
         Begin VB.Menu mnuFileItemn 
            Caption         =   "Project Type B"
            Index           =   2
            Shortcut        =   ^B
         End
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Open Project"
         Index           =   20
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Close Project"
         Index           =   30
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   31
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Save Project"
         Index           =   40
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Save Project As .."
         Index           =   50
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   51
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Exit"
         Index           =   70
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewItem 
         Caption         =   "Data"
         Index           =   10
         Begin VB.Menu mnuDataItem 
            Caption         =   "General data"
            Index           =   1
            Shortcut        =   ^D
         End
         Begin VB.Menu mnuDataItem 
            Caption         =   "Pump Characteristics"
            Index           =   2
            Shortcut        =   ^P
         End
         Begin VB.Menu mnuDataItem 
            Caption         =   "Alignment Data"
            Index           =   3
            Shortcut        =   ^L
         End
      End
      Begin VB.Menu mnuViewItem 
         Caption         =   "Result"
         Index           =   20
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuViewItem 
         Caption         =   "Graph"
         Index           =   30
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuViewItem 
         Caption         =   "-"
         Index           =   31
      End
      Begin VB.Menu mnuViewItem 
         Caption         =   "Show ToolBar"
         Checked         =   -1  'True
         Index           =   40
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuRun 
      Caption         =   "&Run"
      Begin VB.Menu mnuExec 
         Caption         =   "Execute"
         Index           =   10
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpItem 
         Caption         =   "Help On"
         Index           =   10
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUpToolbar 
         Caption         =   "&Hide ToolBar"
         Index           =   0
      End
      Begin VB.Menu mnuPopUpToolbar 
         Caption         =   "&Show ToolBar"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit Program"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuPopNew 
      Caption         =   "popNew"
      Visible         =   0   'False
      Begin VB.Menu mnuType 
         Caption         =   "Project Type A"
         Index           =   1
      End
      Begin VB.Menu mnuType 
         Caption         =   "Project Type B"
         Index           =   2
      End
   End
   Begin VB.Menu mnuPopData 
      Caption         =   "popData"
      Visible         =   0   'False
      Begin VB.Menu mnuDataItem1 
         Caption         =   "General Data"
         Index           =   1
      End
      Begin VB.Menu mnuDataItem1 
         Caption         =   "Pump Characteristics"
         Index           =   2
      End
      Begin VB.Menu mnuDataItem1 
         Caption         =   "Alignment Data"
         Index           =   3
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ij As Integer
Dim iflag As Integer
Dim xxx As String
Dim Answer As Integer
Dim Confirm As Integer
Option Explicit
Dim X As String
'File Menu Indices
Const FILE_NEW = 10
Const NEW_A = 1
Const NEW_B = 2
Const NEW_C = 3
Const FILE_OPEN = 20
Const FILE_CLOSE = 30
Const FILE_SAVE = 40
Const FILE_SAVEAS = 50
'Const FILE_PRINT = 60
Const FILE_EXIT = 70
'View Menu Indices
Const VIEW_DATA = 10
Const DATA_G = 1
Const DATA_P = 2
Const DATA_A = 3
Const VIEW_RESULT = 20
Const VIEW_GRAPH = 30
Const VIEW_TOOLBAR = 40
'Help Menu Indices
Const HELP_HELPON = 10
Const EXEC = 10

Private Sub MDIForm_Load()
MDIForm1.AutoShowChildren = False
cmnDialogOpen.Filter = "SAP Project (*.sap)|*.sap|All Files (*.*)|*.*"
Not_Active
End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuPopUp
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set MDIForm1 = Nothing
   End Sub

Private Sub mnuPopupToolbar_Click(Index As Integer)
   mnuViewItem_Click VIEW_TOOLBAR
End Sub

Private Sub mnuQuit_Click()
   mnuFileItem_Click FILE_EXIT
End Sub

Private Sub mnuDataItem1_Click(Index As Integer)
Select Case Index
 Case DATA_G
   mnuDataItem_Click DATA_G
 Case DATA_P
   mnuDataItem_Click DATA_P
 Case DATA_A
   mnuDataItem_Click DATA_A
End Select
End Sub
Private Sub mnuType_Click(Index As Integer)
Select Case Index
 Case NEW_A
   mnuFileItemn_Click NEW_A
 Case NEW_B
   mnuFileItemn_Click NEW_B
 Case NEW_C
   mnuFileItemn_Click NEW_C
End Select
End Sub

Private Sub mnuFileItem_Click(Index As Integer)
    Select Case Index
        Case FILE_OPEN
         Project_Open
         Case FILE_CLOSE
         Close_Project
         Case FILE_SAVE
         Project_Save
       Case FILE_SAVEAS
         Save_project_As
       Case FILE_EXIT
         End_Project
       End Select
   End Sub

Private Sub mnuDataItem_Click(Index As Integer)
Select Case Index
Case DATA_G
  frmTitle.Show
 Case DATA_P
   MDIForm1.mnuExec.Item(10).Enabled = False
   MDIForm1.tbrMain.Buttons(6).Enabled = False
 If PTYPE = "TYPEA" Then
   frmPumpChA.Show
   'ElseIf PTYPE = "TYPEB" Then
   'frmHGL.Show
  End If
  Case DATA_A
  MDIForm1.mnuExec.Item(10).Enabled = False
  MDIForm1.tbrMain.Buttons(6).Enabled = False
  If PTYPE = "TYPEA" Then
   frmAlignA.Show
  ElseIf PTYPE = "TYPEB" Then
   frmAlignB.Show
  End If
 End Select
End Sub

Private Sub mnuFileItemn_Click(Index As Integer)
Select Case Index
Case NEW_A
 PTYPE = "TYPEA"
Case NEW_B
 PTYPE = "TYPEB"
 SIML = "TYPEB"
Case NEW_C
 PTYPE = "TYPEC"
End Select
  Start_New_Project
 End Sub

' view item except data

Private Sub mnuViewItem_Click(Index As Integer)
Select Case Index

Case VIEW_RESULT
    X = Shell("c:\windows\command\edit.com " & "c:\iisc\sap2.res", vbMaximizedFocus)
Case VIEW_GRAPH
   frmRes.GetExcel
Case VIEW_TOOLBAR
    mnuViewItem(VIEW_TOOLBAR).Checked = Not mnuViewItem(VIEW_TOOLBAR).Checked
    tbrMain.Visible = mnuViewItem(VIEW_TOOLBAR).Checked
    If mnuViewItem(VIEW_TOOLBAR).Checked Then
     mnuPopUpToolbar(0).Enabled = True
     mnuPopUpToolbar(1).Enabled = False
    Else
     mnuPopUpToolbar(0).Enabled = False
     mnuPopUpToolbar(1).Enabled = True
    End If
End Select
End Sub

   Private Sub mnuHelpItem_Click(Index As Integer)
   'Handle  Help event
   Select Case Index
       Case HELP_HELPON
   End Select
   End Sub
   Private Sub mnuExec_Click(Index As Integer)
    
   
   Select Case Index
    Case EXEC
       frmRes.Show
       frmRes.Command2.SetFocus
       frmRes.execute
       End Select
End Sub
   
Private Sub tbrMain_ButtonClick(ByVal Button As MSComCtlLib.Button)
   'Send Toolbar events to the appropriate Menu item
      Select Case Button.Key
      Case "New"
         PopupMenu mnuPopNew
      Case "Open"
         mnuFileItem_Click FILE_OPEN
      Case "Close"
         mnuFileItem_Click FILE_CLOSE
      Case "Save"
         mnuFileItem_Click FILE_SAVE
      Case "SaveAs"
         mnuFileItem_Click FILE_SAVEAS
      Case "Exit"
         mnuFileItem_Click FILE_EXIT
      Case "Run"
         mnuExec_Click EXEC
      Case "Data"
         PopupMenu mnuPopData
      Case "Result"
         mnuViewItem_Click VIEW_RESULT
      Case "Graph"
         mnuViewItem_Click VIEW_GRAPH
      Case "Help"
         mnuHelpItem_Click HELP_HELPON
      End Select
   End Sub

Private Sub mnupop_Click(Index As Integer)
   mnuViewItem_Click VIEW_TOOLBAR
End Sub
  


Private Sub Not_Active()
mnuFileItem.Item(30).Enabled = False
mnuFileItem.Item(40).Enabled = False
mnuFileItem.Item(50).Enabled = False
mnuView.Enabled = False
mnuRun.Enabled = False
tbrMain.Buttons.Item(3).Enabled = False
tbrMain.Buttons.Item(4).Enabled = False
tbrMain.Buttons.Item(5).Enabled = False
tbrMain.Buttons.Item(6).Enabled = False
tbrMain.Buttons.Item(8).Enabled = False
tbrMain.Buttons.Item(9).Enabled = False
tbrMain.Buttons.Item(10).Enabled = False
End Sub
Private Sub Get_Active()
mnuFileItem.Item(30).Enabled = True
mnuFileItem.Item(40).Enabled = True
mnuFileItem.Item(50).Enabled = True
mnuView.Enabled = True
mnuRun.Enabled = True
tbrMain.Buttons.Item(3).Enabled = True
tbrMain.Buttons.Item(4).Enabled = True
tbrMain.Buttons.Item(5).Enabled = True
tbrMain.Buttons.Item(6).Enabled = True
tbrMain.Buttons.Item(8).Enabled = True
'tbrMain.Buttons.Item(9).Enabled = True
'tbrMain.Buttons.Item(10).Enabled = True
End Sub

Private Sub Start_New_Project()
    If Not iflag = 0 Then
     Save_All
     If Answer = 3 Then Exit Sub
    End If
    
' open the starting frame depending
' on the type of the project
    
    For ij = 2 To Forms.Count - 1
     Unload Forms(2)
    Next
    Get_Active
    Reset_All
    OpenFile = ""
    SaveFile = ""
    frmTitle.Show
    iflag = 1
    If PTYPE = "TYPEA" Then
       MDIForm1.mnuDataItem(2).Enabled = True
       MDIForm1.mnuDataItem1(2).Enabled = True
    ElseIf PTYPE = "TYPEB" Then
       MDIForm1.mnuDataItem(2).Enabled = False
       MDIForm1.mnuDataItem1(2).Enabled = False
    End If
End Sub
Private Sub Project_Open()
   If Not iflag = 0 Then
      Save_All
      If Answer = 3 Then Exit Sub
   End If
   For ij = 2 To Forms.Count - 1
     Unload Forms(2)
   Next
   iflag = 0
   OpenFile = ""
   SaveFile = ""
   Not_Active
   Reset_All
   Open_It
   End Sub
   
Private Sub Close_Project()
   If Not iflag = 0 Then
     Save_All
     If Answer = 3 Then Exit Sub
     For ij = 2 To Forms.Count - 1
       Unload Forms(2)
     Next
     iflag = 0
     OpenFile = ""
     SaveFile = ""
   End If
   Not_Active
   Reset_All
End Sub
Private Sub End_Project()
   If Not iflag = 0 Then
     Save_All
     If Answer = 3 Then Exit Sub
     iflag = 0
     OpenFile = ""
     SaveFile = ""
   End If
   End
End Sub

Private Sub Project_Save()
  If OpenFile = "" And SaveFile = "" Then
     Show_It
     Do While Not xxx = "" And Not SaveFile = ""
       Mesg_2
       If Confirm = 1 Then
          xxx = ""
       ElseIf Confirm = 2 Then
          Show_It
       Else
         Exit Sub
       End If
     Loop
   Else
    If OpenFile = "" Then
     OpenFile = SaveFile
    ElseIf SaveFile = "" Then
     SaveFile = OpenFile
    End If
   End If
     If Not SaveFile = "" Then
     Save_Project
   End If
End Sub
Private Sub Save_project_As()
  Show_It
  Do While Not xxx = "" And Not SaveFile = ""
     Mesg_2
     If Confirm = 1 Then
       xxx = ""
     ElseIf Confirm = 2 Then
       Show_It
     Else
       Exit Sub
     End If
  Loop
    If Not SaveFile = "" Then
      OpenFile = SaveFile
      Save_Project
    End If
End Sub

 Private Sub Show_It()
   SaveFile = ""
   cmnDialogOpen.FileName = ""
   cmnDialogOpen.ShowSave
   SaveFile = cmnDialogOpen.FileName
   xxx = Dir(SaveFile)
 End Sub
 
 Private Sub Save_All()
    Answer = 1
    SaveFile = ""
    Mesg_1
    Do Until Not Answer = 1 Or Not SaveFile = ""
     Save_Project_File
     If SaveFile = "" Then
       Mesg_1
     End If
    Loop
    If Answer = 1 Then
     Save_Project
    End If
End Sub
Private Sub Mesg_1()
   Dim intRet  As Integer
   intRet = MsgBox("Project is not saved. Do you want to save it ?", _
                 vbYesNoCancel + vbQuestion, "Response requires")
   Select Case intRet
    Case vbYes
     Answer = 1
    Case vbNo
     Answer = 2
    Case vbCancel
     Answer = 3
   End Select
End Sub
Private Sub Mesg_2()
  Dim intRet  As Integer
  intRet = MsgBox("File already exist. Do you want to overwrite it ?", _
               vbYesNoCancel + vbQuestion, "Response requires")
  Select Case intRet
    Case vbYes
     Confirm = 1
    Case vbNo
     Confirm = 2
    Case vbCancel
     Confirm = 3
  End Select
End Sub

Private Sub Save_Project_File()
  If OpenFile = "" Then
    Save_File
    Do While Not xxx = "" And Not SaveFile = ""
       Mesg_2
       If Confirm = 1 Then
         xxx = ""
       ElseIf Confirm = 2 Then
         Save_File
       Else: Exit Sub
       End If
     Loop
  Else
    SaveFile = OpenFile
  End If
  End Sub
 Private Sub Save_File()
      cmnDialogOpen.FileName = ""
      cmnDialogOpen.ShowSave
      SaveFile = cmnDialogOpen.FileName
      xxx = Dir(SaveFile)
 End Sub
 
 Private Sub Open_It()
      cmnDialogOpen.FileName = ""
      cmnDialogOpen.ShowOpen
      OpenFile = cmnDialogOpen.FileName
      xxx = Dir(OpenFile)
      If xxx = "" Then
       MsgBox "File Not Found"
      ElseIf Not OpenFile = "" Then
         iflag = 1
         Open_Project
         Get_Active
          frmTitle.Show
          If PTYPE = "TYPEA" Then
          MDIForm1.mnuDataItem(2).Enabled = True
          MDIForm1.mnuDataItem1(2).Enabled = True
          ElseIf PTYPE = "TYPEB" Then
            MDIForm1.mnuDataItem(2).Enabled = False
            MDIForm1.mnuDataItem1(2).Enabled = False
          End If
      End If
End Sub
    



          
          
          
         
 



 


