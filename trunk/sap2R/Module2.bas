
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

Attribute VB_Name = "Module2"
'Declaration, Collecting and Writing Data for Project Type B

Public Declare Sub PMPYES Lib "c:\iisc\pumpch.dll" _
Alias "_PMPYES@44" (ByRef A As Long, ByRef B As Single, _
ByRef C As Single, ByRef D As Single, ByRef E As Single, _
ByRef QQ As Single, ByRef HH As Single, ByRef ETA As Single, _
ByRef WH As Single, ByRef WB As Single, ByRef FKNR As Single)

Public Declare Sub PMPNO Lib "c:\iisc\pumpch.dll" _
Alias "_PMPNO@12" (ByRef X As Long, _
ByRef Y As Single, ByRef Z As Single)

Dim TempString As String
Dim NSTR As String
Global SIML As String
Global TYPEN(0 To 75) As String

'Temp Variables
Global NP As Integer
Global NN As Integer





Global NPMAX As Integer
Global NNMAX As Integer
Global Iflag_Err As Integer
Global Iflag_PB As Integer
Global Iflag_But(1 To 20) As Integer

Global NORD  As Integer
Global NCJN  As Integer
Global NDJN  As Integer
Global NRES  As Integer
Global NSOU  As Integer
Global NCDS  As Integer
Global NOBS  As Integer
Global NPMP  As Integer
Global NBST  As Integer

'frmGridPipe
Global IP(0 To 75) As Integer
Global IND1(0 To 75) As Integer
Global IND2(0 To 75) As Integer
Global PDC(0 To 75) As Single
Global PDIA(0 To 75) As Single
Global PLEN(0 To 75) As Single
Global CHST(0 To 75) As Single

'frmGridNode
Global NNO(0 To 75) As Integer
Global NTYPE(0 To 75) As Integer
Global NUSP(0 To 75) As Integer
Global NDSP(0 To 75) As Integer
Global TYPEROW As Integer

'frmGridUS
Global IUSP(0 To 75, 0 To 10) As Integer
Global IDSP(0 To 75, 0 To 10) As Integer
Global NoNode As Integer
Global NPUSRS(0 To 50) As Integer
Global NPUSOB(0 To 50) As Integer
Global NPUSCD(0 To 50) As Integer

'frmgridHGL
Global HGL(0 To 75) As Single

'frmGridRES
  Global IRES(0 To 75) As Integer
  Global RESWL(0 To 50) As Single
  Global RESDC(0 To 50) As Single
  
'frmGridSOV
  Global ISOU(0 To 75) As Integer
  Global SOUWL(0 To 50) As Single
  Global SOUDC(0 To 50) As Single
  Global CODEVK(0 To 50) As String
  Global SOUVK(0 To 10) As Single

'frmGridCDS
  Global ICDS(0 To 75) As Integer
  Global QOCDS(0 To 50) As Single
  Global HGL1CD(0 To 50) As Single
  Global HGL2CD(0 To 50) As Single
  Global NPDSPM(0 To 50) As Integer
  Global NPDSBS(0 To 50) As Integer
  Global NPDSSO(0 To 50) As Integer
 
'frmGridCDS
  Global IOBS(0 To 75) As Integer
  Global QOOBS(0 To 50) As Single
  Global HGL1OB(0 To 50) As Single
  Global HGL2OB(0 To 50) As Single

'frmGridPump
   
   Global KODPHV(0 To 50) As Integer
   Global DLYPH(0 To 20) As Single
   Global TCLOSEP(0 To 20) As Single
   Global TRAPIDP(0 To 20) As Single
   Global TSLOWP(0 To 20) As Single
   Global NPUMPS(0 To 20) As Integer
   Global PUMPDC(0 To 20) As Single
   Global PUMPH(0 To 20) As Single
   Global PUMPSP(0 To 20) As Single
   Global SUMPWL(0 To 20) As Single
   Global CODPMP(0 To 20) As String
   Global IPMP(0 To 75) As Integer
   Global FKNRR(0 To 20) As Single
      
'frmGridbOOST
   Global NB As Integer
   Global KODBSV(0 To 50) As Integer
   Global DLYBS(0 To 20) As Single
   Global TCLOSEB(0 To 20) As Single
   Global TRAPIDB(0 To 20) As Single
   Global TSLOWB(0 To 20) As Single
   Global NBOOST(0 To 20) As Integer
   Global BSTDC(0 To 20) As Single
   Global BSTH(0 To 20) As Single
   Global BSTSP(0 To 20) As Single
   Global HGLSUC(0 To 20) As Single
   Global CODBST(0 To 20) As String
   Global IBST(0 To 75) As Integer
   Global FKNRRB(0 To 20) As Single
   
   'frmPMechY
   Global DCP(0 To 20, 0 To 100) As Single
   Global HEADP(0 To 20, 0 To 100) As Single
   Global EFFP(0 To 20, 0 To 100) As Single
   Global NPUMPCH(0 To 20) As Integer
   Global GD2PMP(0 To 20) As Single
   Global GD2PM(0 To 20) As Single
   Global QRPMP(0 To 20) As Single
   Global HRPMP(0 To 20) As Single
   Global EFFRP(0 To 20) As Single
   Global HSHPMP(0 To 20) As Single
   Global CODNRR(0 To 20) As String
   
   'frmBMechY
   Global DCB(0 To 20, 0 To 100) As Single
   Global HEADB(0 To 20, 0 To 100) As Single
   Global EFFB(0 To 20, 0 To 100) As Single
   Global NBSTCH(0 To 20) As Integer
   Global GD2BST(0 To 20) As Single
   Global GD2BM(0 To 20) As Single
   Global QRBST(0 To 20) As Single
   Global HRBST(0 To 20) As Single
   Global EFFRB(0 To 20) As Single
   Global HSHBST(0 To 20) As Single
   Global CODNRB(0 To 20) As String
      
  'frmpmechn
   Global TYPCHP(0 To 20) As String
   
   'frmbmechn
    Global TYPCHB(0 To 20) As String
   
   'frmgridptrip
   Global ITRIPP(0 To 20) As Integer
   
   'frmgridbtrip
   Global ITRIPB(0 To 20) As Integer
   
   'frmgridwave
   Global WV(0 To 75) As Single
   Global PIPEMAT(0 To 75) As String
   
   'frmSteel
   Global WTHICK(0 To 75) As Single
   Global LTHICK(0 To 75) As Single
   Global GTHICK(0 To 75) As Single
   Global CODEG(0 To 75) As String
   Global CODEL(0 To 75) As String
   
  'frmDI
   Global CODECL(0 To 75) As String
   Global CODEDIL(0 To 75) As String
   Global THICKDIL(0 To 75) As Single
   
  'frmCI
   Global CODECI(0 To 75) As String
   
  'frmPSC and frmBWSC
   Global CORET(0 To 75) As Single
   Global COATT(0 To 75) As Single
   Global STBYV(0 To 75) As Single
   
   'frmAC
   Global THICKAC(0 To 75) As Single
   
   'calculated variables
   Global WHP(0 To 20, 0 To 90) As Single
   Global WBP(0 To 20, 0 To 90) As Single
   Global WHB(0 To 20, 0 To 90) As Single
   Global WBB(0 To 20, 0 To 90) As Single
   
'frmOutPutB
Global NPR1 As Integer
Global NPR2 As Integer
Global NPR3 As Integer
Global CHPR1 As Single
Global CHPR2 As Single
Global CHPR3 As Single
Global ZPR1 As Single
Global ZPR2 As Single
Global ZPR3 As Single
Global TMAX As Single
Global CODSIM As String
Global ISPLOT As Integer

'frmGridCHB
Global NPRNT As Integer
Global NPPRNT(0 To 100) As Integer
Global CHPRNT(0 To 100) As Single
   
'frmAlignB
Global NALIGN(0 To 75) As Integer
Global CHAIN(0 To 75, 0 To 1000) As Single
Global GL(0 To 75, 0 To 1000) As Single

'frmAnalyB
Global CODECS As String
Global CODEPR As String
Global ISEL As Integer

'frmColumnB
Global NLCS As Integer
Global CS_Data(1 To 2, 0 To 2) As Single

'frmavesB
Global CODEAC As String
Global NPAC As Integer
Global GLAC As Single
Global ACC As Single
Global DCAC As Single
Global ACNRV As String
Global KACTYP As Integer
Global DORBY As Single

'frmGridOWSTB
Global NOSTD As Integer
Global OST_Data(0 To 20, 0 To 6) As Single
'frmGridDPCB
Global NDPCVD As Integer
Global DPC_Data(0 To 20, 0 To 2) As Single
'frmGridZVB
Global NZVD As Integer
Global ZV_Data(0 To 20, 0 To 2) As Single
'frmGridInB
Global NNRVD As Integer
Global INR_Data(0 To 20, 0 To 3) As Single
'frmGridAVB
Global NAVD As Integer
Global AVV_Data(0 To 20, 0 To 3) As Single
'frmGridACVB
Global NACVD As Integer
Global AC_Data(0 To 20, 0 To 3) As Single
'frmGridSPB
Global NSSD As Integer
Global SP_Data(0 To 20, 0 To 4) As Single

'frmSRVB
Global NLSRV As Integer
Global CODESV As String
Global SRV_Data(0 To 4, 0 To 7) As Single

'frmVBO
Global CODIV As String
Global CODBIV As String
Global NPIV As Integer
Global CHIV As Single
Global DLYIV As Single
Global TCIV As Single
Global HDELB As Single
Global SZBIV As Single
Global TOBIV As Single
Global DLYBIV As Single
Global TCBIV As Single
Global TOPGB As Single

Global DELV As String
Global RESVT1 As String
Global RESVT2 As String
Global DLYDS As Single
Global TOPDS As Single
Global KODEDS As Integer

'frmPath
Global NPATH As Integer
Global NPPATH(0 To 10) As Integer
Global IPPATH(0 To 10, 0 To 75) As Integer

'============================================
' Validate the nodal information based on type of the node
'=============================================

Public Sub CheckNode()
Dim ii As Integer
Dim jj As Integer

For ii = 1 To NNMAX
  If (NTYPE(ii) = 1) Or (NTYPE(ii) = 6) Or (NTYPE(ii) = 7) Or (NTYPE(ii) = 9) Then
      If (Not NUSP(ii) = 1) Or (Not NDSP(ii) = 1) Then
        Iflag_Err = 1
        MsgBox ("Error in Input of No. of U/S and/or D/S Pipes for Node No: " & ii)
        Exit Sub
      Else
        If (Not IND2(IUSP(ii, 1)) = ii) Then
         Iflag_Err = 1
         MsgBox ("Error in Input of End Nodes of  Pipe No: " & IUSP(ii, 1))
         Exit Sub
        End If
        If (Not IND1(IDSP(ii, 1)) = ii) Then
         Iflag_Err = 1
         MsgBox ("Error in Input of End Nodes of  Pipe No: " & IDSP(ii, 1))
         Exit Sub
        End If
      End If
  ElseIf (NTYPE(ii) = 8) Or (NTYPE(ii) = 5) Then
       If (Not NUSP(ii) = 0) Or (Not NDSP(ii) = 1) Then
        Iflag_Err = 1
        MsgBox ("Error in Input of No. of U/S and/or D/S Pipes for Node No: " & ii)
        Exit Sub
      Else
        If (Not IND1(IDSP(ii, 1)) = ii) Then
         Iflag_Err = 1
         MsgBox ("Error in Input of End Nodes of  Pipe No: " & IDSP(ii, 1))
         Exit Sub
        End If
      End If
  ElseIf (NTYPE(ii) = 4) Then
       If (Not NUSP(ii) = 1) Or (Not NDSP(ii) = 0) Then
       Iflag_Err = 1
        MsgBox ("Error in Input of No. of U/S and/or D/S Pipes for Node No: " & ii)
        Exit Sub
      Else
        If (Not IND2(IUSP(ii, 1)) = ii) Then
         Iflag_Err = 1
         MsgBox ("Error in Input of End Nodes of  Pipe No: " & IUSP(ii, 1))
         Exit Sub
        End If
      End If
  ElseIf (NTYPE(ii) = 2) Then
       If (Not NUSP(ii) > 1) Or (Not NDSP(ii) = 1) Then
        Iflag_Err = 1
        MsgBox ("Error in Input of No. of U/S and/or D/S Pipes for Node No: " & ii)
        Exit Sub
       Else
        If (Not IND1(IDSP(ii, 1)) = ii) Then
         Iflag_Err = 1
         MsgBox ("Error in Input of End Nodes of  Pipe No: " & IDSP(ii, 1))
         Exit Sub
        End If
         For jj = 1 To NUSP(ii)
          If (Not IND2(IUSP(ii, jj)) = ii) Then
           Iflag_Err = 1
           MsgBox ("Error in Input of U/S Pipe at Combining Junction, Node No: " & ii & ", Pipe No: " & IUSP(ii, jj))
           Exit Sub
          End If
         Next
       End If
      
  ElseIf (NTYPE(ii) = 3) Then
       If (Not NUSP(ii) = 1) Or (Not NDSP(ii) > 1) Then
        Iflag_Err = 1
        MsgBox ("Error in Input of No. of U/S and/or D/S Pipes for Node No: " & ii)
        Exit Sub
       Else
        If (Not IND2(IUSP(ii, 1)) = ii) Then
         Iflag_Err = 1
         MsgBox ("Error in Input of End Nodes of  Pipe No: " & IUSP(ii, 1))
         Exit Sub
        End If
         For jj = 1 To NDSP(ii)
          If (Not IND1(IDSP(ii, jj)) = ii) Then
           Iflag_Err = 1
           MsgBox ("Error in Input of D/S Pipe at Dividing Junction, Node No: " & ii & ", Pipe No: " & IDSP(ii, jj))
           Exit Sub
          End If
         Next
       End If
  End If
  Next

End Sub

' Validate the correctness of discharge
'========================================
Public Sub CheckDis()
  Dim ii As Integer
  Dim jj As Integer
  Dim dsum As Single
  
  For ii = 1 To NNMAX
   If (NTYPE(ii) = 1) Or (NTYPE(ii) = 6) Or (NTYPE(ii) = 7) Or (NTYPE(ii) = 9) Then
      If Not Abs(PDC(IUSP(ii, 1)) - PDC(IDSP(ii, 1))) < 0.002 Then
        Iflag_Err = 1
        MsgBox ("U/S and D/S Discharges do not Match at Node No: " & ii)
        Exit Sub
      End If
   ElseIf (NTYPE(ii) = 2) Then
    dsum = 0
    For jj = 1 To NUSP(ii)
     dsum = dsum + PDC(IUSP(ii, jj))
    Next
    If Not Abs(PDC(IDSP(ii, 1)) - dsum) < 0.002 Then
        Iflag_Err = 1
        MsgBox ("U/S and D/S Discharges do not Match at Combining Junction, Node No: " & ii)
        Exit Sub
    End If
   ElseIf (NTYPE(ii) = 3) Then
    dsum = 0
    For jj = 1 To NDSP(ii)
     dsum = dsum + PDC(IDSP(ii, jj))
    Next
    If Not Abs(PDC(IUSP(ii, 1)) - dsum) < 0.002 Then
        Iflag_Err = 1
        MsgBox ("U/S and D/S Discharges do not Match at Dividing Junction, Node No: " & ii)
        Exit Sub
    End If
   End If
Next

End Sub

Public Sub CollectType()
  Dim ii As Integer
  NORD = 0
  NCJN = 0
  NDJN = 0
  NRES = 0
  NSOU = 0
  NCDS = 0
  NOBS = 0
  NPMP = 0
  NBST = 0
  
   For ii = 1 To NNMAX
     Select Case NTYPE(ii)
       Case 1
         NORD = NORD + 1
       Case 2
         NCJN = NCJN + 1
       Case 3
         NDJN = NDJN + 1
       Case 4
         NRES = NRES + 1
       Case 5
         NSOU = NSOU + 1
       Case 6
         NCDS = NCDS + 1
       Case 7
         NOBS = NOBS + 1
       Case 8
         NPMP = NPMP + 1
       Case 9
         NBST = NBST + 1
     End Select
  Next
End Sub
Public Sub Data_Fort_B1()
Write #1, PROJECT
Write #1, PCASE
Write #1, PTYPE
Write #1, NPMAX, NNMAX
For i = 1 To NPMAX
 Write #1, IP(i), IND1(i), IND2(i), PDC(i), PDIA(i), PLEN(i), WV(i), CHST(i)
Next
End Sub
Public Sub Data_Fort_B2()
For i = 1 To NNMAX
  Select Case NTYPE(i)
  Case 1
    NSTR = "ORD"
  Case 2
    NSTR = "CJN"
  Case 3
    NSTR = "DJN"
  Case 4
    NSTR = "RES"
  Case 5
    NSTR = "SOU"
  Case 6
    NSTR = "CDS"
  Case 7
    NSTR = "OBS"
  Case 8
    NSTR = "PMP"
  Case 9
    NSTR = "BST"
  End Select
  Write #1, NNO(i), NSTR, NUSP(i), NDSP(i)
  For j = 1 To NUSP(i)
   Write #1, IUSP(i, j)
  Next
  For j = 1 To NDSP(i)
   Write #1, IDSP(i, j)
  Next
Next
End Sub
Public Sub Data_Save_B2()
For i = 1 To NNMAX
  Write #1, NNO(i), NTYPE(i), NUSP(i), NDSP(i)
  For j = 1 To NUSP(i)
   Write #1, IUSP(i, j)
  Next
  For j = 1 To NDSP(i)
   Write #1, IDSP(i, j)
  Next
Next
End Sub
Public Sub Data_Fort_B3()
For i = 1 To NNMAX
 Write #1, IRES(i), ISOU(i), ICDS(i), IOBS(i), IPMP(i), IBST(i), HGL(i)
Next
Write #1, NORD, NCJN, NDJN, NRES, NSOU, NCDS, NOBS, NPMP, NBST
For jj = 1 To NRES
 Write #1, RESWL(jj), RESDC(jj), NPUSRS(jj)
Next
If NRES = 1 Then
  Write #1, KODEDS, DLYDS, TOPDS
End If
For jj = 1 To NSOU
 Write #1, SOUWL(jj), SOUDC(jj), NPDSSO(jj)
Next
For jj = 1 To NCDS
 Write #1, HGL1CD(jj), HGL2CD(jj), QOCDS(jj), NPUSCD(jj)
Next
For jj = 1 To NOBS
 Write #1, HGL1OB(jj), HGL2OB(jj), QOOBS(jj), NPUSOB(jj)
Next
For jj = 1 To NPMP
 Write #1, NPUMPS(jj), PUMPDC(jj), QRPMP(jj), PUMPH(jj), HRPMP(jj), EFFRP(jj), PUMPSP(jj), CODPMP(jj), GD2PMP(jj), GD2PM(jj), SUMPWL(jj), KODPHV(jj), CODNRR(jj), FKNRR(jj), NPDSPM(jj), ITRIPP(jj)
Next
For jj = 1 To NBST
 Write #1, NBOOST(jj), BSTDC(jj), QRBST(jj), BSTH(jj), HRBST(jj), EFFRB(jj), BSTSP(jj), CODBST(jj), GD2BST(jj), GD2BM(jj), HGLSUC(jj), KODBSV(jj), CODNRB(jj), FKNRRB(jj), NPDSBS(jj), ITRIPB(jj)
Next
End Sub
Public Sub Data_Fort_B4()
For jj = 1 To NPMP
For kk = 1 To 89
    Write #1, WHP(jj, kk), WBP(jj, kk)
  Next
Next
For jj = 1 To NBST
For kk = 1 To 89
    Write #1, WHB(jj, kk), WBB(jj, kk)
  Next
Next
End Sub
Public Sub Data_Fort_B5()
For jj = 1 To NPMP
 If KODPHV(jj) = 1 Then
  Write #1, DLYPH(jj)
 ElseIf KODPHV(jj) = 2 Then
  Write #1, TCLOSEP(jj), DLYPH(jj)
 ElseIf KODPHV(jj) = 3 Then
  Write #1, TRAPIDP(jj), TSLOWP(jj), DLYPH(jj)
 ElseIf KODPHV(jj) = 4 Then
  Write #1, TCLOSEP(jj)
 ElseIf KODPHV(jj) = 5 Then
  Write #1, TRAPIDP(jj), TSLOWP(jj)
 End If
Next

For jj = 1 To NBST
 If KODBSV(jj) = 1 Then
  Write #1, DLYBS(jj)
 ElseIf KODBSV(jj) = 2 Then
  Write #1, TCLOSEB(jj), DLYBS(jj)
 ElseIf KODBSV(jj) = 3 Then
  Write #1, TRAPIDB(jj), TSLOWB(jj), DLYBS(jj)
 ElseIf KODBSV(jj) = 4 Then
  Write #1, TCLOSEB(jj)
 ElseIf KODBSV(jj) = 5 Then
  Write #1, TRAPIDB(jj), TSLOWB(jj)
 End If
Next

End Sub
Public Sub Data_Fort_B6()
For ii = 1 To NPMAX
 Write #1, NALIGN(ii)
 For jj = 1 To NALIGN(ii)
  Write #1, CHAIN(ii, jj), GL(ii, jj)
 Next
Next
End Sub
Public Sub Data_Fort_B7()
Write #1, NPR1, CHPR1, ZPR1
Write #1, NPR2, CHPR2, ZPR2
Write #1, NPR3, CHPR3, ZPR3
Write #1, ISPLOT
Write #1, NPATH
For jj = 1 To NPATH
 Write #1, NPPATH(jj)
 For ii = 1 To NPPATH(jj)
  Write #1, IPPATH(jj, ii)
 Next
Next
Write #1, NPRNT
For i = 1 To NPRNT
 Write #1, NPPRNT(i), CHPRNT(i)
Next
Write #1, CODSIM
If CODSIM = "YES" Then
Write #1, TMAX
End If

End Sub
Public Sub GetData_B()

With frmTitle
 PROJECT = .txtName.Text
 PCASE = .txtCase.Text
End With
 
With frmTranMainB
  NPMAX = .txtNoPipe.Text
  NNMAX = .txtNoNodes.Text
End With

With frmOutPutB
ISPLOT = Val(.txtPNo.Text)
NPR1 = Val(.txtPN1.Text)
NPR2 = Val(.txtPN2.Text)
NPR3 = Val(.txtPN3.Text)
CHPR1 = Val(.txtCh1.Text)
CHPR2 = Val(.txtCh2.Text)
CHPR3 = Val(.txtCh3.Text)
ZPR1 = Val(.txtRL1.Text)
ZPR2 = Val(.txtRL2.Text)
ZPR3 = Val(.txtRL3.Text)
CODSIM = .cmbSimTime.Text
If .cmbSimTime.Text = "YES" Then
TMAX = Val(.txtSimT.Text)
End If
End With

With frmAnalyB
   CODECS = .cmbColSep.Text
   CODEPR = .cmbAnProt.Text
End With

End Sub
Public Sub Write_Device()
Write #1, CODECS
If CODECS = "YES" Then
Write #1, NLCS
For jj = 1 To NLCS
Write #1, CS_Data(jj, 0), CS_Data(jj, 1), CS_Data(jj, 2)
Next
End If

Write #1, CODEPR
If CODEPR = "YES" Then
Write #1, CODEAC
If CODEAC = "YES" Then
Write #1, NPAC, GLAC, ACC, DCAC, ACNRV, KACTYP, DORBY
End If
Write #1, NOSTD
If NOSTD > 0 Then
For i = 0 To NOSTD - 1
 Write #1, OST_Data(i, 0), OST_Data(i, 1), OST_Data(i, 2), OST_Data(i, 3), OST_Data(i, 4), OST_Data(i, 5), OST_Data(i, 6)
Next
End If
Write #1, NZVD
If NZVD > 0 Then
For i = 0 To NZVD - 1
  Write #1, ZV_Data(i, 0), ZV_Data(i, 1), ZV_Data(i, 2)
Next
End If
Write #1, NDPCVD
If NDPCVD > 0 Then
For i = 0 To NDPCVD - 1
 Write #1, DPC_Data(i, 0), DPC_Data(i, 1), DPC_Data(i, 2)
Next
End If
Write #1, NNRVD
If NNRVD > 0 Then
For i = 0 To NNRVD - 1
  Write #1, INR_Data(i, 0), INR_Data(i, 1), INR_Data(i, 2), INR_Data(i, 3)
Next
End If
Write #1, NAVD
If NAVD > 0 Then
For i = 0 To NAVD - 1
  Write #1, AVV_Data(i, 0), AVV_Data(i, 1), AVV_Data(i, 2), AVV_Data(i, 3)
Next
End If
Write #1, NACVD
If NACVD > 0 Then
For i = 0 To NACVD - 1
Write #1, AC_Data(i, 0), AC_Data(i, 1), AC_Data(i, 2), AC_Data(i, 3)
Next
End If
Write #1, NSSD
If NSSD > 0 Then
For i = 0 To NSSD - 1
  Write #1, SP_Data(i, 0), SP_Data(i, 1), SP_Data(i, 2), SP_Data(i, 3), SP_Data(i, 4)
Next
End If
Write #1, NLSRV
If NLSRV > 0 Then
 For jj = 1 To NLSRV
  Write #1, SRV_Data(jj, 0), SRV_Data(jj, 1), SRV_Data(jj, 2), SRV_Data(jj, 3), SRV_Data(jj, 4), SRV_Data(jj, 5), SRV_Data(jj, 6), SRV_Data(jj, 7)
 Next
End If
Write #1, CODIV
If CODIV = "YES" Then
  Write #1, NPIV, CHIV, DLYIV, TCIV
  Write #1, CODBIV
  If CODBIV = "YES" Then
    Write #1, HDELB, SZBIV, DLYBIV, TOBIV, TOPGB, TCBIV
  End If
End If
End If
End Sub
Public Sub Data_Extra_B()
For jj = 1 To NNMAX
Write #1, TYPEN(jj)
Next
For jj = 1 To NPMP
 If CODPMP(jj) = "YES" Then
   Write #1, NPUMPCH(jj), HSHPMP(jj), CODNRR(jj)
   For ii = 1 To NPUMPCH(jj)
    Write #1, DCP(jj, ii), HEADP(jj, ii), EFFP(jj, ii)
   Next
 ElseIf CODPMP(jj) = "NO" Then
   Write #1, TYPCHP(jj)
 End If
Next
For jj = 1 To NBST
 If CODBST(jj) = "YES" Then
   Write #1, NBSTCH(jj), HSHBST(jj), CODNRB(jj)
   For ii = 1 To NBSTCH(jj)
    Write #1, DCB(jj, ii), HEADB(jj, ii), EFFB(jj, ii)
   Next
 ElseIf CODBST(jj) = "NO" Then
   Write #1, TYPCHB(jj)
 End If
Next
Write #1, DELV
For jj = 1 To 14
 Write #1, Iflag_But(jj)
Next
For jj = 1 To NSOU
Write #1, CODEVK(jj)
Next
For jj = 1 To NPMAX
  Write #1, PIPEMAT(jj)
Next
For jj = 1 To NPMAX
  Write #1, WTHICK(jj), CODEL(jj), LTHICK(jj), CODEG(jj), GTHICK(jj)
Next
For jj = 1 To NPMAX
  Write #1, CODECL(jj), CODEDIL(jj), THICKDIL(jj)
Next
For jj = 1 To NPMAX
  Write #1, CODECI(jj)
Next
For jj = 1 To NPMAX
  Write #1, CORET(jj), COATT(jj), STBYV(jj)
Next
For jj = 1 To NPMAX
  Write #1, THICKAC(jj)
Next
End Sub
Public Sub Save_Project()
Open SaveFile For Output As #1
 Write #1, PTYPE
  If PTYPE = "TYPEA" Then
   Fort_Data_A
   Extra_Data_A
  ElseIf PTYPE = "TYPEB" Then
  GetData_B
  Data_Fort_B1
  Data_Save_B2
  Data_Fort_B3
  Data_Fort_B4
  Data_Fort_B5
  Data_Fort_B6
  Data_Fort_B7
  Write_Device
  Data_Extra_B
 End If
 Close (1)
End Sub
Public Sub Fort_Data()
  If PTYPE = "TYPEA" Then
   Data_PUMPCH
   Format_Data_A
  ElseIf PTYPE = "TYPEB" Then
   GetData_B
  End If
  Open "c:\iisc\sap2.dat" For Output As #1
    Data_Fort_B1
    Data_Fort_B2
    Data_Fort_B3
    Data_Fort_B5
    Data_Fort_B7
    Write_Device
  Close (1)
  
  Open "c:\iisc\whwb2.dat" For Output As #1
    Data_Fort_B4
  Close (1)
  
  Open "c:\iisc\align2.dat" For Output As #1
    Data_Fort_B6
  Close (1)
  
End Sub


