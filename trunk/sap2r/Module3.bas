
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

Attribute VB_Name = "Module3"

'Read Data from the Saved Project

Public Sub Data_Read_B()
Input #1, PROJECT
Input #1, PCASE
Input #1, SIML
Input #1, NPMAX, NNMAX
For i = 1 To NPMAX
 Input #1, IP(i), IND1(i), IND2(i), PDC(i), PDIA(i), PLEN(i), WV(i), CHST(i)
Next
For i = 1 To NNMAX
  Input #1, NNO(i), NTYPE(i), NUSP(i), NDSP(i)
  For j = 1 To NUSP(i)
   Input #1, IUSP(i, j)
  Next
  For j = 1 To NDSP(i)
   Input #1, IDSP(i, j)
  Next
Next
For i = 1 To NNMAX
 Input #1, IRES(i), ISOU(i), ICDS(i), IOBS(i), IPMP(i), IBST(i), HGL(i)
Next
Input #1, NORD, NCJN, NDJN, NRES, NSOU, NCDS, NOBS, NPMP, NBST
For jj = 1 To NRES
 Input #1, RESWL(jj), RESDC(jj), NPUSRS(jj)
Next
If NRES = 1 Then
 Input #1, KODEDS, DLYDS, TOPDS
End If
For jj = 1 To NSOU
 Input #1, SOUWL(jj), SOUDC(jj), NPDSSO(jj)
Next
For jj = 1 To NCDS
 Input #1, HGL1CD(jj), HGL2CD(jj), QOCDS(jj), NPUSCD(jj)
Next
For jj = 1 To NOBS
 Input #1, HGL1OB(jj), HGL2OB(jj), QOOBS(jj), NPUSOB(jj)
Next
For jj = 1 To NPMP
 Input #1, NPUMPS(jj), PUMPDC(jj), QRPMP(jj), PUMPH(jj), HRPMP(jj), EFFRP(jj), PUMPSP(jj), CODPMP(jj), GD2PMP(jj), GD2PM(jj), SUMPWL(jj), KODPHV(jj), CODNRR(jj), FKNRR(jj), NPDSPM(jj), ITRIPP(jj)
Next
For jj = 1 To NBST
 Input #1, NBOOST(jj), BSTDC(jj), QRBST(jj), BSTH(jj), HRBST(jj), EFFRB(jj), BSTSP(jj), CODBST(jj), GD2BST(jj), GD2BM(jj), HGLSUC(jj), KODBSV(jj), CODNRB(jj), FKNRRB(jj), NPDSBS(jj), ITRIPB(jj)
Next
For jj = 1 To NPMP
For kk = 1 To 89
    Input #1, WHP(jj, kk), WBP(jj, kk)
  Next
Next
For jj = 1 To NBST
For kk = 1 To 89
    Input #1, WHB(jj, kk), WBB(jj, kk)
  Next
Next
For jj = 1 To NPMP
 If KODPHV(jj) = 1 Then
  Input #1, DLYPH(jj)
 ElseIf KODPHV(jj) = 2 Then
  Input #1, TCLOSEP(jj), DLYPH(jj)
 ElseIf KODPHV(jj) = 3 Then
  Input #1, TRAPIDP(jj), TSLOWP(jj), DLYPH(jj)
 ElseIf KODPHV(jj) = 4 Then
  Input #1, TCLOSEP(jj)
 ElseIf KODPHV(jj) = 5 Then
  Input #1, TRAPIDP(jj), TSLOWP(jj)
 End If
Next
For jj = 1 To NBST
 If KODBSV(jj) = 1 Then
  Input #1, DLYBS(jj)
 ElseIf KODBSV(jj) = 2 Then
  Input #1, TCLOSEB(jj), DLYBS(jj)
 ElseIf KODBSV(jj) = 3 Then
  Input #1, TRAPIDB(jj), TSLOWB(jj), DLYBS(jj)
 ElseIf KODBSV(jj) = 4 Then
  Input #1, TCLOSEB(jj)
 ElseIf KODBSV(jj) = 5 Then
  Input #1, TRAPIDB(jj), TSLOWB(jj)
 End If
Next

For ii = 1 To NPMAX
 Input #1, NALIGN(ii)
 For jj = 1 To NALIGN(ii)
  Input #1, CHAIN(ii, jj), GL(ii, jj)
 Next
Next

Input #1, NPR1, CHPR1, ZPR1
Input #1, NPR2, CHPR2, ZPR2
Input #1, NPR3, CHPR3, ZPR3
Input #1, ISPLOT
Input #1, NPATH
For jj = 1 To NPATH
 Input #1, NPPATH(jj)
For ii = 1 To NPPATH(jj)
 Input #1, IPPATH(jj, ii)
Next
Next
Input #1, NPRNT
For i = 1 To NPRNT
 Input #1, NPPRNT(i), CHPRNT(i)
Next
Input #1, CODSIM
If CODSIM = "YES" Then
Input #1, TMAX
End If
End Sub
Public Sub Read_Device()
Input #1, CODECS
If CODECS = "YES" Then
Input #1, NLCS
For jj = 1 To NLCS
Input #1, CS_Data(jj, 0), CS_Data(jj, 1), CS_Data(jj, 2)
Next
End If
Input #1, CODEPR
If CODEPR = "YES" Then
Input #1, CODEAC
If CODEAC = "YES" Then
Input #1, NPAC, GLAC, ACC, DCAC, ACNRV, KACTYP, DORBY
End If
Input #1, NOSTD
If NOSTD > 0 Then
For i = 0 To NOSTD - 1
 Input #1, OST_Data(i, 0), OST_Data(i, 1), OST_Data(i, 2), OST_Data(i, 3), OST_Data(i, 4), OST_Data(i, 5), OST_Data(i, 6)
Next
End If
Input #1, NZVD
If NZVD > 0 Then
For i = 0 To NZVD - 1
  Input #1, ZV_Data(i, 0), ZV_Data(i, 1), ZV_Data(i, 2)
Next
End If
Input #1, NDPCVD
If NDPCVD > 0 Then
For i = 0 To NDPCVD - 1
 Input #1, DPC_Data(i, 0), DPC_Data(i, 1), DPC_Data(i, 2)
Next
End If
Input #1, NNRVD
If NNRVD > 0 Then
For i = 0 To NNRVD - 1
  Input #1, INR_Data(i, 0), INR_Data(i, 1), INR_Data(i, 2), INR_Data(i, 3)
Next
End If
Input #1, NAVD
If NAVD > 0 Then
For i = 0 To NAVD - 1
  Input #1, AVV_Data(i, 0), AVV_Data(i, 1), AVV_Data(i, 2), AVV_Data(i, 3)
Next
End If
Input #1, NACVD
If NACVD > 0 Then
For i = 0 To NACVD - 1
Input #1, AC_Data(i, 0), AC_Data(i, 1), AC_Data(i, 2), AC_Data(i, 3)
Next
End If
Input #1, NSSD
If NSSD > 0 Then
For i = 0 To NSSD - 1
  Input #1, SP_Data(i, 0), SP_Data(i, 1), SP_Data(i, 2), SP_Data(i, 3), SP_Data(i, 4)
Next
End If
Input #1, NLSRV
If NLSRV > 0 Then
For jj = 1 To NLSRV
Input #1, SRV_Data(jj, 0), SRV_Data(jj, 1), SRV_Data(jj, 2), SRV_Data(jj, 3), SRV_Data(jj, 4), SRV_Data(jj, 5), SRV_Data(jj, 6), SRV_Data(jj, 7)
Next
End If
Input #1, CODIV
If CODIV = "YES" Then
  Input #1, NPIV, CHIV, DLYIV, TCIV
  Input #1, CODBIV
  If CODBIV = "YES" Then
    Input #1, HDELB, SZBIV, DLYBIV, TOBIV, TOPGB, TCBIV
  End If
End If
End If
End Sub
Public Sub Extra_Read_B()
For jj = 1 To NNMAX
Input #1, TYPEN(jj)
Next

For jj = 1 To NPMP
 If CODPMP(jj) = "YES" Then
   Input #1, NPUMPCH(jj), HSHPMP(jj), CODNRR(jj)
   For ii = 1 To NPUMPCH(jj)
    Input #1, DCP(jj, ii), HEADP(jj, ii), EFFP(jj, ii)
   Next
 ElseIf CODPMP(jj) = "NO" Then
   Input #1, TYPCHP(jj)
 End If
Next
For jj = 1 To NBST
 If CODBST(jj) = "YES" Then
   Input #1, NBSTCH(jj), HSHBST(jj), CODNRB(jj)
   For ii = 1 To NBSTCH(jj)
    Input #1, DCB(jj, ii), HEADB(jj, ii), EFFB(jj, ii)
   Next
 ElseIf CODBST(jj) = "NO" Then
   Input #1, TYPCHB(jj)
 End If
Next
Input #1, DELV
For jj = 1 To 14
 Input #1, Iflag_But(jj)
Next
For jj = 1 To NSOU
Input #1, CODEVK(jj)
Next
For jj = 1 To NPMAX
  Input #1, PIPEMAT(jj)
Next
For jj = 1 To NPMAX
  Input #1, WTHICK(jj), CODEL(jj), LTHICK(jj), CODEG(jj), GTHICK(jj)
Next
For jj = 1 To NPMAX
  Input #1, CODECL(jj), CODEDIL(jj), THICKDIL(jj)
Next
For jj = 1 To NPMAX
  Input #1, CODECI(jj)
Next
For jj = 1 To NPMAX
  Input #1, CORET(jj), COATT(jj), STBYV(jj)
Next
For jj = 1 To NPMAX
  Input #1, THICKAC(jj)
Next
End Sub

Public Sub SetData_B()

With frmTitle
 .txtName.Text = PROJECT
 .txtCase.Text = PCASE
End With
 
With frmTranMainB
  .txtNoPipe.Text = NPMAX
   If NPMAX > 0 And Iflag_But(1) = 1 Then
    .cmdDP.Caption = "Change"
   End If
  .txtNoNodes.Text = NNMAX
  If NNMAX > 0 And Iflag_But(2) = 1 Then
    .cmdDN.Caption = "Change"
   End If
End With

With frmOutPutB
.txtPN1.Text = NPR1
.txtPN2.Text = NPR2
.txtPN3.Text = NPR3
.txtCh1.Text = CHPR1
.txtCh2.Text = CHPR2
.txtCh3.Text = CHPR3
.txtRL1.Text = ZPR1
.txtRL2.Text = ZPR2
.txtRL3.Text = ZPR3
.cmbSimTime.Text = CODSIM
If .cmbSimTime.Text = "YES" Then
 .txtSimT.Text = TMAX
End If
End With

With frmHGL
For i = 0 To 8
If Iflag_But(i + 3) = 1 Then
.cmdHGL(i).Caption = "Change"
End If
Next
End With

With frmProtB
If CODEAC = "YES" Then
  .chkAirVes.Value = 1
  .chkSRV.Enabled = False
End If
If NOSTD > 0 Then
  .chkOneWay.Value = 1
End If
If NZVD > 0 Then
  .chkZeroV.Value = 1
  .chkDualPl.Enabled = False
End If
If NDPCVD > 0 Then
  .chkDualPl.Value = 1
  .chkZeroV.Enabled = False
End If
If NNRVD > 0 Then
  .chkInrv.Value = 1
End If
If NAVD > 0 Then
  .chkAirV.Value = 1
End If
If NACVD > 0 Then
  .chkACV.Value = 1
End If
If NSSD > 0 Then
  .chkSP.Value = 1
End If
If NLSRV > 0 Then
  .chkSRV.Value = 1
  .chkAirVes.Enabled = False
End If
If CODIV = "YES" Then
  .chkVBO.Value = 1
End If

End With
With frmOutPutB
    .txtPNo = ISPLOT
End With
With frmPath
  .txtNPath = NPATH
End With
End Sub
Public Sub Open_Project()
Open OpenFile For Input As #1
Input #1, PTYPE
 If PTYPE = "TYPEA" Then
  Read_Data_A
  Open_Project_A
 ElseIf PTYPE = "TYPEB" Then
  Data_Read_B
  Read_Device
  Extra_Read_B
  SetData_B
 End If
Close (1)
End Sub

