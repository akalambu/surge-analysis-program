
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

' Format Project Type A Data for fortran Dump

Attribute VB_Name = "Module4"
Option Explicit

 Dim i As Integer
 Dim j As Integer
 
 Dim NPCH As Long
 Dim HSHUT As Single
 Dim QPMP As Single
 Dim HPMP As Single
 Dim EFF As Single
 Dim FKNR As Single
 Dim ISPCH As Long
 Dim QQ(1 To 100) As Single
 Dim HH(1 To 100) As Single
 Dim ETA(1 To 100) As Single
 Dim WH(1 To 89) As Single
 Dim WB(1 To 89) As Single

'Calculate Pump Characteristics Data
Public Sub Data_PUMPCH()
If CODPCH = "NO" Then
 If TYPCH = "RADIAL" Then
  ISPCH = 1
  FKNRRA = 0.7
 ElseIf TYPCH = "MIXED" Then
  ISPCH = 2
  FKNRRA = 2.2
 ElseIf TYPCH = "AXIAL" Then
  ISPCH = 3
  FKNRRA = 1.1
End If
Call PMPNO(ISPCH, WH(1), WB(1))
For i = 1 To 89
 WHP(1, i) = WH(i)
 WBP(1, i) = WB(i)
 WHP(2, i) = WH(i)
 WBP(2, i) = WB(i)
Next
ElseIf CODPCH = "YES" Then
  NPCH = NPUMPCH(1)
  QPMP = QR / NPUMP
  HPMP = REFH
  EFF = EFFA
  HSHUT = SHUOFF
 For i = 1 To NPUMPCH(1)
   QQ(i) = DCP(1, i)
   HH(i) = HEADP(1, i)
   ETA(i) = EFFP(1, i)
 Next
 NPCH = NPCH
 
 Call PMPYES(NPCH, QPMP, HPMP, EFF, HSHUT, QQ(1), HH(1), ETA(1), WH(1), WB(1), FKNR)
 
For i = 1 To 89
 WHP(1, i) = WH(i)
 WBP(1, i) = WB(i)
 WHP(2, i) = WH(i)
 WBP(2, i) = WB(i)
Next
FKNRRA = FKNR
End If
End Sub

' Format the TYPEA Data for Fortran Program

Public Sub Format_Data_A()
Dim TVEL As Single
GetData_A
If SIML = "PF" Then
NPMAX = 1
NNMAX = 2
IP(1) = 1
IND1(1) = 1
IND2(1) = 2
PDC(1) = QR
PDIA(1) = DIA
PLEN(1) = ALEN
WV(1) = WVA
CHST(1) = CHSTA
NNO(1) = 1
NTYPE(1) = 8
NUSP(1) = 0
NDSP(1) = 1
IDSP(1, 1) = 1
NNO(2) = 2
NTYPE(2) = 4
NUSP(2) = 1
NDSP(2) = 0
IUSP(2, 1) = 1
For i = 1 To NNMAX
 IRES(i) = 0
 ISOU(i) = 0
 ICDS(i) = 0
 IOBS(i) = 0
 IPMP(i) = 0
 IBST(i) = 0
 HGL(i) = 0
Next

IPMP(1) = 1
IRES(2) = 1

NORD = 0
NCJN = 0
NDJN = 0
NRES = 1
NSOU = 0
NCDS = 0
NOBS = 0
NPMP = 1
NBST = 0

RESWL(1) = DELL
RESDC(1) = QR
NPUSRS(1) = 1
KODEDS = 0
DLYDS = 0
TOPDS = 0

NPUMPS(1) = NPUMP
PUMPDC(1) = QR / NPUMP
QRPMP(1) = QR / NPUMP
PUMPH(1) = REFH
HRPMP(1) = REFH
EFFRP(1) = EFFA
PUMPSP(1) = ISPEED
CODPMP(1) = CODPCH
GD2PMP(1) = GDSQP
GD2PM(1) = GDSQM
SUMPWL(1) = DATUM
CODNRR(1) = CODNRRA
FKNRR(1) = FKNRRA
NPDSPM(1) = 1
ITRIPP(1) = 1

NPR1 = 1
NPR2 = 1
NPR3 = 1
ISPLOT = 1
NPATH = 1
NPPATH(1) = 1
IPPATH(1, 1) = 1
For i = 1 To NPRNT
 NPPRNT(i) = 1
Next

' Format Protection Devices
NLCS = 1
CS_Data(1, 0) = 1
NPAC = 1

For i = 0 To NOSTD - 1
 OST_Data(i, 0) = 1
Next
For i = 0 To NSSD - 1
 SP_Data(i, 0) = 1
Next
If CODESV = "YES" Then
 NLSRV = 1
 SRV_Data(1, 0) = 1
 SRV_Data(1, 1) = 0
Else
 NLSRV = 0
End If
For i = 0 To NZVD - 1
 ZV_Data(i, 0) = 1
Next
For i = 0 To NNRVD - 1
 INR_Data(i, 0) = 1
Next
For i = 0 To NAVD - 1
 AVV_Data(i, 0) = 1
Next
For i = 0 To NACVD - 1
 AC_Data(i, 0) = 1
Next
For i = 0 To NDPCVD - 1
 DPC_Data(i, 0) = 1
Next
CODIV = "NO"

ElseIf SIML = "APF" Or SIML = "SPF" Then
NPMAX = 3
NNMAX = 4

IP(1) = 1
IP(2) = 2
IP(3) = 3

IND1(1) = 1
IND2(1) = 2
IND1(2) = 2
IND2(2) = 3
IND1(3) = 4
IND2(3) = 2

PDC(1) = QR / NPUMP
PDIA(1) = DIAP
PLEN(1) = 10
WV(1) = WVP
CHST(1) = CHSTA - 10

PDC(2) = QR
PDIA(2) = DIA
PLEN(2) = ALEN
WV(2) = WVA
CHST(2) = CHSTA

PDC(3) = PDC(1) * (NPUMP - 1)
PDIA(3) = DIAP * Sqr(NPUMP - 1)
PLEN(3) = 10
WV(3) = WVP
CHST(3) = CHSTA - 10

NNO(1) = 1
NTYPE(1) = 8
NUSP(1) = 0
NDSP(1) = 1
IDSP(1, 1) = 1
NNO(2) = 2
NTYPE(2) = 2
NUSP(2) = 2
NDSP(2) = 1
IUSP(2, 1) = 1
IUSP(2, 2) = 3
IDSP(2, 1) = 2
NNO(3) = 3
NTYPE(3) = 4
NUSP(3) = 1
NDSP(3) = 0
IUSP(3, 1) = 2
NNO(4) = 4
NTYPE(4) = 8
NUSP(4) = 0
NDSP(4) = 1
IDSP(4, 1) = 3


For i = 1 To NNMAX
 IRES(i) = 0
 ISOU(i) = 0
 ICDS(i) = 0
 IOBS(i) = 0
 IPMP(i) = 0
 IBST(i) = 0
 HGL(i) = 0
Next
IPMP(1) = 1
TVEL = (QR / NPUMP) / ((3.14 * (DIAP / 1000) ^ 2) / 4#)
HGL(2) = DATUM + REFH - (TVEL ^ 2) / (2 * 9.81)
IRES(3) = 1
IPMP(4) = 2

NORD = 0
NCJN = 1
NDJN = 0
NRES = 1
NSOU = 0
NCDS = 0
NOBS = 0
NPMP = 2
NBST = 0

RESWL(1) = DELL
RESDC(1) = QR
NPUSRS(1) = 2
KODEDS = 0
DLYDS = 0
TOPDS = 0

NPUMPS(1) = 1
NPUMPS(2) = NPUMP - 1
PUMPDC(1) = QR / NPUMP
QRPMP(1) = QR / NPUMP
PUMPDC(2) = PUMPDC(1)
QRPMP(2) = PUMPDC(1)
PUMPH(1) = REFH
HRPMP(1) = REFH
PUMPH(2) = REFH
HRPMP(2) = REFH
EFFRP(1) = EFFA
EFFRP(2) = EFFA

PUMPSP(1) = ISPEED
CODPMP(1) = CODPCH
GD2PMP(1) = GDSQP
GD2PM(1) = GDSQM
SUMPWL(1) = DATUM
CODNRR(1) = CODNRRA
FKNRR(1) = FKNRRA

PUMPSP(2) = ISPEED
CODPMP(2) = CODPCH
GD2PMP(2) = GDSQP
GD2PM(2) = GDSQM
SUMPWL(2) = DATUM
CODNRR(2) = CODNRRA
FKNRR(2) = FKNRRA


NPDSPM(1) = 1
ITRIPP(1) = 1
NPDSPM(2) = 3
If SIML = "APF" Then
 ITRIPP(2) = 1
Else
 ITRIPP(2) = 0
End If


KODPHV(2) = KODPHV(1)
DLYPH(2) = DLYPH(1)
TCLOSEP(2) = TCLOSEP(1)
TRAPIDP(2) = TRAPIDP(1)
TSLOWP(2) = TSLOWP(1)

NPR1 = 2
NPR2 = 2
NPR3 = 2
ISPLOT = 1
NPATH = 1
NPPATH(1) = 1
IPPATH(1, 1) = 2
For i = 1 To NPRNT
 NPPRNT(i) = 2
Next

' Format Protection Devices
NLCS = 1
CS_Data(1, 0) = 2
NPAC = 2

For i = 0 To NOSTD - 1
 OST_Data(i, 0) = 2
Next
For i = 0 To NSSD - 1
 SP_Data(i, 0) = 2
Next
If CODESV = "YES" Then
 NLSRV = 1
 SRV_Data(1, 0) = 2
 SRV_Data(1, 1) = 0
Else
 NLSRV = 0
End If
For i = 0 To NZVD - 1
 ZV_Data(i, 0) = 2
Next
For i = 0 To NNRVD - 1
 INR_Data(i, 0) = 2
Next
For i = 0 To NAVD - 1
 AVV_Data(i, 0) = 2
Next
For i = 0 To NACVD - 1
 AC_Data(i, 0) = 2
Next
For i = 0 To NDPCVD - 1
 DPC_Data(i, 0) = 2
Next
CODIV = "NO"
End If
End Sub

'close a project and be ready to open a new one
Public Sub Reset_All()

If PTYPE = "TYPEA" Then
 NPMAX = 2
 NNMAX = 3
 NLCS = 1
 For i = 1 To 20
  NPUMPCH(i) = 0
 Next
End If

For i = 1 To 14
 Iflag_But(i) = 0
Next

For i = 1 To NPMAX
IP(i) = 0
IND1(i) = 0
IND2(i) = 0
PDC(i) = 0
PDIA(i) = 0
PLEN(i) = 0
CHST(i) = 0
PIPEMAT(i) = ""
WTHICK(i) = 0
CODEL(i) = ""
LTHICK(i) = 0
CODEG(i) = ""
GTHICK(i) = 0
CODECL(i) = ""
CODEDIL(i) = ""
THICKDIL(i) = 0
CODECI(i) = ""
CORET(i) = 0
COATT(i) = 0
STBYV(i) = 0
THICKAC(i) = 0
Next
For i = 1 To NNMAX
NNO(i) = 0
NTYPE(i) = 0
NUSP(i) = 0
NDSP(i) = 0
Next
For i = 1 To NNMAX
For j = 1 To 10
 IUSP(i, j) = 0
 IDSP(i, j) = 0
Next
Next
For i = 1 To NORD
 HGL(i) = 0
Next
For i = 1 To NCJN
 HGL(i) = 0
Next
For i = 1 To NDJN
 HGL(i) = 0
Next
For i = 1 To NRES
RESWL(i) = 0
RESDC(i) = 0
Next
If NRES = 1 Then
 KODEDS = 0
 DLYDS = 0
 TOPDS = 0
End If
For i = 1 To NSOU
 SOUWL(i) = 0
 SOUDC(i) = 0
Next
For i = 1 To NCDS
 HGL1CD(i) = 0
 HGL2CD(i) = 0
Next
For i = 1 To NOBS
HGL1OB(i) = 0
HGL2OB(i) = 0
Next
For i = 1 To NPMP
 NPUMPS(i) = 0
 PUMPDC(i) = 0
 QRPMP(i) = 0
 PUMPH(i) = 0
 HRPMP(i) = 0
 EFFRP(i) = 0
 PUMPSP(i) = 0
 CODPMP(i) = 0
 GD2PMP(i) = 0
 GD2PM(i) = 0
 SUMPWL(i) = 0
 KODPHV(i) = 0
 CODNRR(i) = ""
 FKNRR(i) = 0
 NPDSPM(i) = 0
 ITRIPP(i) = 0
Next
For i = 1 To NBST
 NBOOST(i) = 0
 BSTDC(i) = 0
 QRBST(i) = 0
 BSTH(i) = 0
 HRBST(i) = 0
 EFFRB(i) = 0
 BSTSP(i) = 0
 CODBST(i) = 0
 GD2BST(i) = 0
 GD2BM(i) = 0
 HGLSUC(i) = 0
 KODBSV(i) = 0
 CODNRB(i) = ""
 FKNRRB(i) = 0
 NPDSBS(i) = 0
 ITRIPB(i) = 0
Next
For i = 1 To NPMP
KODPHV(i) = 0
DLYPH(i) = 0
TCLOSEP(i) = 0
TRAPIDP(i) = 0
TSLOWP(i) = 0
Next
For i = 1 To NBST
KODBSV(i) = 0
DLYBS(i) = 0
TCLOSEB(i) = 0
TRAPIDB(i) = 0
TSLOWB(i) = 0
Next
NORD = 0
NCJN = 0
NDJN = 0
NRES = 0
NSOU = 0
NOBS = 0
NCDS = 0
NPMP = 0
NBST = 0

' for ptypea

KODPHV(1) = 0
DLYPH(1) = 0
TCLOSEP(1) = 0
TRAPIDP(1) = 0
TSLOWP(1) = 0

' Column Seperation Data

For i = 1 To NLCS
CS_Data(i, 0) = 0
CS_Data(i, 1) = 0
CS_Data(i, 2) = 0
Next
NLCS = 0

' Protection Devices Data

ISEL = 0
NOSTD = 0
NDPCVD = 0
NZVD = 0
NNRVD = 0
NAVD = 0
NACVD = 0
NSSD = 0

' Prit Data
NPRNT = 0

For i = 1 To 10
 For j = 1 To 50
   IPPATH(i, j) = 0
 Next
Next
For i = 1 To 10
 NPPATH(i) = 0
Next
For i = 1 To NPMAX
NALIGN(i) = 0
Next
NPATH = 0
NNMAX = 0
NPMAX = 0
End Sub

