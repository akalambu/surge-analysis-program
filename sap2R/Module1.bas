
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

Attribute VB_Name = "Module1"
' Declaration, Reading and writing of Data for Project Type A

'MDIform1
Global PTYPE As String
Global OpenFile As String
Global SaveFile As String

'frmTitle

Global PROJECT As String
Global PCASE As String

'frmTranMainA

Global QR As Single
Global DIA As Single
Global ALEN As Single
Global CHSTA As Single
Global REFH As Single
Global DATUM As Single
Global DELL As Single
Global IDELP As String
Global DIAP As Single
Global WVP As Single
Global WVA As Single

'frmPump

Global EFFA As Single
Global ISPEED As Integer
Global GDSQP As Single
Global GDSQM As Single
Global CODNRRA As String
Global NPUMP As Integer
Global CODENP As String
Global CODEPM As String
'
'frmgridpumpch

Global CODPCH As String
Global SHUOFF As Single
Global FKNRRA As Single
Global TYPCH As String

Public Sub GetData_A()

'Title form data
 With frmTitle
  PROJECT = .txtName.Text
  PCASE = .txtCase.Text
 End With

'Transmission Main Data
With frmTranMainA
 QR = Val(.txtDisch.Text)
 DIA = Val(.txtDia.Text)
 ALEN = Val(.txtLeng.Text)
 CHSTA = Val(.txtChain.Text)
 REFH = Val(.txtHead.Text)
 DATUM = Val(.txtSumLev.Text)
 DELL = Val(.txtDelLev.Text)
 PIPEMAT(1) = .cmbPWV.Text

If .cmbPumpPipe.Text = "YES" Then
 IDELP = "YES"
 DIAP = Val(.txtPumpDelDia.Text)
 PIPEMAT(2) = .cmbPWVD.Text
Else
  IDELP = "NO"
End If
End With
 
'Pump Data

With frmPumpA
Dim bkw As Single
CODENP = .cmbPumps.Text
CODEPM = .cmbMach.Text
If .cmbPumps.Text = "YES" Then
NPUMP = Val(.txtPumpNo.Text)
Else
NPUMP = 1
End If
If .cmbMach.Text = "YES" Then
 EFFA = Val(.txtPumpEff.Text)
 ISPEED = Val(.txtPumpSp.Text)
 GDSQP = Val(.txtPumpGd2.Text)
 GDSQM = Val(.txtMotorGD2.Text)
 CODNRRA = .cmbRatchet.Text
Else
 EFFA = 85
 ISPEED = 1440
 bkw = (746# * (QR / NPUMP) * REFH * 1.2) / (75 * 0.85)
 GDSQP = (1# / 3#) * 540# * ((bkw / ISPEED) ^ 1.4)
 GDSQM = (2# / 3#) * 540# * ((bkw / ISPEED) ^ 1.4)
 CODNRRA = "NO"
End If
End With

'Analysis Details

With frmAnalyA
 If .optPowFail.Value = True Then
   SIML = "PF"
  ElseIf .optSingPump.Value = True Then
   SIML = "SPF"
  ElseIf .optAllPump.Value = True Then
   SIML = "APF"
  End If
  CODECS = .cmbColSep.Text
  CODEPR = .cmbAnProt.Text
 End With

'Output Details
 
 With frmOutPutA
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
 
 With frmPumpChA
  CODPCH = .cmbPumpCh.Text
  TYPCH = .cmbType.Text
  SHUOFF = Val(.txtShutOff.Text)
  End With
 End Sub

Public Sub Fort_Data_A()

GetData_A

Write #1, PROJECT
Write #1, PCASE
Write #1, QR, DIA, ALEN, CHSTA, REFH
Write #1, NPUMP, EFFA, ISPEED, GDSQP, GDSQM
Write #1, DIAP, WVP, DATUM, DELL, WVA

Write #1, CODNRRA
Write #1, KODPHV(1)
If KODPHV(1) = 1 Then
 Write #1, DLYPH(1)
ElseIf KODPHV(1) = 2 Then
 Write #1, TCLOSEP(1), DLYPH(1)
ElseIf KODPHV(1) = 3 Then
 Write #1, TRAPIDP(1), TSLOWP(1), DLYPH(1)
ElseIf KODPHV(1) = 4 Then
 Write #1, TCLOSEP(1)
ElseIf KODPHV(1) = 5 Then
 Write #1, TRAPIDP(1), TSLOWP(1)
ElseIf KODPHV(1) = 6 Then
End If
Write #1, SIML
Write #1, CODECS
If CODECS = "YES" Then
 Write #1, CS_Data(1, 1), CS_Data(1, 2)
End If
If CODEPR = "YES" Then
Write #1, CODEAC
If CODEAC = "YES" Then
Write #1, GLAC, ACC, DCAC, ACNRV, KACTYP, DORBY
End If
Write #1, NOSTD
If NOSTD > 0 Then
For i = 0 To NOSTD - 1
 Write #1, OST_Data(i, 1), OST_Data(i, 2), OST_Data(i, 3), OST_Data(i, 4), OST_Data(i, 5), OST_Data(i, 6)
Next
End If
Write #1, NZVD
If NZVD > 0 Then
For i = 0 To NZVD - 1
  Write #1, ZV_Data(i, 1), ZV_Data(i, 2)
Next
End If
Write #1, NDPCVD
If NDPCVD > 0 Then
For i = 0 To NDPCVD - 1
 Write #1, DPC_Data(i, 1), DPC_Data(i, 2)
Next
End If
Write #1, NNRVD
If NNRVD > 0 Then
For i = 0 To NNRVD - 1
  Write #1, INR_Data(i, 1), INR_Data(i, 2), INR_Data(i, 3)
Next
End If
Write #1, NAVD
If NAVD > 0 Then
For i = 0 To NAVD - 1
  Write #1, AVV_Data(i, 1), AVV_Data(i, 2), AVV_Data(i, 3)
Next
End If
Write #1, NACVD
If NACVD > 0 Then
For i = 0 To NACVD - 1
Write #1, AC_Data(i, 1), AC_Data(i, 2), AC_Data(i, 3)
Next
End If
Write #1, NSSD
If NSSD > 0 Then
For i = 0 To NSSD - 1
  Write #1, SP_Data(i, 1), SP_Data(i, 2), SP_Data(i, 3), SP_Data(i, 4)
Next
End If
Write #1, CODESV
If CODESV = "YES" Then
Write #1, SRV_Data(1, 2), SRV_Data(1, 3), SRV_Data(1, 4), SRV_Data(1, 5), SRV_Data(1, 6), SRV_Data(1, 7)
End If
 Else
 Write #1, "NO"
 Write #1, 0
 Write #1, 0
 Write #1, 0
 Write #1, 0
 Write #1, 0
 Write #1, 0
 Write #1, 0
 Write #1, "NO"
 End If
 Write #1, CODSIM
 If CODSIM = "YES" Then
 Write #1, TMAX
 End If
 Write #1, CHPR1, ZPR1
 Write #1, CHPR2, ZPR2
 Write #1, CHPR3, ZPR3
 Write #1, NPRNT
For i = 1 To NPRNT
 Write #1, CHPRNT(i)
Next
End Sub
Public Sub Extra_Data_A()
 Write #1, IDELP, CODENP, CODEPM
 Write #1, CODPCH
 If CODPCH = "NO" Then
   Write #1, TYPCH
 Else
   Write #1, SHUOFF
   Write #1, NPUMPCH(1)
   For i = 1 To NPUMPCH(1)
    Write #1, DCP(1, i), HEADP(1, i), EFFP(1, i)
   Next
 End If
 Write #1, NALIGN(1)
 For i = 1 To NALIGN(1)
 Write #1, CHAIN(1, i), GL(1, i)
 Next
For i = 1 To 2
Write #1, PIPEMAT(i)
Write #1, WTHICK(i), CODEL(i), LTHICK(i), CODEG(i), GTHICK(i)
Write #1, CODECL(i), CODEDIL(i), THICKDIL(i)
Write #1, CODECI(i)
Write #1, CORET(i), COATT(i), STBYV(i)
Write #1, THICKAC(i)
Next
End Sub
Public Sub Read_Data_A()
Input #1, PROJECT
Input #1, PCASE
Input #1, QR, DIA, ALEN, CHSTA, REFH
Input #1, NPUMP, EFFA, ISPEED, GDSQP, GDSQM
Input #1, DIAP, WVP, DATUM, DELL, WVA

Input #1, CODNRRA
Input #1, KODPHV(1)
If KODPHV(1) = 1 Then
Input #1, DLYPH(1)
ElseIf KODPHV(1) = 2 Then
Input #1, TCLOSEP(1), DLYPH(1)
ElseIf KODPHV(1) = 3 Then
Input #1, TRAPIDP(1), TSLOWP(1), DLYPH(1)
ElseIf KODPHV(1) = 4 Then
Input #1, TCLOSEP(1)
ElseIf KODPHV(1) = 5 Then
Input #1, TRAPIDP(1), TSLOWP(1)
ElseIf KODPHV(1) = 6 Then
End If
Input #1, SIML
Input #1, CODECS
If CODECS = "YES" Then
Input #1, CS_Data(1, 1), CS_Data(1, 2)
End If
Input #1, CODEAC
If CODEAC = "YES" Then
Input #1, GLAC, ACC, DCAC, ACNRV, KACTYP, DORBY
End If
Input #1, NOSTD
If NOSTD > 0 Then
For i = 0 To NOSTD - 1
 Input #1, OST_Data(i, 1), OST_Data(i, 2), OST_Data(i, 3), OST_Data(i, 4), OST_Data(i, 5), OST_Data(i, 6)
Next
End If
Input #1, NZVD
If NZVD > 0 Then
For i = 0 To NZVD - 1
  Input #1, ZV_Data(i, 1), ZV_Data(i, 2)
Next
End If
Input #1, NDPCVD
If NDPCVD > 0 Then
For i = 0 To NDPCVD - 1
 Input #1, DPC_Data(i, 1), DPC_Data(i, 2)
Next
End If
Input #1, NNRVD
If NNRVD > 0 Then
For i = 0 To NNRVD - 1
  Input #1, INR_Data(i, 1), INR_Data(i, 2), INR_Data(i, 3)
Next
End If
Input #1, NAVD
If NAVD > 0 Then
For i = 0 To NAVD - 1
  Input #1, AVV_Data(i, 1), AVV_Data(i, 2), AVV_Data(i, 3)
Next
End If
Input #1, NACVD
If NACVD > 0 Then
For i = 0 To NACVD - 1
Input #1, AC_Data(i, 1), AC_Data(i, 2), AC_Data(i, 3)
Next
End If
Input #1, NSSD
If NSSD > 0 Then
For i = 0 To NSSD - 1
  Input #1, SP_Data(i, 1), SP_Data(i, 2), SP_Data(i, 3), SP_Data(i, 4)
Next
End If
Input #1, CODESV
If CODESV = "YES" Then
Input #1, SRV_Data(1, 2), SRV_Data(1, 3), SRV_Data(1, 4), SRV_Data(1, 5), SRV_Data(1, 6), SRV_Data(1, 7)
End If
If (CODEAC = "YES" Or NOSTD > 0 Or NZVD > 0 Or NDPCVD > 0 Or NNRVD > 0 Or NAVD > 0 Or NACVD > 0 Or NSSD > 0 Or CODESV = "YES") Then
 CODEPR = "YES"
Else
 CODEPR = "NO"
End If
Input #1, CODSIM
If CODSIM = "YES" Then
Input #1, TMAX
End If
Input #1, CHPR1, ZPR1
Input #1, CHPR2, ZPR2
Input #1, CHPR3, ZPR3
Input #1, NPRNT
For i = 1 To NPRNT
 Input #1, CHPRNT(i)
Next
Input #1, IDELP, CODENP, CODEPM
Input #1, CODPCH
 If CODPCH = "NO" Then
  Input #1, TYPCH
 Else
   Input #1, SHUOFF
   Input #1, NPUMPCH(1)
   For i = 1 To NPUMPCH(1)
    Input #1, DCP(1, i), HEADP(1, i), EFFP(1, i)
   Next
End If
Input #1, NALIGN(1)
For i = 1 To NALIGN(1)
 Input #1, CHAIN(1, i), GL(1, i)
Next

For i = 1 To 2
Input #1, PIPEMAT(i)
Input #1, WTHICK(i), CODEL(i), LTHICK(i), CODEG(i), GTHICK(i)
Input #1, CODECL(i), CODEDIL(i), THICKDIL(i)
Input #1, CODECI(i)
Input #1, CORET(i), COATT(i), STBYV(i)
Input #1, THICKAC(i)
Next
End Sub
Public Sub Open_Project_A()
 With frmTitle
    .txtName = PROJECT
    .txtCase = PCASE
 End With

 With frmTranMainA
    .txtDisch = QR
    .txtDia = DIA
    .txtLeng = ALEN
    .txtChain = CHSTA
    .txtHead = REFH
    .txtSumLev.Text = DATUM
    .txtDelLev.Text = DELL
    .cmbPWV.Text = PIPEMAT(1)
  If IDELP = "YES" Then
    .txtPumpDelDia.Text = DIAP
    .cmbPWVD.Text = PIPEMAT(2)
  End If
 End With
 
 With frmPumpA
   If CODEPM = "YES" Then
     .txtPumpEff.Text = EFFA
     .txtPumpSp.Text = ISPEED
     .txtPumpGd2.Text = GDSQP
     .txtMotorGD2.Text = GDSQM
     .cmbRatchet.Text = CODNRRA
    End If
    If CODENP = "YES" Then
    .txtPumpNo.Text = NPUMP
    End If
 End With

With frmAnalyA
 If SIML = "PF" Then
 .optPowFail.Value = True
 ElseIf SIML = "SPF" Then
 .optSingPump.Value = True
 ElseIf SIML = "APF" Then
 .optAllPump.Value = True
 End If
End With

With frmProtA
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
End With

'Output Details

With frmOutPutA
.txtCh1.Text = CHPR1
.txtCh2.Text = CHPR2
.txtCh3.Text = CHPR3
.txtRL1.Text = ZPR1
.txtRL2.Text = ZPR2
.txtRL3.Text = ZPR3
If CODSIM = "YES" Then
.txtSimT.Text = TMAX
End If
End With

With frmPumpChA
 .cmbPumpCh = CODPCH
  If CODPCH = "NO" Then
   .cmbType = TYPCH
  Else
   .txtShutOff = SHUOFF
  End If
End With

End Sub

