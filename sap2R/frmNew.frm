
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
Begin VB.Form frmSalesDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Details"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5865
   ScaleWidth      =   7920
   Begin VB.PictureBox crReport 
      Height          =   480
      Left            =   5880
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   44
      Top             =   5280
      Width           =   1200
   End
   Begin VB.TextBox BillNo 
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   5880
      TabIndex        =   41
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox OrderDate 
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   5880
      TabIndex        =   39
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4680
      TabIndex        =   38
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdBill 
      Caption         =   "&Bill"
      Height          =   375
      Left            =   3600
      TabIndex        =   37
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   2520
      TabIndex        =   36
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   375
      Left            =   1440
      TabIndex        =   35
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox PaidAmount 
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4920
      TabIndex        =   34
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox TotalCost 
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1560
      TabIndex        =   33
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox SalesTaxRate 
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4920
      TabIndex        =   32
      Text            =   "000.00"
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox FreightCharges 
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1560
      TabIndex        =   31
      Text            =   "000.00"
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox UnitPrice 
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4920
      TabIndex        =   30
      Top             =   3260
      Width           =   1815
   End
   Begin VB.TextBox UnitSold 
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1560
      TabIndex        =   29
      Top             =   3260
      Width           =   1815
   End
   Begin VB.TextBox ProductID 
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4920
      TabIndex        =   28
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox CategoryID 
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1560
      TabIndex        =   27
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox CustomerID 
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1560
      TabIndex        =   26
      Top             =   360
      Width           =   1695
   End
   Begin VB.Frame frmChNo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   21
      Top             =   4680
      Width           =   3015
      Begin VB.TextBox ChequeNo 
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1200
         TabIndex        =   23
         Text            =   "00000"
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblNo 
         Caption         =   "No:"
         Height          =   375
         Left            =   720
         TabIndex        =   22
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.ComboBox PaymentMode 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1560
      TabIndex        =   20
      Text            =   "Cash"
      Top             =   4800
      Width           =   1815
   End
   Begin VB.TextBox CategoryName 
      ForeColor       =   &H00FF0000&
      Height          =   350
      Left            =   1560
      TabIndex        =   13
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox ProductName 
      ForeColor       =   &H00FF0000&
      Height          =   350
      Left            =   4920
      TabIndex        =   11
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox BillingAddress 
      ForeColor       =   &H00FF0000&
      Height          =   350
      Left            =   1560
      TabIndex        =   7
      Top             =   1560
      Width           =   3975
   End
   Begin VB.TextBox CompanyName 
      ForeColor       =   &H00FF0000&
      Height          =   350
      Left            =   1560
      TabIndex        =   5
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox CustomerName 
      ForeColor       =   &H00FF0000&
      Height          =   350
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   15
      Left            =   6000
      TabIndex        =   43
      Top             =   840
      Width           =   15
   End
   Begin VB.Label lblBil 
      Caption         =   "Bill No."
      Height          =   375
      Left            =   5160
      TabIndex        =   42
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   15
      Left            =   6000
      TabIndex        =   40
      Top             =   360
      Width           =   15
   End
   Begin VB.Label lblFrei 
      Caption         =   "Others"
      Height          =   495
      Left            =   360
      TabIndex        =   25
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label lblProd 
      Caption         =   "Product ID"
      Height          =   255
      Left            =   3840
      TabIndex        =   24
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblMode 
      Caption         =   "Payment Mode"
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label lblCost 
      Caption         =   "Total Cost"
      Height          =   495
      Left            =   360
      TabIndex        =   18
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label lblPaidAmout 
      Caption         =   "Paid Amount"
      Height          =   495
      Left            =   3720
      TabIndex        =   17
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lblTax 
      Caption         =   "Tax percentage "
      Height          =   495
      Left            =   3480
      TabIndex        =   16
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label lblQntity 
      Caption         =   "Quantity"
      Height          =   495
      Left            =   360
      TabIndex        =   15
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lblUnitPrice 
      Caption         =   "Unit Price"
      Height          =   495
      Left            =   3960
      TabIndex        =   14
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblCatagory 
      Caption         =   "Category"
      Height          =   495
      Left            =   360
      TabIndex        =   12
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblProductName 
      Caption         =   "Product Name"
      Height          =   495
      Left            =   3600
      TabIndex        =   10
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblProductID 
      Caption         =   "Category ID"
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblPurchase 
      Caption         =   "Sales Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label lblAddress 
      Caption         =   "Address"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblCompany 
      Caption         =   "Company"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblSupplierID 
      Caption         =   "Customer ID"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblDate 
      Caption         =   "Date"
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmSalesDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub PaymentMode_click()
If Not PaymentMode.Text = "Cash" Then
 frmChNo.Visible = True
Else
 frmChNo.Visible = False
End If
End Sub

