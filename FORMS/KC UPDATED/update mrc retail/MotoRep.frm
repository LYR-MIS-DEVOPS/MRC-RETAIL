VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MotoRep 
   BackColor       =   &H00FF8080&
   Caption         =   "==COMPANY NAME=="
   ClientHeight    =   9750
   ClientLeft      =   165
   ClientTop       =   -2625
   ClientWidth     =   16095
   Icon            =   "MotoRep.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cdlgVan 
      Left            =   240
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   9375
      Width           =   16095
      _ExtentX        =   28390
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "MotoRep.frx":17D2A
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "User Name"
            TextSave        =   "User Name"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Waiting..."
            TextSave        =   "Waiting..."
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "MotoRep.frx":1891C
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1940
            MinWidth        =   1940
            Text            =   "Transactions"
            TextSave        =   "Transactions"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14393
            Text            =   "Waiting..."
            TextSave        =   "Waiting..."
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "MotoRep.frx":191F6
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   882
            MinWidth        =   882
            Text            =   "Date"
            TextSave        =   "Date"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin ResizeXtra.RXPanel tbMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   16095
      _ExtentX        =   28390
      _ExtentY        =   741
      BorderSize      =   40
      BorderColor     =   -2147483644
      RegistryKey     =   "RXB49E"
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      BackColor       =   -2147483644
      Begin lvButton.lvButtons_H BtnMain 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "&Cardex"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   16777215
         cFHover         =   65535
         cBhover         =   16711680
         LockHover       =   2
         cGradient       =   16711680
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H BtnMain 
         Height          =   375
         Index           =   1
         Left            =   1935
         TabIndex        =   7
         Top             =   0
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "P&urchase Order"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   16777215
         cFHover         =   65535
         cBhover         =   16711680
         LockHover       =   2
         cGradient       =   16711680
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H BtnMain 
         Height          =   375
         Index           =   2
         Left            =   3750
         TabIndex        =   6
         Top             =   0
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "S&ales Order"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   16777215
         cFHover         =   65535
         cBhover         =   16711680
         LockHover       =   2
         cGradient       =   16711680
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H BtnMain 
         Height          =   375
         Index           =   3
         Left            =   5565
         TabIndex        =   5
         Top             =   0
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "&Ledger"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   16777215
         cFHover         =   65535
         cBhover         =   16711680
         LockHover       =   2
         cGradient       =   16711680
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H BtnMain 
         Height          =   375
         Index           =   4
         Left            =   7380
         TabIndex        =   4
         Top             =   0
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "&Daily Sales"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   16777215
         cFHover         =   65535
         cBhover         =   16711680
         LockHover       =   2
         cGradient       =   16711680
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H BtnMain 
         Height          =   375
         Index           =   5
         Left            =   9195
         TabIndex        =   3
         Top             =   0
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "&Mgr's Cashflow"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   16777215
         cFHover         =   65535
         cBhover         =   16711680
         LockHover       =   2
         cGradient       =   16711680
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Left            =   13845
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         Max             =   5
         Min             =   1
         Enabled         =   -1  'True
      End
      Begin VB.Label lblUser 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   10320
         TabIndex        =   9
         Top             =   60
         Width           =   1935
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         FillColor       =   &H00FF8080&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   11760
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.Menu mnuPI 
      Caption         =   "Personnel &Information"
      Begin VB.Menu mSubPI 
         Caption         =   "&Employees Profile"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mSubPI 
         Caption         =   "&User Profile"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu mSubPI 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mSubPI 
         Caption         =   "&On-Account Customer"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu mSubPI 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mSubPI 
         Caption         =   "&Supplier"
         Enabled         =   0   'False
         Index           =   6
      End
      Begin VB.Menu mSubPI 
         Caption         =   "S&hipper"
         Enabled         =   0   'False
         Index           =   7
      End
   End
   Begin VB.Menu mnuSP 
      Caption         =   "Spare &Parts"
      Begin VB.Menu mSubSP 
         Caption         =   "&Ordering"
         Enabled         =   0   'False
         Index           =   1
         Begin VB.Menu subOrder 
            Caption         =   "Purchase Order"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu subOrder 
            Caption         =   "Branch Purchase Order"
            Index           =   2
         End
         Begin VB.Menu subOrder 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu subOrder 
            Caption         =   "Summary of P.O."
            Enabled         =   0   'False
            Index           =   4
         End
         Begin VB.Menu subOrder 
            Caption         =   "Summary of Branch P.O."
            Index           =   5
         End
         Begin VB.Menu subOrder 
            Caption         =   "Back Order Summary"
            Index           =   6
         End
         Begin VB.Menu subOrder 
            Caption         =   "Top Selling Spare Parts"
            Index           =   7
         End
      End
      Begin VB.Menu mSubSP 
         Caption         =   "&Receiving"
         Enabled         =   0   'False
         Index           =   2
         Begin VB.Menu subReceive 
            Caption         =   "Receiving Form"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu subReceive 
            Caption         =   "Purchase Return"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu subReceive 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu subReceive 
            Caption         =   "Receiving Summary"
            Enabled         =   0   'False
            Index           =   4
         End
         Begin VB.Menu subReceive 
            Caption         =   "Purchase Return Summary"
            Enabled         =   0   'False
            Index           =   5
         End
      End
      Begin VB.Menu mSubSP 
         Caption         =   "&Sales"
         Enabled         =   0   'False
         Index           =   3
         Begin VB.Menu subSales 
            Caption         =   "Sales Order"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu subSales 
            Caption         =   "Warehouse"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu subSales 
            Caption         =   "Sales Return"
            Enabled         =   0   'False
            Index           =   3
         End
         Begin VB.Menu subSales 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu subSales 
            Caption         =   "Sales Order Summary"
            Enabled         =   0   'False
            Index           =   5
         End
         Begin VB.Menu subSales 
            Caption         =   "Sales Return Summary"
            Enabled         =   0   'False
            Index           =   6
         End
      End
      Begin VB.Menu mSubSP 
         Caption         =   "Stock &Monitoring"
         Enabled         =   0   'False
         Index           =   4
         Begin VB.Menu SPsubStocks 
            Caption         =   "CARDEX"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu SPsubStocks 
            Caption         =   "CARDEX [RED]"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu SPsubStocks 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu SPsubStocks 
            Caption         =   "Merchandise Inventory"
            Enabled         =   0   'False
            Index           =   4
         End
         Begin VB.Menu SPsubStocks 
            Caption         =   "Monthly Inventory Summary"
            Enabled         =   0   'False
            Index           =   5
         End
         Begin VB.Menu SPsubStocks 
            Caption         =   "Spare Parts Price List"
            Enabled         =   0   'False
            Index           =   6
         End
      End
      Begin VB.Menu mSubSP 
         Caption         =   "&Updates"
         Enabled         =   0   'False
         Index           =   5
         Begin VB.Menu SPsubADJ 
            Caption         =   "Update Location"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu SPsubADJ 
            Caption         =   "Update Spare Parts"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu SPsubADJ 
            Caption         =   "Price Adjustment"
            Enabled         =   0   'False
            Index           =   3
         End
      End
   End
   Begin VB.Menu mnuSrvc 
      Caption         =   "&Service Center"
      Begin VB.Menu SubSRVC 
         Caption         =   "&Job Orders"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu SubSRVC 
         Caption         =   "&Requisition"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu SubSRVC 
         Caption         =   "Customer &Purchase Order"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu SubSRVC 
         Caption         =   "&Customer Service Record"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu SubSRVC 
         Caption         =   "&Mechanic Service Income"
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu SubSRVC 
         Caption         =   "Service Income &Summary"
         Enabled         =   0   'False
         Index           =   6
      End
   End
   Begin VB.Menu mnuTRN 
      Caption         =   "&Transactions"
      Begin VB.Menu mSubTR 
         Caption         =   "Sales &Invoice"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mSubTR 
         Caption         =   "&Official Receipt"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu mSubTR 
         Caption         =   "&Provisional Receipt"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu mSubTR 
         Caption         =   "&Adjustments"
         Enabled         =   0   'False
         Index           =   4
         Begin VB.Menu subAdj 
            Caption         =   "A/R DMCM"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu subAdj 
            Caption         =   "Inventory Adjustment"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu subAdj 
            Caption         =   "Clerical Adjustment"
            Enabled         =   0   'False
            Index           =   3
         End
         Begin VB.Menu subAdj 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu subAdj 
            Caption         =   "A/R DMCM Summary"
            Enabled         =   0   'False
            Index           =   5
         End
         Begin VB.Menu subAdj 
            Caption         =   "Inv. Adj. Summary"
            Enabled         =   0   'False
            Index           =   6
         End
         Begin VB.Menu subAdj 
            Caption         =   "Clerical Adj. Summary"
            Enabled         =   0   'False
            Index           =   7
         End
      End
      Begin VB.Menu mSubTR 
         Caption         =   "&Disbursement"
         Enabled         =   0   'False
         Index           =   5
         Begin VB.Menu subDisb 
            Caption         =   "Petty Cash Voucher"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu subDisb 
            Caption         =   "Cash Voucher"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu subDisb 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu subDisb 
            Caption         =   "PCV Summary"
            Enabled         =   0   'False
            Index           =   4
         End
         Begin VB.Menu subDisb 
            Caption         =   "Cash Voucher Summary"
            Enabled         =   0   'False
            Index           =   5
         End
      End
      Begin VB.Menu mSubTR 
         Caption         =   "&Cash Flow"
         Enabled         =   0   'False
         Index           =   6
         Begin VB.Menu subCashFlow 
            Caption         =   "Cashier's Cash Flow"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu subCashFlow 
            Caption         =   "Manager's Cash Flow"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu subCashFlow 
            Caption         =   "Advances for Liquidation"
            Enabled         =   0   'False
            Index           =   3
         End
         Begin VB.Menu subCashFlow 
            Caption         =   "Cash Return"
            Enabled         =   0   'False
            Index           =   4
         End
         Begin VB.Menu subCashFlow 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu subCashFlow 
            Caption         =   "Checks Summary"
            Enabled         =   0   'False
            Index           =   6
         End
      End
      Begin VB.Menu mSubTR 
         Caption         =   "Accounts &Receivable"
         Enabled         =   0   'False
         Index           =   7
         Begin VB.Menu subAR 
            Caption         =   "Aging of Accounts Receivables"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu subAR 
            Caption         =   "Statement of Accounts"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu subAR 
            Caption         =   "Ledger"
            Enabled         =   0   'False
            Index           =   3
         End
         Begin VB.Menu subAR 
            Caption         =   "Down Payment Summary"
            Enabled         =   0   'False
            Index           =   4
         End
         Begin VB.Menu subAR 
            Caption         =   "Cash Quick Service w/ Credit Invoice"
            Enabled         =   0   'False
            Index           =   5
         End
         Begin VB.Menu subAR 
            Caption         =   "Aging (New Version)"
            Enabled         =   0   'False
            Index           =   6
         End
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mSubRep 
         Caption         =   "Daily Transactions"
         Enabled         =   0   'False
         Index           =   1
         Begin VB.Menu subRptDailySales 
            Caption         =   "Daily Sales Report"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu subRptDailySales 
            Caption         =   "Daily Cash Collection"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu subRptDailySales 
            Caption         =   "Daily Disbursement Report"
            Enabled         =   0   'False
            Index           =   3
         End
      End
      Begin VB.Menu mSubRep 
         Caption         =   "Spare Parts Inventory"
         Enabled         =   0   'False
         Index           =   2
         Begin VB.Menu subRptSP 
            Caption         =   "Merchandise Inventory"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu subRptSP 
            Caption         =   "Detailed Receiving Report"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu subRptSP 
            Caption         =   "Detailed Sales Report At Selling Price"
            Enabled         =   0   'False
            Index           =   3
         End
         Begin VB.Menu subRptSP 
            Caption         =   "Summary of Inv. Adjustment"
            Enabled         =   0   'False
            Index           =   4
         End
         Begin VB.Menu subRptSP 
            Caption         =   "Summary of Sales Return"
            Enabled         =   0   'False
            Index           =   5
         End
         Begin VB.Menu subRptSP 
            Caption         =   "Summary of Purchase Return"
            Enabled         =   0   'False
            Index           =   6
         End
         Begin VB.Menu subRptSP 
            Caption         =   "Detailed Sales Report At Cost Price"
            Enabled         =   0   'False
            Index           =   7
         End
      End
      Begin VB.Menu mSubRep 
         Caption         =   "Service Center"
         Enabled         =   0   'False
         Index           =   3
         Begin VB.Menu subRptSrv 
            Caption         =   "Daily Service Income"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu subRptSrv 
            Caption         =   "Mechanic's Service Income"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu subRptSrv 
            Caption         =   "Service Income Summary"
            Enabled         =   0   'False
            Index           =   3
         End
      End
      Begin VB.Menu mSubRep 
         Caption         =   "Monthly Reports"
         Enabled         =   0   'False
         Index           =   4
         Begin VB.Menu SubRptMonthly 
            Caption         =   "Monthly Sales Report"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu SubRptMonthly 
            Caption         =   "Monthly Collection Report"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu SubRptMonthly 
            Caption         =   "Monthly Disbursement Report"
            Enabled         =   0   'False
            Index           =   3
         End
         Begin VB.Menu SubRptMonthly 
            Caption         =   "Monthly Service Report"
            Enabled         =   0   'False
            Index           =   4
         End
         Begin VB.Menu SubRptMonthly 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu SubRptMonthly 
            Caption         =   "Sales Distribution by Parts"
            Enabled         =   0   'False
            Index           =   6
         End
         Begin VB.Menu SubRptMonthly 
            Caption         =   "Sales Database Report"
            Enabled         =   0   'False
            Index           =   7
         End
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "T&ools"
      Begin VB.Menu mSubTools 
         Caption         =   "Backup"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mSubTools 
         Caption         =   "Restore"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu mSubTools 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mSubTools 
         Caption         =   "Audit Tools"
         Enabled         =   0   'False
         Index           =   4
         Begin VB.Menu AudTools 
            Caption         =   "Inventory Auditing"
            Enabled         =   0   'False
            Index           =   1
         End
      End
      Begin VB.Menu mSubTools 
         Caption         =   "Settings"
         Enabled         =   0   'False
         Index           =   5
         Begin VB.Menu mSetPrn 
            Caption         =   "Printing Preference"
            Enabled         =   0   'False
         End
         Begin VB.Menu mSetOpt 
            Caption         =   "Options"
            Enabled         =   0   'False
         End
      End
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "&Quit"
   End
End
Attribute VB_Name = "MotoRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'maayong buntag dha kanimo
'salamat sa pag uban kanamo karng adlawa
'sa among pagsagud sa trabaho nga imong gi salig kanamo,
'among pangayaon Lord nga tagaan ni nmo ug igong kusog nga mabuhat namo among mga trabaho nga malipayon
'tabangi mi Lord nga mahuman namo among mga trabaho, ug nag hinaot mi nga kini mka palipay kanimo.
'pasaylua mi sa among mga sala Lord.

'**********************************
'Lord maayong adlaw diha kanimo sa kanunay ikaw nagauban kanamo
'kami nagapasalamat sa gasa sa kinabuhi na walay sawang gihatag nimo kanamo,
'salamat usab lord sa mga tao na nahumong instrumento na nagtabang sa usag usa kanamo
'na sila permi ubanan nimo ug dan-agan mo sa kahayag sa imong hulagway na sila magpadayon sa ilang pagtuman
'sa imong kasuguan.

Private Sub AudTools_Click(Index As Integer)
    Select Case Index
        Case 1
            AuditInvntry.Show
            tbMain.Visible = False
    End Select
End Sub

Private Sub BtnMain_Click(Index As Integer)
    Select Case Index
        Case 0
            SPCardex.Show vbModal
        Case 1
            SPordering.Show
            tbMain.Visible = False
        Case 2
            SPSorder.Show
            tbMain.Visible = False
        Case 3
            Ledger.Show
            tbMain.Visible = False
        Case 4
            SPSalesRpt.Show
            tbMain.Visible = False
        Case 5
            CashFlowMgr.Show
            tbMain.Visible = False
    End Select
End Sub

Private Sub MDIForm_Load()
    'sBar.Panels(9).Text = Format(Now, "Mmm. dd, yyyy")
    sBar.Panels(9).Text = ServerTime
    MotoRep.Caption = SysBranch & " v15.8.03"
    'v14.3.13 = Fix DP & A/R DMCM w/ Inv. Type(Credit&Cash) error, Parts Desc. Upto 150 chars., Disbursement Refund allow CR
    'V14.6.30 = NEW DSN + LOCAL     use in Audit dept. for running previous backup
    'V14.7.15 = Allocate check to different customer. 2 Printing Preference New OR and Receipt controlled in table settings
    'v14.7.24 = Allocate check to Job Order and Payment can use Customer Name with check allocation even if Check customer name is different. PO from Branch can view Gensan
    'V15.1.2  = TRA for Van Selling, Import Stock Transfer to VAN, Recall of last Date Covered
    'v15.2.17 = fix VAN Sales Return & MRC2 statement of accts, VAN3 on Receipts check allocations
    'v15.3.12 = Branch PO is now Warehouse ready
    'v15.3.22 = fIX BUG On Online Saving og PO, Use table sprecdiscrepancy(REMARKS)
    'V15.8.03 = WAREHOUSE BACK ORDER AND FREEZE COST IN EVERY TRANSACTION CONNECTING TO SPITR
    If InStr(SysBranch, "VAN") >= 1 Then
        mSubTR(1).Caption = "Trust Receipt"
        If InStr(SysBranch, "MRC VAN") >= 1 Then
            mSubTools(2).Caption = "Restore Mobile DB"
        End If
    End If
    FrDate = DateValue("1/1/1977")
    ToDate = DateValue("1/1/1977")
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Ans As Integer, i As Integer
    Dim MyForms As String
    
    MyForms = ""
    
    For i = 0 To Forms.Count - 1
        MyForms = Forms(i).Caption
    Next i
    If MyForms <> SysBranch & " v15.3.22" Then
        MsgBox "Please Close the following forms:" & vbCrLf & MyForms, vbExclamation + vbOKOnly, "Warning"
        Cancel = True
        Exit Sub
    End If
    Ans = MsgBox("Are you sure you want to quit?", vbQuestion + vbYesNo, "Quit")
    If Ans = vbYes Then
        'Call AstaLaVista
        LYRcon.Close
        If conVAN1.State = adStateOpen Then conVAN1.Close
        If conVAN2.State = adStateOpen Then conVAN2.Close
        If conMRC1.State = adStateOpen Then conMRC1.Close
    Else
        Cancel = True
    End If

End Sub

Private Sub MDIForm_Resize()
    Dim lWidth As Long
    Dim BtnSize As Integer, i As Integer

    If Me.WindowState = 1 Then Exit Sub
    lWidth = Screen.Width \ Screen.TwipsPerPixelX
    If lWidth > 850 And Me.Width > 14000 Then
        BtnSize = 1800
    Else
        BtnSize = 1455
    End If
    For i = 0 To 5
        BtnMain(i).Width = BtnSize
        BtnMain(i).Left = 120 + (BtnMain(i).Width * i)
    Next i
    
    UpDown1.Left = Me.ScaleWidth - 300
    Shape1.Left = Me.ScaleWidth - 2450
    lblUser.Left = Me.ScaleWidth - 2400
End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub

Private Sub mSetOpt_Click()
    SEToptions.Show vbModal
End Sub

Private Sub mSetPrn_Click()
    TLSprintPref.Show
    MotoRep.tbMain.Visible = False
End Sub

Private Sub mSubPI_Click(Index As Integer)
    Select Case Index
        Case 1
            Employees.Show
            MotoRep.tbMain.Visible = False
        Case 2  'User
            EMPuser.Show vbModal
            'MotoRep.tbMain.Enabled = False
        Case 4  'Branch
            AcctCustomer.Show vbModal
        Case 6  'Supplier
            EMPsupplier.Show vbModal
        Case 7  'Shipper
            EMPshipper.Show vbModal
    End Select
End Sub

Private Sub mSubTools_Click(Index As Integer)
    Dim myFile As String, myFile2 As String
    Dim Qry As String, BranchDB As String
    
    On Error GoTo LocalErr
    Select Case Index
        
        Case 1  'Backup
            myFile = Dir("C:\MIS\BACKUP\MYSQLDUMP.EXE")
            myFile2 = Dir("C:\MIS\BACKUP\RAR.EXE")
            If UCase$(myFile) <> "MYSQLDUMP.EXE" Or UCase$(myFile2) <> "RAR.EXE" Then
                MsgBox "Backup File Systems missing, Please consult IT Support Group.", vbExclamation + vbOKOnly, "Warning"
                Exit Sub
            End If
            Qry = "SELECT settings.Value From settings Where settings.DESCRIPTION = 'BACKUP' AND settings.REMARKS = 'FILENAME'"
            myFile = GetFieldValue(Qry)
            With cdlgVan
                .CancelError = True
                .Filter = "Backup Files (*.RAR)|(*.RAR)"
                .InitDir = App.Path
                myFile = myFile & Format(Now, "mmddyyhhmm")
                .filename = myFile & ".RAR"
                .ShowSave
                myFile2 = .filename
            End With
            If UCase$(SysBranch) = "MOBILE VAN-2 [LHC403]" Then
                BranchDB = "van1mobile"
            ElseIf UCase$(SysBranch) = "MOBILE VAN-56 [AAH1233]" Then
                BranchDB = "van2mobile"
            ElseIf UCase$(SysBranch) = "MOBILE VAN-55 [AAH1250]" Then
                BranchDB = "van3mobile"
            ElseIf UCase$(SysBranch) = "MRC 2 WHOLESALE" Then
                BranchDB = "sqlmrc2tgm"
            End If
            
            Open "Backup.bat" For Output As #1
            Print #1, "@echo off"
            Print #1, "echo Processing your BACKUP for " & myFile & ", Please wait....."
            Print #1, "C:\MIS\BACKUP\mysqldump -hlocalhost -uroot -plaborday " & BranchDB & " > C:\MIS\BACKUP\" & myFile & ".SQL"
            Print #1, "echo This may take several minutes... Please wait...."
            Print #1, "C:\MIS\BACKUP\rar a -plaborday " & "C:\MIS\BACKUP\" & myFile & ".rar C:\MIS\BACKUP\" & myFile & ".SQL"
            Print #1, "COPY C:\MIS\BACKUP\" & myFile & ".RAR " & myFile2
            Print #1, "echo BACKUP is now finished"
            Print #1, "del C:\MIS\BACKUP\" & myFile & ".SQL"
            Print #1, "del Backup.bat"
            Close #1
            Shell "Backup.bat", vbNormalNoFocus
        
        Case 2  'Restore
        
            myFile = Dir("C:\MIS\BACKUP\MYSQL.EXE")
            myFile2 = Dir("C:\MIS\BACKUP\RAR.EXE")
            If UCase$(myFile) <> "MYSQL.EXE" Or UCase$(myFile2) <> "RAR.EXE" Then
                MsgBox "Restore File System missing, Please consult IT Support Group.", vbExclamation + vbOKOnly, "Warning"
                Exit Sub
            End If
            If UCase$(SysBranch) <> "MRC VAN-2 [LHC403]" And UCase$(SysBranch) <> "MRC VAN-56 [AAH1233]" And UCase$(SysBranch) <> "MRC VAN-55 [AAH1250]" Then
                MsgBox "Sorry, Restore is not applicable to current system branch", vbExclamation + vbOKOnly, "Warning"
                Exit Sub
            End If
            With cdlgVan
                .filename = ""
                .CancelError = True
                .Filter = "VAN Backup (*.rar)|*.rar"
                .InitDir = App.Path
                .ShowOpen
                myFile2 = .filename
            End With
            
            If UCase$(SysBranch) = "MRC VAN-2 [LHC403]" Then
                myFile = "LHC403VAN"
                BranchDB = "van1mobile"
                
            ElseIf UCase$(SysBranch) = "MRC VAN-56 [AAH1233]" Then
                myFile = "AAH1233VAN"
                BranchDB = "van2mobile"
                
            ElseIf UCase$(SysBranch) = "MRC VAN-55 [AAH1250]" Then
                myFile = "AAH1250VAN"
                BranchDB = "van3mobile"
            End If
            If InStr(myFile2, myFile) = 0 Then
                MsgBox "Incompatible Backup file! please select " & myFile & " format!", vbExclamation + vbOKOnly, "Warning"
                Exit Sub
            End If
            Open "Restore.bat" For Output As #1
            Print #1, "@echo off"
            Print #1, "echo Restoring your BACKUP for " & myFile & ", Please wait....."
            Print #1, "C:\MIS\BACKUP\rar e -plaborday " & myFile2 '& " " & <FILE PATH>
            Print #1, "echo Please wait...Updating server databases"
            Print #1, "C:\MIS\BACKUP\mysql -uroot -plaborday " & BranchDB & " < " & Left(myFile2, Len(myFile2) - 3) & "SQL"
            Print #1, "echo Data Successfully loaded"
            Print #1, "pause"
            Print #1, "del Restore.bat"
            Close #1
            Shell "Restore.bat", vbNormalNoFocus
        
        Case 5  'Settings
            'Settings.Show vbModal
    End Select
    Exit Sub
LocalErr:
    If Err.Number <> 32755 Then
        MsgBox ErrMsg & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Error " & Err.Number
    End If
End Sub

Private Sub mSubTR_Click(Index As Integer)
    Select Case Index
        Case 1
            If InStr(SysBranch, "VAN") >= 1 Then
                TRA.Show
            Else
                Invoice.Show
            End If
            MotoRep.tbMain.Visible = False
        Case 2
            Receipt.Show 'vbModal
            MotoRep.tbMain.Visible = False
        Case 3
            arPR.Show
            MotoRep.tbMain.Visible = False
    End Select
End Sub

Private Sub SPsubADJ_Click(Index As Integer)
    Select Case Index
        Case 1  'SP location adjustment
            SPLocAdj.Show
            MotoRep.tbMain.Visible = False
        Case 2
            SPupdate.Show
            MotoRep.tbMain.Visible = False
        Case 3  'Pricing
            SPpricing.Show
            MotoRep.tbMain.Visible = False
    End Select
End Sub

Private Sub SPsubStocks_Click(Index As Integer)
    Dim Qry As String, MyPos As String
    
    Select Case Index
        Case 1  'Cardex
            SPCardex.Show vbModal
        Case 2  'Cardex Red
            SPCardexRed.Show vbModal
        Case 4  'Merchandise Inventory
            Qry = "SELECT employees.`POSITION` From Employees Where Employees.EmpId = " & UserID & ""
            MyPos = GetFieldValue(Qry)
            If UserID = 2 Or UserID = 4 Or UserID = 6 Or UserID = 25 Or UCase$(MyPos) = "MANAGER" Or UCase$(MyPos) = "MANAGER" Then
                SPMerInvCost.Show
            Else
                SPrptINV.Show
            End If
            MotoRep.tbMain.Visible = False
        Case 5  'Inventory Monitoring Report
            SPInventory.Show
            MotoRep.tbMain.Visible = False
        Case 6  'Price List
            PRNmc.MySource = "PRICELIST"
            PRNmc.Show
            MotoRep.tbMain.Visible = False
    End Select
End Sub


Private Sub subAdj_Click(Index As Integer)

    Select Case Index
        Case 1  'AR DMCM
            DMCMadj.Show
            MotoRep.tbMain.Visible = False
        Case 2
            InvntryADJ.Show
            MotoRep.tbMain.Visible = False
        Case 3
            ClericalADJ.Show vbModal
        Case 5  'DMCM Summary
            'DmCmSumm.Show
            FrmSummary.mySrc = "DMCM [A/R]"
            FrmSummary.Show
            MotoRep.tbMain.Visible = False
        Case 6
            FrmSummary.mySrc = "Inventory Adjustment"
            FrmSummary.Show
            'InvAdjSumm.Show
            MotoRep.tbMain.Visible = False
        Case 7
            FrmSummary.mySrc = "Clerical Adjustment"
            FrmSummary.Show
            'ClrAdjSumm.Show
            MotoRep.tbMain.Visible = False
    End Select
End Sub

Private Sub subAR_Click(Index As Integer)
    Select Case Index
        Case 1
            'SPrptAR.Show
            AgingNew.Show
        Case 2
            If SysBranch = "MRC 2 WHOLESALE" Then
                BillingNew.Show
            Else
                BillingProcess.Show
            End If
        Case 3
            Ledger.Show
        Case 4
            DPSummary.Show
        Case 5
            QuickSrvInvoice.Show
        Case 6
            AgingNew.Show
    End Select
    MotoRep.tbMain.Visible = False
    
End Sub

Private Sub subCashFlow_Click(Index As Integer)
    Select Case Index
        Case 1  'Cashier Cash Flow
            CashFlowCashier.Show
            MotoRep.tbMain.Visible = False
        Case 2  'Mgr Cash flow
            CashFlowMgr.Show
            MotoRep.tbMain.Visible = False
        Case 3
            ExpAdvLiq.Show vbModal
        Case 4
            ExpCashRet.Show vbModal
        Case 6  'Check Summary
            arCheckSumm.Show
            MotoRep.tbMain.Visible = False
    End Select
End Sub

Private Sub subDisb_Click(Index As Integer)
    Select Case Index
        Case 1  'PCV
            Disbursement.Show vbModal
        Case 2  'CV
            CashVoucher.Show vbModal
        Case 4  'PCV summary
            DisburseSumm.Show
            MotoRep.tbMain.Visible = False
        Case 5  'CV summary
            'CVsummary.ViewList = True
            'CVsummary.Show
            FrmSummary.mySrc = "Cash Voucher"
            FrmSummary.Show
            MotoRep.tbMain.Visible = False
    End Select
End Sub
Private Sub subOrder_Click(Index As Integer)
    Select Case Index
    
        'Old------------------------------
        'Case 1  'PO
        '    SPordering.Show
        '    MotoRep.tbMain.Visible = False
        'Case 2
            'separator
        'Case 3
            'Summary
        '    SPorderSum.Show
        '    MotoRep.tbMain.Visible = False
        
        'Case 4 'Back Order
        '    SPbckOrdSum.Show
        '    MotoRep.tbMain.Visible = False
    
        'Case 5 'Top Selling spare parts
        '    SPTopSales.Show
        '    MotoRep.tbMain.Visible = False
        
        
        'New-------------------------------
        Case 1  'PO
            SPordering.Show
            MotoRep.tbMain.Visible = False
        Case 2
            'Branch P.O
            'BranchPO.Show
            'MotoRep.tbMain.Visible = False
            '3 - Separator
        Case 4
            'Summary
            SPorderSum.Show
            MotoRep.tbMain.Visible = False
        Case 5
            'Summary Branch PO
            FrmSummary.mySrc = "Branch Purchase Order"
            FrmSummary.Show
            MotoRep.tbMain.Visible = False
            
        Case 6 'Back Order
'            SPbckOrdSum.Show
'            MotoRep.tbMain.Visible = False
            frmHSummary.MySource = "MRC WAREHOUSE BACK ORDER"
            frmHSummary.Show
            MotoRep.tbMain.Visible = False
        Case 7 'Top Selling spare parts
            SPTopSales.Show
            MotoRep.tbMain.Visible = False
    End Select
End Sub

Private Sub subReceive_Click(Index As Integer)
    Select Case Index
        Case 1  'Receive SP
            SPincoming.Show
            MotoRep.tbMain.Visible = False
        Case 2  'Purchase Return
            PurchaseReturn.Show
            MotoRep.tbMain.Visible = False
        Case 4  'Receive Summary
            FrmSummary.mySrc = "Received Spare Parts"
            FrmSummary.Show
            'SPRecSummary.Show
            MotoRep.tbMain.Visible = False
        Case 5  'Receive Summary
            'PurchRetSum.Show
            FrmSummary.mySrc = "Purchase Return"
            FrmSummary.Show
            MotoRep.tbMain.Visible = False
    End Select
End Sub

Private Sub subRptDailySales_Click(Index As Integer)
    Select Case Index
        Case 1 'Daily Sales Report
            SPSalesRpt.Show
            MotoRep.tbMain.Visible = False
        Case 2  'Daily Cash Collection
            DailyCash.Show
            MotoRep.tbMain.Visible = False
        Case 6  'Inventory
            SPInventory.Show
            MotoRep.tbMain.Visible = False
    End Select
End Sub

Private Sub SubRptMonthly_Click(Index As Integer)
    Select Case Index
        Case 1  'Sales
            MonthlySales.Show
            MotoRep.tbMain.Visible = False
        Case 2
            MonthlyColl.Show
            MotoRep.tbMain.Visible = False
        Case 6
            'Sales Distribution by Parts
            SalesDistribution.Show
            MotoRep.tbMain.Visible = False
        Case 7
            'MonthlySalesAdj.Show
            'MonthlySalesAdjBreakDown.Show
            SalesDB.Show
            
            MotoRep.tbMain.Visible = False
    End Select
End Sub

Private Sub subRptSP_Click(Index As Integer)
    Select Case Index
        Case 1 'Merchandise
            SPrptINV.Show
        Case 2 'Detailed receiving
            InvDtlRecRpt.Show
        Case 3 'Sales report
            InvDtlSalRpt.Show
        Case 4  'Summary of Inv. Adjustment
            'InvAdjSumm.Show
            FrmSummary.mySrc = "Inventory Adjustment"
            FrmSummary.Show
        Case 5  'Summary of Sales Return
            FrmSummary.mySrc = "Sales Return"
            FrmSummary.Show
            'SalesRetSum.Show
        Case 6  'Summary of Purchase Return
            FrmSummary.mySrc = "Purchase Return"
            FrmSummary.Show
            'PurchRetSum.Show
        Case 7
             InvDtlCostSalRpt.Show
    End Select
    MotoRep.tbMain.Visible = False
End Sub

Private Sub subSales_Click(Index As Integer)
    Select Case Index
        Case 1  'sales order
            SPSorder.Show
            MotoRep.tbMain.Visible = False
        Case 2  'warehouse
            SPWarehouse.Show
            MotoRep.tbMain.Visible = False
        Case 3  'Sales Return
            SalesReturn.Show
            MotoRep.tbMain.Visible = False
        Case 5  'Sales Order Summary
            FrmSummary.mySrc = "Sales Order"
            FrmSummary.Show
            'SPSOrdSummary.Show
            MotoRep.tbMain.Visible = False
        Case 6
            FrmSummary.mySrc = "Sales Return"
            FrmSummary.Show
            'SalesRetSum.Show
            MotoRep.tbMain.Visible = False
            
    End Select
End Sub

Private Sub SubSRVC_Click(Index As Integer)
    Select Case Index
        Case 1
            JobOrder.Show
            MotoRep.tbMain.Visible = False
        Case 2
            Requisition.Show vbModal
        Case 3
            FrmSummary.mySrc = "Customer P.O."
            FrmSummary.Show
            'POSummary.Show
        Case 4
            SRVCrecord.Show
            MotoRep.tbMain.Visible = False
        Case 5
            FrmSummary.mySrc = "Mechanic Income"
            FrmSummary.Show
            'SvcMechSumm.Show
    End Select
End Sub
