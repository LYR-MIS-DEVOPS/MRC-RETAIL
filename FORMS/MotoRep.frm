VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MotoRep 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   Caption         =   "==COMPANY NAME=="
   ClientHeight    =   8850
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   12000
   Icon            =   "MotoRep.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MotoRep.frx":57E2
   WindowState     =   2  'Maximized
   Begin VB.PictureBox imgBG 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11940
      TabIndex        =   2
      Top             =   660
      Visible         =   0   'False
      Width           =   12000
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1080
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MotoRep.frx":9170C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MotoRep.frx":9211E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MotoRep.frx":92B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MotoRep.frx":93542
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MotoRep.frx":93F54
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MotoRep.frx":944EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   8400
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   794
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "MotoRep.frx":94A88
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "User Name"
            TextSave        =   "User Name"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3704
            MinWidth        =   3704
            Text            =   "Waiting..."
            TextSave        =   "Waiting..."
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "MotoRep.frx":9567A
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
            Object.Width           =   6006
            Text            =   "Waiting..."
            TextSave        =   "Waiting..."
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "MotoRep.frx":95F54
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
   Begin MSComDlg.CommonDialog cdlgVan 
      Left            =   240
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar toolBarButton 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   1164
      ButtonWidth     =   2461
      ButtonHeight    =   1005
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "SALES ORDER"
            Key             =   "SALESORDER"
            Object.ToolTipText     =   "SALES ORDER"
            Object.Tag             =   "XYF"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "RECEIVE ITEM"
            Key             =   "RECEIVE"
            Object.ToolTipText     =   "REVEIVE ITEM"
            Object.Tag             =   "XYF"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "TRANSFER ITEM"
            Key             =   "TRANSFER"
            Object.ToolTipText     =   "TRANSFER ITEM"
            Object.Tag             =   "XYF"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "WAREHOUSE"
            Key             =   "WAREHOUSE"
            Object.ToolTipText     =   "WAREHOUSE"
            Object.Tag             =   "XYF"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "LOCATION"
            Key             =   "LOCATION"
            Object.ToolTipText     =   "LOCATION"
            Object.Tag             =   "XYF"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "EXTRACT SALES"
            Key             =   "EXTRACTSALES"
            Object.ToolTipText     =   "EXTRACT SALES"
            Object.Tag             =   "XYF"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   1440
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   12240
      m_MDI_Height    =   9720
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
         Caption         =   "&Promo Card Info"
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu mSubPI 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mSubPI 
         Caption         =   "&Supplier"
         Enabled         =   0   'False
         Index           =   7
      End
      Begin VB.Menu mSubPI 
         Caption         =   "S&hipper"
         Enabled         =   0   'False
         Index           =   8
      End
   End
   Begin VB.Menu mnuSP 
      Caption         =   "Spare &Parts"
      Begin VB.Menu mSubSP 
         Caption         =   "&Ordering"
         Enabled         =   0   'False
         Index           =   1
         Begin VB.Menu subOrder 
            Caption         =   "Warehouse Purchase Order"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu subOrder 
            Caption         =   "Replenishment Order"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu subOrder 
            Caption         =   "Partsmatrix Order"
            Enabled         =   0   'False
            Index           =   3
         End
         Begin VB.Menu subOrder 
            Caption         =   "Supplier Purchase Order"
            Enabled         =   0   'False
            Index           =   4
         End
         Begin VB.Menu subOrder 
            Caption         =   "Davao Warehouse Purchase Order"
            Enabled         =   0   'False
            Index           =   5
         End
         Begin VB.Menu subOrder 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu subOrder 
            Caption         =   "Summary of P.O."
            Enabled         =   0   'False
            Index           =   7
         End
         Begin VB.Menu subOrder 
            Caption         =   "Summary of Branch P.O."
            Index           =   8
            Visible         =   0   'False
         End
         Begin VB.Menu subOrder 
            Caption         =   "P.O. Back Order Summary"
            Enabled         =   0   'False
            Index           =   9
         End
         Begin VB.Menu subOrder 
            Caption         =   "Top Selling Spare Parts"
            Index           =   10
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mSubSP 
         Caption         =   "S&tock Incoming/Outgoing"
         Enabled         =   0   'False
         Index           =   2
         Begin VB.Menu subReceive 
            Caption         =   "Receiving Form"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu subReceive 
            Caption         =   "Stock Transfer"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu subReceive 
            Caption         =   "Purchase Return"
            Enabled         =   0   'False
            Index           =   3
         End
         Begin VB.Menu subReceive 
            Caption         =   "Process Discrepancy"
            Enabled         =   0   'False
            Index           =   4
         End
         Begin VB.Menu subReceive 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu subReceive 
            Caption         =   "Receiving Summary"
            Enabled         =   0   'False
            Index           =   6
         End
         Begin VB.Menu subReceive 
            Caption         =   "Transfer Summary"
            Index           =   7
         End
         Begin VB.Menu subReceive 
            Caption         =   "Purchase Return Summary"
            Enabled         =   0   'False
            Index           =   8
         End
         Begin VB.Menu subReceive 
            Caption         =   "Receiving Discrepancy Summary"
            Index           =   9
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
         Begin VB.Menu SPsubStocks 
            Caption         =   "Inventory Control Summary"
            Enabled         =   0   'False
            Index           =   7
         End
         Begin VB.Menu SPsubStocks 
            Caption         =   "-"
            Index           =   8
         End
         Begin VB.Menu SPsubStocks 
            Caption         =   "Parts Matrix Strandard Qty."
            Enabled         =   0   'False
            Index           =   9
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
         Begin VB.Menu SPsubADJ 
            Caption         =   "Update Warehouse"
            Enabled         =   0   'False
            Index           =   4
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
         Caption         =   "&Trust Receipt"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu mSubTR 
         Caption         =   "&Adjustments"
         Enabled         =   0   'False
         Index           =   5
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
            Caption         =   "Settlement of Discrepancy Adj."
            Enabled         =   0   'False
            Index           =   3
         End
         Begin VB.Menu subAdj 
            Caption         =   "Clerical Adjustment"
            Enabled         =   0   'False
            Index           =   4
         End
         Begin VB.Menu subAdj 
            Caption         =   "Adjustment Request"
            Enabled         =   0   'False
            Index           =   5
         End
         Begin VB.Menu subAdj 
            Caption         =   "Cancellation of Transaction"
            Enabled         =   0   'False
            Index           =   6
         End
         Begin VB.Menu subAdj 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu subAdj 
            Caption         =   "A/R DMCM Summary"
            Enabled         =   0   'False
            Index           =   8
         End
         Begin VB.Menu subAdj 
            Caption         =   "Inv. Adj. Summary"
            Enabled         =   0   'False
            Index           =   9
         End
         Begin VB.Menu subAdj 
            Caption         =   "Clerical Adj. Summary"
            Enabled         =   0   'False
            Index           =   10
         End
      End
      Begin VB.Menu mSubTR 
         Caption         =   "&Disbursement"
         Enabled         =   0   'False
         Index           =   6
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
         Index           =   7
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
         Index           =   8
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
            Caption         =   "Payment Allocation"
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
         Begin VB.Menu subRptDailySales 
            Caption         =   "Daily Cash Deposit Report"
            Enabled         =   0   'False
            Index           =   4
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
            Caption         =   "Detailed Outgoing Report"
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
         Begin VB.Menu SubRptMonthly 
            Caption         =   "Extracted Sales"
            Enabled         =   0   'False
            Index           =   8
         End
      End
      Begin VB.Menu mSubRep 
         Caption         =   "Accounting Reports"
         Enabled         =   0   'False
         Index           =   5
         Begin VB.Menu mnuAcctReport 
            Caption         =   "Detailed Inventory Outgoing"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu mnuAcctReport 
            Caption         =   "Detailed Inventory Incoming"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu mnuAcctReport 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuAcctReport 
            Caption         =   "Sales Book"
            Enabled         =   0   'False
            Index           =   4
         End
         Begin VB.Menu mnuAcctReport 
            Caption         =   "Sales Return"
            Enabled         =   0   'False
            Index           =   5
         End
         Begin VB.Menu mnuAcctReport 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu mnuAcctReport 
            Caption         =   "Service Receipts Book"
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
         Index           =   4
         Begin VB.Menu AudTools 
            Caption         =   "Inventory Auditing"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu mnuBARCODE 
            Caption         =   "Barcode Inventory"
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
'sa among pagsugud sa trabaho nga imong gi salig kanamo,
'among pangayaon Lord nga tagaan mi nmo ug igong kusog nga mabuhat namo among mga trabaho nga malipayon
'tabangi mi Lord nga mahuman namo among mga trabaho, ug nag hinaot mi nga kini mka palipay kanimo.
'pasaylua mi sa among mga sala Lord.

'**********************************
'Lord maayong adlaw diha kanimo sa kanunay ikaw nagauban kanamo
'kami nagapasalamat sa gasa sa kinabuhi na walay sawang gihatag nimo kanamo,
'salamat usab lord sa mga tao na nahimong instrumento na nagtabang sa usag usa kanamo
'na sila permi ubanan nimo ug dan-agan mo sa kahayag sa imong hulagway na sila magpadayon sa ilang pagtuman
'sa imong kasuguan.
Public programVersion As String

Private Sub AudTools_Click(Index As Integer)
    Select Case Index
        Case 1
            AuditInvntry.Show
            'tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
    End Select
End Sub

Private Sub BtnMain_Click(Index As Integer)
    Select Case Index
        Case 0
            SPCardex.Show vbModal
        Case 1
            SPordering.Show
            'tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 2
            SPSorder.Show
            'tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 3
            Ledger.Show
            'tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 4
            SPSalesRpt.Show
            'tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 5
            CashFlowMgr.Show
            'tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
    End Select
End Sub

Private Sub MDIForm_Load()
    'sBar.Panels(9).Text = Format(Now, "Mmm. dd, yyyy")
    sBar.Panels(9).Text = ServerTime
    programVersion = " v26.01.12"
    MotoRep.Caption = SysBranch + " " + programVersion
    'v14.3.13 = Fix DP & A/R DMCM w/ Inv. Type(Credit&Cash) error, Parts Desc. Upto 150 chars., Disbursement Refund allow CR
    'V14.6.30 = NEW DSN + LOCAL     use in Audit dept. for running previous backup
    'V14.7.15 = Allocate check to different customer. 2 Printing Preference New OR and Receipt controlled in table settings
    'v14.7.24 = Allocate check to Job Order and Payment can use Customer Name with check allocation even if Check customer name is different. PO from Branch can view Gensan
    'V15.1.2  = TRA for Van Selling, Import Stock Transfer to VAN, Recall of last Date Covered
    'v15.2.17 = fix VAN Sales Return & MRC2 statement of accts, VAN3 on Receipts check allocations
    'v15.3.12 = Branch PO is now Warehouse ready
    'v15.3.22 = fIX BUG On Online Saving og PO, Use table sprecdiscrepancy(REMARKS)
    'V15.8.03 = WAREHOUSE BACK ORDER AND FREEZE COST IN EVERY TRANSACTION CONNECTING TO SPITR
    'v15.11.17 = receiving discrepancies and fix receiving backorder naay bug
    'v15.11.25 = naa nay on acct JO sa 3s og nafix na ang mgdoble sa re-print nga invDetails
    'v16.01.23 = billing new retail
    'v16.05.30 = integrated
    'v16.06.15 = 3s has supplier PO
    'v16.09.01 = Applied Wholesale discount.
    'V16.10.8  = fIX sales order bug on dealer Price + Allow Gas Island for On-Account Wholesale Price
    'V16.12.13 = APPLY WHOLESALE PRICE to selected branch using a value in table settings
    'v16.11.16 = New Billing with Payment Allocation
    'v17.01.17 = PO direct to supplier for 3S branch na fix
    'v17.02.15 = Final Print of PO to supplier fix
    'v17.04.07 = Allow System Date on MRC Sobrecary
    'v17.06.08 = - Fix Stock Transfer, Fix MGR Cashflow 3S signatories, Add new branch background to default MRClogo.
    'v17.16.23 = Remove wholesale discounting using On Account. Only manager can process wholesale
    'v17.07.01 = Fix OR Amount in words from Receive to Payable.
    'v17.07.21 = Fix Sales Order -> On Account w/ DownPayment
    'v17.12.14 = Invoice ug Job Order -> Nag'Add ug Branchname "YAMAHA MOTORHUB ULAS" condition sa printing
    'v18.01.08 = gtanggal ang 5k limit para makadiscount MC Depot Rizal Only
    'v18.04.17 = Replenishment Order
    'v18.06.28 = Update sa PO to WAREHOUSE ug Replenishment Order
    'v18.07.31 = Clerical Adjustment and TRA
    'v18.08.10 = Updated Replenishment Order
    'v18.08.28 = Updated MI and Detailed Sales Report
    'v18.09.25 = Daily Cash Collection & Deposit Monitoring Summary Print
    'v18.09.27 = Updated Daily Cash Collection & Deposit Monitoring Summary Print
    'v18.11.07 = Update PO
    'v18.11.08 = Online Payment
    'v18.11.22 = catalog path in sales order (Find button)
    'v18.11.29 = additional Saving validation in managers cash flow validation form
    'v19.01.09 = gTanggal ang workstation sa where clause sa SAVESALESORDER na declaration
    'v19.02.01 = gFix ang Discounting sa SALES ORDER
    'v19.03.04 = gFix ang Warehouse PO Overflow Error
    'v19.03.20 = Fix the Receipt transaction
    'v19.03.29 = Updates for Cash Deposit Summary
    'v19.04.10 = Updates for PO to WAREHOUSE
    'v19.05.14 = Updates modConnection
    'v19.05.25 = Updates DMCM Adjustment
    'v19.06.01 = Updates sa DMCM ADJUSTMENT FORM sa validation sa ADJNO
    'v19.06.06 = Updates sa Clerical Adjustment JobOrder Payment Type Validation
    'v19.06.17 = Updates sa sales order na WS MAX Discount
    'v19.06.25 = Update sa DMCM SUMMARY PRINTABLE FORM
    'v19.07.25 = New Feature Accounting Reports in Reports Menu
    'v19.08.13 = Updates sa Monthly Sales Report
    'v19.08.30 = Updates sa Replenish Order
    'v19.09.04 = additional features: PO TO warehouse Agdao
    'v19.09.13 = additional features: Receiving SP FROM warehouse Agdao
    'v19.09.24 = modify the Davao Warehouse Agdao Purchase Order
    'v19.10.08 = modify the Provisional
    'v20.01.10 = SPinxst back to normal From WHQTY-WSWHQTY to WHQTY
    'v20.01.21 = SPinxst change from cost to storeprice
    'v20.02.08 = Monthly Cash Collection Report
    'v20.03.07 = Update in SPSORDER,INVADJ,PARTSTRANS,SDISCINVTRYADJ,SPBRANCHINCOMING,SPINCOMING,SPINCOMINGSUPPLIER,SPINCOMINGAGDAO,SPORDERING,SPORDERINGAGDAO,TRA
    'v20.05.26 = Fix bugs sa Monthly Cash Collection
    'v20.06.04 = update the sales order in editing prices. Now: only can edit when the New Price is greater than sa Original Price
    'v20.08.08 = update in cashvoucher summary (Show Form) bugs
    'v20.08.28 = update in inventory adj, official receipt and disbursement (Account Title)
    'v20.09.03 = update in disbursement (update journal)
    'v20.09.11 = update in sales order
    'v20.11.05 = FIXED BUG OF TRUNCATE AMOUNT ERROR (empty data to be stored in DB)in PCV
    'v20.11.14 = Disbursement Module 2.0
    'v20.11.18 = Added check box EXPENSE in CashVoucher
    'v20.12.10 = Update Disbursement
    'v20.12.18 = Added TIN and SUPPLIER on DISBURSEMENT MODULE and 01/04/2021 update in head lease
    'v21.02.05 = Added features Adjustment Request
    'v21.03.11 = Added features Cancellation Module
    'v21.03.29 = Fix Sales Return and Cash flow
    'v21.04.26 = Fix Disbursement Display Record
    'v21.06.10 = Fix Manager Cashflow in Sales Return
    'v21.08.23 = add validation in managers cashflow and sales report regarding in No SP per Invoice
    'v21.09.23 = fixing cash voucher ..disable checkbox of expense ang tof
    'v21.10.06 = fix bug in disbursement sales return/fund
    'v21.10.11 = fix bug in disbursement sales return/fund -customerid ang problem
    'v21.10.14 = fix bug in disbursement sales return/fund - cash or credit ang problem
    'v21.10.15 = updates on adjustment request and cancellation
    'v21.11.10 = updates on printpref in collection receipt
    'v21.12.03 = dmcm added validation
    'v22.02.10 = dmcm fixing query in ledger
    'v22.12.28 = Disbursement add administrative list
    'v23.09.19 = Fix Merchandise Inventory Report w/incomplete description
    'v23.10.28 = Hide Cost and TotalCost in STF and Advanced Search in Sales Order
    'v24.03.21 = Add Key Accounts Discount for Sales Order and get Back the Branch Ordering Module, and Add Superseding in Sales Order
    'v24.05.18 = Remove unit price and  total price column on branch PO to warehouse
    'v24.05.27 = Remove feature superseeding parts on sales order
    'v24.06.06 = Changes on PO to Warehouse module
    '               1. Fixed formula for QH, BO, SQ is >=
    '               2. Back Order Qty should be zero starting June 1, 2024
    '               3. Change ST to "Inbound Intransit Data"
    '               4. Formula for Suggested Order Quantity(SOQ) is (QH + BO + INTRANSIT) - SQ = SOQ
    '               5. Auto cancel of Back orders for Processed POs
    'V24.06.10 = Fix wherein invoices from previous years with the same series of invoice number are included on the data extracted.
    'V24.09.11 (1.0.2) = CHANGE SALES INVOICE FORMAT
    'V24.09.16 (1.1.2) = ADD STORE CODE ON SALES INVOICE NUMBER
    'V24.09.27 (2.1.2) = Service detail added on Quick Service Official Receipt
    'V24.10.28 (2.2.2) = Fixed cash collection and daily collection print preview invoices ticket : MIS241028-33164
    'V24.11.04 (2.3.2) = Fixed processing collection receipt nga dili maka print : MIS241031-02168
    'V24.11.22 (2.4.2) = AUTO CANCEL BACKORDER UPON PROCESSING PURCHASE REQUEST AND CHANGE FORMULA ON SOQ TO : MIS241120-00488
    'v25.02.20 (2.5.2) = fix cash return printing and ang printer pref at ticket : MIS250219-29160 and remove store code at sale invoicing
    'v25.02.20 (2.5.1) = fix bug on DMCM adjustment module by fixing mismatching Balance on ledger  and invoice details with reference : MIS250314-13164
    'v25.03.29 (2.5.2) = fix cash return to add approve by and receeived by with ticket reference : MIS250329-11164
    'v25.05.16 (2.5.3) = fix collection receipt printing vonfliction to official receipt : MIS250513-00158
    'v25.06.18 (2.6.3) = fix upon printing on Receive froom branch print will automtically exits : MIS250618-02713
    'v25.06.24 (2.6.4) = fix cash return expand display to view approved by and recieved by : MIS250625-11163
    'v25.07.17 (2.7.4) = create branch table on sp_mrcretails to get the available store code for module invoicing
    'v25.08.13 (2.8.4) = fix issuing of on account service using cheque : MIS250811-22164
    'v25.09.05 (2.9.4) = replace branch caption on cash return from motorep.caption to query add colum as storecode on bracnh table to display on cash invoice
    'v26.01.12 (2.9.5) = fixed printing pref for Official Receipt on others by modifying query and add condition ticket : MIS251219-30158 MODIFIED BY BEN
    'v26.10.24 (2.10.5) = add function on TRA for issuance of TRA in retail branches for their wholesale customers: MIS261015-01388
    If InStr(SysBranch, "VAN") >= 1 Then
        mSubTR(1).Caption = "Trust Receipt"
        If InStr(SysBranch, "LYR MOTORPARTS CORP. - VAN") >= 1 Then
            mSubTools(2).Caption = "Restore Mobile DB"
        End If
    End If
    
    If InStr(SysBranchCategory, "3S") > 0 Then
        subOrder(3).Visible = True
    End If
    
    FrDate = DateValue("1/1/1977")
    ToDate = DateValue("1/1/1977")
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo er
    Dim Ans As Integer, i As Integer
    Dim MyForms As String
    
    MyForms = ""
    
    For i = 0 To Forms.Count - 1
        MyForms = Forms(i).Caption
    Next i
    If MyForms <> MotoRep.Caption Then
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
Exit Sub
er:
    End
End Sub

Private Sub MDIForm_Resize()
'    Dim lWidth As Long
'    Dim BtnSize As Integer, i As Integer
'
'    If Me.WindowState = 1 Then Exit Sub
'    lWidth = Screen.Width \ Screen.TwipsPerPixelX
'    If lWidth > 850 And Me.Width > 14000 Then
'        BtnSize = 1800
'    Else
'        BtnSize = 1455
'    End If
'    For i = 0 To 5
'        BtnMain(i).Width = BtnSize
'        BtnMain(i).Left = 120 + (BtnMain(i).Width * i)
'    Next i
'
'    UpDown1.Left = Me.ScaleWidth - 300
'    Shape1.Left = Me.ScaleWidth - 2450
'    lblUser.Left = Me.ScaleWidth - 2400
On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    Me.WindowState = 2 'Maximized
    DoEvents
    imgBG.Left = 0
    imgBG.Width = Me.Width
    imgBG.Height = Me.Height
    
    'lblBranchName.Left = 0
    'lblBranchName.Width = Me.Width

End Sub

Private Sub mnuDIO_Click()
    
End Sub

Private Sub mnuAcctReport_Click(Index As Integer)
    Select Case Index
        Case 1
            FrmSummary.mySrc = "Detailed Inventory Outgoing"
            FrmSummary.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 2
            FrmSummary.mySrc = "Detailed Inventory Incoming"
            FrmSummary.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 4
            FrmSummary.mySrc = "Sales Book"
            FrmSummary.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 5
            FrmSummary.mySrc = "Sales Return Report"
            FrmSummary.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 7
            FrmSummary.mySrc = "Service Receipts Book"
            FrmSummary.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
    End Select
End Sub

Private Sub mnuBARCODE_Click()
 frmMain.Show
 
 
End Sub

'Private Sub mnuInvCtrlSum_Click()
'    frmInventoryCtrl.Show
'    'MotoRep.tbMain.Visible = False
'    MotoRep.imgBG.Visible = False
'    MotoRep.toolBarButton.Visible = False
'End Sub

'Private Sub mnuOverallQty_click()
''    frmOvrallQty.Show
''    'MotoRep.tbMain.Visible = False
''    MotoRep.imgBG.Visible = False
''    MotoRep.toolBarButton.Visible = False
'End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub

Private Sub mSetOpt_Click()
    SEToptions.Show vbModal
End Sub

Private Sub mSetPrn_Click()
    TLSprintPref.Show
    'MotoRep.tbMain.Visible = False
    MotoRep.imgBG.Visible = False
    MotoRep.toolBarButton.Visible = False
End Sub

Private Sub mSubPI_Click(Index As Integer)
    Select Case Index
        Case 1
            Employees.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 2  'User
            EMPuserNew.Show vbModal
            'MotoRep.tbMain.Enabled = False
        Case 4  'Branch
            AcctCustomer.Show vbModal
        Case 5  'Branch
            frmPromoCard.Show vbModal
        Case 7 'Supplier
            EMPsupplier.Show vbModal
        Case 8  'Shipper
            EMPshipper.Show vbModal
    End Select
End Sub

Private Sub mSubTools_Click(Index As Integer)
    Dim rsBackUp As New ADODB.Recordset
    Dim myFile As String, myFile2 As String
    Dim Qry As String, BranchDB As String
    
    On Error GoTo LocalErr
    Select Case Index
        
        Case 1  'Backup
            myFile = Dir("C:\MIS\BACKUP\MYSQLDUMP.EXE")
            myFile2 = Dir("C:\MIS\BACKUP\RAR.EXE")
            If UCase$(myFile) <> "MYSQLDUMP.EXE" Or UCase$(myFile2) <> "RAR.EXE" Then
                MsgBox "Backup File Systems missing, Please consult MIS Help Desk.", vbExclamation + vbOKOnly, "Warning"
                Exit Sub
            End If
            Qry = "SELECT settings.* From settings Where settings.DESCRIPTION = 'BACKUP' AND settings.UPDATEDBY = 2"
            With rsBackUp
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    myFile = IIf(IsNull(!Value), "", !Value)
                    BranchDB = IIf(IsNull(!Remarks), "", !Remarks)
                Else
                    myFile = ""
                    BranchDB = ""
                End If
                .Close
            End With
            If myFile = "" Or BranchDB = "" Then
                MsgBox "Backup File Settings missing, Please consult MIS Help Desk.", vbExclamation + vbOKOnly, "Warning"
                Exit Sub
            End If
            
            With cdlgVan
                .CancelError = True
                .Filter = "Backup Files (*.RAR)|(*.RAR)"
                .InitDir = App.Path
                myFile = myFile & Format(Now, "mmddyyhhmm")
                .filename = myFile & ".RAR"
                .ShowSave
                myFile2 = .filename
            End With
            
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
            If UCase$(SysBranch) <> "LYR MOTORPARTS CORP. - VAN-2 [LHC403]" And UCase$(SysBranch) <> "LYR MOTORPARTS CORP. - VAN-56 [AAH1233]" And UCase$(SysBranch) <> "LYR MOTORPARTS CORP. - VAN-55 [AAH1250]" Then
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
            
            If UCase$(SysBranch) = "LYR MOTORPARTS CORP. - VAN-2 [LHC403]" Then
                myFile = "LHC403VAN"
                BranchDB = "van1mobile"
                
            ElseIf UCase$(SysBranch) = "LYR MOTORPARTS CORP. - VAN-56 [AAH1233]" Then
                myFile = "AAH1233VAN"
                BranchDB = "van2mobile"
                
            ElseIf UCase$(SysBranch) = "LYR MOTORPARTS CORP. - VAN-55 [AAH1250]" Then
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
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 2
            Receipt.Show 'vbModal
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 3
            arPR.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 4
            TRA.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
    End Select
    
End Sub

Private Sub SPsubADJ_Click(Index As Integer)
Dim myFile As String, myFile2 As String, myFile3 As String
    Select Case Index
        Case 1  'SP location adjustment
            SPLocAdj.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 2
            Qry = "call sp_promotionalonline.sp_promocard_update_CustToBranch('" & sysBranchCode & "')"
            With rsAnyUse
                If .State = adStateOpen Then .Close
                If .RecordCount > 0 Then
                    .MoveFirst
                    If !level = "Granted" Then
                        MsgBox "" & !message & "", vbInformation + vbOKOnly, "Validation"
                    Else
                        MsgBox "" & !message & "", vbInformation + vbOKOnly, "Validation"
                    End If
                End If
                .Close
            End With
            'SPupdate.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 3  'Pricing
            SPpricing.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 4  'Update Warehouse
            
            myFile = Dir("F:\Dropbox\MRCWAREHOUSE\MYSQL.EXE")
            myFile2 = Dir("F:\Dropbox\MRCWAREHOUSE\RAR.EXE")
            myFile3 = Dir("F:\Dropbox\MRCWAREHOUSE\MRCWAREHOUSE.RAR")
            If UCase$(myFile) <> "MYSQL.EXE" Or UCase$(myFile2) <> "RAR.EXE" Or UCase$(myFile3) <> "MRCWAREHOUSE.RAR" Then
                MsgBox "Restore File System missing, Please consult IT Support Group.", vbExclamation + vbOKOnly, "Warning"
                Exit Sub
            End If

            myFile = "mrcwarehouse"
            myFile2 = "F:\Dropbox\MRCWAREHOUSE\mrcwarehouse.rar"
            
            Open "Restore.bat" For Output As #1
            Print #1, "@echo off"
            Print #1, "echo Restoring your BACKUP for LYR MOTORPARTS WAREHOUSE, Please wait....."
            Print #1, "F:\Dropbox\MRCWAREHOUSE\rar e -plaborday " & myFile2
            Print #1, "echo Please wait...Updating server databases"
            Print #1, "F:\Dropbox\MRCWAREHOUSE\mysql -uroot -plaborday " & myFile & " < " & Left(myFile2, Len(myFile2) - 3) & "SQL"
            Print #1, "del *.sql"
            Print #1, "echo Data Successfully loaded"
            Print #1, "pause"
            Print #1, "del Restore.bat"
            Close #1
            Shell "Restore.bat", vbNormalNoFocus
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
                SPMerInvCost.Show
            End If
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 5  'Inventory Monitoring Report
            'SPInventory.Show
            frmHSummary.MySource = "MONTHLY MERCHANDISE INVENTORY"
            frmHSummary.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 6  'Price List
            PRNmc.MySource = "PRICELIST"
            PRNmc.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 7
            frmInventoryCtrl.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 9
            BranchSQTY.Show
            'frmOvrallQty.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        
    End Select
End Sub


Private Sub subAdj_Click(Index As Integer)

    Select Case Index
        Case 1  'AR DMCM
            DMCMadj.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 2
            InvntryADJ.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
'*************** CODE UPDATED 10/11/17 *************
        Case 3
            SDiscInvtryAdj.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
'*************** CODE UPDATED 10/11/17 *************
        Case 4
            ClericalADJ.Show vbModal
        Case 5
            frmOtherAdjustment.Show
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 6
            frmCancellation.Show vbModal
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 8  'DMCM Summary
            'DmCmSumm.Show
            FrmSummary.mySrc = "DMCM [A/R]"
            FrmSummary.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 9
            FrmSummary.mySrc = "Inventory Adjustment"
            FrmSummary.Show
            'InvAdjSumm.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 10
            FrmSummary.mySrc = "Clerical Adjustment"
            FrmSummary.Show
            'ClrAdjSumm.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
    End Select
End Sub

Private Sub subAR_Click(Index As Integer)
    Select Case Index
        Case 1
            'SPrptAR.Show
'            If SysBranchID = 21 Then 'wholesale
'                frmAgingSum.Show
'            Else
'                AgingNew.Show
'                'frmAgingSum.Show
'            End If
            frmAgingSum.Show
        Case 2
            If SysBranch = "LYR MOTORPARTS CORP. - WHOLESALE" Then
                BillingNew.Show
            Else
                'BillingProcess.Show
                BillingNewRetail.Show
            End If
        Case 3
            Ledger.Show
        Case 4
            DPSummary.Show
        Case 5
            QuickSrvInvoice.Show
        Case 6
            'AgingNew.Show
            frmAllocationSummary.Show
    End Select
    'MotoRep.tbMain.Visible = False
    MotoRep.imgBG.Visible = False
    MotoRep.toolBarButton.Visible = False
    
End Sub

Private Sub subCashFlow_Click(Index As Integer)
    Select Case Index
        Case 1  'Cashier Cash Flow
            CashFlowCashier.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 2  'Mgr Cash flow
            CashFlowMgr.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 3
            ExpAdvLiq.Show vbModal
        Case 4
            ExpCashRet.Show vbModal
        Case 6  'Check Summary
            arCheckSumm.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
    End Select

End Sub

Private Sub subDisb_Click(Index As Integer)
    Select Case Index
        Case 1  'PCV
            DoEvents
            DisbursementV2.Show vbModal
        Case 2  'CV
            CashVoucher.Show vbModal
        Case 4  'PCV summary
            DisburseSumm.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 5  'CV summary
            'CVsummary.ViewList = True
            'CVsummary.Show
            FrmSummary.mySrc = "Cash Voucher"
            FrmSummary.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
    End Select
End Sub
Private Sub subOrder_Click(Index As Integer)
    Select Case Index
        
        'New-------------------------------
        Case 1  'PO
            SPordering.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 2  'Replenishment PO
            BranchReplPO.Show
            'BranchIdentifying.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 3  'Replenishment PO
            BranchIdentifying.Show
            'BranchIdentifying.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 4
            'Supplier P.O. INCLUSIVE TO 3S BRANCHES
            frmSupplierPO.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
            '3 - Separator
        Case 5
            'Davao Warehouse Purchase Order
            SPorderingWHAgdao.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
            '3 - Separator
        Case 7
            'Summary
            'SPorderSum.Show
            'MotoRep.tbMain.Visible = False
            FrmSummary.mySrc = "Branch Order"
            FrmSummary.Show
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 8
            'Summary Branch PO
            FrmSummary.mySrc = "Branch Purchase Order"
            FrmSummary.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
            
        Case 9 'Back Order
'            SPbckOrdSum.Show
'            MotoRep.tbMain.Visible = False
            frmHSummary.MySource = "LYR MOTORPARTS WAREHOUSE BACK ORDER"
            frmHSummary.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 10 'Top Selling spare parts
            SPTopSales.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
    End Select
End Sub

Private Sub subReceive_Click(Index As Integer)
    Select Case Index
        Case 1  'Receive SP
            SPincoming.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.toolBarButton.Visible = False
            MotoRep.imgBG.Visible = False
        Case 2  'Stock Transfer
            ParTrans.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.toolBarButton.Visible = False
            MotoRep.imgBG.Visible = False
        Case 3  'Purchase Return
            PurchaseReturn.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.toolBarButton.Visible = False
            MotoRep.imgBG.Visible = False
    '********** UPDATED BY 10/11/17 *********
        Case 4  'Processing Discrepancy
            frmSDiscrepancy.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.toolBarButton.Visible = False
            MotoRep.imgBG.Visible = False
    '******* END UPDATED BY 10/11/17 *********
        Case 6  'Receive Summary
            FrmSummary.mySrc = "Received Spare Parts"
            FrmSummary.Show
            'SPRecSummary.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.toolBarButton.Visible = False
            MotoRep.imgBG.Visible = False
        Case 7  'transfer Summary
            'FrmSummary.mySrc = "Received Spare Parts"
            TranSummary.Show
            'SPRecSummary.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.toolBarButton.Visible = False
            MotoRep.imgBG.Visible = False
        Case 8  'Receive Summary
            'PurchRetSum.Show
            FrmSummary.mySrc = "Purchase Return"
            FrmSummary.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.toolBarButton.Visible = False
            MotoRep.imgBG.Visible = False
        Case 9 'INVENTORY DISCREPANCY SUMMARY
            frmHSummary.MySource = "RECEIVING DISCREPANCY"
            frmHSummary.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.toolBarButton.Visible = False
            MotoRep.imgBG.Visible = False
    End Select
End Sub

Private Sub subRptDailySales_Click(Index As Integer)
    Select Case Index
        Case 1 'Daily Sales Report
            SPSalesRpt.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 2  'Daily Cash Collection
            DailyCash.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 4
            FrmSummary.mySrc = "Daily Cash Deposit Report"
            FrmSummary.Show
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 6  'Inventory
            SPInventory.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
    End Select
End Sub

Private Sub SubRptMonthly_Click(Index As Integer)
    Select Case Index
        Case 1  'Sales
            MonthlySales.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 2
            MonthlyColl.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 6
            'Sales Distribution by Parts
            SalesDistribution.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 7
            'MonthlySalesAdj.Show
            'MonthlySalesAdjBreakDown.Show
            SalesDB.Show
            
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
'*************** CODE UPDATED 11/27/17 *****************
        Case 8
            frmHSummary.MySource = "EXTRACTSALES"
            frmHSummary.Show
            
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
'*********** END CODE UPDATED 11/27/17 *****************
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
'            FrmSummary.mySrc = "Purchase Return"
'            FrmSummary.Show
            'PurchRetSum.Show
            frmHSummary.MySource = "DETAILED STOCKTRANSFER"
            frmHSummary.Show
        Case 7
             InvDtlCostSalRpt.Show
    End Select
    'MotoRep.tbMain.Visible = False
    MotoRep.imgBG.Visible = False
    MotoRep.toolBarButton.Visible = False
End Sub

Private Sub subSales_Click(Index As Integer)
    Select Case Index
        Case 1  'sales order
            SPSorder.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 2  'warehouse
            SPWarehouse.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 3  'Sales Return
            SalesReturn.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 5  'Sales Order Summary
            FrmSummary.mySrc = "Sales Order"
            FrmSummary.Show
            'SPSOrdSummary.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 6
            FrmSummary.mySrc = "Sales Return"
            FrmSummary.Show
            'SalesRetSum.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
            
    End Select
End Sub

Private Sub SubSRVC_Click(Index As Integer)
    Select Case Index
        Case 1
            JobOrder.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 2
            Requisition.Show vbModal
        Case 3
            FrmSummary.mySrc = "Customer P.O."
            FrmSummary.Show
            'POSummary.Show
        Case 4
            SRVCrecord.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 5
            FrmSummary.mySrc = "Mechanic Income"
            FrmSummary.Show
            'SvcMechSumm.Show
    End Select
End Sub

Private Sub toolBarButton_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    If Button.Visible = False Then Exit Sub
    If Button.Key = "SALESORDER" Then
        SPSorder.Show
        MotoRep.toolBarButton.Visible = False
        MotoRep.imgBG.Visible = False
    ElseIf Button.Key = "RECEIVE" Then
        frmChoose.MySource = "SELECTEDRECEIVE"
        frmChoose.Show vbModal
        
        'MotoRep.toolBarButton.Visible = False
        'MotoRep.imgBG.Visible = False
    ElseIf Button.Key = "TRANSFER" Then
        ParTrans.Show
        MotoRep.toolBarButton.Visible = False
        MotoRep.imgBG.Visible = False
    ElseIf Button.Key = "WAREHOUSE" Then
        SPWarehouse.Show
        MotoRep.toolBarButton.Visible = False
        MotoRep.imgBG.Visible = False
    ElseIf Button.Key = "LOCATION" Then
        UpdateLoc.Show vbModal
    ElseIf Button.Key = "EXTRACTSALES" Then
        frmHSummary.MySource = "EXTRACTSALES"
        frmHSummary.Show
        MotoRep.toolBarButton.Visible = False
        MotoRep.imgBG.Visible = False
    End If
    
End Sub
