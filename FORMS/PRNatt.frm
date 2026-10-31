VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form PRNatt 
   Caption         =   "Attachments"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9450
   Icon            =   "PRNatt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   9450
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Tag             =   "xy"
      Top             =   480
      Width           =   9255
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "E&xit Report"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "PRNatt.frx":08CA
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   9720
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   9690
      m_MDI_Height    =   7590
   End
End
Attribute VB_Name = "PRNatt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mcLyr As CRAXDRT.Application
Dim mcReport As CRAXDRT.Report

Dim rsPrn As New ADODB.Recordset
Public QryPrn As String
Public MySource As String
Public ParamValue1 As String
Public ParamValue2 As String
Public ParamValue3 As String
Public ParamValue4 As String
Public ParamValue5 As String
Public ParamValue6 As String
Public ParamValue7 As String
Public ParamValue8 As String

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)
    
    If MySource = "JOB ORDER" Then
        If Val(JobOrder.txtPrintCopy.Text) = 0 Then
            LYRcon.Execute "UPDATE SrvcJobOrder SET DATEPROMISED = '" & Format(Date, "yyyy-mm-dd") & "', PRINTCOPY = 1 Where srvcJobOrder.JOBORDERID = " & Val(JobOrder.txtJOid.Text) & ";"
        Else
            LYRcon.Execute "UPDATE SrvcJobOrder SET PRINTCOPY = " & Val(JobOrder.txtPrintCopy.Text) + 1 & " Where srvcJobOrder.JOBORDERID = " & Val(JobOrder.txtJOid.Text) & ";"
        End If
        JobOrder.txtPrintCopy.Text = Val(JobOrder.txtPrintCopy.Text) + 1
    
    End If

End Sub

Private Sub Form_Load()
    If MySource = "ATTACHMENT" Then
        Dim crAttch As New rptORattachment
        
        rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
        crAttch.DiscardSavedData
        crAttch.Database.SetDataSource rsPrn
        CRViewer1.ReportSource = crAttch
        CRViewer1.ViewReport
        CRViewer1.Zoom 50
        CRViewer1.DisplayGroupTree = False
        
    ElseIf MySource = "LEDGER" Then
        Dim crLdgr As New rptLedger
        
        Me.Caption = "LEDGER"
        rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
        crLdgr.DiscardSavedData
        crLdgr.Database.SetDataSource rsPrn
        crLdgr.ParameterFields(1).ClearCurrentValueAndRange
        crLdgr.ParameterFields(1).AddCurrentValue ParamValue1
        crLdgr.ParameterFields(2).ClearCurrentValueAndRange
        crLdgr.ParameterFields(2).AddCurrentValue ParamValue2
        crLdgr.ParameterFields(3).ClearCurrentValueAndRange
        crLdgr.ParameterFields(3).AddCurrentValue ParamValue3
        crLdgr.ParameterFields(4).ClearCurrentValueAndRange
        crLdgr.ParameterFields(4).AddCurrentValue ParamValue4
        crLdgr.ParameterFields(5).ClearCurrentValueAndRange
        crLdgr.ParameterFields(5).AddCurrentValue ParamValue5
        crLdgr.ParameterFields(6).ClearCurrentValueAndRange
        crLdgr.ParameterFields(6).AddCurrentValue ParamValue6
        crLdgr.ParameterFields(7).ClearCurrentValueAndRange
        crLdgr.ParameterFields(7).AddCurrentValue ParamValue7
        crLdgr.ParameterFields(8).ClearCurrentValueAndRange
        crLdgr.ParameterFields(8).AddCurrentValue ParamValue8
        crLdgr.ParameterFields(9).ClearCurrentValueAndRange
        crLdgr.ParameterFields(9).AddCurrentValue SysBranch
        crLdgr.ParameterFields(10).ClearCurrentValueAndRange
        crLdgr.ParameterFields(10).AddCurrentValue SysAddress
        CRViewer1.ReportSource = crLdgr
        CRViewer1.ViewReport
        CRViewer1.DisplayGroupTree = False
        'CRViewer1.EnableExportButton = False
        CRViewer1.EnableExportButton = True
        
    ElseIf MySource = "SPINVOICE" Then
        Dim crSPinv As New rptInvoice
        
        rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
        crSPinv.DiscardSavedData
        crSPinv.Database.SetDataSource rsPrn
        crSPinv.ParameterFields(1).ClearCurrentValueAndRange
        crSPinv.ParameterFields(1).AddCurrentValue ParamValue1
        crSPinv.ParameterFields(2).ClearCurrentValueAndRange
        crSPinv.ParameterFields(2).AddCurrentValue ParamValue2
        crSPinv.ParameterFields(3).ClearCurrentValueAndRange
        crSPinv.ParameterFields(3).AddCurrentValue ParamValue3
        crSPinv.ParameterFields(4).ClearCurrentValueAndRange
        crSPinv.ParameterFields(4).AddCurrentValue ParamValue4
        crSPinv.ParameterFields(5).ClearCurrentValueAndRange
        crSPinv.ParameterFields(5).AddCurrentValue ParamValue5
        crSPinv.ParameterFields(6).ClearCurrentValueAndRange
        crSPinv.ParameterFields(6).AddCurrentValue ParamValue6
        crSPinv.ParameterFields(7).ClearCurrentValueAndRange
        crSPinv.ParameterFields(7).AddCurrentValue SysBranch
        crSPinv.ParameterFields(8).ClearCurrentValueAndRange
        crSPinv.ParameterFields(8).AddCurrentValue SysAddress
        
        CRViewer1.ReportSource = crSPinv
        CRViewer1.ViewReport
        CRViewer1.DisplayGroupTree = False
        CRViewer1.EnableExportButton = True
        CRViewer1.Zoom 75
    ElseIf MySource = "JOB ORDER" Then
        
        
            Call DisplayJobORderRPT
        
    ElseIf MySource = "TRA" Then
        Dim crTRAsalorder As New TRAorder
    
        rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
        crSPinv.DiscardSavedData
        crTRAsalorder.Database.SetDataSource rsPrn
        crTRAsalorder.ParameterFields(1).ClearCurrentValueAndRange
        crTRAsalorder.ParameterFields(1).AddCurrentValue ParamValue1
        crTRAsalorder.ParameterFields(2).ClearCurrentValueAndRange
        crTRAsalorder.ParameterFields(2).AddCurrentValue ParamValue2
        crTRAsalorder.ParameterFields(3).ClearCurrentValueAndRange
        crTRAsalorder.ParameterFields(3).AddCurrentValue ParamValue3
        crTRAsalorder.ParameterFields(4).ClearCurrentValueAndRange
        crTRAsalorder.ParameterFields(4).AddCurrentValue ParamValue4
        crTRAsalorder.ParameterFields(5).ClearCurrentValueAndRange
        crTRAsalorder.ParameterFields(5).AddCurrentValue ParamValue5
        crTRAsalorder.ParameterFields(6).ClearCurrentValueAndRange
        crTRAsalorder.ParameterFields(6).AddCurrentValue ParamValue6
        crTRAsalorder.ParameterFields(7).ClearCurrentValueAndRange
        crTRAsalorder.ParameterFields(7).AddCurrentValue SysBranch
        
        CRViewer1.ReportSource = crTRAsalorder
        CRViewer1.ViewReport
        CRViewer1.DisplayGroupTree = False
        CRViewer1.EnableExportButton = True
        CRViewer1.Zoom 75
    
    
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsPrn.State = adStateOpen Then rsPrn.Close
End Sub


Sub DisplayJobORderRPT()
      Dim txtH As TextObject
      Dim txtD As TextObject
      Dim lnObjD As LineObject
      Dim PageNumber As FieldObject
      Dim Qry As String, Brn As String, Adrs As String, i As Integer, lstRow As Integer
      
      Set M_Proj = New CRAXDRT.Application
      Set CF_JOB = M_Proj.NewReport
          
      'Report settings
      CF_JOB.PaperSize = crPaperLetter
      CF_JOB.PaperOrientation = crPortrait
      CF_JOB.TopMargin = 500
      
      'Section 1: Report Header
      CF_JOB.Sections(1).Suppress = True
      'Section 2: Page Header
      CF_JOB.Sections(2).Height = 700
      Set txtH = CF_JOB.Sections(2).AddTextObject(SysBranch, 500, 0)
      txtH.Width = 11000: txtH.Height = 300: txtH.Font.Name = "ARIAL": txtH.Font.Size = 12: txtH.Font.Bold = True: txtH.HorAlignment = crLeftAlign
      Set txtH = CF_JOB.Sections(2).AddTextObject("JOB ORDER", 8500, 0)
      txtH.Width = 2000: txtH.Height = 300: txtH.Font.Name = "ARIAL": txtH.Font.Size = 12: txtH.Font.Bold = True: txtH.HorAlignment = crRightAlign
      Set txtH = CF_JOB.Sections(2).AddTextObject(SysAddress, 500, 300)
      txtH.Width = 11000: txtH.Height = 300: txtH.Font.Name = "ARIAL": txtH.Font.Size = 10: txtH.Font.Bold = True: txtH.HorAlignment = crLeftAlign
      'section 3: Details
      With JobOrder
          CF_JOB.Sections(3).Height = 13000
          Set txtD = CF_JOB.Sections(3).AddTextObject("CUSTOMER: " & .txtCustomer.Text, 500, 0)
          txtD.Width = 5000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crLeftAlign
          Set txtD = CF_JOB.Sections(3).AddTextObject("ADDRESS : " & .txtAddress.Text, 500, 300)
          txtD.Width = 5000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crLeftAlign
          Set txtD = CF_JOB.Sections(3).AddTextObject("SERVICE TYPE: " & .cboType.Text, 500, 600)
          txtD.Width = 5000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crLeftAlign
          Set txtD = CF_JOB.Sections(3).AddTextObject("PAYMENT TYPE: " & .chkType.Caption, 500, 900)
          txtD.Width = 5000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crLeftAlign
          
          Set txtD = CF_JOB.Sections(3).AddTextObject("JOB ORDER#:", 8000, 0)
          txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crRightAlign
          Set txtD = CF_JOB.Sections(3).AddTextObject(.txtJobNo.Text, 9600, 0)       'JO#
          txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crLeftAlign
          Set txtD = CF_JOB.Sections(3).AddTextObject("DATE IN:", 8000, 300)   'Date
          txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crRightAlign
          Set txtD = CF_JOB.Sections(3).AddTextObject(.txtDateRec.Text, 9600, 300)     'Date IN
          txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crLeftAlign
          Set txtD = CF_JOB.Sections(3).AddTextObject("DATE OUT:", 8000, 600)   'Date OUT
          txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crRightAlign
          Set txtD = CF_JOB.Sections(3).AddTextObject(IIf(IsDate(.txtDateProm.Text), .txtDateProm.Text, Format(Date, "MM/DD/YYYY")), 9600, 600)   'Date OUT
          txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crLeftAlign
          
          
          Set lnObjD = CF_JOB.Sections(3).AddLineObject(500, 1500, 11000, 1500) 'FRAME
          lnObjD.LineThickness = 1
          Set lnObjD = CF_JOB.Sections(3).AddLineObject(500, 1800, 11000, 1800)
          lnObjD.LineThickness = 1
          Set lnObjD = CF_JOB.Sections(3).AddLineObject(500, 2100, 11000, 2100)
          lnObjD.LineThickness = 1
          Set lnObjD = CF_JOB.Sections(3).AddLineObject(500, 1500, 500, 2100) 'column1
          lnObjD.LineThickness = 1
          Set lnObjD = CF_JOB.Sections(3).AddLineObject(2500, 1500, 2500, 2100) 'column2
          lnObjD.LineThickness = 1
          Set lnObjD = CF_JOB.Sections(3).AddLineObject(4500, 1500, 4500, 2100) 'column3
          lnObjD.LineThickness = 1
          Set lnObjD = CF_JOB.Sections(3).AddLineObject(6500, 1500, 6500, 2100) 'column4
          lnObjD.LineThickness = 1
          Set lnObjD = CF_JOB.Sections(3).AddLineObject(8500, 1500, 8500, 2100) 'column5
          lnObjD.LineThickness = 1
          Set lnObjD = CF_JOB.Sections(3).AddLineObject(9700, 1500, 9700, 2100) 'column6
          lnObjD.LineThickness = 1
          Set lnObjD = CF_JOB.Sections(3).AddLineObject(11000, 1500, 11000, 2100) 'column7
          lnObjD.LineThickness = 1
    
          Set txtD = CF_JOB.Sections(3).AddTextObject("MC MODEL", 500, 1550)
          txtD.Width = 2000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crHorCenterAlign
          Set txtD = CF_JOB.Sections(3).AddTextObject("ENGINE NO.", 2500, 1550)
          txtD.Width = 2000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crHorCenterAlign
          Set txtD = CF_JOB.Sections(3).AddTextObject("CHASSIS NO.", 4500, 1550)
          txtD.Width = 2000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crHorCenterAlign
          Set txtD = CF_JOB.Sections(3).AddTextObject("PLATE NO.", 6500, 1550)
          txtD.Width = 2000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crHorCenterAlign
          Set txtD = CF_JOB.Sections(3).AddTextObject("USAGE", 8500, 1550)
          txtD.Width = 1200: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crHorCenterAlign
          Set txtD = CF_JOB.Sections(3).AddTextObject("MILEAGE", 9700, 1550)
          txtD.Width = 1200: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crHorCenterAlign
          'Display content of FlxMC
          Set txtD = CF_JOB.Sections(3).AddTextObject(.flxMC.TextMatrix(1, 0), 500, 1850)
          txtD.Width = 2000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crHorCenterAlign
          Set txtD = CF_JOB.Sections(3).AddTextObject(.flxMC.TextMatrix(1, 1), 2500, 1850)
          txtD.Width = 2000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crHorCenterAlign
          Set txtD = CF_JOB.Sections(3).AddTextObject(.flxMC.TextMatrix(1, 2), 4500, 1850)
          txtD.Width = 2000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crHorCenterAlign
          Set txtD = CF_JOB.Sections(3).AddTextObject(.flxMC.TextMatrix(1, 3), 6500, 1850)
          txtD.Width = 2000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crHorCenterAlign
          Set txtD = CF_JOB.Sections(3).AddTextObject(.flxMC.TextMatrix(1, 4), 8500, 1850)
          txtD.Width = 1200: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crHorCenterAlign
          Set txtD = CF_JOB.Sections(3).AddTextObject(.flxMC.TextMatrix(1, 5), 9700, 1850)
          txtD.Width = 1200: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crHorCenterAlign
          
          Set txtD = CF_JOB.Sections(3).AddTextObject("J O B   R E Q U E S T", 500, 2200)
          txtD.Width = 5000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crHorCenterAlign: txtD.Font.Bold = True
          Set txtD = CF_JOB.Sections(3).AddTextObject("FSC NUMBER", 5500, 2200)
          txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crHorCenterAlign: txtD.Font.Bold = True
          Set txtD = CF_JOB.Sections(3).AddTextObject("LABOR", 7000, 2200)
          txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crHorCenterAlign: txtD.Font.Bold = True
          Set txtD = CF_JOB.Sections(3).AddTextObject("DISCOUNT", 8500, 2200)
          txtD.Width = 1200: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crHorCenterAlign: txtD.Font.Bold = True
          Set txtD = CF_JOB.Sections(3).AddTextObject("TOTAL", 9700, 2200)
          txtD.Width = 1200: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crHorCenterAlign: txtD.Font.Bold = True
      End With
      With JobOrder.flxJob
        lstRow = 2200
        For i = 1 To .Rows - 1
            lstRow = lstRow + 300
            Set txtD = CF_JOB.Sections(3).AddTextObject(.TextMatrix(i, 1), 700, lstRow)
            txtD.Width = 5000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crLeftAlign
            Set txtD = CF_JOB.Sections(3).AddTextObject(.TextMatrix(i, 2), 5500, lstRow)
            txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crLeftAlign
            Set txtD = CF_JOB.Sections(3).AddTextObject(.TextMatrix(i, 3), 7000, lstRow)
            txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crHorCenterAlign
            Set txtD = CF_JOB.Sections(3).AddTextObject(.TextMatrix(i, 4), 8500, lstRow)
            txtD.Width = 1200: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crRightAlign
            Set txtD = CF_JOB.Sections(3).AddTextObject(.TextMatrix(i, 5), 9700, lstRow)
            txtD.Width = 1200: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crRightAlign
        Next i
     End With
     With JobOrder
        lstRow = lstRow + 400
        Set txtD = CF_JOB.Sections(3).AddTextObject("TOTAL:", 8500, lstRow)
        txtD.Width = 1200: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crRightAlign: txtD.Font.Bold = True
        Set txtD = CF_JOB.Sections(3).AddTextObject(.txtTotal.Text, 9700, lstRow)
        txtD.Width = 1200: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crRightAlign: txtD.Font.Bold = True
        If Val(Replace(.txtDiscAmt.Text, ",", "")) > 0 Then
            lstRow = lstRow + 300
            Set txtD = CF_JOB.Sections(3).AddTextObject("LESS " & .txtPctD.Text & " DISCOUNT:", 8000, lstRow)
            txtD.Width = 1700: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crRightAlign
            Set txtD = CF_JOB.Sections(3).AddTextObject(.txtDiscAmt.Text, 9700, lstRow)
            txtD.Width = 1200: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crRightAlign
            lstRow = lstRow + 300
            Set txtD = CF_JOB.Sections(3).AddTextObject("GRAND TOTAL:", 8500, lstRow)
            txtD.Width = 1200: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crRightAlign: txtD.Font.Bold = True
            Set txtD = CF_JOB.Sections(3).AddTextObject(.txtGrandTtl.Text, 9700, lstRow)
            txtD.Width = 1200: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crRightAlign: txtD.Font.Bold = True
        End If
        If .txtRemarks.Text <> "" Then
            Set txtD = CF_JOB.Sections(3).AddTextObject("Remarks: " & .txtRemarks.Text, 1000, 5000)
            txtD.Width = 8000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crLeftAlign
        End If
        Set txtD = CF_JOB.Sections(3).AddTextObject("O.R. NO. ", 9000, 5800)
        txtD.Width = 1000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
        Set lnObjD = CF_JOB.Sections(3).AddLineObject(10000, 6050, 11000, 6050) 'OR LINE
        lnObjD.LineThickness = 1
        
        Set txtD = CF_JOB.Sections(3).AddTextObject("Prepared By: " & .txtPrepby.Text, 500, 5300)
        txtD.Width = 5000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crLeftAlign
        Set txtD = CF_JOB.Sections(3).AddTextObject("Mechanic: " & .cboMechanic.Text, 500, 5600)
        txtD.Width = 5000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crLeftAlign
        Set txtD = CF_JOB.Sections(3).AddTextObject("Date Printed: " & Format(Now, "mm/dd/yyyy hh:mm ampm"), 500, 5900)
        txtD.Width = 5000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 8:  txtD.HorAlignment = crLeftAlign
    
    End With
      'Display Report in Viewer
      CRViewer1.ReportSource = CF_JOB
      CRViewer1.DisplayGroupTree = False
      CRViewer1.ViewReport
      CRViewer1.Zoom 100
      Me.Width = MotoRep.Width - 1000
      CRViewer1.Width = Me.Width - 500
      

End Sub

