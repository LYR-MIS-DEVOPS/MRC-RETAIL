VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmPrint 
   ClientHeight    =   5595
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6780
   LockControls    =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Tag             =   "xyf"
      Top             =   480
      Width           =   6735
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
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
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "Exit"
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
      Image           =   "frmPrint.frx":0000
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   240
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   7020
      m_MDI_Height    =   6165
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public QryPrn As String
Public MySource As String
Dim rsPrint As New ADODB.Recordset
Dim Ans As Integer
Public ParamValue1 As String
Public ParamValue2 As String
Public ParamValue3 As String
Public ParamValue4 As String
Public ParamValue5 As String
Public ParamValue6 As String
Public ParamValue7 As String
Public ParamValue8 As String
Public ParamValue9 As String
Public ParamValue10 As String
'******************* CODE UPDATED: 08/04/2017 ******************
Public ParamValue11 As String
Public ParamValue12 As String
Public ParamValue13 As String
Public ParamValue14 As String
Public ParamValue15 As String
Public ParamValue16 As String
'*************** END CODE UPDATED: 08/04/2017 ******************

Private Sub cmdExit_Click()
    Unload Me
    'findBBO.Show
End Sub


Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)
    If MySource = "Branch Order" Then
        LYRcon.Execute "UPDATE SPORDER SET SPORDER.PRINTCOPY = SPORDER.PRINTCOPY+1 " & _
                   "WHERE SPORDER.REFERENCE='" & BranchReplPO.txtPO.Text & "';"
    ElseIf MySource = "ADJUSTMENT REQUEST" Then
        If frmOtherAdjustment.txtRequestStatus.Text = "PENDING" Then
            Ans = MsgBox("Do you want to Print? This will Finalized your Request!", vbQuestion + vbYesNo, "Print")
            If Ans = vbNo Then Exit Sub
            Call frmOtherAdjustment.ProcessingAR
        End If
    End If
End Sub

Private Sub Form_Load()
    
    Select Case MySource
'        Case "PICKBBO"
'            Dim crPICKBBO As New rptPrntBBO
'
'            If rsPrint.State = adStateOpen Then rsPrint.Close
'            rsPrint.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
'            crPICKBBO.DiscardSavedData
'            crPICKBBO.Database.SetDataSource rsPrint
'            crPICKBBO.ParameterFields(1).ClearCurrentValueAndRange
'            crPICKBBO.ParameterFields(1).AddCurrentValue ParamValue1
'            CRViewer1.ReportSource = crPICKBBO
'            CRViewer1.ViewReport
'            CRViewer1.DisplayGroupTree = False
'            CRViewer1.EnableExportButton = False
'
        Case "RECEIVING DISCREPANCY PRINT"
            Dim crBrnchRecDis As New rptRecDis
            
            If rsPrint.State = adStateOpen Then rsPrint.Close
            rsPrint.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crBrnchRecDis.DiscardSavedData
            crBrnchRecDis.Database.SetDataSource rsPrint
            crBrnchRecDis.ParameterFields(1).ClearCurrentValueAndRange
            crBrnchRecDis.ParameterFields(1).AddCurrentValue ParamValue1
            crBrnchRecDis.ParameterFields(2).ClearCurrentValueAndRange
            crBrnchRecDis.ParameterFields(2).AddCurrentValue ParamValue2
            crBrnchRecDis.ParameterFields(3).ClearCurrentValueAndRange
            crBrnchRecDis.ParameterFields(3).AddCurrentValue ParamValue3
            CRViewer1.ReportSource = crBrnchRecDis
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
'********************** CODE UPDATED: 08/04/2017 ******************************
        Case "SETTLEMENT DISCREPANCY"
            Dim crSDiscrepancy As New rptSDiscrepancy
            
            If rsPrint.State = adStateOpen Then rsPrint.Close
            rsPrint.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crSDiscrepancy.DiscardSavedData
            crSDiscrepancy.Database.SetDataSource rsPrint
            crSDiscrepancy.ParameterFields(1).ClearCurrentValueAndRange
            crSDiscrepancy.ParameterFields(1).AddCurrentValue ParamValue1
            crSDiscrepancy.ParameterFields(2).ClearCurrentValueAndRange
            crSDiscrepancy.ParameterFields(2).AddCurrentValue ParamValue2
            crSDiscrepancy.ParameterFields(3).ClearCurrentValueAndRange
            crSDiscrepancy.ParameterFields(3).AddCurrentValue ParamValue3
            crSDiscrepancy.ParameterFields(4).ClearCurrentValueAndRange
            crSDiscrepancy.ParameterFields(4).AddCurrentValue ParamValue4
            crSDiscrepancy.ParameterFields(5).ClearCurrentValueAndRange
            crSDiscrepancy.ParameterFields(5).AddCurrentValue ParamValue5
            crSDiscrepancy.ParameterFields(6).ClearCurrentValueAndRange
            crSDiscrepancy.ParameterFields(6).AddCurrentValue ParamValue6
            crSDiscrepancy.ParameterFields(7).ClearCurrentValueAndRange
            crSDiscrepancy.ParameterFields(7).AddCurrentValue ParamValue7
            crSDiscrepancy.ParameterFields(8).ClearCurrentValueAndRange
            crSDiscrepancy.ParameterFields(8).AddCurrentValue ParamValue8
            crSDiscrepancy.ParameterFields(9).ClearCurrentValueAndRange
            crSDiscrepancy.ParameterFields(9).AddCurrentValue ParamValue9
            crSDiscrepancy.ParameterFields(10).ClearCurrentValueAndRange
            crSDiscrepancy.ParameterFields(10).AddCurrentValue ParamValue10
            crSDiscrepancy.ParameterFields(11).ClearCurrentValueAndRange
            crSDiscrepancy.ParameterFields(11).AddCurrentValue ParamValue11
            crSDiscrepancy.ParameterFields(12).ClearCurrentValueAndRange
            crSDiscrepancy.ParameterFields(12).AddCurrentValue ParamValue12
            crSDiscrepancy.ParameterFields(13).ClearCurrentValueAndRange
            crSDiscrepancy.ParameterFields(13).AddCurrentValue ParamValue13

            CRViewer1.ReportSource = crSDiscrepancy
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
'****************** END CODE UPDATED: 08/04/2017 ******************************
'********************** CODE UPDATED: 03/16/2018 ******************************
        Case "Branch Order"
            Dim crReplOrder As New rptPOReplOrd
            
            If rsPrint.State = adStateOpen Then rsPrint.Close
            rsPrint.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crReplOrder.DiscardSavedData
            crReplOrder.Database.SetDataSource rsPrint
            crReplOrder.ParameterFields(1).ClearCurrentValueAndRange
            crReplOrder.ParameterFields(1).AddCurrentValue ParamValue1
            crReplOrder.ParameterFields(2).ClearCurrentValueAndRange
            crReplOrder.ParameterFields(2).AddCurrentValue ParamValue2
            crReplOrder.ParameterFields(3).ClearCurrentValueAndRange
            crReplOrder.ParameterFields(3).AddCurrentValue ParamValue3
            crReplOrder.ParameterFields(4).ClearCurrentValueAndRange
            crReplOrder.ParameterFields(4).AddCurrentValue ParamValue4
            crReplOrder.ParameterFields(5).ClearCurrentValueAndRange
            crReplOrder.ParameterFields(5).AddCurrentValue ParamValue5
            crReplOrder.ParameterFields(6).ClearCurrentValueAndRange
            crReplOrder.ParameterFields(6).AddCurrentValue ParamValue6
            crReplOrder.ParameterFields(7).ClearCurrentValueAndRange
            crReplOrder.ParameterFields(7).AddCurrentValue ParamValue7
            crReplOrder.ParameterFields(8).ClearCurrentValueAndRange
            crReplOrder.ParameterFields(8).AddCurrentValue ParamValue8
            crReplOrder.ParameterFields(9).ClearCurrentValueAndRange
            crReplOrder.ParameterFields(9).AddCurrentValue ParamValue9
            crReplOrder.ParameterFields(10).ClearCurrentValueAndRange
            crReplOrder.ParameterFields(10).AddCurrentValue ParamValue10
            crReplOrder.ParameterFields(11).ClearCurrentValueAndRange
            crReplOrder.ParameterFields(11).AddCurrentValue ParamValue11
            CRViewer1.ReportSource = crReplOrder
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
'****************** END CODE UPDATED: 03/16/2018 ******************************
        Case "ADJUSTMENT REQUEST"
            Dim crADJUSTMENTREQ As New rptAdjustmentRequest
            
            If rsPrint.State = adStateOpen Then rsPrint.Close
            rsPrint.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crADJUSTMENTREQ.DiscardSavedData
            crADJUSTMENTREQ.Database.SetDataSource rsPrint
            crADJUSTMENTREQ.ParameterFields(1).ClearCurrentValueAndRange
            crADJUSTMENTREQ.ParameterFields(1).AddCurrentValue ParamValue1
            crADJUSTMENTREQ.ParameterFields(2).ClearCurrentValueAndRange
            crADJUSTMENTREQ.ParameterFields(2).AddCurrentValue ParamValue2
            crADJUSTMENTREQ.ParameterFields(3).ClearCurrentValueAndRange
            crADJUSTMENTREQ.ParameterFields(3).AddCurrentValue ParamValue3
            crADJUSTMENTREQ.ParameterFields(4).ClearCurrentValueAndRange
            crADJUSTMENTREQ.ParameterFields(4).AddCurrentValue ParamValue4
            CRViewer1.ReportSource = crADJUSTMENTREQ
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
    Case "ADJUSTMENT REQUEST SCRIPT"
            Dim crADJUSTMENTREQSCRIPT As New rptAdjustmentRequestProcessScript
            
            If rsPrint.State = adStateOpen Then rsPrint.Close
            rsPrint.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crADJUSTMENTREQSCRIPT.DiscardSavedData
            crADJUSTMENTREQSCRIPT.Database.SetDataSource rsPrint
            crADJUSTMENTREQSCRIPT.ParameterFields(1).ClearCurrentValueAndRange
            crADJUSTMENTREQSCRIPT.ParameterFields(1).AddCurrentValue ParamValue1
            crADJUSTMENTREQSCRIPT.ParameterFields(2).ClearCurrentValueAndRange
            crADJUSTMENTREQSCRIPT.ParameterFields(2).AddCurrentValue ParamValue2
            crADJUSTMENTREQSCRIPT.ParameterFields(3).ClearCurrentValueAndRange
            crADJUSTMENTREQSCRIPT.ParameterFields(3).AddCurrentValue ParamValue3
            crADJUSTMENTREQSCRIPT.ParameterFields(4).ClearCurrentValueAndRange
            crADJUSTMENTREQSCRIPT.ParameterFields(4).AddCurrentValue ParamValue4
            CRViewer1.ReportSource = crADJUSTMENTREQSCRIPT
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
        Case "CANCELLATION MODULE"
            Dim crCM As New rptCancellationModule
            
            If rsPrint.State = adStateOpen Then rsPrint.Close
            rsPrint.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crCM.DiscardSavedData
            crCM.Database.SetDataSource rsPrint
            crCM.ParameterFields(1).ClearCurrentValueAndRange
            crCM.ParameterFields(1).AddCurrentValue ParamValue1
            crCM.ParameterFields(2).ClearCurrentValueAndRange
            crCM.ParameterFields(2).AddCurrentValue ParamValue2
            crCM.ParameterFields(3).ClearCurrentValueAndRange
            crCM.ParameterFields(3).AddCurrentValue ParamValue3
            crCM.ParameterFields(4).ClearCurrentValueAndRange
            crCM.ParameterFields(4).AddCurrentValue ParamValue4
            CRViewer1.ReportSource = crCM
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MySource = "" Then
        findBBO.Show
    ElseIf MySource = "BRANCH RECEIVING DISCREPANCY PRINT" Then
        Exit Sub
    End If
End Sub
