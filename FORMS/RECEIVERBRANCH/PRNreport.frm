VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form PRNmc 
   Caption         =   "Packing List Print Preview"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   735
   ClientWidth     =   8025
   Icon            =   "PRNreport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6345
   ScaleWidth      =   8025
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1560
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PRNreport.frx":6852
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport cRPT 
      Left            =   840
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7815
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
      Image           =   "PRNreport.frx":9004
      cBack           =   -2147483633
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "Print"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "PRNmc"
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
    Dim i As Integer, Ttl As Integer
    
    Unload Me
    SPincoming.Visible = True
End Sub

Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)
    
    If MySource = "SPPO" Or MySource = "SPPOrepl" Then
        SPordering.txtPrnCopy.Text = Val(SPordering.txtPrnCopy.Text) + 1
        'LYRcon.Execute "UPDATE SPORDER SET SPORDER.PRINTCOPY = [SPORDER].[PRINTCOPY]+1 " & _
                       "WHERE (((SPORDER.REFERENCE)='" & SPordering.txtPONumber.Text & "'));"
        'mySQL
        LYRcon.Execute "UPDATE SPORDER SET SPORDER.PRINTCOPY = SPORDER.PRINTCOPY+1 " & _
                       "WHERE SPORDER.REFERENCE='" & SPordering.txtPONumber.Text & "';"
        
    ElseIf MySource = "SERVICEDAILYINCOME" Then
        SRVCincome.txtPrintCopy.Text = Val(SRVCincome.txtPrintCopy.Text) + 1
        'LYRcon.Execute "UPDATE SERVICEINCOME SET SERVICEINCOME.PRINTCOPY = [SERVICEINCOME].[PRINTCOPY]+1 " & _
                       "WHERE (((SERVICEINCOME.SRVCINCOMEID)=" & SRVCincome.txtSrvcID.Text & "));"
        'mySQL
        LYRcon.Execute "UPDATE SERVICEINCOME SET SERVICEINCOME.PRINTCOPY = SERVICEINCOME.PRINTCOPY+1 " & _
                       "WHERE SERVICEINCOME.SRVCINCOMEID=" & SRVCincome.txtSrvcID.Text & ";"
        
    ElseIf MySource = "SERVICEDAILYINCOMENEW" Then
        SRVCincomeNew.txtPrintCopy.Text = Val(SRVCincomeNew.txtPrintCopy.Text) + 1
        'LYRcon.Execute "UPDATE SERVICEINCOME SET SERVICEINCOME.PRINTCOPY = [SERVICEINCOME].[PRINTCOPY]+1 " & _
                       "WHERE (((SERVICEINCOME.SRVCINCOMEID)=" & SRVCincomeNew.txtSrvcID.Text & "));"
        'mySQL
        LYRcon.Execute "UPDATE SERVICEINCOME SET SERVICEINCOME.PRINTCOPY = SERVICEINCOME.PRINTCOPY+1 " & _
                       "WHERE SERVICEINCOME.SRVCINCOMEID=" & SRVCincomeNew.txtSrvcID.Text & ";"
        
    ElseIf MySource = "SPRECEIVE" Then
        Call SPincoming.FinalSPrec
    End If
End Sub

Private Sub Form_Load()
'    On Error Resume Next
    
    Select Case MySource
        
        Case "SPRECEIVE"
            Dim crSPRec As New rptPackingList
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crSPRec.DiscardSavedData
            crSPRec.Database.SetDataSource rsPrn
            crSPRec.ParameterFields(1).ClearCurrentValueAndRange
            crSPRec.ParameterFields(1).AddCurrentValue ParamValue1
            crSPRec.ParameterFields(2).ClearCurrentValueAndRange
            crSPRec.ParameterFields(2).AddCurrentValue ParamValue2
            crSPRec.ParameterFields(3).ClearCurrentValueAndRange
            crSPRec.ParameterFields(3).AddCurrentValue ParamValue3
            crSPRec.ParameterFields(4).ClearCurrentValueAndRange
            crSPRec.ParameterFields(4).AddCurrentValue ParamValue4
            crSPRec.ParameterFields(5).ClearCurrentValueAndRange
            crSPRec.ParameterFields(5).AddCurrentValue ParamValue5
            crSPRec.ParameterFields(6).ClearCurrentValueAndRange
            crSPRec.ParameterFields(6).AddCurrentValue ParamValue6
            crSPRec.ParameterFields(7).ClearCurrentValueAndRange
            crSPRec.ParameterFields(7).AddCurrentValue ParamValue7
            crSPRec.ParameterFields(8).ClearCurrentValueAndRange
            crSPRec.ParameterFields(8).AddCurrentValue ParamValue8
            CRViewer1.ReportSource = crSPRec
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
            
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    QryPrn = ""
    MySource = ""
    If rsPrn.State = adStateOpen Then rsPrn.Close
End Sub

Private Sub Form_Resize()
    With CRViewer1
        .Top = cmdExit.Height + 50
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight '- MotoRep.sBar.Height
    End With
End Sub




