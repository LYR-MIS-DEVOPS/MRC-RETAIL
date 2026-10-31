VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHSummary 
   Caption         =   "Summary"
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   15060
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar Pbar 
      Height          =   135
      Left            =   120
      TabIndex        =   22
      Tag             =   "xy"
      Top             =   9480
      Visible         =   0   'False
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxChk 
      Height          =   4695
      Left            =   16200
      TabIndex        =   20
      Top             =   1200
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   8281
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   13560
      Top             =   0
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   7560
      TabIndex        =   12
      Tag             =   "xyf"
      Top             =   480
      Width           =   7335
      Begin VB.TextBox txtFind 
         Height          =   375
         Left            =   5040
         TabIndex        =   14
         Tag             =   "xyf"
         Top             =   240
         Width           =   2175
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   390
         Left            =   3720
         TabIndex        =   13
         Tag             =   "xyf"
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   688
         Caption         =   "Find Next"
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
         Image           =   "frmHSummary.frx":0000
         cBack           =   -2147483633
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSum 
      Height          =   8175
      Left            =   120
      TabIndex        =   5
      Tag             =   "xyf"
      Top             =   1320
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   14420
      _Version        =   393216
      RowHeightMin    =   300
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Tag             =   "xyf"
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "E&xit Summary"
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
      Image           =   "frmHSummary.frx":0A12
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdShow 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Tag             =   "xyf"
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "Sho&w Form"
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
      Image           =   "frmHSummary.frx":1424
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Tag             =   "xyf"
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&Print Preview"
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
      Image           =   "frmHSummary.frx":19BE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdFilter 
      Height          =   390
      Left            =   3600
      TabIndex        =   3
      Tag             =   "xyf"
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   688
      Caption         =   "&Filter Records"
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
      Image           =   "frmHSummary.frx":23D0
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCheck 
      Height          =   390
      Left            =   14400
      TabIndex        =   21
      Tag             =   "xyf"
      Top             =   0
      Visible         =   0   'False
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   688
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
      ImgAlign        =   4
      Image           =   "frmHSummary.frx":2DE2
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Tag             =   "xyf"
      Top             =   480
      Visible         =   0   'False
      Width           =   7335
      Begin MSComCtl2.DTPicker Dfrom 
         Height          =   300
         Left            =   1440
         TabIndex        =   7
         Tag             =   "xyf"
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         Format          =   99090433
         CurrentDate     =   42107
      End
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   390
         Left            =   6120
         TabIndex        =   8
         Tag             =   "xyf"
         Top             =   195
         Visible         =   0   'False
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   688
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
         ImgAlign        =   4
         Image           =   "frmHSummary.frx":37F4
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker Dto 
         Height          =   300
         Left            =   3840
         TabIndex        =   9
         Tag             =   "xyf"
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         Format          =   98959361
         CurrentDate     =   42107
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   390
         Left            =   5707
         TabIndex        =   10
         Tag             =   "xyf"
         Top             =   195
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   688
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
         ImgAlign        =   4
         Image           =   "frmHSummary.frx":4206
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdPrev 
         Height          =   390
         Left            =   3360
         TabIndex        =   11
         Tag             =   "xyf"
         Top             =   195
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   688
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
         ImgAlign        =   4
         Image           =   "frmHSummary.frx":4C18
         cBack           =   -2147483633
      End
      Begin VB.Label Label1 
         Caption         =   "Date Covered:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Tag             =   "xyf"
         Top             =   285
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Tag             =   "XYF"
      Top             =   480
      Visible         =   0   'False
      Width           =   7335
      Begin VB.TextBox txtYear 
         Height          =   315
         Left            =   3135
         TabIndex        =   17
         Tag             =   "XYF"
         Top             =   240
         Width           =   720
      End
      Begin VB.ComboBox cboMonth 
         Height          =   315
         Left            =   1440
         TabIndex        =   16
         Tag             =   "XYF"
         Top             =   240
         Width           =   1695
      End
      Begin MSComCtl2.UpDown UpdYear 
         Height          =   315
         Left            =   3840
         TabIndex        =   18
         Tag             =   "XYF"
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Value           =   2015
         OrigLeft        =   11310
         OrigTop         =   480
         OrigRight       =   11565
         OrigBottom      =   795
         Max             =   2020
         Min             =   2005
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "For the month of"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   300
         Width           =   1215
      End
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   12840
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   15300
      m_MDI_Height    =   10380
      OriginalBorder  =   10
      OriginalCaption =   31
   End
   Begin VB.Menu mnuSummary 
      Caption         =   "Summary"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmHSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSummary As New ADODB.Recordset
Dim DailyRange As Boolean, MonthlyRange As Boolean, FixRange As Boolean
Dim ActRow As Integer, fRow As Integer, fCol As Integer, MaxCol As Integer, ShowDate As Integer
Public MySource As String


Private Sub cboMonth_Click()
    If cboMonth.ListIndex = -1 Then Exit Sub
    
    'Identify source whose monthly reports is from Day 1 to Last day of the month
    If mySrc = "CLOSED ACCOUNTS FOR AGENT" Or mySrc = "Monthly Collection" Or mySrc = "Monthly Inventory" Then
        Dfrom.Value = DateSerial(Val(txtYear.Text), cboMonth.ListIndex + 1, 1)
        Dto.Value = DateAdd("M", 1, Dfrom.Value)
        Dto.Value = DateAdd("D", -1, Dto.Value)
    Else    'Use for reports whose date covered is from 7th day to 6th of the following month
        Dfrom.Value = DateSerial(Val(txtYear.Text), cboMonth.ListIndex + 1, 7)
        Dto.Value = DateSerial(Year(DateAdd("M", 1, Dfrom.Value)), Month(DateAdd("M", 1, Dfrom.Value)), 6)
    End If
    Call cmdRefresh_Click
End Sub

Private Sub cmdCheck_Click()
    Dim rsChk As New ADODB.Recordset
    Dim Qry As String, i As Integer, j As Integer
    
    If MySource = "BRANCH BACK ORDER" Then
        Qry = "SELECT backorder.REFERENCE, backorder.BRANCHSPOID, sporderbranch.BRANCHNAME, sporderbranch.branchID, backorder.SPAREPARTID " & _
              "From backorder INNER JOIN sporderbranch ON (backorder.REFERENCE = sporderbranch.REFERENCE) AND (backorder.BRANCHSPOID = sporderbranch.BRANCHSPOID) Order By backorder.TRANSDATE"
        'Display Query in Flx
        If rsChk.State = adStateOpen Then rsChk.Close
        rsChk.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If rsChk.RecordCount = 0 Then   'No records to show
            flxChk.Clear
            flxChk.Rows = 2
            flxChk.Cols = MaxCol
            Call RXResizer1_ResizeComplete
            Exit Sub
        End If
        Set flxChk.DataSource = rsChk   'With records
        'Find Branch in flxChk and display in flxsum
        Pbar.Value = 1
        Pbar.Max = flxSum.Rows
        Pbar.Visible = True
        For i = 1 To flxSum.Rows - 1
            For j = 1 To flxChk.Rows - 1
                If flxSum.TextMatrix(i, 2) = flxChk.TextMatrix(j, 1) And flxSum.TextMatrix(i, 10) = flxChk.TextMatrix(j, 2) And flxSum.TextMatrix(i, 9) = flxChk.TextMatrix(j, 5) Then
                    flxSum.TextMatrix(i, 7) = flxChk.TextMatrix(j, 3)
                    LYRcon.Execute "UPDATE backorder SET BRANCHNAME = '" & flxChk.TextMatrix(j, 3) & "', branchID = " & Val(flxChk.TextMatrix(j, 4)) & " " & _
                        "Where backorder.BRANCHSPOID = " & Val(flxSum.TextMatrix(i, 10)) & " AND backorder.SPAREPARTID = " & Val(flxSum.TextMatrix(i, 9)) & " AND backorder.REFERENCE = '" & flxSum.TextMatrix(i, 2) & "'"
                    Exit For
                End If
            Next j
            Pbar.Value = Pbar.Value + 1
            DoEvents
        Next i
        Pbar.Visible = False
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxSum
        For i = fRow To .Rows - 1
            For j = fCol To MaxCol - 1
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = MaxCol - 1: .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = MaxCol - 1, i + 1, i)
                    fCol = IIf(j = MaxCol - 1, 1, j + 1)
                    ActRow = i
                    Exit Sub
                End If
            Next j
            fCol = 1
        Next i
    End With
    MsgBox "No match found!", vbInformation + vbOKOnly, "Find"
    fRow = 1
    fCol = 1

End Sub

Private Sub cmdNext_Click()
    cmdNext.Visible = False
    cmdPrev.Visible = True
    Dto.Visible = True
    ShowDate = 2
End Sub

Private Sub cmdPrev_Click()
    Dto.Value = Dfrom.Value
    cmdNext.Left = cmdPrev.Left
    cmdNext.Visible = True
    cmdPrev.Visible = False
    Dto.Visible = False
    ShowDate = 1
End Sub

Private Sub cmdPrint_Click()
    Dim Qry As String
    
    
    If MySource = "BRANCH BACK ORDER" Then
        Qry = "SELECT DATE_FORMAT(backorder.TRANSDATE,'%c/%d/%Y'), backorder.REFERENCE, spareparts.PARTNUMBER, spareparts.DESCRIPTION, backorder.BACKORDERQTY, backorder.BOBALANCE, backorder.BRANCHNAME, backorder.SUPPORDERREF " & _
              "From backorder INNER JOIN spareparts ON (backorder.SPAREPARTID = spareparts.SPAREPARTID) " & _
              "WHERE backorder.TRANSDATE >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' AND backorder.TRANSDATE <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' ORDER BY backorder.TRANSDATE DESC"
    
        frmReport.MySource = "BRANCH BACK ORDER"
        frmReport.QryPrn = Qry
        frmReport.ParamValue1 = "MOTO-R REPUBLIC CORPORATION"
        frmReport.ParamValue2 = "SOBRECARY ST. TAGUM CITY"
        frmReport.Show vbModal
        
    End If
End Sub

Private Sub cmdRefresh_Click()
    Dim Qry As String
    
    
    If MySource = "PRODUCTS" Then
        Qry = "SELECT products.BRAND, products.MODEL, products.DESCRIPTION, products.CLASSIFICATION, products.PRODTYPE, products.PRODINIT, products.ACCTITLE, products.PRODID " & _
            "From products Order By products.BRAND, products.MODEL"
    
    ElseIf MySource = "BRANCH BACK ORDER" Then
        Qry = "SELECT DATE_FORMAT(backorder.TRANSDATE,'%c/%d/%Y'), backorder.REFERENCE, spareparts.PARTNUMBER, spareparts.DESCRIPTION, backorder.BACKORDERQTY, backorder.BOBALANCE, backorder.BRANCHNAME, backorder.SUPPORDERREF, backorder.SPAREPARTID, backorder.BRANCHSPOID " & _
              "From backorder INNER JOIN spareparts ON (backorder.SPAREPARTID = spareparts.SPAREPARTID) " & _
              "WHERE backorder.TRANSDATE >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' AND backorder.TRANSDATE <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' ORDER BY backorder.TRANSDATE DESC"
    ElseIf MySource = "CLOSED ACCOUNTS FOR AGENT" Then
        Qry = "SELECT CONCAT(agents.LASTNAME, ', ', agents.FIRSTNAME) AS AGENT, agentorders.RANK, DATE_FORMAT(unitordered.DATECLOSED,'%c/%d/%Y'), unitordered.ACCOUNTNO, CONCAT(customer.LASTNAME, ', ', customer.FIRSTNAME) AS customername, CONCAT(customer.ADDRESS, ' ', customer.TOWN) AS custAddress, concat(productline.BRAND, ' ', productline.DESCRIPTION, ' ', unitordered.MODEL) AS UnitDesc " & _
              "From agingprev INNER JOIN unitordered ON (agingprev.ORDERID = unitordered.ORDERID) INNER JOIN agentorders ON (agingprev.ORDERID = agentorders.ORDERID) INNER JOIN agents ON (agentorders.AGENTID = agents.AGENTID) INNER JOIN customer ON (unitordered.CUSTOMERID = customer.CUSTOMERID) INNER JOIN productline ON (unitordered.PRODID = productline.PRODID) " & _
              "Where (agentorders.RANK = 'USM' OR agentorders.RANK = 'USS') AND unitordered.DATECLOSED >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' AND unitordered.DATECLOSED < '" & Format(Dto.Value, "YYYY-MM-DD") & "' AND agingprev.REMARKS = 'CLOSED' AND agingprev.`STATUS` = 'GOOD' Order By agentorders.RANK, AGENT"
    
    ElseIf MySource = "MRC WAREHOUSE BACK ORDER" Then
        Qry = "SELECT date_format(spitr.TRANSDATE, '%M %d, %Y'), spitr.REFERENCE, spareparts.PARTNUMBER, spareparts.DESCRIPTION,  spitr.`ORDER`, spitr.RECEIVE,  spitr.BALANCE, spitr.SPAREPARTID " & _
                "From spareparts INNER JOIN spitr ON (spareparts.SPAREPARTID = spitr.SPAREPARTID) " & _
                "Where spitr.TRANSTYPE = 'order' AND spitr.BALANCE > 0 AND spitr.TRANSDATE >= '2015-05-01';"
    End If
    
    'Display Query in Flx
    If rsSummary.State = adStateOpen Then rsSummary.Close
    rsSummary.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsSummary.RecordCount = 0 Then   'No records to show
        flxSum.Clear
        flxSum.Rows = 2
        flxSum.Cols = MaxCol
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxSum.DataSource = rsSummary   'With records
    ActRow = 1
    Call RXResizer1_ResizeComplete
          
End Sub

Private Sub cmdShow_Click()
    If MySource = "PRODUCTS" Then
        Call frmproductEnt.ShowRecord(Val(flxSum.TextMatrix(ActRow, 8)))
        frmproductEnt.Show vbModal
    End If
End Sub



Private Sub Dfrom_Change()
    If ShowDate = 1 Then Dto.Value = Dfrom.Value
End Sub

Private Sub flxSum_Click()
    ActRow = flxSum.Row
End Sub

Private Sub flxSum_EnterCell()
    ActRow = flxSum.Row
End Sub

Private Sub Form_Load()

    Me.Caption = MySource & " SUMMARY"
    Timer1.Enabled = True
End Sub

Private Sub RXResizer1_ResizeComplete()
    With flxSum
        If MySource = "PRODUCTS" Then
            .FormatString = "   |^BRAND|<MODEL|<DESCRIPTION|<CLASSIFICATION|<PRODUCT TYPE|<PROD. INITIAL|<ACCOUNT TITLE"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.12
            .ColWidth(2) = (.Width - 300) * 0.12
            .ColWidth(3) = (.Width - 300) * 0.2
            .ColWidth(4) = (.Width - 300) * 0.15
            .ColWidth(5) = (.Width - 300) * 0.12
            .ColWidth(6) = (.Width - 300) * 0.08
            .ColWidth(7) = (.Width - 300) * 0.18
            .ColWidth(8) = 0
        ElseIf MySource = "BRANCH BACK ORDER" Then
           .FormatString = " |^DATE|^REF.NO.|<PARTNUMBER|<DESCRIPTION|^BBO QTY.|^BBO BAL.|<BRANCH |<P.O. REF#. "
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.08
            .ColWidth(2) = (.Width - 300) * 0.08
            .ColWidth(3) = (.Width - 300) * 0.18
            .ColWidth(4) = (.Width - 300) * 0.2
            .ColWidth(5) = (.Width - 300) * 0.08
            .ColWidth(6) = (.Width - 300) * 0.08
            .ColWidth(7) = (.Width - 300) * 0.15
            .ColWidth(8) = (.Width - 300) * 0.12
            .ColWidth(9) = 0
            .ColWidth(10) = 0
        ElseIf MySource = "CLOSED ACCOUNTS FOR AGENT" Then
            .FormatString = "  |<NAME OF AGENT|^RANK|^DATE CLOSED|^ACCOUNTNO|<CUSTOMER NAME|<ADDRESS|<UNIT DESCRIPTION"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.15
            .ColWidth(2) = (.Width - 300) * 0.08
            .ColWidth(3) = (.Width - 300) * 0.08
            .ColWidth(4) = (.Width - 300) * 0.1
            .ColWidth(5) = (.Width - 300) * 0.18
            .ColWidth(6) = (.Width - 300) * 0.18
            .ColWidth(7) = (.Width - 300) * 0.18
        ElseIf MySource = "MRC WAREHOUSE BACK ORDER" Then
            .FormatString = " |^DATE|^REF. NO.|<PARTNUMBER|<DESCRIPTION|^ORDER QTY.|^RECEIVE QTY.|^B.O. QTY"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.08
            .ColWidth(2) = (.Width - 300) * 0.08
            .ColWidth(3) = (.Width - 300) * 0.18
            .ColWidth(4) = (.Width - 300) * 0.2
            .ColWidth(5) = (.Width - 300) * 0.08
            .ColWidth(6) = (.Width - 300) * 0.08
            .ColWidth(7) = (.Width - 300) * 0.15
            
            .ColWidth(8) = 0
            .ColWidth(9) = 0
        End If
    End With
    
End Sub

Private Sub Timer1_Timer()
    Dim i As Integer
    
    If MySource = "PRODUCTS" Then
        FixRange = True
        DailyRange = False
        MonthlyRange = False
        MaxCol = 8
    ElseIf MySource = "BRANCH BACK ORDER" Then
        FixRange = False
        DailyRange = True
        MonthlyRange = False
        MaxCol = 9
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto.Value = Date
        ShowDate = 2
        cmdCheck.Visible = True
    ElseIf MySource = "CLOSED ACCOUNTS FOR AGENT" Then
        FixRange = False
        DailyRange = False
        MonthlyRange = True
        MaxCol = 7
    ElseIf MySource = "MRC WAREHOUSE BACK ORDER" Then
        FixRange = False
        DailyRange = False
        MonthlyRange = True
        MaxCol = 10
    End If
    
    
        
    If DailyRange Then
        Frame1.Visible = True
        Frame3.Visible = False
    ElseIf MonthlyRange Then
        Frame1.Visible = False
        Frame3.Visible = True
        For i = 1 To 12
            cboMonth.AddItem MonthName(i)
        Next i
        UpdYear.Value = Year(Date)
        txtYear.Text = Year(Date)
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto.Value = DateAdd("d", -1, DateAdd("m", 1, Dfrom.Value))
        cboMonth.ListIndex = Month(Dfrom.Value) - 1
    ElseIf FixRange Then
        Frame1.Visible = False
        Frame3.Visible = False
    End If
    Call cmdRefresh_Click
    Timer1.Enabled = False
End Sub
