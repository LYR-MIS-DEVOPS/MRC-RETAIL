VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form DailyCash 
   Caption         =   "Daily Cash Collection"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   11520
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid FlxChkAlloc 
      Height          =   1300
      Left            =   11640
      TabIndex        =   16
      Top             =   3480
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2302
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      RowHeightMin    =   300
      AllowUserResizing=   1
      FormatString    =   "ORNO|<CHKPRID|<SPSORDERID|<BRANCH|<INVOICE|<CUSTOMERID"
   End
   Begin MSFlexGridLib.MSFlexGrid flxInvCancel 
      Height          =   735
      Left            =   840
      TabIndex        =   11
      Top             =   8280
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1296
      _Version        =   393216
      FixedCols       =   0
      FormatString    =   "REFERENCE NO.|^REF. TYPE"
   End
   Begin MSFlexGridLib.MSFlexGrid flxDC 
      Height          =   6975
      Left            =   0
      TabIndex        =   6
      Tag             =   "XY"
      Top             =   1200
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   12303
      _Version        =   393216
      Cols            =   15
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   16777215
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   $"DailyCollection.frx":0000
   End
   Begin CRVIEWERLibCtl.CRViewer vwrDaily 
      Height          =   7695
      Left            =   0
      TabIndex        =   3
      Tag             =   "XY"
      Top             =   480
      Width           =   11415
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
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Tag             =   "XY"
      Top             =   400
      Width           =   11295
      Begin MSComCtl2.DTPicker dFrom 
         Height          =   315
         Left            =   1920
         TabIndex        =   5
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   103284737
         CurrentDate     =   39055
      End
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   375
         Left            =   3720
         TabIndex        =   12
         ToolTipText     =   "Execute"
         Top             =   210
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         Image           =   "DailyCollection.frx":0097
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   375
         Left            =   5880
         TabIndex        =   13
         ToolTipText     =   "Execute"
         Top             =   210
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         Image           =   "DailyCollection.frx":0631
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dTo 
         Height          =   315
         Left            =   4200
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   103284737
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdPrev 
         Height          =   375
         Left            =   6360
         TabIndex        =   15
         ToolTipText     =   "Execute"
         Top             =   210
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         Image           =   "DailyCollection.frx":0BCB
         cBack           =   -2147483633
      End
      Begin VB.Label lblHeading 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Select Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   2
         Left            =   -720
         TabIndex        =   1
         Top             =   285
         Width           =   2505
      End
      Begin VB.Label lblHeading 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Select Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   -720
         TabIndex        =   2
         Top             =   315
         Width           =   2505
      End
   End
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4560
      Top             =   8280
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4560
      Top             =   7800
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   8280
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   7800
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "DailyCollection.frx":1165
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "DailyCollection.frx":16FF
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Grid View"
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
      Image           =   "DailyCollection.frx":1C99
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   10
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "DailyCollection.frx":2233
      cBack           =   -2147483633
   End
   Begin VB.Label lblP 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   8280
      Visible         =   0   'False
      Width           =   255
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   480
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11760
      m_MDI_Height    =   8745
   End
   Begin VB.Menu mnuDaily 
      Caption         =   "Daily Cash"
      Visible         =   0   'False
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel O.R."
      End
      Begin VB.Menu mnuAttach 
         Caption         =   "Show Attachment"
      End
   End
End
Attribute VB_Name = "DailyCash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRPT As New ADODB.Recordset

Dim BtnIndx As Integer
Dim ActRow As Integer
Dim MQry As String, ThisBranch As String
Dim prevDCode As Integer

Private Sub cmdNext_Click()
    cmdNext.Tag = "2"
    cmdNext.Visible = False
    cmdPrev.Left = cmdNext.Left
    cmdPrev.Top = cmdNext.Top
    cmdPrev.Visible = True
    dTo.Visible = True
    cmdRefresh.Visible = True
    dTo.Value = dFrom.Value
End Sub

Private Sub cmdPrev_Click()
    cmdNext.Tag = "1"
    cmdNext.Visible = True
    cmdPrev.Visible = False
    dTo.Visible = False
    cmdRefresh.Visible = False
    Call RefreshRecord(1)
End Sub

Private Sub cmdRefresh_Click()
    Call RefreshRecord(2)
End Sub

Private Sub cmdTrans_Click(Index As Integer)
'On Error GoTo er
    Dim Qry As String, strDate As String
    
    strDate = Format$(Now(), "hh:mm AM/PM")
    
    Select Case Index
        Case 0  'Grid View
            BtnIndx = 1
            tmrFloat.Enabled = True
            Call PrintPreView(False)
        Case 1  'Exit
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
            Unload Me
        Case 2  'Print Preview
            If SysBranch <> "MOTORHUB SPAREPARTS TAGUM" And SysBranch <> "BINOS POLOMOLOK SPAREPARTS" And SysBranch <> "BINOS SPAREPARTS STO. TOMAS" And SysBranch <> "BINOS SPAREPARTS SAMAL" And SysBranch <> "MOTORHUB SPAREPARTS MARBEL" And SysBranch <> "BINOS SPAREPARTS TAGUM" And SysBranch <> "YAMAHA MOTORHUB ULAS" Then
                If Val(cmdNext.Tag) = 1 And Date = dFrom.Value Then
                    Qry = "SELECT * From payments Where payments.UPDATEDBY = " & UserID & " AND payments.PAYDATE = '" & Format(Date, "yyyy-mm-dd") & "';"
                    If CtrlNumExist(Qry) Then
                        MsgBox "You are restricted to view this report!" & vbCrLf & "Please contact your OIC for further notice.", vbExclamation + vbOKOnly, "Warning"
                       Exit Sub
                    End If
                ElseIf Val(cmdNext.Tag) = 2 And Date >= dFrom.Value And Date <= dTo.Value Then
                    Qry = "SELECT * From payments Where payments.UPDATEDBY = " & UserID & " AND payments.PAYDATE = '" & Format(Date, "yyyy-mm-dd") & "';"
                    If CtrlNumExist(Qry) Then
                        MsgBox "You are restricted to view this report!" & vbCrLf & "Please contact your OIC for further notice.", vbExclamation + vbOKOnly, "Warning"
                        Exit Sub
                    End If
                End If
            Else
                If strDate >= "04:00 PM" Then
                    If Val(cmdNext.Tag) = 1 And Date = dFrom.Value Then
                        Qry = "SELECT * From payments Where payments.UPDATEDBY = " & UserID & " AND payments.PAYDATE = '" & Format(Date, "yyyy-mm-dd") & "';"
                        If CtrlNumExist(Qry) Then
                            MsgBox "You are restricted to view this report!" & vbCrLf & "Please contact your OIC for further notice.", vbExclamation + vbOKOnly, "Warning"
                           Exit Sub
                        End If
                    ElseIf Val(cmdNext.Tag) = 2 And Date >= dFrom.Value And Date <= dTo.Value Then
                        Qry = "SELECT * From payments Where payments.UPDATEDBY = " & UserID & " AND payments.PAYDATE = '" & Format(Date, "yyyy-mm-dd") & "';"
                        If CtrlNumExist(Qry) Then
                            MsgBox "You are restricted to view this report!" & vbCrLf & "Please contact your OIC for further notice.", vbExclamation + vbOKOnly, "Warning"
                            Exit Sub
                        End If
                    End If
                Else
                    If Val(cmdNext.Tag) = 1 And Date = dFrom.Value Then
                        Qry = "SELECT * From payments Where payments.UPDATEDBY = " & UserID & " AND payments.PAYDATE = '" & Format(Date, "yyyy-mm-dd") & "';"
                        'If CtrlNumExist(Qry) Then
                        '    MsgBox "You are restricted to view this report!" & vbCrLf & "Please contact your OIC for further notice.", vbExclamation + vbOKOnly, "Warning"
                        '   Exit Sub
                        'End If
                    ElseIf Val(cmdNext.Tag) = 2 And Date >= dFrom.Value And Date <= dTo.Value Then
                        Qry = "SELECT * From payments Where payments.UPDATEDBY = " & UserID & " AND payments.PAYDATE = '" & Format(Date, "yyyy-mm-dd") & "';"
                        'If CtrlNumExist(Qry) Then
                        '    MsgBox "You are restricted to view this report!" & vbCrLf & "Please contact your OIC for further notice.", vbExclamation + vbOKOnly, "Warning"
                        '    Exit Sub
                        'End If
                    End If
                End If
            End If
            
            BtnIndx = 3
            tmrFloatRev.Enabled = True
            Call DisplayReport
            Call PrintPreView(True)
        Case 3  'Exit
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
            Unload Me
    End Select
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->cmdTrans_Click()"
End Sub

Private Sub RefreshRecord(dCode As Integer)
On Error GoTo er
    prevDCode = dCode
    Dim i As Integer
    
    If dCode = 1 Then
        'MYsq
        MQry = "SELECT PAYMENTS.INVOICENO, IF(PAYMENTS.CRNO<>'', PAYMENTS.CRNO, PAYMENTS.ORNO) as RECEIPT, PAYMENTS.CUSTOMERNAME, PAYMENTS.DOWNPAYMENT, PAYMENTS.SPAREPARTS, PAYMENTS.SERVICE, PAYMENTS.OTHERS, PAYMENTS.REMARKS, concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.'))) AS CASHIER, PAYMENTS.PAYDATE, PAYMENTS.CRNO, PAYMENTS.PAYMENTID,ifnull(PAYMENTS.PAYTYPE,'') as PAYTYPE,round(ifnull(allocation.TAXWHELDAMOUNT,0),2) as TAXWHELDAMOUNT  " & _
               "From Employees INNER JOIN PAYMENTS ON (EMPLOYEES.EMPID = PAYMENTS.UPDATEDBY) " & _
               " Left Join  allocation on payments.ALLOCATIONID = allocation.ALLOCATIONID And ifnull(payments.PAYTYPE,'') <> '' " & _
               " WHERE PAYMENTS.PAYDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' Order By CONCAT(EMPLOYEES.LASTNAME, ', ' , EMPLOYEES.FIRSTNAME, SPACE(1), IF(ISNULL(EMPLOYEES.MIDDLENAME), ' ', LEFT(EMPLOYEES.MIDDLENAME,1)), IF(ISNULL(EMPLOYEES.MIDDLENAME), ' ','.')), PAYMENTS.ORNO, PAYMENTS.CRNO, PAYMENTS.INVOICENO;"
    Else
        'MQry = "SELECT PAYMENTS.INVOICENO, PAYMENTS.ORNO, PAYMENTS.CUSTOMERNAME, PAYMENTS.DOWNPAYMENT, PAYMENTS.SPAREPARTS, PAYMENTS.SERVICE, PAYMENTS.OTHERS, PAYMENTS.REMARKS, [LASTNAME] & ' ' & Left$([FIRSTNAME],1) & '.' AS CASHIER, PAYMENTS.PAYDATE " & _
              "FROM EMPLOYEES INNER JOIN PAYMENTS ON EMPLOYEES.EMPID = PAYMENTS.UPDATEDBY Where (((PAYMENTS.PAYDATE) >= #" & dFrom.Value & "# And (PAYMENTS.PAYDATE) <= #" & dTo.Value & "#)) ORDER BY PAYMENTS.PAYDATE, [LASTNAME] & ' ' & Left$([FIRSTNAME],1) & '.',  PAYMENTS.ORNO, PAYMENTS.INVOICENO;"
       'mySQL
       MQry = "SELECT PAYMENTS.INVOICENO, IF(PAYMENTS.CRNO<>'', PAYMENTS.CRNO, PAYMENTS.ORNO) as RECEIPT, PAYMENTS.CUSTOMERNAME, PAYMENTS.DOWNPAYMENT, PAYMENTS.SPAREPARTS, PAYMENTS.SERVICE, PAYMENTS.OTHERS, PAYMENTS.REMARKS, concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.'))) AS CASHIER, PAYMENTS.PAYDATE, PAYMENTS.CRNO, PAYMENTS.PAYMENTID,ifnull(PAYMENTS.PAYTYPE,'') as PAYTYPE,round(ifnull(allocation.TAXWHELDAMOUNT,0),2) as TAXWHELDAMOUNT " & _
              " FROM EMPLOYEES INNER JOIN PAYMENTS ON EMPLOYEES.EMPID = PAYMENTS.UPDATEDBY " & _
              " Left Join  allocation on payments.ALLOCATIONID = allocation.ALLOCATIONID And ifnull(payments.PAYTYPE,'') <> '' " & _
              " Where PAYMENTS.PAYDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And PAYMENTS.PAYDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' Order By CONCAT(EMPLOYEES.LASTNAME, ', ' , EMPLOYEES.FIRSTNAME, SPACE(1), IF(ISNULL(EMPLOYEES.MIDDLENAME), ' ', LEFT(EMPLOYEES.MIDDLENAME,1)), IF(ISNULL(EMPLOYEES.MIDDLENAME), ' ','.')), PAYMENTS.ORNO, PAYMENTS.CRNO, PAYMENTS.INVOICENO;"
    End If
    If rsRPT.State = adStateOpen Then rsRPT.Close
    
    ' 0     1              2             3                4                5               6          7         8             9           10         11       12     13     14
    '   |^  Invoice No.  |^Receipt No. |<Customer Name  |> Down Payment  |>Spare Parts   |>Service  |>Others  |>TAX w/HELD  |<Remarks   |<Cashier   |<Date  | 12   | 13   | PAYTYPE
    With rsRPT
        .Open MQry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            flxDC.Rows = 1
            Do While Not .EOF
                '              0          1                  2                   3                                             4                                           5                                         6                                         7                      8                          9                  10               11                                        12                  13                14
                flxDC.AddItem vbTab & !InvoiceNO & vbTab & !Receipt & vbTab & !CustomerName & vbTab & Format$(!DownPayment, "###,##0.00") & vbTab & Format$(!SPAREPARTS, "###,##0.00") & vbTab & Format$(!Service, "###,##0.00") & vbTab & Format$(!Others, "###,##0.00") & vbTab & !TAXWHELDAMOUNT & vbTab & !Remarks & vbTab & !Cashier & vbTab & Format$(!PayDate, "mm/dd/yyyy") & vbTab & !CRNO & vbTab & !PaymentID & vbTab & !PayType
                .MoveNext
            Loop
        Else
            flxDC.Rows = 1
            flxDC.Rows = 2
        End If
        .Close
    End With

Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->RefreshRecord()"
End Sub

Private Sub dFrom_CloseUp()
    If cmdNext.Tag = "1" Then
        Call RefreshRecord(1)
    End If
End Sub

Private Sub dTo_CloseUp()
    If dTo.Value < dFrom.Value Then
        MsgBox "Please choose a date not less than " & dFrom.Value, vbExclamation + vbOKOnly, "Validation Error"
        dTo.Value = dFrom.Value
    End If
End Sub

Private Sub flxdc_Click()
    If flxDC.Col = 0 And flxDC.Row > 0 Then
        ActRow = flxDC.Row
        Call flxdc_EnterCell
    Else
        ActRow = flxDC.Row
    End If
End Sub

Private Sub flxdc_EnterCell()
    ActRow = flxDC.Row
    With flxDC
        .Row = ActRow: .Col = 0
        .RowSel = ActRow: .ColSel = .Cols - 1
        If Not .RowIsVisible(.Row) Then .TopRow = .Row
    End With
End Sub

Private Sub flxdc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo er
    If ActRow = 0 Then
        ActRow = flxDC.Row
    End If

    If ActRow <= 0 Then
        ActRow = 0
        MsgBox "Please select a record!", vbCritical, "Error->mnuCancel_Click()"
        Exit Sub
        
    ElseIf ActRow > flxDC.Rows Then
        ActRow = 0
        MsgBox "Please select a record!", vbCritical, "Error->mnuCancel_Click()"
        Exit Sub
    End If
    
    If flxDC.TextMatrix(ActRow, 2) = "" Or InStr(UCase$(flxDC.TextMatrix(ActRow, 9)), "CANCELLED") > 0 Then Exit Sub    'No O.R. or Cancelled O.R.
    If DateValue(dFrom.Value) <> Date Then
        If InStr(UCase$(flxDC.TextMatrix(ActRow, 9)), "ATTACHMENT") = 0 And InStr(UCase$(flxDC.TextMatrix(ActRow, 9)), "CHECK") = 0 Then Exit Sub
    End If
    If Button = vbRightButton Then
        If DateValue(dFrom.Value) <> Date Then
            mnuCancel.Enabled = False
        Else
            mnuCancel.Enabled = True
        End If
        If InStr(UCase$(flxDC.TextMatrix(ActRow, 9)), "ATTACHMENT") > 0 Or InStr(UCase$(flxDC.TextMatrix(ActRow, 9)), "CHECK") > 0 Then
            mnuAttach.Enabled = True
        Else
            mnuAttach.Enabled = False
        End If
        Call PopupMenu(mnuDaily)
    End If
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->flxdc_MouseDown()"
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    dFrom.Value = Date
    cmdNext.Tag = "1"
    Call RefreshRecord(1)
    BtnIndx = 1
    tmrFloat.Enabled = True
    Call PrintPreView(False)
    
    If SysBranch = "MOTO-R REPUBLIC CORP." Then
        ThisBranch = "MRC1"
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - WHOLESALE" Then
        ThisBranch = "MRC2"
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - GANTE" Then
        ThisBranch = "MRC2 RETAIL"
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - VAN SELLING" Then
        ThisBranch = "VAN1"
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - VAN 2" Then
        ThisBranch = "VAN2"
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - ULAS" Then
        ThisBranch = "ULAS"
    ElseIf SysBranch = "MC PARTS DEPOT TORIL" Then
        ThisBranch = "TORIL"
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - PANABO" Then
        ThisBranch = "PANABO"
    End If
    
    MotoRep.sBar.Panels(6).Text = "Daily Cash Collection"
End Sub

Private Sub DisplayReport()
    Dim crDaily As New rptDailyCash
    If rsRPT.State = adStateOpen Then rsRPT.Close
    rsRPT.Open MQry, LYRcon, adOpenKeyset, adLockOptimistic
    crDaily.DiscardSavedData
    crDaily.ParameterFields(1).ClearCurrentValueAndRange
    crDaily.ParameterFields(1).AddCurrentValue SysBranch
    crDaily.ParameterFields(2).ClearCurrentValueAndRange
    crDaily.ParameterFields(2).AddCurrentValue SysAddress
    crDaily.Database.SetDataSource rsRPT
    vwrDaily.ReportSource = crDaily
    vwrDaily.ViewReport
    vwrDaily.EnableExportButton = True
End Sub

Private Sub PrintPreView(Ppvw As Boolean)
    If Ppvw Then
        Frame1.Visible = False
        flxDC.Visible = False
        vwrDaily.Visible = True
    Else
        vwrDaily.Visible = False
        Frame1.Visible = True
        flxDC.Visible = True
    End If
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsRPT.State = adStateOpen Then rsRPT.Close
End Sub

Private Sub mnuAttach_Click()
On Error GoTo er
    Dim AQry As String
    ' Updated By SMC 2016-08-11
    ' allocation table is added to AQry Query
    If InStr(UCase$(flxDC.TextMatrix(ActRow, 9)), "CHECK") > 0 Then
        AQry = " SELECT PAYMENTS.CUSTOMERNAME, PAYMENTS.SPAREPARTS + PAYMENTS.SERVICE + PAYMENTS.OTHERS AS ORAMOUNT, PAYMENTS.CRNO, PAYMENTS.PAYDATE, ORDETAILS.INVOICENO, ORDETAILS.JOBORDERNO, ORDETAILS.TRANSDATE, ORDETAILS.ACCTAMT, " & _
               " ORDETAILS.AMTPAID, ORDETAILS.TTLPAYMENT, ORDETAILS.Balance, round(ifnull(allocation.TAXWHELDAMOUNT,0),2) as  TAXWHELDAMOUNT  " & _
               " From PAYMENTS INNER JOIN ORDETAILS ON (PAYMENTS.PAYMENTID = ORDETAILS.PAYMENTID) " & _
               " Left Join  allocation  on payments.ALLOCATIONID =  allocation.ALLOCATIONID " & _
               " Where PAYMENTS.PAYMENTID = " & Val(flxDC.TextMatrix(ActRow, 13)) & " Order By ORDETAILS.JOBORDERNO, ORDETAILS.InvoiceNO;"
        PRNmc.MySource = "CRATTACHMENT"
        PRNmc.QryPrn = AQry
        PRNmc.ParamValue3 = flxDC.TextMatrix(ActRow, 9)
        PRNmc.Show
        
    Else
        
        AQry = "SELECT PAYMENTS.CUSTOMERNAME, PAYMENTS.SPAREPARTS + PAYMENTS.SERVICE + PAYMENTS.OTHERS AS ORAMOUNT, PAYMENTS.CRNO, PAYMENTS.PAYDATE, ORDETAILS.INVOICENO, ORDETAILS.JOBORDERNO, ORDETAILS.TRANSDATE, ORDETAILS.ACCTAMT, " & _
               " ORDETAILS.AMTPAID, ORDETAILS.TTLPAYMENT, ORDETAILS.Balance, round(ifnull(allocation.TAXWHELDAMOUNT,0),2) as  TAXWHELDAMOUNT " & _
               " From PAYMENTS INNER JOIN ORDETAILS ON (PAYMENTS.PAYMENTID = ORDETAILS.PAYMENTID)  " & _
               " Left Join allocation  on payments.ALLOCATIONID =  allocation.ALLOCATIONID " & _
               " Where PAYMENTS.PAYMENTID = " & Val(flxDC.TextMatrix(ActRow, 13)) & " Order By ORDETAILS.JOBORDERNO, ORDETAILS.InvoiceNO;"
        'InputBox AQry, AQry, AQry
        PRNmc.MySource = "ORATTACHMENT"
        PRNmc.QryPrn = AQry
        PRNmc.Show
        
    End If
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->mnuAttach_Click()"
End Sub

Private Sub CancelCheckAlloc(ChkID As Long, MyCR As String)
    Dim rsAlloc As New ADODB.Recordset
    Dim Qry As String
    Dim i As Integer
    
    'Get records from check allocation
    Qry = "SELECT checkalloc.ORNO, checkalloc.CHKSPRID, checkalloc.SPSORDERID, checkalloc.BRANCH, checkalloc.REFNO, checkalloc.CUSTOMERID, checkalloc.ALLOCAMT " & _
          "From checkalloc Where checkalloc.CHKSPRID = " & ChkID & " AND checkalloc.ORNO = '" & MyCR & "';"
    With rsAlloc
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            FlxChkAlloc.Rows = 1
            .MoveFirst
            Do While Not .EOF
                FlxChkAlloc.AddItem !ORNO & vbTab & !CHKSPRID & vbTab & !SPSORDERID & vbTab & !Branch & vbTab & !RefNo & vbTab & !CUSTOMERID & vbTab & !AllocAmt
                .MoveNext
            Loop
        Else
            .Close
            Exit Sub
        End If
        .Close
    End With
    'loop thru flxCHKalloc: 0-CR#, 1-ChksPRid, 2-SPSOrderID, 3-Branch, 4-INV#, 5-CustID, 6-Alloc Amt
    With FlxChkAlloc
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 3) = ThisBranch Then
                'update ledger REFTYPE="CANCELLED", CUSTOMERID=0
                LYRcon.Execute "UPDATE Ledger SET REFTYPE = 'CANCELLED', CUSTOMERID = 0 " & _
                "Where ledger.SPINVNO = '" & .TextMatrix(i, 4) & "' AND ledger.REFNO = '" & .TextMatrix(i, 0) & "';"
                'UPDATE ONACCTERMS  SET TTLAMTPAID, BALANCE, STATUS
                LYRcon.Execute "UPDATE onaccterms SET TTLAMTPAID = onaccterms.TTLAMTPAID - " & Val(.TextMatrix(i, 6)) & ", BALANCE = onaccterms.BALANCE + " & Val(.TextMatrix(i, 6)) & ", `STATUS` = 'OPEN' " & _
                "Where onaccterms.SPINVNO = '" & .TextMatrix(i, 4) & "' AND onaccterms.SPSORDERID = " & Val(.TextMatrix(i, 2)) & ";"
                'UPDATE CHECK ALLOCATION SET ORNO, ORDATE, OR AMT TO NULL
                LYRcon.Execute "UPDATE checkalloc SET ORAMT = 0, ORNO = '', ORDATE = Null " & _
                "Where checkalloc.CHKSPRID = " & Val(.TextMatrix(i, 1)) & " AND checkalloc.REFNO = '" & .TextMatrix(i, 4) & "';"
                'UPDATE CHECK BALANCE
                LYRcon.Execute "UPDATE checks SET CHECKBAL = checks.CHECKBAL + " & Val(.TextMatrix(i, 6)) & " " & _
                "Where checks.CHKSPRID = " & Val(.TextMatrix(i, 1)) & ";"
            Else
                If ThisBranch = "MRC2" Then
                    If .TextMatrix(i, 3) = "MRC1" Then
                        If ConnectMRC1 Then
                            'update ledger REFTYPE="CANCELLED", CUSTOMERID=0
                            conMRC1.Execute "UPDATE Ledger SET REFTYPE = 'CANCELLED', CUSTOMERID = 0 " & _
                            "Where ledger.SPINVNO = '" & .TextMatrix(i, 4) & "' AND ledger.REFNO = '" & .TextMatrix(i, 0) & "';"
                            'UPDATE ONACCTERMS  SET TTLAMTPAID, BALANCE, STATUS
                            conMRC1.Execute "UPDATE onaccterms SET TTLAMTPAID = onaccterms.TTLAMTPAID - " & Val(.TextMatrix(i, 6)) & ", BALANCE = onaccterms.BALANCE + " & Val(.TextMatrix(i, 6)) & ", `STATUS` = 'OPEN' " & _
                            "Where onaccterms.SPINVNO = '" & .TextMatrix(i, 4) & "' AND onaccterms.SPSORDERID = " & Val(.TextMatrix(i, 2)) & ";"
                            'UPDATE CHECK ALLOCATION SET ORNO, ORDATE, OR AMT TO NULL
                            conMRC1.Execute "UPDATE checkalloc SET ORAMT = 0, ORNO = '', ORDATE = Null " & _
                            "Where checkalloc.CHKSPRID = " & Val(.TextMatrix(i, 1)) & " AND checkalloc.REFNO = '" & .TextMatrix(i, 4) & "';"
                            'UPDATE CHECK BALANCE
                            conMRC1.Execute "UPDATE checks SET CHECKBAL = checks.CHECKBAL + " & Val(.TextMatrix(i, 6)) & " " & _
                            "Where checks.CHKSPRID = " & Val(.TextMatrix(i, 1)) & ";"
                        Else
                            MsgBox "Error connecting to MRC Sobrecary!" & vbCrLf & "Failure to update Invoice#: " & .TextMatrix(i, 4) & "!" & vbCrLf & _
                                   "Please inform your System Administrator!", vbExclamation + vbOKOnly, "MRC1 Connection Failed"
                        End If
                    ElseIf .TextMatrix(i, 3) = "VAN1" Then
                        If ConnectVAN1 Then
                            'update ledger REFTYPE="CANCELLED", CUSTOMERID=0
                            conVAN1.Execute "UPDATE Ledger SET REFTYPE = 'CANCELLED', CUSTOMERID = 0 " & _
                            "Where ledger.SPINVNO = '" & .TextMatrix(i, 4) & "' AND ledger.REFNO = '" & .TextMatrix(i, 0) & "';"
                            'UPDATE ONACCTERMS  SET TTLAMTPAID, BALANCE, STATUS
                            conVAN1.Execute "UPDATE onaccterms SET TTLAMTPAID = onaccterms.TTLAMTPAID - " & Val(.TextMatrix(i, 6)) & ", BALANCE = onaccterms.BALANCE + " & Val(.TextMatrix(i, 6)) & ", `STATUS` = 'OPEN' " & _
                            "Where onaccterms.SPINVNO = '" & .TextMatrix(i, 4) & "' AND onaccterms.SPSORDERID = " & Val(.TextMatrix(i, 2)) & ";"
                            'UPDATE CHECK ALLOCATION SET ORNO, ORDATE, OR AMT TO NULL
                            conVAN1.Execute "UPDATE checkalloc SET ORAMT = 0, ORNO = '', ORDATE = Null " & _
                            "Where checkalloc.CHKSPRID = " & Val(.TextMatrix(i, 1)) & " AND checkalloc.REFNO = '" & .TextMatrix(i, 4) & "';"
                            'UPDATE CHECK BALANCE
                            conVAN1.Execute "UPDATE checks SET CHECKBAL = checks.CHECKBAL + " & Val(.TextMatrix(i, 6)) & " " & _
                            "Where checks.CHKSPRID = " & Val(.TextMatrix(i, 1)) & ";"
                        Else
                            MsgBox "Error connecting to VAN SELLING 1!" & vbCrLf & "Failure to update Invoice#: " & .TextMatrix(i, 4) & "!" & vbCrLf & _
                                   "Please inform your System Administrator!", vbExclamation + vbOKOnly, "VAN1 Connection Failed"
                        End If
                    ElseIf .TextMatrix(i, 3) = "VAN2" Then
                        If ConnectVAN2 Then
                            'update ledger REFTYPE="CANCELLED", CUSTOMERID=0
                            conVAN2.Execute "UPDATE Ledger SET REFTYPE = 'CANCELLED', CUSTOMERID = 0 " & _
                            "Where ledger.SPINVNO = '" & .TextMatrix(i, 4) & "' AND ledger.REFNO = '" & .TextMatrix(i, 0) & "';"
                            'UPDATE ONACCTERMS  SET TTLAMTPAID, BALANCE, STATUS
                            conVAN2.Execute "UPDATE onaccterms SET TTLAMTPAID = onaccterms.TTLAMTPAID - " & Val(.TextMatrix(i, 6)) & ", BALANCE = onaccterms.BALANCE + " & Val(.TextMatrix(i, 6)) & ", `STATUS` = 'OPEN' " & _
                            "Where onaccterms.SPINVNO = '" & .TextMatrix(i, 4) & "' AND onaccterms.SPSORDERID = " & Val(.TextMatrix(i, 2)) & ";"
                            'UPDATE CHECK ALLOCATION SET ORNO, ORDATE, OR AMT TO NULL
                            conVAN2.Execute "UPDATE checkalloc SET ORAMT = 0, ORNO = '', ORDATE = Null " & _
                            "Where checkalloc.CHKSPRID = " & Val(.TextMatrix(i, 1)) & " AND checkalloc.REFNO = '" & .TextMatrix(i, 4) & "';"
                            'UPDATE CHECK BALANCE
                            conVAN2.Execute "UPDATE checks SET CHECKBAL = checks.CHECKBAL + " & Val(.TextMatrix(i, 6)) & " " & _
                            "Where checks.CHKSPRID = " & Val(.TextMatrix(i, 1)) & ";"
                        Else
                            MsgBox "Error connecting to VAN SELLING 2!" & vbCrLf & "Failure to update Invoice#: " & .TextMatrix(i, 4) & "!" & vbCrLf & _
                                   "Please inform your System Administrator!", vbExclamation + vbOKOnly, "VAN2 Connection Failed"
                        End If
                
                    End If
                    LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.SPAREPARTS = 0, PAYMENTS.SERVICE = 0, PAYMENTS.REMARKS = 'CANCELLED' WHERE PAYMENTS.CRNO='" & MyCR & "';"
                    LYRcon.Execute "UPDATE CHECKALLOC SET CHECKALLOC.ORAMT = 0, CHECKALLOC.ORNO =  Null , CHECKALLOC.ORDATE = Null WHERE CHECKALLOC.ORNO='" & MyCR & "' AND CHECKALLOC.CHKSPRID=" & ChkID & ";"
                    LYRcon.Execute "UPDATE CHECKS SET CHECKS.CHECKBAL = CHECKS.CHECKBAL+" & Val(.TextMatrix(i, 6)) & " WHERE CHECKS.CHKSPRID=" & ChkID & ";"
                
                
                End If  'IF THISBRANCH=MRC2
            End If
        Next i
    End With
    
  '  LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
        "VALUES ('" & Format(Date, "yyyy-mm-dd") & "', 'CANCEL', 'COLLECTION RECEIPT', '" & MyCR & "', '" & Format(flxDC.TextMatrix(ActRow, 11), "yyyy-mm-dd") & "', '" & Left$(MyReason, 200) & "', " & UserID & ");"
    
    
    
    
    flxDC.TextMatrix(ActRow, 5) = "0.00"
    flxDC.TextMatrix(ActRow, 9) = "CANCELLED"
    MsgBox "C.R. Number " & MyCR & " successfully cancelled!", vbInformation + vbOKOnly, "Cancel"
    
    
End Sub

Private Sub mnuCancel_Click()
MsgBox "Not applicable to this Transaction. Please use Cancellation Module", vbInformation + vbOKOnly, "Validation"
'On Error GoTo er
'    Dim rsLedger As New ADODB.Recordset
'    Dim Qry As String, MyOR As String, InvBal As Single, TtlAmtPd As Currency, MyCredit As Currency, MyDebit As Currency
'    Dim ORAmt As Single, CancelType As Integer, CustID As Integer, i As Integer, MyJOid As Long, MyChkID As Long
'    Dim myRem As String, MyAcc As String, MyCurrentCr As Long, MyCR As String
'
'    If ActRow <= 0 Then
'        ActRow = 0
'        MsgBox "Please select a record!", vbCritical, "Error->mnuCancel_Click()"
'        Exit Sub
'
'    ElseIf ActRow > flxDC.Rows Then
'        ActRow = 0
'        MsgBox "Please select a record!", vbCritical, "Error->mnuCancel_Click()"
'        Exit Sub
'    End If
'
'
'    'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=74));"
'    'mySQL
'    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=74;"
'    MyAcc = GetFieldValue(Qry)
'    If MyAcc <> "FULL ACCESS" Then
'        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
'        Exit Sub
'    End If
'
'    If flxDC.TextMatrix(ActRow, 11) <> "" Then
'        MyCR = flxDC.TextMatrix(ActRow, 2)
'        MyOR = ""
'    Else
'        MyOR = flxDC.TextMatrix(ActRow, 2)
'        MyCR = ""
'    End If
'
'    If MyOR <> "" Then
'        Ans = MsgBox("Do you want to Cancel O.R.?" & vbCrLf & vbCrLf & _
'            "O.R. Date : " & flxDC.TextMatrix(ActRow, 11) & vbCrLf & _
'            "O.R. Number : " & MyOR & vbCrLf & _
'            "Customer Name: " & flxDC.TextMatrix(ActRow, 3), vbQuestion + vbYesNo, "Cancel")
'    Else
'        Ans = MsgBox("Do you want to Cancel C.R.?" & vbCrLf & vbCrLf & _
'            "C.R. Date : " & flxDC.TextMatrix(ActRow, 11) & vbCrLf & _
'            "C.R. Number : " & MyCR & vbCrLf & _
'            "Customer Name: " & flxDC.TextMatrix(ActRow, 3), vbQuestion + vbYesNo, "Cancel")
'    End If
'
'    If Ans = vbNo Then Exit Sub
'    MyReason = InputBox("Please indicate reason for Cancellation of " & IIf(MyOR <> "", "O.R.:", "C.R.:"), "Reason for Cancel")
'    MyReason = Replace(MyReason, "'", "`")
'    If MyReason = "" Then
'        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Receipt failed!", vbExclamation + vbOKOnly, "Cancel Failed"
'        Exit Sub
'    End If
'    'IDENTIFY TYPE OF CANCELLATION: 1-DP, 2-SERVICE, 3-OTHERS, 4-SP:ON ACCOUNT, 5-SP:SERVICE
'    CancelType = 0
'    If Val(flxDC.TextMatrix(ActRow, 4)) > 0 Then
'        CancelType = 1
'        ORAmt = Val(Replace(flxDC.TextMatrix(ActRow, 4), ",", ""))
'    ElseIf Val(flxDC.TextMatrix(ActRow, 6)) > 0 Then
'        'Qry = "SELECT LEDGER.* From LEDGER WHERE (((LEDGER.REFNO)='" & MyOR & "'));"
'        'mySQL
'        If MyOR <> "" Then
'            Qry = "SELECT LEDGER.* From LEDGER WHERE LEDGER.REFNO='" & MyOR & "';"
'        Else
'            Qry = "SELECT LEDGER.* From LEDGER WHERE LEDGER.REFNO='" & MyCR & "';"
'        End If
'        If CtrlNumExist(Qry) Then
'            CancelType = 4
'            If Val(flxDC.TextMatrix(ActRow, 5)) > 0 Then
'                ORAmt = Val(Replace(flxDC.TextMatrix(ActRow, 5), ",", "")) + Val(Replace(flxDC.TextMatrix(ActRow, 6), ",", ""))
'            Else
'                ORAmt = Val(Replace(flxDC.TextMatrix(ActRow, 6), ",", ""))
'            End If
'        Else
'            CancelType = 2
'            ORAmt = Val(Replace(flxDC.TextMatrix(ActRow, 6), ",", ""))
'        End If
'    ElseIf Val(flxDC.TextMatrix(ActRow, 7)) > 0 Then
'        CancelType = 3
'        ORAmt = Val(Replace(flxDC.TextMatrix(ActRow, 7), ",", ""))
'    ElseIf Val(Replace(flxDC.TextMatrix(ActRow, 5), ",", "")) > 0 Then
'        ORAmt = Val(Replace(flxDC.TextMatrix(ActRow, 5), ",", ""))
'        'Qry = "SELECT LEDGER.* From LEDGER WHERE (((LEDGER.REFNO)='" & MyOR & "'));"
'        'mySQL
'        If InStr(UCase$(flxDC.TextMatrix(ActRow, 9)), "CHECK") > 0 Then 'Possible use of Check Allocation assigned to other branch
'            Qry = "SELECT payments.CHECKPAYMENT From payments Where payments.CRNO = '" & MyCR & "';"
'            MyChkID = Val(GetFieldValue(Qry))
'            If MyChkID > 1 Then
'                Call CancelCheckAlloc(MyChkID, MyCR)
'                Exit Sub
'            End If
'        End If
'        If MyOR <> "" Then
'            Qry = "SELECT LEDGER.* From LEDGER WHERE LEDGER.REFNO='" & MyOR & "';"
'        Else
'            Qry = "SELECT LEDGER.* From LEDGER WHERE LEDGER.REFNO='" & MyCR & "';"
'        End If
'
'        If CtrlNumExist(Qry) Then
'            CancelType = 4
'            If MyOR <> "" Then
'                Qry = "SELECT payments.CUSTOMERID From payments Where payments.ORNO = '" & MyOR & "';"
'            Else
'                 Qry = "SELECT payments.CUSTOMERID From payments Where payments.CRNO = '" & MyCR & "';"
'            End If
'            CustID = GetFieldValue(Qry)
'        Else
'            'Qry = "SELECT SPINVOICE.* From SPINVOICE WHERE (((SPINVOICE.ORNO)='" & MyOR & "'));"
'            'mySQL
'            If MyOR <> "" Then
'                Qry = "SELECT SPINVOICE.* From SPINVOICE WHERE SPINVOICE.ORNO='" & MyOR & "';"
'            Else
'                Qry = "SELECT SPINVOICE.* From SPINVOICE WHERE SPINVOICE.ORNO='" & MyCR & "';"
'            End If
'            If CtrlNumExist(Qry) Then CancelType = 5
'        End If
'    End If
'    If CancelType = 0 Then
'        MsgBox "Unable to identify Receipt Type..." & vbCrLf & "Please inform your System Administrator!", vbOKOnly + vbExclamation, "Cancel Denied"
'        Exit Sub
'    End If
'
'    If CancelType = 1 Then 'Down Payment
'        If MyOR <> "" Then
'            Qry = "SELECT LEDGER.* From LEDGER WHERE LEDGER.REFNO='" & MyOR & "' AND LEDGER.REFTYPE='Payment';"
'            If CtrlNumExist(Qry) Then
'                MsgBox "Cannot Cancel O.R. because an Invoice has already been issued!", vbExclamation + vbOKOnly, "Cancel"
'                Exit Sub
'            End If
'            'mysql
'            LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.DOWNPAYMENT = 0, PAYMENTS.REMARKS = 'CANCELLED' WHERE PAYMENTS.ORNO='" & MyOR & "';"
'            LYRcon.Execute "UPDATE DPLEDGER SET DPLEDGER.STATUS = 'CANCELLED', DPLEDGER.CUSTOMERID = 0 WHERE DPLEDGER.ORNO='" & MyOR & "';"
'            LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
'                            "VALUES ('" & Format(Date, "yyyy-mm-dd") & "', 'CANCELLED', 'OFFICIAL RECEIPT', '" & MyOR & "', '" & Format(flxDC.TextMatrix(ActRow, 11), "yyyy-mm-dd") & "', '" & Left$(MyReason, 200) & "', " & UserID & ");"
'            flxDC.TextMatrix(ActRow, 4) = "0.00"
'            flxDC.TextMatrix(ActRow, 9) = "CANCELLED"
'            MsgBox "O.R. Number " & MyOR & " successfully cancelled!", vbInformation + vbOKOnly, "Cancel"
'        Else
'            Qry = "SELECT LEDGER.* From LEDGER WHERE LEDGER.REFNO='" & MyCR & "' AND LEDGER.REFTYPE='Payment';"
'            If CtrlNumExist(Qry) Then
'                MsgBox "Cannot Cancel C.R. because an Invoice has already been issued!", vbExclamation + vbOKOnly, "Cancel"
'                Exit Sub
'            End If
'            'mysql
'            LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.DOWNPAYMENT = 0, PAYMENTS.REMARKS = 'CANCELLED' WHERE PAYMENTS.CRNO='" & MyCR & "';"
'            LYRcon.Execute "UPDATE DPLEDGER SET DPLEDGER.STATUS = 'CANCELLED', DPLEDGER.CUSTOMERID = 0 WHERE DPLEDGER.ORNO='" & MyCR & "';"
'            LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
'                            "VALUES ('" & Format(Date, "yyyy-mm-dd") & "', 'CANCELLED', 'COLLECTION RECEIPT', '" & MyCR & "', '" & Format(flxDC.TextMatrix(ActRow, 11), "yyyy-mm-dd") & "', '" & Left$(MyReason, 200) & "', " & UserID & ");"
'            flxDC.TextMatrix(ActRow, 4) = "0.00"
'            flxDC.TextMatrix(ActRow, 9) = "CANCELLED"
'            MsgBox "C.R. Number " & MyCR & " successfully cancelled!", vbInformation + vbOKOnly, "Cancel"
'        End If
'        Exit Sub
'    End If
'    If CancelType = 2 Then  'Service
'        If MyOR <> "" Then
'            LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.SERVICE = 0, PAYMENTS.REMARKS = 'CANCELLED' WHERE PAYMENTS.ORNO='" & MyOR & "';"
'            Qry = "SELECT SRVCJOBORDER.* From SRVCJOBORDER WHERE SRVCJOBORDER.ORNO='" & MyOR & "';"
'            If CtrlNumExist(Qry) Then
'                LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.ORNO = '', SRVCJOBORDER.PAYMENTDONE = 0, SRVCJOBORDER.BALANCE = SRVCJOBORDER.BALANCE+" & ORAmt & _
'                                " WHERE SRVCJOBORDER.ORNO='" & MyOR & "';"
'            Else
'                Call CancelServiceOR
'            End If
'            LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
'                            "VALUES ('" & Format(Date, "yyyy-mm-dd") & "', 'CANCEL', 'OFFICIAL RECEIPT', '" & MyOR & "', '" & Format(flxDC.TextMatrix(ActRow, 11), "yyyy-mm-dd") & "', '" & Left$(MyReason, 200) & "', " & UserID & ");"
'            flxDC.TextMatrix(ActRow, 6) = "0.00"
'            flxDC.TextMatrix(ActRow, 9) = "CANCELLED"
'            MsgBox "O.R. Number " & MyOR & " successfully cancelled!", vbInformation + vbOKOnly, "Cancel"
'        Else
'            LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.SERVICE = 0, PAYMENTS.REMARKS = 'CANCELLED' WHERE PAYMENTS.CRNO='" & MyCR & "';"
'            Qry = "SELECT SRVCJOBORDER.* From SRVCJOBORDER WHERE SRVCJOBORDER.ORNO='" & MyCR & "';"
'            If CtrlNumExist(Qry) Then
'                LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.ORNO = '', SRVCJOBORDER.PAYMENTDONE = 0, SRVCJOBORDER.BALANCE = SRVCJOBORDER.BALANCE+" & ORAmt & _
'                                " WHERE SRVCJOBORDER.ORNO='" & MyCR & "';"
'            Else
'                Call CancelServiceOR
'            End If
'            LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
'                            "VALUES ('" & Format(Date, "yyyy-mm-dd") & "', 'CANCEL', 'COLLECTION RECEIPT', '" & MyCR & "', '" & Format(flxDC.TextMatrix(ActRow, 11), "yyyy-mm-dd") & "', '" & Left$(MyReason, 200) & "', " & UserID & ");"
'            flxDC.TextMatrix(ActRow, 6) = "0.00"
'            flxDC.TextMatrix(ActRow, 9) = "CANCELLED"
'            MsgBox "O.R. Number " & MyOR & " successfully cancelled!", vbInformation + vbOKOnly, "Cancel"
'        End If
'        Exit Sub
'    End If
'    If CancelType = 3 Then  'Others
'        myRem = flxDC.TextMatrix(ActRow, 9) 'Get Remarks
'        If MyOR <> "" Then
'            LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.OTHERS = 0, PAYMENTS.REMARKS = 'CANCELLED' WHERE PAYMENTS.ORNO='" & MyOR & "';"
'            If InStr(myRem, "Other Accts") > 0 Then 'On Account Payment
'                'New Code 08/12/2010
'                 Qry = "SELECT ONACCTPCV.VOUCHERID From OnAcctPCV WHERE ONACCTPCV.ORNO='" & MyOR & "';"
'                 MyChkID = Val(GetFieldValue(Qry))
'
'                 Qry = "SELECT DISBURSEMENT.AMOUNT From Disbursement " & _
'                       "WHERE DISBURSEMENT.VOUCHERID=" & MyChkID & " AND DISBURSEMENT.DISBURSETYPE Not Like 'Cancelled%';"
'                 TtlAmtPd = Val(GetFieldValue(Qry))
'
'                 Qry = "SELECT ONACCTPCV.Balance From OnAcctPCV WHERE ONACCTPCV.ORNO='" & MyOR & "';"
'                 AcctBlnce = Val(GetFieldValue(Qry))
'
'                 MyBalnce = TtlAmtPd - AcctBlnce
'                 LYRcon.Execute "UPDATE OnAcctPCV SET BALANCE = " & MyBalnce & ", ORNO = null Where ONACCTPCV.ORNO = '" & MyOR & "';"
'
'
'            End If
'            LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
'                            "VALUES ('" & Format(Date, "yyyy-mm-dd") & "', 'CANCEL', 'OFFICIAL RECEIPT', '" & MyOR & "', '" & Format(flxDC.TextMatrix(ActRow, 11), "yyyy-mm-dd") & "', '" & Left$(MyReason, 200) & "', " & UserID & ");"
'
'            flxDC.TextMatrix(ActRow, 7) = "0.00"
'            flxDC.TextMatrix(ActRow, 9) = "CANCELLED"
'            MsgBox "O.R. Number " & MyOR & " successfully cancelled!", vbInformation + vbOKOnly, "Cancel"
'        Else
'            LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.OTHERS = 0, PAYMENTS.REMARKS = 'CANCELLED' WHERE PAYMENTS.CRNO='" & MyCR & "';"
'            If InStr(myRem, "Other Accts") > 0 Then 'On Account Payment
'                'New Code 08/12/2010
'                 Qry = "SELECT ONACCTPCV.VOUCHERID From OnAcctPCV WHERE ONACCTPCV.ORNO='" & MyCR & "';"
'                 MyChkID = Val(GetFieldValue(Qry))
'
'                 Qry = "SELECT DISBURSEMENT.AMOUNT From Disbursement " & _
'                       "WHERE DISBURSEMENT.VOUCHERID=" & MyChkID & " AND DISBURSEMENT.DISBURSETYPE Not Like 'Cancelled%';"
'                 TtlAmtPd = Val(GetFieldValue(Qry))
'
'                 Qry = "SELECT ONACCTPCV.Balance From OnAcctPCV WHERE ONACCTPCV.ORNO='" & MyCR & "';"
'                 AcctBlnce = Val(GetFieldValue(Qry))
'
'                 MyBalnce = TtlAmtPd - AcctBlnce
'                 LYRcon.Execute "UPDATE OnAcctPCV SET BALANCE = " & MyBalnce & ", ORNO = null Where ONACCTPCV.ORNO = '" & MyCR & "';"
'            End If
'            LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
'                            "VALUES ('" & Format(Date, "yyyy-mm-dd") & "', 'CANCEL', 'COLLECTION RECEIPT', '" & MyCR & "', '" & Format(flxDC.TextMatrix(ActRow, 11), "yyyy-mm-dd") & "', '" & Left$(MyReason, 200) & "', " & UserID & ");"
'
'            flxDC.TextMatrix(ActRow, 7) = "0.00"
'            flxDC.TextMatrix(ActRow, 9) = "CANCELLED"
'            MsgBox "C.R. Number " & MyCR & " successfully cancelled!", vbInformation + vbOKOnly, "Cancel"
'        End If
'        Exit Sub
'    End If
'
'
'    'Updated By SMC
'    '2016-08-09
'    If Trim(flxDC.TextMatrix(ActRow, 14)) <> "" Then ' Payment Type
'        Dim PaymentID As String
'        Dim tmpRs As New ADODB.Recordset
'
'        PaymentID = flxDC.TextMatrix(ActRow, 13)
'
'        Call modAllocation.sp_payments_cancel(PaymentID, Left$(MyReason, 200), tmpRs)
'        If tmpRs.State = 0 Then
'            MsgBox "PAYMENT CANCELLATION  WAS UNSUCCESSFUL EXPECTED PARAMETER WAS NOT ARRIVED." & vbCrLf & "PLEASE INFORM THE MIS DEPARTMENT.", vbCritical, "Error->mnuCancel_Click()"
'            Exit Sub
'        Else
'            If tmpRs.RecordCount > 0 Then
'                If tmpRs!level = "Error" Then
'                    MsgBox tmpRs!message, vbCritical, "Error->mnuCancel_Click()"
'                Else
'                    MsgBox tmpRs!message, vbInformation + vbOKOnly, "Cancel"
'                End If
'            End If
'        End If
'        If tmpRs.State <> 0 Then tmpRs.Close
'
'        'Refresh the record
'        Call RefreshRecord(prevDCode)
'    Else
'        If CancelType = 4 Then  'SP: On Account
'            If MyOR <> "" Then
'                LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.SPAREPARTS = 0, PAYMENTS.SERVICE = 0, PAYMENTS.REMARKS = 'CANCELLED' WHERE PAYMENTS.ORNO='" & MyOR & "';"
'                Qry = "SELECT * From Ledger Where ledger.REFNO = '" & MyOR & "' AND  Ledger.CUSTOMERID = " & CustID & " AND ledger.REFTYPE = 'Payment';"
'            Else
'                LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.SPAREPARTS = 0, PAYMENTS.SERVICE = 0, PAYMENTS.REMARKS = 'CANCELLED' WHERE PAYMENTS.CRNO='" & MyCR & "';"
'                Qry = "SELECT * From Ledger Where ledger.REFNO = '" & MyCR & "' AND  Ledger.CUSTOMERID = " & CustID & " AND ledger.REFTYPE = 'Payment';"
'            End If
'            With rsLedger   'Update Ledger
'                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
'                If .RecordCount > 0 Then
'                    .MoveFirst
'                    flxInvCancel.Rows = 1
'                    CustID = !CUSTOMERID
'                    Do While Not .EOF
'                        MyJOid = IIf(IsNull(!JOBID), 0, !JOBID)
'                        flxInvCancel.AddItem !SPINVNO & vbTab & IIf(MyJOid > 0, "JOB", "INV")
'                        '!RefType = "CANCELLED"
'                        '!CUSTOMERID = 0
'                        .MoveNext
'                    Loop
'                End If
'                .Close
'            End With
'            LYRcon.Execute "UPDATE Ledger SET Ledger.REFTYPE = 'CANCELLED', Ledger.CUSTOMERID = 0 WHERE Ledger.REFNO='" & IIf(MyOR <> "", MyOR, MyCR) & "';"
'            For i = 1 To flxInvCancel.Rows - 1
'                TtlAmtPd = 0
'                MyCurrentCr = 0
'                If flxInvCancel.TextMatrix(i, 1) = "INV" Then
'                    Qry = "SELECT LEDGER.* FROM LEDGER WHERE LEDGER.SPINVNO='" & flxInvCancel.TextMatrix(i, 0) & "' AND LEDGER.CUSTOMERID=" & CustID & " ORDER BY LEDGER.SEQNO;"
'                    InvBal = GetInvAmt(flxInvCancel.TextMatrix(i, 0))
'                Else
'                    Qry = "SELECT ONACCTERMS.ACCTAMOUNT FROM ONACCTERMS WHERE (ONACCTERMS.SPINVNO) Is Null AND ONACCTERMS.JOBORDERNO='" & flxInvCancel.TextMatrix(i, 0) & "';"
'                    InvBal = GetFieldValue(Qry)
'                End If
'                'Get credit per invoice number  USE REFNO IN CRITERIA
'                Qry = "SELECT sum(LEDGER.Credit) as ACCreditAmt From Ledger WHERE LEDGER.SPINVNO='" & flxInvCancel.TextMatrix(i, 0) & "' AND LEDGER.CUSTOMERID=" & CustID & " AND LEDGER.REFTYPE <> 'CANCELLED' ORDER BY LEDGER.SEQNO;"
'                MyCredit = Val(GetFieldValue(Qry))
'                TtlAmtPd = TtlAmtPd + MyCredit
'                Qry = "SELECT LEDGER.Debit From Ledger WHERE LEDGER.SPINVNO='" & flxInvCancel.TextMatrix(i, 0) & "' AND LEDGER.Refno='" & MyOR & "' AND LEDGER.CUSTOMERID=" & CustID & " ORDER BY LEDGER.SEQNO;"
'                MyDebit = Val(GetFieldValue(Qry))
'
'                If MyCredit > 0 Then
'                    InvBal = InvBal - MyCredit
'                ElseIf MyDebit > 0 Then
'                    InvBal = InvBal + MyDebit
'                End If
'                LYRcon.Execute "UPDATE Ledger SET OSbal = " & InvBal & " Where ledger.SPINVNO = '" & flxInvCancel.TextMatrix(i, 0) & "' AND ledger.REFNO = '" & IIf(MyOR <> "", MyOR, MyCR) & "';"
'
'                'Update Onaccount table
'                If flxInvCancel.TextMatrix(i, 1) = "INV" Then
'                    LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.TTLAMTPAID = " & TtlAmtPd & ", ONACCTERMS.BALANCE = " & InvBal & ", ONACCTERMS.STATUS = 'OPEN', ONACCTERMS.LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
'                                   "WHERE ONACCTERMS.SPINVNO='" & flxInvCancel.TextMatrix(i, 0) & "' AND onaccterms.CUSTOMERID = " & CustID & ";"
'                Else
'                    LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.TTLAMTPAID = " & TtlAmtPd & ", ONACCTERMS.BALANCE = " & InvBal & ", ONACCTERMS.STATUS = 'OPEN', ONACCTERMS.LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
'                                   "WHERE ONACCTERMS.JOBORDERNO='" & flxInvCancel.TextMatrix(i, 0) & "' AND ONACCTERMS.SPINVNO Is Null;"
'                End If
'            Next i
'
'            '***************************************************************************************
'            'Transfer this code Up to Get correct balance of credit amount
'            'Update Ledger's RefType to CANCELLED and CustomerID=0 where RefNo=MyOR
'            'LYRcon.Execute "UPDATE Ledger SET Ledger.REFTYPE = 'CANCELLED', Ledger.CUSTOMERID = 0 WHERE (((Ledger.REFNO)='" & MyOR & "'));"
'            'mySql
'            'LYRcon.Execute "UPDATE Ledger SET Ledger.REFTYPE = 'CANCELLED', Ledger.CUSTOMERID = 0 WHERE Ledger.REFNO='" & MyOR & "';"
'            '****************************************************************************************
'
'            'Update Check Details if Check Payment
'            'Qry = "SELECT PAYMENTS.CHECKPAYMENT From PAYMENTS WHERE (((PAYMENTS.ORNO)= '" & MyOR & "'));"
'            'mySQL
'            If MyOR <> "" Then
'                Qry = "SELECT PAYMENTS.CHECKPAYMENT From PAYMENTS WHERE PAYMENTS.ORNO= '" & MyOR & "';"
'            Else
'                Qry = "SELECT PAYMENTS.CHECKPAYMENT From PAYMENTS WHERE PAYMENTS.CRNO= '" & MyCR & "';"
'            End If
'            If GetFieldValue(Qry) >= 1 Then 'CheckPayment=True
'                Qry = "SELECT CHECKALLOC.CHKSPRID, CHECKALLOC.ORNO From CHECKALLOC WHERE CHECKALLOC.ORNO='" & IIf(MyOR <> "", MyOR, MyCR) & "';"
'                MyChkID = Val(GetFieldValue(Qry))
'                LYRcon.Execute "UPDATE CHECKALLOC SET CHECKALLOC.ORAMT = 0, CHECKALLOC.ORNO =  Null , CHECKALLOC.ORDATE = Null WHERE CHECKALLOC.ORNO='" & IIf(MyOR <> "", MyOR, MyCR) & "' AND CHECKALLOC.CHKSPRID=" & MyChkID & ";"
'                LYRcon.Execute "UPDATE CHECKS SET CHECKS.CHECKBAL = CHECKS.CHECKBAL+" & ORAmt & " WHERE CHECKS.CHKSPRID=" & MyChkID & ";"
'            End If
'            LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = CUSTOMER.OSBAL+" & ORAmt & " WHERE CUSTOMER.CUSTOMERID=" & CustID & ";"
'            If MyOR <> "" Then
'                LYRcon.Execute "INSERT INTO AUDITRANS (TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
'                                "VALUES ('" & Format(Date, "yyyy-mm-dd") & "', 'CANCEL', 'OFFICIAL RECEIPT', '" & MyOR & "', '" & Format(flxDC.TextMatrix(ActRow, 11), "yyyy-mm-dd") & "', '" & Left$(MyReason, 200) & "', " & UserID & ");"
'                flxDC.TextMatrix(ActRow, 5) = "0.00"
'                flxDC.TextMatrix(ActRow, 9) = "CANCELLED"
'                MsgBox "O.R. Number " & MyOR & " successfully cancelled!", vbInformation + vbOKOnly, "Cancel"
'            Else
'                LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
'                                "VALUES ('" & Format(Date, "yyyy-mm-dd") & "', 'CANCEL', 'COLLECTION RECEIPT', '" & MyCR & "', '" & Format(flxDC.TextMatrix(ActRow, 11), "yyyy-mm-dd") & "', '" & Left$(MyReason, 200) & "', " & UserID & ");"
'
'                flxDC.TextMatrix(ActRow, 5) = "0.00"
'                flxDC.TextMatrix(ActRow, 9) = "CANCELLED"
'                MsgBox "C.R. Number " & MyCR & " successfully cancelled!", vbInformation + vbOKOnly, "Cancel"
'            End If
'            Exit Sub
'        End If
'
'        If CancelType = 5 Then  'SP:Service
'            If MyOR <> "" Then
'                LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.SPAREPARTS = 0, PAYMENTS.REMARKS = 'CANCELLED' WHERE PAYMENTS.ORNO='" & MyOR & "';"
'                LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.ORNO = Null WHERE SPINVOICE.ORNO='" & MyOR & "';"
'                LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
'                                "VALUES ('" & Format(Date, "yyyy-mm-dd") & "', 'CANCEL', 'OFFICIAL RECEIPT', '" & MyOR & "', '" & Format(flxDC.TextMatrix(ActRow, 11), "yyyy-mm-dd") & "', '" & Left$(MyReason, 200) & "', " & UserID & ");"
'                flxDC.TextMatrix(ActRow, 5) = "0.00"
'                flxDC.TextMatrix(ActRow, 9) = "CANCELLED"
'                MsgBox "O.R. Number " & MyOR & " successfully cancelled!", vbInformation + vbOKOnly, "Cancel"
'            Else
'                LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.SPAREPARTS = 0, PAYMENTS.REMARKS = 'CANCELLED' WHERE PAYMENTS.CRNO='" & MyCR & "';"
'                LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.ORNO = Null WHERE SPINVOICE.ORNO='" & MyCR & "';"
'                LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
'                                "VALUES ('" & Format(Date, "yyyy-mm-dd") & "', 'CANCEL', 'COLLECTION RECEIPT', '" & MyCR & "', '" & Format(flxDC.TextMatrix(ActRow, 11), "yyyy-mm-dd") & "', '" & Left$(MyReason, 200) & "', " & UserID & ");"
'                flxDC.TextMatrix(ActRow, 5) = "0.00"
'                flxDC.TextMatrix(ActRow, 9) = "CANCELLED"
'                MsgBox "C.R. Number " & MyCR & " successfully cancelled!", vbInformation + vbOKOnly, "Cancel"
'            End If
'            Exit Sub
'        End If
'    End If
'
'Exit Sub
'er:
'    MsgBox Err.Description, vbCritical, "Error->mnuCancel_Click()"
End Sub

Function GetInvAmt(InvNo As String) As Single
    Dim rsInv As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT SPINVOICE.AMOUNT From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & InvNo & "'));"
    'mySQL
    Qry = "SELECT SPINVOICE.AMOUNT From SPINVOICE WHERE SPINVOICE.SPINVNO='" & InvNo & "';"
    With rsInv
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            GetInvAmt = !Amount
        Else
            GetInvAmt = 0
        End If
        .Close
    End With
    
End Function

Private Sub CancelServiceOR()
    Dim rsJO As New ADODB.Recordset
    Dim MyJo As String, Qry As String, MyOR As String
    Dim ORAmt As Single
    
    'MyJo = Mid$(flxDC.TextMatrix(ActRow, 9), 5, 5)
    'ORAmt = PureNum(flxDC.TextMatrix(ActRow, 6))
    'MyOR = flxDC.TextMatrix(ActRow, 2)
    'Qry = "SELECT SRVCJOBORDER.* From SRVCJOBORDER WHERE (((SRVCJOBORDER.JOBORDERNO)='" & MyJo & "'));"
    'mySQL
    'Qry = "SELECT SRVCJOBORDER.* From SRVCJOBORDER WHERE SRVCJOBORDER.JOBORDERNO='" & MyJo & "';"
    'With rsJO
    '    .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    '    If .RecordCount > 0 Then
    '        .MoveFirst
    '        MyOR = Replace(!ORNO, MyOR, "")
    '        !ORNO = MyOR
    '        'Access
    '        '!PaymentDone = False
    '        'mySQL
    '        !PaymentDone = 0
    '        !BALANCE = !BALANCE + ORAmt
    '        .Update
    '    End If
    '    .Close
    'End With
    'End If
    
    MyJo = Mid$(flxDC.TextMatrix(ActRow, 9), 5, 5)
    ORAmt = PureNum(flxDC.TextMatrix(ActRow, 6))
    MyOR = flxDC.TextMatrix(ActRow, 2)
                
    Qry = "SELECT SRVCJOBORDER.ORNO From SRVCJOBORDER WHERE SRVCJOBORDER.JOBORDERNO='" & MyJo & "';"
    MyOR = Replace(CStr(GetFieldValue(Qry)), MyOR, "")
                
    Qry = "SELECT SRVCJOBORDER.BALANCE From SRVCJOBORDER WHERE SRVCJOBORDER.JOBORDERNO='" & MyJo & "';"
    SvcBlnce = Val(GetFieldValue(Qry)) + ORAmt
                
    LYRcon.Execute "UPDATE srvcjoborder SET ORNO = '" & MyOR & "', PAYMENTDONE = 0, BALANCE = " & SvcBlnce & " Where srvcjoborder.JOBORDERNO = '" & MyJo & "';"
    
    
End Sub

Private Sub RXResizer1_ResizeComplete()
On Error GoTo er
    Dim i As Integer
    ' 0     1              2             3                4                5               6          7         8             9           10         11       12     13     14
    '   |^  Invoice No.  |^Receipt No. |<Customer Name  |> Down Payment  |>Spare Parts   |>Service  |>Others  |>TAX w/HELD  |<Remarks   |<Cashier   |<Date  | 12   | 13   | PAYTYPE
    With flxDC
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.06
        .ColWidth(2) = (.Width - 300) * 0.06
        .ColWidth(3) = (.Width - 300) * 0.15
        .ColWidth(4) = (.Width - 300) * 0.07
        .ColWidth(5) = (.Width - 300) * 0.07
        .ColWidth(6) = (.Width - 300) * 0.07
        .ColWidth(7) = (.Width - 300) * 0.07
        .ColWidth(8) = (.Width - 300) * 0.07
        .ColWidth(9) = (.Width - 300) * 0.13
        .ColWidth(10) = (.Width - 300) * 0.12
        .ColWidth(11) = (.Width - 300) * 0.09
        .ColWidth(12) = 0
        .ColWidth(13) = 0
        .ColWidth(14) = 0
    End With
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->RXResizer1_ResizeComplete()"
End Sub

Private Sub tmrDrop_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 4 Then
            tmrDrop.Enabled = False
        End If
    End If
End Sub

Private Sub tmrDropRev_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 2 Then
            tmrDropRev.Enabled = False
            'cmdTrans(0).SetFocus
        End If
    End If
End Sub

Private Sub tmrFloat_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top - 50
    If cmdTrans(BtnIndx).Top <= -375 Then
        cmdTrans(BtnIndx).Visible = False
        BtnIndx = BtnIndx - 1
        If BtnIndx = -1 Then
            tmrFloat.Enabled = False
            BtnIndx = 2 'Index of Button view grid
            cmdTrans(2).Top = -375
            cmdTrans(2).Left = 0
            cmdTrans(2).Visible = True
            cmdTrans(3).Top = -375
            cmdTrans(3).Left = cmdTrans(2).Width
            cmdTrans(3).Visible = True
            tmrDrop.Enabled = True
        End If
    End If
End Sub

Private Sub tmrFloatRev_Timer()
    
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top - 50
    If cmdTrans(BtnIndx).Top <= -375 Then
        cmdTrans(BtnIndx).Visible = False
        BtnIndx = BtnIndx - 1
        If BtnIndx = 1 Then 'Index of Button Exit
            tmrFloatRev.Enabled = False
            BtnIndx = 0
            cmdTrans(0).Visible = True
            cmdTrans(1).Visible = True
            tmrDropRev.Enabled = True
        End If
    End If
End Sub

