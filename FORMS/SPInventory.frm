VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form SPInventory 
   Caption         =   "Spare Parts Monthly Inventory"
   ClientHeight    =   9150
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   14430
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtSQL 
      Height          =   285
      Left            =   8520
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
   End
   Begin MSComctlLib.ProgressBar Pbar 
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Tag             =   "XYF"
      Top             =   9150
      Visible         =   0   'False
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5760
      Top             =   0
   End
   Begin VB.TextBox txtYr 
      Height          =   285
      Left            =   7800
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtMo 
      Height          =   285
      Left            =   6960
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6240
      Top             =   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSP 
      Height          =   7815
      Left            =   120
      TabIndex        =   11
      Tag             =   "XYF"
      Top             =   1320
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   13785
      _Version        =   393216
      RowHeightMin    =   285
      BackColorBkg    =   16777215
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Tag             =   "XY"
      Top             =   720
      Width           =   14175
      Begin VB.ComboBox cboCutOff 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1680
         TabIndex        =   12
         Tag             =   "XYF"
         Top             =   172
         Width           =   3135
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   12120
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   195
         Width           =   1935
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   10800
         TabIndex        =   6
         Tag             =   "XY"
         Top             =   150
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
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
         Image           =   "SPInventory.frx":0000
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdNxtCutOff 
         Height          =   375
         Left            =   4920
         TabIndex        =   18
         Tag             =   "XY"
         Top             =   150
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Next Cut-Off"
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
         Image           =   "SPInventory.frx":059A
         cBack           =   -2147483633
      End
      Begin VB.Label Label1 
         Caption         =   "Period Covered:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   210
         Width           =   1575
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
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
      Image           =   "SPInventory.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   8
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
      Image           =   "SPInventory.frx":10CE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   3600
      TabIndex        =   9
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
      Image           =   "SPInventory.frx":1668
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxOrd 
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   9720
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1931
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      FormatString    =   "Date   |<PO Number    |<Orders"
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13680
      TabIndex        =   17
      Tag             =   "XY"
      Top             =   9150
      Width           =   615
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   360
      Top             =   8640
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14670
      m_MDI_Height    =   9720
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MONTHLY INVENTORY"
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
      Height          =   375
      Index           =   2
      Left            =   9720
      TabIndex        =   2
      Top             =   450
      Width           =   4500
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "MOTO>R REPUBLIC CORP."
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
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   450
      Width           =   2615
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "MOTO>R REPUBLIC CORP."
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
      Index           =   1
      Left            =   150
      TabIndex        =   1
      Top             =   475
      Width           =   2615
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MONTHLY INVENTORY"
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
      Left            =   9735
      TabIndex        =   3
      Top             =   480
      Width           =   4500
   End
   Begin VB.Menu mnuINV 
      Caption         =   "SP Inventory"
      Visible         =   0   'False
      Begin VB.Menu mnuProcess 
         Caption         =   "Process Inventory"
      End
   End
End
Attribute VB_Name = "SPInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSPinv As New ADODB.Recordset
Dim ActRow As Integer, fRow As Integer, fCol As Integer
Private Sub cboCutOff_Click()
    Dim MoYr() As String
    
    If cboCutOff.ListIndex = -1 Or cboCutOff.Text = "" Then Exit Sub
    MoYr = Split(cboCutOff.Text, "-")
    txtYr.Text = Trim(MoYr(1))
    txtMo.Text = Month(DateValue(UCase$(MoYr(0)) & "1, " & txtYr.Text))
    Call DisplayInventory
End Sub

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxSP
        For i = fRow To .Rows - 1
            For j = fCol To 12
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 2: .Row = i
                    .ColSel = 13:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 13, i + 1, i)
                    fCol = IIf(j = 13, 2, j + 1)
                    Exit Sub
                End If
            Next j
            fCol = 2
        Next i
    End With
    MsgBox "No match found!", vbInformation + vbOKOnly, "Find"
    fRow = 1
    fCol = 2
End Sub

Private Sub cmdNxtCutOff_Click()
    Dim PrvMo As Integer, PrvYr As Integer, Ans As Integer
    Dim Qry As String
    Dim CutOffDate As Date
    
    'Validate if Next Period wla pa nag lapse
    Qry = "SELECT SPINVENTORY.YEAR, SPINVENTORY.MONTH From SPInventory " & _
          "GROUP BY SPINVENTORY.YEAR, SPINVENTORY.MONTH ORDER BY SPINVENTORY.YEAR DESC , SPINVENTORY.MONTH DESC;"
    PrvYr = Val(GetFieldValue(Qry))
    Qry = "SELECT SPINVENTORY.MONTH, SPINVENTORY.YEAR From SPInventory " & _
          "GROUP BY SPINVENTORY.MONTH, SPINVENTORY.YEAR ORDER BY SPINVENTORY.YEAR DESC , SPINVENTORY.MONTH DESC;"
    PrvMo = Val(GetFieldValue(Qry))
    If PrvYr = 0 Then
        Timer2.Enabled = True
        Exit Sub
    End If
    CutOffDate = DateAdd("m", 2, DateSerial(PrvYr, PrvMo, 1))
    CutOffDate = DateAdd("d", -1, CutOffDate)
    If Date < CutOffDate Then
        MsgBox "Cut-Off will due on " & CutOffDate & vbCrLf & "Inventory can be process after the cutoff", vbInformation + vbOKOnly, "Inventory"
        Exit Sub
    End If
    txtMo.Text = Month(CutOffDate)
    txtYr.Text = Year(CutOffDate)
    Ans = MsgBox("Do you want to process inventory for the period " & MonthName(txtMo.Text) & " - " & txtYr.Text & "?", vbQuestion + vbYesNo, "Process Inventory")
    If Ans = vbNo Then Exit Sub
    'Display all SP in flxSP
    cboCutOff.AddItem MonthName(txtMo.Text) & " - " & txtYr.Text
    cboCutOff.Text = MonthName(txtMo.Text) & " - " & txtYr.Text
    Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"
    If rsSPinv.State = adStateOpen Then rsSPinv.Close
    rsSPinv.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsSPinv.RecordCount = 0 Then
        flxSP.Clear
        flxSP.Rows = 2
        flxSP.Cols = 12
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxSP.DataSource = rsSPinv
    ActRow = 1
    Call RXResizer1_ResizeComplete
    Call ProcessInventory
    
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim dCov As String, Qry As String
    
    
    Select Case Index
        Case 1  'Filter
            For i = 0 To 2
                 With SPfilter.cboLookIn(i)
                   .AddItem "PART NUMBER"
                   .AddItem "DESCRIPTION"
                   .AddItem "BEG. INV."
                   .AddItem "PURCHASES"
                   .AddItem "PURCHASE RETURN"
                   .AddItem "SALES"
                   .AddItem "SALES RETURN"
                   .AddItem "ADJUSTMENTS"
                   .AddItem "INV. END"
                   .AddItem "BACK ORDER"
                 End With
            Next i
            SPfilter.fldNames = "SPAREPARTS.PARTNUMBER*SPAREPARTS.DESCRIPTION*SPINVENTORY.BEGINV*SPINVENTORY.PURCHASES*SPINVENTORY.PURRET*SPINVENTORY.SALES*SPINVENTORY.SALRET*SPINVENTORY.ADJ*SPINVENTORY.ENDINV*SPINVENTORY.BACKORDER"
            SPfilter.fldNdx = "23456789" 'start w/ 0
            SPfilter.fldDates = ""
            SPfilter.fldSupplier = 10
            SPfilter.MySource = "SPINVENTORY"
            SPfilter.Show vbModal
            
        Case 2  'Print
            PRNmc.MySource = "SPINVENTORY"
            PRNmc.QryPrn = txtSQL.Text
            PRNmc.ParamValue1 = "Period Covered: " & cboCutOff.Text
            '***Access
            'PRNmc.ParamValue2 = GetFieldValue("SELECT SETTINGS.VALUE From SETTINGS WHERE (((SETTINGS.DESCRIPTION)='EMP DEFAULT VALUE') AND ((SETTINGS.REMARKS)='STOCK RECEIVER'));")
            'PRNmc.ParamValue3 = GetFieldValue("SELECT SETTINGS.VALUE From SETTINGS WHERE (((SETTINGS.DESCRIPTION)='EMP DEFAULT VALUE') AND ((SETTINGS.REMARKS)='MANAGER'));")
            
            'mySQL
            PRNmc.ParamValue2 = GetFieldValue("SELECT SETTINGS.VALUE From SETTINGS WHERE SETTINGS.DESCRIPTION='EMP DEFAULT VALUE' AND SETTINGS.REMARKS='STOCK RECEIVER';")
            PRNmc.ParamValue3 = GetFieldValue("SELECT SETTINGS.VALUE From SETTINGS WHERE SETTINGS.DESCRIPTION='EMP DEFAULT VALUE' AND SETTINGS.REMARKS='MANAGER';")
           
            
            PRNmc.ParamValue4 = "STOCK RECEIVER"
            PRNmc.ParamValue5 = "OPERATIONS MANAGER"
            PRNmc.Show
        Case 3  'Exit
            Unload Me
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
    End Select
End Sub


Private Sub flxSP_EnterCell()
    ActRow = flxSP.Row
End Sub


Private Sub flxSP_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        Call PopupMenu(mnuINV)
    End If
End Sub

Private Sub Form_Load()
    Dim rsInv As New ADODB.Recordset
    Dim Qry As String
    
    Qry = "SELECT SPINVENTORY.MONTH, SPINVENTORY.YEAR, Count(SPINVENTORY.SPAREPARTID) AS CountOfSPAREPARTID " & _
          "From SPInventory GROUP BY SPINVENTORY.MONTH, SPINVENTORY.YEAR ORDER BY SPINVENTORY.YEAR DESC , SPINVENTORY.MONTH DESC;"
    With rsInv
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cboCutOff.AddItem MonthName(!Month) & " - " & !Year
                .MoveNext
            Loop
        End If
        .Close
    End With
    If cboCutOff.ListCount > 0 Then
        Timer1.Enabled = True
    Else
        Timer2.Enabled = True
    End If
    MotoRep.sBar.Panels(6).Text = "Spare Parts Monthly Inventory Summary"
    'MotoRep.tbMain.Visible = False
    MotoRep.imgBG.Visible = False
    MotoRep.toolBarButton.Visible = False
    ActRow = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsSPinv.State = adStateOpen Then rsSPinv.Close
End Sub


Private Sub mnuProcess_Click()
    Dim rsCardex As New ADODB.Recordset
    Dim dStart As Date, dEnd As Date
    Dim pPct As Single, PrevMo As Integer, PrevYr As Integer, myIN As Integer, myOUT As Integer, PrvBal As Integer
    Dim Qry As String, Qinv As String
    Dim i As Integer, bI As Integer, S As Integer, sR As Integer, P As Integer, pR As Integer, Adj As Integer, eI As Integer, bO As Integer
    
    dStart = DateSerial(txtYr.Text, Val(txtMo.Text), 1)
    dEnd = DateAdd("d", -1, DateAdd("m", 1, dStart))
    PrevMo = Month(DateAdd("d", -1, dStart))
    PrevYr = Year(DateAdd("d", -1, dStart))
        
    '***Access
    'Qry = "SELECT SPITR.TRANSDATE, SPITR.RECEIVE, SPITR.OUT, SPITR.TRANSTYPE, SPITR.LASTUPDATE From SPITR " & _
          "WHERE (((SPITR.TRANSDATE)>=#" & dStart & "# And (SPITR.TRANSDATE)<=#" & dEnd & "#) AND ((SPITR.TRANSTYPE)='RECEIVE' Or (SPITR.TRANSTYPE)='PURCHASE RETURN' Or (SPITR.TRANSTYPE)='SALES RETURN' Or (SPITR.TRANSTYPE)='INVENTORY ADJ.' Or (SPITR.TRANSTYPE)='OUT') AND ((SPITR.REMARKS) Is Null Or (SPITR.REMARKS)<>'Temp') AND ((SPITR.SPAREPARTID)=" & flxSP.TextMatrix(ActRow, 1) & ")) ORDER BY SPITR.TRANSDATE, SPITR.LASTUPDATE;"
    
    'mySQL
    Qry = "SELECT SPITR.TRANSDATE, SPITR.RECEIVE, SPITR.`OUT`, SPITR.TRANSTYPE, SPITR.LASTUPDATE From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(dStart, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(dEnd, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='RECEIVE' Or SPITR.TRANSTYPE='PURCHASE RETURN' Or SPITR.TRANSTYPE='SALES RETURN' Or SPITR.TRANSTYPE='INVENTORY ADJ.' Or SPITR.TRANSTYPE='OUT' AND (SPITR.REMARKS) Is Null Or SPITR.REMARKS<>'Temp' AND SPITR.SPAREPARTID=" & flxSP.TextMatrix(ActRow, 1) & " ORDER BY SPITR.TRANSDATE, SPITR.LASTUPDATE;"

    
    With rsCardex
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        '***Access
        'Qinv = "SELECT SPINVENTORY.ENDINV From SPInventory WHERE (((SPINVENTORY.SPAREPARTID)=" & flxSP.TextMatrix(ActRow, 1) & ") AND ((SPINVENTORY.MONTH)=" & PrevMo & ") AND ((SPINVENTORY.YEAR)=" & PrevYr & "));"
        
        'mySQL
        Qinv = "SELECT SPINVENTORY.ENDINV From SPInventory WHERE (((SPINVENTORY.SPAREPARTID)=" & flxSP.TextMatrix(ActRow, 1) & ") AND ((SPINVENTORY.MONTH)=" & PrevMo & ") AND ((SPINVENTORY.YEAR)=" & PrevYr & "));"
        
        bI = Val(GetFieldValue(Qinv))
        If .RecordCount > 0 Then
            S = 0: sR = 0: P = 0: pR = 0: Adj = 0: eI = bI: PrvBal = bI
            .MoveFirst
            Do While Not .EOF
                If !TRANSTYPE = "INVENTORY ADJ." Then
                    If !RECEIVE > 0 Then
                        Adj = Adj + !RECEIVE
                        myIN = !RECEIVE
                        myOUT = 0
                    ElseIf !OUT > 0 Then
                        Adj = Adj - !OUT
                        myOUT = !OUT
                        myIN = 0
                    End If
                Else
                    If !TRANSTYPE = "OUT" Then
                        S = S + !OUT
                        myOUT = !OUT
                        myIN = 0
                    ElseIf !TRANSTYPE = "RECEIVE" Then
                        P = P + !RECEIVE
                        myOUT = 0
                        myIN = !RECEIVE
                    ElseIf !TRANSTYPE = "SALES RETURN" Then
                        sR = sR + !RECEIVE
                        myOUT = 0
                        myIN = !RECEIVE
                    ElseIf !TRANSTYPE = "PURCHASE RETURN" Then
                        pR = pR + !OUT
                        myOUT = !OUT
                        myIN = 0
                    End If
                End If
                PrvBal = PrvBal + (myIN - myOUT)
                .MoveNext
            Loop
            eI = PrvBal
        Else
            S = 0: sR = 0: P = 0: pR = 0: Adj = 0: eI = bI
        End If
        .Close
    End With
    bO = GetBackOrder(flxSP.TextMatrix(ActRow, 1))
    'Insert Into table Spare Parts Inventory
    LYRcon.Execute "INSERT INTO SPINVENTORY ( SPAREPARTID, MONTH, YEAR, BEGINV, SALES, SALRET, PURCHASES, PURRET, ADJ, ENDINV, BACKORDER ) " & _
                   "VALUES (" & flxSP.TextMatrix(ActRow, 1) & ", " & txtMo.Text & ", " & txtYr.Text & ", " & bI & ", " & S & ", " & sR & ", " & P & ", " & pR & ", " & Adj & ", " & eI & ", " & bO & ");"
    'display in flx
    With flxSP
        .TextMatrix(ActRow, 4) = bI
        .TextMatrix(ActRow, 5) = P
        .TextMatrix(ActRow, 6) = pR
        .TextMatrix(ActRow, 7) = S
        .TextMatrix(ActRow, 8) = sR
        .TextMatrix(ActRow, 9) = Adj
        .TextMatrix(ActRow, 10) = eI
        .TextMatrix(ActRow, 11) = bO
    End With
    MsgBox "Finish processing inventory for " & MonthName(txtMo.Text) & " " & txtYr.Text & "!", vbInformation + vbOKOnly, flxSP.TextMatrix(ActRow, 2)
    
End Sub

Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer
    
    lblHeading(2).Left = Frame1.Left + (Frame1.Width - lblHeading(2).Width) - 30
    lblHeading(3).Left = Frame1.Left + (Frame1.Width - lblHeading(3).Width)
    
    With flxSP             '   1       2          3               4          5         6           7          8           9        10         11       12       13
        .FormatString = "   |  |< PART NUMBER  |<DESCRIPTION|^BEG. INV.|^PURCHASES|^PURCH. RET. |^SALES|^SALES RET.|^ADJUSTMENT|^INV. END|^BACK ORDER|^MONTH |^YEAR"
        .ColWidth(0) = 300
        .ColWidth(1) = 0
        For i = 2 To 11
            .ColWidth(i) = (.Width - 300) * 0.08
        Next i
        .ColWidth(2) = (.Width - 300) * 0.12
        .ColWidth(3) = (.Width - 300) * 0.18
        .ZOrder
    End With
    
    txtFind.Top = cmdFind.Top
End Sub

Private Sub Timer1_Timer()
    cboCutOff.ListIndex = 0
    'Call cboCutOff_Click
    Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
    Dim Ans As Integer, StartDate As Date
    Dim Qry As String
    
    'Display all SP in flxSP
    Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"
    If rsSPinv.State = adStateOpen Then rsSPinv.Close
    rsSPinv.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsSPinv.RecordCount = 0 Then
        flxSP.Clear
        flxSP.Rows = 2
        flxSP.Cols = 12
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxSP.DataSource = rsSPinv
    ActRow = 1
    flxSP.Visible = True
    Call RXResizer1_ResizeComplete
    
    'Identify 1st month-year cutoff
    '***access
    'Qry = "SELECT SPITR.TRANSDATE From SPITR WHERE (((SPITR.TRANSTYPE)='RECEIVE')) ORDER BY SPITR.TRANSDATE;"
    
    'mySQL
    Qry = "SELECT SPITR.TRANSDATE From SPITR WHERE SPITR.TRANSTYPE='RECEIVE' ORDER BY SPITR.TRANSDATE;"
    
    StartDate = GetFieldValue(Qry)
    txtMo.Text = Month(StartDate)
    txtYr.Text = Year(StartDate)
    
    Ans = MsgBox("Do you want to Process Inventory for the Month of " & MonthName(txtMo.Text) & " - " & txtYr.Text, vbQuestion + vbYesNo, MonthName(txtMo.Text) & " Cuf-off")
    If Ans = vbNo Then
        Timer2.Enabled = False
        Exit Sub
    End If
    cboCutOff.AddItem MonthName(txtMo.Text) & " - " & txtYr.Text
    cboCutOff.Text = MonthName(txtMo.Text) & " - " & txtYr.Text
    Call ProcessInventory
    Timer2.Enabled = False
End Sub

Private Sub txtFind_Change()
    fRow = 1
    fCol = 2
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

Private Sub ProcessInventory()
    Dim rsCardex As New ADODB.Recordset
    Dim dStart As Date, dEnd As Date
    Dim pPct As Single, PrevMo As Long, PrevYr As Long, myIN As Long, myOUT As Long, PrvBal As Long
    Dim Qry As String, Qinv As String
    Dim i As Long, bI As Long, S As Long, sR As Long, P As Long, pR As Long, Adj As Long, eI As Long, bO As Long
    
    'Lock objects
    Frame1.Enabled = False
    cmdTrans(1).Enabled = False
    cmdTrans(2).Enabled = False
    cmdTrans(3).Enabled = False
    flxSP.Enabled = False
    
    dStart = DateSerial(txtYr.Text, Val(txtMo.Text), 1)
    dEnd = DateAdd("d", -1, DateAdd("m", 1, dStart))
    PrevMo = Month(DateAdd("d", -1, dStart))
    PrevYr = Year(DateAdd("d", -1, dStart))
        
    Pbar.Value = 0
    Pbar.Max = flxSP.Rows - 1
    Pbar.Visible = True
    For i = 1 To flxSP.Rows - 1
        'Start Progress bar
        'flxCardex.Rows = 1
        '***Access
        'Qry = "SELECT SPITR.TRANSDATE, SPITR.RECEIVE, SPITR.OUT, SPITR.TRANSTYPE, SPITR.LASTUPDATE From SPITR " & _
              "WHERE (((SPITR.TRANSDATE)>=#" & dStart & "# And (SPITR.TRANSDATE)<=#" & dEnd & "#) AND ((SPITR.TRANSTYPE)='RECEIVE' Or (SPITR.TRANSTYPE)='PURCHASE RETURN' Or (SPITR.TRANSTYPE)='SALES RETURN' Or (SPITR.TRANSTYPE)='INVENTORY ADJ.' Or (SPITR.TRANSTYPE)='OUT') AND ((SPITR.REMARKS) Is Null Or (SPITR.REMARKS)<>'Temp') AND ((SPITR.SPAREPARTID)=" & flxSP.TextMatrix(i, 1) & ")) ORDER BY SPITR.TRANSDATE, SPITR.LASTUPDATE;"
        
        'mySQL
        'Qry = "SELECT SPITR.TRANSDATE, SPITR.RECEIVE, SPITR.`OUT`, SPITR.TRANSTYPE, SPITR.LASTUPDATE From SPITR " & _
              "WHERE SPITR.TRANSDATE>='" & Format(dStart, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(dEnd, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='RECEIVE' Or SPITR.TRANSTYPE='PURCHASE RETURN' Or SPITR.TRANSTYPE='SALES RETURN' Or SPITR.TRANSTYPE='INVENTORY ADJ.' Or SPITR.TRANSTYPE='OUT' AND (SPITR.REMARKS) Is Null Or SPITR.REMARKS<>'Temp' AND SPITR.SPAREPARTID=" & flxSP.TextMatrix(i, 1) & " ORDER BY SPITR.TRANSDATE, SPITR.LASTUPDATE;"
        
        Qry = "SELECT SPITR.TRANSDATE, SPITR.RECEIVE, SPITR.OUT, SPITR.TRANSTYPE, SPITR.LASTUPDATE From SPITR " & _
              "WHERE (((SPITR.TRANSDATE)>='" & Format(dStart, "yyyy-mm-dd") & "' And (SPITR.TRANSDATE)<='" & Format(dEnd, "yyyy-mm-dd") & "') AND ((SPITR.TRANSTYPE)='RECEIVE' Or (SPITR.TRANSTYPE)='PURCHASE RETURN' Or (SPITR.TRANSTYPE)='SALES RETURN' Or (SPITR.TRANSTYPE)='INVENTORY ADJ.' Or (SPITR.TRANSTYPE)='OUT') AND ((SPITR.REMARKS) Is Null Or (SPITR.REMARKS)<>'Temp') AND ((SPITR.SPAREPARTID)=" & flxSP.TextMatrix(i, 1) & ")) ORDER BY SPITR.TRANSDATE, SPITR.LASTUPDATE;"
     
        With rsCardex
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            '***Access
            'Qinv = "SELECT SPINVENTORY.ENDINV From SPInventory WHERE (((SPINVENTORY.SPAREPARTID)=" & flxSP.TextMatrix(i, 1) & ") AND ((SPINVENTORY.MONTH)=" & PrevMo & ") AND ((SPINVENTORY.YEAR)=" & PrevYr & "));"
            
            'mySQL
            Qinv = "SELECT SPINVENTORY.ENDINV From SPInventory WHERE (((SPINVENTORY.SPAREPARTID)=" & flxSP.TextMatrix(i, 1) & ") AND ((SPINVENTORY.MONTH)=" & PrevMo & ") AND ((SPINVENTORY.YEAR)=" & PrevYr & "));"
            
            bI = Val(GetFieldValue(Qinv))
            If .RecordCount > 0 Then
                S = 0: sR = 0: P = 0: pR = 0: Adj = 0: eI = bI: PrvBal = bI
                .MoveFirst
                Do While Not .EOF
                    If !TRANSTYPE = "INVENTORY ADJ." Then
                        If !RECEIVE > 0 Then
                            Adj = Adj + !RECEIVE
                            myIN = !RECEIVE
                            myOUT = 0
                        ElseIf !OUT > 0 Then
                            Adj = Adj - !OUT
                            myOUT = !OUT
                            myIN = 0
                        End If
                    Else
                        If !TRANSTYPE = "OUT" Then
                            S = S + !OUT
                            myOUT = !OUT
                            myIN = 0
                        ElseIf !TRANSTYPE = "RECEIVE" Then
                            P = P + !RECEIVE
                            myOUT = 0
                            myIN = !RECEIVE
                        ElseIf !TRANSTYPE = "SALES RETURN" Then
                            sR = sR + !RECEIVE
                            myOUT = 0
                            myIN = !RECEIVE
                        ElseIf !TRANSTYPE = "PURCHASE RETURN" Then
                            pR = pR + !OUT
                            myOUT = !OUT
                            myIN = 0
                        End If
                    End If
                    PrvBal = PrvBal + (myIN - myOUT)
                    .MoveNext
                Loop
                eI = PrvBal
            Else
                S = 0: sR = 0: P = 0: pR = 0: Adj = 0: eI = bI
            End If
            .Close
        End With
        bO = GetBackOrder(flxSP.TextMatrix(i, 1))
        'Insert Into table Spare Parts Inventory
        LYRcon.Execute "INSERT INTO SPINVENTORY ( SPAREPARTID, MONTH, YEAR, BEGINV, SALES, SALRET, PURCHASES, PURRET, ADJ, ENDINV, BACKORDER ) " & _
                       "VALUES (" & flxSP.TextMatrix(i, 1) & ", " & txtMo.Text & ", " & txtYr.Text & ", " & bI & ", " & S & ", " & sR & ", " & P & ", " & pR & ", " & Adj & ", " & eI & ", " & bO & ");"
        Pbar.Value = Pbar.Value + 1
        DoEvents
        pPct = (Pbar.Value / Pbar.Max)
        lblTotal.Caption = Format$(pPct, "##0.0%")
        DoEvents
    Next i
    DoEvents
    
    MsgBox "Finish processing inventory for " & MonthName(txtMo.Text) & " " & txtYr.Text & "!", vbInformation + vbOKOnly, "Update Complete"
    Pbar.Visible = False
    lblTotal.Caption = ""
    Frame1.Enabled = True
    cmdTrans(1).Enabled = True
    cmdTrans(2).Enabled = True
    cmdTrans(3).Enabled = True
    flxSP.Enabled = True
    Call DisplayInventory
End Sub

Function GetBackOrder(MyID As Integer) As Integer
    Dim rsOrder As New ADODB.Recordset
    Dim rsRec As New ADODB.Recordset
    Dim Qry As String, MyPO As String
    Dim xRow As Integer, xBal As Integer, i As Integer, dEnd As Date
    
    'Clear FlxStock
    dEnd = DateAdd("d", -1, DateAdd("m", 1, DateSerial(txtYr.Text, Val(txtMo.Text), 1)))
    '***Access
    'Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.ORDER From SPITR " & _
          "WHERE (((SPITR.TRANSDATE)<=#" & dEnd & "#) AND ((SPITR.TRANSTYPE)='ORDER') AND ((SPITR.SPAREPARTID)=" & MyId & ")) ORDER BY SPITR.TRANSDATE;"
    
    'mySQL
    Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.`ORDER` From SPITR " & _
          "WHERE SPITR.TRANSDATE<='" & Format(dEnd) & "' AND SPITR.TRANSTYPE='ORDER' AND SPITR.SPAREPARTID=" & MyID & " ORDER BY SPITR.TRANSDATE;"
    
    flxOrd.Rows = 1
    With rsOrder
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                flxOrd.AddItem Format(!TRANSDATE, "mm/dd/yyyy") & vbTab & !REFERENCE & vbTab & !Order
                .MoveNext
            Loop
        End If
        .Close
    End With
    xBal = 0
    With flxOrd
        For i = 1 To .Rows - 1
            'Display PO in Stock Card and obtain Balance
            xBal = xBal + Val(.TextMatrix(i, 2))
            'flxStock.AddItem .TextMatrix(i, 0) & vbTab & .TextMatrix(i, 1) & vbTab & vbTab & .TextMatrix(i, 2) & vbTab & vbTab & xBal   '.TextMatrix(i, 2)
            'Acquire Packing list for P.O.
            MyPO = .TextMatrix(i, 1)
            '***Access
            'Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.RECEIVE, SPITR.PACKINGLIST From SPITR " & _
                  "WHERE (((SPITR.TRANSDATE)<=#" & dEnd & "#) AND ((SPITR.PACKINGLIST)='" & MyPO & "') AND ((SPITR.TRANSTYPE)='RECEIVE') AND ((SPITR.SPAREPARTID)=" & MyId & ")) ORDER BY SPITR.TRANSDATE;"
            
            'mySQL
            Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.RECEIVE, SPITR.PACKINGLIST From SPITR " & _
                  "WHERE SPITR.TRANSDATE<='" & Format(dEnd, "yyyy-mm-dd") & "' AND SPITR.PACKINGLIST='" & MyPO & "' AND SPITR.TRANSTYPE='RECEIVE' AND SPITR.SPAREPARTID=" & MyID & " ORDER BY SPITR.TRANSDATE;"
            
            rsRec.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If rsRec.RecordCount > 0 Then
                rsRec.MoveFirst
                Do While Not rsRec.EOF
                    xBal = xBal - rsRec!RECEIVE
                    rsRec.MoveNext
                Loop
            End If
            rsRec.Close
        Next i
    End With
    GetBackOrder = xBal
End Function

Private Sub DisplayInventory()
    Dim Qry As String
    
    '***Access
    'txtSQL.Text = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPINVENTORY.BEGINV, SPINVENTORY.PURCHASES, SPINVENTORY.PURRET, SPINVENTORY.SALES, SPINVENTORY.SALRET, SPINVENTORY.ADJ, SPINVENTORY.ENDINV, SPINVENTORY.BACKORDER " & _
                  "FROM SPINVENTORY INNER JOIN SPAREPARTS ON SPINVENTORY.SPAREPARTID = SPAREPARTS.SPAREPARTID Where (((SPInventory.Month) = " & txtMo.Text & ") And ((SPInventory.Year) = " & txtYr.Text & ")) ORDER BY SPAREPARTS.PARTNUMBER;"
    
    'mySQL
    txtSQL.Text = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPINVENTORY.BEGINV, SPINVENTORY.PURCHASES, SPINVENTORY.PURRET, SPINVENTORY.SALES, SPINVENTORY.SALRET, SPINVENTORY.ADJ, SPINVENTORY.ENDINV, SPINVENTORY.BACKORDER " & _
                  "FROM SPINVENTORY INNER JOIN SPAREPARTS ON SPINVENTORY.SPAREPARTID = SPAREPARTS.SPAREPARTID Where SPInventory.Month = " & txtMo.Text & " And SPInventory.Year = " & txtYr.Text & " ORDER BY SPAREPARTS.PARTNUMBER;"
    
    '***Access
    'Qry = "SELECT SPINVENTORY.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPINVENTORY.BEGINV, SPINVENTORY.PURCHASES, SPINVENTORY.PURRET, SPINVENTORY.SALES, SPINVENTORY.SALRET, SPINVENTORY.ADJ, SPINVENTORY.ENDINV, SPINVENTORY.BACKORDER, SPINVENTORY.MONTH, SPINVENTORY.YEAR " & _
          "FROM SPINVENTORY INNER JOIN SPAREPARTS ON SPINVENTORY.SPAREPARTID = SPAREPARTS.SPAREPARTID Where (((SPInventory.Month) = " & txtMo.Text & ") And ((SPInventory.Year) = " & txtYr.Text & ")) ORDER BY SPAREPARTS.PARTNUMBER;"
    'mySQL
    Qry = "SELECT SPINVENTORY.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPINVENTORY.BEGINV, SPINVENTORY.PURCHASES, SPINVENTORY.PURRET, SPINVENTORY.SALES, SPINVENTORY.SALRET, SPINVENTORY.ADJ, SPINVENTORY.ENDINV, SPINVENTORY.BACKORDER, SPINVENTORY.MONTH, SPINVENTORY.YEAR " & _
          "FROM SPINVENTORY INNER JOIN SPAREPARTS ON SPINVENTORY.SPAREPARTID = SPAREPARTS.SPAREPARTID Where SPInventory.Month = " & txtMo.Text & " And SPInventory.Year = " & txtYr.Text & " ORDER BY SPAREPARTS.PARTNUMBER;"
    
    
    If rsSPinv.State = adStateOpen Then rsSPinv.Close
    rsSPinv.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsSPinv.RecordCount = 0 Then
        flxSP.Clear
        flxSP.Rows = 2
        flxSP.Cols = 12
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxSP.DataSource = rsSPinv
    ActRow = 1
    Call RXResizer1_ResizeComplete
    
End Sub
Public Sub ShowFilteredRecords()
    Dim Qry1 As String, Qry2 As String
    
    
    'Add Criteria in default Select
    '***Access
    'Qry1 = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPINVENTORY.BEGINV, SPINVENTORY.PURCHASES, SPINVENTORY.PURRET, SPINVENTORY.SALES, SPINVENTORY.SALRET, SPINVENTORY.ADJ, SPINVENTORY.ENDINV, SPINVENTORY.BACKORDER " & _
                  "FROM SPINVENTORY INNER JOIN SPAREPARTS ON SPINVENTORY.SPAREPARTID = SPAREPARTS.SPAREPARTID " & txtSQL.Text & " And ((SPInventory.Month) = " & txtMo.Text & " And ((SPInventory.Year) = " & txtYr.Text & ")) ORDER BY SPAREPARTS.PARTNUMBER;"
    'Qry2 = "SELECT SPINVENTORY.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPINVENTORY.BEGINV, SPINVENTORY.PURCHASES, SPINVENTORY.PURRET, SPINVENTORY.SALES, SPINVENTORY.SALRET, SPINVENTORY.ADJ, SPINVENTORY.ENDINV, SPINVENTORY.BACKORDER, SPINVENTORY.MONTH, SPINVENTORY.YEAR " & _
                  "FROM SPINVENTORY INNER JOIN SPAREPARTS ON SPINVENTORY.SPAREPARTID = SPAREPARTS.SPAREPARTID " & txtSQL.Text & " And ((SPInventory.Month) = " & txtMo.Text & " And ((SPInventory.Year) = " & txtYr.Text & ")) ORDER BY SPAREPARTS.PARTNUMBER;"
    
    'mySQL change here
    
    Qry1 = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPINVENTORY.BEGINV, SPINVENTORY.PURCHASES, SPINVENTORY.PURRET, SPINVENTORY.SALES, SPINVENTORY.SALRET, SPINVENTORY.ADJ, SPINVENTORY.ENDINV, SPINVENTORY.BACKORDER " & _
                  "FROM SPINVENTORY INNER JOIN SPAREPARTS ON SPINVENTORY.SPAREPARTID = SPAREPARTS.SPAREPARTID " & txtSQL.Text & " And SPInventory.Month = " & txtMo.Text & " And SPInventory.Year = " & txtYr.Text & " ORDER BY SPAREPARTS.PARTNUMBER;"
    
    Qry2 = "SELECT SPINVENTORY.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPINVENTORY.BEGINV, SPINVENTORY.PURCHASES, SPINVENTORY.PURRET, SPINVENTORY.SALES, SPINVENTORY.SALRET, SPINVENTORY.ADJ, SPINVENTORY.ENDINV, SPINVENTORY.BACKORDER, SPINVENTORY.MONTH, SPINVENTORY.YEAR " & _
                  "FROM SPINVENTORY INNER JOIN SPAREPARTS ON SPINVENTORY.SPAREPARTID = SPAREPARTS.SPAREPARTID " & txtSQL.Text & " And SPInventory.Month = " & txtMo.Text & " And SPInventory.Year = " & txtYr.Text & " ORDER BY SPAREPARTS.PARTNUMBER;"
    
    'Assign Complete Query back to txtSQL
    txtSQL.Text = Qry1
        
    If rsSPinv.State = adStateOpen Then rsSPinv.Close
    rsSPinv.Open Qry2, LYRcon, adOpenKeyset, adLockOptimistic
    If rsSPinv.RecordCount = 0 Then
        flxSP.Clear
        flxSP.Rows = 2
        flxSP.Cols = 12
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxSP.DataSource = rsSPinv
    ActRow = 1
    Call RXResizer1_ResizeComplete
    
End Sub
