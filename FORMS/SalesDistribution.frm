VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form SalesDistribution 
   Caption         =   "Sales Distribution Report"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   -510
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   120
      TabIndex        =   12
      Tag             =   "xyf"
      Top             =   8760
      Visible         =   0   'False
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin TabDlg.SSTab TabSales 
      Height          =   7815
      Left            =   120
      TabIndex        =   5
      Tag             =   "xyf"
      Top             =   840
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   13785
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "DETAILS"
      TabPicture(0)   =   "SalesDistribution.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "flxSD"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "GRAPH PREVIEW"
      TabPicture(1)   =   "SalesDistribution.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSChart1"
      Tab(1).Control(1)=   "cboClass"
      Tab(1).Control(2)=   "lblBranch"
      Tab(1).Control(3)=   "Label2"
      Tab(1).ControlCount=   4
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   6735
         Left            =   -74760
         OleObjectBlob   =   "SalesDistribution.frx":0038
         TabIndex        =   13
         Tag             =   "xyf"
         Top             =   840
         Width           =   14535
      End
      Begin VB.ComboBox cboClass 
         Height          =   315
         Left            =   -73320
         TabIndex        =   15
         Tag             =   "xyf"
         Top             =   480
         Width           =   2175
      End
      Begin VB.Frame Frame2 
         Height          =   2655
         Left            =   120
         TabIndex        =   11
         Tag             =   "xyf"
         Top             =   5040
         Width           =   14655
         Begin MSFlexGridLib.MSFlexGrid flxSP 
            Height          =   2535
            Left            =   120
            TabIndex        =   19
            Tag             =   "xyf"
            Top             =   120
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   4471
            _Version        =   393216
            Cols            =   9
            FixedCols       =   0
            RowHeightMin    =   300
            BackColorBkg    =   -2147483633
            SelectionMode   =   1
            AllowUserResizing=   1
            FormatString    =   $"SalesDistribution.frx":23DB
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flxSD 
         Height          =   4095
         Left            =   120
         TabIndex        =   10
         Tag             =   "xyf"
         Top             =   960
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   7223
         _Version        =   393216
         Rows            =   13
         Cols            =   11
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         FormatString    =   $"SalesDistribution.frx":2480
      End
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   340
         Width           =   14775
         Begin VB.ComboBox cboMonth 
            Height          =   315
            Left            =   7560
            TabIndex        =   22
            Tag             =   "xyf"
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdClass 
            Caption         =   "Update Classification"
            Height          =   375
            Left            =   11160
            TabIndex        =   20
            Tag             =   "xyf"
            Top             =   150
            Width           =   1695
         End
         Begin VB.ComboBox cboBranch 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Tag             =   "xyf"
            Top             =   230
            Width           =   1695
         End
         Begin VB.ComboBox cboYear 
            Height          =   315
            Left            =   4680
            TabIndex        =   9
            Tag             =   "xyf"
            Top             =   260
            Width           =   1335
         End
         Begin VB.CommandButton cmdReprocess 
            Caption         =   "Reprocess"
            Height          =   375
            Left            =   12975
            TabIndex        =   7
            Tag             =   "xyf"
            Top             =   150
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "SELECT MONTH:"
            Height          =   255
            Index           =   2
            Left            =   6240
            TabIndex        =   21
            Tag             =   "xyf"
            Top             =   290
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "SELECT BRANCH:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   290
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "SELECT YEAR:"
            Height          =   255
            Index           =   0
            Left            =   3480
            TabIndex        =   8
            Tag             =   "xyf"
            Top             =   290
            Width           =   1815
         End
      End
      Begin VB.Label lblBranch 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   -63480
         TabIndex        =   18
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "CLASSIFICATION:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   14
         Tag             =   "xyf"
         Top             =   480
         Width           =   1695
      End
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   0
      TabIndex        =   0
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
      Image           =   "SalesDistribution.frx":2526
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   120
      Top             =   8280
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   15480
      m_MDI_Height    =   9720
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SALES DISTRIBUTION"
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
      Left            =   10560
      TabIndex        =   3
      Tag             =   "XYF"
      Top             =   600
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
      TabIndex        =   1
      Tag             =   "XYF"
      Top             =   600
      Width           =   2610
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
      TabIndex        =   2
      Tag             =   "XYF"
      Top             =   600
      Width           =   2610
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SALES DISTRIBUTION"
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
      Left            =   10605
      TabIndex        =   4
      Tag             =   "XYF"
      Top             =   600
      Width           =   4500
   End
   Begin VB.Menu mnuSD 
      Caption         =   "Sales Distribution"
      Visible         =   0   'False
      Begin VB.Menu mnuVerify 
         Caption         =   "Verify Classification"
      End
      Begin VB.Menu mnuDetails 
         Caption         =   "Show Sales Details"
      End
   End
   Begin VB.Menu mnuClass 
      Caption         =   "Parts Class"
      Visible         =   0   'False
      Begin VB.Menu mnuUpdateClass 
         Caption         =   "Update Classification to"
         Begin VB.Menu mnuClassHonda 
            Caption         =   "HONDA GENUINE"
         End
         Begin VB.Menu mnuClassYamaha 
            Caption         =   "YAMAHA GENUINE"
         End
         Begin VB.Menu mnuClassSuzuki 
            Caption         =   "SUZUKI GENUINE"
         End
         Begin VB.Menu mnuClassReplacement 
            Caption         =   "REPLACEMENT"
         End
         Begin VB.Menu mnuClassAccessories 
            Caption         =   "ACCESSORIES"
         End
         Begin VB.Menu mnuClassPerformance 
            Caption         =   "PERFORMANCE"
         End
         Begin VB.Menu mnuClassApparel 
            Caption         =   "RIDER APPARELS"
         End
         Begin VB.Menu mnuClassTIRES 
            Caption         =   "TIRES"
         End
         Begin VB.Menu mnuClassOIL 
            Caption         =   "OIL"
         End
      End
   End
End
Attribute VB_Name = "SalesDistribution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ActRow As Integer
Public SDCloaded As Boolean

Private Sub cboBranch_Click()
    flxSP.Rows = 1
    flxSP.Rows = 2
End Sub

Private Sub cboClass_Click()
    If cboClass.ListIndex = 0 Then
        Call ShowGraphMainAll
    Else
        Call ShowGraphMainIndiv
    End If
End Sub


Private Sub cboMonth_Click()
    Dim i As Integer
    
    If cboMonth.ListIndex >= 1 Then
        flxSD.Rows = 1
        flxSD.AddItem Left$(UCase$(MonthName(cboMonth.ListIndex)), 3) & ". " & cboYear.Text
    ElseIf cboMonth.ListIndex = 0 Then
        flxSD.Rows = 1
        For i = 1 To 12
            If DateSerial(Val(cboYear.Text), i, 1) > DateSerial(Year(Date), Month(Date), 1) Then Exit For
            flxSD.AddItem Left$(UCase$(MonthName(i)), 3) & ". " & cboYear.Text
        Next i
    End If
    flxSP.Rows = 1
    flxSP.Rows = 2
        
End Sub

Private Sub cboYear_Click()
    Dim i As Integer, j As Integer
    Dim Qry As String
    
    flxSD.Rows = 1
    For i = 1 To 12
        If DateSerial(Val(cboYear.Text), i, 1) > DateSerial(Year(Date), Month(Date), 1) Then Exit For
        flxSD.AddItem Left$(UCase$(MonthName(i)), 3) & ". " & cboYear.Text
    Next i
    flxSP.Rows = 1
    flxSP.Rows = 2
    
    'Display Records if available
'    With flxSD
'        pBar.Value = 0
'        pBar.Max = (.Rows - 1) * (.Cols - 1)
'        pBar.Visible = True
'        For i = 1 To .Cols - 1
'            For j = 1 To .Rows - 1
'                Qry = "SELECT topsalesparts.AMOUNT From topsalesparts Where topsalesparts.`MONTH` = " & j & " AND " & _
'                      "topsalesparts.`YEAR` = " & Val(cboYear.Text) & " AND topsalesparts.PARTSCLASS = '" & .TextMatrix(0, i) & "';"
'                .TextMatrix(j, i) = GetFieldValue(Qry)
'                pBar.Value = pBar.Value + 1
'            Next j
'        Next i
'        pBar.Visible = False
'    End With
    
End Sub


Private Sub cmdClass_Click()
    SDCloaded = True
    SalesDistClass.Show
End Sub

Private Sub cmdExit_Click()
        
    If SDCloaded Then Unload SalesDistClass
    'MotoRep.tbMain.Visible = True
    'MotoRep.imgBG.Visible = True
    MotoRep.toolBarButton.Visible = True
    Unload Me
End Sub

Private Sub cmdReprocess_Click()
    Dim i As Integer, j As Integer, Amount As Currency, SalRet As Currency, Ans As Integer
    Dim DateStart As Date, DateEnd As Date
    Dim Qry As String, QRet As String, myRem As String
    
    Ans = MsgBox("Do you want to Reprocess Sales Distribution?", vbQuestion + vbYesNo, "Reprocess")
    If Ans = vbNo Then Exit Sub
    
    
    flxSP.Rows = 1
    flxSP.Rows = 2
    TabSales.Enabled = False
    cmdExit.Enabled = False
    With flxSD
        If cboMonth.ListIndex >= 1 Then 'process single month only
            Pbar.Value = 0
            Pbar.Max = .Cols - 1
            Pbar.Visible = True
            For j = 1 To 10      'Columns of Sales
                DateStart = DateSerial(Int(cboYear.Text), cboMonth.ListIndex, 1)
                DateEnd = DateAdd("d", -1, DateAdd("m", 1, DateValue(DateStart)))
                If .TextMatrix(0, j) = "OTHERS" Then
                    Qry = "SELECT SUM((spitr.OUT * spitr.Price)) As TtlAmt From spitr INNER JOIN spareparts ON (spitr.SPAREPARTID = spareparts.SPAREPARTID) " & _
                          "Where spitr.TRANSDATE >= '" & Format(DateStart, "yyyy-mm-dd") & "' AND spitr.TRANSDATE <= '" & Format(DateEnd, "yyyy-mm-dd") & "' AND spitr.OUT > 0 AND spitr.REFERENCE LIKE 'SI#%' AND " & _
                          "spareparts.REMARKS <> 'HONDA GENUINE' AND spareparts.REMARKS <> 'YAMAHA GENUINE' AND spareparts.REMARKS <> 'SUZUKI GENUINE' AND spareparts.REMARKS <> 'REPLACEMENT' AND " & _
                          "spareparts.REMARKS <> 'ACCESSORIES' AND spareparts.REMARKS <> 'PERFORMANCE' AND spareparts.REMARKS <> 'RIDER APPARELS' AND spareparts.REMARKS <> 'TIRES' AND spareparts.REMARKS <> 'OIL';"
                    QRet = "SELECT Sum((SPITR.RECEIVE * spitr.PRICE)) AS TtlAmtRet From spitr INNER JOIN spareparts ON (spitr.SPAREPARTID = spareparts.SPAREPARTID) " & _
                          "Where spitr.TRANSDATE >= '" & Format(DateStart, "yyyy-mm-dd") & "' AND spitr.TRANSDATE <= '" & Format(DateEnd, "yyyy-mm-dd") & "' AND spitr.RECEIVE > 0 AND spitr.TRANSTYPE = 'SALES RETURN' AND " & _
                          "spareparts.REMARKS <> 'HONDA GENUINE' AND spareparts.REMARKS <> 'YAMAHA GENUINE' AND spareparts.REMARKS <> 'SUZUKI GENUINE' AND spareparts.REMARKS <> 'REPLACEMENT' AND " & _
                          "spareparts.REMARKS <> 'ACCESSORIES' AND spareparts.REMARKS <> 'PERFORMANCE' AND spareparts.REMARKS <> 'RIDER APPARELS' AND spareparts.REMARKS <> 'TIRES' AND spareparts.REMARKS <> 'OIL';"
                Else
                    If InStr(.TextMatrix(0, j), "GEN.") > 0 Then
                        myRem = Replace$(.TextMatrix(0, j), "GEN.", "GENUINE")
                    ElseIf InStr(.TextMatrix(0, j), "APPARELS") > 0 Then
                        myRem = "RIDER APPARELS"
                    Else
                        myRem = .TextMatrix(0, j)
                    End If
                    Qry = "SELECT SUM((spitr.OUT * spitr.PRICE)) AS TtlAmt FROM spitr INNER JOIN spareparts ON (spitr.SPAREPARTID = spareparts.SPAREPARTID) " & _
                    "WHERE spitr.TRANSDATE >= '" & Format(DateStart, "yyyy-mm-dd") & "' AND spitr.TRANSDATE <= '" & Format(DateEnd, "yyyy-mm-dd") & "' AND  spareparts.REMARKS = '" & myRem & "' AND spitr.OUT > 0 AND spitr.REFERENCE LIKE 'SI#%';"
                    
                    QRet = "SELECT Sum((SPITR.RECEIVE * spitr.PRICE)) AS TtlAmtRet FROM spitr INNER JOIN spareparts ON (spitr.SPAREPARTID = spareparts.SPAREPARTID) " & _
                    "WHERE spitr.TRANSDATE >= '" & Format(DateStart, "yyyy-mm-dd") & "' AND spitr.TRANSDATE <= '" & Format(DateEnd, "yyyy-mm-dd") & "' AND  spareparts.REMARKS = '" & myRem & "' AND spitr.RECEIVE > 0 AND spitr.TRANSTYPE = 'SALES RETURN';"
                    
                End If
                If cboBranch.ListIndex = 0 Then
                    Amount = Val(GetFieldValue(Qry))
                    SalRet = Val(GetFieldValue(QRet))
                ElseIf cboBranch.ListIndex = 1 Then
                    Amount = Val(GetFieldValueBranch(Qry, "ULAS"))
                    SalRet = Val(GetFieldValueBranch(QRet, "ULAS"))
                ElseIf cboBranch.ListIndex = 2 Then
                    Amount = Val(GetFieldValueBranch(Qry, "TORIL"))
                    SalRet = Val(GetFieldValueBranch(QRet, "TORIL"))
                ElseIf cboBranch.ListIndex = 3 Then
                    Amount = Val(GetFieldValueBranch(Qry, "YAMAHA"))
                    SalRet = Val(GetFieldValueBranch(QRet, "YAMAHA"))
                ElseIf cboBranch.ListIndex = 4 Then
                    Amount = Val(GetFieldValueBranch(Qry, "VAN1"))
                    SalRet = Val(GetFieldValueBranch(QRet, "VAN1"))
                ElseIf cboBranch.ListIndex = 5 Then
                    Amount = Val(GetFieldValueBranch(Qry, "VAN2"))
                    SalRet = Val(GetFieldValueBranch(QRet, "VAN2"))
                ElseIf cboBranch.ListIndex = 6 Then
                    Amount = Val(GetFieldValueBranch(Qry, "MRC1"))
                    SalRet = Val(GetFieldValueBranch(QRet, "MRC1"))
                ElseIf cboBranch.ListIndex = 7 Then
                    Amount = Val(GetFieldValueBranch(Qry, "PANABO"))
                    SalRet = Val(GetFieldValueBranch(QRet, "PANABO"))
                End If
                .TextMatrix(1, j) = Format(Amount - SalRet, "#,##0.00")
                Pbar.Value = Pbar.Value + 1
                DoEvents
            Next j
            Pbar.Visible = False
        Else
                
            'older code
            Pbar.Value = 0
            Pbar.Max = (.Rows - 1) * (.Cols - 1)
            Pbar.Visible = True
            
            For j = 1 To 10      'Columns of Sales
                For i = 1 To .Rows - 1
                    DateStart = DateSerial(Int(cboYear.Text), i, 1)
                    DateEnd = DateAdd("d", -1, DateAdd("m", 1, DateValue(DateStart)))
                    If .TextMatrix(0, j) = "OTHERS" Then
                        Qry = "SELECT SUM((spitr.OUT * spitr.Price)) As TtlAmt From spitr INNER JOIN spareparts ON (spitr.SPAREPARTID = spareparts.SPAREPARTID) " & _
                              "Where spitr.TRANSDATE >= '" & Format(DateStart, "yyyy-mm-dd") & "' AND spitr.TRANSDATE <= '" & Format(DateEnd, "yyyy-mm-dd") & "' AND spitr.OUT > 0 AND spitr.REFERENCE LIKE 'SI#%' AND " & _
                              "spareparts.REMARKS <> 'HONDA GENUINE' AND spareparts.REMARKS <> 'YAMAHA GENUINE' AND spareparts.REMARKS <> 'SUZUKI GENUINE' AND spareparts.REMARKS <> 'REPLACEMENT' AND " & _
                              "spareparts.REMARKS <> 'ACCESSORIES' AND spareparts.REMARKS <> 'PERFORMANCE' AND spareparts.REMARKS <> 'RIDER APPARELS' AND spareparts.REMARKS <> 'TIRES' AND spareparts.REMARKS <> 'OIL';"
                        QRet = "SELECT Sum((SPITR.RECEIVE * spitr.PRICE)) AS TtlAmtRet From spitr INNER JOIN spareparts ON (spitr.SPAREPARTID = spareparts.SPAREPARTID) " & _
                              "Where spitr.TRANSDATE >= '" & Format(DateStart, "yyyy-mm-dd") & "' AND spitr.TRANSDATE <= '" & Format(DateEnd, "yyyy-mm-dd") & "' AND spitr.RECEIVE > 0 AND spitr.TRANSTYPE = 'SALES RETURN' AND " & _
                              "spareparts.REMARKS <> 'HONDA GENUINE' AND spareparts.REMARKS <> 'YAMAHA GENUINE' AND spareparts.REMARKS <> 'SUZUKI GENUINE' AND spareparts.REMARKS <> 'REPLACEMENT' AND " & _
                              "spareparts.REMARKS <> 'ACCESSORIES' AND spareparts.REMARKS <> 'PERFORMANCE' AND spareparts.REMARKS <> 'RIDER APPARELS' AND spareparts.REMARKS <> 'TIRES' AND spareparts.REMARKS <> 'OIL';"
                    Else
                        If InStr(.TextMatrix(0, j), "GEN.") > 0 Then
                            myRem = Replace$(.TextMatrix(0, j), "GEN.", "GENUINE")
                        ElseIf InStr(.TextMatrix(0, j), "APPARELS") > 0 Then
                            myRem = "RIDER APPARELS"
                        Else
                            myRem = .TextMatrix(0, j)
                        End If
                        Qry = "SELECT SUM((spitr.OUT * spitr.PRICE)) AS TtlAmt FROM spitr INNER JOIN spareparts ON (spitr.SPAREPARTID = spareparts.SPAREPARTID) " & _
                        "WHERE spitr.TRANSDATE >= '" & Format(DateStart, "yyyy-mm-dd") & "' AND spitr.TRANSDATE <= '" & Format(DateEnd, "yyyy-mm-dd") & "' AND  spareparts.REMARKS = '" & myRem & "' AND spitr.OUT > 0 AND spitr.REFERENCE LIKE 'SI#%';"
                        
                        QRet = "SELECT Sum((SPITR.RECEIVE * spitr.PRICE)) AS TtlAmtRet FROM spitr INNER JOIN spareparts ON (spitr.SPAREPARTID = spareparts.SPAREPARTID) " & _
                        "WHERE spitr.TRANSDATE >= '" & Format(DateStart, "yyyy-mm-dd") & "' AND spitr.TRANSDATE <= '" & Format(DateEnd, "yyyy-mm-dd") & "' AND  spareparts.REMARKS = '" & myRem & "' AND spitr.RECEIVE > 0 AND spitr.TRANSTYPE = 'SALES RETURN';"
                        
                    End If
                    If cboBranch.ListIndex = 0 Then
                        Amount = Val(GetFieldValue(Qry))
                        SalRet = Val(GetFieldValue(QRet))
                    ElseIf cboBranch.ListIndex = 1 Then
                        Amount = Val(GetFieldValueBranch(Qry, "ULAS"))
                        SalRet = Val(GetFieldValueBranch(QRet, "ULAS"))
                    ElseIf cboBranch.ListIndex = 2 Then
                        Amount = Val(GetFieldValueBranch(Qry, "TORIL"))
                        SalRet = Val(GetFieldValueBranch(QRet, "TORIL"))
                    ElseIf cboBranch.ListIndex = 3 Then
                        Amount = Val(GetFieldValueBranch(Qry, "YAMAHA"))
                        SalRet = Val(GetFieldValueBranch(QRet, "YAMAHA"))
                    ElseIf cboBranch.ListIndex = 4 Then
                        Amount = Val(GetFieldValueBranch(Qry, "VAN1"))
                        SalRet = Val(GetFieldValueBranch(QRet, "VAN1"))
                    ElseIf cboBranch.ListIndex = 5 Then
                        Amount = Val(GetFieldValueBranch(Qry, "VAN2"))
                        SalRet = Val(GetFieldValueBranch(QRet, "VAN2"))
                    ElseIf cboBranch.ListIndex = 6 Then
                        Amount = Val(GetFieldValueBranch(Qry, "MRC1"))
                        SalRet = Val(GetFieldValueBranch(QRet, "MRC1"))
                    ElseIf cboBranch.ListIndex = 7 Then
                        Amount = Val(GetFieldValueBranch(Qry, "PANABO"))
                        SalRet = Val(GetFieldValueBranch(QRet, "PANABO"))
                    End If
                    .TextMatrix(i, j) = Format(Amount - SalRet, "#,##0.00")
                    Pbar.Value = Pbar.Value + 1
                    DoEvents
                Next i
            Next j
            Pbar.Visible = False
        End If
    End With
    
    MsgBox "Sales Distribution successfully processed!", vbInformation + vbOKOnly, "Done"
    TabSales.Enabled = True
    cmdExit.Enabled = True
    
End Sub


Private Sub flxSD_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        If flxSD.Col > 0 And flxSD.Row > 0 And Val(Replace(flxSD.TextMatrix(flxSD.Row, flxSD.Col), ",", "")) > 0 Then
            Call PopupMenu(mnuSD)
        End If
    End If
End Sub

Private Sub flxSP_EnterCell()
    ActRow = flxSP.Row
End Sub

Private Sub flxSP_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        If Val(flxSP.TextMatrix(ActRow, 8)) > 0 Then
            If flxSP.TextMatrix(ActRow, 4) = "HONDA GENUINE" Then mnuClassHonda.Enabled = False Else mnuClassHonda.Enabled = True
            If flxSP.TextMatrix(ActRow, 4) = "YAMAHA GENUINE" Then mnuClassYamaha.Enabled = False Else mnuClassYamaha.Enabled = True
            If flxSP.TextMatrix(ActRow, 4) = "SUZUKI GENUINE" Then mnuClassSuzuki.Enabled = False Else mnuClassSuzuki.Enabled = True
            If flxSP.TextMatrix(ActRow, 4) = "REPLACEMENT" Then mnuClassReplacement.Enabled = False Else mnuClassReplacement.Enabled = True
            If flxSP.TextMatrix(ActRow, 4) = "ACCESSORIES" Then mnuClassAccessories.Enabled = False Else mnuClassAccessories.Enabled = True
            If flxSP.TextMatrix(ActRow, 4) = "PERFORMANCE" Then mnuClassPerformance.Enabled = False Else mnuClassPerformance.Enabled = True
            If flxSP.TextMatrix(ActRow, 4) = "RIDER APPARELS" Then mnuClassApparel.Enabled = False Else mnuClassApparel.Enabled = True
            Call PopupMenu(mnuClass)
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    For i = Year(Date) To 2009 Step -1
        cboYear.AddItem i
    Next i
    cboMonth.AddItem "All"
    For i = 1 To 12
        cboMonth.AddItem MonthName(i)
    Next i
    With cboClass
        .AddItem "ALL PRODUCTS"
        .AddItem "HONDA GENUINE"
        .AddItem "YAMAHA GENUINE"
        .AddItem "SUZUKI GENUINE"
        .AddItem "REPLACEMENT"
        .AddItem "ACCESSORIES"
        .AddItem "PERFORMANCE"
        .AddItem "RIDER APPARELS"
        .AddItem "TIRES"
        .AddItem "OIL"
        .AddItem "OTHERS"
    End With
    With cboBranch
'        .AddItem "MRC2 GANTE"
        .AddItem "LYR MOTORPARTS CORP. - ULAS"
'        .AddItem "MC PARTS DEPOT"
'        .AddItem "YAMAHA MOTORHUB"
'        .AddItem "VAN SELLING I"
'        .AddItem "VAN SELLING II"
'        .AddItem "MRC1 SOBRECARY"
'        .AddItem "MRC PANABO"
'        .AddItem "MRC GENSAN"
'        .AddItem "LYR SHOP DAVAO"
'       .AddItem "MRC STO. TOMAS"
        .ListIndex = 0
    End With
    TabSales.Tab = 0 '

'    Call ConnectUlas
'    Call ConnectYamaha
'    Call ConnectToril
'    Call ConnectVAN1
'    Call ConnectMRC1
'    Call ConnectPanabo
'    Call ConnectVAN2
'    Call ConnectMRCGensan
'    Call ConnectStoTomas
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If conUlas.State = adStateOpen Then conUlas.Close
   If conYM.State = adStateOpen Then conYM.Close
   If conToril.State = adStateOpen Then conToril.Close
   If conVAN1.State = adStateOpen Then conVAN1.Close
   If conVAN2.State = adStateOpen Then conVAN2.Close
   If conMRC1.State = adStateOpen Then conMRC1.Close
   If conPanabo.State = adStateOpen Then conPanabo.Close
   If conGenSan.State = adStateOpen Then conGenSan.Close
   If conStoTomas.State = adStateOpen Then conStoTomas.Close
End Sub

Private Sub mnuClassAccessories_Click()
    Dim Ans As Integer, i As Integer, MyID As Long
    
    Ans = MsgBox("Do you want to update " & flxSP.TextMatrix(ActRow, 2) & " classification to ACCESSORIES?", vbQuestion + vbYesNo, "Update")
    If Ans = vbNo Then Exit Sub
    
    MyID = Val(flxSP.TextMatrix(ActRow, 8))
    For i = 1 To flxSP.Rows - 1
        If Val(flxSP.TextMatrix(i, 8)) = MyID Then
            flxSP.TextMatrix(i, 4) = "ACCESSORIES"
        End If
    Next i
    'update DB
    If cboBranch.ListIndex = 0 Then     'MOTO-R MAIN"
        LYRcon.Execute "UPDATE spareparts SET REMARKS = 'ACCESSORIES' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 1 Then 'MOTO-R ULAS"
        conUlas.Execute "UPDATE spareparts SET REMARKS = 'ACCESSORIES' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 2 Then 'MC PARTS DEPOT"
        conToril.Execute "UPDATE spareparts SET REMARKS = 'ACCESSORIES' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 3 Then 'YAMAHA MOTORHUB"
        conYM.Execute "UPDATE spareparts SET REMARKS = 'ACCESSORIES' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 4 Then  'VAN SELLING"
        conVAN1.Execute "UPDATE spareparts SET REMARKS = 'ACCESSORIES' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 5 Then  'VAN SELLING2"
        conVAN1.Execute "UPDATE spareparts SET REMARKS = 'ACCESSORIES' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 6 Then 'MRC2 GANTE"
        conMRC1.Execute "UPDATE spareparts SET REMARKS = 'ACCESSORIES' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 7 Then 'panabo"
        conPanabo.Execute "UPDATE spareparts SET REMARKS = 'ACCESSORIES' Where spareparts.SPAREPARTID = " & MyID & ";"
    End If

    MsgBox flxSP.TextMatrix(ActRow, 2) & " successfully updated to ACCESSORIES"
End Sub

Private Sub mnuClassApparel_Click()
    Dim Ans As Integer, i As Integer, MyID As Long
    
    Ans = MsgBox("Do you want to update " & flxSP.TextMatrix(ActRow, 2) & " classification to RIDER APPARELS?", vbQuestion + vbYesNo, "Update")
    If Ans = vbNo Then Exit Sub
    
    MyID = Val(flxSP.TextMatrix(ActRow, 8))
    For i = 1 To flxSP.Rows - 1
        If Val(flxSP.TextMatrix(i, 8)) = MyID Then
            flxSP.TextMatrix(i, 4) = "RIDER APPARELS"
        End If
    Next i                  'RIDER APPARELS'
    'update DB
    If cboBranch.ListIndex = 0 Then     'MOTO-R MAIN"
        LYRcon.Execute "UPDATE spareparts SET REMARKS = 'RIDER APPARELS' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 1 Then 'MOTO-R ULAS"
        conUlas.Execute "UPDATE spareparts SET REMARKS = 'RIDER APPARELS' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 2 Then 'MC PARTS DEPOT"
        conToril.Execute "UPDATE spareparts SET REMARKS = 'RIDER APPARELS' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 3 Then 'YAMAHA MOTORHUB"
        conYM.Execute "UPDATE spareparts SET REMARKS = 'RIDER APPARELS' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 4 Then  'VAN SELLING"
        conVAN1.Execute "UPDATE spareparts SET REMARKS = 'RIDER APPARELS' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 5 Then  'VAN SELLING2"
        conVAN1.Execute "UPDATE spareparts SET REMARKS = 'RIDER APPARELS' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 6 Then 'MRC2 GANTE"
        conMRC1.Execute "UPDATE spareparts SET REMARKS = 'RIDER APPARELS' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 7 Then 'panabo"
        conPanabo.Execute "UPDATE spareparts SET REMARKS = 'RIDER APPARELS' Where spareparts.SPAREPARTID = " & MyID & ";"
    End If

    MsgBox flxSP.TextMatrix(ActRow, 2) & " successfully updated to RIDER APPARELS"

End Sub

Private Sub mnuClassHonda_Click()
    Dim Ans As Integer, i As Integer, MyID As Long
    
    Ans = MsgBox("Do you want to update " & flxSP.TextMatrix(ActRow, 2) & " classification to HONDA GENUINE?", vbQuestion + vbYesNo, "Update")
    If Ans = vbNo Then Exit Sub
    
    MyID = Val(flxSP.TextMatrix(ActRow, 8))
    For i = 1 To flxSP.Rows - 1
        If Val(flxSP.TextMatrix(i, 8)) = MyID Then
            flxSP.TextMatrix(i, 4) = "HONDA GENUINE"
        End If
    Next i
    'update DB
    If cboBranch.ListIndex = 0 Then     'MOTO-R MAIN"
        LYRcon.Execute "UPDATE spareparts SET REMARKS = 'HONDA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 1 Then 'MOTO-R ULAS"
        conUlas.Execute "UPDATE spareparts SET REMARKS = 'HONDA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 2 Then 'MC PARTS DEPOT"
        conToril.Execute "UPDATE spareparts SET REMARKS = 'HONDA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 3 Then 'YAMAHA MOTORHUB"
        conYM.Execute "UPDATE spareparts SET REMARKS = 'HONDA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 4 Then  'VAN SELLING"
        conVAN1.Execute "UPDATE spareparts SET REMARKS = 'HONDA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 5 Then  'VAN SELLING2"
        conVAN1.Execute "UPDATE spareparts SET REMARKS = 'HONDA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 6 Then 'MRC2 GANTE"
        conMRC1.Execute "UPDATE spareparts SET REMARKS = 'HONDA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 7 Then 'panabo"
        conPanabo.Execute "UPDATE spareparts SET REMARKS = 'HONDA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    End If

    MsgBox flxSP.TextMatrix(ActRow, 2) & " successfully updated to HONDA GENUINE"

End Sub

Private Sub mnuClassOIL_Click()
    Dim Ans As Integer, i As Integer, MyID As Long
    
    Ans = MsgBox("Do you want to update " & flxSP.TextMatrix(ActRow, 2) & " classification to OIL?", vbQuestion + vbYesNo, "Update")
    If Ans = vbNo Then Exit Sub
    
    MyID = Val(flxSP.TextMatrix(ActRow, 8))
    For i = 1 To flxSP.Rows - 1
        If Val(flxSP.TextMatrix(i, 8)) = MyID Then
            flxSP.TextMatrix(i, 4) = "OIL"
        End If
    Next i
    'update DB
    If cboBranch.ListIndex = 0 Then     'MOTO-R MAIN"
        LYRcon.Execute "UPDATE spareparts SET REMARKS = 'OIL' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 1 Then 'MOTO-R ULAS"
        conUlas.Execute "UPDATE spareparts SET REMARKS = 'OIL' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 2 Then 'MC PARTS DEPOT"
        conToril.Execute "UPDATE spareparts SET REMARKS = 'OIL' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 3 Then 'YAMAHA MOTORHUB"
        conYM.Execute "UPDATE spareparts SET REMARKS = 'OIL' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 4 Then  'VAN SELLING"
        conVAN1.Execute "UPDATE spareparts SET REMARKS = 'OIL' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 5 Then  'VAN SELLING2"
        conVAN1.Execute "UPDATE spareparts SET REMARKS = 'OIL' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 6 Then 'MRC2 GANTE"
        conMRC1.Execute "UPDATE spareparts SET REMARKS = 'OIL' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 7 Then 'panabo"
        conPanabo.Execute "UPDATE spareparts SET REMARKS = 'OIL' Where spareparts.SPAREPARTID = " & MyID & ";"
    End If

    MsgBox flxSP.TextMatrix(ActRow, 2) & " successfully updated to OIL"

End Sub

Private Sub mnuClassPerformance_Click()
    Dim Ans As Integer, i As Integer, MyID As Long
    
    Ans = MsgBox("Do you want to update " & flxSP.TextMatrix(ActRow, 2) & " classification to PERFORMANCE?", vbQuestion + vbYesNo, "Update")
    If Ans = vbNo Then Exit Sub
    
    MyID = Val(flxSP.TextMatrix(ActRow, 8))
    For i = 1 To flxSP.Rows - 1
        If Val(flxSP.TextMatrix(i, 8)) = MyID Then
            flxSP.TextMatrix(i, 4) = "PERFORMANCE"
        End If
    Next i
    'update DB
    If cboBranch.ListIndex = 0 Then     'MOTO-R MAIN"
        LYRcon.Execute "UPDATE spareparts SET REMARKS = 'PERFORMANCE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 1 Then 'MOTO-R ULAS"
        conUlas.Execute "UPDATE spareparts SET REMARKS = 'PERFORMANCE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 2 Then 'MC PARTS DEPOT"
        conToril.Execute "UPDATE spareparts SET REMARKS = 'PERFORMANCE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 3 Then 'YAMAHA MOTORHUB"
        conYM.Execute "UPDATE spareparts SET REMARKS = 'PERFORMANCE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 4 Then  'VAN SELLING"
        conVAN1.Execute "UPDATE spareparts SET REMARKS = 'PERFORMANCE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 5 Then  'VAN SELLING2"
        conVAN1.Execute "UPDATE spareparts SET REMARKS = 'PERFORMANCE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 6 Then 'MRC2 GANTE"
        conMRC1.Execute "UPDATE spareparts SET REMARKS = 'PERFORMANCE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 7 Then 'panabo"
        conPanabo.Execute "UPDATE spareparts SET REMARKS = 'PERFORMANCE' Where spareparts.SPAREPARTID = " & MyID & ";"
    End If

    MsgBox flxSP.TextMatrix(ActRow, 2) & " successfully updated to PERFORMANCE"

End Sub

Private Sub mnuClassReplacement_Click()
    Dim Ans As Integer, i As Integer, MyID As Long
    
    Ans = MsgBox("Do you want to update " & flxSP.TextMatrix(ActRow, 2) & " classification to REPLACEMENT?", vbQuestion + vbYesNo, "Update")
    If Ans = vbNo Then Exit Sub
    
    MyID = Val(flxSP.TextMatrix(ActRow, 8))
    For i = 1 To flxSP.Rows - 1
        If Val(flxSP.TextMatrix(i, 8)) = MyID Then
            flxSP.TextMatrix(i, 4) = "REPLACEMENT"
        End If
    Next i
    'update DB
    If cboBranch.ListIndex = 0 Then     'MOTO-R MAIN"
        LYRcon.Execute "UPDATE spareparts SET REMARKS = 'REPLACEMENT' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 1 Then 'MOTO-R ULAS"
        conUlas.Execute "UPDATE spareparts SET REMARKS = 'REPLACEMENT' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 2 Then 'MC PARTS DEPOT"
        conToril.Execute "UPDATE spareparts SET REMARKS = 'REPLACEMENT' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 3 Then 'YAMAHA MOTORHUB"
        conYM.Execute "UPDATE spareparts SET REMARKS = 'REPLACEMENT' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 4 Then  'VAN SELLING"
        conVAN1.Execute "UPDATE spareparts SET REMARKS = 'REPLACEMENT' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 5 Then  'VAN SELLING2"
        conVAN1.Execute "UPDATE spareparts SET REMARKS = 'REPLACEMENT' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 6 Then 'MRC2 GANTE"
        conMRC1.Execute "UPDATE spareparts SET REMARKS = 'REPLACEMENT' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 7 Then 'panabo"
        conPanabo.Execute "UPDATE spareparts SET REMARKS = 'REPLACEMENT' Where spareparts.SPAREPARTID = " & MyID & ";"
    End If

    MsgBox flxSP.TextMatrix(ActRow, 2) & " successfully updated to REPLACEMENT"

End Sub

Private Sub mnuClassSuzuki_Click()
    Dim Ans As Integer, i As Integer, MyID As Long
    
    Ans = MsgBox("Do you want to update " & flxSP.TextMatrix(ActRow, 2) & " classification to SUZUKI GENUINE?", vbQuestion + vbYesNo, "Update")
    If Ans = vbNo Then Exit Sub
    
    MyID = Val(flxSP.TextMatrix(ActRow, 8))
    For i = 1 To flxSP.Rows - 1
        If Val(flxSP.TextMatrix(i, 8)) = MyID Then
            flxSP.TextMatrix(i, 4) = "SUZUKI GENUINE"
        End If
    Next i
    'update DB
    'update DB
    If cboBranch.ListIndex = 0 Then     'MOTO-R MAIN"
        LYRcon.Execute "UPDATE spareparts SET REMARKS = 'SUZUKI GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 1 Then 'MOTO-R ULAS"
        conUlas.Execute "UPDATE spareparts SET REMARKS = 'SUZUKI GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 2 Then 'MC PARTS DEPOT"
        conToril.Execute "UPDATE spareparts SET REMARKS = 'SUZUKI GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 3 Then 'YAMAHA MOTORHUB"
        conYM.Execute "UPDATE spareparts SET REMARKS = 'SUZUKI GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 4 Then  'VAN SELLING"
        conVAN1.Execute "UPDATE spareparts SET REMARKS = 'SUZUKI GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 5 Then  'VAN SELLING2"
        conVAN1.Execute "UPDATE spareparts SET REMARKS = 'SUZUKI GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 6 Then 'MRC2 GANTE"
        conMRC1.Execute "UPDATE spareparts SET REMARKS = 'SUZUKI GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 7 Then 'panabo"
        conPanabo.Execute "UPDATE spareparts SET REMARKS = 'SUZUKI GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    End If

    MsgBox flxSP.TextMatrix(ActRow, 2) & " successfully updated to SUZUKI GENUINE"

End Sub

Private Sub mnuClassTIRES_Click()
    Dim Ans As Integer, i As Integer, MyID As Long
    
    Ans = MsgBox("Do you want to update " & flxSP.TextMatrix(ActRow, 2) & " classification to TIRES?", vbQuestion + vbYesNo, "Update")
    If Ans = vbNo Then Exit Sub
    
    MyID = Val(flxSP.TextMatrix(ActRow, 8))
    For i = 1 To flxSP.Rows - 1
        If Val(flxSP.TextMatrix(i, 8)) = MyID Then
            flxSP.TextMatrix(i, 4) = "TIRES"
        End If
    Next i
    'update DB
    If cboBranch.ListIndex = 0 Then     'MOTO-R MAIN"
        LYRcon.Execute "UPDATE spareparts SET REMARKS = 'TIRES' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 1 Then 'MOTO-R ULAS"
        conUlas.Execute "UPDATE spareparts SET REMARKS = 'TIRES' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 2 Then 'MC PARTS DEPOT"
        conToril.Execute "UPDATE spareparts SET REMARKS = 'TIRES' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 3 Then 'YAMAHA MOTORHUB"
        conYM.Execute "UPDATE spareparts SET REMARKS = 'TIRES' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 4 Then  'VAN SELLING"
        conVAN1.Execute "UPDATE spareparts SET REMARKS = 'TIRES' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 5 Then  'VAN SELLING2"
        conVAN1.Execute "UPDATE spareparts SET REMARKS = 'TIRES' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 6 Then 'MRC2 GANTE"
        conMRC1.Execute "UPDATE spareparts SET REMARKS = 'TIRES' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 7 Then 'panabo"
        conPanabo.Execute "UPDATE spareparts SET REMARKS = 'TIRES' Where spareparts.SPAREPARTID = " & MyID & ";"
    End If

    MsgBox flxSP.TextMatrix(ActRow, 2) & " successfully updated to TIRES"

End Sub

Private Sub mnuClassYamaha_Click()
    Dim Ans As Integer, i As Integer, MyID As Long
    
    Ans = MsgBox("Do you want to update " & flxSP.TextMatrix(ActRow, 2) & " classification to YAMAHA GENUINE?", vbQuestion + vbYesNo, "Update")
    If Ans = vbNo Then Exit Sub
    
    MyID = Val(flxSP.TextMatrix(ActRow, 8))
    For i = 1 To flxSP.Rows - 1
        If Val(flxSP.TextMatrix(i, 8)) = MyID Then
            flxSP.TextMatrix(i, 4) = "YAMAHA GENUINE"
        End If
    Next i
    'update DB
    If cboBranch.ListIndex = 0 Then     'MOTO-R MAIN"
        LYRcon.Execute "UPDATE spareparts SET REMARKS = 'YAMAHA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 1 Then 'MOTO-R ULAS"
        conUlas.Execute "UPDATE spareparts SET REMARKS = 'YAMAHA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 2 Then 'MC PARTS DEPOT"
        conToril.Execute "UPDATE spareparts SET REMARKS = 'YAMAHA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 3 Then 'YAMAHA MOTORHUB"
        conYM.Execute "UPDATE spareparts SET REMARKS = 'YAMAHA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 4 Then  'VAN SELLING"
        conVAN1.Execute "UPDATE spareparts SET REMARKS = 'YAMAHA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 5 Then  'VAN SELLING2"
        conVAN1.Execute "UPDATE spareparts SET REMARKS = 'YAMAHA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 6 Then 'MRC2 GANTE"
        conMRC1.Execute "UPDATE spareparts SET REMARKS = 'YAMAHA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf cboBranch.ListIndex = 7 Then 'panabo"
        conPanabo.Execute "UPDATE spareparts SET REMARKS = 'YAMAHA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    End If

    MsgBox flxSP.TextMatrix(ActRow, 2) & " successfully updated to YAMAHA GENUINE"

End Sub

Private Sub mnuDetails_Click()
    Dim rsPart As New ADODB.Recordset
    Dim Qry As String, MyRow As Integer, myRem As String, Ans As Integer, MyTotal As Currency
    Dim DateStart As Date, DateEnd As Date
    
    Ans = MsgBox("Do you want to Display Details?", vbQuestion + vbYesNo, "Sales Distribution")
    If Ans = vbNo Then Exit Sub
    
    flxSP.Rows = 1
    MyRow = flxSD.Row
    myRem = flxSD.TextMatrix(0, flxSD.Col)
    If InStr(myRem, "GEN.") > 0 Then
        myRem = Replace$(myRem, "GEN.", "GENUINE")
    ElseIf InStr(myRem, "APPAREL") > 0 Then
        myRem = "RIDER APPARELS"
    End If
    MyTotal = 0
    TabSales.Enabled = False
    cmdExit.Enabled = False
    DateStart = DateSerial(Int(cboYear.Text), MyRow, 1)
    DateEnd = DateAdd("d", -1, DateAdd("m", 1, DateValue(DateStart)))
    
    If myRem = "OTHERS" Then
        Qry = "SELECT spitr.TRANSDATE, spitr.REFERENCE, spareparts.PARTNUMBER, spareparts.DESCRIPTION, spareparts.REMARKS, spitr.OUT,  spitr.PRICE, spareparts.SPAREPARTID  From spitr INNER JOIN spareparts ON (spitr.SPAREPARTID = spareparts.SPAREPARTID) " & _
              "Where spitr.TRANSDATE >= '" & Format(DateStart, "yyyy-mm-dd") & "' AND spitr.TRANSDATE <= '" & Format(DateEnd, "yyyy-mm-dd") & "' AND spitr.OUT > 0 AND spitr.REFERENCE LIKE 'SI#%' AND " & _
              "spareparts.REMARKS <> 'HONDA GENUINE' AND spareparts.REMARKS <> 'YAMAHA GENUINE' AND spareparts.REMARKS <> 'SUZUKI GENUINE' AND spareparts.REMARKS <> 'REPLACEMENT' AND " & _
              "spareparts.REMARKS <> 'ACCESSORIES' AND spareparts.REMARKS <> 'PERFORMANCE' AND spareparts.REMARKS <> 'RIDER APPARELS' AND spareparts.REMARKS <> 'TIRES' AND spareparts.REMARKS <> 'OIL' ORDER BY spitr.TRANSDATE, spitr.REFERENCE;"
    Else
        Qry = "SELECT spitr.TRANSDATE, spitr.REFERENCE, spareparts.PARTNUMBER, spareparts.DESCRIPTION, spareparts.REMARKS, spitr.OUT,  spitr.PRICE, spareparts.SPAREPARTID FROM spitr INNER JOIN spareparts ON (spitr.SPAREPARTID = spareparts.SPAREPARTID) " & _
                "WHERE spitr.TRANSDATE >= '" & Format(DateStart, "yyyy-mm-dd") & "' AND spitr.TRANSDATE <= '" & Format(DateEnd, "yyyy-mm-dd") & "' AND  spareparts.REMARKS = '" & myRem & "' AND spitr.OUT > 0 AND spitr.REFERENCE LIKE 'SI#%' ORDER BY spitr.TRANSDATE, spitr.REFERENCE;"
    End If
    If cboBranch.ListIndex = 0 Then     'MOTO-R MAIN"
        rsPart.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    ElseIf cboBranch.ListIndex = 1 Then 'MOTO-R ULAS"
        rsPart.Open Qry, conUlas, adOpenKeyset, adLockOptimistic
    ElseIf cboBranch.ListIndex = 2 Then 'MC PARTS DEPOT"
        rsPart.Open Qry, conToril, adOpenKeyset, adLockOptimistic
    ElseIf cboBranch.ListIndex = 3 Then 'YAMAHA MOTORHUB"
        rsPart.Open Qry, conYM, adOpenKeyset, adLockOptimistic
    ElseIf cboBranch.ListIndex = 4 Then  'VAN SELLING1"
        rsPart.Open Qry, conVAN1, adOpenKeyset, adLockOptimistic
    ElseIf cboBranch.ListIndex = 5 Then  'VAN SELLING2"
        rsPart.Open Qry, conVAN2, adOpenKeyset, adLockOptimistic
    ElseIf cboBranch.ListIndex = 6 Then 'MRC2 GANTE"
        rsPart.Open Qry, conMRC1, adOpenKeyset, adLockOptimistic
    ElseIf cboBranch.ListIndex = 7 Then 'PANABO
        rsPart.Open Qry, conPanabo, adOpenKeyset, adLockOptimistic
    End If
    With rsPart
        If .RecordCount > 0 Then
            .MoveFirst
            Pbar.Value = 0
            Pbar.Max = .RecordCount
            Pbar.Visible = True
            Do While Not .EOF
                flxSP.AddItem Format(!TRANSDATE, "mm/dd/yyyy") & vbTab & !REFERENCE & vbTab & !PartNumber & vbTab & !Description & vbTab & IIf(IsNull(!Remarks), "", !Remarks) & vbTab & !OUT & vbTab & Format(!Price, "#,##0.00") & vbTab & Format(!OUT * !Price, "#,##0.00") & vbTab & !SparepartID
                MyTotal = MyTotal + (!OUT * !Price)
                Pbar.Value = Pbar.Value + 1
                DoEvents
                .MoveNext
            Loop
            flxSP.AddItem "SALES"
            flxSP.TextMatrix(flxSP.Rows - 1, 6) = "TOTAL"
            flxSP.TextMatrix(flxSP.Rows - 1, 7) = Format(MyTotal, "#,##0.00")
            flxSP.Col = 1: flxSP.Row = flxSP.Rows - 1: flxSP.ColSel = 8: flxSP.RowSel = flxSP.Rows - 1: flxSP.FillStyle = flexFillRepeat: flxSP.CellBackColor = RGB(250, 250, 180)
            flxSP.Col = 2
            Pbar.Visible = True
        Else
            flxSP.AddItem ""
        End If
        .Close
    End With
    Pbar.Visible = False
    'Display Sales Return
    MyTotal = 0
    If myRem = "OTHERS" Then
        Qry = "SELECT spitr.TRANSDATE, spitr.REFERENCE, spareparts.PARTNUMBER, spareparts.DESCRIPTION, spareparts.REMARKS, spitr.RECEIVE,  spitr.PRICE, spareparts.SPAREPARTID  From spitr INNER JOIN spareparts ON (spitr.SPAREPARTID = spareparts.SPAREPARTID) " & _
              "Where spitr.TRANSDATE >= '" & Format(DateStart, "yyyy-mm-dd") & "' AND spitr.TRANSDATE <= '" & Format(DateEnd, "yyyy-mm-dd") & "' AND spitr.RECEIVE > 0 AND spitr.TRANSTYPE = 'SALES RETURN' AND " & _
              "spareparts.REMARKS <> 'HONDA GENUINE' AND spareparts.REMARKS <> 'YAMAHA GENUINE' AND spareparts.REMARKS <> 'SUZUKI GENUINE' AND spareparts.REMARKS <> 'REPLACEMENT' AND " & _
              "spareparts.REMARKS <> 'ACCESSORIES' AND spareparts.REMARKS <> 'PERFORMANCE' AND spareparts.REMARKS <> 'RIDER APPARELS' AND spareparts.REMARKS <> 'TIRES' AND spareparts.REMARKS <> 'OIL' ORDER BY spitr.TRANSDATE, spitr.REFERENCE;"
    Else
        Qry = "SELECT spitr.TRANSDATE, spitr.REFERENCE, spareparts.PARTNUMBER, spareparts.DESCRIPTION, spareparts.REMARKS, spitr.RECEIVE,  spitr.PRICE, spareparts.SPAREPARTID FROM spitr INNER JOIN spareparts ON (spitr.SPAREPARTID = spareparts.SPAREPARTID) " & _
              "WHERE spitr.TRANSDATE >= '" & Format(DateStart, "yyyy-mm-dd") & "' AND spitr.TRANSDATE <= '" & Format(DateEnd, "yyyy-mm-dd") & "' AND  spareparts.REMARKS = '" & myRem & "' AND spitr.RECEIVE > 0 AND spitr.TRANSTYPE = 'SALES RETURN' ORDER BY spitr.TRANSDATE, spitr.REFERENCE;"
    End If
    If cboBranch.ListIndex = 0 Then     'MOTO-R MAIN"
        rsPart.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    ElseIf cboBranch.ListIndex = 1 Then 'MOTO-R ULAS"
        rsPart.Open Qry, conUlas, adOpenKeyset, adLockOptimistic
    ElseIf cboBranch.ListIndex = 2 Then 'MC PARTS DEPOT"
        rsPart.Open Qry, conToril, adOpenKeyset, adLockOptimistic
    ElseIf cboBranch.ListIndex = 3 Then 'YAMAHA MOTORHUB"
        rsPart.Open Qry, conYM, adOpenKeyset, adLockOptimistic
    ElseIf cboBranch.ListIndex = 4 Then  'VAN SELLING1"
        rsPart.Open Qry, conVAN1, adOpenKeyset, adLockOptimistic
    ElseIf cboBranch.ListIndex = 5 Then  'VAN SELLING2"
        rsPart.Open Qry, conVAN2, adOpenKeyset, adLockOptimistic
    ElseIf cboBranch.ListIndex = 6 Then 'MRC2 GANTE"
        rsPart.Open Qry, conMRC1, adOpenKeyset, adLockOptimistic
    ElseIf cboBranch.ListIndex = 7 Then 'PANABO
        rsPart.Open Qry, conPanabo, adOpenKeyset, adLockOptimistic
    End If
    With rsPart
        If .RecordCount > 0 Then
            .MoveFirst
            Pbar.Value = 0
            Pbar.Max = .RecordCount
            Pbar.Visible = True
            Do While Not .EOF
                flxSP.AddItem Format(!TRANSDATE, "mm/dd/yyyy") & vbTab & !REFERENCE & vbTab & !PartNumber & vbTab & !Description & vbTab & IIf(IsNull(!Remarks), "", !Remarks) & vbTab & !RECEIVE & vbTab & Format(!Price, "#,##0.00") & vbTab & Format(!RECEIVE * !Price, "#,##0.00") & vbTab & !SparepartID
                MyTotal = MyTotal + (!RECEIVE * !Price)
                Pbar.Value = Pbar.Value + 1
                DoEvents
                .MoveNext
            Loop
            flxSP.AddItem "SALES RETURN"
            flxSP.TextMatrix(flxSP.Rows - 1, 6) = "TOTAL"
            flxSP.TextMatrix(flxSP.Rows - 1, 7) = Format(MyTotal, "#,##0.00")
            flxSP.Col = 1: flxSP.Row = flxSP.Rows - 1: flxSP.ColSel = 8: flxSP.RowSel = flxSP.Rows - 1: flxSP.FillStyle = flexFillRepeat: flxSP.CellBackColor = RGB(250, 250, 180)
            flxSP.Col = 2
            Pbar.Visible = False
        End If
        .Close
    End With
    
    TabSales.Enabled = True
    cmdExit.Enabled = True
    
    
End Sub

Private Sub mnuVerify_Click()
    Dim Qry As String, MyRow As Integer, myRem As String, Ans As Integer
    Dim DateStart As Date, DateEnd As Date
    
    Ans = MsgBox("Do you want to Verify Spare Parts Classification?", vbQuestion + vbYesNo, "Verify Classification")
    If Ans = vbNo Then Exit Sub
    
    MyRow = flxSD.Row
    myRem = flxSD.TextMatrix(0, flxSD.Col)
    DateStart = DateSerial(Int(cboYear.Text), MyRow, 1)
    DateEnd = DateAdd("d", -1, DateAdd("m", 1, DateValue(DateStart)))
    
    If myRem = "OTHERS" Then
        SalesDistClass.spQry = "SELECT DISTINCT spareparts.PARTNUMBER, spareparts.DESCRIPTION, spareparts.REMARKS, spitr.SPAREPARTID From spitr INNER JOIN spareparts ON (spitr.SPAREPARTID = spareparts.SPAREPARTID) " & _
        "Where spitr.TRANSDATE >= '" & Format(DateStart, "yyyy-mm-dd") & "' AND spitr.TRANSDATE <= '" & Format(DateEnd, "yyyy-mm-dd") & "' AND spitr.OUT > 0 AND spitr.REFERENCE LIKE 'SI#%' AND spareparts.REMARKS <> 'HONDA GENUINE' AND spareparts.REMARKS <> 'YAMAHA GENUINE' AND " & _
        "spareparts.REMARKS <> 'SUZUKI GENUINE' AND spareparts.REMARKS <> 'REPLACEMENT' AND spareparts.REMARKS <> 'ACCESSORIES' AND spareparts.REMARKS <> 'PERFORMANCE' AND spareparts.REMARKS <> 'RIDER APPARELS' Order By spareparts.PARTNUMBER;"
    Else
        SalesDistClass.spQry = "SELECT DISTINCT spareparts.PARTNUMBER, spareparts.DESCRIPTION, spareparts.REMARKS, spitr.SPAREPARTID From spitr INNER JOIN spareparts ON (spitr.SPAREPARTID = spareparts.SPAREPARTID) " & _
        "Where spitr.TRANSDATE >= '" & Format(DateStart, "yyyy-mm-dd") & "' AND spitr.TRANSDATE <= '" & Format(DateEnd, "yyyy-mm-dd") & "' AND spareparts.REMARKS = '" & myRem & "' AND spitr.OUT > 0 AND spitr.REFERENCE LIKE 'SI#%' Order By spareparts.PARTNUMBER;"
    End If
    SDCloaded = True
    SalesDistClass.Show
End Sub

Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer
    
    For i = 0 To 10
        flxSD.ColWidth(i) = flxSD.Width * 0.088
    Next i
    flxSD.Height = flxSD.CellHeight * 14
    Frame2.Top = flxSD.Top + flxSD.Height + 100
    Frame2.Height = TabSales.Height - (flxSD.Top + flxSD.Height + 400)
    With flxSP
        .ColWidth(0) = .Width * 0.1
        .ColWidth(1) = .Width * 0.1
        .ColWidth(2) = .Width * 0.15
        .ColWidth(3) = .Width * 0.2
        .ColWidth(4) = .Width * 0.15
        .ColWidth(5) = .Width * 0.08
        .ColWidth(6) = .Width * 0.1
        .ColWidth(7) = .Width * 0.1
        .ColWidth(8) = 0
    End With
End Sub


Private Sub TabSales_Click(PreviousTab As Integer)
    
    If PreviousTab = 1 Then Exit Sub
    If TabSales.Tab = 1 Then
        Call ShowGraphMainAll
        lblBranch.Caption = cboBranch.Text
    End If
        
End Sub

Private Sub ShowGraphMainAll()
    Dim X() As Variant, Amt As Currency
    Dim i As Integer, j As Integer, Tmp As String
    
    
    ReDim X(0 To flxSD.Rows - 1, 0 To 10) As Variant
    'ReDim X(0 To flxSD.Rows - 1, 0 To 8) As Variant
    For i = 1 To 10
        X(0, i) = flxSD.TextMatrix(0, i)
    Next i
    
    For i = 1 To flxSD.Rows - 1
        For j = 0 To 10
            X(i, j) = IIf(j <> 0, Val(Replace(flxSD.TextMatrix(i, j), ",", "")), flxSD.TextMatrix(i, j))
        Next j
    Next i
    
    MSChart1.ChartData = X
    MSChart1.ChartType = VtChChartType2dLine

End Sub

Private Sub ShowGraphMainIndiv()
    Dim X() As Variant, Amt As Currency
    Dim i As Integer, j As Integer, MyCol As Integer
    
    
    ReDim X(0 To flxSD.Rows - 1, 0 To 1) As Variant
    MyCol = cboClass.ListIndex
    
    
    X(0, 1) = flxSD.TextMatrix(0, MyCol)
    For i = 1 To flxSD.Rows - 1
        X(i, 0) = flxSD.TextMatrix(i, 0)
    Next i
    For i = 1 To flxSD.Rows - 1
        X(i, 1) = Val(Replace(flxSD.TextMatrix(i, MyCol), ",", ""))
    Next i
    
    MSChart1.ChartData = X
    MSChart1.ChartType = VtChChartType2dLine

End Sub

