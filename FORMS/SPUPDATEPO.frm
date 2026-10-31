VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form SPUPDATEPO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Branch Purchase Order"
   ClientHeight    =   8625
   ClientLeft      =   6585
   ClientTop       =   1515
   ClientWidth     =   8175
   Icon            =   "SPUPDATEPO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   8175
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDetails 
      Height          =   945
      Left            =   210
      TabIndex        =   16
      Top             =   9840
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   1667
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin MSFlexGridLib.MSFlexGrid flxMain 
      Height          =   1005
      Left            =   270
      TabIndex        =   10
      Top             =   9120
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   1773
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
   End
   Begin VB.Frame Frame1 
      Height          =   8055
      Left            =   30
      TabIndex        =   0
      Top             =   540
      Width           =   8115
      Begin TabDlg.SSTab stabBk 
         Height          =   7365
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   12991
         _Version        =   393216
         Style           =   1
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         TabCaption(0)   =   "Branch P.O"
         TabPicture(0)   =   "SPUPDATEPO.frx":058A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblp"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label3"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "flexUpdatePO"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cmdUpdatePO"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "PrgBar"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "frRmBackup(1)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "flexAdd"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "chkMark"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).ControlCount=   9
         Begin VB.CheckBox chkMark 
            Height          =   225
            Left            =   180
            TabIndex        =   12
            Tag             =   "XYF"
            Top             =   2910
            Width           =   225
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexAdd 
            Height          =   1635
            Left            =   90
            TabIndex        =   9
            Top             =   2880
            Width           =   7545
            _ExtentX        =   13309
            _ExtentY        =   2884
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   5
         End
         Begin VB.Frame frRmBackup 
            Caption         =   "Select Branch to Update Purchase Order"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1875
            Index           =   1
            Left            =   90
            TabIndex        =   4
            Top             =   540
            Width           =   7575
            Begin VB.ComboBox cmbBranch 
               Height          =   315
               Left            =   1170
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   1380
               Width           =   3675
            End
            Begin lvButton.lvButtons_H cmdShowPO 
               Height          =   375
               Left            =   4920
               TabIndex        =   6
               Top             =   1350
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   661
               Caption         =   "Verify P.O"
               CapAlign        =   2
               BackStyle       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
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
               Image           =   "SPUPDATEPO.frx":05A6
               cBack           =   -2147483633
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Select Branch"
               Height          =   195
               Left            =   1170
               TabIndex        =   8
               Top             =   1140
               Width           =   1005
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Select Branch to Update Purchase Order"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   1170
               TabIndex        =   7
               Top             =   690
               Width           =   3690
            End
            Begin VB.Image Image1 
               Appearance      =   0  'Flat
               Height          =   720
               Left            =   150
               Picture         =   "SPUPDATEPO.frx":0B40
               Top             =   720
               Width           =   720
            End
         End
         Begin MSComctlLib.ProgressBar PrgBar 
            Height          =   195
            Left            =   90
            TabIndex        =   3
            Top             =   4590
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   344
            _Version        =   393216
            Appearance      =   1
         End
         Begin lvButton.lvButtons_H cmdUpdatePO 
            Height          =   375
            Left            =   4980
            TabIndex        =   11
            Top             =   2430
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   661
            Caption         =   "Update P.O"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            Image           =   "SPUPDATEPO.frx":493CA
            Enabled         =   0   'False
            cBack           =   -2147483633
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexUpdatePO 
            Height          =   2085
            Left            =   120
            TabIndex        =   17
            Top             =   5130
            Width           =   7395
            _ExtentX        =   13044
            _ExtentY        =   3678
            _Version        =   393216
            Cols            =   4
            FixedCols       =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   4
         End
         Begin VB.Label Label3 
            Caption         =   "Added P.O"
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
            Height          =   285
            Left            =   120
            TabIndex        =   18
            Top             =   4830
            Width           =   1335
         End
         Begin VB.Label lblp 
            Height          =   315
            Left            =   7200
            TabIndex        =   15
            Top             =   4530
            Width           =   465
         End
         Begin VB.Label Label1 
            Caption         =   "Purchase Order"
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
            Height          =   285
            Left            =   90
            TabIndex        =   14
            Top             =   2580
            Width           =   1335
         End
      End
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "E&xit"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Image           =   "SPUPDATEPO.frx":4A21C
      cBack           =   -2147483633
   End
   Begin VB.Label lblChk 
      Caption         =   "ü"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1860
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "SPUPDATEPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsUPO As New ADODB.Recordset
Dim MRC1Connect As Boolean, MRC2Connect As Boolean, UlasConnect As Boolean, TorilConnect As Boolean, Van2Connect As Boolean, StoTomasConnect As Boolean
Dim PanaboConnect As Boolean, BinosPolConnect As Boolean, YamahaConnect As Boolean, Van1Connect As Boolean, GensanConnect As Boolean, SamalConnect As Boolean, MarbelConnect As Boolean



Private Sub chkMark_Click()
 Dim i As Integer
    If chkMark.Value = 0 Then
        For i = 1 To flexAdd.Rows - 1
            flexAdd.TextMatrix(i, 0) = ""
        Next i
    Else
    With flexAdd
        For i = 1 To .Rows - 1
            'If Val(flexAdd.TextMatrix(i, 7)) > 0 Then
                .Row = i
                .Col = 0
                .CellFontName = "WINGDINGS"
                .CellFontSize = 12
                .Text = lblChk.Caption
                .CellAlignment = 4
                .TextMatrix(i, 0) = lblChk.Caption
            'End If
        Next i
    End With
    End If
End Sub
Private Sub cmdShowPO_Click()
    Call LoadBranchPO
    If flexAdd.TextMatrix(1, 1) <> "" Then cmdUpdatePO.Enabled = True
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Sub LoadBranch()
    With cmbBranch
        If ConnectUlas Then
            .AddItem "MOTO-R ULAS"
            .ItemData(.ListCount - 1) = 13
            UlasConnect = True
        Else
            UlasConnect = False
        End If
        If ConnectYamaha Then
            .AddItem "YAMAHA MOTORHUB"
            .ItemData(.ListCount - 1) = 14
            YamahaConnect = True
        Else
            YamahaConnect = False
        End If
        If ConnectToril Then
            .AddItem "MC PARTS DEPOT"
            .ItemData(.ListCount - 1) = 15
            TorilConnect = True
        Else
            TorilConnect = False
        End If
'         .AddItem "VAN I"
'         .ItemData(.ListCount - 1) = 16
'         .AddItem "VAN II"
'         .ItemData(.ListCount - 1) = 17
        If ConnectPanabo Then
            .AddItem "MRC PANABO"
            .ItemData(.ListCount - 1) = 19
            PanaboConnect = True
        Else
            PanaboConnect = False
        End If
        If ConnectHondaPol Then
            .AddItem "BINOS POLOMOLOK"
            .ItemData(.ListCount - 1) = 20
            BinosPolConnect = True
        Else
            BinosPolConnect = False
        End If
        If ConnectMRC1 Then
            .AddItem "MRC1 SOBRECARY"
            .ItemData(.ListCount - 1) = 12
            MRC1Connect = True
        Else
            MRC1Connect = False
        End If
        If ConnectMRCGensan Then
            .AddItem "MRC GENSAN"
            .ItemData(.ListCount - 1) = 23
            GensanConnect = True
        Else
            GensanConnect = False
        End If
        If ConnectStoTomas Then
            .AddItem "BINOS STO. TOMAS"
            .ItemData(.ListCount - 1) = 24
            StoTomasConnect = True
        Else
            StoTomasConnect = False
        End If
        If ConnectSamal Then
            .AddItem "BINOS BABAK SAMAL"
            .ItemData(.ListCount - 1) = 25
            SamalConnect = True
        Else
            SamalConnect = False
        End If
            
        If ConnectMarbel Then
            .AddItem "MOTORHUB MARBEL"
            .ItemData(.ListCount - 1) = 26
            MarbelConnect = True
        Else
            MarbelConnect = False
        End If
        
    End With
End Sub

Sub LoadBranchPO()
    Dim rsChk As New ADODB.Recordset
    Dim Qry As String
    Dim iRw As Long
    
    With rsUPO
        'Select Order table from Branches
        Qry = "SELECT sporder.ORDERDATE, sporder.REFERENCE, sporder.ORDERTYPE " & _
              "From sporder where sporder.CONFIRMNO IS NULL AND sporder.ORDERDATE > '2011/07/01';"
              
        If .State = adStateOpen Then .Close
        
        'Select Branches to open
        If cmbBranch.Text = "MOTO-R ULAS" Then
            If UlasConnect Then
                .Open Qry, conUlas, adOpenStatic, adLockOptimistic
            End If
        ElseIf cmbBranch.Text = "YAMAHA MOTORHUB" Then
            If YamahaConnect Then
                .Open Qry, conYM, adOpenStatic, adLockOptimistic
            End If
        ElseIf cmbBranch.Text = "MC PARTS DEPOT" Then
            If TorilConnect Then
                .Open Qry, conToril, adOpenStatic, adLockOptimistic
            End If
        ElseIf cmbBranch.Text = "MRC PANABO" Then
            If PanaboConnect Then
                .Open Qry, conPanabo, adOpenStatic, adLockOptimistic
            End If
        ElseIf cmbBranch.Text = "BINOS POLOMOLOK" Then
            If BinosPolConnect Then
                .Open Qry, conPol, adOpenStatic, adLockOptimistic
            End If
        ElseIf cmbBranch.Text = "MRC GENSAN" Then
            If GensanConnect Then
                .Open Qry, conGenSan, adOpenKeyset, adLockOptimistic
            End If
        ElseIf cmbBranch.Text = "BINOS STO. TOMAS" Then
            If StoTomasConnect Then
                .Open Qry, conStoTomas, adOpenKeyset, adLockOptimistic
            End If
        ElseIf cmbBranch.Text = "BINOS BABAK SAMAL" Then
            If SamalConnect Then
                .Open Qry, conSamal, adOpenKeyset, adLockOptimistic
            End If
        ElseIf cmbBranch.Text = "MOTORHUB MARBEL" Then
            If MarbelConnect Then
                .Open Qry, conMarbel, adOpenKeyset, adLockOptimistic
            End If
    
        ElseIf cmbBranch.Text = "MRC1 SOBRECARY" Then
            If MRC1Connect Then
                Qry = "SELECT sporder.ORDERDATE, sporder.REFERENCE, sporder.ORDERTYPE " & _
                      "From sporder where sporder.CONFIRMNO IS NULL AND sporder.ORDERDATE > '2011/07/01' AND sporder.SUPPLIERID = 119;"
                .Open Qry, conMRC1, adOpenStatic, adLockOptimistic
            End If
        End If
        
        If .RecordCount = 0 Then: Exit Sub
        .MoveFirst
        flxMain.Rows = 1
        flexAdd.Rows = 1
        'Add P.O record to flex
        PrgBar.Visible = True
        lblp.Visible = True
        PrgBar.Value = 0
        PrgBar.Max = .RecordCount
        .MoveFirst
        For iRw = 1 To .RecordCount
                flxMain.AddItem !ORDERDATE & vbTab & cmbBranch.Text & vbTab & !REFERENCE & vbTab & !ORDERTYPE
                
                'Verify records if already added
                Qry = "SELECT * From sporderbranch where sporderbranch.REFERENCE = '" & flxMain.TextMatrix(iRw, 2) & "' AND sporderbranch.BRANCHID = " & cmbBranch.ItemData(cmbBranch.ListIndex) & ";"
                If rsChk.State = adStateOpen Then rsChk.Close
                rsChk.Open Qry, LYRcon, adOpenStatic, adLockOptimistic
                If rsChk.RecordCount = 0 Then
                     'Add record if not Addedd
                     flexAdd.AddItem vbTab & !ORDERDATE & vbTab & cmbBranch.Text & vbTab & !REFERENCE & vbTab & !ORDERTYPE
                 End If
             .MoveNext
              PrgBar.Value = PrgBar.Value + 1
              PrcntRec = PrgBar.Value / PrgBar.Max
              lblp.Caption = Format(PrcntRec, "###%")
              
        Next iRw
        
        lblp.Caption = Format(PrcntRec, "100%")
        PrgBar.Visible = False
        lblp.Visible = False
        
        Call FormatFlxAdd
    End With
    
End Sub

'Sub CheckIfLoaded()
'    Dim rsChk As New ADODB.Recordset
'    Dim nrw As Long, Qry As String
'        With cmbBranch
'            For nrw = 1 To flxMain.Rows
'                Qry = ""
'                    If rsChk.State = adStateOpen Then rsChk.Close
'                         rsChk.Open Qry, LYRcon, adOpenStatic, adLockReadOnly
                    
'                    If rsChk.RecordCount = 0 Then
'                        flexAdd.AddItem ""
'                    End If
'             Next nrw
'        End With
'End Sub

Private Sub cmdUpdatePO_Click()
    Call UpDatePO
    cmdUpdatePO.Enabled = False
End Sub

Sub UpDatePO()
    Dim rsCup As New ADODB.Recordset
    Dim iRw As Long, uRw As Long, POTotAmount As Currency
    Dim rsDtl As New ADODB.Recordset
    
    With flexAdd
        PrgBar.Value = 0
        PrgBar.Max = .Rows - 1
        
        For iRw = 1 To .Rows - 1
            
            If .TextMatrix(iRw, 0) <> "" Then
                flxDetails.Clear
                flxDetails.Rows = 1
                LYRcon.Execute "Insert Into sporderbranch(ORDERDATE, BRANCHNAME, REFERENCE, ORDERTYPE, BRANCHID)VALUES('" & Format(.TextMatrix(iRw, 1), "yyyy-mm-dd") & "', '" & .TextMatrix(iRw, 2) & "', '" & .TextMatrix(iRw, 3) & "', '" & .TextMatrix(iRw, 4) & "', " & cmbBranch.ItemData(cmbBranch.ListIndex) & ");"
                DoEvents
                Qry = "SELECT spitr.REFERENCE, spitr.SPAREPARTID, spitr.TRANSTYPE, spitr.`ORDER`, spitr.PRICE " & _
                    "From spitr Where spitr.REFERENCE = '" & .TextMatrix(iRw, 3) & "' AND spitr.TRANSTYPE <> 'CANCELLED';"
                If rsDtl.State = adStateOpen Then rsDtl.Close
                    'If cmbBranch.ItemData(cmbBranch.ListIndex) = 13 Then
                    '    rsDtl.Open Qry, BrchConMU, adOpenStatic, adLockOptimistic
                    'ElseIf cmbBranch.ItemData(cmbBranch.ListIndex) = 14 Then
                    '    rsDtl.Open Qry, BrchConYM, adOpenStatic, adLockOptimistic
                    'ElseIf cmbBranch.ItemData(cmbBranch.ListIndex) = 15 Then
                    '    rsDtl.Open Qry, BrchConMD, adOpenStatic, adLockOptimistic
                    'End If
                    'Select Branches to open
                    If cmbBranch.Text = "MOTO-R ULAS" Then
                        If UlasConnect Then
                            rsDtl.Open Qry, conUlas, adOpenStatic, adLockOptimistic
                        End If
                    ElseIf cmbBranch.Text = "YAMAHA MOTORHUB" Then
                        If YamahaConnect Then
                            rsDtl.Open Qry, conYM, adOpenStatic, adLockOptimistic
                        End If
                    ElseIf cmbBranch.Text = "MC PARTS DEPOT" Then
                        If TorilConnect Then
                            rsDtl.Open Qry, conToril, adOpenStatic, adLockOptimistic
                        End If
                    ElseIf cmbBranch.Text = "MRC PANABO" Then
                        If PanaboConnect Then
                            rsDtl.Open Qry, conPanabo, adOpenStatic, adLockOptimistic
                        End If
                    ElseIf cmbBranch.Text = "BINOS POLOMOLOK" Then
                        If BinosPolConnect Then
                            rsDtl.Open Qry, conPol, adOpenStatic, adLockOptimistic
                        End If
                    ElseIf cmbBranch.Text = "MRC1 SOBRECARY" Then
                        If MRC1Connect Then
                            rsDtl.Open Qry, conMRC1, adOpenStatic, adLockOptimistic
                        End If
                    ElseIf cmbBranch.Text = "MRC GENSAN" Then
                        If GensanConnect Then
                            rsDtl.Open Qry, conGenSan, adOpenKeyset, adLockOptimistic
                        End If
                    ElseIf cmbBranch.Text = "BINOS STO. TOMAS" Then
                        If StoTomasConnect Then
                            rsDtl.Open Qry, conStoTomas, adOpenKeyset, adLockOptimistic
                        End If
                    ElseIf cmbBranch.Text = "BINOS BABAK SAMAL" Then
                        If SamalConnect Then
                            rsDtl.Open Qry, conSamal, adOpenKeyset, adLockOptimistic
                        End If
                    ElseIf cmbBranch.Text = "MOTORHUB MARBEL" Then
                        If MarbelConnect Then
                            rsDtl.Open Qry, conMarbel, adOpenKeyset, adLockOptimistic
                        End If
                    End If
                    
                
                Set flxDetails.Recordset = rsDtl
                DoEvents
                For uRw = 0 To flxDetails.Rows - 1
                    POTotAmount = Val(flxDetails.TextMatrix(uRw, 3)) * Val(flxDetails.TextMatrix(uRw, 4))
                    LYRcon.Execute "INSERT INTO spbranchdtl(REFERENCE, SPAREPARTID, TRANSTYPE, `ORDER`, BRANCHNAME, ORDERBAL, PRICE, TOTAMOUNT)" & _
                        "VALUES('" & flxDetails.TextMatrix(uRw, 0) & "', " & Val(flxDetails.TextMatrix(uRw, 1)) & ", '" & flxDetails.TextMatrix(uRw, 2) & "', " & _
                        "" & Val(flxDetails.TextMatrix(uRw, 3)) & ", '" & cmbBranch.Text & "', " & Val(flxDetails.TextMatrix(uRw, 3)) & ", " & Val(flxDetails.TextMatrix(uRw, 4)) & ", " & POTotAmount & ");"
                Next uRw
                DoEvents
                'Verify records if already added
                Qry = "SELECT sporderbranch.ORDERDATE, sporderbranch.BRANCHNAME, sporderbranch.REFERENCE, sporderbranch.ORDERTYPE From sporderbranch where sporderbranch.REFERENCE = '" & .TextMatrix(iRw, 3) & "' AND sporderbranch.BRANCHID = " & cmbBranch.ItemData(cmbBranch.ListIndex) & ";"
                If rsCup.State = adStateOpen Then rsCup.Close
                rsCup.Open Qry, LYRcon, adOpenStatic, adLockOptimistic
                If rsCup.RecordCount > 0 Then
                     rsCup.MoveFirst
                     If flexUpdatePO.Rows = 2 And flexUpdatePO.TextMatrix(1, 1) = "" Then
                       flexUpdatePO.Rows = 1: Call FormatFlxPO
                     End If
                     flexUpdatePO.AddItem rsCup!ORDERDATE & vbTab & rsCup!BRANCHNAME & vbTab & rsCup!REFERENCE & vbTab & rsCup!ORDERTYPE
                     flexUpdatePO.FixedRows = 1
                End If
                
            End If
                PrgBar.Value = PrgBar.Value + 1
                PrcntRec = PrgBar.Value / PrgBar.Max
                lblp.Caption = Format(PrcntRec, "###%")
        Next iRw
        lblp.Caption = "100%"
        MsgBox ("P.O Successfully Update"), vbInformation, "P.O"
    End With
End Sub
Private Sub flexAdd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        With flexAdd
            If .TextMatrix(flexAdd.Row, 0) = "" Then
                .Col = 0
                .CellFontName = "WINGDINGS"
                .CellFontSize = 12
                .Text = lblChk.Caption
                .CellAlignment = 4
            Else
                .TextMatrix(.Row, 0) = ""
            End If
        End With
    End If
End Sub
Private Sub Form_Load()
    Call LoadBranch
    Call FormatFlxPO
    Call FormatFlxAdd
End Sub
Sub FormatFlx()
    With flxMain
        If .Rows = 1 Then .Rows = 2
         .FixedRows = 1
        .FormatString = "   P.O DATE |^ SUPPLIER NAME |^ REFERENCE NO. |^ ORDER TYPE"
        .ColWidth(0) = .Width * 0.18
        .ColWidth(1) = .Width * 0.3
        .ColWidth(2) = .Width * 0.3
        .ColWidth(3) = .Width * 0.3
    End With
End Sub
Sub FormatFlxPO()
    With flexUpdatePO
        'If .Rows = 1 Then .Rows = 2
        
        .FormatString = "   P.O DATE |^ SUPPLIER NAME |^ REFERENCE NO. |^ ORDER TYPE"
        .ColWidth(0) = .Width * 0.18
        .ColWidth(1) = .Width * 0.3
        .ColWidth(2) = .Width * 0.3
        .ColWidth(3) = .Width * 0.21
    End With
End Sub
Sub FormatFlxAdd()
    With flexAdd
        If .Rows = 1 Then .Rows = 2
         .FixedRows = 1
        .FormatString = "  |^   P.O DATE |^ SUPPLIER NAME |^ REFERENCE NO. |^ ORDER TYPE"
        .ColWidth(0) = .Width * 0.05
        .ColWidth(1) = .Width * 0.21
        .ColWidth(2) = .Width * 0.21
        .ColWidth(3) = .Width * 0.21
        .ColWidth(4) = .Width * 0.31
        
    End With
End Sub
