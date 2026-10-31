VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form SEToptions 
   Caption         =   "Options"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7890
   Icon            =   "SEToptionsOld.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPartsD 
      Height          =   1335
      Left            =   720
      TabIndex        =   28
      Top             =   3480
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2355
      _Version        =   393216
      BackColor       =   16777215
      FixedRows       =   0
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      GridColorUnpopulated=   16777215
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtVAT 
      Height          =   285
      Left            =   4080
      TabIndex        =   27
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   840
      Top             =   5040
   End
   Begin VB.ListBox lstChange 
      Height          =   2010
      ItemData        =   "SEToptionsOld.frx":058A
      Left            =   240
      List            =   "SEToptionsOld.frx":058C
      TabIndex        =   7
      Top             =   5880
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   3480
      TabIndex        =   1
      Tag             =   "XYF"
      Top             =   4800
      Width           =   4095
      Begin lvButton.lvButtons_H cmdApply 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Tag             =   "XYF"
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Caption         =   "&Apply"
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
         Image           =   "SEToptionsOld.frx":058E
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdClose 
         Height          =   375
         Left            =   2160
         TabIndex        =   3
         Tag             =   "XYF"
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Caption         =   "&Close"
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
         Image           =   "SEToptionsOld.frx":0B28
         cBack           =   -2147483633
      End
   End
   Begin TabDlg.SSTab tabOpt 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Tag             =   "XYF"
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Disbursement"
      TabPicture(0)   =   "SEToptionsOld.frx":10C2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraDisb"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Labor Charge"
      TabPicture(1)   =   "SEToptionsOld.frx":10DE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Customer Type"
      TabPicture(2)   =   "SEToptionsOld.frx":10FA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&Discount"
      TabPicture(3)   =   "SEToptionsOld.frx":1116
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame4 
         Height          =   4095
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   7335
         Begin VB.TextBox txtPartNo 
            Height          =   285
            Left            =   480
            TabIndex        =   26
            Top             =   2640
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.TextBox txtDPrice 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5400
            TabIndex        =   25
            Top             =   2640
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox txtDFctr 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4560
            MaxLength       =   50
            TabIndex        =   20
            Top             =   840
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.ComboBox cboSPType 
            Height          =   315
            Left            =   600
            TabIndex        =   19
            Top             =   840
            Visible         =   0   'False
            Width           =   3975
         End
         Begin MSFlexGridLib.MSFlexGrid flxWDisc 
            Height          =   1410
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   2487
            _Version        =   393216
            Cols            =   3
            FixedCols       =   0
            RowHeightMin    =   300
            BackColorBkg    =   -2147483633
            ScrollBars      =   2
            FormatString    =   "      |<     S P A R E P A R T    T Y P E                                     |^DISCOUNT FACTOR             "
         End
         Begin MSFlexGridLib.MSFlexGrid flxSPec 
            Height          =   1755
            Left            =   120
            TabIndex        =   23
            Top             =   2280
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   3096
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            RowHeightMin    =   300
            BackColorBkg    =   -2147483633
            ScrollBars      =   2
            FormatString    =   "      |<P A R T N U M B E R                                |^C O S T (VAT inclusive)|^Dicounted Price      "
         End
         Begin VB.Label Label3 
            Caption         =   "Specific Sparepart Discount"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   2040
            Width           =   5295
         End
         Begin VB.Label Label2 
            Caption         =   "Wholesale Discount Formula = COST * Discount Factor"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   240
            Width           =   4935
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4095
         Left            =   -74880
         TabIndex        =   14
         Top             =   360
         Width           =   7335
         Begin VB.ComboBox cboPayType 
            Height          =   315
            Left            =   4320
            TabIndex        =   17
            Top             =   840
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.TextBox txtCustType 
            Height          =   285
            Left            =   600
            MaxLength       =   50
            TabIndex        =   16
            Top             =   960
            Visible         =   0   'False
            Width           =   3615
         End
         Begin MSFlexGridLib.MSFlexGrid flxCust 
            Height          =   3810
            Left            =   120
            TabIndex        =   15
            Top             =   190
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   6720
            _Version        =   393216
            Cols            =   3
            FixedCols       =   0
            RowHeightMin    =   300
            BackColorBkg    =   -2147483633
            ScrollBars      =   2
            FormatString    =   "      |<     C U S T O M E R    T Y P E                                |^PAYMENT TYPE                            "
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4095
         Left            =   -74880
         TabIndex        =   9
         Top             =   360
         Width           =   7335
         Begin VB.TextBox txtJDesc 
            Height          =   285
            Left            =   240
            TabIndex        =   13
            Top             =   4320
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CheckBox chkLaborXst 
            Caption         =   "Check1"
            Height          =   195
            Left            =   360
            TabIndex        =   12
            Top             =   4800
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtJAmt 
            Height          =   285
            Left            =   2040
            TabIndex        =   11
            Top             =   4440
            Visible         =   0   'False
            Width           =   1455
         End
         Begin MSFlexGridLib.MSFlexGrid flxLabor 
            Height          =   3810
            Left            =   120
            TabIndex        =   10
            Top             =   195
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   6720
            _Version        =   393216
            Cols            =   3
            FixedCols       =   0
            RowHeightMin    =   280
            BackColorBkg    =   -2147483633
            ScrollBars      =   2
            FormatString    =   "    |<                                               J O B    D E S C R I P T I O N                  |>AMOUNT        "
         End
      End
      Begin VB.Frame fraDisb 
         Height          =   4095
         Left            =   -74880
         TabIndex        =   4
         Tag             =   "XYF"
         Top             =   360
         Width           =   7335
         Begin VB.CheckBox ChkAcctXst 
            Caption         =   "Check1"
            Height          =   255
            Left            =   4320
            TabIndex        =   8
            Top             =   3600
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.ComboBox cboAcctTitle 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3330
            Left            =   240
            Style           =   1  'Simple Combo
            TabIndex        =   5
            Tag             =   "XYF"
            Top             =   480
            Width           =   3975
         End
         Begin VB.Label Label1 
            Caption         =   "ACCOUNT TITLES"
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
            Left            =   240
            TabIndex        =   6
            Tag             =   "XYF"
            Top             =   240
            Width           =   3855
         End
      End
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   0
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   8010
      m_MDI_Height    =   6195
   End
   Begin VB.Menu mnuLaborCharge 
      Caption         =   "Labor Charge"
      Visible         =   0   'False
      Begin VB.Menu mnuLCadd 
         Caption         =   "Add Row"
      End
      Begin VB.Menu mnuLCdel 
         Caption         =   "Delete Row"
      End
   End
   Begin VB.Menu mnuCust 
      Caption         =   "Customer Type"
      Visible         =   0   'False
      Begin VB.Menu mnuCustAdd 
         Caption         =   "Add Row"
      End
      Begin VB.Menu mnuCustDel 
         Caption         =   "Delete Row"
      End
   End
   Begin VB.Menu mnuWDisc 
      Caption         =   "Wholesale Discount"
      Visible         =   0   'False
      Begin VB.Menu mnuWhlAdd 
         Caption         =   "Add Row"
      End
      Begin VB.Menu mnuWhlDel 
         Caption         =   "Delete Row"
      End
   End
   Begin VB.Menu mnuSpec 
      Caption         =   "Specific SP Discount"
      Visible         =   0   'False
      Begin VB.Menu mnuSpcAdd 
         Caption         =   "Add Row"
      End
      Begin VB.Menu mnuSpcDel 
         Caption         =   "Delete Row"
      End
   End
End
Attribute VB_Name = "SEToptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsPartsD As New ADODB.Recordset
Dim AccTtlSrch As Boolean, FullAccess As Boolean, PrtSrch As Boolean, PrtBln As Boolean
Dim ActRow As Integer, ActCol As Integer

Private Sub cboAcctTitle_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    
    If AccTtlSrch Then
        AccTtlSrch = False
        Exit Sub
    End If
    With cboAcctTitle
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            'rs.MoveFirst
            'rs.Move i
            'MySearchID = IIf(IsNull(rs.Fields(0)), 0, rs.Fields(0))
            ChkAcctXst.Value = 1
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                AccTtlSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            ChkAcctXst.Value = 0
            'MySearchID = 0
        End If
    End With
End Sub

Private Sub cboAcctTitle_Click()
    Call cboAcctTitle_Change
End Sub

Private Sub cboAcctTitle_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Ans As Integer
    
    If KeyCode = vbKeyDelete Then
        AccTtlSrch = True
        If cboAcctTitle.List(cboAcctTitle.ListIndex) = "" Then Exit Sub
        Ans = MsgBox("Do you want to remove " & cboAcctTitle.List(cboAcctTitle.ListIndex) & " as Account Title?", vbQuestion + vbYesNo, "Remove")
        If Ans = vbNo Then Exit Sub
        cboAcctTitle.RemoveItem cboAcctTitle.ListIndex
        lstChange.AddItem "Account Title"
    End If
    If KeyCode = vbKeyBack Then AccTtlSrch = True
End Sub

Private Sub cboAcctTitle_KeyPress(KeyAscii As Integer)
    Dim Ans As Integer, MyNdx As Integer
    
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If ChkAcctXst.Value = 0 Then
            Ans = MsgBox("Do you want to add " & cboAcctTitle.Text & "?", vbQuestion + vbYesNo, "New Account Title")
            If Ans = vbNo Then Exit Sub
            MyNdx = GetAcctTitleInsertNo(cboAcctTitle.Text)
            cboAcctTitle.AddItem cboAcctTitle.Text, MyNdx
            cboAcctTitle.Text = ""
            lstChange.AddItem "Account Title"
        End If
    End If
End Sub

Function GetAcctTitleInsertNo(MyTxt As String) As Integer
    Dim i As Integer
    
    With cboAcctTitle
        For i = 0 To .ListCount - 1
            If MyTxt < .List(i) Then
                GetAcctTitleInsertNo = i
                Exit Function
            End If
        Next i
        GetAcctTitleInsertNo = .ListCount
    End With
End Function


Private Sub cboPayType_GotFocus()
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboPayType.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboPayType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cboPayType.Text = Empty
        flxCust.TextMatrix(ActRow, 1) = cboPayType.Tag
        cboPayType.Tag = "skiplostfocus"
        flxCust.SetFocus
        flxCust.Col = 1
        cboPayType.Visible = False
    End If
End Sub

Private Sub cboPayType_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxCust.SetFocus
        If flxCust.TextMatrix(ActRow, 1) = "" Then
            flxCust.Col = 1
            flxCust.Row = ActRow
        ElseIf cboPayType.Text = "" Then
            flxCust.Col = 2
            flxCust.Row = ActRow
        ElseIf cboPayType.Text <> "" And flxCust.TextMatrix(ActRow, 1) <> "" And ActRow = flxCust.Rows - 1 Then
            flxCust.AddItem ""
            flxCust.Row = flxCust.Rows - 1
            flxCust.Col = 1
        Else
            flxCust.Row = ActRow + 1
            flxCust.Col = 1
        End If
    End If
End Sub

Private Sub cboPayType_LostFocus()
    
    If cboPayType.Tag = "skiplostfocus" Or cboPayType.Text = "" Then
        cboPayType.Tag = ""
        cboPayType.Visible = False
        Exit Sub
    End If
    If cboPayType.ListIndex = -1 Then
        MsgBox "Please select a payment type from the list!", vbExclamation + vbOKOnly, "Payment Type"
        Exit Sub
        cboPayType.Visible = False
    End If
    flxCust.TextMatrix(ActRow, 2) = cboPayType.Text
    cboPayType.Visible = False
    lstChange.AddItem "Customer Type"
End Sub

Private Sub cboSPtype_GotFocus()
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboSPType.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboSPtype_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cboSPType.Text = Empty
        flxWDisc.TextMatrix(ActRow, 1) = cboSPType.Tag
        cboSPType.Tag = "skiplostfocus"
        flxWDisc.SetFocus
        flxWDisc.Col = 1
        cboSPType.Visible = False
    End If
End Sub

Private Sub cboSPtype_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxWDisc.SetFocus
        If cboSPType.Text = "" Then
            flxWDisc.Col = 1
            flxWDisc.Row = ActRow
        ElseIf flxWDisc.TextMatrix(ActRow, 2) = "" Then
            flxWDisc.Col = 2
            flxWDisc.Row = ActRow
        ElseIf cboSPType.Text <> "" And flxWDisc.TextMatrix(ActRow, 2) <> "" And ActRow = flxWDisc.Rows - 1 Then
            flxWDisc.AddItem ""
            flxWDisc.Row = flxWDisc.Rows - 1
            flxWDisc.Col = 1
        Else
            flxWDisc.Row = ActRow + 1
            flxWDisc.Col = 1
        End If
    End If

End Sub

Private Sub cboSPType_LostFocus()
    If cboSPType.Tag = "skiplostfocus" Or cboSPType.Text = "" Then
        cboSPType.Tag = ""
        cboSPType.Visible = False
        Exit Sub
    End If
    If cboSPType.ListIndex = -1 Then
        MsgBox "Please select a Sparepart type from the list!", vbExclamation + vbOKOnly, "Sparepart Type"
        Exit Sub
        cboSPType.Visible = False
    End If
    flxWDisc.TextMatrix(ActRow, 1) = cboSPType.Text
    cboSPType.Visible = False
    lstChange.AddItem "Wholesale Discount"
End Sub

Private Sub cmdApply_Click()
    Dim rsSettings As New ADODB.Recordset
    Dim i As Integer, Ans As Integer
    Dim MyAcct As String
    'Check individual Objects for changes
    
    'Account Title
    If ChangesMade("Account Title") Then
        LYRcon.Execute "DELETE DISBURSEACCTDESC.* FROM DISBURSEACCTDESC;"
        With cboAcctTitle
            For i = 0 To .ListCount - 1
                MyAcct = Replace$(.List(i), "'", "`")
                LYRcon.Execute "INSERT INTO DISBURSEACCTDESC ( ACCOUNTDESCRIPTION, UPDATEDBY, LASTUPDATE ) " & _
                               "VALUES ('" & MyAcct & "', " & UserID & ", #" & Date & "#);"
            Next i
        End With
    End If
    If ChangesMade("Labor Charge") Then
        LYRcon.Execute "DELETE MCSERVICELABOR.* FROM MCSERVICELABOR;"
        With flxLabor
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) <> "" And .TextMatrix(i, 2) <> "" Then
                    MyAcct = Replace$(.TextMatrix(i, 1), "'", "`")
                    LYRcon.Execute "INSERT INTO MCSERVICELABOR ( JOBREQUEST, LABORCOST ) " & _
                                   "VALUES ('" & MyAcct & "', " & Val(Replace(.TextMatrix(i, 2), ",", "")) & ");"
                End If
            Next i
        End With
    End If
    
    If ChangesMade("Customer Type") Then
        LYRcon.Execute "DELETE SETTINGS.* From SETTINGS WHERE (((SETTINGS.DESCRIPTION)='Customer Type'));"
        With rsSettings
            .Open "Settings", LYRcon, adOpenKeyset, adLockOptimistic
            For i = 1 To flxCust.Rows - 1
                If flxCust.TextMatrix(i, 1) <> "" And flxCust.TextMatrix(i, 2) <> "" Then
                    .AddNew
                    !Description = "Customer Type"
                    MyAcct = Replace$(flxCust.TextMatrix(i, 1), "'", "`")
                    !Value = MyAcct
                    !Remarks = flxCust.TextMatrix(i, 2)
                    !UpdatedBy = UserID
                    !LastUpdate = Now
                    .Update
                End If
            Next i
            .Close
        End With
    End If
    
    If ChangesMade("Wholesale Discount") Then
        LYRcon.Execute "DELETE SETTINGS.* From SETTINGS WHERE (((SETTINGS.DESCRIPTION)='Wholesale Discount'));"
        With rsSettings
            .Open "Settings", LYRcon, adOpenKeyset, adLockOptimistic
            For i = 1 To flxWDisc.Rows - 1
                If flxWDisc.TextMatrix(i, 1) <> "" And flxWDisc.TextMatrix(i, 2) <> "" Then
                    .AddNew
                    !Description = "Wholesale Discount"
                    !Value = flxWDisc.TextMatrix(i, 2)
                    !Remarks = flxWDisc.TextMatrix(i, 1)
                    !UpdatedBy = UserID
                    !LastUpdate = Now
                    .Update
                End If
            Next i
            .Close
        End With
    End If
    
    If ChangesMade("SP Disclusion") Then
        LYRcon.Execute "DELETE SETTINGS.* From SETTINGS WHERE (((SETTINGS.DESCRIPTION)='SP Disclusion'));"
        With rsSettings
            .Open "Settings", LYRcon, adOpenKeyset, adLockOptimistic
            For i = 1 To flxSPec.Rows - 1
                If flxSPec.TextMatrix(i, 4) <> "" And flxSPec.TextMatrix(i, 3) <> "" Then
                    .AddNew
                    !Description = "SP Disclusion"
                    !Value = flxSPec.TextMatrix(i, 3)
                    !Remarks = flxSPec.TextMatrix(i, 4)
                    !UpdatedBy = UserID
                    !LastUpdate = Now
                    .Update
                End If
            Next i
            .Close
        End With
    End If
    
    Ans = MsgBox("Changes has been Saved!" & vbCrLf & vbCrLf & "Do you want to Close Options?", vbQuestion + vbYesNo, "Close")
    If Ans = vbNo Then Exit Sub
    Unload Me
End Sub

Function ChangesMade(MyObj As String) As Boolean
    Dim i As Integer
    
    With lstChange
        For i = 0 To .ListCount - 1
            If MyObj = .List(i) Then
                ChangesMade = True
                Exit Function
            End If
        Next i
        ChangesMade = False
    End With
End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub flxCust_Click()
    ActRow = flxCust.Row
    ActCol = flxCust.Col
    
    If ActCol = 0 Or ActRow = 0 Or Not FullAccess Then Exit Sub
    Select Case ActCol
    Case 1
        With txtCustType
            If flxCust.TextMatrix(ActRow, 1) <> "" Then .Tag = flxCust.TextMatrix(ActRow, 1)
            .Top = flxCust.CellTop + flxCust.Top
            .Left = flxCust.CellLeft + flxCust.Left
            .Width = flxCust.CellWidth
            .Text = flxCust.Text
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
    Case 2
        With cboPayType
            If flxCust.TextMatrix(ActRow, 2) <> "" Then .Tag = flxCust.TextMatrix(ActRow, 2)
            .Top = flxCust.CellTop + flxCust.Top
            .Left = flxCust.CellLeft + flxCust.Left
            .Width = flxCust.CellWidth
            .Text = Replace(flxCust.Text, ",", "")
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
    End Select
End Sub

Private Sub flxCust_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxCust_Click
End Sub

Private Sub flxCust_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If FullAccess Then
        If Button = vbRightButton Then
            Call PopupMenu(mnuCust)
        End If
    End If
End Sub

Private Sub flxLabor_Click()
    ActRow = flxLabor.Row
    ActCol = flxLabor.Col
    
    If ActCol = 0 Or Not FullAccess Then Exit Sub
    Select Case ActCol
    Case 1
        With txtJDesc
            If flxLabor.TextMatrix(ActRow, 1) <> "" Then .Tag = flxLabor.TextMatrix(ActRow, 1)
            .Top = flxLabor.CellTop + flxLabor.Top
            .Left = flxLabor.CellLeft + flxLabor.Left
            .Width = flxLabor.CellWidth
            .Text = flxLabor.Text
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
    Case 2
        With txtJAmt
            If flxLabor.TextMatrix(ActRow, 2) <> "" Then .Tag = flxLabor.TextMatrix(ActRow, 2)
            .Top = flxLabor.CellTop + flxLabor.Top
            .Left = flxLabor.CellLeft + flxLabor.Left
            .Width = flxLabor.CellWidth
            .Text = Replace(flxLabor.Text, ",", "")
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
    End Select

End Sub

Private Sub flxLabor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxLabor_Click
End Sub

Private Sub flxLabor_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If FullAccess Then
        If Button = vbRightButton Then
            Call PopupMenu(mnuLaborCharge)
        End If
    End If
End Sub

Private Sub flxPartsD_Click()
    Dim Cost As Currency, CostVat As Currency
    
    With flxPartsD
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        PrtSrch = True
        txtPartNo.Text = .TextMatrix(.Row, 2)
        'Insert Additional retrieval here
        Cost = flxPartsD.TextMatrix(flxPartsD.Row, 3)
        CostVat = Cost + (Cost * (Val(txtVat.Text) / 100))
        flxSPec.TextMatrix(ActRow, 2) = Format(CostVat, "#,##0.00") 'Cost
        flxSPec.TextMatrix(ActRow, 4) = flxPartsD.TextMatrix(flxPartsD.Row, 1) 'SPID
    End With
End Sub

Private Sub flxPartsD_EnterCell()
    Dim Cost As Currency, CostVat As Currency
    
    With flxPartsD
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        PrtSrch = True
        txtPartNo.Text = .TextMatrix(.Row, 2)
        'Insert Additional retrieval here
        Cost = flxPartsD.TextMatrix(flxPartsD.Row, 3)
        CostVat = Cost + (Cost * (Val(txtVat.Text) / 100))
        flxSPec.TextMatrix(ActRow, 2) = Format(CostVat, "#,##0.00") 'Cost
        flxSPec.TextMatrix(ActRow, 4) = flxPartsD.TextMatrix(flxPartsD.Row, 1) 'SPID
    End With
End Sub

Private Sub flxPartsD_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Tmp() As String
    Dim i As Integer
    
    If KeyCode = vbKeyEscape Then
        If txtPartNo.Tag <> "" Then
            Tmp = Split(txtPartNo.Tag, "|")
            flxSPec.TextMatrix(ActRow, 1) = Tmp(0)
            flxSPec.TextMatrix(ActRow, 2) = Tmp(1)
            flxSPec.TextMatrix(ActRow, 4) = Tmp(2)
        End If
        txtPartNo.Visible = False
        flxPartsD.Visible = False
        flxSPec.SetFocus
        flxSPec.Col = 1
    End If

End Sub

Private Sub flxPartsD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtPartNo_LostFocus
        If flxSPec.TextMatrix(ActRow, 4) <> "" And flxSPec.TextMatrix(ActRow, 3) <> "" Then
            If ActRow = flxSPec.Rows - 1 Then flxSPec.AddItem ""
            flxSPec.SetFocus
            flxSPec.Row = flxSPec.Rows - 1
            flxSPec.Col = 1
        ElseIf flxSPec.TextMatrix(ActRow, 3) = "" Then   'No COST
            flxSPec.SetFocus
            flxSPec.Col = 3
        End If
        flxPartsD.Visible = False
    End If
End Sub

Private Sub flxSPec_Click()
    ActRow = flxSPec.Row
    ActCol = flxSPec.Col
    
    If ActCol = 0 Or Not FullAccess Then Exit Sub
    Select Case ActCol
    Case 1
        With txtPartNo
            If flxSPec.TextMatrix(ActRow, 1) <> "" Then .Tag = flxSPec.TextMatrix(ActRow, 1) & "|" & flxSPec.TextMatrix(ActRow, 2) & "|" & flxSPec.TextMatrix(ActRow, 4)
            .Top = flxSPec.CellTop + flxSPec.Top
            .Left = flxSPec.CellLeft + flxSPec.Left
            .Width = flxSPec.CellWidth
            .Text = flxSPec.Text
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
    Case 3
        With txtDPrice
            If flxSPec.TextMatrix(ActRow, 2) <> "" Then .Tag = flxSPec.TextMatrix(ActRow, 2)
            .Top = flxSPec.CellTop + flxSPec.Top
            .Left = flxSPec.CellLeft + flxSPec.Left
            .Width = flxSPec.CellWidth
            .Text = Replace(flxSPec.Text, ",", "")
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
    End Select

End Sub

Private Sub flxSPec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxSPec_Click
End Sub

Private Sub flxSPec_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If FullAccess Then
        If Button = vbRightButton Then
            Call PopupMenu(mnuSpec)
        End If
    End If
End Sub

Private Sub flxWDisc_Click()
    ActRow = flxWDisc.Row
    ActCol = flxWDisc.Col
    
    If ActCol = 0 Or Not FullAccess Then Exit Sub
    Select Case ActCol
    Case 1
        With cboSPType
            If flxWDisc.TextMatrix(ActRow, 1) <> "" Then .Tag = flxWDisc.TextMatrix(ActRow, 1)
            .Top = flxWDisc.CellTop + flxWDisc.Top
            .Left = flxWDisc.CellLeft + flxWDisc.Left
            .Width = flxWDisc.CellWidth
            .Text = flxWDisc.Text
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
    Case 2
        With txtDFctr
            If flxWDisc.TextMatrix(ActRow, 2) <> "" Then .Tag = flxWDisc.TextMatrix(ActRow, 2)
            .Top = flxWDisc.CellTop + flxWDisc.Top
            .Left = flxWDisc.CellLeft + flxWDisc.Left
            .Width = flxWDisc.CellWidth
            .Text = Replace(flxWDisc.Text, ",", "")
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
    End Select

End Sub

Private Sub flxWDisc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxWDisc_Click
End Sub

Private Sub flxWDisc_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If FullAccess Then
        If Button = vbRightButton Then
            Call PopupMenu(mnuWDisc)
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim Qry As String, MyAcc As String
    
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=96));"
    MyAcc = GetFieldValue(Qry)
    FullAccess = IIf(MyAcc = "FULL ACCESS", True, False)
    If Not FullAccess Then cmdApply.Enabled = False
    flxSPec.ColWidth(4) = 0
    Timer1.Enabled = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If rsPartsD.State = adStateOpen Then rsPartsD.Close
End Sub

Private Sub mnuCustAdd_Click()
    If flxCust.TextMatrix(flxCust.Rows - 1, 1) = "" Or flxCust.TextMatrix(flxCust.Rows - 1, 2) = "" Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        Exit Sub
    End If
    flxCust.AddItem ""
    flxCust.SetFocus
    flxCust.Row = flxCust.Rows - 1
    flxCust.Col = 1
End Sub

Private Sub mnuCustDel_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If Ans = vbYes Then
        If flxCust.Rows = 2 Then
            flxCust.Rows = 1
            flxCust.Rows = 2
        Else
            flxCust.RemoveItem ActRow
        End If
        lstChange.AddItem "Customer Type"
    End If
End Sub

Private Sub mnuLCadd_Click()
    If flxLabor.TextMatrix(flxLabor.Rows - 1, 1) = "" Or flxLabor.TextMatrix(flxLabor.Rows - 1, 2) = "" Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        Exit Sub
    End If
    flxLabor.AddItem ""
    flxLabor.SetFocus
    flxLabor.Row = flxLabor.Rows - 1
    flxLabor.Col = 1
End Sub

Private Sub mnuLCdel_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If Ans = vbYes Then
        If flxLabor.Rows = 2 Then
            flxLabor.Rows = 1
            flxLabor.Rows = 2
        Else
            flxLabor.RemoveItem ActRow
        End If
        lstChange.AddItem "Labor Charge"
    End If
End Sub

Private Sub mnuSpcAdd_Click()
    If flxSPec.TextMatrix(flxSPec.Rows - 1, 4) = "" Or flxSPec.TextMatrix(flxSPec.Rows - 1, 3) = "" Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        Exit Sub
    End If
    flxSPec.AddItem ""
    flxSPec.SetFocus
    flxSPec.Row = flxSPec.Rows - 1
    flxSPec.Col = 1
End Sub

Private Sub mnuSpcDel_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If Ans = vbYes Then
        If flxSPec.Rows = 2 Then
            flxSPec.Rows = 1
            flxSPec.Rows = 2
        Else
            flxSPec.RemoveItem ActRow
        End If
        lstChange.AddItem "SP Disclusion"
    End If
End Sub

Private Sub mnuWhlAdd_Click()
    If flxWDisc.TextMatrix(flxWDisc.Rows - 1, 1) = "" Or flxWDisc.TextMatrix(flxWDisc.Rows - 1, 2) = "" Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        Exit Sub
    End If
    flxWDisc.AddItem ""
    flxWDisc.SetFocus
    flxWDisc.Row = flxWDisc.Rows - 1
    flxWDisc.Col = 1
End Sub

Private Sub mnuWhlDel_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If Ans = vbYes Then
        If flxWDisc.Rows = 2 Then
            flxWDisc.Rows = 1
            flxWDisc.Rows = 2
        Else
            flxWDisc.RemoveItem ActRow
        End If
        lstChange.AddItem "Wholesale Discount"
    End If
End Sub

Private Sub Timer1_Timer()
    Dim rsAccTitle As New ADODB.Recordset
    Dim rsLaborChrg As New ADODB.Recordset
    Dim rsCType As New ADODB.Recordset
    Dim Qry As String, i As Integer, CostVat As Currency
    
    With rsAccTitle
        Qry = "SELECT DISBURSEACCTDESC.ACCOUNTDESCRIPTION From DISBURSEACCTDESC ORDER BY DISBURSEACCTDESC.ACCOUNTDESCRIPTION;"
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cboAcctTitle.AddItem !AccountDescription
                .MoveNext
            Loop
        End If
        .Close
    End With
    With rsLaborChrg
        Qry = "SELECT MCSERVICELABOR.* From MCSERVICELABOR ORDER BY MCSERVICELABOR.JOBREQUEST;"
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxLabor.Rows = 1
            Do While Not .EOF
                flxLabor.AddItem vbTab & !JobRequest & vbTab & IIf(IsNull(!LaborCost), "", Format(!LaborCost, "#,##0.00"))
                .MoveNext
            Loop
        End If
        .Close
    End With
    With rsCType
        Qry = "SELECT SETTINGS.* From SETTINGS " & _
              "WHERE (((SETTINGS.DESCRIPTION)='Customer Type')) ORDER BY SETTINGS.DESCRIPTION;"
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxCust.Rows = 1
            Do While Not .EOF
                flxCust.AddItem vbTab & !Value & vbTab & !Remarks
                .MoveNext
            Loop
        End If
        .Close
    End With
    With rsCType
        Qry = "SELECT SETTINGS.* From SETTINGS " & _
              "WHERE (((SETTINGS.DESCRIPTION)='Wholesale Discount')) ORDER BY SETTINGS.VALUE;"
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxWDisc.Rows = 1
            Do While Not .EOF
                flxWDisc.AddItem vbTab & !Remarks & vbTab & !Value
                .MoveNext
            Loop
        End If
        .Close
    End With
    With rsCType
        Qry = "SELECT SETTINGS.* From SETTINGS " & _
              "WHERE (((SETTINGS.DESCRIPTION)='SP Disclusion')) ORDER BY SETTINGS.VALUE;"
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxSPec.Rows = 1
            Do While Not .EOF
                flxSPec.AddItem vbTab & vbTab & vbTab & !Value & vbTab & !Remarks
                .MoveNext
            Loop
        End If
        .Close
    End With
        
        
    'Default Values:
    With cboPayType
        .AddItem "CASH"
        .AddItem "ON-ACCOUNT"
        .AddItem "CASH & ON-ACCOUNT"
    End With
    
    Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.COST From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"
    rsPartsD.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxPartsD.DataSource = rsPartsD
    flxPartsD.Refresh
    With flxPartsD
        .Width = flxSPec.ColWidth(1)
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = flxSPec.ColWidth(1)
        .ColWidth(3) = 0
    End With
    
    
    With rsCType
        Qry = "SELECT DISTINCT SUPPLIER.SPTYPE From SUPPLIER ORDER BY SUPPLIER.SPTYPE;"
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cboSPType.AddItem !SPtype
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    Qry = "SELECT SETTINGS.VALUE From SETTINGS WHERE (((SETTINGS.DESCRIPTION)='VAT'));"
    txtVat.Text = GetFieldValue(Qry)
    
    With flxSPec
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 4)) > 0 Then
                rsPartsD.MoveFirst
                rsPartsD.Find "[SPAREPARTID] = " & Val(.TextMatrix(i, 4))
                If Not rsPartsD.EOF Then
                    .TextMatrix(i, 1) = rsPartsD!PartNumber
                    CostVat = rsPartsD!Cost
                    CostVat = CostVat + (CostVat * Val(txtVat.Text) / 100)
                    .TextMatrix(i, 2) = Format(CostVat, "#,##0.00")
                End If
            End If
        Next i
    End With
    
    
    Timer1.Enabled = False
End Sub

Private Sub txtCustType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtCustType.Text = Empty
        flxCust.TextMatrix(ActRow, 1) = txtCustType.Tag
        txtCustType.Tag = "skiplostfocus"
        flxCust.SetFocus
        flxCust.Col = 1
        txtCustType.Visible = False
    End If
End Sub

Private Sub txtCustType_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxCust.SetFocus
        If txtCustType.Text = "" Then
            flxCust.Col = 1
            flxCust.Row = ActRow
        ElseIf flxCust.TextMatrix(ActRow, 2) = "" Then
            flxCust.Col = 2
            flxCust.Row = ActRow
        ElseIf flxCust.TextMatrix(ActRow, 2) <> "" And txtCustType.Text <> "" And ActRow = flxCust.Rows - 1 Then
            flxCust.AddItem ""
            flxCust.Row = flxCust.Rows - 1
            flxCust.Col = 1
        Else
            flxCust.Row = ActRow + 1
            flxCust.Col = 1
        End If
    End If
End Sub

Private Sub txtCustType_LostFocus()
    Dim i As Integer
    
    If txtCustType.Tag = "skiplostfocus" Or txtCustType.Text = "" Then
        txtCustType.Tag = ""
        txtCustType.Visible = False
        Exit Sub
    End If
    With flxCust
        For i = 1 To .Rows - 1
            If UCase$(txtCustType.Text) = UCase$(.TextMatrix(i, 1)) Then
                MsgBox "Customer Type already exist at row " & i & "!", vbExclamation + vbOKOnly, "Customer Type"
                txtCustType.Text = ""
                txtCustType.Visible = False
                Exit Sub
            End If
        Next i
    End With
    flxCust.TextMatrix(ActRow, 1) = txtCustType.Text
    txtCustType.Visible = False
    lstChange.AddItem "Labor Charge"

End Sub

Private Sub txtDFctr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtDFctr.Text = Empty
        flxWDisc.TextMatrix(ActRow, 2) = txtDFctr.Tag
        txtDFctr.Tag = "skiplostfocus"
        flxWDisc.SetFocus
        flxWDisc.Col = 1
        txtDFctr.Visible = False
    End If
End Sub

Private Sub txtDFctr_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        flxWDisc.SetFocus
        If flxWDisc.TextMatrix(ActRow, 1) = "" Then
            flxWDisc.Col = 1
            flxWDisc.Row = ActRow
        ElseIf txtDFctr.Text = "" Then
            flxWDisc.Col = 2
            flxWDisc.Row = ActRow
        ElseIf flxWDisc.TextMatrix(ActRow, 1) <> "" And txtDFctr.Text <> "" And ActRow = flxWDisc.Rows - 1 Then
            flxWDisc.AddItem ""
            flxWDisc.Row = flxWDisc.Rows - 1
            flxWDisc.Col = 1
        Else
            flxWDisc.Row = ActRow + 1
            flxWDisc.Col = 1
        End If
    End If
End Sub

Private Sub txtDFctr_LostFocus()
    
    If txtDFctr.Tag = "skiplostfocus" Or txtDFctr.Text = "" Then
        txtDFctr.Tag = ""
        txtDFctr.Visible = False
        Exit Sub
    End If
    flxWDisc.TextMatrix(ActRow, 2) = txtDFctr.Text
    txtDFctr.Visible = False
    lstChange.AddItem "Wholesale Discount"
End Sub

Private Sub txtDPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtDPrice.Text = Empty
        flxSPec.TextMatrix(ActRow, 3) = txtDPrice.Tag
        txtDPrice.Tag = "skiplostfocus"
        flxSPec.SetFocus
        flxSPec.Col = 3
        txtDPrice.Visible = False
    End If
End Sub

Private Sub txtDPrice_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        flxSPec.SetFocus
        If flxSPec.TextMatrix(ActRow, 1) = "" Then
            flxSPec.Col = 1
            flxSPec.Row = ActRow
        ElseIf txtDPrice.Text = "" Then
            flxSPec.Col = 3
            flxSPec.Row = ActRow
        ElseIf txtDPrice.Text <> "" And flxSPec.TextMatrix(ActRow, 1) <> "" And ActRow = flxSPec.Rows - 1 Then
            flxSPec.AddItem ""
            flxSPec.Row = flxSPec.Rows - 1
            flxSPec.Col = 1
        Else
            flxSPec.Row = ActRow + 1
            flxSPec.Col = 1
        End If
    End If
End Sub

Private Sub txtDPrice_LostFocus()
    If txtDPrice.Tag = "skiplostfocus" Or txtDPrice.Text = "" Then
        txtDPrice.Tag = ""
        txtDPrice.Visible = False
        Exit Sub
    End If
    flxSPec.TextMatrix(ActRow, 3) = Format(txtDPrice.Text, "###0.00")
    txtDPrice.Visible = False
    lstChange.AddItem "SP Disclusion"
End Sub

Private Sub txtJAmt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtJAmt.Text = Empty
        flxLabor.TextMatrix(ActRow, 2) = txtJAmt.Tag
        txtJAmt.Tag = "skiplostfocus"
        flxLabor.SetFocus
        flxLabor.Col = 2
        txtJAmt.Visible = False
    End If
End Sub

Private Sub txtJAmt_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        flxLabor.SetFocus
        If flxLabor.TextMatrix(ActRow, 1) = "" Then
            flxLabor.Col = 1
            flxLabor.Row = ActRow
        ElseIf flxLabor.TextMatrix(ActRow, 2) = "" Then
            flxLabor.Col = 2
            flxLabor.Row = ActRow
        ElseIf flxLabor.TextMatrix(ActRow, 2) <> "" And flxLabor.TextMatrix(ActRow, 1) <> "" And ActRow = flxLabor.Rows - 1 Then
            flxLabor.AddItem ""
            flxLabor.Row = flxLabor.Rows - 1
            flxLabor.Col = 1
        Else
            flxLabor.Row = ActRow + 1
            flxLabor.Col = 1
        End If
    End If
    
End Sub

Private Sub txtJAmt_LostFocus()
    If txtJAmt.Tag = "skiplostfocus" Then
        txtJAmt.Tag = ""
        txtJAmt.Visible = False
        Exit Sub
    End If
    flxLabor.TextMatrix(ActRow, 2) = IIf(Val(txtJAmt.Text) > 0, Format(txtJAmt.Text, "#,##0.00"), "")
    txtJAmt.Visible = False
    lstChange.AddItem "Labor Charge"
End Sub

Private Sub txtJDesc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtJDesc.Text = Empty
        flxLabor.TextMatrix(ActRow, 1) = txtJDesc.Tag
        txtJDesc.Tag = "skiplostfocus"
        flxLabor.SetFocus
        flxLabor.Col = 1
        txtJDesc.Visible = False
    End If
End Sub

Private Sub txtJDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxLabor.SetFocus
        If flxLabor.TextMatrix(ActRow, 1) = "" Then
            flxLabor.Col = 1
            flxLabor.Row = ActRow
        ElseIf flxLabor.TextMatrix(ActRow, 2) = "" Then
            flxLabor.Col = 2
            flxLabor.Row = ActRow
        ElseIf flxLabor.TextMatrix(ActRow, 2) <> "" And flxLabor.TextMatrix(ActRow, 1) <> "" And ActRow = flxLabor.Rows - 1 Then
            flxLabor.AddItem ""
            flxLabor.Row = flxLabor.Rows - 1
            flxLabor.Col = 1
        Else
            flxLabor.Row = ActRow + 1
            flxLabor.Col = 1
        End If
    End If
End Sub

Private Sub txtJDesc_LostFocus()
    Dim i As Integer
    
    If txtJDesc.Tag = "skiplostfocus" Or txtJDesc.Text = "" Then
        txtJDesc.Tag = ""
        txtJDesc.Visible = False
        Exit Sub
    End If
    With flxLabor
        For i = 1 To .Rows - 1
            If UCase$(txtJDesc.Text) = UCase$(.TextMatrix(i, 1)) Then
                MsgBox "Job Description already exist at row " & i & "!", vbExclamation + vbOKOnly, "Service Center Labor Charge"
                txtJDesc.Text = ""
                txtJDesc.Visible = False
                Exit Sub
            End If
        Next i
    End With
    flxLabor.TextMatrix(ActRow, 1) = txtJDesc.Text
    txtJDesc.Visible = False
    lstChange.AddItem "Labor Charge"
End Sub

Private Sub txtPartNo_Change()
    Dim Cost As Currency, CostVat As Currency
    
    'On Error GoTo LocalErr
    With flxPartsD
        If PrtSrch Then
            PrtSrch = False
            Exit Sub
        End If
        .Top = txtPartNo.Height + txtPartNo.Top + Frame4.Top + tabOpt.Top
        .Left = txtPartNo.Left + Frame4.Left + tabOpt.Left
        .Width = txtPartNo.Width
        .Visible = True
        .ZOrder
        If txtPartNo.Text = "" Then .Visible = False: Exit Sub
        rsPartsD.Find "[PARTNUMBER] LIKE '" + txtPartNo.Text + "%'", , adSearchForward, 1
        If Not rsPartsD.EOF Then
            If Not PrtBln Then
                txtPartNo.Tag = Len(txtPartNo.Text)
                PrtBln = True
            End If
            .TopRow = rsPartsD.AbsolutePosition - 1
            .Row = rsPartsD.AbsolutePosition - 1
            .Col = 0: .ColSel = 3
            txtPartNo.Text = rsPartsD!PartNumber
            'Insert Additional retrieval here
            Cost = flxPartsD.TextMatrix(flxPartsD.Row, 3)
            CostVat = Cost + (Cost * (Val(txtVat.Text) / 100))
            flxSPec.TextMatrix(ActRow, 2) = Format(CostVat, "#,##0.00") 'Cost
            flxSPec.TextMatrix(ActRow, 4) = flxPartsD.TextMatrix(flxPartsD.Row, 1) 'SPID
            txtPartNo.SelStart = Val(txtPartNo.Tag)
            txtPartNo.SelLength = Len(txtPartNo.Text)
            PrtBln = False
        Else
            flxSPec.TextMatrix(ActRow, 2) = ""
            flxSPec.TextMatrix(ActRow, 4) = ""
            .Visible = False
        End If
    End With
End Sub

Private Sub txtPartNo_GotFocus()
    With flxPartsD
        .Top = txtPartNo.Height + txtPartNo.Top + Frame4.Top + tabOpt.Top
        .Left = txtPartNo.Left + Frame4.Left + tabOpt.Left
        .Width = txtPartNo.Width
    End With
End Sub

Private Sub txtPartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim Tmp() As String
    
    Select Case KeyCode
        Case vbKeyDown
            txtPartNo.Tag = "skiplostfocus"
            flxPartsD.Visible = True
            flxPartsD.SetFocus
        Case vbKeyEscape
            txtPartNo.Text = Empty
            Tmp = Split(txtPartNo.Tag, "|")
            flxSPec.TextMatrix(ActRow, 1) = Tmp(0)
            flxSPec.TextMatrix(ActRow, 2) = Tmp(1)
            flxSPec.TextMatrix(ActRow, 4) = Tmp(2)
            txtPartNo.Tag = "skiplostfocus"
            flxSPec.SetFocus
            flxSPec.Col = 1
            txtPartNo.Visible = False
            flxPartsD.Visible = False
        Case vbKeyDelete
            PrtSrch = True
            If rsPartsD.State = adStateOpen Then rsPartsD.MoveFirst
        Case vbKeyBack
            PrtSrch = True
            If rsPartsD.State = adStateOpen Then rsPartsD.MoveFirst
    End Select
End Sub

Private Sub txtPartNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxSPec.SetFocus
        If flxSPec.TextMatrix(ActRow, 4) <> "" And flxSPec.TextMatrix(ActRow, 3) <> "" Then
            If ActRow = flxSPec.Rows - 1 Then flxSPec.AddItem ""
            flxSPec.SetFocus
            flxSPec.Row = flxSPec.Rows - 1
            flxSPec.Col = 1
        ElseIf flxSPec.TextMatrix(ActRow, 3) = "" Then   'No Discount Price
            flxSPec.SetFocus
            flxSPec.Col = 3
        End If
        flxPartsD.Visible = False
    End If
End Sub

Private Sub txtPartNo_LostFocus()
    Dim i As Integer
    
    If txtPartNo.Tag = "skiplostfocus" Then
        txtPartNo.Tag = ""
        Exit Sub
    End If
    If flxSPec.TextMatrix(ActRow, 4) = "" And txtPartNo.Text <> "" Then
        MsgBox "Unrecognized Part Number!" & vbCrLf & "Please select a Part Number from list.", vbExclamation + vbOKOnly, "Warning"
        txtPartNo.Visible = False
        flxPartsD.Visible = False
        flxSPec.TextMatrix(ActRow, 2) = ""
        Exit Sub
    End If
    flxSPec.TextMatrix(ActRow, 1) = txtPartNo.Text
    txtPartNo.Visible = False
    flxPartsD.Visible = False
    lstChange.AddItem "SP Disclusion"
End Sub

