VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form SEToptions 
   Caption         =   "Options"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7860
   Icon            =   "SEToptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   7860
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFind 
      Height          =   375
      Left            =   1200
      TabIndex        =   26
      Tag             =   "XYF"
      Top             =   5040
      Width           =   2655
   End
   Begin MSFlexGridLib.MSFlexGrid flxAdj 
      Height          =   975
      Left            =   1080
      TabIndex        =   21
      Top             =   7920
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   1720
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      FormatString    =   "    |<description  |<value           |<remarks |<new desc |<new value|<new remarks"
   End
   Begin VB.TextBox txtVAT 
      Height          =   285
      Left            =   4080
      TabIndex        =   18
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   960
      Top             =   5520
   End
   Begin VB.ListBox lstChange 
      Height          =   2010
      ItemData        =   "SEToptions.frx":058A
      Left            =   240
      List            =   "SEToptions.frx":058C
      TabIndex        =   7
      Top             =   5880
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   4320
      TabIndex        =   1
      Tag             =   "XYF"
      Top             =   4800
      Width           =   3255
      Begin lvButton.lvButtons_H cmdApply 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Tag             =   "XYF"
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         Caption         =   "Apply"
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
         Image           =   "SEToptions.frx":058E
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdClose 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Tag             =   "XYF"
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
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
         Image           =   "SEToptions.frx":0B28
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
      Tab             =   2
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
      TabPicture(0)   =   "SEToptions.frx":10C2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraDisb"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Labor Charge"
      TabPicture(1)   =   "SEToptions.frx":10DE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Customer Type"
      TabPicture(2)   =   "SEToptions.frx":10FA
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Address"
      TabPicture(3)   =   "SEToptions.frx":1116
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame4 
         Height          =   4095
         Left            =   -74880
         TabIndex        =   19
         Tag             =   "xyf"
         Top             =   360
         Width           =   7335
         Begin VB.TextBox txtRegion 
            Height          =   285
            Left            =   5640
            TabIndex        =   24
            Top             =   1560
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtProv 
            Height          =   285
            Left            =   3000
            TabIndex        =   23
            Top             =   960
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.TextBox txtCity 
            Height          =   285
            Left            =   720
            TabIndex        =   22
            Top             =   1080
            Visible         =   0   'False
            Width           =   2175
         End
         Begin MSFlexGridLib.MSFlexGrid flxAdrs 
            Height          =   3735
            Left            =   120
            TabIndex        =   20
            Tag             =   "xyf"
            Top             =   240
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   6588
            _Version        =   393216
            Cols            =   4
            FixedCols       =   0
            RowHeightMin    =   300
            BackColorBkg    =   -2147483633
            FormatString    =   "      |<CITY/TOWN                                |<PROVINCE                                         |<REGION       "
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4095
         Left            =   120
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
   Begin lvButton.lvButtons_H cmdFind 
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Tag             =   "XYF"
      Top             =   5040
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      Caption         =   "FIND:"
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
      Image           =   "SEToptions.frx":1132
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   120
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   7980
      m_MDI_Height    =   6150
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
   Begin VB.Menu mnuAddress 
      Caption         =   "Address"
      Visible         =   0   'False
      Begin VB.Menu mnuNewAdrs 
         Caption         =   "New Address"
      End
      Begin VB.Menu mnuRemAdrs 
         Caption         =   "Remove Address"
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
Dim ActRow As Integer, ActCol As Integer, fRow As Integer, fCol As Integer

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


Private Sub cmdApply_Click()
    Dim rsSettings As New ADODB.Recordset
    Dim i As Integer, Ans As Integer
    Dim MyAcct As String
    'Check individual Objects for changes
    
    'Validate
    If ChangesMade("ADDRESS") Then
        With flxAdj
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 0)) = 2 Then 'AddNew address
                    If .TextMatrix(i, 4) = "" Or .TextMatrix(i, 5) = "" Or .TextMatrix(i, 6) = "" Then
                        MsgBox "Incomplete entry on Address!", vbCritical + vbOKOnly, "Warning"
                        Exit Sub
                    End If
                End If
            Next i
        End With
    End If
    
    
    
    'Account Title
    If ChangesMade("Account Title") Then
        LYRcon.Execute "DELETE FROM DISBURSEACCTDESC;"
        With cboAcctTitle
            For i = 0 To .ListCount - 1
                MyAcct = Replace$(.List(i), "'", "`")
                'LYRcon.Execute "INSERT INTO DISBURSEACCTDESC ( ACCOUNTDESCRIPTION, UPDATEDBY, LASTUPDATE ) " & _
                               "VALUES ('" & MyAcct & "', " & UserID & ", #" & Date & "#);"
                'mySQL
                LYRcon.Execute "INSERT INTO DISBURSEACCTDESC ( ACCOUNTDESCRIPTION, UPDATEDBY, LASTUPDATE ) " & _
                               "VALUES ('" & MyAcct & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
                
            Next i
        End With
    End If
    If ChangesMade("Labor Charge") Then
        LYRcon.Execute "DELETE FROM MCSERVICELABOR;"
        With flxLabor
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) <> "" And .TextMatrix(i, 2) <> "" Then
                    MyAcct = Replace$(.TextMatrix(i, 1), "'", "`")
                    'LYRcon.Execute "INSERT INTO MCSERVICELABOR ( JOBREQUEST, LABORCOST ) " & _
                                   "VALUES ('" & MyAcct & "', " & Val(Replace(.TextMatrix(i, 2), ",", "")) & ");"
                    'mySQL
                    LYRcon.Execute "INSERT INTO MCSERVICELABOR ( JOBREQUEST, LABORCOST ) " & _
                                   "VALUES ('" & MyAcct & "', " & Val(Replace(.TextMatrix(i, 2), ",", "")) & ");"
                    
                End If
            Next i
        End With
    End If
    
    If ChangesMade("Customer Type") Then
        'LYRcon.Execute "DELETE SETTINGS.* From SETTINGS WHERE (((SETTINGS.DESCRIPTION)='Customer Type'));"
        
        'mySQL
        LYRcon.Execute "DELETE From SETTINGS WHERE SETTINGS.DESCRIPTION='Customer Type';"
        
        'With rsSettings
            '.Open "Settings", LYRcon, adOpenKeyset, adLockOptimistic
            For i = 1 To flxCust.Rows - 1
                If flxCust.TextMatrix(i, 1) <> "" And flxCust.TextMatrix(i, 2) <> "" Then
                    '.AddNew
                    '!Description = "Customer Type"
                    'MyAcct = Replace$(flxCust.TextMatrix(i, 1), "'", "`")
                    '!Value = MyAcct
                    '!Remarks = flxCust.TextMatrix(i, 2)
                    '!UPDATEDBY = UserID
                    '!LASTUPDATE = Now
                    '.Update
                    
                    'mySQL
                    MyAcct = Replace$(flxCust.TextMatrix(i, 1), "'", "`")
                    LYRcon.Execute "Insert Into settings(DESCRIPTION, VALUE, REMARKS, UPDATEDBY, LASTUPDATE)" & _
                        "VALUES('Customer Type', '" & MyAcct & "', '" & flxCust.TextMatrix(i, 2) & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
                    
                End If
            Next i
            '.Close
        'End With
    End If
    'flxadj  1-Region/desc, 2-City/value, 3-Prov/remarks, 4-new desc/region, 5-new value/city, 6-new remarks/prov.
    If ChangesMade("ADDRESS") Then
        With flxAdj
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 0)) = 1 Then 'Edit address
                    LYRcon.Execute "Update settings SET VALUE = '" & .TextMatrix(i, 5) & "', DESCRIPTION = 'REGION " & .TextMatrix(i, 4) & "', REMARKS = '" & .TextMatrix(i, 6) & "', UPDATEDBY = " & UserID & ", LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
                    "Where settings.DESCRIPTION = '" & .TextMatrix(i, 1) & "' AND settings.VALUE = '" & .TextMatrix(i, 2) & "' AND settings.REMARKS = '" & .TextMatrix(i, 3) & "';"
                ElseIf Val(.TextMatrix(i, 0)) = 2 Then 'AddNew address
                    LYRcon.Execute "INSERT INTO settings(DESCRIPTION, VALUE, REMARKS, UPDATEDBY, LASTUPDATE) " & _
                    "VALUES('REGION " & .TextMatrix(i, 4) & "', '" & .TextMatrix(i, 5) & "', '" & .TextMatrix(i, 6) & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
                ElseIf Val(.TextMatrix(i, 0)) = 3 Then 'REMOVE address
                    LYRcon.Execute "Delete From settings Where settings.DESCRIPTION = '" & .TextMatrix(i, 1) & "' AND settings.VALUE = '" & .TextMatrix(i, 2) & "' AND settings.REMARKS = '" & .TextMatrix(i, 3) & "';"
                End If
            Next i
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

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    If tabOpt.Tab = 3 Then 'ADDRESS
        TCol = 4
        With flxAdrs
            For i = fRow To .Rows - 1
                For j = fCol To TCol - 1
                    If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                        .Col = 0: .Row = i
                        .ColSel = TCol - 1: .RowSel = i
                        .TopRow = i
                        fRow = IIf(j = TCol - 1, i + 1, i)
                        fCol = IIf(j = TCol - 1, 1, j + 1)
                        ActRow = i
                        Exit Sub
                    End If
                Next j
                fCol = 1
            Next i
        End With
        MsgBox "No match found!", vbInformation + vbOKOnly, "Find"
    ElseIf tabOpt.Tab = 2 Then 'CUSTOMER TYPE
        TCol = 3
        With flxCust
            For i = fRow To .Rows - 1
                For j = fCol To TCol - 1
                    If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                        .Col = 0: .Row = i
                        .ColSel = TCol - 1: .RowSel = i
                        .TopRow = i
                        fRow = IIf(j = TCol - 1, i + 1, i)
                        fCol = IIf(j = TCol - 1, 1, j + 1)
                        ActRow = i
                        Exit Sub
                    End If
                Next j
                fCol = 1
            Next i
        End With
        MsgBox "No match found!", vbInformation + vbOKOnly, "Find"
    ElseIf tabOpt.Tab = 1 Then 'LABOR CHARGE
        TCol = 3
        With flxLabor
            For i = fRow To .Rows - 1
                For j = fCol To TCol - 1
                    If InStr(UCase$(.TextMatrix(i, j)), txtFind.Text) > 0 Then
                        .Col = 0: .Row = i
                        .ColSel = TCol - 1: .RowSel = i
                        .TopRow = i
                        fRow = IIf(j = TCol - 1, i + 1, i)
                        fCol = IIf(j = TCol - 1, 1, j + 1)
                        ActRow = i
                        Exit Sub
                    End If
                Next j
                fCol = 1
            Next i
        End With
        MsgBox "No match found!", vbInformation + vbOKOnly, "Find"
    End If
    fRow = 1
    fCol = 1

End Sub

Private Sub flxAdrs_Click()
    ActRow = flxAdrs.Row
    ActCol = flxAdrs.Col
    
    If ActCol = 0 Or ActRow = 0 Or Not FullAccess Then Exit Sub
    Select Case ActCol
    Case 1
        With txtCity
            If flxAdrs.TextMatrix(ActRow, 1) <> "" Then .Tag = flxAdrs.TextMatrix(ActRow, 1)
            .Top = flxAdrs.CellTop + flxAdrs.Top
            .Left = flxAdrs.CellLeft + flxAdrs.Left
            .Width = flxAdrs.CellWidth
            .Text = flxAdrs.Text
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
    Case 2
        With txtProv
            If flxAdrs.TextMatrix(ActRow, 2) <> "" Then .Tag = flxAdrs.TextMatrix(ActRow, 2)
            .Top = flxAdrs.CellTop + flxAdrs.Top
            .Left = flxAdrs.CellLeft + flxAdrs.Left
            .Width = flxAdrs.CellWidth
            .Text = Replace(flxAdrs.Text, ",", "")
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
    Case 3
        With txtRegion
            If flxAdrs.TextMatrix(ActRow, 3) <> "" Then .Tag = flxAdrs.TextMatrix(ActRow, 3)
            .Top = flxAdrs.CellTop + flxAdrs.Top
            .Left = flxAdrs.CellLeft + flxAdrs.Left
            .Width = flxAdrs.CellWidth
            .Text = Replace(flxAdrs.Text, ",", "")
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
    End Select

End Sub

Private Sub flxAdrs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxAdrs_Click
End Sub

Private Sub flxAdrs_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If FullAccess Then
        If Button = vbRightButton Then
            Call PopupMenu(mnuAddress)
        End If
    End If
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

Private Sub Form_Load()
    Dim Qry As String, MyAcc As String
    
    'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=96));"
    'mySQL
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=96;"
    
    MyAcc = GetFieldValue(Qry)
    FullAccess = IIf(MyAcc = "FULL ACCESS", True, False)
    If Not FullAccess Then cmdApply.Enabled = False
    flxAdj.Rows = 1
        
    
    flxAdrs.Cols = 5
    flxAdrs.ColWidth(4) = 0
    Timer1.Enabled = True
    
    '    |<description  |<value           |<remarks |<new desc |<new value|<new remarks
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

Private Sub mnuNewAdrs_Click()
    If flxAdrs.TextMatrix(flxAdrs.Rows - 1, 1) = "" Or flxAdrs.TextMatrix(flxAdrs.Rows - 1, 2) = "" Or flxAdrs.TextMatrix(flxAdrs.Rows - 1, 3) = "" Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        Exit Sub
    End If
    flxAdj.AddItem "2"
    flxAdrs.AddItem ""
    flxAdrs.TextMatrix(flxAdrs.Rows - 1, 4) = flxAdj.Rows - 1
    flxAdrs.SetFocus
    flxAdrs.Row = flxAdrs.Rows - 1
    flxAdrs.Col = 1
    flxAdrs.TopRow = flxAdrs.Rows - 1
End Sub

Private Sub mnuRemAdrs_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to remove " & flxAdrs.TextMatrix(ActRow, 1) & "?", vbQuestion + vbYesNo, "Remove")
    If Ans = vbYes Then
        With flxAdrs
            If Val(.TextMatrix(ActRow, 4)) > 0 Then
                If flxAdj.TextMatrix(Val(.TextMatrix(ActRow, 4)), 0) = 1 Then
                    flxAdj.TextMatrix(Val(.TextMatrix(ActRow, 4)), 0) = 3
                ElseIf flxAdj.TextMatrix(Val(.TextMatrix(ActRow, 4)), 0) = 2 Then
                    If flxAdj.Rows = 2 Then
                        flxAdj.Rows = 1
                    Else
                        flxAdj.RemoveItem Val(.TextMatrix(ActRow, 4))
                    End If
                End If
            ElseIf .TextMatrix(ActRow, 4) = "" Then
                flxAdj.AddItem 3 & vbTab & "REGION " & .TextMatrix(ActRow, 3) & vbTab & .TextMatrix(ActRow, 1) & vbTab & .TextMatrix(ActRow, 2)
            End If
            If .Rows = 2 Then
                .Rows = 1
                .Rows = 2
            Else
                .RemoveItem ActRow
            End If
            lstChange.AddItem "ADDRESS"
        End With
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
    Dim rsAdrs As New ADODB.Recordset
    
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
        'Qry = "SELECT SETTINGS.* From SETTINGS " & _
              "WHERE (((SETTINGS.DESCRIPTION)='Customer Type')) ORDER BY SETTINGS.DESCRIPTION;"
        'mySQL
        Qry = "SELECT SETTINGS.* From SETTINGS " & _
              "WHERE SETTINGS.DESCRIPTION='Customer Type' ORDER BY SETTINGS.DESCRIPTION;"
        
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
    
    With rsAdrs
        Qry = "SELECT SETTINGS.* From SETTINGS " & _
              "WHERE SETTINGS.DESCRIPTION like 'REGION%' ORDER BY settings.DESCRIPTION, settings.REMARKS, settings.Value;"
        
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxAdrs.Rows = 1
            Do While Not .EOF
                flxAdrs.AddItem vbTab & !Value & vbTab & !Remarks & vbTab & Replace(!Description, "REGION ", "")
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
    
    
    Timer1.Enabled = False
End Sub

Private Sub txtCity_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtCity.Text = Empty
        flxAdrs.TextMatrix(ActRow, 1) = txtCity.Tag
        txtCity.Tag = "skiplostfocus"
        flxAdrs.SetFocus
        flxAdrs.Col = 1
        txtCity.Visible = False
    End If

End Sub

Private Sub txtCity_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxAdrs.SetFocus
        If txtCity.Text = "" Then
            flxAdrs.Col = 1
            flxAdrs.Row = ActRow
        ElseIf flxAdrs.TextMatrix(ActRow, 2) = "" Then
            flxAdrs.Col = 2
            flxAdrs.Row = ActRow
        ElseIf flxAdrs.TextMatrix(ActRow, 2) <> "" And txtCity.Text <> "" And ActRow = flxAdrs.Rows - 1 Then
            flxAdrs.AddItem ""
            flxAdrs.Row = flxAdrs.Rows - 1
            flxAdrs.Col = 1
        Else
            flxAdrs.Row = ActRow + 1
            flxAdrs.Col = 1
        End If
    End If
End Sub

Private Sub txtCity_LostFocus()
    Dim i As Integer
    
    If txtCity.Tag = "skiplostfocus" Or txtCity.Text = "" Then
        txtCity.Tag = ""
        txtCity.Visible = False
        Exit Sub
    End If
    With flxAdrs
        If .TextMatrix(ActRow, 4) = "" And txtCity.Text <> .TextMatrix(ActRow, 1) Then 'This row was not yet edited
            'Add old record to flxadj  1-Region/desc, 2-City/value, 3-Prov/remarks, 4-new desc/region, 5-new value/city, 6-new remarks/prov.
            flxAdj.AddItem 1 & vbTab & "REGION " & .TextMatrix(ActRow, 3) & vbTab & .TextMatrix(ActRow, 1) & vbTab & .TextMatrix(ActRow, 2) & vbTab & .TextMatrix(ActRow, 3) & vbTab & txtCity.Text & vbTab & .TextMatrix(ActRow, 2) 'store new City/Town in flxadj col 5
            'Assign new row of flxadj to column 4 to flxadrs
            .TextMatrix(ActRow, 4) = flxAdj.Rows - 1
            lstChange.AddItem "ADDRESS"
        ElseIf Val(.TextMatrix(ActRow, 4)) > 0 Then 'Row already edited
            'store new City/Town in flxadj col 5 using flxadrs.col(4) as row
            flxAdj.TextMatrix(Val(.TextMatrix(ActRow, 4)), 5) = txtCity.Text
            'flxadrs: 1-city/value, 2-prov./remarks, 3- region/desc. 4-new remarks
            lstChange.AddItem "ADDRESS"
        End If
    End With
    flxAdrs.TextMatrix(ActRow, 1) = txtCity.Text
    txtCity.Visible = False

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


Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
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


Private Sub txtProv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtProv.Text = Empty
        flxAdrs.TextMatrix(ActRow, 2) = txtProv.Tag
        txtProv.Tag = "skiplostfocus"
        flxAdrs.SetFocus
        flxAdrs.Col = 2
        txtProv.Visible = False
    End If
End Sub

Private Sub txtProv_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxAdrs.SetFocus
        If txtProv.Text = "" Then
            flxAdrs.Col = 2
            flxAdrs.Row = ActRow
        ElseIf flxAdrs.TextMatrix(ActRow, 3) = "" Then
            flxAdrs.Col = 3
            flxAdrs.Row = ActRow
        ElseIf flxAdrs.TextMatrix(ActRow, 3) <> "" And txtProv.Text <> "" And ActRow = flxAdrs.Rows - 1 Then
            flxAdrs.AddItem ""
            flxAdrs.Row = flxAdrs.Rows - 1
            flxAdrs.Col = 1
        Else
            flxAdrs.Row = ActRow + 1
            flxAdrs.Col = 1
        End If
    End If
End Sub

Private Sub txtProv_LostFocus()
    Dim i As Integer
    
    If txtProv.Tag = "skiplostfocus" Or txtProv.Text = "" Then
        txtProv.Tag = ""
        txtProv.Visible = False
        Exit Sub
    End If
    With flxAdrs
        If .TextMatrix(ActRow, 4) = "" And txtProv.Text <> .TextMatrix(ActRow, 2) Then 'This row was not yet edited
            'Add old record to flxadj  1-Region/desc, 2-City/value, 3-Prov/remarks, 4-new desc/region, 5-new value/city, 6-new remarks/prov.
            flxAdj.AddItem 1 & vbTab & "REGION " & .TextMatrix(ActRow, 3) & vbTab & .TextMatrix(ActRow, 1) & vbTab & .TextMatrix(ActRow, 2) & vbTab & .TextMatrix(ActRow, 3) & vbTab & .TextMatrix(ActRow, 1) & vbTab & txtProv.Text  'store new City/Town in flxadj col 5
            'Assign new row of flxadj to column 4 to flxadrs
            .TextMatrix(ActRow, 4) = flxAdj.Rows - 1
            lstChange.AddItem "ADDRESS"
        ElseIf Val(.TextMatrix(ActRow, 4)) > 0 Then 'Row already edited
            'store new City/Town in flxadj col 5 using flxadrs.col(4) as row
            flxAdj.TextMatrix(Val(.TextMatrix(ActRow, 4)), 6) = txtProv.Text
            'flxadrs: 1-city/value, 2-prov./remarks, 3- region/desc. 4-new remarks
            lstChange.AddItem "ADDRESS"
        End If
    End With
    flxAdrs.TextMatrix(ActRow, 2) = txtProv.Text
    txtProv.Visible = False

End Sub


Private Sub txtRegion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtRegion.Text = Empty
        flxAdrs.TextMatrix(ActRow, 3) = txtRegion.Tag
        txtRegion.Tag = "skiplostfocus"
        flxAdrs.SetFocus
        flxAdrs.Col = 3
        txtRegion.Visible = False
    End If
End Sub

Private Sub txtRegion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxAdrs.SetFocus
        If txtRegion.Text = "" Then
            flxAdrs.Col = 3
            flxAdrs.Row = ActRow
        ElseIf flxAdrs.TextMatrix(ActRow, 3) = "" Then
            flxAdrs.Col = 3
            flxAdrs.Row = ActRow
        ElseIf flxAdrs.TextMatrix(ActRow, 3) <> "" And txtRegion.Text <> "" And ActRow = flxAdrs.Rows - 1 Then
            flxAdrs.AddItem ""
            flxAdrs.Row = flxAdrs.Rows - 1
            flxAdrs.Col = 1
        Else
            flxAdrs.Row = ActRow + 1
            flxAdrs.Col = 1
        End If
    End If

End Sub

Private Sub txtRegion_LostFocus()
    Dim i As Integer
    
    If txtRegion.Tag = "skiplostfocus" Or txtRegion.Text = "" Then
        txtRegion.Tag = ""
        txtRegion.Visible = False
        Exit Sub
    End If
    With flxAdrs
        If .TextMatrix(ActRow, 4) = "" And txtRegion.Text <> .TextMatrix(ActRow, 3) Then 'This row was not yet edited
            'Add old record to flxadj  1-Region/desc, 2-City/value, 3-Prov/remarks, 4-new desc/region, 5-new value/city, 6-new remarks/prov.
            flxAdj.AddItem 1 & vbTab & "REGION " & .TextMatrix(ActRow, 3) & vbTab & .TextMatrix(ActRow, 1) & vbTab & .TextMatrix(ActRow, 2) & vbTab & txtRegion.Text & vbTab & .TextMatrix(ActRow, 1) & vbTab & .TextMatrix(ActRow, 2)   'store new City/Town in flxadj col 5
            'Assign new row of flxadj to column 4 to flxadrs
            .TextMatrix(ActRow, 4) = flxAdj.Rows - 1
            lstChange.AddItem "ADDRESS"
        ElseIf Val(.TextMatrix(ActRow, 4)) > 0 Then 'Row already edited
            'store new City/Town in flxadj col 5 using flxadrs.col(4) as row
            flxAdj.TextMatrix(Val(.TextMatrix(ActRow, 4)), 4) = txtRegion.Text
            'flxadrs: 1-city/value, 2-prov./remarks, 3- region/desc. 4-new remarks
            lstChange.AddItem "ADDRESS"
        End If
    End With
    flxAdrs.TextMatrix(ActRow, 3) = txtRegion.Text
    txtRegion.Visible = False

End Sub
