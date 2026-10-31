VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form JOBonacct 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "On-Account Customer - Service"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   4560
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7725
      Begin VB.ComboBox cboInd 
         Height          =   315
         Left            =   2340
         TabIndex        =   11
         Top             =   550
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.TextBox txtAddress 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2340
         MaxLength       =   50
         TabIndex        =   5
         Top             =   900
         Width           =   4935
      End
      Begin VB.TextBox txtCustID 
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.OptionButton OptType 
         Caption         =   "COMPANY"
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   3
         Top             =   210
         Width           =   1575
      End
      Begin VB.OptionButton OptType 
         Caption         =   "INDIVIDUAL"
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   2
         Top             =   210
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.ComboBox cboComp 
         Height          =   315
         Left            =   2340
         TabIndex        =   1
         Top             =   550
         Width           =   4935
      End
      Begin VB.Label lblCustName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER NAME"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   720
         TabIndex        =   8
         Top             =   610
         Width           =   1425
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ADDRESS"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   930
         Width           =   2025
      End
      Begin VB.Label lblI_fld 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER TYPE:"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   6
         Top             =   240
         Width           =   1545
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxMC 
      Height          =   3135
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5530
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   375
      Left            =   6480
      TabIndex        =   10
      Top             =   4560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&OK"
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
      Image           =   "JOBonacct.frx":0000
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "JOBonacct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsComp As New ADODB.Recordset
Dim rsIndv As New ADODB.Recordset
Dim CustSrch As Boolean, IndvSrch As Boolean ', CustBln As Boolean

Private Sub DisplaySrvcAcct()
    Dim rsAcctSrvc As New ADODB.Recordset
    Dim Qry As String, i As Integer
    
    'Qry = "SELECT SRVCUSTOMER.CUSTOMER, SRVCUSTOMER.DESIGNATION, SRVCUSTOMER.BRAND, SRVCUSTOMER.MODEL, SRVCUSTOMER.ENGINE, SRVCUSTOMER.CHASSIS, SRVCUSTOMER.PLATENO, SRVCUSTOMER.ACCOUNTNO " & _
          "From SRVCUSTOMER Where (((SRVCUSTOMER.CUSTOMERID) = " & Val(txtCustID.Text) & ")) ORDER BY SRVCUSTOMER.CUSTOMER;"
    'mySQL
    Qry = "SELECT SRVCUSTOMER.CUSTOMER, SRVCUSTOMER.DESIGNATION, SRVCUSTOMER.BRAND, SRVCUSTOMER.MODEL, SRVCUSTOMER.ENGINE, SRVCUSTOMER.CHASSIS, SRVCUSTOMER.PLATENO, SRVCUSTOMER.ACCOUNTNO " & _
          "From SRVCUSTOMER Where SRVCUSTOMER.CUSTOMERID = " & Val(txtCustID.Text) & " ORDER BY SRVCUSTOMER.CUSTOMER;"
    
    
    rsAcctSrvc.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsAcctSrvc.RecordCount = 0 Then
        flxMC.Clear
        flxMC.Rows = 2
        flxMC.FormatString = "   |^CUSTOMER |^DESIGNATION  |^BRAND |^MODEL |^ENGINE |^CHASSIS|^PLATENO "
    Else
        Set flxMC.DataSource = rsAcctSrvc
    End If
    flxMC.ColWidth(0) = 300
    For i = 1 To 7
        flxMC.ColWidth(i) = 2000
    Next i
    flxMC.ColWidth(8) = 0
End Sub


Private Sub cboComp_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If CustSrch Then
        CustSrch = False
        Exit Sub
    End If
    With cboComp
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsComp.MoveFirst
            rsComp.Move i
            txtCustID.Text = rsComp!CUSTOMERID
            txtAddress.Text = rsComp!Address
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                CustSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtCustID.Text = ""
            txtAddress.Text = ""
        End If
    End With
End Sub

Private Sub cboComp_Click()
    Call cboComp_Change
End Sub

Private Sub cboComp_GotFocus()
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboComp.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboComp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then CustSrch = True
End Sub

Private Sub cboComp_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        txtAddress.SetFocus
    End If
End Sub

Private Sub cboComp_LostFocus()
    If Val(txtCustID.Text) > 0 Then Call DisplaySrvcAcct
End Sub

Private Sub cboInd_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If IndvSrch Then
        IndvSrch = False
        Exit Sub
    End If
    With cboInd
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsIndv.MoveFirst
            rsIndv.Move i
            txtCustID.Text = rsIndv!CUSTOMERID
            txtAddress.Text = rsIndv!Address
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                IndvSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtCustID.Text = ""
            txtAddress.Text = ""
        End If
    End With
End Sub

Private Sub cboInd_Click()
    Call cboInd_Change
End Sub

Private Sub cboInd_GotFocus()
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboInd.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboInd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then IndvSrch = True
End Sub

Private Sub cboInd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        txtAddress.SetFocus
    End If
End Sub

Private Sub cboInd_LostFocus()
    If Val(txtCustID.Text) > 0 Then Call DisplaySrvcAcct
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Timer1.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsComp.State = adStateOpen Then rsComp.Close
    If rsIndv.State = adStateOpen Then rsIndv.Close
End Sub

Private Sub OptType_Click(Index As Integer)
    If OptType(0).Value = True Then
        cboComp.Visible = True
        cboInd.Visible = False
        cboComp.Text = ""
    Else
        cboInd.Visible = True
        cboComp.Visible = False
        cboInd.Text = ""
    End If
    txtCustID.Text = ""
    txtAddress.Text = ""
    Call ClearFlx
End Sub

Private Sub Timer1_Timer()
    Dim Qry As String
        
    'Qry = "SELECT DISTINCT CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, SRVCUSTOMER.CTYPE, CUSTOMER.CUSTOMERID " & _
          "FROM CUSTOMER INNER JOIN SRVCUSTOMER ON CUSTOMER.CUSTOMERID = SRVCUSTOMER.CUSTOMERID " & _
          "Where (((SRVCUSTOMER.CTYPE) = 'COMPANY') And ((CUSTOMER.APPLYSERVICE) = True)) ORDER BY CUSTOMER.CUSTOMERNAME;"
    
    'mySQL
    Qry = "SELECT DISTINCT CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, SRVCUSTOMER.CTYPE, CUSTOMER.CUSTOMERID " & _
          "FROM CUSTOMER INNER JOIN SRVCUSTOMER ON CUSTOMER.CUSTOMERID = SRVCUSTOMER.CUSTOMERID " & _
          "Where SRVCUSTOMER.CTYPE = 'COMPANY' And CUSTOMER.APPLYSERVICE = 1 ORDER BY CUSTOMER.CUSTOMERNAME;"
    
    
    With rsComp
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cboComp.AddItem !CustomerName
                .MoveNext
            Loop
        End If
    End With
    
    'Qry = "SELECT DISTINCT CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, SRVCUSTOMER.CTYPE, CUSTOMER.CUSTOMERID " & _
          "FROM CUSTOMER INNER JOIN SRVCUSTOMER ON CUSTOMER.CUSTOMERID = SRVCUSTOMER.CUSTOMERID " & _
          "WHERE (((SRVCUSTOMER.CTYPE)='INDIVIDUAL') AND ((CUSTOMER.APPLYSERVICE)=True)) ORDER BY CUSTOMER.CUSTOMERNAME;"
    'mySQL
    Qry = "SELECT DISTINCT CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, SRVCUSTOMER.CTYPE, CUSTOMER.CUSTOMERID " & _
          "FROM CUSTOMER INNER JOIN SRVCUSTOMER ON CUSTOMER.CUSTOMERID = SRVCUSTOMER.CUSTOMERID " & _
          "WHERE SRVCUSTOMER.CTYPE='INDIVIDUAL' AND CUSTOMER.APPLYSERVICE=1 ORDER BY CUSTOMER.CUSTOMERNAME;"

    
    With rsIndv
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cboInd.AddItem !CustomerName
                .MoveNext
            Loop
        End If
    End With
    Call ClearFlx
    OptType(0).Value = True
    Timer1.Enabled = False
End Sub

Private Sub ClearFlx()
    Dim i As Integer
    
    'set flxmc formatting
    flxMC.Clear
    flxMC.Rows = 2
    flxMC.FormatString = "   |^CUSTOMER |^DESIGNATION  |^BRAND |^MODEL |^ENGINE |^CHASSIS|^PLATENO "
    flxMC.ColWidth(0) = 300
    For i = 1 To 7
        flxMC.ColWidth(i) = 2000
    Next i
    flxMC.ColWidth(8) = 0

End Sub

