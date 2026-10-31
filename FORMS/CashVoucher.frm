VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form CashVoucher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Voucher"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7785
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7785
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkExpense 
      Caption         =   "EXPENSE"
      Height          =   195
      Left            =   1320
      TabIndex        =   36
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CheckBox chkTOF 
      Caption         =   "TOF"
      Height          =   195
      Left            =   480
      TabIndex        =   35
      Top             =   3480
      Width           =   735
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxCV 
      Height          =   735
      Left            =   3240
      TabIndex        =   34
      Top             =   4440
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtPrnCopy 
      Height          =   285
      Left            =   2880
      TabIndex        =   33
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame fraBrowser 
      Height          =   615
      Left            =   240
      TabIndex        =   27
      Top             =   4440
      Visible         =   0   'False
      Width           =   2415
      Begin lvButton.lvButtons_H cmdPrev 
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   165
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
         ImgAlign        =   1
         Image           =   "CashVoucher.frx":0000
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   375
         Left            =   495
         TabIndex        =   29
         Top             =   165
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
         ImgAlign        =   1
         Image           =   "CashVoucher.frx":059A
         cBack           =   -2147483633
      End
      Begin VB.Label lblRec 
         Caption         =   "Voucher x of y"
         Height          =   255
         Left            =   960
         TabIndex        =   30
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox txtPclar1 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   2
      Left            =   480
      TabIndex        =   26
      Top             =   2850
      Width           =   4815
   End
   Begin VB.TextBox txtPclar1 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   480
      TabIndex        =   25
      Top             =   2565
      Width           =   4815
   End
   Begin VB.TextBox txtPclar1 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   24
      Top             =   2280
      Width           =   4815
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1320
      Top             =   5280
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1920
      Top             =   5280
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2520
      Top             =   5280
   End
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3120
      Top             =   5280
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   2  'Center
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   5640
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txtCVno 
      Height          =   285
      Left            =   6240
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtTo 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Top             =   1230
      Width           =   2415
   End
   Begin VB.TextBox txtPosition 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   18
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Edit Voucher"
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
      Image           =   "CashVoucher.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   19
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "E&xit Voucher"
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
      Image           =   "CashVoucher.frx":10CE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Save Voucher"
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
      Image           =   "CashVoucher.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   8400
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Cancel Voucher"
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
      Image           =   "CashVoucher.frx":1C02
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&New Voucher"
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
      Image           =   "CashVoucher.frx":219C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   32
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Print Voucher"
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
      Image           =   "CashVoucher.frx":2736
      cBack           =   -2147483633
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Left            =   420
      Top             =   2235
      Width           =   4935
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Voucher"
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
      Left            =   435
      TabIndex        =   22
      Top             =   690
      Width           =   1800
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Voucher"
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
      Left            =   480
      TabIndex        =   23
      Top             =   720
      Width           =   1800
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6120
      TabIndex        =   17
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblReprint 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5160
      TabIndex        =   16
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CV#"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   15
      Top             =   840
      Width           =   495
   End
   Begin VB.Line Line7 
      Index           =   2
      X1              =   5640
      X2              =   7680
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line7 
      Index           =   1
      X1              =   3120
      X2              =   5160
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line7 
      Index           =   0
      X1              =   240
      X2              =   2640
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label lblRecBy 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   5640
      TabIndex        =   14
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label lblAprvBy 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label lblPrepBy 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Received by:"
      Height          =   255
      Left            =   5640
      TabIndex        =   11
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Approved by:"
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Prepared by:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lblSum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   5640
      TabIndex        =   8
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   5520
      X2              =   7560
      Y1              =   3645
      Y2              =   3645
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A M O U N T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5520
      TabIndex        =   7
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P A R T I C U L A R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   6
      Top             =   1940
      Width           =   3015
   End
   Begin VB.Line Line5 
      X1              =   5520
      X2              =   5520
      Y1              =   1920
      Y2              =   3650
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      Index           =   2
      X1              =   240
      X2              =   7560
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      Index           =   1
      X1              =   240
      X2              =   7560
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   7560
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line3 
      X1              =   6120
      X2              =   7560
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line2 
      X1              =   480
      X2              =   2880
      Y1              =   1845
      Y2              =   1845
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   2880
      Y1              =   1515
      Y2              =   1515
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      Height          =   255
      Left            =   5520
      TabIndex        =   5
      Top             =   1335
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1260
      Width           =   255
   End
End
Attribute VB_Name = "CashVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mcLyr As CRAXDRT.Application
Dim mcReport As CRAXDRT.Report
Dim rsCV As New ADODB.Recordset
Public CVtype As String
Dim BtnIndx As Integer, FrstBtnIndx As Integer, iRec As Integer
Dim AddNew As Boolean

Public Sub chkDtls(myCVType As String)
Dim Qry As String, MyNo As String

    CVtype = myCVType
    If CVtype = "EXPENSE" Then
        '***Access
        'Qry = "SELECT CASHVOUCHER.* From CASHVOUCHER " & _
              "WHERE (((CASHVOUCHER.CVDATE)=#" & Date & "#) AND ((CASHVOUCHER.REMARKS)='" & CVtype & "'));"
        'mySQL Format(, "yyyy-mm-dd")
        
        Qry = "SELECT CASHVOUCHER.* From CASHVOUCHER " & _
              "WHERE CASHVOUCHER.CVDATE='" & Format(lblDate.Caption, "YYYY-MM-DD") & "' AND CASHVOUCHER.REMARKS='" & CVtype & "';"
        MyNo = GetCVNumber(Qry)
        If MyNo = 0 Then
            Call DisplayCurrentPettyCash
        Else
            Call ShowRecord(MyNo)
            'Call BtnCV(False)
        End If
    ElseIf CVtype = "CASH-OUT LYR" Then
        Call BtnCV(True)
        txtTO.Text = "LORENZO Y. RUBINOS SR."
        'txtCVNO.Text = AutoCVNo
        lblDate.Caption = Date
        txtPclar1(1).Text = "CASH-OUT LYR"
        lblPrepBy.Caption = GetSettings("DISBURSER")
        lblAprvBy.Caption = GetSettings("MANAGER")
        AddNew = True
    ElseIf CVtype = "TOF" Then
        Qry = "SELECT CASHVOUCHER.* From CASHVOUCHER " & _
              "WHERE CASHVOUCHER.CVDATE='" & Format(lblDate.Caption, "YYYY-MM-DD") & "' AND CASHVOUCHER.REMARKS='" & CVtype & "';"
        MyNo = GetCVNumber(Qry)
        If MyNo = 0 Then
            Call BtnCV(True)
            'txtCVNO.Text = AutoCVNo
            lblDate.Caption = Date
            txtPclar1(1).Text = "TRANSFER OF FUNDS"
            lblPrepBy.Caption = GetSettings("DISBURSER")
            lblAprvBy.Caption = GetSettings("MANAGER")
            AddNew = True
        Else
            Call ShowRecord(txtCVNO.Text)
            'Call BtnCV(False)
        End If
    End If
End Sub
Private Sub BtnCV(SaveCancel As Boolean)
    Dim i As Integer
    
    If SaveCancel Then
        For i = 0 To 3
            cmdTrans(i).Visible = False
        Next i
        cmdTrans(4).Visible = True
        cmdTrans(4).Left = 0
        cmdTrans(5).Visible = True
        cmdTrans(5).Left = cmdTrans(4).Width
    Else
        cmdTrans(0).Left = 0
        cmdTrans(0).Visible = True
        For i = 1 To 3
            cmdTrans(i).Left = cmdTrans(0).Width * i
            cmdTrans(i).Visible = True
        Next i
        cmdTrans(4).Visible = False
        cmdTrans(5).Visible = False
    End If
End Sub

Function AutoCVNo() As String
    Dim rsInvNo As New ADODB.Recordset
    Dim Qry As String
    Dim MyNo As Long
    
    Qry = "SELECT CASHVOUCHER.VOUCHERNO From CashVoucher ORDER BY CASHVOUCHER.VOUCHERNO DESC;"

    With rsInvNo
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MyNo = IIf(IsNull(.Fields(0)), 1, Val(.Fields(0)) + 1)
        Else
            MyNo = 1
        End If
        .Close
    End With
    AutoCVNo = Format(MyNo, "0000")
End Function

Function GetCVNumber(MyQry As String) As Long
    Dim rsCVNO As New ADODB.Recordset
    
    With rsCVNO
        .Open MyQry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            GetCVNumber = !VOUCHERNO
        Else
            GetCVNumber = 0
        End If
        .Close
    End With
End Function

Function ValidateCV()
    Dim eMsg As String
    
    eMsg = ""
    
    If txtCVNO.Text = "" Then
        eMsg = "Incomplete entry on field CV#"
    Else
        '***Access
        'If AddNew And CtrlNumExist("SELECT VOUCHERNO From CASHVOUCHER WHERE (((VOUCHERNO)=" & Val(txtCVno.Text) & "));") Then
        'mysql
        If AddNew And CtrlNumExist("SELECT VOUCHERNO From CASHVOUCHER WHERE VOUCHERNO='" & txtCVNO.Text & "';") Then
            eMsg = "CV Number " & txtCVNO.Text & " already exist"
        End If
    End If
    If txtTO.Text = "" Then eMsg = eMsg & vbCrLf & "Incomplete entry on field Payee (To)."
    If CVtype <> "CASH-OUT LYR" Then
        If txtPosition.Text = "" Then eMsg = eMsg & vbCrLf & "Incomplete entry on field Payee's Position."
    End If
    If txtPclar1(0).Text = "" And txtPclar1(1).Text = "" And txtPclar1(2).Text = "" Then eMsg = eMsg & vbCrLf & "Incomplete entry on field Particular."
    If txtAmount.Text = "" Then eMsg = eMsg & vbCrLf & "Incomplete entry on field Amount"
    ValidateCV = eMsg
End Function

Public Sub ShowRecord(cvno As String)
    Dim Qry As String
    Dim Prtclr As Variant
    
    '***Access
    'Qry = "SELECT CASHVOUCHER.* From CASHVOUCHER WHERE (((CASHVOUCHER.VOUCHERNO)=" & CVno & "));"
    'mySQL
    Qry = "SELECT CASHVOUCHER.* From CASHVOUCHER WHERE CASHVOUCHER.VOUCHERNO='" & cvno & "';"
    With rsCV
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            lblDate.Caption = Format$(!CVDate, "mm/dd/yyyy")
            txtCVNO.Text = !VOUCHERNO
            txtTO.Text = !PaidTo
            txtPosition.Text = IIf(IsNull(!Position), "", !Position)
            txtAmount.Text = Format$(!Amount, "P##,###,##0.00")
            lblSum.Caption = Format$(!Amount, "P##,###,##0.00")
            lblPrepBy.Caption = !PREPAREDBY
            lblAprvBy.Caption = !APPROVEDBY
            lblRecBy.Caption = !PaidTo
            txtPrnCopy.Text = !PRINTCOPY
            Prtclr = Split(!PARTICULAR, "|")
            For i = LBound(Prtclr) To UBound(Prtclr)
                txtPclar1(i).Text = Prtclr(i)
            Next i
            
            If IIf(IsNull(!Remarks), "", !Remarks) = "TOF" Then
                chkTOF.Value = 1
            ElseIf IIf(IsNull(!Remarks), "", !Remarks) = "EXPENSE" Then
                chkExpense.Value = 1
            Else
                chkExpense.Value = 0
            End If
        End If
    End With
End Sub

Private Sub DisplayCurrentPettyCash()
    Dim rsCVsum As New ADODB.Recordset
    Dim Qry As String, ViewOnly As Boolean
    Dim Prtclr As Variant
    Dim i As Integer, Ans As Integer

    '***Access
    'Qry = "SELECT Sum(DISBURSEMENT.AMOUNT) AS SumOfAMOUNT " & _
          "From DISBURSEMENT WHERE (((DISBURSEMENT.CVDATE)=#" & lblDate.Caption & "#));"
    'mySQL
    Qry = "SELECT Sum(DISBURSEMENT.AMOUNT) AS SumOfAMOUNT " & _
          "From DISBURSEMENT WHERE DISBURSEMENT.CVDATE='" & Format(lblDate.Caption, "yyyy-mm-dd") & "';"
 
    With rsCVsum
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtAmount.Text = IIf(IsNull(.Fields(0)), "P 0.00", Format$(.Fields(0), "P#,###,##0.00"))
            lblSum.Caption = IIf(IsNull(.Fields(0)), "P 0.00", Format$(.Fields(0), "P#,###,##0.00"))
        End If
        .Close
    End With
    txtTO.Text = GetSettings("DISBURSER")
    txtPosition.Text = "Disbursement Officer"
    txtPclar1(1).Text = "Petty Cash"
    lblPrepBy.Caption = txtTO.Text
    lblRecBy.Caption = txtTO.Text
    lblAprvBy.Caption = GetSettings("MANAGER")
'    Call LockData
    txtCVNO.Locked = False
    'txtCVNO.Text = AutoCVNo
    txtAmount.Locked = False
    'Display Save Cancel
    'cmdTrans(0).Visible = False
    'FrstBtnIndx = 1
    'BtnIndx = 3
    'tmrFloat.Enabled = True
    'AddNew = True
End Sub
Private Sub DisplayCashOutLYR()
    txtTO.Text = "Lorenzo Y. Rubinos Sr."
    txtPosition.Text = "General Manager"
    txtPclar1(0).Text = "Cash Out"
    lblDate.Caption = Date
    lblPrepBy.Caption = GetSettings("DISBURSER")
    lblAprvBy.Caption = GetSettings("MANAGER")
    lblRecBy.Caption = "Lorenzo Y. Rubinos Sr."
    cmdProcess.Caption = "&Save"
    chkCVLYR.Visible = True
End Sub


Private Sub UpdatePrintCopy()
    Dim rsCVprn As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT CASHVOUCHER.PRINTCOPY From CASHVOUCHER " & _
          "WHERE (((CASHVOUCHER.VOUCHERNO)=" & txtCVno.Text & "));"
    'mySQL
    Qry = "SELECT CASHVOUCHER.PRINTCOPY From CASHVOUCHER " & _
          "WHERE CASHVOUCHER.VOUCHERNO=" & txtCVNO.Text & ";"
    With rsCVprn
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            .Fields(0) = IIf(IsNull(.Fields(0)), 0, .Fields(0)) + 1
        End If
        .Close
    End With
End Sub

Private Sub PrintCV()
    Dim myMsg As String, PrnVal(0 To 12) As String
    Dim flx As MSFlexGrid
    Dim Copies As String
    
    If Val(txtPrnCopy.Text) > 0 Then
        Ans = MsgBox("You have already printed " & txtPrnCopy.Text & IIf(Val(txtPrnCopy.Text) > 1, " Copies ", " Copy ") & "of CV#-" & txtCVNO.Text & vbCrLf & _
              "Do you want to continue?", vbQuestion + vbYesNo, "Print")
        If Ans = vbNo Then Exit Sub
    End If
    txtPrnCopy.Text = Val(txtPrnCopy.Text) + 1
    Select Case Val(txtPrnCopy.Text)
        Case 1
            Copies = ""
        Case 2
            Copies = "2nd copy"
        Case 3
            Copies = "3rd copy"
        Case Else
            Copies = Val(txtPrnCopy.Text) & "th copy"
    End Select
    PrnVal(0) = "" 'txtCVno.Text
    PrnVal(1) = txtTO.Text
    PrnVal(2) = txtPosition.Text
    PrnVal(3) = lblDate.Caption
    PrnVal(4) = txtPclar1(0).Text
    PrnVal(5) = txtPclar1(1).Text
    PrnVal(6) = txtPclar1(2).Text
    PrnVal(7) = txtAmount.Text
    PrnVal(8) = "" 'Format(Now, "mm/dd/yy hh:mm ampm") & " " & Copies      'DTCpy(myMsg)
    PrnVal(9) = ""  'lblSum.Caption
    PrnVal(10) = "" 'lblPrepBy.Caption
    PrnVal(11) = "" 'lblAprvBy.Caption
    PrnVal(12) = "" 'lblRecBy.Caption
    Call PrintMyForm(PrnVal, "Cash Voucher", flx)
    'LYRcon.Execute "UPDATE CASHVOUCHER SET CASHVOUCHER.PRINTCOPY = [CASHVOUCHER].[PRINTCOPY]+1 " & _
                   "WHERE (((CASHVOUCHER.VOUCHERNO)=" & txtCVno.Text & "));"
    'mySQL
    LYRcon.Execute "UPDATE CASHVOUCHER SET CASHVOUCHER.PRINTCOPY = CASHVOUCHER.PRINTCOPY+1 " & _
                   "WHERE CASHVOUCHER.VOUCHERNO=" & txtCVNO.Text & ";"
End Sub

Private Sub SaveVoucher()
    Dim Amt As Currency
    'Access
    
    'With rsCV
    '    If AddNew Then
    '        .Open "CashVoucher", LYRcon, adOpenKeyset, adLockOptimistic
    '        .AddNew
    '    Else
    '        Qry = "SELECT CASHVOUCHER.* From CASHVOUCHER WHERE (((CASHVOUCHER.VOUCHERNO)=" & txtCVno.Text & "));"
    '        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    '        .MoveFirst
    '    End If
    '    .Fields("VoucherNo") = txtCVno.Text
    '    .Fields("CVDate") = lblDate.Caption
    '    .Fields("PaidTo") = txtTO.Text
    '    .Fields("Position") = txtPosition.Text
    '    .Fields("Amount") = PureNum(txtAmount.Text)
    '    .Fields("Particular") = txtPclar1(0).Text & "|" & txtPclar1(1).Text & "|" & txtPclar1(2).Text
    '    .Fields("PreparedBy") = lblPrepBy.Caption
    '    .Fields("ApprovedBy") = lblAprvBy.Caption
    '    .Fields("PrintCopy") = 0
    '    .Fields("Remarks") = CVtype
    '    .Fields("UpdatedBy") = UserID
    '    .Fields("LastUpdate") = Now
    '    .Update
    '    .Close
    'End With
    Amt = Val(Replace(Replace(txtAmount.Text, "P", ""), ",", ""))
        
    If AddNew Then
        LYRcon.Execute "Insert Into cashvoucher( VOUCHERNO, CVDATE, PAIDTO, `POSITION`, " & _
            "AMOUNT, PARTICULAR, PREPAREDBY, APPROVEDBY, PRINTCOPY, REMARKS, UPDATEDBY, " & _
            "LASTUPDATE)VALUES('" & txtCVNO.Text & "', '" & Format(lblDate.Caption, "yyyy-mm-dd") & "', " & _
            "'" & txtTO.Text & "', '" & txtPosition.Text & "', " & Amt & ", " & _
            "'" & txtPclar1(0).Text & "|" & txtPclar1(1).Text & "|" & txtPclar1(2).Text & "', " & _
            "'" & lblPrepBy.Caption & "', '" & lblAprvBy.Caption & "', 0, '" & CVtype & "', " & _
            "" & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
    Else
        LYRcon.Execute "Update CashVoucher SET VOUCHERNO = '" & txtCVNO.Text & "', " & _
            "CVDATE = '" & Format(lblDate.Caption, "yyyy-mm-dd") & "', PAIDTO = '" & txtTO.Text & "', " & _
            "`POSITION` = '" & txtPosition.Text & "', AMOUNT = " & Amt & ", " & _
            "PARTICULAR = '" & txtPclar1(0).Text & "|" & txtPclar1(1).Text & "|" & txtPclar1(2).Text & "', " & _
            "PREPAREDBY = '" & lblPrepBy.Caption & "', APPROVEDBY = '" & lblAprvBy.Caption & "', " & _
            "PRINTCOPY = 0, REMARKS = '" & CVtype & "', UPDATEDBY = " & UserID & ", " & _
            "LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' WHERE CASHVOUCHER.VOUCHERNO=" & txtCVNO.Text & ";"
    End If
    MsgBox "Cash Voucher successfully saved!", vbInformation + vbOKOnly, "Save"
End Sub

Private Sub chkExpense_Click()
    If chkExpense.Value = 1 Then
        CVtype = "EXPENSE"
        Call chkDtls(CVtype)
    Else
        CVtype = ""
        Call ClearEntries
    End If
End Sub

Private Sub chkTOF_Click()
    If chkTOF.Value = 1 Then
        CVtype = "TOF"
        Call chkDtls(CVtype)
    Else
        CVtype = ""
        Call ClearEntries
    End If
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim Ans As Integer, i As Long, cvno As String
    Dim vMsg As String, Qry As String, MyAcc As String
    
    Select Case Index
        Case 0  'Edit
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=58));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=58;"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
        
            If Val(txtPrnCopy.Text) > 0 Then
                MsgBox "Cannot edit voucher because a printed copy was already made.", vbExclamation + vbOKOnly, "Edit"
                Exit Sub
            End If
            If txtCVNO.Text = "" Then
                MsgBox "No Cash Voucher available for edit", vbExclamation + vbOKOnly, "Edit"
                Exit Sub
            End If
            
            FrstBtnIndx = 0
            BtnIndx = 3
            tmrFloat.Enabled = True
            AddNew = False
            Call LockData(False)
            txtCVNO.Tag = txtCVNO.Text
        Case 1  'New
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=58));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=58;"
            
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            cmdTrans(0).Visible = False
            Call LockData(False)
            FrstBtnIndx = 1
            BtnIndx = 3
            tmrFloat.Enabled = True
            AddNew = True
            Call ClearEntries
            txtCVNO.Text = AutoCVNo
            txtTO.SetFocus
        Case 2  'Print
            Ans = MsgBox("Do you want to Print Voucher?", vbQuestion + vbYesNo, "Print")
            If Ans = vbNo Then Exit Sub
            Call PrintCV
            'LYRcon.Execute "UPDATE CASHVOUCHER SET CASHVOUCHER.PRINTCOPY = [CASHVOUCHER].[PRINTCOPY]+1 " & _
                           "WHERE (((CASHVOUCHER.VOUCHERNO)=" & txtPrnCopy.Text & "));"
            'mySQL
            'LYRcon.Execute "UPDATE CASHVOUCHER SET CASHVOUCHER.PRINTCOPY = CASHVOUCHER.PRINTCOPY+1 " & _
                           "WHERE CASHVOUCHER.VOUCHERNO=" & txtPrnCopy.Text & ";"
       
        Case 3  'Exit
            Unload Me
        Case 4  'Save
            vMsg = ""
            vMsg = ValidateCV
            If vMsg <> "" Then
                MsgBox vMsg, vbExclamation + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            
            Ans = MsgBox("Do you want to Save Voucher?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            'Save Process here
            Call SaveVoucher
            'Ask for Print
            Ans = MsgBox("Do you want to Print Voucher?", vbQuestion + vbYesNo, "Print")
            If Ans = vbYes Then
                Call PrintCV
                FrstBtnIndx = 1
            Else
                FrstBtnIndx = 0
            End If
            Call LockData(True)
            BtnIndx = 5
            tmrFloatRev.Enabled = True
        Case 5  'Cancel
            Ans = MsgBox("Do you want to cancel voucher?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            If AddNew Then
                Call ClearEntries
            Else
                cvno = Val(txtCVNO.Tag)
                Call ShowRecord(cvno)
            End If
            Call LockData(True)
            BtnIndx = 5
            tmrFloatRev.Enabled = True
    End Select
End Sub

Private Sub LockData(bln As Boolean)
    txtCVNO.Locked = bln
    txtTO.Locked = bln
    txtPosition.Locked = bln
    txtAmount.Locked = bln
    txtPclar1(0).Locked = bln
    txtPclar1(1).Locked = bln
    txtPclar1(2).Locked = bln
    chkTOF.Enabled = Not bln
    chkExpense.Enabled = Not bln
End Sub
 
Private Sub ClearEntries()
    txtCVNO.Text = ""
    txtTO.Text = ""
    txtPosition.Text = ""
    txtAmount.Text = ""
    txtPclar1(0).Text = ""
    txtPclar1(1).Text = ""
    txtPclar1(2).Text = ""
    txtPrnCopy.Text = "0"
    chkTOF.Value = 0
    chkExpense.Value = 0
End Sub

Private Sub Form_Load()
    'txtCVno.Text = AutoCVNo
    lblDate.Caption = Date
    lblPrepBy.Caption = GetSettings("DISBURSER")
    lblAprvBy.Caption = GetSettings("MANAGER")
    Call LockData(True)
    Call BtnCV(False)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsCV.State = adStateOpen Then rsCV.Close
    CVtype = ""
End Sub

Private Sub txtAmount_Change()
    lblSum.Caption = txtAmount.Text
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys "{TAB}"
    If KeyCode = vbKeyUp Then SendKeys "+{TAB}"
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789.,P"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


Private Sub txtAmount_LostFocus()
    txtAmount.Text = Format(txtAmount.Text, "#,##0.00")
    lblSum.Caption = Format(txtAmount.Text, "#,##0.00")
End Sub

Private Sub txtCVno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys "{TAB}"
    If KeyCode = vbKeyUp Then SendKeys "+{TAB}"
End Sub

Private Sub txtCVno_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789_"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


Private Sub txtPclar1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys "{TAB}"
    If KeyCode = vbKeyUp Then SendKeys "+{TAB}"
End Sub

Private Sub txtPclar1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


Private Sub txtPosition_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys "{TAB}"
    If KeyCode = vbKeyUp Then SendKeys "+{TAB}"
End Sub

Private Sub txtPosition_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtTo_Change()
    lblRecBy.Caption = txtTO.Text
End Sub

Private Sub txtTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys "{TAB}"
    If KeyCode = vbKeyUp Then SendKeys "+{TAB}"
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub tmrDrop_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 6 Then 'last index + 1
            tmrDrop.Enabled = False
        End If
    End If
End Sub

Private Sub tmrDropRev_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 4 Then
            tmrDropRev.Enabled = False
        End If
    End If
End Sub

Private Sub tmrFloat_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top - 50
    If cmdTrans(BtnIndx).Top <= -375 Then
        cmdTrans(BtnIndx).Visible = False
        BtnIndx = BtnIndx - 1
        If BtnIndx = FrstBtnIndx - 1 Then
            tmrFloat.Enabled = False
            BtnIndx = 4
            cmdTrans(4).Top = -375
            cmdTrans(4).Left = 0
            cmdTrans(4).Visible = True
            cmdTrans(5).Top = -375
            cmdTrans(5).Left = cmdTrans(4).Width
            cmdTrans(5).Visible = True
            tmrDrop.Enabled = True
        End If
    End If
End Sub


Private Sub tmrFloatRev_Timer()
    Dim i As Integer, myLeft As Integer
    
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top - 50
    If cmdTrans(BtnIndx).Top <= -375 Then
        cmdTrans(BtnIndx).Visible = False
        BtnIndx = BtnIndx - 1
        If BtnIndx = 3 Then
            tmrFloatRev.Enabled = False
            BtnIndx = FrstBtnIndx
            For i = FrstBtnIndx To 3
                cmdTrans(i).Visible = True
                'Display Buttons from leftmost part of form
                If i = FrstBtnIndx Then
                    cmdTrans(i).Left = 0
                    myLeft = cmdTrans(i).Width
                Else
                    cmdTrans(i).Left = myLeft
                    myLeft = myLeft + cmdTrans(i).Width
                End If
            Next i
            tmrDropRev.Enabled = True
        End If
    End If
End Sub

