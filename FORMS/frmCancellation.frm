VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCancellation 
   Caption         =   "Cancellation"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12840
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   12840
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   12960
      TabIndex        =   35
      Text            =   "HoldBackColor"
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   12960
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   12615
      Begin VB.ComboBox cboApprovedBy 
         Height          =   315
         Left            =   8280
         TabIndex        =   33
         Top             =   240
         Width           =   4215
      End
      Begin VB.TextBox txtPreparedBy 
         Height          =   285
         Left            =   1800
         TabIndex        =   28
         Tag             =   "XY"
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "APPROVED BY:"
         Height          =   375
         Left            =   6360
         TabIndex        =   30
         Tag             =   "XY"
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "PREPARED BY:"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   12615
      Begin VB.TextBox txtReason 
         Height          =   1845
         Left            =   1800
         MaxLength       =   5000
         TabIndex        =   26
         Tag             =   "XY"
         Top             =   960
         Width           =   5895
      End
      Begin VB.Frame Frame4 
         Height          =   2295
         Left            =   7800
         TabIndex        =   22
         Top             =   480
         Width           =   4695
         Begin VB.ComboBox cboCustomerName 
            Height          =   315
            ItemData        =   "frmCancellation.frx":0000
            Left            =   960
            List            =   "frmCancellation.frx":0002
            TabIndex        =   38
            Top             =   240
            Width           =   3615
         End
         Begin VB.TextBox txtCustDate 
            Height          =   285
            Left            =   960
            TabIndex        =   31
            Tag             =   "XY"
            Top             =   990
            Width           =   3615
         End
         Begin VB.TextBox txtAmount 
            Height          =   285
            Left            =   960
            TabIndex        =   24
            Tag             =   "XY"
            Top             =   600
            Width           =   3615
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            Caption         =   "Note: Please Double Check the Name, Amount and Date if this information was correct before to cancel this Transaction."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   735
            Left            =   240
            TabIndex        =   36
            Top             =   1440
            Width           =   4335
         End
         Begin VB.Label Label12 
            Caption         =   "DATE:"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   1035
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "AMOUNT:"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   645
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "NAME:"
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox txtProgValue 
         Height          =   285
         Left            =   1800
         TabIndex        =   20
         Tag             =   "XY"
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label9 
         Caption         =   "REASON:"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "PROGRAM VALUE:"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Note: If the Transaction Type is Cash Flow, Please encode the Date in this format: YYYY-MM-DD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   12375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   12615
      Begin VB.ComboBox cboStatus 
         Height          =   315
         ItemData        =   "frmCancellation.frx":0004
         Left            =   1920
         List            =   "frmCancellation.frx":000E
         TabIndex        =   37
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtCancellationNo 
         Height          =   285
         Left            =   10560
         TabIndex        =   15
         Tag             =   "XY"
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox cboTransType 
         Height          =   315
         Left            =   1920
         TabIndex        =   13
         Top             =   650
         Width           =   2775
      End
      Begin VB.TextBox txtAdjType 
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Tag             =   "XY"
         Top             =   360
         Width           =   2775
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   360
         Left            =   10560
         TabIndex        =   16
         Tag             =   "XY"
         Top             =   650
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE:"
         Height          =   375
         Left            =   9000
         TabIndex        =   17
         Tag             =   "XY"
         Top             =   675
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "TRANSACTION TYPE:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   675
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "STATUS:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1000
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "ADJUSTMENT TYPE:"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "CANCELLATION NO.:"
         Height          =   375
         Left            =   8640
         TabIndex        =   18
         Tag             =   "XY"
         Top             =   360
         Width           =   1935
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   4500
      TabIndex        =   5
      Top             =   0
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   661
      Caption         =   "&Print"
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
      Image           =   "frmCancellation.frx":0022
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Add New"
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
      Image           =   "frmCancellation.frx":05BC
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   7
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "Su&mmary"
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
      Image           =   "frmCancellation.frx":0B56
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6120
      TabIndex        =   8
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "E&xit"
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
      Image           =   "frmCancellation.frx":10F0
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   7620
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Save"
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
      Image           =   "frmCancellation.frx":168A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   9120
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Cancel"
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
      Image           =   "frmCancellation.frx":1C24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1500
      TabIndex        =   11
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Edit"
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
      Image           =   "frmCancellation.frx":21BE
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   13320
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   13080
      m_MDI_Height    =   6180
   End
End
Attribute VB_Name = "frmCancellation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CID As Long

Dim arrPOSCODE() As String, arrCCODE() As String, arrValidation() As String
Dim mainRs As New ADODB.Recordset
Dim cboRs As New ADODB.Recordset
Dim tmpRs As New ADODB.Recordset
Dim Ans As String
Dim Qry As String
Dim strPartial As String, strTotal As String
Dim formAction As String, myTicketNo As String
Dim myFile As String, POSCODE As String, CCODE As String

Dim i As Long, j As Long
Dim ActRow As Integer, ActCol As Integer
Dim EMPID As Long

Dim StatSrch As Boolean
Dim PrtBln As Boolean, PrtSrch As Boolean, ProperExit As Boolean, myUpload As Boolean

Public Sub ShowRecord(myCancelNo As String)
On Error GoTo er

    Screen.MousePointer = vbHourglass
    
    DoEvents
    Qry = "call sp_mrcretails.sp_cancellation_module_dtls('" & sysBranchCode & "', '" & myCancelNo & "')"
    With tmpRs
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
    
            txtAdjType.Text = IIf(IsNull(!AdjType), "", !AdjType)
            txtCancellationNo.Text = IIf(IsNull(!CANCELNO), "", !CANCELNO)
            cboTransType.Text = IIf(IsNull(!TRANSTYPE), "", !TRANSTYPE)
            txtDate.Text = Format(!DateCreated, "MM/DD/YYYY")
            cboStatus.Text = IIf(IsNull(!STATUS), "", !STATUS)
            txtProgValue.Text = IIf(IsNull(!PROGVALUE), "", !PROGVALUE)
            txtReason.Text = IIf(IsNull(!REASON), "", !REASON) 'IIf(IsNull(!date_final), "__/__/____", Format(!date_final, "MM/DD/YYYY"))
            cboCustomerName.Text = IIf(IsNull(!Name), "", !Name)
            txtAmount.Text = IIf(IsNull(!Amount), "", !Amount)
            txtCustDate.Text = IIf(IsNull(!Date), "", !Date)
            txtPreparedBy = IIf(IsNull(!PreparedBy), "", !PreparedBy)
            cboApprovedBy.Text = IIf(IsNull(!APPROVEDBY), "", !APPROVEDBY)
            
            CID = !CANCELLATIONID
            CCODE = !CCODE
            EMPID = !UPDATEDBY
        End If
        .Close
    End With

    Screen.MousePointer = vbDefault
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->ShowRecord()")
End Sub


Function ValidateEntry() As Boolean
  ValidateEntry = False
    
    If Trim(cboTransType.Text) = "" Then
        MsgBox "TRANSACTION TYPE IS REQUIRED!", vbExclamation, "Validation"
        cboTransType.SetFocus
        Exit Function
    End If
    
    If Trim(cboStatus.Text) = "" Then
        MsgBox "STATUS IS REQUIRED!", vbExclamation, "Validation"
        cboStatus.SetFocus
        Exit Function
    End If
    
    If Trim(txtProgValue.Text) = "" Then
        MsgBox "PROGRAM VALUE IS REQUIRED!", vbExclamation, "Validation"
        txtProgValue.SetFocus
        Exit Function
    End If
    
    If Trim(txtReason.Text) = "" Then
        MsgBox "REASON IS REQUIRED!", vbExclamation, "Validation"
        txtReason.SetFocus
        Exit Function
    End If
    
    If Trim(cboApprovedBy.Text) = "" Then
        MsgBox "APPROVED BY IS REQUIRED!", vbExclamation, "Validation"
        cboApprovedBy.SetFocus
        Exit Function
    End If
    
    If Trim(txtCancellationNo.Text) = "" Then
        MsgBox "CANCELLATION NUMBER IS REQUIRED!", vbExclamation, "Validation"
        txtCancellationNo.SetFocus
        Exit Function
    End If
    
    If Trim(cboCustomerName.Text) = "" Then
        MsgBox "NAME IS REQUIRED!", vbExclamation, "Validation"
        cboCustomerName.SetFocus
        Exit Function
    End If
    
    If Trim(txtAmount.Text) = "" Then
        MsgBox "AMOUNT IS REQUIRED!", vbExclamation, "Validation"
        txtAmount.SetFocus
        Exit Function
    End If
    
    If Trim(txtCustDate.Text) = "" Then
        MsgBox "DATE IS REQUIRED!", vbExclamation, "Validation"
        txtCustDate.SetFocus
        Exit Function
    End If
    
    ValidateEntry = True

End Function

Function ValidateAdj(myAdjustment) As Boolean
    ValidateAdj = False
 
    Qry = "call sp_mrcretails.sp_cancellation_module_validation('" & sysBranchCode & "','" & CCODE & "','" & myAdjustment & "')"
    With tmpRs
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If !CLEVEL = "Error" Then
                cboCustomerName.Text = ""
                txtAmount.Text = ""
                txtCustDate.Text = ""
                MsgBox !CMESSAGE, vbInformation + vbOKOnly, "Validation"
                Exit Function
            Else
                If cboTransType.Text = "CASHIER CASHFLOW" Or cboTransType.Text = "MANAGER CASHFLOW" Then
                    DoEvents
                     txtCustDate.Text = !CDate
                    If tmpRs.State = 1 Then
                        Call modObjTools.addToListObjectWithArrayGivenField(cboCustomerName, tmpRs, arrValidation, "CAMOUNT", "CNAME")
                    End If
                Else
                    cboCustomerName.Text = !CNAME
                    txtAmount.Text = !CAMOUNT
                    txtCustDate.Text = !CDate
                End If
                ValidateAdj = True
            End If
        Else
            Exit Function
        End If
    End With
End Function


Public Function getCancellationNo() As String
Dim cntCancellation As Long

    cntCancellation = Val(GetFieldValue("Select MAX(CANCELLATIONID) as SEQNO From cancellation"))
    cntCancellation = cntCancellation + 1
    
    getCancellationNo = "C-" & sysBranchCode & Year(Now()) & cntCancellation

End Function

Sub clearform()
    'frmHeader
    txtAdjType.Text = "CANCELLATION"
    cboTransType.Text = ""
    cboStatus.Text = ""
    txtCancellationNo.Text = ""
    txtDate.Text = Format(Now, "mm/dd/yyyy")
    'frmBody
    txtProgValue.Text = ""
    txtReason.Text = ""
    cboCustomerName.Text = ""
    txtAmount.Text = ""
    txtCustDate.Text = ""
    'frmFooter
    txtPreparedBy.Text = ""
    cboApprovedBy.Text = ""
    
    CID = 0
    POSCODE = ""
    CCODE = ""
    EMPID = 0
End Sub
Public Sub CboDisplay()
On Error GoTo er
    
    'Get Data from database
    DoEvents
    Qry = "Select UCASE(`VALUE`) as `VALUE`, REMARKS From settings Where (REPLACE(DESCRIPTION,'AR - ','') = '" & sysBranchCode & "' or REPLACE(DESCRIPTION,'AR - ','') = 'ADMIN') and REMARKS IN ('BM','OIC','AOM')"
    If cboRs.State = adStateOpen Then cboRs.Close
    cboRs.Open Qry, LYRcon, adOpenStatic, adLockReadOnly
    If cboRs.State = 1 Then
        Call modObjTools.addToListObjectWithArrayGivenField(cboApprovedBy, cboRs, arrPOSCODE, "REMARKS", "VALUE")
    End If
    If cboRs.State <> 0 Then cboRs.Close
    
    'Get Data from database
    DoEvents
    Qry = "Select UCASE(`VALUE`) as `VALUE`, REMARKS From settings Where DESCRIPTION = 'CANCELLATION MODULE' "
    If cboRs.State = adStateOpen Then cboRs.Close
    cboRs.Open Qry, LYRcon, adOpenStatic, adLockReadOnly
    If cboRs.State = 1 Then
        Call modObjTools.addToListObjectWithArrayGivenField(cboTransType, cboRs, arrCCODE, "VALUE", "REMARKS")
    End If
    If cboRs.State <> 0 Then cboRs.Close
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplay()")
End Sub

Private Sub SetButtonVisibility(DataEntry As Boolean)
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->VIEW SUMM
'3 ->EXIT
'4 ->PRINT
'5 ->SAVE
'6 ->CANCEL

    'hide all button
    For i = 0 To 4
        cmdTrans(i).Visible = False
    Next i

    If DataEntry Then
        cmdTrans(5).Visible = True
        cmdTrans(6).Visible = True
    Else
        For i = 0 To 4
            cmdTrans(i).Visible = True
            cmdTrans(5).Visible = False
            cmdTrans(6).Visible = False
        Next i
    End If
End Sub


Private Sub cboApprovedBy_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cboApprovedBy
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboApprovedBy-Change()")
End Sub

Private Sub cboApprovedBy_Click()
    POSCODE = arrPOSCODE(cboApprovedBy.ListIndex, 0)
End Sub

Private Sub cboApprovedBy_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboApprovedBy_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboApprovedBy, KeyAscii, ValidChars)
End Sub

Private Sub cboApprovedBy_LostFocus()
Dim cbolistIndex As Integer

    cbolistIndex = modObjTools.getListIndex(cboApprovedBy, cboApprovedBy.Text)
    
    If cbolistIndex < 0 Then
        cboApprovedBy.Text = ""
        cboTransType.SetFocus
        MsgBox "The Selected Names were not found on the List!", vbInformation + vbOKOnly, "Validation"
    Else
        cboApprovedBy.ListIndex = cbolistIndex
        Call cboApprovedBy_Click
    End If
End Sub

Private Sub cboCustomerName_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cboCustomerName
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboCustomerName-Change()")
End Sub

Private Sub cboCustomerName_Click()
    txtAmount.Text = arrValidation(cboCustomerName.ListIndex, 0)
End Sub

Private Sub cboCustomerName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboCustomerName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboCustomerName, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        Call cboCustomerName_Click
    End If
End Sub

Private Sub cboStatus_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cboStatus
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboStatus-Change()")
End Sub

Private Sub cboStatus_Click()
Call cboStatus_Change
End Sub

Private Sub cboStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboStatus, KeyAscii, ValidChars)
End Sub

Private Sub cboTransType_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cboTransType
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboTransType-Change()")
End Sub

Private Sub cboTransType_Click()
    CCODE = arrCCODE(cboTransType.ListIndex, 0)
End Sub

Private Sub cboTransType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboTransType_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboTransType, KeyAscii, ValidChars)
End Sub

Private Sub cboTransType_LostFocus()
Dim cbolistIndex As Integer

    cbolistIndex = modObjTools.getListIndex(cboTransType, cboTransType.Text)
    
    If cbolistIndex < 0 Then
        cboTransType.Text = ""
        txtProgValue.SetFocus
        MsgBox "The Selected Transaction were not found on the List!", vbInformation + vbOKOnly, "Validation"
    Else
        cboTransType.ListIndex = cbolistIndex
        Call cboTransType_Click
    End If
End Sub

Private Sub cmdTrans_Click(Index As Integer)
On Error GoTo er
Dim myDefinition As String

    If Index = 0 Then 'add
        Call clearform
        Call LockData(True)
        Call SetButtonVisibility(True)
        formAction = "add"
        txtPreparedBy.Text = GetEmployeeName
        EMPID = UserID
        cboStatus.Text = "PENDING"
        'txtCancellationNo.Text = getCancellationNo
        cboTransType.SetFocus
    ElseIf Index = 1 Then 'edit
        If cboStatus.Text = "FINAL" Then
            MsgBox "This Transaction is already final", vbExclamation + vbOKOnly, "Edit"
            Exit Sub
        ElseIf Val(CID) = 0 Then
            MsgBox "No Cancellation ID found!", vbExclamation + vbOKOnly, "Edit"
            Exit Sub
        End If
        Call SetButtonVisibility(True)
        Call LockData(True)
        formAction = "edit"
    ElseIf Index = 2 Then 'view summary
        Unload Me
        FrmSummary.mySrc = "CANCELLATION MODULE"
        FrmSummary.Show
    ElseIf Index = 3 Then 'print
        Dim empLastName As String
        Dim chkExists As String
        Dim empPosition As String
        Dim ApproverPosition As String

        If txtCancellationNo.Text = "" Then
            MsgBox "No Records Found", vbExclamation + vbOKOnly, "PRINT"
            Exit Sub
        End If

        Qry = "call sp_mrcretails.sp_cancellation_module_dtls('" & sysBranchCode & "', '" & txtCancellationNo.Text & "')"
        frmPrint.QryPrn = Qry

        frmPrint.ParamValue1 = GetFieldValue("Select `VALUE` From settings Where DESCRIPTION = 'ADJUSTMENT REQUEST' and REMARKS = 'AUDIT DIVISION HEAD' ")

        ApproverPosition = GetFieldValue("Select REMARKS From settings Where UCASE(VALUE) = '" & cboApprovedBy.Text & "' and DESCRIPTION like '%AR -%' ")
        If ApproverPosition <> "" Then
            If ApproverPosition = "BM" Then
                frmPrint.ParamValue3 = "Branch Manager"
            ElseIf ApproverPosition = "OIC" Then
                frmPrint.ParamValue3 = "Branch Operataion In-Chief"
            ElseIf ApproverPosition = "AOM" Then
                frmPrint.ParamValue3 = "Area Operation Manager"
            End If
        End If

        empLastName = GetFieldValue("Select LASTNAME From employees where EMPID = " & Val(EMPID) & " ")
        chkExists = GetFieldValue("Select REMARKS From settings Where Value like '%" & empLastName & "%' and REMARKS IN ('BM','OIC','AOM') and DESCRIPTION LIKE '%AR -%' ")
        If chkExists <> "" Then
            If chkExists = "BM" Then
                frmPrint.ParamValue3 = "Branch Manager"
            ElseIf chkExists = "OIC" Then
                frmPrint.ParamValue3 = "Branch Operataion In-Chief"
            ElseIf chkExists = "AOM" Then
                frmPrint.ParamValue3 = "Area Operation Manager"
            End If
        Else
            empPosition = GetFieldValue("Select POSITION From employees Where EMPID = " & Val(EMPID) & " ")
            If empPosition <> "" Then
                frmPrint.ParamValue2 = empPosition
            End If
        End If
        frmPrint.ParamValue4 = SysBranch
        Unload Me
        frmPrint.MySource = "CANCELLATION MODULE"
        frmPrint.Show
    ElseIf Index = 4 Then 'exit
        formAction = "EXIT"
        Call clearform
        formAction = ""
        Unload Me
    ElseIf Index = 5 Then 'save
        If ValidateEntry() = False Then Exit Sub

        Ans = MsgBox("Do you want to save?", vbQuestion + vbYesNo, "Save")

        If Ans = vbNo Then Exit Sub

        If formAction = "add" Then
            Qry = "call sp_mrcretails.sp_cancellation_module_create('" & sysBranchCode & "','" & CCODE & "','" & txtCancellationNo.Text & "'," & _
                  "'" & txtProgValue.Text & "','" & txtReason.Text & "','" & cboStatus.Text & "','" & Val(UserID) & "','" & cboApprovedBy.Text & "','" & cboCustomerName.Text & "','" & txtAmount.Text & "','" & Format(txtCustDate.Text, "YYYY-MM-DD") & "')"
            With mainRs
                If .State = adStateOpen Then .Close
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    CID = !CID
                    DoEvents
                    MsgBox !message, vbInformation + vbOKOnly, "Save"
                    If !level = "Error" Then
                        .Close
                        Exit Sub
                    End If
                End If
                .Close
            End With
        ElseIf formAction = "edit" Then
            Qry = "call sp_mrcretails.sp_cancellation_module_update('" & sysBranchCode & "','" & CCODE & "','" & txtCancellationNo.Text & "'," & _
                  "'" & txtProgValue.Text & "','" & txtReason.Text & "','" & cboStatus.Text & "','" & Val(UserID) & "','" & cboApprovedBy.Text & "','" & cboCustomerName.Text & "','" & txtAmount.Text & "','" & Format(txtCustDate.Text, "YYYY-MM-DD") & "'," & Val(CID) & ")"
            With mainRs
                If .State = adStateOpen Then .Close
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    DoEvents
                    MsgBox !message, vbInformation + vbOKOnly, "Save"
                    If !level = "Error" Then
                        .Close
                        Exit Sub
                    End If
                End If
                .Close
            End With
        End If
        Call SetButtonVisibility(False)
        Call LockData(False)
        formAction = ""
    ElseIf Index = 6 Then 'cancel
        Ans = MsgBox("Do you want to cancel?", vbQuestion + vbYesNo)
        If Ans = vbNo Then Exit Sub
        Call SetButtonVisibility(False)
        Call LockData(False)
        formAction = ""
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmdTrans-Click()")

End Sub

Private Sub Form_Load()
    Call LockData(False)
    txtDate.Text = Format$(Now, "mm/dd/yyyy")
    txtCancellationNo.Text = ""
    txtAdjType.Text = "CANCELLATION"
    Call CboDisplay
    '===================BUTTON POSITION===================
    If cmdTrans(0).Visible = False Then
        'Align Save Button
        cmdTrans(5).Left = cmdTrans(0).Left
        cmdTrans(5).Top = cmdTrans(0).Top
        'Align Cancel Button
        cmdTrans(6).Left = cmdTrans(1).Left
        cmdTrans(6).Top = cmdTrans(1).Top
    End If
End Sub

Sub LockData(DataEntry As Boolean)
    DataEntry = Not DataEntry
    
    'frmHeader
    txtAdjType.Locked = DataEntry
    cboTransType.Locked = DataEntry
    cboStatus.Locked = DataEntry
    txtCancellationNo.Locked = DataEntry
    txtDate.Enabled = Not DataEntry
    'frmBody
    txtProgValue.Locked = DataEntry
    txtReason.Enabled = Not DataEntry
    cboCustomerName.Locked = DataEntry
    txtAmount.Locked = DataEntry
    txtCustDate.Locked = DataEntry
    'frmFooter
    txtPreparedBy.Locked = DataEntry
    cboApprovedBy.Locked = DataEntry
    
    If txtAdjType.Locked Then
        'frmHeader
        txtAdjType.BackColor = txtHoldGrayBackColor.BackColor
        cboTransType.BackColor = txtHoldGrayBackColor.BackColor
        cboStatus.BackColor = txtHoldGrayBackColor.BackColor
        txtCancellationNo.BackColor = txtHoldGrayBackColor.BackColor
        'frmBody
        txtProgValue.BackColor = txtHoldGrayBackColor.BackColor
        txtReason.BackColor = txtHoldGrayBackColor.BackColor
        cboCustomerName.BackColor = txtHoldGrayBackColor.BackColor
        txtAmount.BackColor = txtHoldGrayBackColor.BackColor
        txtCustDate.BackColor = txtHoldGrayBackColor.BackColor
        'frmFooter
        txtPreparedBy.BackColor = txtHoldGrayBackColor.BackColor
        cboApprovedBy.BackColor = txtHoldGrayBackColor.BackColor
    Else
        'frmHeader
        txtAdjType.Locked = True
        cboTransType.BackColor = txtHoldWhiteBackColor.BackColor
        cboStatus.BackColor = txtHoldWhiteBackColor.BackColor
        txtCancellationNo.BackColor = txtHoldWhiteBackColor.BackColor
        txtDate.Enabled = False
        'frmBody
        txtProgValue.BackColor = txtHoldWhiteBackColor.BackColor
        txtReason.BackColor = txtHoldWhiteBackColor.BackColor
        cboCustomerName.Locked = True
        txtAmount.Locked = True
        txtCustDate.Locked = True
        'frmFooter
        txtPreparedBy.Locked = True
        cboApprovedBy.BackColor = txtHoldWhiteBackColor.BackColor
    End If
End Sub

Private Sub txtProgValue_KeyPress(KeyAscii As Integer)
    'KeyAscii = Asc(Chr(KeyAscii))
    'Call ObjKeyTrapping.Validation(txtProgValue, KeyAscii, ValidSQLChars)
    If KeyAscii = 13 Then
        Call txtProgValue_LostFocus
    End If
End Sub


Private Sub txtProgValue_LostFocus()
    If ValidateAdj(txtProgValue.Text) = False Then Exit Sub
    If cboTransType.Text = "CASHIER CASHFLOW" Or cboTransType.Text = "MANAGER CASHFLOW" Then
        cboCustomerName.Locked = False
        cboCustomerName.BackColor = txtHoldWhiteBackColor.BackColor
        cboCustomerName.SetFocus
    Else
        txtReason.SetFocus
    End If
End Sub

Private Sub txtReason_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(Chr(KeyAscii))
    Call ObjKeyTrapping.Validation(txtReason, KeyAscii, ValidSQLChars)
    If KeyAscii = 13 Then
        cboApprovedBy.SetFocus
    End If
End Sub
