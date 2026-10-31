VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmPromoCard 
   Caption         =   "Customer's Promo Card Information"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCCID 
      Height          =   285
      Left            =   7680
      TabIndex        =   33
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "CLUB INFORMATION"
      Height          =   1215
      Left            =   120
      TabIndex        =   28
      Top             =   3600
      Width           =   8775
      Begin VB.ComboBox cboCustClub 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         TabIndex        =   10
         Top             =   240
         Width           =   2535
      End
      Begin VB.ComboBox cboCustClubName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         TabIndex        =   11
         Top             =   720
         Width           =   6495
      End
      Begin lvButton.lvButtons_H cmdViewClubInfo 
         Height          =   375
         Left            =   5400
         TabIndex        =   31
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         Caption         =   "&"
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
         Image           =   "frmPromoCard.frx":0000
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdNewClub 
         Height          =   375
         Left            =   4800
         TabIndex        =   32
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         Caption         =   "&"
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
         cFHover         =   255
         LockHover       =   2
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "frmPromoCard.frx":059A
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label Label7 
         Caption         =   "CUSTOMER'S CLUB:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "CUSTOMER CLUB NAME:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   825
         Width           =   2055
      End
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   9120
      TabIndex        =   25
      Text            =   "HoldBackColor"
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "Add &New"
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
      Image           =   "frmPromoCard.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   12
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "E&xit Form"
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
      Image           =   "frmPromoCard.frx":10CE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   4440
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
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
      Image           =   "frmPromoCard.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   5880
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
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
      Image           =   "frmPromoCard.frx":1C02
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   15
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "&Find"
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
      Image           =   "frmPromoCard.frx":219C
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame1 
      Caption         =   "BASIC INFORMATION"
      Height          =   3015
      Left            =   120
      TabIndex        =   16
      Top             =   480
      Width           =   8775
      Begin VB.TextBox txtClubPoints 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6000
         TabIndex        =   37
         Top             =   2160
         Width           =   2655
      End
      Begin VB.ComboBox cboCardType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6000
         TabIndex        =   35
         Top             =   2520
         Width           =   2655
      End
      Begin VB.TextBox txtCustStatus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         TabIndex        =   9
         Top             =   2520
         Width           =   2535
      End
      Begin VB.ComboBox cboCustName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         TabIndex        =   2
         Top             =   720
         Width           =   6495
      End
      Begin VB.ComboBox cboCustGender 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4080
         TabIndex        =   5
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtCustPoints 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         TabIndex        =   8
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox txtCustEmail 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         TabIndex        =   7
         Top             =   1800
         Width           =   6495
      End
      Begin VB.TextBox txtCustContactNo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6720
         TabIndex        =   6
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtCustAddress 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         TabIndex        =   3
         Top             =   1080
         Width           =   6495
      End
      Begin VB.TextBox txtCustAge 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         TabIndex        =   4
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtCustNo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         TabIndex        =   1
         Top             =   360
         Width           =   2535
      End
      Begin lvButton.lvButtons_H btnConsolPts 
         Height          =   375
         Left            =   6600
         TabIndex        =   38
         Tag             =   "XYF"
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Caption         =   "CHECK STATUS"
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
         Image           =   "frmPromoCard.frx":2736
         cBack           =   -2147483633
      End
      Begin VB.Label lblConnection 
         Alignment       =   2  'Center
         Caption         =   "ONLINE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7200
         TabIndex        =   39
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "CLUB PTS:"
         Height          =   255
         Left            =   4800
         TabIndex        =   36
         Top             =   2250
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "CARD TYPE:"
         Height          =   255
         Left            =   4920
         TabIndex        =   34
         Top             =   2625
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "CUSTOMER'S STATUS:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2625
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "CUSTOMER'S POINTS:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2250
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "CUSTOMER'S EMAIL:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1875
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "CONTACT NO.:"
         Height          =   255
         Left            =   5520
         TabIndex        =   22
         Top             =   1515
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "CUSTOMER'S AGE:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1530
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "GENDER:"
         Height          =   255
         Left            =   3240
         TabIndex        =   20
         Top             =   1530
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "CUSTOMER'S ADDRESS:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1185
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "CUSTOMER'S NAME:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   780
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "CUSTOMER'S NO.:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   375
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmPromoCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Ans As String
Dim Qry As String
Dim arrCUSTNAME() As String
Dim arrCLUBNAME() As String
Dim i As Long, j As Long, CustID As Long
Dim ActRow As Integer, ActCol As Integer
Dim mainRs As New ADODB.Recordset
Dim StatSrch As Boolean
Dim strPartial As String, strTotal As String
Dim formAction As String
Dim PrtBln As Boolean, PrtSrch As Boolean, ProperExit As Boolean

Public Sub ShowRecord(MyNo As String)

'    Call ConnectOnline
'    If modConnection.TestConnection() = False Then Exit Sub

    Qry = "call sp_promotionalonline.sp_promocard_details('" & sysBranchCode & "','" & MyNo & "')"
    With mainRs
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            'Frame 1
            txtCustNo.Text = IIf(IsNull(!CUSTOMERNO), "", !CUSTOMERNO)
            cboCustName.Text = IIf(IsNull(!CustomerName), "", !CustomerName)
            txtCustAddress.Text = IIf(IsNull(!CUSTOMERADDRESS), "", !CUSTOMERADDRESS)
            txtCustAge.Text = IIf(IsNull(!CUSTOMERAGE), "", !CUSTOMERAGE)
            cboCustGender.Text = IIf(IsNull(!CUSTOMERGENDER), "", !CUSTOMERGENDER)
            txtCustContactNo.Text = IIf(IsNull(!CONTACTNO), "", !CONTACTNO)
            txtCustEmail.Text = IIf(IsNull(!CUSTOMEREMAIL), "", !CUSTOMEREMAIL)
            cboCustClub.Text = IIf(IsNull(!CLUBTYPE), "", !CLUBTYPE)
            txtCustPoints.Text = IIf(IsNull(!CUSTOMERPOINTS), "", !CUSTOMERPOINTS)
            txtCustStatus.Text = IIf(IsNull(!STATUS), "", !STATUS)
            CustID = IIf(IsNull(!CUSTOMERID), 0, !CUSTOMERID)
            cboCustClubName.Text = IIf(IsNull(!CLUBNAME), "", !CLUBNAME)
            txtCCID.Text = IIf(IsNull(!CCID), "", !CCID)
            cboCardType.Text = IIf(IsNull(!CARDTYPE), "", !CARDTYPE)
            txtClubPoints.Text = IIf(IsNull(!CLUBPOINTS), 0, !CLUBPOINTS)
        End If
        .Close
    End With
    If cboCardType.Text = "CLUB CARD" Then
        cmdViewClubInfo.Enabled = True
    Else
        cmdViewClubInfo.Enabled = False
        cmdNewClub.Enabled = False
    End If
End Sub


Function ValidateEntry() As Boolean
  ValidateEntry = False
    
    If cboCustName.Text <> "" Then
    
        If Val(CustID) = 0 Then
            MsgBox "UNRECOGNIZED CUSTOMER! PLEASE SELECT ON THE LIST", vbExclamation, "Validation"
            cboCustName.SetFocus
            Exit Function
        End If
        
        If Trim(cboCustName.Text) = "" Then
            MsgBox "CUSTOMERNAME IS REQUIRED!", vbExclamation, "Validation"
            cboCustName.SetFocus
            Exit Function
        End If
        
    End If
    
    If Trim(txtCustAge.Text) = "" Then
        MsgBox "AGE IS REQUIRED!", vbExclamation, "Validation"
        txtCustAge.SetFocus
        Exit Function
    End If
    
    If Trim(cboCustGender.Text) = "" Then
        MsgBox "GENDER IS REQUIRED!", vbExclamation, "Validation"
        cboCustGender.SetFocus
        Exit Function
    End If

    If Trim(cboCardType.Text) = "" Then
        MsgBox "CARD TYPE IS REQUIRED!", vbExclamation, "Validation"
        cboCardType.SetFocus
        Exit Function
    End If
    
    If cboCustClub <> "" Then
    
        If Trim(cboCustClubName.Text) = "" Then
            MsgBox "CLUB NAME IS REQUIRED!", vbExclamation, "Validation"
            cboCustClubName.SetFocus
            Exit Function
        End If
    
        If Trim(txtCCID.Text) = "" Or Val(txtCCID.Text) = 0 Then
            MsgBox "UNRECOGNIZED CLUB NAME! PLEASE SELECT ON THE LIST", vbExclamation, "Validation"
            txtCCID.SetFocus
            Exit Function
        End If
        
    End If
    ValidateEntry = True

End Function

Public Sub CboDisplayClubName()
On Error GoTo er
    'Get Data from database
'    Call ConnectOnline
'    If modConnection.TestConnection() = False Then Exit Sub
    
    DoEvents
    Qry = "call sp_promotionalonline.sp_promocard_getListClubName('" & Trim(cboCustClub.Text) & "')"
    With mainRs
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .State = 1 Then
            Call modObjTools.addToListObjectWithArrayGivenField(cboCustClubName, mainRs, arrCLUBNAME, "CCID", "CLUBNAME")
        End If
        If .State <> 0 Then .Close
    End With
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplay()")
End Sub

Public Sub LoadData()
'    Call ConnectOnline
'    If modConnection.TestConnection() = False Then
'        lblConnection.Caption = "OFFLINE"
'        Exit Sub
'    End If
    
    Qry = "call sp_promotionalonline.sp_promocard_getClubType()"
    With mainRs
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboCustClub.Clear
            Do While Not .EOF
                cboCustClub.AddItem !Value
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    cboCustGender.AddItem ""
    cboCustGender.AddItem "MALE"
    cboCustGender.AddItem "FEMALE"
    
    cboCardType.AddItem ""
    cboCardType.AddItem "MEMBERSHIP CARD"
    cboCardType.AddItem "CLUB CARD"
    
    lblConnection.Caption = "ONLINE"
End Sub
Private Sub SetButtonVisibility(DataEntry As Boolean)
'cmdTrans Legend
'0 ->ADD
'1 ->FIND
'2 ->EXIT
'3 ->SAVE
'4 ->CANCEL

    'hide all button
    For i = 0 To 4
        cmdTrans(i).Visible = False
    Next i

    If DataEntry Then
        cmdTrans(3).Visible = True
        cmdTrans(4).Visible = True
    Else
        For i = 0 To 4
            cmdTrans(i).Visible = True
            cmdTrans(3).Visible = False
            cmdTrans(4).Visible = False
        Next i
    End If
End Sub

Sub Clearform()

    'Frame 1
    txtCustNo.Text = ""
    cboCustName.Text = ""
    txtCustAddress.Text = ""
    txtCustAge.Text = ""
    cboCustGender.Text = ""
    txtCustContactNo.Text = ""
    txtCustEmail.Text = ""
    cboCustClub.Text = ""
    txtCustPoints.Text = ""
    txtCustStatus.Text = ""
    CustID = 0
    cboCustClubName.Text = ""
    txtCCID.Text = ""
    cboCardType.Text = ""
    txtClubPoints.Text = ""
End Sub

Public Sub CboDisplay()
On Error GoTo er

    If formAction = "add" Then
        'Get Data from database
        DoEvents
        Qry = "call sp_promotionalonline.sp_promocard_getCustomer('" & sysBranchCode & "')"
        With mainRs
            If mainRs.State = adStateOpen Then mainRs.Close
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .State = 1 Then
            If .RecordCount > 0 Then
                Call modObjTools.fillArray(modMotoR.AllArrCust, mainRs)
            Else
                MsgBox "No Record Found", vbInformation + vbOKOnly, "Validation"
                Exit Sub
            End If
        End If
        If .State <> 0 Then .Close
        
        'Load Data to combo box
        Call SetCboListData
        End With
    End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplay()")
End Sub

Sub SetCboListData()
On Error GoTo er

    DoEvents
    Call modObjTools.addToListObjectUsingArrayWithCols(cboCustName, arrCUSTNAME, modMotoR.AllArrCust, 2, 1)

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->SetCboListData()")
End Sub

Sub LockData(DataEntry As Boolean)
    DataEntry = Not DataEntry

    txtCustNo.Locked = DataEntry
    cboCustName.Locked = DataEntry
    txtCustAddress.Locked = DataEntry
    txtCustAge.Locked = DataEntry
    cboCustGender.Locked = DataEntry
    txtCustContactNo.Locked = DataEntry
    txtCustEmail.Locked = DataEntry
    cboCustClub.Locked = DataEntry
    txtCustPoints.Locked = DataEntry
    txtCustStatus.Locked = DataEntry
    cboCustClubName.Locked = DataEntry
    cboCardType.Locked = DataEntry
    txtClubPoints.Locked = DataEntry
    
    If txtCustAddress.Locked Then
        txtCustNo.BackColor = txtHoldGrayBackColor.BackColor
        cboCustName.BackColor = txtHoldGrayBackColor.BackColor
        txtCustAddress.BackColor = txtHoldGrayBackColor.BackColor
        txtCustAge.BackColor = txtHoldGrayBackColor.BackColor
        cboCustGender.BackColor = txtHoldGrayBackColor.BackColor
        txtCustContactNo.BackColor = txtHoldGrayBackColor.BackColor
        txtCustEmail.BackColor = txtHoldGrayBackColor.BackColor
        cboCustClub.BackColor = txtHoldGrayBackColor.BackColor
        txtCustPoints.BackColor = txtHoldGrayBackColor.BackColor
        txtCustStatus.BackColor = txtHoldGrayBackColor.BackColor
        cboCustClubName.BackColor = txtHoldGrayBackColor.BackColor
        cboCardType.BackColor = txtHoldGrayBackColor.BackColor
        txtClubPoints.BackColor = txtHoldGrayBackColor.BackColor
    Else
        txtCustNo.Locked = True
        cboCustName.BackColor = txtHoldWhiteBackColor.BackColor
        txtCustAddress.Locked = True
        txtCustAge.BackColor = txtHoldWhiteBackColor.BackColor
        cboCustGender.BackColor = txtHoldWhiteBackColor.BackColor
        txtCustContactNo.BackColor = txtHoldWhiteBackColor.BackColor
        txtCustEmail.BackColor = txtHoldWhiteBackColor.BackColor
        'cboCustClub.BackColor = txtHoldWhiteBackColor.BackColor
        'cboCustClubName.BackColor = txtHoldWhiteBackColor.BackColor
        cboCardType.BackColor = txtHoldWhiteBackColor.BackColor
        txtCustPoints.Locked = True
        txtCustStatus.Locked = True
        txtClubPoints.Locked = True
    End If
    
End Sub

Private Sub btnConsolPts_Click()
    Call ConnectOnline
    If modConnection.TestConnection() = False Then Exit Sub
    
    If txtCustNo.Text <> "" Then
        Qry = "call sp_promotionalonline.sp_promocard_CheckStatus('" & sysBranchCode & "'," & Val(CustID) & ",'" & txtCustNo.Text & "')"
        
        With mainRs
            If .State = adStateOpen Then .Close
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            
            If .RecordCount > 0 Then
                .MoveFirst
                
                If !level = "Granted" Then
                    MsgBox "" & !message & "", vbInformation + vbOKOnly, "VALIDATION"
                    txtCustStatus.Text = !STATUS
                Else
                   MsgBox "" & !message & "", vbInformation + vbOKOnly, "VALIDATION"
                    .Close
                    Exit Sub
                End If
            End If
            .Close
        End With
    Else
        MsgBox "No Customer No. is Found!", vbInformation + vbOKOnly, "VALIDATION"
        Exit Sub
    End If
End Sub

Private Sub cboCardType_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cboCardType
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboCardType-Change()")
End Sub

Private Sub cboCardType_Click()
If cboCardType.Text = "MEMBERSHIP CARD" Or cboCardType.Text = "" Then
    cboCustClub.Locked = True
    cboCustClubName.Locked = True
    cboCustClub.BackColor = txtHoldGrayBackColor.BackColor
    cboCustClubName.BackColor = txtHoldGrayBackColor.BackColor
    cboCustClub.Text = ""
    cboCustClubName.Text = ""
ElseIf cboCardType.Text = "CLUB CARD" Then
    cboCustClub.Locked = False
    cboCustClubName.Locked = False
    cboCustClub.BackColor = txtHoldWhiteBackColor.BackColor
    cboCustClubName.BackColor = txtHoldWhiteBackColor.BackColor
End If
Call cboCardType_Change
cboCustClub.SetFocus
End Sub

Private Sub cboCardType_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboCardType_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboCardType, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        Call cboCardType_Click
    End If
End Sub

Private Sub cboCustClub_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cboCustClub
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboCustClub-Change()")
End Sub

Private Sub cboCustClub_Click()
If cboCustClub.Text <> "" Then
    cmdNewClub.Enabled = True
    Call CboDisplayClubName
Else
    cboCustClubName.Clear
    txtCCID.Text = ""
    cmdNewClub.Enabled = False
End If
End Sub

Private Sub cboCustClub_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboCustClub_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboCustClub, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        Call cboCustClub_Click
    End If
End Sub


Private Sub cboCustClubName_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cboCustClubName
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboCustClubName-Change()")
End Sub

Private Sub cboCustClubName_Click()
If cboCustClubName.Text <> "" Then
    If cboCustClubName.Locked = False Then
        If cboCustClubName.ListIndex = -1 Then
            MsgBox "Unrecognized Club Name! Please Select on the List", vbInformation + vbOKOnly, "VALIDATION"
            txtCCID.Text = ""
            cboCustClubName.Text = ""
            cboCustClubName.SetFocus
        Else
            txtCCID.Text = arrCLUBNAME(cboCustClubName.ListIndex, 0)
            cmdViewClubInfo.Enabled = True
        End If
    End If
Else
    txtCCID.Text = ""
    cmdViewClubInfo.Enabled = False
End If
End Sub

Private Sub cboCustClubName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub



Private Sub cboCustClubName_KeyPress(KeyAscii As Integer)
Dim cbolistIndex As Integer

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(cmbBranch, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        cbolistIndex = modObjTools.getListIndex(cboCustClubName, cboCustClubName.Text)
        
        If cbolistIndex < 0 Then
            Call cboCustClubName_Click
        Else
            cboCustClubName.ListIndex = cbolistIndex
            Call cboCustClubName_Click
        End If
    
        
    End If
End Sub

Private Sub cboCustClubName_LostFocus()
Call cboCustClubName_Click
End Sub

Private Sub cboCustGender_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cboCustGender
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboCustGender-Change()")
End Sub

Private Sub cboCustGender_Click()
Call cboCustGender_Change
txtCustContactNo.SetFocus
End Sub

Private Sub cboCustGender_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboCustGender_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboCustGender, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        Call cboCustGender_Click
    End If
End Sub

Private Sub cboCustName_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cboCustName
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboCustName-Change()")
End Sub

Private Sub cboCustName_Click()
    DoEvents
    If cboCustName.ListIndex = -1 Then
        CustID = 0
        txtCustAddress.Text = ""
    Else
        CustID = arrCUSTNAME(cboCustName.ListIndex, 0)
        txtCustAddress.Text = arrCUSTNAME(cboCustName.ListIndex, 2)
    End If
    
    If CustID > 0 Then
        txtCustAge.SetFocus
    Else
        MsgBox "No Customer Found!", vbInformation + vbOKOnly, "Validation"
        Exit Sub
    End If
End Sub

Private Sub cboCustName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboCustName_KeyPress(KeyAscii As Integer)
Dim cbolistIndex As Integer
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(cmbBranch, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        

        cbolistIndex = modObjTools.getListIndex(cboCustName, cboCustName.Text)

        If cbolistIndex < 0 Then
            cboCustName.SetFocus
        Else
            cboCustName.ListIndex = cbolistIndex
        End If
        Call cboCustName_Click
        
    End If
End Sub

Private Sub cmdNewClub_Click()
frmClubInfo.cmdTrans(1).Visible = False
frmClubInfo.cmdTrans(0).Visible = True
frmClubInfo.myClubType = cboCustClub.Text
frmClubInfo.Show vbModal
End Sub

Private Sub cmdTrans_Click(Index As Integer)
Dim localRs As New ADODB.Recordset
On Error GoTo er

    If Index = 0 Then 'add
'        If lblConnection.Caption = "ONLINE" Then
            formAction = "add"
            Call Clearform
            Call CboDisplay
            Call LockData(True)
            Call SetButtonVisibility(True)
            cmdNewClub.Enabled = False
            cmdViewClubInfo.Enabled = False
'        Else
'            MsgBox "THE CONNECTION WAS OFFLINE!", vbInformation + vbOKOnly, "VALIDATION"
'            Exit Sub
'        End If
    ElseIf Index = 1 Then 'find
'        If lblConnection.Caption = "ONLINE" Then
            Set frmSearch.FrmSource = frmPromoCard
            frmSearch.MySource = "PromoCard"
            Qry = "call sp_promotionalonline.sp_promocard_find('" & sysBranchCode & "')"
            frmSearch.QrySearch = Qry
            frmSearch.lblCap.Caption = "Search for Promo Card Customer"
            frmSearch.Show vbModal
'        Else
'            MsgBox "THE CONNECTION WAS OFFLINE!", vbInformation + vbOKOnly, "VALIDATION"
'            Exit Sub
'        End If
    ElseIf Index = 2 Then
        formAction = "EXIT"
        ProperExit = True
        'Call clearform
        formAction = ""
        Unload Me
    ElseIf Index = 3 Then 'save
        If ValidateEntry() = False Then Exit Sub

        Ans = MsgBox("Do you want to save?", vbQuestion + vbYesNo, "Save")
        
        If Ans = vbNo Then Exit Sub
        
        If formAction = "add" Then
'            Call ConnectOnline
'            If modConnection.TestConnection() = False Then Exit Sub
'            DoEvents
'            Qry = "call sp_promotionalonline.sp_promocard_create('" & sysBranchCode & "'," & Val(CustID) & ",'" & Trim(cboCustName.Text) & "','" & Trim(txtCustAddress.Text) & "'," & Val(txtCustAge.Text) & ",'" & Trim(cboCustGender.Text) & "','" & Trim(txtCustContactNo.Text) & "','" & Trim(txtCustEmail.Text) & "'," & Val(txtCCID.Text) & "," & Val(UserID) & ",'" & Trim(cboCardType.Text) & "')"
'
'            With mainRs
'                If .State = adStateOpen Then .Close
'                .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
'                If .RecordCount > 0 Then
'                    If !level = "Granted" Then
'                        txtCustNo.Text = IIf(IsNull(!CUSTNO), "", !CUSTNO)
'                        txtCustStatus.Text = IIf(IsNull(!STATUS), "", !STATUS)
'                        MsgBox "" & !message & "", vbInformation + vbOKOnly, "" & !level & ""
'                    Else
'                        MsgBox "" & !message & "", vbInformation + vbOKOnly, "" & !level & ""
'                        .Close
'                        Exit Sub
'                    End If
'                End If
'                .Close
'            End With

            'save sa local
            Qry = "call sp_promotionalonline.sp_promocard_create('" & sysBranchCode & "'," & Val(CustID) & ",'" & Trim(cboCustName.Text) & "','" & Trim(txtCustAddress.Text) & "'," & Val(txtCustAge.Text) & ",'" & Trim(cboCustGender.Text) & "','" & Trim(txtCustContactNo.Text) & "','" & Trim(txtCustEmail.Text) & "'," & Val(txtCCID.Text) & "," & Val(UserID) & ",'" & Trim(cboCardType.Text) & "')"
            With localRs
            If .State = adStateOpen Then .Close
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    If !level = "Granted" Then
                        txtCustNo.Text = IIf(IsNull(!CUSTNO), "", !CUSTNO)
                        txtCustStatus.Text = IIf(IsNull(!STATUS), "", !STATUS)
                        MsgBox "" & !message & "", vbInformation + vbOKOnly, "" & !level & ""
                    Else
                        MsgBox "" & !message & "", vbInformation + vbOKOnly, "" & !level & ""
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
        cmdViewClubInfo.Enabled = True
        cmdNewClub.Enabled = False
    
    ElseIf Index = 4 Then 'cancel
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

Private Sub cmdViewClubInfo_Click()
frmClubInfo.cmdTrans(1).Visible = True
frmClubInfo.cmdTrans(0).Visible = False
frmClubInfo.myClubType = cboCustClub.Text
frmClubInfo.ShowRecord (Val(txtCCID.Text))
frmClubInfo.Show vbModal
End Sub

Private Sub Form_Load()
    Call LockData(False)
    Call LoadData
       
    '===================BUTTON POSITION===================
    If cmdTrans(0).Visible = False Then
        'Align Save Button
        cmdTrans(3).Left = cmdTrans(0).Left
        cmdTrans(3).Top = cmdTrans(0).Top
        'Align Cancel Button
        cmdTrans(4).Left = cmdTrans(1).Left
        cmdTrans(4).Top = cmdTrans(1).Top
    End If
End Sub

Private Sub txtCustAge_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtCustAge, KeyAscii, NumbersOnly)
    If KeyAscii = 13 Then
        cboCustGender.SetFocus
    End If
End Sub

Private Sub txtCustContactNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtCustContactNo, KeyAscii, NumbersOnly)
    If KeyAscii = 13 Then
        txtCustEmail.SetFocus
    End If
End Sub


Private Sub txtCustEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboCardType.SetFocus
    End If
End Sub
