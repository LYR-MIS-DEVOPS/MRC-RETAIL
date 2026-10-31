VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form SPAddnew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Spare Part"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6270
   Icon            =   "SPAddNew.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAppmodel 
      Height          =   285
      Left            =   2040
      TabIndex        =   10
      Top             =   4200
      Width           =   3735
   End
   Begin VB.Frame Frame3 
      Height          =   135
      Left            =   360
      TabIndex        =   38
      Top             =   3720
      Width           =   5415
   End
   Begin VB.TextBox txtSPFid 
      Height          =   285
      Left            =   4680
      TabIndex        =   37
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtVat 
      Height          =   285
      Left            =   4680
      TabIndex        =   36
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtMarkUp 
      Height          =   285
      Left            =   4680
      TabIndex        =   35
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cboFactor 
      Height          =   315
      Left            =   2040
      TabIndex        =   33
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox txtPartNo 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   625
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   360
      TabIndex        =   31
      Top             =   4560
      Width           =   5415
      Begin lvButton.lvButtons_H cmdOk 
         Height          =   375
         Left            =   1200
         TabIndex        =   11
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
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
         Image           =   "SPAddNew.frx":08CA
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   375
         Left            =   2760
         TabIndex        =   12
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
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
         Image           =   "SPAddNew.frx":0E64
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00C00000&
      Height          =   2055
      Left            =   1080
      TabIndex        =   1
      Top             =   5520
      Width           =   4215
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "400 - Accessories"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   6
         Left            =   2520
         TabIndex        =   30
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "300 - Apparels"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   840
         TabIndex        =   29
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "200 - Frame Parts"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   28
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "100 - Engine Parts"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   27
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Line Line7 
         X1              =   720
         X2              =   480
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line6 
         X1              =   480
         X2              =   480
         Y1              =   480
         Y2              =   1320
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Brand or Model"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   26
         Top             =   960
         Width           =   2655
      End
      Begin VB.Line Line5 
         X1              =   720
         X2              =   600
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line4 
         X1              =   600
         X2              =   600
         Y1              =   1080
         Y2              =   600
      End
      Begin VB.Line Line3 
         X1              =   1320
         X2              =   1320
         Y1              =   480
         Y2              =   600
      End
      Begin VB.Line Line2 
         X1              =   1320
         X2              =   600
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Other Description + Supplier Code"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   25
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Classification Code [CCC] + Description [DD]"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   24
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         X1              =   2160
         X2              =   2160
         Y1              =   480
         Y2              =   720
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CCCDD - BBB - XXXSSS.."
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
         Left            =   360
         TabIndex        =   23
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.TextBox txtMaxDisc 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      TabIndex        =   9
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox txtPrice 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Top             =   2760
      Width           =   2535
   End
   Begin VB.TextBox txtCost 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox txtType 
      Alignment       =   2  'Center
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1690
      Width           =   2535
   End
   Begin VB.TextBox txtSupCode 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   1335
      Width           =   1095
   End
   Begin VB.TextBox txtSupplier 
      Alignment       =   2  'Center
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1335
      Width           =   2535
   End
   Begin VB.TextBox txtPartNo3 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   4440
      TabIndex        =   13
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox txtPartNo2 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   3120
      MaxLength       =   10
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox txtDesc 
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2040
      MaxLength       =   250
      TabIndex        =   4
      Top             =   980
      Width           =   3735
   End
   Begin VB.TextBox txtPartNo1 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "APPLICABLE MODEL "
      Height          =   255
      Left            =   0
      TabIndex        =   39
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label lblGDesc 
      Alignment       =   1  'Right Justify
      Caption         =   "PRICE FACTOR"
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   34
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "PART NUMBER"
      Height          =   255
      Left            =   480
      TabIndex        =   32
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblGDesc 
      Alignment       =   1  'Right Justify
      Caption         =   "MAX. DISCOUNT"
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   22
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label lblGDesc 
      Alignment       =   1  'Right Justify
      Caption         =   "RETAIL PRICE"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   21
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label lblGDesc 
      Alignment       =   1  'Right Justify
      Caption         =   "COST"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   20
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lblGDesc 
      Alignment       =   1  'Right Justify
      Caption         =   "TYPE"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   19
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblGDesc 
      Alignment       =   1  'Right Justify
      Caption         =   "SUPPLIER"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   17
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblPartNO 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   16
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label lblGDesc 
      Alignment       =   1  'Right Justify
      Caption         =   "DESCRIPTION"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   15
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblGPnum 
      Alignment       =   1  'Right Justify
      Caption         =   "PART NUMBER"
      Height          =   285
      Left            =   480
      TabIndex        =   14
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "SPAddnew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsFactor As New ADODB.Recordset
Dim PrtSrch As Boolean, PrtBln As Boolean
Public sRow As Integer
Public MySource As String
Dim FctrSrch As Boolean, UseflxNewSP As Boolean
Dim flxNewRow As Integer

Private Sub cboFactor_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    Dim NewPrice As Currency, CostVat As Currency, MyCost As Currency

    If FctrSrch Then
        FctrSrch = False
        Exit Sub
    End If
    With cboFactor
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsFactor.MoveFirst
            rsFactor.Move i
            txtSPFid.Text = rsFactor!SPFid
            txtMarkUp.Text = rsFactor!Markup
            txtVat.Text = rsFactor!Vat
            txtMaxDisc.Text = rsFactor!MaxDiscount
            If Val(txtCost.Text) > 0 Then
                MyCost = PureNum(txtCost.Text)
                CostVat = MyCost + (MyCost * (Val(txtVat.Text) / 100))
                NewPrice = CostVat + (CostVat * (Val(txtMarkUp.Text) / 100))
                txtPrice.Text = Format$(Int(NewPrice + 0.6) + 1, "####0.00")
                txtPrice.Tag = NewPrice
            End If
            txtPrice.Locked = True
            txtMaxDisc.Locked = True
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                FctrSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtSPFid.Text = ""
            txtMarkUp.Text = ""
            txtVat.Text = ""
            txtMaxDisc.Text = ""
            If txtPrice.Text = txtPrice.Tag Then txtPrice.Text = ""
            txtPrice.Locked = False
            txtMaxDisc.Locked = False
        End If
    End With

End Sub

Private Sub cboFactor_Click()
    Call cboFactor_Change
End Sub

Private Sub cboFactor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then FctrSrch = True
End Sub

Private Sub cboFactor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtCost.SetFocus
End Sub

Private Sub cboFactor_LostFocus()
    If cboFactor.Text = "" Then
        txtSPFid.Text = ""
        txtMarkUp.Text = ""
        txtVat.Text = ""
        txtMaxDisc.Text = ""
        If txtPrice.Text = txtPrice.Tag Then txtPrice.Text = ""
        txtPrice.Locked = False
        txtMaxDisc.Locked = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Dim i As Integer
    Dim Tmp() As String
        
    If MySource <> "SPORDER COST" Then
        If Val(SPordering.flxParts.TextMatrix(sRow, 6)) = 0 Then
            SPordering.flxParts.TextMatrix(sRow, 1) = ""
        Else
            If SPordering.flxSrch.Tag <> "" Then
                Tmp = Split(SPordering.flxSrch.Tag, "|")
                For i = LBound(Tmp) To UBound(Tmp)
                    SPordering.flxParts.TextMatrix(sRow, i + 1) = Tmp(i)
                Next i
            End If
        End If
    End If
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsParts As New ADODB.Recordset
    Dim MyVal As String, myPno As String
    Dim Ans As Integer, mySPID As Long
    
    If Trim(txtAppModel.Text) = "" Then MsgBox ("Incomplete entry on Applicable Model"), vbExclamation: Exit Sub

    If MySource = "SPORDER COST" Then   'Validate only if new SP
        With SPordering
            If Not UseflxNewSP Then   'Existing PartNo update cost in flxparts
                .flxParts.TextMatrix(sRow, 4) = Format$(txtCost.Text, "#,##0.00")
                .flxParts.TextMatrix(sRow, 9) = txtPrice.Text
                .flxParts.TextMatrix(sRow, 10) = txtSPFid.Text
                .flxParts.TextMatrix(sRow, 11) = txtMaxDisc.Text
                .flxParts.TextMatrix(sRow, 7) = .flxParts.TextMatrix(sRow, 7) & "| Ecost"
                
            Else
                .flxParts.TextMatrix(sRow, 4) = Format$(txtCost.Text, "#,##0.00")
                .flxNewSP.TextMatrix(flxNewRow, 4) = txtCost.Text
                .flxNewSP.TextMatrix(flxNewRow, 5) = txtPrice.Text
                .flxNewSP.TextMatrix(flxNewRow, 6) = txtSPFid.Text
                .flxNewSP.TextMatrix(flxNewRow, 7) = txtMaxDisc.Text
                
            End If
            ' Add Applicable model
            .flxParts.TextMatrix(sRow, 13) = txtAppModel.Text
        
        End With
        Unload Me
        Exit Sub
    End If
    MyVal = ValidatePartNO
    If MyVal <> "" Then
        MsgBox MyVal, vbCritical + vbOKOnly, "Validation Error"
        Exit Sub
    End If
    With SPordering
        myPno = txtPartNo.Text
        If flxNewRow = 0 Then
            .flxNewSP.AddItem myPno & vbTab & txtDesc.Text & vbTab & txtSupCode.Text & vbTab & txtType.Text & vbTab & txtCost.Text & vbTab & txtPrice.Text & vbTab & txtSPFid.Text & vbTab & txtMaxDisc.Text & vbTab & sRow & vbTab & vbTab & txtAppModel.Text
    
        Else
            .flxNewSP.TextMatrix(flxNewRow, 0) = myPno
            .flxNewSP.TextMatrix(flxNewRow, 1) = txtDesc.Text
            .flxNewSP.TextMatrix(flxNewRow, 2) = txtSupCode.Text
            .flxNewSP.TextMatrix(flxNewRow, 3) = txtType.Text
            .flxNewSP.TextMatrix(flxNewRow, 4) = txtCost.Text
            .flxNewSP.TextMatrix(flxNewRow, 5) = txtPrice.Text
            .flxNewSP.TextMatrix(flxNewRow, 6) = txtSPFid.Text
            .flxNewSP.TextMatrix(flxNewRow, 7) = txtMaxDisc.Text
            .flxNewSP.TextMatrix(flxNewRow, 8) = sRow
        End If
        'Display New SpareParts in Order
        .txtPartNo.Text = myPno
        .flxParts.TextMatrix(sRow, 1) = myPno
        .flxParts.TextMatrix(sRow, 2) = txtDesc.Text
        .flxParts.TextMatrix(sRow, 4) = Format$(txtCost.Text, "#,##0.00")
        If Val(.flxParts.TextMatrix(sRow, 6)) > 0 Then
            .flxDel.AddItem .flxParts.TextMatrix(sRow, 6)
            .flxParts.TextMatrix(sRow, 6) = ""
        End If
        If flxNewRow = 0 Then .flxParts.TextMatrix(sRow, 7) = "NEW:" & SPordering.flxNewSP.Rows - 1
                
        'Save spareparts to source disk
        
        'Add applicable model
        .flxNewSP.TextMatrix(flxNewRow, 10) = txtAppModel.Text
        .flxParts.TextMatrix(sRow, 13) = txtAppModel.Text

    End With
        
    Unload Me
End Sub

Function ValidatePartNO() As String
    Dim eMsg As String, Qry As String
    
    eMsg = ""
    If InStr(txtType.Text, "GENUINE") > 0 Then
        If txtPartNo.Text = "" Then
            eMsg = "Incomplete entry on field Part Number"
        Else    'Validate existing Part Number
            'Qry = "SELECT SPAREPARTS.PARTNUMBER From SPAREPARTS WHERE (((SPAREPARTS.PARTNUMBER)='" & txtPartNo.Text & "'));"
            
            'mySQL
            Qry = "SELECT SPAREPARTS.PARTNUMBER From SPAREPARTS WHERE SPAREPARTS.PARTNUMBER='" & txtPartNo.Text & "';"
            
            If CtrlNumExist(Qry) Then
                eMsg = "Part Number " & txtPartNo.Text & " already exist"
            Else
                If PartsInSPcon(txtPartNo.Text) = True Then eMsg = "Part Number " & txtPartNo.Text & " already exist"
            End If
        End If
    Else
        If lblPartNO.Caption = "" Then
            eMsg = "Incomplete entry on field Part Number"
        Else
            'Qry = "SELECT SPAREPARTS.PARTNUMBER From SPAREPARTS WHERE (((SPAREPARTS.PARTNUMBER)='" & lblPartNO.Caption & "'));"
            'mySQL
            Qry = "SELECT SPAREPARTS.PARTNUMBER From SPAREPARTS WHERE SPAREPARTS.PARTNUMBER='" & lblPartNO.Caption & "';"
            
            If CtrlNumExist(Qry) Then
                eMsg = "Part Number " & lblPartNO.Caption & " already exist"
            Else
                If PartsInSPcon(txtPartNo.Text) = True Then eMsg = "Part Number " & lblPartNO.Caption & " already exist"
            End If
        End If
    End If
    If txtDesc.Text = "" Then eMsg = eMsg & vbCrLf & "Incomplete entry on field Description"
    If Val(txtMaxDisc.Text) > 100 Then eMsg = eMsg & vbCrLf & "Maximum Discount cannot exceed 100"
End Function

Function PartsInSPcon(MyPartNo As String) As Boolean
    Dim rsFindSP As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT SOURCESP.PARTNUMBER From SOURCESP WHERE (((SOURCESP.PARTNUMBER)='" & MyPartNo & "'));"
    'mySQL
    Qry = "SELECT SOURCESP.PARTNUMBER From SOURCESP WHERE SOURCESP.PARTNUMBER='" & MyPartNo & "';"
    
    With rsFindSP
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            PartsInSPcon = IIf(IsNull(!PartNumber), False, True)
        Else
            PartsInSPcon = False
        End If
        .Close
    End With
End Function

Private Sub Form_Load()
    Dim rsXParts As New ADODB.Recordset
    Dim Qry As String, myPN As String, myRem As String, myFactorDesc As String
    Dim Dash1 As Integer, Dash2 As Integer, SepPos As Integer, Dgts As Integer
    
    
    With SPordering
        myRem = .flxParts.TextMatrix(sRow, 7)
        If MySource = "SPORDER COST" Then   'Change in cost
            txtPartNo1.Visible = False
            txtPartNo2.Visible = False
            txtPartNo3.Visible = False
            lblPartNO.Visible = False
            lblGPnum.Visible = False
            Label3.Visible = True
            txtPartNo.Visible = True
            Label3.Top = txtPartNo.Top
            txtType.Locked = True
            txtDesc.Locked = True
            txtPartNo.Locked = True
            txtSupplier.Locked = True
            If .txtScode.Text <> "" Then
                'Qry = "SELECT SUPPLIER.SPTYPE From SUPPLIER WHERE (((SUPPLIER.SUPPLIERCODE)='" & .txtScode.Text & "'));"
                'mySQL
                Qry = "SELECT SUPPLIER.SPTYPE From SUPPLIER WHERE SUPPLIER.SUPPLIERCODE='" & .txtScode.Text & "';"
                
                txtType.Text = GetFieldValue(Qry)
            End If
            txtDesc.Text = .flxParts.TextMatrix(sRow, 2)
            txtPartNo.Text = .flxParts.TextMatrix(sRow, 1)
            txtSupplier.Text = .cboSupplier.Text
            txtSupCode.Text = .txtScode.Text
            'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.COST, SPAREPARTS.STOREPRICE, SPAREPARTS.SPFID, SPAREPARTS.SUPPLIERCODE, SPAREPARTS.MAXDISCOUNT " & _
                  "From SPAREPARTS WHERE (((SPAREPARTS.PARTNUMBER)='" & txtPartNo.Text & "'));"
            'mySQL
            'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.COST, SPAREPARTS.STOREPRICE, SPAREPARTS.SPFID, SPAREPARTS.SUPPLIERCODE, SPAREPARTS.MAXDISCOUNT " & _
                  "From SPAREPARTS WHERE SPAREPARTS.PARTNUMBER='" & txtPartNo.Text & "';"
            
            'mySQL
            Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.COST, SPAREPARTS.STOREPRICE, SPAREPARTS.SPFID, SPAREPARTS.SUPPLIERCODE, SPAREPARTS.MAXDISCOUNT, SPAREPARTS.APPLICMODEL " & _
                  "From SPAREPARTS WHERE SPAREPARTS.PARTNUMBER='" & txtPartNo.Text & "';"
          
            rsXParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If rsXParts.RecordCount > 0 Then
                rsXParts.MoveFirst
                If Val(.flxParts.TextMatrix(sRow, 4)) > 0 Then
                    txtCost.Text = .flxParts.TextMatrix(sRow, 4)
                Else
                    txtCost.Text = Format$(rsXParts!Cost, "#####0.00")
                End If
                If Val(.flxParts.TextMatrix(sRow, 9)) > 0 Then
                    txtPrice.Text = Format$(.flxParts.TextMatrix(sRow, 9), "#####0.00")
                Else
                    txtPrice.Text = Format$(rsXParts!StorePrice, "#####0.00")
                End If
                If Val(.flxParts.TextMatrix(sRow, 11)) > 0 Then
                    txtMaxDisc.Text = .flxParts.TextMatrix(sRow, 11)
                Else
                    txtMaxDisc.Text = rsXParts!MaxDiscount
                End If
                If Val(.flxParts.TextMatrix(sRow, 10)) > 0 Then
                    txtSPFid.Text = .flxParts.TextMatrix(sRow, 10)
                Else
                    txtSPFid.Text = rsXParts!SPFid
                End If
                
                If Not IsNull(rsXParts!SUPPLIERCODE) Then
                    If rsXParts!SUPPLIERCODE <> txtSupCode.Text And Len(rsXParts!SUPPLIERCODE) <> 0 Then
                        MsgBox "Warning!!! Part Number " & txtPartNo.Text & " has differnt Supplier!" & vbCrLf & _
                        "It is advice to create a different Part Number!", vbExclamation + vbOKOnly, "Different Supplier"
                    End If
                End If
            Else    'new SP and Previously entered in flxNew
                If InStr(myRem, "NEW:") > 0 Then
                    SepPos = InStr(myRem, "|")
                    Dgts = IIf(SepPos = 0, Len(myRem) - 4, SepPos - 5)
                    flxNewRow = Val(Mid$(myRem, 5, Dgts))
                    txtCost.Text = Val(Replace(.flxNewSP.TextMatrix(flxNewRow, 4), ",", ""))
                    txtPrice.Text = .flxNewSP.TextMatrix(flxNewRow, 5)
                    txtMaxDisc.Text = .flxNewSP.TextMatrix(flxNewRow, 7)
                    txtSPFid.Text = .flxNewSP.TextMatrix(flxNewRow, 6)
                    
                Else
                    txtCost.Text = Val(Replace(.flxParts.TextMatrix(sRow, 4), ",", ""))
                    txtPrice.Text = .flxParts.TextMatrix(sRow, 9)
                    txtMaxDisc.Text = .flxParts.TextMatrix(sRow, 11)
                    txtSPFid.Text = .flxParts.TextMatrix(sRow, 10)
                End If
            End If
            rsXParts.Close
            'setFocus alternative
            'cboFactor.TabIndex = 0
            txtPartNo.TabIndex = 0
        Else
                'Hide partno for replacement
            txtPartNo1.Visible = False
            txtPartNo2.Visible = False
            txtPartNo3.Visible = False
            lblPartNO.Visible = False
            lblGPnum.Visible = False
            Label3.Visible = True
            txtPartNo.Visible = True
            Label3.Top = txtPartNo.Top
            txtPartNo.Text = .flxParts.TextMatrix(sRow, 1)
            txtSupplier.Text = .cboSupplier.Text
            txtSupCode.Text = .txtScode.Text
            If .txtScode.Text <> "" Then
                'Qry = "SELECT SUPPLIER.SPTYPE From SUPPLIER WHERE (((SUPPLIER.SUPPLIERCODE)='" & .txtScode.Text & "'));"
               
                
                'mySQL
                Qry = "SELECT SUPPLIER.SPTYPE From SUPPLIER WHERE SUPPLIER.SUPPLIERCODE='" & .txtScode.Text & "';"
                 txtType.Text = GetFieldValue(Qry)
            End If
            txtDesc.TabIndex = 0
        End If
        
        'Add Applicable model
        If .flxParts.TextMatrix(sRow, 13) <> "" Then
            txtAppModel.Text = .flxParts.TextMatrix(sRow, 13)
        End If
        
    End With
    Qry = "SELECT SPFACTOR.* From SPFACTOR ORDER BY SPFACTOR.SPFDESC;"
    With rsFactor
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboFactor.Clear
            Do While Not .EOF
                cboFactor.AddItem .Fields("SPFdesc")
                If MySource = "SPORDER COST" And Val(txtSPFid.Text) > 0 Then
                    If Val(txtSPFid.Text) = !SPFid Then myFactorDesc = !SPFDesc
                End If
                .MoveNext
            Loop
        End If
    End With
    cboFactor.Text = myFactorDesc
    If myRem = "" Then
        flxNewRow = 0
        UseflxNewSP = False
    Else
        If InStr(myRem, "NEW:") > 0 Then
            SepPos = InStr(myRem, "|")
            Dgts = IIf(SepPos = 0, Len(myRem) - 4, SepPos - 5)
            flxNewRow = Val(Mid$(myRem, 5, Dgts))
            UseflxNewSP = True
        Else
            SPordering.flxDel.AddItem SPordering.flxParts.TextMatrix(sRow, 7)
            flxNewRow = 0
            UseflxNewSP = False
        End If
    End If
    
    
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Frame1.Visible = False
    lblGPnum.ForeColor = vbBlack
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsFactor.State = adStateOpen Then rsFactor.Close
End Sub

Private Sub lblGPnum_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Frame1.Top = txtDesc.Top
    Frame1.Visible = True
    Frame1.ZOrder
    lblGPnum.ForeColor = vbRed
End Sub



Private Sub txtAppmodel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then txtMaxDisc.SetFocus
    If KeyCode = vbKeyDown Then cmdOK.SetFocus
End Sub
Private Sub txtAppmodel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOK.SetFocus
End Sub

Private Sub txtCost_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        txtPrice.SetFocus
    ElseIf KeyCode = vbKeyUp Then
        cboFactor.SetFocus
    End If
End Sub

Private Sub txtCost_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then txtPrice.SetFocus
End Sub

Private Sub txtCost_LostFocus()
    Dim NewPrice As Currency, CostVat As Currency, MyCost As Currency
    
    If Not IsNumeric(txtCost.Text) Then
        MsgBox "Invalid Value on field Cost!", vbExclamation + vbOKOnly, "Validation Error"
        Exit Sub
    End If
    If Val(txtSPFid.Text) > 0 Then
        MyCost = PureNum(txtCost.Text)
        CostVat = MyCost + (MyCost * (Val(txtVat.Text) / 100))
        NewPrice = CostVat + (CostVat * (Val(txtMarkUp.Text) / 100))
        txtPrice.Text = Format$(Int(NewPrice + 0.6) + 1, "####0.00")
        txtPrice.Tag = NewPrice
    End If
End Sub

Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        If InStr(txtType.Text, "GENUINE") > 0 Then
            txtPartNo.SetFocus
        Else
            txtPartNo1.SetFocus
        End If
    End If
    If KeyCode = vbKeyDown Then txtSupplier.SetFocus
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtSupplier.SetFocus
End Sub

Private Sub txtMaxDisc_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyUp Then txtPrice.SetFocus
    'If KeyCode = vbKeyDown Then cmdOK.SetFocus
    If KeyCode = vbKeyUp Then txtPrice.SetFocus
    If KeyCode = vbKeyDown Then txtAppModel.SetFocus
    
End Sub

Private Sub txtMaxDisc_KeyPress(KeyAscii As Integer)
    'KeyAscii = Asc(UCase$(Asc(KeyAscii)))
    'If KeyAscii = 13 Then cmdOK.SetFocus
    If KeyAscii = 13 Then txtAppModel.SetFocus
End Sub

Private Sub txtPartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then txtDesc.SetFocus
End Sub

Private Sub txtPartNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtDesc.SetFocus
End Sub

Private Sub txtPartNo1_Change()
    If txtPartNo1.Text = "" And txtPartNo2.Text = "" And txtPartNo3.Text = "" Then
        lblPartNO.Caption = ""
    Else
        lblPartNO.Caption = txtPartNo1.Text & IIf(txtPartNo1.Text <> "" And txtPartNo2.Text <> "", "-", "") & txtPartNo2.Text & IIf((txtPartNo1.Text <> "" Or txtPartNo2.Text <> "") And txtPartNo3.Text <> "", "-", "") & txtPartNo3.Text
    End If
End Sub

Private Sub txtPartNo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyRight Then txtPartNo2.SetFocus
End Sub

Private Sub txtPartNo1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtPartNo2.SetFocus
End Sub

Private Sub txtPartNo2_Change()
    If txtPartNo1.Text = "" And txtPartNo2.Text = "" And txtPartNo3.Text = "" Then
        lblPartNO.Caption = ""
    Else
        lblPartNO.Caption = txtPartNo1.Text & IIf(txtPartNo1.Text <> "" And txtPartNo2.Text <> "", "-", "") & txtPartNo2.Text & IIf((txtPartNo1.Text <> "" Or txtPartNo2.Text <> "") And txtPartNo3.Text <> "", "-", "") & txtPartNo3.Text
    End If
End Sub

Private Sub txtPartNo2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyRight Then txtPartNo2.SetFocus
    If KeyCode = vbKeyLeft Then txtPartNo1.SetFocus
End Sub

Private Sub txtPartNo2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtPartNo3.SetFocus
End Sub

Private Sub txtPartNo3_Change()
    If txtPartNo1.Text = "" And txtPartNo2.Text = "" And txtPartNo3.Text = "" Then
        lblPartNO.Caption = ""
    Else
        lblPartNO.Caption = txtPartNo1.Text & IIf(txtPartNo1.Text <> "" And txtPartNo2.Text <> "", "-", "") & txtPartNo2.Text & IIf((txtPartNo1.Text <> "" Or txtPartNo2.Text <> "") And txtPartNo3.Text <> "", "-", "") & txtPartNo3.Text
    End If
End Sub

Private Sub txtPartNo3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyLeft Then txtPartNo2.SetFocus
End Sub

Private Sub txtPartNo3_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtDesc.SetFocus
End Sub

Private Sub txtPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then txtMaxDisc.SetFocus
    If KeyCode = vbKeyUp Then txtCost.SetFocus
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then txtMaxDisc.SetFocus
End Sub

Private Sub txtSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then txtDesc.SetFocus
    If KeyCode = vbKeyDown Then txtType.SetFocus
End Sub

Private Sub txtSupplier_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtType.SetFocus
End Sub

Private Sub txtType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then txtSupplier.SetFocus
    If KeyCode = vbKeyDown Then cboFactor.SetFocus
End Sub

Private Sub txtType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cboFactor.SetFocus
End Sub
