VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form EMPsupplier 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supplier Information"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   150
   ClientWidth     =   6615
   Icon            =   "EMPsupplier.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2160
      Top             =   5160
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2640
      Top             =   5160
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3720
      Top             =   5160
   End
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3240
      Top             =   5160
   End
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   6045
      Begin VB.TextBox txtSPtype 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1052
         Width           =   3150
      End
      Begin VB.TextBox txtAddress 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1398
         Width           =   3150
      End
      Begin VB.TextBox txtSCode 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   3
         TabIndex        =   2
         Top             =   706
         Width           =   3150
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   2880
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtMobile 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   25
         TabIndex        =   9
         Top             =   3128
         Width           =   3150
      End
      Begin VB.TextBox txtEmail 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   10
         Top             =   3480
         Width           =   3150
      End
      Begin VB.TextBox txtTelNo 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1744
         Width           =   3150
      End
      Begin VB.TextBox txtSName 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   1
         Top             =   360
         Width           =   3150
      End
      Begin VB.TextBox txtPosition 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   25
         TabIndex        =   8
         Top             =   2782
         Width           =   3150
      End
      Begin VB.TextBox txtInCharge 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   7
         Top             =   2436
         Width           =   3150
      End
      Begin VB.TextBox txtFaxNo 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2090
         Width           =   3150
      End
      Begin VB.Label lblI_fld 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SPAREPART TYPE"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   9
         Left            =   660
         TabIndex        =   31
         Top             =   1112
         Width           =   1440
      End
      Begin VB.Label lblI_fld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER CODE"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   3
         Left            =   855
         TabIndex        =   30
         Top             =   766
         Width           =   1290
      End
      Begin VB.Label lblI_fld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FAX NO."
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   2
         Left            =   1515
         TabIndex        =   22
         Top             =   2150
         Width           =   630
      End
      Begin VB.Label lblI_fld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TEL. NO."
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   21
         Top             =   1804
         Width           =   675
      End
      Begin VB.Label lblI_fld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EMAIL ADDRESS"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   20
         Top             =   3540
         Width           =   1305
      End
      Begin VB.Label lblI_fld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MOBILE NO."
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   8
         Left            =   1215
         TabIndex        =   15
         Top             =   3188
         Width           =   930
      End
      Begin VB.Label lblI_fld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER NAME"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   4
         Left            =   840
         TabIndex        =   14
         Top             =   420
         Width           =   1305
      End
      Begin VB.Label lblI_fld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   5
         Left            =   1365
         TabIndex        =   13
         Top             =   1458
         Width           =   780
      End
      Begin VB.Label lblI_fld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONTACT PERSON"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   6
         Left            =   660
         TabIndex        =   12
         Top             =   2496
         Width           =   1485
      End
      Begin VB.Label lblI_fld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "POSITION"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   7
         Left            =   1380
         TabIndex        =   11
         Top             =   2842
         Width           =   765
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   6120
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EMPsupplier.frx":3072
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EMPsupplier.frx":3F4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EMPsupplier.frx":6086
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EMPsupplier.frx":6F60
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EMPsupplier.frx":7E3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EMPsupplier.frx":8D14
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   24
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
      Image           =   "EMPsupplier.frx":BD96
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   25
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
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
      Image           =   "EMPsupplier.frx":C330
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   4320
      TabIndex        =   26
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
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
      Image           =   "EMPsupplier.frx":C8CA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   5760
      TabIndex        =   27
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
      Image           =   "EMPsupplier.frx":CE64
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   7200
      TabIndex        =   28
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
      Image           =   "EMPsupplier.frx":D3FE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   2880
      TabIndex        =   29
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
      Image           =   "EMPsupplier.frx":D998
      cBack           =   -2147483633
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SUPPLIER INFORMATION"
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
      Left            =   3795
      TabIndex        =   18
      Top             =   540
      Width           =   2415
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
      Left            =   240
      TabIndex        =   16
      Top             =   545
      Width           =   3500
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
      Left            =   270
      TabIndex        =   17
      Top             =   570
      Width           =   3500
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SUPPLIER INFORMATION"
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
      Left            =   3825
      TabIndex        =   19
      Top             =   570
      Width           =   2415
   End
End
Attribute VB_Name = "EMPsupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSupplier As New ADODB.Recordset
Dim AddNew As Boolean
Dim BtnIndx As Integer

Sub ClearEntries()
    
    txtEntry(0).Text = ""
    txtSName.Text = ""
    txtScode.Text = ""
    txtAddress.Text = ""
    txtTelNo.Text = ""
    txtFaxNo.Text = ""
    txtInCharge.Text = ""
    txtPosition.Text = ""
    txtMobile.Text = ""
    txtEmail.Text = ""
End Sub

Sub LockData(Bln As Boolean)
    
    txtEntry(0).Locked = Bln
    txtSName.Locked = Bln
    txtScode.Locked = Bln
    txtAddress.Locked = Bln
    txtTelNo.Locked = Bln
    txtFaxNo.Locked = Bln
    txtInCharge.Locked = Bln
    txtPosition.Locked = Bln
    txtMobile.Locked = Bln
    txtEmail.Locked = Bln
End Sub

Sub SaveSupplier()
    Dim X As Integer
    Dim Qry As String
    
    '***Access
    'If AddNew Then
    '    rsSupplier.Open "SUPPLIER", LYRcon, adOpenKeyset, adLockOptimistic
    '    rsSupplier.AddNew
    'Else
    '    Qry = "SELECT SUPPLIER.* From SUPPLIER " & _
    '          "WHERE (((SUPPLIER.SUPPLIERID)=" & txtEntry(0).Text & "));"
    '    rsSupplier.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    '    rsSupplier.MoveFirst
    'End If
    'With rsSupplier
    '    !SUPPLIERNAME = txtSName.Text
    '    !SUPPLIERCODE = txtScode.Text
    '    !SPtype = txtSPtype.Text
    '    !Address = txtAddress.Text
    '    !CONTACTNO = txtTelNo.Text
    '    !FAXNO = txtFaxNo.Text
    '    !ContactPerson = txtInCharge.Text
    '    !CONPERPOSITION = txtPosition.Text
    '    !MOBILENO = txtMobile.Text
    '    !EMAIL = txtEmail.Text
    '    !UPDATEDBY = UserID
    '    !LASTUPDATE = Date
    '    .Update
    '    If AddNew Then txtEntry(0).Text = !SUPPLIERID
    '    .Close
    'End With
    
    'mySQL Statement
    If AddNew Then
        
        LYRcon.Execute "Insert Into supplier(SUPPLIERNAME, SUPPLIERCODE, SPTYPE, ADDRESS, CONTACTNO, FAXNO," & _
            "CONTACTPERSON, CONPERPOSITION, MOBILENO, EMAIL, UPDATEDBY, LASTUPDATE)" & _
            "VALUES('" & txtSName.Text & "', '" & txtScode.Text & "', '" & txtSPtype.Text & "', '" & txtAddress.Text & "', " & _
            "'" & txtTelNo.Text & "', '" & txtFaxNo.Text & "', '" & txtInCharge.Text & "', '" & txtPosition.Text & "', " & _
            "'" & txtMobile.Text & "', '" & txtEmail.Text & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
        
        'Display Supplier Id
        Qry = "SELECT supplier.SUPPLIERID From supplier Where supplier.SUPPLIERNAME = '" & txtSName.Text & "' AND supplier.SUPPLIERCODE = '" & txtScode.Text & "'"
            txtEntry(0).Text = GetFieldValue(Qry)
       
    Else
        
        LYRcon.Execute "Update supplier SET SUPPLIERNAME = '" & txtSName.Text & "', SUPPLIERCODE = '" & txtScode.Text & "', SPTYPE = '" & txtSPtype.Text & "', " & _
            "ADDRESS = '" & txtAddress.Text & "', CONTACTNO = '" & txtTelNo.Text & "', FAXNO = '" & txtFaxNo.Text & "', CONTACTPERSON = '" & txtInCharge.Text & "', " & _
            "CONPERPOSITION = '" & txtPosition.Text & "', MOBILENO = '" & txtMobile.Text & "', EMAIL = '" & txtEmail.Text & "', UPDATEDBY = " & UserID & ", LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' WHERE SUPPLIER.SUPPLIERID=" & txtEntry(0).Text & ";"
        
    End If
End Sub

Sub ShowRecord(SUPPLIERID As Integer)
    Dim Qry As String
    Dim X As Integer
    
    '***Access
    'Qry = "SELECT SUPPLIER.* From SUPPLIER " & _
          "WHERE (((SUPPLIER.SUPPLIERid)=" & SUPPLIERID & "));"
    'mySQL
    Qry = "SELECT SUPPLIER.* From SUPPLIER " & _
          "WHERE SUPPLIER.SUPPLIERid=" & SUPPLIERID & ";"
    
    rsSupplier.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    With rsSupplier
        .MoveFirst
        txtSName.Text = !SUPPLIERNAME
        txtScode.Text = IIf(IsNull(!SUPPLIERCODE), "", !SUPPLIERCODE)
        txtSPtype.Text = IIf(IsNull(!SPtype), "", !SPtype)
        txtAddress.Text = !Address
        txtTelNo.Text = IIf(IsNull(!CONTACTNO), "", !CONTACTNO)
        txtFaxNo.Text = IIf(IsNull(!FAXNO), "", !FAXNO)
        txtInCharge.Text = IIf(IsNull(!ContactPerson), "", !ContactPerson)
        txtPosition.Text = IIf(IsNull(!CONPERPOSITION), "", !CONPERPOSITION)
        txtMobile.Text = IIf(IsNull(!MOBILENO), "", !MOBILENO)
        txtEmail.Text = IIf(IsNull(!EMAIL), "", !EMAIL)
        txtEntry(0).Text = !SUPPLIERID
        .Close
    End With
End Sub
Function ValidateEntry() As String
    Dim myMsg As String, Qry As String
    
    myMsg = ""
    If txtSName.Text = "" Then
        myMsg = "No data on field SUPPLIER NAME"
    Else
        If AddNew Then
            '***Access
            'Qry = "SELECT SUPPLIER.SUPPLIERNAME From SUPPLIER WHERE (((SUPPLIER.SUPPLIERNAME)='" & txtSName.Text & "'));"
            'mySQL
            Qry = "SELECT SUPPLIER.SUPPLIERNAME From SUPPLIER WHERE SUPPLIER.SUPPLIERNAME='" & txtSName.Text & "';"
            If CtrlNumExist(Qry) Then myMsg = "SUPPLIER NAME Already exist!"
        End If
    End If
    'Validate Existing Supplier code
    If txtScode.Text = "" Then
        myMsg = myMsg & vbCrLf & "No data on field SUPPLIER CODE"
    Else
        If AddNew Then
            '***Access
            'Qry = "SELECT SUPPLIER.SUPPLIERCODE From SUPPLIER WHERE (((SUPPLIER.SUPPLIERCODE)='" & txtScode.Text & "'));"
            'mySQL
            Qry = "SELECT SUPPLIER.SUPPLIERCODE From SUPPLIER WHERE SUPPLIER.SUPPLIERCODE='" & txtScode.Text & "';"
            If CtrlNumExist(Qry) Then myMsg = "SUPPLIER CODE Already exist!"
        End If
    End If
    If txtSPtype.Text = "" Then myMsg = myMsg & vbCrLf & "No data on field SPAREPART TYPE"
    If txtAddress.Text = "" Then myMsg = myMsg & vbCrLf & "No data on field ADDRESS"
    ValidateEntry = myMsg
End Function

Private Sub BtnSaveCancel(Bln As Boolean)
    Dim i As Integer
    
    If Bln Then
        For i = 0 To 3
            cmdTrans(i).Visible = False
        Next i
        cmdTrans(4).Visible = True
        cmdTrans(5).Visible = True
        cmdTrans(4).Left = 0
        cmdTrans(5).Left = cmdTrans(4).Width
    Else
        For i = 0 To 3
            cmdTrans(i).Visible = True
        Next i
        cmdTrans(4).Visible = False
        cmdTrans(5).Visible = False
    End If
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim vMsg As String, QrySrch As String, Qry As String, MyAcc As String
    Dim SUPPLIERID As Integer, Ans As Integer
    
    Select Case Index
        Case 0  '   "new"
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=11));"
            'mySQl
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=11;"
            
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
        
            If DispBtn Then
                Call BtnSaveCancel(True)
            Else
                BtnIndx = 3
                tmrFloat.Enabled = True
            End If
                
            Call LockData(False)
            Call ClearEntries
            AddNew = True
            txtSName.SetFocus
        Case 1      '"edit"
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=11));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=11;"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            If txtEntry(0).Text = "" Then
                MsgBox "No records found!", vbExclamation + vbOKOnly, "Warning"
                txtSName.SetFocus
                Exit Sub
            End If
            If DispBtn Then
                Call BtnSaveCancel(True)
            Else
                BtnIndx = 3
                tmrFloat.Enabled = True
            End If
            Call LockData(False)
            AddNew = False
            txtSName.SetFocus
        Case 2          '"find"
            Set frmSearch.FrmSource = EMPsupplier
            QrySrch = "SELECT SUPPLIER.SUPPLIERID, SUPPLIER.SUPPLIERNAME " & _
                      "From SUPPLIER ORDER BY SUPPLIER.SUPPLIERNAME;"

            frmSearch.QrySearch = QrySrch
            frmSearch.lblCap.Caption = "Search for SUPPLIER"
            frmSearch.Show vbModal
        Case 3          '"quit"
            MotoRep.tbMain.Visible = True
            Unload Me
        
        Case 4              '"save"
            vMsg = ValidateEntry
            If vMsg <> "" Then
                MsgBox vMsg, vbOKOnly + vbCritical, "Validation Error"
                txtSName.SetFocus
                Exit Sub
            End If
            Ans = MsgBox("Do you want to save supplier?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then
                txtSName.SetFocus
                Exit Sub
            End If
            Call SaveSupplier
            MsgBox "Record successfully saved!", vbInformation + vbOKOnly, "Save"
            Call LockData(True)
            txtSName.SetFocus
            If DispBtn Then
                Call BtnSaveCancel(False)
            Else
                BtnIndx = 5
                tmrFloatRev.Enabled = True
            End If
        Case 5          '"cancel"
            Ans = MsgBox("Are you sure you want to cancel?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            If AddNew Then
                Call ClearEntries
            Else
                SUPPLIERID = txtEntry(0).Text
                Call ClearEntries
                Call ShowRecord(SUPPLIERID)
            End If
            Call LockData(True)
            txtSName.SetFocus
            If DispBtn Then
                Call BtnSaveCancel(False)
            Else
                BtnIndx = 5
                tmrFloatRev.Enabled = True
            End If
    End Select
End Sub
Private Sub Form_Load()
    
    Call LockData(True)
    
    MotoRep.sBar.Panels(6).Text = "SUPPLIER Information entry form"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
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
        If BtnIndx = -1 Then
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
    Dim i As Integer
    
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top - 50
    If cmdTrans(BtnIndx).Top <= -375 Then
        cmdTrans(BtnIndx).Visible = False
        BtnIndx = BtnIndx - 1
        If BtnIndx = 3 Then
            tmrFloatRev.Enabled = False
            BtnIndx = 0
            For i = 0 To 3
                cmdTrans(i).Visible = True
            Next i
            tmrDropRev.Enabled = True
        End If
    End If
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtEntry_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        SendKeys "{TAB}"
    ElseIf KeyCode = 38 Then
        SendKeys "+{TAB}"
    End If
End Sub

Private Sub txtFaxNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtInCharge_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtMobile_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtPosition_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtSCode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtSName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtSPtype_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtTelNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
