VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBackUp 
   Caption         =   "Backup"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4350
   Icon            =   "LogID.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4350
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar PbaR 
      Height          =   135
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   2040
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   3975
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   405
      Left            =   1560
      TabIndex        =   0
      Top             =   2160
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   714
      Caption         =   "OK"
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
      Image           =   "LogID.frx":058A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   405
      Left            =   2880
      TabIndex        =   1
      Top             =   2160
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   714
      Caption         =   "Cancel"
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
      Image           =   "LogID.frx":0B24
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxT 
      Height          =   6375
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   11245
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      FormatString    =   "^TABLE NAME                                                    "
   End
   Begin VB.Label Label1 
      Caption         =   "Select Drive and Folder:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   150
      Width           =   1935
   End
End
Attribute VB_Name = "frmBackUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim db As DAO.Database
    Dim i As Integer, Ans As Integer
    Dim mySrc As String, TargetFile As String, Qry As String, MyTime As String, MyAcc As String
    
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=90));"
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    
    'validate if Blank DB Exist
    On Error GoTo LocalErr
    mySrc = Dir(App.Path & "\blankmtr.mdb")
    If mySrc = "" Then
        MsgBox "Some Components in backup is missing!" & vbCrLf & "Please contact MIS Department", vbInformation + vbOKOnly
        Exit Sub
    End If
    If Dir1.Path = App.Path Then Exit Sub
    TargetFile = Dir1.Path & IIf(Right(Dir1.Path, 1) <> "\", "\MTREPTGM.MDB", "MTREPTGM.MDB")
    'Save remarks
    Qry = "SELECT SETTINGS.DESCRIPTION From SETTINGS WHERE (((SETTINGS.DESCRIPTION)='LAST BACKUP'));"
    MyTime = Now
    If CtrlNumExist(Qry) Then
        LYRcon.Execute "UPDATE SETTINGS SET SETTINGS.REMARKS = '" & MyTime & "',  SETTINGS.UPDATEDBY = " & UserID & ", SETTINGS.LASTUPDATE = #" & Now & "# WHERE (((SETTINGS.DESCRIPTION)='LAST BACKUP'));"
    Else
        LYRcon.Execute "INSERT INTO SETTINGS ( DESCRIPTION, REMARKS, UPDATEDBY, LASTUPDATE) " & _
                       "VALUES ('LAST BACKUP', '" & MyTime & "', " & UserID & ", #" & Now & "#);"
    End If
    If Dir(TargetFile) = "" Then
        FileCopy mySrc, TargetFile
    Else
        Ans = MsgBox("A copy of the Database already exist in Target Drive" & vbCrLf & "Do you want to continue?", vbQuestion + vbYesNo)
        If Ans = vbNo Then Exit Sub
        FileCopy mySrc, TargetFile
    End If
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
    
    Pbar.Value = 0
    With flxT
        Pbar.Max = .Rows
        For i = 1 To .Rows - 1
            Call CopyTable(TargetFile, .TextMatrix(i, 0), .TextMatrix(i, 0))
            Pbar.Value = Pbar.Value + 1
            DoEvents
        Next i
    End With
    Set db = DBEngine.OpenDatabase(TargetFile, True, False, ";pwd=")
    db.NewPassword "", "LCRJR"
    db.Close
    MsgBox "Backup successfully made!", vbInformation + vbOKOnly
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
    Unload Me
    Exit Sub
LocalErr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error " & Err.Number
    Unload Me
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    Dir1.Path = "C:\"
    With flxT
        .Rows = 1
        .AddItem "ADJCLERICAL"
        .AddItem "ADJCNTR"
        .AddItem "ADJDMCM"
        .AddItem "AGINGAR"
        .AddItem "AUDITINVDTLS"
        .AddItem "AUDITINVEN"
        .AddItem "AUDITRANS"
        .AddItem "BACKORDERDTLS"
        .AddItem "BEGINDESC"
        .AddItem "BILLING"
        .AddItem "BILLINGDTLS"
        .AddItem "BRANCH"
        .AddItem "CASHFLOW"
        .AddItem "CASHFLOWMGR"
        .AddItem "CASHVOUCHER"
        .AddItem "CFATTACHMENT"
        .AddItem "CHECKALLOC"
        .AddItem "CHECKS"
        .AddItem "CUSTOMER"
        .AddItem "CUSTPO"
        .AddItem "DISBURSEACCTDESC"
        .AddItem "DISBURSEMENT"
        .AddItem "DPLEDGER"
        .AddItem "EMPLOYEES"
        .AddItem "INVADJCHARGES"
        .AddItem "INVADJDETAILS"
        .AddItem "INVNTRYADJ"
        .AddItem "LEDGER"
        .AddItem "MCRECON"
        .AddItem "MCSERVICELABOR"
        .AddItem "MSALES"
        .AddItem "ONACCTERMS"
        .AddItem "ONACCTPCV"
        .AddItem "ORDETAILS"
        .AddItem "PAYMENTS"
        .AddItem "PCUSER"
        .AddItem "PREVCUSTOMER"
        .AddItem "PRINTPREF"
        .AddItem "PURCHASERET"
        .AddItem "PURCHRETDTLS"
        .AddItem "REQUISITION"
        .AddItem "SETTINGS"
        .AddItem "SHIPPER"
        .AddItem "SOURCESP"
        .AddItem "SPAREPARTS"
        .AddItem "SPFACTOR"
        .AddItem "SPGFACTOR"
        .AddItem "SPINVENTORY"
        .AddItem "SPINVOICE"
        .AddItem "SPITR"
        .AddItem "SPORDER"
        .AddItem "SPRECEIVE"
        .AddItem "SPSALESORDER"
        .AddItem "SPSORDERDETAIL"
        .AddItem "SRETURNDTLS"
        .AddItem "SRETURNSP"
        .AddItem "SRVCJOBORDER"
        .AddItem "SRVCUSTOMER"
        .AddItem "SRVJOBDETAILS"
        .AddItem "SUPPLIER"
        .AddItem "TRANSFERLOC"
        .AddItem "TRANSLOCDTLS"
        .AddItem "USERACCESS"
    End With
End Sub

Private Sub CopyTable(ByVal DBPathName$, ByVal TargetTable$, ByVal NewTableName$)
On Error GoTo ErrorHandler:
    LYRcon.Execute "SELECT " & TargetTable & ".* INTO " & TargetTable & " IN '" & DBPathName & "' From " & NewTableName
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical + vbOKOnly, "Error " & Err.Number & " in CopyTable!"
End Sub

