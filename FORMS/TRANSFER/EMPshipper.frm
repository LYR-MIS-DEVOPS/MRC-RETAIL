VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form EMPshipper 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shipper Information"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   150
   ClientWidth     =   6480
   Icon            =   "EMPshipper.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   6480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3120
      Top             =   4800
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3120
      Top             =   5280
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3720
      Top             =   4800
   End
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3720
      Top             =   5280
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   6045
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtEntry 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   9
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   9
         Top             =   3120
         Width           =   3150
      End
      Begin VB.TextBox txtEntry 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   7
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   7
         Top             =   2412
         Width           =   3150
      End
      Begin VB.TextBox txtEntry 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   8
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   8
         Top             =   2760
         Width           =   3150
      End
      Begin VB.TextBox txtEntry 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   2280
         MaxLength       =   25
         TabIndex        =   3
         Top             =   1044
         Width           =   3150
      End
      Begin VB.TextBox txtEntry 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   2
         Top             =   702
         Width           =   3150
      End
      Begin VB.TextBox txtEntry 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   1
         Top             =   360
         Width           =   3150
      End
      Begin VB.TextBox txtEntry 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   2280
         MaxLength       =   25
         TabIndex        =   6
         Top             =   2070
         Width           =   3150
      End
      Begin VB.TextBox txtEntry 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   2280
         MaxLength       =   25
         TabIndex        =   5
         Top             =   1728
         Width           =   3150
      End
      Begin VB.TextBox txtEntry 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   2280
         MaxLength       =   25
         TabIndex        =   4
         Top             =   1386
         Width           =   3150
      End
      Begin VB.Label lblI_fld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HELPER (2)"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   3
         Left            =   1275
         TabIndex        =   22
         Top             =   3180
         Width           =   870
      End
      Begin VB.Label lblI_fld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FAX NO."
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   2
         Left            =   1515
         TabIndex        =   21
         Top             =   1446
         Width           =   630
      End
      Begin VB.Label lblI_fld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TEL. NO."
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   1
         Left            =   1470
         TabIndex        =   20
         Top             =   1104
         Width           =   675
      End
      Begin VB.Label lblI_fld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HELPER (1)"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   0
         Left            =   1275
         TabIndex        =   19
         Top             =   2820
         Width           =   870
      End
      Begin VB.Label lblI_fld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DRIVER"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   8
         Left            =   1560
         TabIndex        =   14
         Top             =   2472
         Width           =   615
      End
      Begin VB.Label lblI_fld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SHIPPER NAME"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   4
         Left            =   930
         TabIndex        =   13
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label lblI_fld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   5
         Left            =   1365
         TabIndex        =   12
         Top             =   762
         Width           =   780
      End
      Begin VB.Label lblI_fld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TRUCK NAME"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   6
         Left            =   1080
         TabIndex        =   11
         Top             =   1788
         Width           =   1065
      End
      Begin VB.Label lblI_fld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PLATE NO."
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   7
         Left            =   1305
         TabIndex        =   10
         Top             =   2130
         Width           =   840
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
            Picture         =   "EMPshipper.frx":3072
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EMPshipper.frx":3F4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EMPshipper.frx":6086
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EMPshipper.frx":6F60
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EMPshipper.frx":7E3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EMPshipper.frx":8D14
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
      Image           =   "EMPshipper.frx":BD96
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
      Image           =   "EMPshipper.frx":C330
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
      Image           =   "EMPshipper.frx":C8CA
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
      Image           =   "EMPshipper.frx":CE64
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
      Image           =   "EMPshipper.frx":D3FE
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
      Image           =   "EMPshipper.frx":D998
      cBack           =   -2147483633
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SHIPPER INFORMATION"
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
      TabIndex        =   17
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
      TabIndex        =   15
      Top             =   545
      Width           =   3500
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   6300
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   6300
      Y1              =   4560
      Y2              =   4560
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
      TabIndex        =   16
      Top             =   570
      Width           =   3500
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SHIPPER INFORMATION"
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
      TabIndex        =   18
      Top             =   570
      Width           =   2415
   End
End
Attribute VB_Name = "EMPshipper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsShipper As New ADODB.Recordset
Dim AddNew As Boolean
Dim BtnIndx As Integer

Sub ClearEntries()
    Dim i As Integer
    
    For i = 0 To 9
        txtEntry(i).Text = ""
    Next i
End Sub

Sub LockData(Bln As Boolean)
    Dim i As Integer
    
    For i = 1 To 9
        txtEntry(i).Locked = Bln
    Next i
End Sub

Sub tbSaveCancel(Bln As Boolean)
    Dim i As Integer

    For i = 1 To 4
        tbShipper.Buttons(i).Visible = Bln
    Next i
    For i = 5 To 12
        tbShipper.Buttons(i).Visible = Not Bln
    Next i
End Sub
Sub SaveShipper()
    'Dim X As Integer
    '***Access
    Dim Qry As String
    
    'If AddNew Then
    '    rsShipper.Open "SHIPPER", LYRcon, adOpenKeyset, adLockOptimistic
    '    rsShipper.AddNew
    'Else
    '    Qry = "SELECT SHIPPER.* From SHIPPER " & _
              "WHERE (((SHIPPER.SHIPPERID)=" & txtEntry(0).Text & "));"
        
    '    rsShipper.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    '    rsShipper.MoveFirst
    'End If
    'With rsShipper
    '    For X = 1 To 9
    '        .Fields(X) = txtEntry(X).Text
    '    Next X
    '    .Fields("UpdatedBy") = UserID
    '    .Fields("LastUpdate") = Date
    '    .Update
    '    If AddNew Then txtEntry(0).Text = !SHIPPERID
    '    .Close
    
    'mySQL Statement
       
    If AddNew Then
        LYRcon.Execute "Insert Into shipper(SHIPPERNAME, ADDRESS, CONTACTNO, FAXNO, TRUCKNAME, PLATENO, DRIVER, " & _
            "HELPER1, HELPER2, UPDATEDBY, LASTUPDATE)VALUES('" & txtEntry(1).Text & "', '" & txtEntry(2).Text & "', " & _
            "'" & txtEntry(3).Text & "', " & txtEntry(4).Text & "', " & txtEntry(5).Text & "', '" & txtEntry(6).Text & "', " & _
            "'" & txtEntry(7).Text & "', '" & txtEntry(8).Text & "', '" & txtEntry(9).Text & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "')"
        Qry = "SELECT shipper.SHIPPERID, shipper.SHIPPERNAME " & _
            "From shipper Where shipper.SHIPPERNAME = '" & txtEntry(1).Text & "';"
        txtEntry(0).Text = GetFieldValue(Qry)
    Else
        LYRcon.Execute "Update shipper SET SHIPPERNAME = '" & txtEntry(1).Text & "', ADDRESS = '" & txtEntry(2).Text & "', " & _
            "CONTACTNO = '" & txtEntry(3).Text & "', FAXNO = '" & txtEntry(4).Text & "', TRUCKNAME = '" & txtEntry(5).Text & "', " & _
            "PLATENO = '" & txtEntry(6).Text & "', DRIVER = '" & txtEntry(7).Text & "', HELPER1 = '" & txtEntry(8).Text & "', " & _
            "HELPER2 = '" & txtEntry(9).Text & "', UPDATEDBY = " & UserID & ", LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "'" & _
            " Where shipper.SHIPPERID = " & txtEntry(0).Text & ""
    End If
    'End With
End Sub
Sub ShowRecord(SHIPPERID As Integer)
    Dim Qry As String
    Dim X As Integer
    '***Access
    'Qry = "SELECT SHIPPER.* From SHIPPER " & _
          "WHERE (((SHIPPER.SHIPPERid)=" & SHIPPERID & "));"
    'mySQL
    Qry = "SELECT SHIPPER.* From SHIPPER " & _
          "WHERE SHIPPER.SHIPPERid=" & SHIPPERID & ";"
    rsShipper.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    With rsShipper
        .MoveFirst
        For X = 0 To 9
            txtEntry(X).Text = IIf(IsNull(.Fields(X)), "", .Fields(X))
        Next X
        .Close
    End With
End Sub
Function ValidateEntry() As String
    Dim myMsg As String
    
    myMsg = ""
    If txtEntry(1).Text = "" Then myMsg = "No data on field SHIPPER NAME"
    If txtEntry(2).Text = "" Then myMsg = myMsg & vbCrLf & "No data on field ADDRESS"
    If txtEntry(5).Text = "" Then myMsg = "No data on field TRUCK NAME"
    If txtEntry(7).Text = "" Then myMsg = myMsg & vbCrLf & "No data on field DRIVER NAME"
    
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
    Dim SHIPPERID As Integer, Ans As Integer
    
    Select Case Index
        Case 0  '   "new"
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=12));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=12;"
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
            txtEntry(1).SetFocus
        Case 1      '"edit"
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=12));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=12;"
            
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            If txtEntry(0).Text = "" Then
                MsgBox "No records found!", vbExclamation + vbOKOnly, "Warning"
                txtEntry(1).SetFocus
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
            txtEntry(1).SetFocus
        Case 2          '"find"
            Set frmSearch.FrmSource = EMPshipper
            QrySrch = "SELECT SHIPPER.SHIPPERID, SHIPPER.SHIPPERNAME " & _
                      "From SHIPPER ORDER BY SHIPPER.SHIPPERNAME;"

            frmSearch.QrySearch = QrySrch
            frmSearch.lblCap.Caption = "Search for Shipper"
            frmSearch.Show vbModal
        Case 3          '"quit"
            MotoRep.tbMain.Visible = True
            Unload Me
        
        Case 4              '"save"
            vMsg = ValidateEntry
            If vMsg <> "" Then
                MsgBox vMsg, vbOKOnly + vbCritical, "Validation Error"
                txtEntry(1).SetFocus
                Exit Sub
            End If
            Ans = MsgBox("Do you want to save shipper?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then
                txtEntry(1).SetFocus
                Exit Sub
            End If
            Call SaveShipper
            MsgBox "Record successfully saved!", vbInformation + vbOKOnly, "Save"
            Call LockData(True)
            txtEntry(1).SetFocus
            If DispBtn Then
                Call BtnSaveCancel(False)
            Else
                BtnIndx = 5
                tmrFloatRev.Enabled = True
            End If
        Case 5          '"cancel"
            Ans = MsgBox("Do you want to cancel Shipper?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            If AddNew Then
                Call ClearEntries
            Else
                SHIPPERID = txtEntry(0).Text
                Call ClearEntries
                Call ShowRecord(SHIPPERID)
            End If
            Call LockData(True)
            txtEntry(1).SetFocus
            If DispBtn Then
                Call BtnSaveCancel(False)
            Else
                BtnIndx = 5
                tmrFloatRev.Enabled = True
            End If
    End Select
End Sub
Private Sub Form_Load()
    Dim i As Integer
    
   
    Call LockData(True)
    
    MotoRep.sBar.Panels(6).Text = "SHIPPER Information entry form"
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



Private Sub TBcmd_Click(Index As Integer)
End Sub

Private Sub tbSHIPPER_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim vMsg As String, QrySrch As String
    Dim SHIPPERID As Integer, Ans As Integer
    
    Select Case Button.Key
        Case "addnew"
            Call LockData(False)
            Call ClearEntries
            Call tbSaveCancel(True)
            AddNew = True
            txtEntry(1).SetFocus
        Case "edit"
            If txtEntry(0).Text = "" Then
                MsgBox "No records found!", vbExclamation + vbOKOnly, "Warning"
                txtEntry(1).SetFocus
                Exit Sub
            End If
            Call LockData(False)
            Call tbSaveCancel(True)
            AddNew = False
            txtEntry(1).SetFocus
        Case "save"
            vMsg = ValidateEntry
            If vMsg <> "" Then
                MsgBox vMsg, vbOKOnly + vbCritical, "Validation Error"
                txtEntry(1).SetFocus
                Exit Sub
            End If
            Ans = MsgBox("Do you want to save shipper?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then
                txtEntry(1).SetFocus
                Exit Sub
            End If
            Call SaveShipper
            MsgBox "Record successfully saved!", vbInformation + vbOKOnly, "Save"
            Call tbSaveCancel(False)
            Call LockData(True)
            txtEntry(1).SetFocus
        Case "cancel"
            Ans = MsgBox("Are you sure you want to cancel?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            If AddNew Then
                Call ClearEntries
            Else
                SHIPPERID = txtEntry(0).Text
                Call ClearEntries
                Call ShowRecord(SHIPPERID)
            End If
            Call LockData(True)
            Call tbSaveCancel(False)
            txtEntry(1).SetFocus
        Case "search"
            Set frmSearch.FrmSource = EMPshipper
            QrySrch = "SELECT SHIPPER.SHIPPERID, SHIPPER.SHIPPERNAME " & _
                      "From SHIPPER ORDER BY SHIPPER.SHIPPERNAME;"

            frmSearch.QrySearch = QrySrch
            frmSearch.lblCap.Caption = "Search for SHIPPER"
            frmSearch.Show vbModal
        Case "quit"
            MotoRep.tbMain.Visible = True
            Unload Me
    End Select

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






