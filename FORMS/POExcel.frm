VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form POexcel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import P.O. from Excel"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ClipControls    =   0   'False
   Icon            =   "POExcel.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   3550
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog Cdlg 
      Left            =   120
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin lvButton.lvButtons_H cmdImport 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Import"
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
      Image           =   "POExcel.frx":058A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      Image           =   "POExcel.frx":0B24
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxPO 
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4260
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   16777215
      ScrollBars      =   2
      FormatString    =   ">N O.    |< P A R T   N U M B E R               |^Q T Y.           "
   End
   Begin lvButton.lvButtons_H cmdOpen 
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Open"
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
      Image           =   "POExcel.frx":10BE
      cBack           =   -2147483633
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
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
      Left            =   3240
      TabIndex        =   5
      Top             =   2820
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "TOTAL:"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   2820
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Part Number to be copied:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "POexcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As DAO.Database
Dim rs As DAO.Recordset
Private filepath As String
Private sheetname As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdImport_Click()
    Dim i As Integer
    
    cmdImport.Enabled = False
    cmdOpen.Enabled = False
    cmdCancel.Enabled = False
    With frmSupplierPO
        pBar.Visible = True
        pBar.Value = 0
        pBar.Max = flxPO.Rows - 1
        .flxParts.Rows = 1
        For i = 1 To flxPO.Rows - 1
            .flxParts.AddItem flxPO.TextMatrix(i, 0) & vbTab & flxPO.TextMatrix(i, 1) & vbTab & vbTab & flxPO.TextMatrix(i, 2)
            pBar.Value = pBar.Value + 1
            DoEvents
        Next i
    End With
    pBar.Visible = False
    Unload Me
    DoEvents
    Call frmSupplierPO.ProcessImportedPO
    
End Sub

Private Sub cmdOpen_Click()
    Dim X As Integer, TtlQty As Long
    Dim zQty As Integer
    Dim InvldRow As String
    
    On Error GoTo LocalErr
    With Cdlg
        .CancelError = True
        .Filter = "(*.xls)|*.xls"
        .ShowOpen
        filepath = .filename
        If filepath = "" Then
            MsgBox "Please select a File Name to continue!", vbExclamation + vbOKOnly, "Import P.O."
            Exit Sub
        End If
    End With
    sheetname = "POSheet$"
    Set db = OpenDatabase(filepath, False, False, "Excel 12.0 Xml;IMEX=1;HDR=yes;")
    Set rs = db.OpenRecordset(sheetname)
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Screen.MousePointer = 11
        X = 1
        TtlQty = 0
        zQty = 0
        flxPO.Rows = 1
        While rs.EOF <> True
            If IsNull(rs.Fields(0)) Or IsNull(rs.Fields(1)) Then
                zQty = zQty + 1
                InvldRow = IIf(InvldRow = "", "Row " & X + 1, InvldRow & ", " & X + 1)
            Else
                If rs.Fields(0) = "" Or Val(rs.Fields(1)) = 0 Then
                    zQty = zQty + 1
                    InvldRow = IIf(InvldRow = "", "Row " & X + 1, InvldRow & ", " & X + 1)
                Else
                    flxPO.AddItem X & vbTab & rs.Fields(0) & vbTab & rs.Fields(1)
                    TtlQty = TtlQty + rs.Fields(1)
                    X = X + 1
                End If
            End If
            rs.MoveNext
        Wend
        Screen.MousePointer = 0
        lblTotal.Caption = TtlQty
        If zQty > 0 Then
            MsgBox zQty & IIf(zQty > 1, " Records", " Record") & " found Invalid!" & vbCrLf & _
            "Record is either blank Part number or zero quantity." & vbCrLf & "Please check source at the following rows:" & vbCrLf & _
            InvldRow, vbExclamation + vbOKOnly, "Import P.O."
        End If
    Else
        MsgBox "No records found in file!", vbExclamation + vbOKOnly, "Import P.O."
    End If
    Exit Sub
LocalErr:
    Screen.MousePointer = 0
    If Err.Number = 32755 Then Exit Sub
    If Err.Number = 3125 Then
        MsgBox "No P.O. Sheet found in File!" & vbCrLf & "Please use POSheet Worksheet w/ Part Number and Qty.", vbCritical + vbOKOnly, "Error"
    Else
        MsgBox Err.Description & vbCrLf & "Please contact System Administrator", vbCritical + vbOKOnly, "Error"
    End If
End Sub

Private Sub Form_Load()
    Call cmdOpen_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set rs = Nothing
    Set db = Nothing
End Sub
