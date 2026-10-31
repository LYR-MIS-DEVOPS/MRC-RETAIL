VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form TranSummary 
   Caption         =   "Spare Parts Transfer Summary"
   ClientHeight    =   8535
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11925
   ControlBox      =   0   'False
   Icon            =   "TranSummary2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8535
   ScaleWidth      =   11925
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSConfirmTab 
      Height          =   1575
      Left            =   4200
      TabIndex        =   16
      Tag             =   "XYF"
      Top             =   3000
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   2778
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Confirmation"
      TabPicture(0)   =   "TranSummary2.frx":17D2A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdCancel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdOK"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtConfirmDate"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin MSMask.MaskEdBox txtConfirmDate 
         Height          =   375
         Left            =   2120
         TabIndex        =   17
         Tag             =   "XYF"
         Top             =   420
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin lvButton.lvButtons_H cmdOK 
         Height          =   375
         Left            =   840
         TabIndex        =   18
         Tag             =   "XYF"
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "&CONFIRM"
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
         Image           =   "TranSummary2.frx":17D46
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   375
         Left            =   2040
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "&CANCEL"
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
         Image           =   "TranSummary2.frx":182E0
         cBack           =   -2147483633
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Change Confirmation Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Tag             =   "XYF"
         Top             =   480
         Width           =   1935
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxREC 
      Height          =   7095
      Left            =   120
      TabIndex        =   15
      Tag             =   "XYF"
      Top             =   1350
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12515
      _Version        =   393216
      RowHeightMin    =   285
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Tag             =   "XY"
      Top             =   720
      Width           =   11655
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   9600
         TabIndex        =   4
         Tag             =   "XYF"
         Top             =   195
         Width           =   1935
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   375
         Left            =   5400
         TabIndex        =   3
         ToolTipText     =   "Execute"
         Top             =   150
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
         Image           =   "TranSummary2.frx":1887A
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   8280
         TabIndex        =   5
         Tag             =   "XY"
         Top             =   150
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Find Next"
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
         Image           =   "TranSummary2.frx":18E14
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dTo 
         Height          =   315
         Left            =   3720
         TabIndex        =   6
         Top             =   180
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   79822849
         CurrentDate     =   39322
      End
      Begin MSComCtl2.DTPicker dFrom 
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   180
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   79822849
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   375
         Left            =   3240
         TabIndex        =   13
         ToolTipText     =   "Execute"
         Top             =   150
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
         Image           =   "TranSummary2.frx":193AE
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdPrev 
         Height          =   375
         Left            =   5880
         TabIndex        =   14
         ToolTipText     =   "Execute"
         Top             =   150
         Visible         =   0   'False
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
         Image           =   "TranSummary2.frx":19948
         cBack           =   -2147483633
      End
      Begin VB.Label Label1 
         Caption         =   "Date Covered:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   210
         Width           =   1335
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&Show Form"
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
      Image           =   "TranSummary2.frx":19EE2
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   9
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&Refresh "
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
      Image           =   "TranSummary2.frx":1A47C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3600
      TabIndex        =   10
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&Print Summary"
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
      Image           =   "TranSummary2.frx":1AA16
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5400
      TabIndex        =   11
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "E&xit Report"
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
      Image           =   "TranSummary2.frx":1AFB0
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   360
      Top             =   8640
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   12165
      m_MDI_Height    =   9105
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Spare Parts Transfer Summary"
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
      Left            =   7200
      TabIndex        =   0
      Top             =   440
      Width           =   4500
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Spare Parts Transfer Summary"
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
      Left            =   7215
      TabIndex        =   1
      Top             =   480
      Width           =   4500
   End
   Begin VB.Menu mnuTrans 
      Caption         =   "SP Transfer"
      Visible         =   0   'False
      Begin VB.Menu mnuConfirm 
         Caption         =   "Confirm Delivery"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel Transfer"
      End
   End
End
Attribute VB_Name = "TranSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSPtrans As New ADODB.Recordset
Dim ActRow As Integer, fRow As Integer, fCol As Integer
Dim MyAccess As String
Private Sub cmdCancel_Click()
    SSConfirmTab.Visible = False
End Sub

Private Sub cmdOK_Click()
    Dim rsTransfer As New ADODB.Recordset
    Dim Ans As Integer, i As Integer, myDateFinal As String, STFID As Long
    Dim Qry As String, CanRemarks As String
        
    Ans = MsgBox("Do you want to Confirm ths Transfer?", vbQuestion + vbYesNo, "Validation")
    If Ans = vbNo Then Exit Sub
    
    If flxREC.TextMatrix(ActRow, 2) <> "" Then
        MyRef = "TR#" & Val(flxREC.TextMatrix(ActRow, 4))
        'Qry = "SELECT SPITR.SPAREPARTID, SPITR.OUT From SPITR WHERE (((SPITR.REFERENCE)='" & MyRef & "'));"
        'mySQL
        Qry = "SELECT stf.TRANSFERID,stf.DELIVEREDDATE From stocktransfer as stf WHERE stf.TRANSFERID='" & Val(flxREC.TextMatrix(ActRow, 10)) & "';"
        
        With rsTransfer
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    STFID = !TRANSFERID
                    If flxREC.TextMatrix(ActRow, 3) = "" Then
                        LYRcon.Execute "UPDATE STOCKTRANSFER SET STOCKTRANSFER.DELIVEREDDATE = '" & Format(txtConfirmDate.Text, "yyyy-mm-dd") & "', STOCKTRANSFER.UPDATEDBY = " & UserID & ", STOCKTRANSFER.LASTUPDATE = '" & Format(Now, "yyyy-mm-dd") & "' " & _
                            "WHERE STOCKTRANSFER.TRANSFERID=" & STFID & ";"
                    Else
                        MsgBox "Spareparts Transfer Ref.# " & flxREC.TextMatrix(ActRow, 4) & " is already Confirm !", vbInformation + vbOKOnly, "CONFIRM"
                        .Close
                        SSConfirmTab.Visible = False
                        Exit Sub
                     End If
                    .MoveNext
                Loop
            End If
            .Close
        End With
    Else
        MsgBox "Spareparts Transfer Ref.# " & flxREC.TextMatrix(ActRow, 4) & " is not yet Final !", vbInformation + vbOKOnly, "CONFIRM"
        Exit Sub
    End If
    
    MsgBox "Spareparts Transfer Ref.# " & flxREC.TextMatrix(ActRow, 4) & " was Confirm!", vbInformation + vbOKOnly, "CONFIRM"
    SSConfirmTab.Visible = False
    Call cmdRefresh_Click
End Sub

Private Sub mnuconfirm_Click()
    SSConfirmTab.Visible = True
    txtConfirmDate.Text = Format(Now, "mm/dd/yyyy")
    With SSConfirmTab
        .Top = (Me.Height - .Height) \ 2
        .Left = (Me.Width - .Width) \ 2
        .Visible = True
        .ZOrder
    End With
End Sub

Private Sub RefreshRecord(dCode As Integer)
    Dim Qry As String, i As Integer
    
    If dCode = 1 Then
        If showcost Then
            'Qry = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.TRANSFERID, STOCKTRANSFER.TOTALCOST, STOCKTRANSFER.RECBRANCH, STOCKTRANSFER.ADDRESS, STOCKTRANSFER.PONO, STOCKTRANSFER.STATUS, STOCKTRANSFER.REMARKS " & _
                  "From STOCKTRANSFER Where (((STOCKTRANSFER.TRANSDATE) = #" & dFrom.Value & "#)) ORDER BY STOCKTRANSFER.TRANSFERID;"
            'mySQL
            Qry = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.DATEFINAL, STOCKTRANSFER.DELIVEREDDATE, STOCKTRANSFER.TRANSNO, (STOCKTRANSFER.TOTALCOST)*1.12 as TOTALCOSTING, STOCKTRANSFER.RECBRANCH, STOCKTRANSFER.ADDRESS, STOCKTRANSFER.PONO, STOCKTRANSFER.STATUS, STOCKTRANSFER.REMARKS, STOCKTRANSFER.TRANSFERID " & _
                  "From STOCKTRANSFER Where (((STOCKTRANSFER.TRANSDATE) = '" & Format(dFrom.Value, "yyyy-mm-dd") & "')) ORDER BY STOCKTRANSFER.TRANSFERID desc;"
            
        Else
            'Qry = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.TRANSFERID, STOCKTRANSFER.TOTALCOST, STOCKTRANSFER.RECBRANCH, STOCKTRANSFER.ADDRESS, STOCKTRANSFER.PONO, STOCKTRANSFER.STATUS, STOCKTRANSFER.REMARKS " & _
                  "From STOCKTRANSFER Where (((STOCKTRANSFER.TRANSDATE) = #" & dFrom.Value & "#)) ORDER BY STOCKTRANSFER.TRANSFERID;"
            'mySQL
            Qry = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.DATEFINAL, STOCKTRANSFER.DELIVEREDDATE, STOCKTRANSFER.TRANSNO, (STOCKTRANSFER.TOTALCOST)*1.12 as TOTALCOSTING, STOCKTRANSFER.RECBRANCH, STOCKTRANSFER.ADDRESS, STOCKTRANSFER.PONO, STOCKTRANSFER.STATUS, STOCKTRANSFER.REMARKS, STOCKTRANSFER.TRANSFERID " & _
                  "From STOCKTRANSFER Where (((STOCKTRANSFER.TRANSDATE) = '" & Format(dFrom.Value, "yyyy-mm-dd") & "')) ORDER BY STOCKTRANSFER.TRANSFERID desc;"
            
        End If
    Else
        If showcost Then
            'Qry = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.TRANSFERID, STOCKTRANSFER.TOTALCOST, STOCKTRANSFER.RECBRANCH, STOCKTRANSFER.ADDRESS, STOCKTRANSFER.PONO, STOCKTRANSFER.STATUS, STOCKTRANSFER.REMARKS " & _
                  "From STOCKTRANSFER Where STOCKTRANSFER.TRANSDATE >= #" & dFrom.Value & "# And STOCKTRANSFER.TRANSDATE <= #" & dTo.Value & "#  ORDER BY STOCKTRANSFER.TRANSFERID;"
            'myMYSQL
            Qry = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.DATEFINAL, STOCKTRANSFER.DELIVEREDDATE, STOCKTRANSFER.TRANSNO, (STOCKTRANSFER.TOTALCOST)*1.12 as TOTALCOSTING, STOCKTRANSFER.RECBRANCH, STOCKTRANSFER.ADDRESS, STOCKTRANSFER.PONO, STOCKTRANSFER.STATUS, STOCKTRANSFER.REMARKS, STOCKTRANSFER.TRANSFERID " & _
                  "From STOCKTRANSFER Where STOCKTRANSFER.TRANSDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And STOCKTRANSFER.TRANSDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "'  ORDER BY STOCKTRANSFER.TRANSFERID desc;"
            
        Else
            'Qry = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.TRANSFERID, STOCKTRANSFER.RECBRANCH, STOCKTRANSFER.ADDRESS, STOCKTRANSFER.PONO, STOCKTRANSFER.STATUS, STOCKTRANSFER.REMARKS " & _
            "From STOCKTRANSFER Where STOCKTRANSFER.TRANSDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & " And STOCKTRANSFER.TRANSDATE <= #" & dTo.Value & "#  ORDER BY STOCKTRANSFER.TRANSFERID;"
            'mySQL
            Qry = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.DATEFINAL, STOCKTRANSFER.DELIVEREDDATE, STOCKTRANSFER.TRANSNO, STOCKTRANSFER.RECBRANCH, STOCKTRANSFER.ADDRESS, STOCKTRANSFER.PONO, STOCKTRANSFER.STATUS, STOCKTRANSFER.REMARKS, STOCKTRANSFER.TRANSFERID " & _
            "From STOCKTRANSFER Where STOCKTRANSFER.TRANSDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And STOCKTRANSFER.TRANSDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "'  ORDER BY STOCKTRANSFER.TRANSFERID desc;"
                
        End If
    End If
    If rsSPtrans.State = adStateOpen Then rsSPtrans.Close
    rsSPtrans.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsSPtrans.RecordCount = 0 Then
        flxREC.Clear
        flxREC.Rows = 2
        flxREC.Cols = 9
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxREC.DataSource = rsSPtrans
    flxREC.Visible = True
    Call RXResizer1_ResizeComplete
End Sub

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxREC
        For i = fRow To .Rows - 1
            For j = fCol To 8
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 8:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 11, i + 1, i)
                    fCol = IIf(j = 11, 1, j + 1)
                    Exit Sub
                End If
            Next j
            fCol = 1
        Next i
    End With
    MsgBox "No match found!", vbInformation + vbOKOnly, "Find"
    fRow = 1
    fCol = 1
End Sub

Private Sub cmdNext_Click()
    cmdNext.Tag = "2"
    cmdNext.Visible = False
    cmdPrev.Left = cmdNext.Left
    cmdPrev.Top = cmdNext.Top
    cmdPrev.Visible = True
    dTo.Visible = True
    cmdRefresh.Visible = True
    dTo.Value = dFrom.Value
End Sub

Private Sub cmdPrev_Click()
    cmdNext.Tag = "1"
    cmdNext.Visible = True
    cmdPrev.Visible = False
    dTo.Visible = False
    cmdRefresh.Visible = False
    Call RefreshRecord(1)
End Sub

Private Sub cmdRefresh_Click()
    DateStart = dFrom.Value
    DateEnd = dTo.Value
    Call RefreshRecord(2)
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim Qry As String, Ans As Integer
    
    Select Case Index
        Case 0  'Show Packing No
            
            If flxREC.Rows = 2 Then ActRow = 1
'            If rsSPtrans.RecordCount > 0 And ActRow <> 0 Then
'                ParTrans.Show
'                Call ParTrans.ShowRecord(Val(flxREC.TextMatrix(ActRow, 3)))
'                Unload Me
'            Else
'                ParTrans.Show
'                Unload Me
'            End If
            
            If flxREC.TextMatrix(flxREC.Rows - 1, 1) = "" Then
                MsgBox "No selected  data to show"
            Else
                If showcost Then
                    Call ParTrans.ShowRecord(flxREC.TextMatrix(ActRow, 4), Val(flxREC.TextMatrix(ActRow, 11)))
                Else
                    Call ParTrans.ShowRecord(flxREC.TextMatrix(ActRow, 4), Val(flxREC.TextMatrix(ActRow, 10)))
                End If
                Unload Me
            End If
            
        Case 1  'Refresh
            Call RefreshRecord(Val(cmdNext.Tag))
        Case 2  'Print
            If showcost Then
                Ans = MsgBox("Do you want to include PENDING Status?", vbQuestion + vbYesNo, "Summary")
                If Ans = vbYes Then
                    If dCode = 1 Then
                        'PRNmc.QryPrn = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.TRANSFERID, STOCKTRANSFER.RECBRANCH, STOCKTRANSFER.TOTALCOST, STOCKTRANSFER.STATUS " & _
                              "From STOCKTRANSFER Where (((STOCKTRANSFER.TRANSDATE) = #" & dFrom.Value & "#)) ORDER BY STOCKTRANSFER.TRANSFERID;"
                        'mySQL
                        PRNmc.QryPrn = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.DATEFINAL, STOCKTRANSFER.DELIVEREDDATE, STOCKTRANSFER.TRANSNO, (STOCKTRANSFER.TOTALCOST)*1.12 as TOTALCOSTING, STOCKTRANSFER.RECBRANCH, STOCKTRANSFER.ADDRESS, STOCKTRANSFER.PONO, STOCKTRANSFER.STATUS, STOCKTRANSFER.REMARKS, STOCKTRANSFER.TRANSFERID " & _
                                        "From STOCKTRANSFER Where STOCKTRANSFER.TRANSDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' ORDER BY STOCKTRANSFER.TRANSFERID;"

                        'PRNmc.QryPrn = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.TRANSNO, STOCKTRANSFER.RECBRANCH, (STOCKTRANSFER.TOTALCOST)*1.12 as TOTALCOSTING, STOCKTRANSFER.STATUS " & _
                        '      "From STOCKTRANSFER Where (((STOCKTRANSFER.TRANSDATE) = '" & Format(dFrom.Value, "yyyy-mm-dd") & "')) ORDER BY STOCKTRANSFER.TRANSFERID;"
                        
                    Else
                        'PRNmc.QryPrn = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.TRANSFERID, STOCKTRANSFER.RECBRANCH, STOCKTRANSFER.TOTALCOST, STOCKTRANSFER.STATUS " & _
                              "From STOCKTRANSFER Where STOCKTRANSFER.TRANSDATE >= #" & dFrom.Value & "# And STOCKTRANSFER.TRANSDATE <= #" & dTo.Value & "#  ORDER BY STOCKTRANSFER.TRANSFERID;"
                        'mySQL
                        PRNmc.QryPrn = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.DATEFINAL, STOCKTRANSFER.DELIVEREDDATE, STOCKTRANSFER.TRANSNO, (STOCKTRANSFER.TOTALCOST)*1.12 as TOTALCOSTING, STOCKTRANSFER.RECBRANCH, STOCKTRANSFER.ADDRESS, STOCKTRANSFER.PONO, STOCKTRANSFER.STATUS, STOCKTRANSFER.REMARKS, STOCKTRANSFER.TRANSFERID " & _
                                        "From STOCKTRANSFER Where STOCKTRANSFER.TRANSDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And STOCKTRANSFER.TRANSDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' ORDER BY STOCKTRANSFER.TRANSFERID;"

                        'PRNmc.QryPrn = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.TRANSNO, STOCKTRANSFER.RECBRANCH, (STOCKTRANSFER.TOTALCOST)*1.12 as TOTALCOSTING, STOCKTRANSFER.STATUS " & _
                        '      "From STOCKTRANSFER Where STOCKTRANSFER.TRANSDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And STOCKTRANSFER.TRANSDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "'  ORDER BY STOCKTRANSFER.TRANSFERID;"
                        
                    End If
                Else
                    If dCode = 1 Then
                        'PRNmc.QryPrn = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.TRANSFERID, STOCKTRANSFER.RECBRANCH, STOCKTRANSFER.TOTALCOST, STOCKTRANSFER.STATUS " & _
                              "From STOCKTRANSFER Where (((STOCKTRANSFER.TRANSDATE) = #" & dFrom.Value & "#) And ((STOCKTRANSFER.STATUS) = 'FINAL')) ORDER BY STOCKTRANSFER.TRANSFERID;"
                        'mySQL
                        PRNmc.QryPrn = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.DATEFINAL, STOCKTRANSFER.DELIVEREDDATE, STOCKTRANSFER.TRANSNO, (STOCKTRANSFER.TOTALCOST)*1.12 as TOTALCOSTING, STOCKTRANSFER.RECBRANCH, STOCKTRANSFER.ADDRESS, STOCKTRANSFER.PONO, STOCKTRANSFER.STATUS, STOCKTRANSFER.REMARKS, STOCKTRANSFER.TRANSFERID " & _
                                        "From STOCKTRANSFER Where STOCKTRANSFER.TRANSDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And STOCKTRANSFER.STATUS = 'FINAL' ORDER BY STOCKTRANSFER.TRANSFERID;"

                        'PRNmc.QryPrn = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.TRANSNO, STOCKTRANSFER.RECBRANCH, (STOCKTRANSFER.TOTALCOST)*1.12 as TOTALCOSTING, STOCKTRANSFER.STATUS " & _
                        '      "From STOCKTRANSFER Where (((STOCKTRANSFER.TRANSDATE) = '" & Format(dFrom.Value, "yyyy-mm-dd") & "') And ((STOCKTRANSFER.STATUS) = 'FINAL')) ORDER BY STOCKTRANSFER.TRANSFERID;"
                        
                    Else
                        'PRNmc.QryPrn = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.TRANSFERID, STOCKTRANSFER.RECBRANCH, STOCKTRANSFER.TOTALCOST, STOCKTRANSFER.STATUS " & _
                              "From STOCKTRANSFER Where STOCKTRANSFER.TRANSDATE >= #" & dFrom.Value & "# And STOCKTRANSFER.TRANSDATE <= #" & dTo.Value & "# And STOCKTRANSFER.STATUS = 'FINAL' ORDER BY STOCKTRANSFER.TRANSFERID;"
                        'mySQL
                        PRNmc.QryPrn = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.DATEFINAL, STOCKTRANSFER.DELIVEREDDATE, STOCKTRANSFER.TRANSNO, (STOCKTRANSFER.TOTALCOST)*1.12 as TOTALCOSTING, STOCKTRANSFER.RECBRANCH, STOCKTRANSFER.ADDRESS, STOCKTRANSFER.PONO, STOCKTRANSFER.STATUS, STOCKTRANSFER.REMARKS, STOCKTRANSFER.TRANSFERID " & _
                                        "From STOCKTRANSFER Where STOCKTRANSFER.TRANSDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And STOCKTRANSFER.TRANSDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' And STOCKTRANSFER.STATUS = 'FINAL' ORDER BY STOCKTRANSFER.TRANSFERID;"

                        'PRNmc.QryPrn = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.TRANSNO, STOCKTRANSFER.RECBRANCH, (STOCKTRANSFER.TOTALCOST)*1.12 as TOTALCOSTING, STOCKTRANSFER.STATUS " & _
                        '      "From STOCKTRANSFER Where STOCKTRANSFER.TRANSDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And STOCKTRANSFER.TRANSDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' And STOCKTRANSFER.STATUS = 'FINAL' ORDER BY STOCKTRANSFER.TRANSFERID;"
                        
                    End If
                End If
            Else
                Ans = MsgBox("Do you want to include PENDING Status?", vbQuestion + vbYesNo, "Summary")
                If Ans = vbYes Then
                    If dCode = 1 Then
                        'PRNmc.QryPrn = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.TRANSFERID, STOCKTRANSFER.RECBRANCH, STOCKTRANSFER.TOTALCOST, STOCKTRANSFER.STATUS " & _
                              "From STOCKTRANSFER Where (((STOCKTRANSFER.TRANSDATE) = #" & dFrom.Value & "#)) ORDER BY STOCKTRANSFER.TRANSFERID;"
                        'mySQL
                        PRNmc.QryPrn = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.DATEFINAL, STOCKTRANSFER.DELIVEREDDATE, STOCKTRANSFER.TRANSNO, (STOCKTRANSFER.TOTALCOST)*1.12 as TOTALCOSTING, STOCKTRANSFER.RECBRANCH, STOCKTRANSFER.ADDRESS, STOCKTRANSFER.PONO, STOCKTRANSFER.STATUS, STOCKTRANSFER.REMARKS, STOCKTRANSFER.TRANSFERID " & _
                                        "From STOCKTRANSFER Where STOCKTRANSFER.TRANSDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' ORDER BY STOCKTRANSFER.TRANSFERID;"
                        'PRNmc.QryPrn = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.TRANSNO, STOCKTRANSFER.RECBRANCH, (STOCKTRANSFER.TOTALCOST)*1.12 as TOTALCOSTING, STOCKTRANSFER.STATUS " & _
                        '      "From STOCKTRANSFER Where (((STOCKTRANSFER.TRANSDATE) = '" & Format(dFrom.Value, "yyyy-mm-dd") & "')) ORDER BY STOCKTRANSFER.TRANSFERID;"
                        
                    Else
                        'PRNmc.QryPrn = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.TRANSFERID, STOCKTRANSFER.RECBRANCH, STOCKTRANSFER.TOTALCOST, STOCKTRANSFER.STATUS " & _
                              "From STOCKTRANSFER Where STOCKTRANSFER.TRANSDATE >= #" & dFrom.Value & "# And STOCKTRANSFER.TRANSDATE <= #" & dTo.Value & "#  ORDER BY STOCKTRANSFER.TRANSFERID;"
                        'mySQL
                        PRNmc.QryPrn = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.DATEFINAL, STOCKTRANSFER.DELIVEREDDATE, STOCKTRANSFER.TRANSNO, (STOCKTRANSFER.TOTALCOST)*1.12 as TOTALCOSTING, STOCKTRANSFER.RECBRANCH, STOCKTRANSFER.ADDRESS, STOCKTRANSFER.PONO, STOCKTRANSFER.STATUS, STOCKTRANSFER.REMARKS, STOCKTRANSFER.TRANSFERID " & _
                                        "From STOCKTRANSFER Where STOCKTRANSFER.TRANSDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And STOCKTRANSFER.TRANSDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "'  ORDER BY STOCKTRANSFER.TRANSFERID;"

                        'PRNmc.QryPrn = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.TRANSNO, STOCKTRANSFER.RECBRANCH, (STOCKTRANSFER.TOTALCOST)*1.12 as TOTALCOSTING, STOCKTRANSFER.STATUS " & _
                        '      "From STOCKTRANSFER Where STOCKTRANSFER.TRANSDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And STOCKTRANSFER.TRANSDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "'  ORDER BY STOCKTRANSFER.TRANSFERID;"
                        
                    End If
                Else
                    If dCode = 1 Then
                        'PRNmc.QryPrn = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.TRANSFERID, STOCKTRANSFER.RECBRANCH, STOCKTRANSFER.TOTALCOST, STOCKTRANSFER.STATUS " & _
                              "From STOCKTRANSFER Where (((STOCKTRANSFER.TRANSDATE) = #" & dFrom.Value & "#) And ((STOCKTRANSFER.STATUS) = 'FINAL')) ORDER BY STOCKTRANSFER.TRANSFERID;"
                        'mySQL
                        PRNmc.QryPrn = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.DATEFINAL, STOCKTRANSFER.DELIVEREDDATE, STOCKTRANSFER.TRANSNO, (STOCKTRANSFER.TOTALCOST)*1.12 as TOTALCOSTING, STOCKTRANSFER.RECBRANCH, STOCKTRANSFER.ADDRESS, STOCKTRANSFER.PONO, STOCKTRANSFER.STATUS, STOCKTRANSFER.REMARKS, STOCKTRANSFER.TRANSFERID " & _
                                        "From STOCKTRANSFER Where STOCKTRANSFER.TRANSDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And ((STOCKTRANSFER.STATUS) = 'FINAL')) ORDER BY STOCKTRANSFER.TRANSFERID;"

                        'PRNmc.QryPrn = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.TRANSNO, STOCKTRANSFER.RECBRANCH, (STOCKTRANSFER.TOTALCOST)*1.12 as TOTALCOSTING, STOCKTRANSFER.STATUS " & _
                        '      "From STOCKTRANSFER Where (((STOCKTRANSFER.TRANSDATE) = '" & Format(dFrom.Value, "yyyy-mm-dd") & "') And ((STOCKTRANSFER.STATUS) = 'FINAL')) ORDER BY STOCKTRANSFER.TRANSFERID;"
                        
                    Else
                        'PRNmc.QryPrn = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.TRANSFERID, STOCKTRANSFER.RECBRANCH, STOCKTRANSFER.TOTALCOST, STOCKTRANSFER.STATUS " & _
                              "From STOCKTRANSFER Where STOCKTRANSFER.TRANSDATE >= #" & dFrom.Value & "# And STOCKTRANSFER.TRANSDATE <= #" & dTo.Value & "# And STOCKTRANSFER.STATUS = 'FINAL' ORDER BY STOCKTRANSFER.TRANSFERID;"
                        'mySQL
                        PRNmc.QryPrn = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.DATEFINAL, STOCKTRANSFER.DELIVEREDDATE, STOCKTRANSFER.TRANSNO, (STOCKTRANSFER.TOTALCOST)*1.12 as TOTALCOSTING, STOCKTRANSFER.RECBRANCH, STOCKTRANSFER.ADDRESS, STOCKTRANSFER.PONO, STOCKTRANSFER.STATUS, STOCKTRANSFER.REMARKS, STOCKTRANSFER.TRANSFERID " & _
                                        "From STOCKTRANSFER Where STOCKTRANSFER.TRANSDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And STOCKTRANSFER.TRANSDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' And STOCKTRANSFER.STATUS = 'FINAL' ORDER BY STOCKTRANSFER.TRANSFERID;"

                        'PRNmc.QryPrn = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.TRANSNO, STOCKTRANSFER.RECBRANCH, (STOCKTRANSFER.TOTALCOST)*1.12 as TOTALCOSTING, STOCKTRANSFER.STATUS " & _
                        '      "From STOCKTRANSFER Where STOCKTRANSFER.TRANSDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And STOCKTRANSFER.TRANSDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' And STOCKTRANSFER.STATUS = 'FINAL' ORDER BY STOCKTRANSFER.TRANSFERID;"
                        
                    End If
                End If
            End If
                PRNmc.MySource = "SUMMARY"
                Qry = "SELECT BRANCH.BRANCHNAME From BRANCH WHERE (((BRANCH.SYSTEMBRANCH)=True));"
                PRNmc.ParamValue1 = GetFieldValue(Qry)
                Qry = "SELECT BRANCH.ADDRESS From BRANCH WHERE (((BRANCH.SYSTEMBRANCH)=True));"
                PRNmc.ParamValue2 = GetFieldValue(Qry)
                PRNmc.ParamValue3 = "For the period " & dFrom.Value & " to " & dTo.Value
                PRNmc.Show
        Case 3  'Exit
            
            'LogIn.Show vbModal
            'MotoRep.tbMain.Visible = True
            Unload Me
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
                        
    End Select
End Sub

Private Sub dFrom_CloseUp()
    If cmdNext.Tag = "1" Then
        Call RefreshRecord(1)
    End If
End Sub

Private Sub dTo_CloseUp()
    If dTo.Value < dFrom.Value Then
        MsgBox "Please choose a date not less than " & dFrom.Value, vbExclamation + vbOKOnly, "Validation Error"
        dTo.Value = dFrom.Value
    End If
End Sub

Private Sub flxREC_Click()
    Call flxrec_EnterCell
End Sub

Private Sub flxrec_EnterCell()
    ActRow = flxREC.Row
End Sub

Private Sub flxREC_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim cOlSW As Integer
    If showcost = True Then
        cOlSW = 9
    Else
        cOlSW = 8
    End If
    
    If flxREC.TextMatrix(ActRow, cOlSW) = "FINAL" Then
        
        If Button = vbRightButton Then
            Call PopupMenu(mnuTrans)
        End If
    End If
End Sub

Private Sub Form_Load()
    MotoRep.sBar.Panels(6).Text = "TRANSFER SUMMARY"
'    MotoRep.toolBarButton.Visible = False
'    MotoRep.imgBG.Visible = False
    If DateStart = "" Then
        If Day(Date) <= 7 Then
            dFrom.Value = DateAdd("m", -1, DateSerial(Year(Date), Month(Date), 1))
        Else
            dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        End If
        dTo.Value = Date
    Else
        dFrom.Value = DateValue(DateStart)
        dTo.Value = DateValue(DateEnd)
    End If
    cmdNext.Tag = "2"
    Call RefreshRecord(2)
    'MotoRep.sBar.Panels(6).Text = "Spare Parts Receiving Summary"
    'MotoRep.tbMain.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'MotoRep.sBar.Panels(6).Text = ""
    If rsSPtrans.State = adStateOpen Then rsSPtrans.Close
    MotoRep.sBar.Panels(6).Text = ""
'    MotoRep.toolBarButton.Visible = True
'    'MotoRep.imgBG.Visible = True
End Sub

Private Sub mnuCancel_Click()
    MsgBox "Not applicable to this Transaction", vbInformation + vbOKOnly, "Validation"
'    Dim rsTransfer As New ADODB.Recordset
'    Dim Ans As Integer, i As Integer, MyQty As Integer, SPid As Long
'    Dim Qry As String, CanRemarks As String
'
'    If InStr(SysBranch, "WAREHOUSE") = 0 Then
'        MsgBox "Not Yet Available", vbInformation, "Cancel"
'        Exit Sub
'    End If
'    Qry = "SELECT employees.UserType From employees Where employees.EmpId = '" & UserID & "';"
'    MyAccess = GetFieldValue(Qry)
'                If MyAccess <> "Admin" Then
'                    MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
'                    Exit Sub
'                End If
'
'    Ans = MsgBox("Do you want to Cancel Transfer?", vbQuestion + vbYesNo, "Cancel")
'    If Ans = vbNo Then Exit Sub
'
'    MyRef = "TR#" & Val(flxREC.TextMatrix(ActRow, 4))
'    'Qry = "SELECT SPITR.SPAREPARTID, SPITR.OUT From SPITR WHERE (((SPITR.REFERENCE)='" & MyRef & "'));"
'    'mySQL
'    Qry = "SELECT SPITR.SPAREPARTID, SPITR.OUT From SPITR WHERE SPITR.REFERENCE='" & MyRef & "';"
'
'    With rsTransfer
'        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
'        If .RecordCount > 0 Then
'            .MoveFirst
'            Do While Not .EOF
'                SPid = !SparepartID
'                MyQty = !OUT
'                'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = [SPAREPARTS].[QTYONHAND]+" & MyQty & ", SPAREPARTS.UPDATEDBY = " & UserID & ", SPAREPARTS.LASTUPDATE = #" & Now & "# " & _
'                               "WHERE (((SPAREPARTS.SPAREPARTID)=" & SPid & "));"
'
'                'mySQL
'                LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = SPAREPARTS.QTYONHAND+" & MyQty & ", SPAREPARTS.UPDATEDBY = " & UserID & ", SPAREPARTS.LASTUPDATE = '" & Format(Now, "yyyy-mm-dd") & "' " & _
'                    "WHERE SPAREPARTS.SPAREPARTID=" & SPid & ";"
'
'               .MoveNext
'            Loop
'            LYRcon.Execute "Update stocktransdtl SET QTY = 0 Where stocktransdtl.TRANSFERID = " & Val(flxREC.TextMatrix(ActRow, 10)) & ";"
'
'        End If
'        .Close
'    End With
'    'Access
'    'LYRcon.Execute "UPDATE STOCKTRANSFER SET STOCKTRANSFER.STATUS = 'CANCELLED' WHERE (((STOCKTRANSFER.TRANSFERID)=" & Val(flxREC.TextMatrix(ActRow, 2)) & "));"
'    'LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSTYPE = 'CANCELLED' WHERE (((SPITR.REFERENCE)='" & MyRef & "'));"
'
'   'mySQL
'    LYRcon.Execute "UPDATE STOCKTRANSFER SET STOCKTRANSFER.STATUS = 'CANCELLED' WHERE STOCKTRANSFER.TRANSFERID=" & Val(flxREC.TextMatrix(ActRow, 10)) & ";"
'    LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSTYPE = 'CANCELLED' WHERE SPITR.REFERENCE='" & MyRef & "';"
'
'    With flxREC
'        .TextMatrix(ActRow, 8) = "CANCELLED"
'        .Col = 1: .Row = ActRow: .ColSel = 9: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
'        .Col = 2
'    End With
'          'CANCELLED ORDER ALLOCATION
'          CanRemarks = "Confirmation Number: CANCELLED"
'
''          LYRcon.Execute "Update spordrconfirm SET FINALQT = 2, REMARKS = '" & CanRemarks & "', SQSERVE = 0 " & _
''            "Where spordrconfirm.BRANCHNAME = '" & flxREC.TextMatrix(ActRow, 4) & "' AND spordrconfirm.CONFIRMNO = '" & MyRef & "';"
''
''          Qry = "SELECT spordrconfirm.ORDERID from spordrconfirm where spordrconfirm.CONFIRMNO = '" & MyRef & "';"
''            STOrderID = Val(GetFieldValue(Qry))
''
''          LYRcon.Execute "Update spbrchconfirmdtl SET TRANSTYPE = 'CANCELLED' Where spbrchconfirmdtl.ORDERID = " & STOrderID & ";"
''
'    MsgBox "Spareparts Transfer Ref.# " & flxREC.TextMatrix(ActRow, 4) & " CANCELLED!", vbInformation + vbOKOnly, "Cancel"
End Sub

Private Sub RXResizer1_ResizeComplete()
    lblHeading(2).Left = Frame1.Left + (Frame1.Width - lblHeading(2).Width) - 30
    lblHeading(3).Left = Frame1.Left + (Frame1.Width - lblHeading(3).Width)
    
    With flxREC
        If showcost Then
            .FormatString = "   |^REF. DATE|^FINAL DATE |^DELIVERY DATE |^REF. NO.|>TOTAL PRICE|<RECEIVING BRANCH|<ADDRESS|^P.O. NO.|<STATUS |<REMARKS"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1
            .ColWidth(2) = (.Width - 300) * 0.1
            .ColWidth(3) = (.Width - 300) * 0.1
            .ColWidth(4) = (.Width - 300) * 0.1
            .ColWidth(5) = (.Width - 300) * 0.1
            .ColWidth(6) = (.Width - 300) * 0.2
            .ColWidth(7) = (.Width - 300) * 0.2
            .ColWidth(8) = (.Width - 300) * 0.1
            .ColWidth(9) = (.Width - 300) * 0.1
            .ColWidth(10) = (.Width - 300) * 0.18
            .ColWidth(11) = 0 '(.Width - 300) * 0.18
            If rsSPtrans.RecordCount > 0 Then
                For i = 1 To .Rows - 1
                    .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                    .TextMatrix(i, 2) = Format$(.TextMatrix(i, 2), "mm/dd/yyyy")
                    .TextMatrix(i, 3) = Format$(.TextMatrix(i, 3), "mm/dd/yyyy")
                    .TextMatrix(i, 5) = Format$(.TextMatrix(i, 5), "#,##0.00")
                    'If Len(.TextMatrix(i, 3)) < 4 Then .TextMatrix(i, 3) = Format(.TextMatrix(i, 3), "0000")
                    If .TextMatrix(i, 9) = "PENDING" Then 'Not yet printed
                        .Col = 1: .Row = i: .ColSel = 10: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                        .Col = 2
                    End If
                    If InStr(.TextMatrix(i, 9), "CANCELLED") > 0 Then 'Cancelled
                        .Col = 1: .Row = i: .ColSel = 10: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
                        .Col = 2
                    End If
                Next i
            End If
        Else
            .FormatString = "   |^REF. DATE|^FINAL DATE |^DELIVERY DATE |^REF. NO.|<RECEIVING BRANCH|<ADDRESS|^P.O. NO.|<STATUS |<REMARKS"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1
            .ColWidth(2) = (.Width - 300) * 0.1
            .ColWidth(3) = (.Width - 300) * 0.1
            .ColWidth(4) = (.Width - 300) * 0.1
            .ColWidth(5) = (.Width - 300) * 0.2
            .ColWidth(6) = (.Width - 300) * 0.2
            .ColWidth(7) = (.Width - 300) * 0.1
            .ColWidth(8) = (.Width - 300) * 0.1
            .ColWidth(9) = (.Width - 300) * 0.18
            .ColWidth(10) = 0 '(.Width - 300) * 0.18
            If rsSPtrans.State = 0 Then Exit Sub
            If rsSPtrans.RecordCount > 0 Then
                For i = 1 To .Rows - 1
                    .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                    .TextMatrix(i, 2) = Format$(.TextMatrix(i, 2), "mm/dd/yyyy")
                    .TextMatrix(i, 3) = Format$(.TextMatrix(i, 3), "mm/dd/yyyy")
                    'If Len(.TextMatrix(i, 3)) < 4 Then .TextMatrix(i, 3) = Format(.TextMatrix(i, 3), "0000")
                    If .TextMatrix(i, 8) = "PENDING" Then 'Not yet printed
                        .Col = 1: .Row = i: .ColSel = 9: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                        .Col = 2
                    End If
                    If InStr(.TextMatrix(i, 8), "CANCELLED") > 0 Then 'Cancelled
                        .Col = 1: .Row = i: .ColSel = 9: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
                        .Col = 2
                    End If
                Next i
            End If
        End If
    End With
    
    'cmdFind.Height = cmdRefresh.Height
    cmdFind.Top = cmdRefresh.Top
    txtFind.Top = cmdFind.Top
End Sub

Private Sub txtFind_Change()
    fRow = 1
    fCol = 1
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub
