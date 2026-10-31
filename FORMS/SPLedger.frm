VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Begin VB.Form SPLedger 
   Caption         =   "Customer Name"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7545
   Icon            =   "SPLedger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid flxLedger 
      Height          =   2175
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColorBkg    =   16777215
      ScrollBars      =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "^ DATE|^REF. NO.|>DEBIT |>CREDIT |>BALANCE"
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Tag             =   "XYF"
      Top             =   0
      Width           =   7335
      Begin VB.TextBox txtBal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4920
         TabIndex        =   12
         Tag             =   "XYF"
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtDue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4920
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   525
         Width           =   2295
      End
      Begin VB.TextBox txtTerms 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtInvAmt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   810
         Width           =   1455
      End
      Begin VB.TextBox txtInvDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   525
         Width           =   1455
      End
      Begin VB.TextBox txtInvNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Tag             =   "XYF"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "OUTSTANDING BAL."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   3000
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   870
         Width           =   1815
      End
      Begin VB.Label lblDue 
         Alignment       =   1  'Right Justify
         Caption         =   "DUE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   555
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "TERMS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3480
         TabIndex        =   4
         Tag             =   "XYF"
         Top             =   255
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "INV. AMOUNT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Tag             =   "XYF"
         Top             =   870
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "INV. DATE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Tag             =   "XYF"
         Top             =   555
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "INVOICE NO."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Tag             =   "XYF"
         Top             =   255
         Width           =   1335
      End
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   5880
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   7665
      m_MDI_Height    =   3915
   End
End
Attribute VB_Name = "SPLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mySrc As Form

Private Sub Form_Activate()
    If mySrc.Caption = "Official Receipt" Then
        Me.Move mySrc.Left, mySrc.Height - (mySrc.Height \ 4)
    End If
End Sub


Private Sub Form_Load()
    Dim rsLedger As New ADODB.Recordset
    Dim Qry As String
    Dim i As Integer
    
    If mySrc.Caption = "Official Receipt" Then
        txtInvNo.Text = mySrc.txtRefNo.Text
        Me.Caption = mySrc.cboAcctCust.Text
        With mySrc.flxAcct
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) = txtInvNo.Text Then
                    txtInvDate.Text = .TextMatrix(i, 12)
                    txtInvAmt.Text = Format$(.TextMatrix(i, 13), "#,##0.00")
                    txtTerms.Text = .TextMatrix(i, 3) & IIf(Val(.TextMatrix(i, 3)) > 1, " Payments ", " Payment ") & LCase(.TextMatrix(i, 2))
                    lblDue.Caption = .TextMatrix(i, 2) & " DUE"
                    txtDue.Text = Format$(.TextMatrix(i, 5), "#,##0.00")
                    txtBal.Text = Format$(.TextMatrix(i, 8), "#,##0.00")
                    Exit For
                End If
            Next i
        End With
        Qry = "SELECT LEDGER.REFDATE, LEDGER.REFNO, LEDGER.DEBIT, LEDGER.CREDIT, LEDGER.OSBAL From LEDGER " & _
              "Where (((LEDGER.SPINVNO) = '" & mySrc.txtRefNo.Text & "') And ((LEDGER.CUSTOMERID) = " & mySrc.txtCustomerID.Text & ")) ORDER BY LEDGER.SEQNO;"
        With rsLedger
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                flxLedger.Rows = 1
                Do While Not .EOF
                    flxLedger.AddItem !RefDate & vbTab & !RefNo & vbTab & IIf(!Debit <> 0, Format$(!Debit, "#,##0.00"), "") & vbTab & IIf(!Credit <> 0, Format$(!Credit, "#,##0.00"), "") & vbTab & Format$(!OSbal, "#,##0.00")
                    .MoveNext
                Loop
            End If
            .Close
        End With
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set mySrc = Nothing
End Sub

Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer
    
    With flxLedger
        For i = 0 To 4
            .ColWidth(i) = (.Width - 100) \ 5
        Next i
    End With
End Sub
