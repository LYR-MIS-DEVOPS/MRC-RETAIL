VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form SPWarehouseOld 
   Caption         =   "Spare Parts Warehouse"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12510
   ControlBox      =   0   'False
   Icon            =   "SPWarehouseOld.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   12510
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid flxWH 
      Height          =   6615
      Left            =   9240
      TabIndex        =   31
      Tag             =   "XYF"
      Top             =   840
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   11668
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      RowHeightMin    =   330
      BackColorBkg    =   16777215
      ScrollBars      =   2
      SelectionMode   =   1
      FormatString    =   "^ ORDER NO. |^ FRONTLINER                "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtInvNo 
      Height          =   285
      Left            =   6120
      TabIndex        =   27
      Top             =   7680
      Visible         =   0   'False
      Width           =   375
   End
   Begin lvButton.lvButtons_H cmdFind 
      Height          =   315
      Left            =   6000
      TabIndex        =   23
      Tag             =   "XYF"
      Top             =   1440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Caption         =   "Find"
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
      Image           =   "SPWarehouseOld.frx":164A
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtFind 
      Height          =   315
      Left            =   6960
      TabIndex        =   22
      Tag             =   "XYF"
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox txtSPSorderID 
      Height          =   285
      Left            =   840
      TabIndex        =   21
      Top             =   7680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   720
      Left            =   240
      TabIndex        =   16
      Tag             =   "XYF"
      Top             =   6720
      Width           =   8895
      Begin VB.TextBox txtFrnt 
         BackColor       =   &H80000003&
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
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   240
         Width           =   3255
      End
      Begin lvButton.lvButtons_H cmdFrontLine 
         Height          =   495
         Left            =   8040
         TabIndex        =   24
         Tag             =   "XYF"
         ToolTipText     =   "Frontline"
         Top             =   165
         Visible         =   0   'False
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12648447
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "SPWarehouseOld.frx":1BE4
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdPrint 
         Height          =   495
         Left            =   7440
         TabIndex        =   25
         Tag             =   "XYF"
         ToolTipText     =   "Print"
         Top             =   165
         Visible         =   0   'False
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12648447
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "SPWarehouseOld.frx":23C7
         cBack           =   -2147483633
      End
      Begin VB.Label lblRemarks 
         Height          =   375
         Left            =   4920
         TabIndex        =   26
         Tag             =   "XYF"
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label lblORno 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   1755
         TabIndex        =   20
         Tag             =   "XYF"
         Top             =   240
         Width           =   2685
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "FRONTLINER:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Tag             =   "XYF"
         Top             =   285
         Width           =   1215
      End
   End
   Begin VB.TextBox txtPrintCopy 
      Height          =   285
      Left            =   5280
      TabIndex        =   15
      Top             =   7680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtWCRid 
      Height          =   285
      Left            =   1320
      TabIndex        =   14
      Top             =   7080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdAltA 
      Caption         =   "&A"
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Tag             =   "XY"
      Top             =   10920
      Width           =   135
   End
   Begin VB.Timer tmrQtyTip 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2040
      Top             =   8040
   End
   Begin VB.TextBox txtQtyTip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   600
      TabIndex        =   12
      Text            =   "Qty. On Hand: 0"
      Top             =   8160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtAcctNo 
      Height          =   285
      Left            =   8640
      TabIndex        =   11
      Top             =   10560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid flxOrd 
      Height          =   4815
      Left            =   240
      TabIndex        =   10
      Tag             =   "XY"
      Top             =   1800
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   8493
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   -2147483633
      ScrollBars      =   0
      SelectionMode   =   1
      FormatString    =   "^     |^ QTY. |^ UNIT |<     PART NUMBER         |<    DESCRIPTION            |<LOCATION   "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   660
      Left            =   240
      TabIndex        =   3
      Tag             =   "XYF"
      Top             =   720
      Width           =   8895
      Begin VB.TextBox txtOrderno 
         Alignment       =   2  'Center
         BackColor       =   &H80000003&
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   1
         Tag             =   "XYF"
         Top             =   180
         Width           =   2055
      End
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         BackColor       =   &H80000003&
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
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   2
         Tag             =   "XYF"
         Top             =   180
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ORDER NO."
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
         Index           =   6
         Left            =   120
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   210
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ORDER DATE"
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
         Left            =   5160
         TabIndex        =   4
         Tag             =   "XYF"
         Top             =   210
         Width           =   1455
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1080
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPWarehouseOld.frx":2961
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPWarehouseOld.frx":383B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPWarehouseOld.frx":4715
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPWarehouseOld.frx":55EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPWarehouseOld.frx":64C9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "E&xit Order"
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
      Image           =   "SPWarehouseOld.frx":73A3
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Refresh"
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
      Image           =   "SPWarehouseOld.frx":793D
      cBack           =   -2147483633
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SPAREPARTS WAREHOUSE"
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
      Left            =   9360
      TabIndex        =   29
      Top             =   480
      Width           =   3000
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SPAREPARTS WAREHOUSE"
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
      Left            =   9390
      TabIndex        =   30
      Top             =   510
      Width           =   3000
   End
   Begin VB.Label lblChk 
      Caption         =   "ü"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   28
      Top             =   6480
      Visible         =   0   'False
      Width           =   255
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   10800
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   12630
      m_MDI_Height    =   8055
      OriginalBorder  =   2
      OriginalCaption =   27
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "ORDERED SPARE PARTS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   8
      Top             =   1545
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
      TabIndex        =   6
      Top             =   450
      Width           =   2615
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
      TabIndex        =   7
      Top             =   480
      Width           =   2615
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "ORDERED SPARE PARTS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   270
      TabIndex        =   9
      Top             =   1575
      Width           =   2415
   End
   Begin VB.Menu mnuRows 
      Caption         =   "Spare Parts Invoice"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Row       Alt+A"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Row"
      End
   End
   Begin VB.Menu mnuStatus 
      Caption         =   "Just Scroll"
      Visible         =   0   'False
      Begin VB.Menu Mark1 
         Caption         =   "Mark Item"
      End
      Begin VB.Menu Mark0 
         Caption         =   "Unmark Item"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu MarkAll 
         Caption         =   "Mark All Item"
      End
      Begin VB.Menu UnmarkAll 
         Caption         =   "Unmark All Item"
      End
   End
End
Attribute VB_Name = "SPWarehouseOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ActRow As Integer, ActCol As Integer
Dim fRow As Integer, fCol As Integer, WHrow As Integer

Private Sub ClearEntries()
    txtOrderno.Text = ""
    txtDate.Text = ""
    lblRemarks.Caption = ""
    txtSPSorderID.Text = ""
    txtInvNo.Text = ""
    flxOrd.Rows = 1
    flxOrd.Rows = 2
    txtPrintCopy = ""
End Sub
                                                                

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxSI
        For i = fRow To .Rows - 1
            For j = fCol To 9
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 9:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 9, i + 1, i)
                    fCol = IIf(j = 9, 1, j + 1)
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
                                
Private Sub cmdFrontLine_Click()
    Dim i As Integer
    
    'Validate items if all are found
    With flxOrd
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) <> "" Then
                '***Access
                'LYRcon.Execute "UPDATE SPSORDERDETAIL SET SPSORDERDETAIL.STATUS = 'FOUND' " & _
                           "WHERE (((SPSORDERDETAIL.SPSORDERID)=" & txtSPSorderID.Text & ") AND ((SPSORDERDETAIL.SPAREPARTID)=" & flxOrd.TextMatrix(i, 6) & "));"
                'mySQL
                LYRcon.Execute "UPDATE SPSORDERDETAIL SET SPSORDERDETAIL.STATUS = 'FOUND' " & _
                           "WHERE SPSORDERDETAIL.SPSORDERID=" & txtSPSorderID.Text & " AND SPSORDERDETAIL.SPAREPARTID=" & flxOrd.TextMatrix(i, 6) & ";"
                
            Else
                '***Access
                'LYRcon.Execute "UPDATE SPSORDERDETAIL SET SPSORDERDETAIL.STATUS = 'BLANK' " & _
                           "WHERE (((SPSORDERDETAIL.SPSORDERID)=" & txtSPSorderID.Text & ") AND ((SPSORDERDETAIL.SPAREPARTID)=" & flxOrd.TextMatrix(i, 6) & "));"
                'mySQL
                LYRcon.Execute "UPDATE SPSORDERDETAIL SET SPSORDERDETAIL.STATUS = 'BLANK' " & _
                           "WHERE SPSORDERDETAIL.SPSORDERID=" & txtSPSorderID.Text & " AND SPSORDERDETAIL.SPAREPARTID=" & flxOrd.TextMatrix(i, 6) & ";"
                
            End If
        Next i
    End With
    '***Access
    'LYRcon.Execute "UPDATE SPSALESORDER SET SPSALESORDER.DESTINATION = 'WAREHOUSE' " & _
                   "WHERE (((SPSALESORDER.SPSORDERID)=" & txtSPSorderID.Text & "));"
    'mySQL
    LYRcon.Execute "UPDATE SPSALESORDER SET SPSALESORDER.DESTINATION = 'WAREHOUSE' " & _
                   "WHERE SPSALESORDER.SPSORDERID=" & txtSPSorderID.Text & ";"
    
    MsgBox "Sales Order items updated to Frontline!", vbInformation + vbOKOnly, "Updated to Frontline"
    
End Sub
                                            
Private Sub cmdPrint_Click()
    Dim PrnVal(0 To 22) As String
    Dim i As Integer
    
    With flxOrd
        flxPRN.Rows = 1
        For i = 1 To flxOrd.Rows - 1
            flxPRN.AddItem vbTab & .TextMatrix(i, 1) & vbTab & .TextMatrix(i, 3) & vbTab & .TextMatrix(i, 5)
        Next i
    End With
    PrnVal(0) = "Order# " & txtOrderno.Text & "  " & txtDate.Text
    PrnVal(1) = "Partsman: " & txtFrnt.Text
    PrnVal(2) = "QTY"
    PrnVal(3) = "PART NUMBER"
    PrnVal(4) = "LOCATION"
    Call PrintMyForm(PrnVal, "WAREHOUSE PICKING", flxPRN)
    MsgBox "Order No. " & txtOrderno.Text & " already Printed!"
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim Ans As Integer, i As Integer
    Dim myMsg As String, PrnVal(0 To 12) As String
    
    Select Case Index
        
        Case 1  'Refresh
            Call DisplayPendingOrders
        Case 2  'Exit
            'MotoRep.tbMain.Visible = True
            ''MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
            Unload Me
    End Select
            
End Sub

                              
Private Sub flxOrd_Click()
    ActRow = flxOrd.Row
    ActCol = flxOrd.Col
End Sub

Private Sub flxOrd_DblClick()
    ActRow = flxOrd.Row
    ActCol = flxOrd.Col
                         
    If ActCol = 0 Then
        If flxOrd.TextMatrix(ActRow, 0) = "" Then
            flxOrd.Row = ActRow: flxOrd.Col = 0: flxOrd.CellFontName = "Wingdings": flxOrd.CellFontSize = 12: flxOrd.Text = lblChk.Caption
        Else
            flxOrd.TextMatrix(ActRow, 0) = ""
        End If
    End If
End Sub
                                                    
Private Sub flxOrd_EnterCell()
    ActRow = flxOrd.Row
    ActCol = flxOrd.Col
End Sub

Private Sub flxOrd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxOrd_DblClick
End Sub
                                                  
Private Sub flxOrd_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        Call PopupMenu(mnuStatus)
    End If
End Sub


Private Sub DisplayPendingOrders()
    Dim rsPending As New ADODB.Recordset
    Dim Qry As String
    
    'old Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.DESTINATION From SPSALESORDER " & _
          "WHERE (((SPSALESORDER.INVOICENO) Is Null) AND ((SPSALESORDER.EMPID)=" & EmpId & ") AND ((SPSALESORDER.ORDERDATE)=#" & Date & "#));"
    '***Access
    'Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.FRONTLINER, SPSALESORDER.SPSORDERID From SPSALESORDER " & _
          "WHERE (((SPSALESORDER.DESTINATION)='WAREHOUSE') AND ((SPSALESORDER.INVOICENO) Is Null) AND ((SPSALESORDER.ORDERDATE)=#" & Date & "#));"
    'mySQL
    Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.FRONTLINER, SPSALESORDER.SPSORDERID From SPSALESORDER " & _
          "WHERE SPSALESORDER.DESTINATION='WAREHOUSE' AND (SPSALESORDER.INVOICENO) Is Null AND SPSALESORDER.ORDERDATE='" & Format(Date, "yyyy-mm-dd") & "';"

    With rsPending
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxWH.Rows = 1
            Do While Not .EOF
                flxWH.AddItem !ORDERNO & vbTab & !Frontliner & vbTab & !SPSORDERID
                .MoveNext
            Loop
        Else
            flxWH.Rows = 1
        End If
        .Close
    End With
    
End Sub

Function GetMarkings() As String
    Dim i As Integer, MyM As String
    
    With flxOrd
        MyM = ""
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) = "" Then
                MyM = IIf(MyM = "", "0", MyM & "|0")
            Else
                MyM = IIf(MyM = "", "1", MyM & "|1")
            End If
        Next i
    End With
    GetMarkings = MyM
End Function

Private Sub flxwh_Click()
    Dim rsOrder As New ADODB.Recordset
    Dim rsOrdDtls As New ADODB.Recordset
    Dim Qry As String, MyMark As String
    Dim i As Integer
    Dim MyTmp() As String
    
    If flxWH.Cols = 0 Or flxWH.Row = 0 Then Exit Sub
    'Store marks on flxHW
    If txtOrderno.Text <> "" Then
        MyMark = GetMarkings
        For i = 1 To flxWH.Rows - 1
            If txtSPSorderID.Text = flxWH.TextMatrix(i, 2) Then
                flxWH.TextMatrix(i, 3) = MyMark
                Exit For
            End If
        Next i
    End If
    WHrow = flxWH.Row
    '***Access
    'Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.ORDERDATE, SPSALESORDER.SPSORDERID, SPSALESORDER.FRONTLINER " & _
          "From SPSALESORDER WHERE (((SPSALESORDER.SPSORDERID)=" & flxWH.TextMatrix(flxWH.Row, 2) & "));"
    'mySQL
    Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.ORDERDATE, SPSALESORDER.SPSORDERID, SPSALESORDER.FRONTLINER " & _
          "From SPSALESORDER WHERE SPSALESORDER.SPSORDERID=" & flxWH.TextMatrix(flxWH.Row, 2) & ";"
    
    With rsOrder
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtOrderno.Text = !ORDERNO
            txtDate.Text = !ORDERDATE
            txtFrnt.Text = !Frontliner
            txtSPSorderID.Text = !SPSORDERID
        Else
            txtOrderno.Text = ""
            txtDate.Text = ""
            txtFrnt.Text = ""
            txtSPSorderID.Text = ""
        End If
        .Close
    End With
    If txtSPSorderID.Text = "" Then Exit Sub
    '***Access
    'Qry = "SELECT SPSORDERDETAIL.STATUS, SPSORDERDETAIL.QTY, SPSORDERDETAIL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.LOCATION, SPSORDERDETAIL.SPAREPARTID " & _
          "FROM SPSORDERDETAIL INNER JOIN SPAREPARTS ON SPSORDERDETAIL.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE (((SPSORDERDETAIL.SPSORDERID)=" & txtSPSorderID.Text & "));"
    'mySQL
    Qry = "SELECT SPSORDERDETAIL.STATUS, SPSORDERDETAIL.QTY, SPSORDERDETAIL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.LOCATION, SPSORDERDETAIL.SPAREPARTID " & _
          "FROM SPSORDERDETAIL INNER JOIN SPAREPARTS ON SPSORDERDETAIL.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE SPSORDERDETAIL.SPSORDERID=" & txtSPSorderID.Text & ";"
    
    With rsOrdDtls
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxOrd.Rows = 1
            Do While Not .EOF
                flxOrd.AddItem IIf(IsNull(!STATUS), "", !STATUS) & vbTab & !Qty & vbTab & !Unit & vbTab & !PartNumber & vbTab & !Description & vbTab & IIf(IsNull(!Location), "", !Location) & vbTab & !SparepartID
                If flxOrd.TextMatrix(flxOrd.Rows - 1, 0) = "FOUND" Or flxOrd.TextMatrix(flxOrd.Rows - 1, 0) = "DROP" Then
                    flxOrd.Row = flxOrd.Rows - 1: flxOrd.Col = 0: flxOrd.CellFontName = "Wingdings": flxOrd.CellFontSize = 12: flxOrd.Text = lblChk.Caption
                Else
                    flxOrd.TextMatrix(flxOrd.Rows - 1, 0) = ""
                End If
                .MoveNext
            Loop
            'Display Updated markings
            If flxWH.TextMatrix(WHrow, 3) <> "" Then
                MyTmp = Split(flxWH.TextMatrix(WHrow, 3), "|")
                For i = 1 To flxOrd.Rows - 1
                    If MyTmp(i - 1) = "1" Then
                        flxOrd.Row = i: flxOrd.Col = 0: flxOrd.CellFontName = "Wingdings": flxOrd.CellFontSize = 12: flxOrd.Text = lblChk.Caption
                    End If
                Next i
            End If
        Else
            flxOrd.Rows = 1
            flxOrd.Rows = 2
        End If
        .Close
    End With
    cmdPrint.Visible = True
    cmdFrontLine.Visible = True
End Sub

Private Sub Form_Load()

    Call DisplayPendingOrders
    txtDate.Text = Format(ServerTime, "mm-dd-yyyy")
    MotoRep.sBar.Panels(6).Text = Me.Caption
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
End Sub


Private Sub Mark0_Click()
    flxOrd.TextMatrix(ActRow, 0) = ""
End Sub

Private Sub Mark1_Click()
    'If ActCol = 0 Then
    flxOrd.Row = ActRow: flxOrd.Col = 0: flxOrd.CellFontName = "Wingdings": flxOrd.CellFontSize = 12: flxOrd.Text = lblChk.Caption
    'End If
End Sub

Private Sub MarkAll_Click()
    Dim i As Integer
                                 
    For i = 1 To flxOrd.Rows - 1
        flxOrd.Row = i: flxOrd.Col = 0: flxOrd.CellFontName = "Wingdings": flxOrd.CellFontSize = 12: flxOrd.Text = lblChk.Caption
    Next i
End Sub

Private Sub RXResizer1_ResizeComplete()
    With flxOrd
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.15
        .ColWidth(2) = (.Width - 300) * 0.15
        .ColWidth(3) = (.Width - 300) * 0.25
        .ColWidth(4) = (.Width - 300) * 0.25
        .ColWidth(5) = (.Width - 300) * 0.2
        .ColWidth(6) = 0
'        .Col = 6: .Row = 1
    End With
    With flxWH
        .ColWidth(0) = .Width * 0.35
        .ColWidth(1) = .Width * 0.65
    End With
    lblHeading(2).Left = flxWH.Left + flxWH.Width - lblHeading(2).Width
    lblHeading(3).Left = flxWH.Left + flxWH.Width - lblHeading(3).Width - 30
    lblHeading(4).Top = flxOrd.Top - lblHeading(4).Height + 30
    lblHeading(5).Top = flxOrd.Top - lblHeading(5).Height
    lblHeading(4).Left = flxOrd.Left
    lblHeading(5).Left = flxOrd.Left + 30
End Sub


Private Sub txtFind_Change()
    fRow = 1
    fCol = 1
End Sub


Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        txtOrderno.SetFocus
    ElseIf KeyCode = vbKeyDown Then
        flxOrd.SetFocus
        flxOrd.Col = 4  'partnumber
    End If
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxOrd.SetFocus
        flxOrd.Col = 4  'partnumber
    End If
End Sub

Private Sub txtOrderNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        txtAddress.SetFocus
    ElseIf KeyCode = vbKeyDown Then
        txtDate.SetFocus
    End If
End Sub

Private Sub txtOrderNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDate.SetFocus
End Sub

Private Sub UnmarkAll_Click()
    Dim i As Integer
    
    For i = 1 To flxOrd.Rows - 1
        flxOrd.TextMatrix(i, 0) = ""
    Next i
End Sub
