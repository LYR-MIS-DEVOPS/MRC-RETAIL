VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmOtherAdjustment 
   Caption         =   "Adjustment Request"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   13215
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   13320
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   13320
      TabIndex        =   35
      Text            =   "HoldBackColor"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame frmFooter 
      Height          =   975
      Left            =   120
      TabIndex        =   30
      Tag             =   "XYF"
      Top             =   9840
      Width           =   12975
      Begin VB.ComboBox txtApprvBy 
         Height          =   315
         Left            =   1440
         TabIndex        =   38
         Top             =   360
         Width           =   4815
      End
      Begin VB.TextBox txtAuditApprvBy 
         Height          =   285
         Left            =   8160
         TabIndex        =   34
         Tag             =   "XY"
         Top             =   600
         Width           =   4695
      End
      Begin VB.TextBox txtMISIncharge 
         Height          =   285
         Left            =   8160
         TabIndex        =   32
         Tag             =   "XY"
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Label12 
         Caption         =   "APPROVED BY:"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Tag             =   "XY"
         Top             =   375
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "AUDIT APPROVED BY:"
         Height          =   255
         Left            =   6360
         TabIndex        =   33
         Tag             =   "XY"
         Top             =   615
         Width           =   1815
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "MIS IN-CHARGE:"
         Height          =   255
         Left            =   6360
         TabIndex        =   31
         Tag             =   "XY"
         Top             =   255
         Width           =   1815
      End
   End
   Begin VB.Frame frmBody 
      Height          =   7695
      Left            =   120
      TabIndex        =   25
      Tag             =   "XYF"
      Top             =   2160
      Width           =   12975
      Begin VB.TextBox txtMISDefinition 
         Height          =   3405
         IMEMode         =   3  'DISABLE
         Left            =   720
         MaxLength       =   10000
         MultiLine       =   -1  'True
         PasswordChar    =   "#"
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Tag             =   "XY"
         Top             =   4200
         Width           =   12015
      End
      Begin VB.TextBox txtExplanation 
         Height          =   3405
         Left            =   720
         MaxLength       =   10000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Tag             =   "XY"
         Top             =   480
         Width           =   12015
      End
      Begin VB.Label Label9 
         Caption         =   "MIS DATA SCRIPT DEFINITION:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Tag             =   "XY"
         Top             =   3960
         Width           =   2535
      End
      Begin VB.Label Label8 
         Caption         =   "EXPLANATION:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Tag             =   "XY"
         Top             =   240
         Width           =   1215
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6060
      TabIndex        =   0
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
      Image           =   "frmOtherAdjustment.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   1
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
      Image           =   "frmOtherAdjustment.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   4560
      TabIndex        =   2
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
      Image           =   "frmOtherAdjustment.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   7680
      TabIndex        =   3
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
      Image           =   "frmOtherAdjustment.frx":10CE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   9180
      TabIndex        =   4
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
      Image           =   "frmOtherAdjustment.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   7
      Left            =   10680
      TabIndex        =   5
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
      Image           =   "frmOtherAdjustment.frx":1C02
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1500
      TabIndex        =   6
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
      Image           =   "frmOtherAdjustment.frx":219C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   7
      Top             =   0
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   661
      Caption         =   "&Process Script"
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
      Image           =   "frmOtherAdjustment.frx":2736
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Frame frmHeader 
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Tag             =   "XYF"
      Top             =   360
      Width           =   12975
      Begin lvButton.lvButtons_H cmdScriptUpload 
         Height          =   375
         Left            =   9360
         TabIndex        =   23
         Tag             =   "XY"
         Top             =   1350
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "UPLOAD SCRIPT"
         CapAlign        =   2
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
         cBack           =   -2147483633
      End
      Begin VB.TextBox txtScriptFileName 
         Height          =   285
         Left            =   10920
         TabIndex        =   22
         Tag             =   "XY"
         Top             =   1400
         Width           =   1935
      End
      Begin VB.TextBox txtControlNo 
         Height          =   285
         Left            =   10920
         TabIndex        =   21
         Tag             =   "XY"
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtTicketNo 
         Height          =   285
         Left            =   10920
         TabIndex        =   20
         Tag             =   "XY"
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtRequestStatus 
         Height          =   285
         Left            =   1800
         TabIndex        =   14
         Tag             =   "XY"
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtSubjRequest 
         Height          =   285
         Left            =   1800
         TabIndex        =   11
         Tag             =   "XY"
         Top             =   600
         Width           =   7215
      End
      Begin VB.TextBox txtRequestBy 
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Tag             =   "XY"
         Top             =   240
         Width           =   7215
      End
      Begin MSMask.MaskEdBox txtDateRequest 
         Height          =   360
         Left            =   1800
         TabIndex        =   16
         Tag             =   "XY"
         Top             =   1320
         Width           =   1815
         _ExtentX        =   3201
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
      Begin MSMask.MaskEdBox txtDateFinal 
         Height          =   360
         Left            =   10920
         TabIndex        =   24
         Tag             =   "XY"
         Top             =   960
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
         Caption         =   "DATE FINAL:"
         Height          =   375
         Left            =   9240
         TabIndex        =   19
         Tag             =   "XY"
         Top             =   1020
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "CONTROL NO.:"
         Height          =   375
         Left            =   9240
         TabIndex        =   18
         Tag             =   "XY"
         Top             =   675
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "TICKET NO.:"
         Height          =   375
         Left            =   8880
         TabIndex        =   17
         Tag             =   "XY"
         Top             =   285
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "DATE REQUEST:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1380
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "REQUEST STATUS:"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   1020
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "SUBJ. OF REQUEST:"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   680
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "REQUEST BY:"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   280
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog cdlgVan 
      Left            =   13680
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   13680
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   13455
      m_MDI_Height    =   11520
   End
   Begin VB.Menu mnuStatus 
      Caption         =   "Popupemnu"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Row"
      End
   End
End
Attribute VB_Name = "frmOtherAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ARID As Long

Dim arrPOSCODE() As String
Dim tmpRs  As New ADODB.Recordset
Dim mainRs As New ADODB.Recordset
Dim cboRs As New ADODB.Recordset

Dim Ans As String
Dim Qry As String
Dim strPartial As String, strTotal As String
Dim formAction As String, myTicketNo As String
Dim myFile As String, POSCODE As String, myMISDifinitionScript As String

Dim i As Long, j As Long
Dim ActRow As Integer, ActCol As Integer
Dim EMPID As Long

Dim StatSrch As Boolean
Dim PrtBln As Boolean, PrtSrch As Boolean, ProperExit As Boolean, myUpload As Boolean

Public Sub CboDisplay()
On Error GoTo er
    
    'Get Data from database
    DoEvents
    Qry = "Select UCASE(`VALUE`) as `VALUE`, REMARKS From settings Where (REPLACE(DESCRIPTION,'AR - ','') = '" & sysBranchCode & "' or REPLACE(DESCRIPTION,'AR - ','') = 'ADMIN') and REMARKS IN ('BM','OIC','AOM')"
    If cboRs.State = adStateOpen Then cboRs.Close
    cboRs.Open Qry, LYRcon, adOpenStatic, adLockReadOnly
    If cboRs.State = 1 Then
        Call modObjTools.addToListObjectWithArrayGivenField(txtApprvBy, cboRs, arrPOSCODE, "REMARKS", "VALUE")
    End If
    If cboRs.State <> 0 Then cboRs.Close
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplay()")
End Sub

Function GetFileNameFromPath(strFullPath As String) As String
    GetFileNameFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))
End Function

Public Sub ProcessingAR()
On Error GoTo er
    
    If txtRequestStatus.Text = "PENDING" Then
        LYRcon.Execute "Update other_adjustment SET req_status = 'PROCESSING' Where TICKETNO = '" & txtTicketNo.Text & "' "
        txtRequestStatus.Text = "PROCESSING"
        MsgBox "Your Request is now processing!", vbInformation + vbOKOnly, "Update"
    End If

    cmdTrans(1).Enabled = False
    cmdTrans(2).Enabled = True
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->ProcessingAR()")
End Sub
Public Sub ShowRecord(myTicketNo As String)
On Error GoTo er

    Screen.MousePointer = vbHourglass
    
    DoEvents
    Qry = "call sp_mrcretails.sp_other_adjustment_dtls('" & sysBranchCode & "', '" & myTicketNo & "')"
    With tmpRs
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtRequestBy.Text = IIf(IsNull(!EMPLOYEENAME), "", !EMPLOYEENAME)
            txtSubjRequest.Text = IIf(IsNull(!subj_request), "", !subj_request)
            txtRequestStatus.Text = IIf(IsNull(!req_status), "", !req_status)
            txtDateRequest.Text = Format(!req_date, "MM/DD/YYYY")
            txtTicketNo.Text = IIf(IsNull(!ticketno), "<AUTO>", !ticketno)
            txtControlNo.Text = IIf(IsNull(!controlno), "", !controlno)
            txtDateFinal.Text = IIf(IsNull(!date_final), "__/__/____", Format(!date_final, "MM/DD/YYYY"))
            txtScriptFileName.Text = IIf(IsNull(!script_file), "", !script_file)
            txtExplanation.Text = IIf(IsNull(!explanation), "", !explanation)
            txtMISDefinition.Text = IIf(IsNull(!mis_definition), "", !mis_definition)
            myMISDifinitionScript = IIf(IsNull(!mis_definition), "", !mis_definition)
            txtApprvBy.Text = IIf(IsNull(!approved_by), "", !approved_by)
            txtMISIncharge.Text = IIf(IsNull(!mis_incharge), "", !mis_incharge)
            txtAuditApprvBy.Text = IIf(IsNull(!aud_apprvd_by), "", !aud_apprvd_by)
            ARID = !oa_id
            EMPID = !ENCODEDBY
            
            If !req_status = "PROCESSING" Then
                cmdTrans(1).Enabled = False
                cmdTrans(2).Enabled = True
            Else
                cmdTrans(1).Enabled = True
                cmdTrans(2).Enabled = False
            End If
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
    
    If Trim(txtSubjRequest.Text) = "" Then
        MsgBox "SUBJECT OF REQUEST IS REQUIRED!", vbExclamation, "Validation"
        txtSubjRequest.SetFocus
        Exit Function
    End If
    
    If Trim(txtExplanation.Text) = "" Then
        MsgBox "EXPLANATION IS REQUIRED!", vbExclamation, "Validation"
        txtExplanation.SetFocus
        Exit Function
    End If
        
    If Trim(txtDateRequest.Text) = "__/__/____" Then
        MsgBox "REQUESTE DATE IS REQUIRED!", vbExclamation, "Validation"
        txtDateRequest.SetFocus
        Exit Function
    End If
    
    If Trim(txtApprvBy.Text) = "" Then
        MsgBox "APPROVED BY IS REQUIRED!", vbExclamation, "Validation"
        txtApprvBy.SetFocus
        Exit Function
    End If
    
    If Trim(txtTicketNo.Text) = "<AUTO>" Then
        MsgBox "TICKET NO IS REQUIRED! YOU NEED TO CLOSE THE MODULE THEN OPEN AGAIN", vbExclamation, "Validation"
        txtTicketNo.SetFocus
        Exit Function
    End If
    
    If formAction = "PROCESS" Then
        If Trim(txtControlNo.Text) = "" Then
            MsgBox "CONTROL NO. IS REQUIRED!", vbExclamation, "Validation"
            txtControlNo.SetFocus
            Exit Function
        End If
        
        If Trim(txtMISDefinition.Text) = "" Then
            MsgBox "MIS DATA SCRIPT DEFINITION IS REQUIRED!", vbExclamation, "Validation"
            txtMISDefinition.SetFocus
            Exit Function
        End If
'        If Trim(txtAuditApprvBy.Text) = "" Then
'            MsgBox "AUDIT APPROVED IS REQUIRED!", vbExclamation, "Validation"
'            txtAuditApprvBy.SetFocus
'            Exit Function
'        End If
        If Trim(txtMISIncharge.Text) = "" Then
            MsgBox "MIS In-Charge IS REQUIRED!", vbExclamation, "Validation"
            txtMISIncharge.SetFocus
            Exit Function
        End If
        
        If Trim(txtScriptFileName.Text) = "" Then
            MsgBox "SCRIPT FILE NAME IS REQUIRED!", vbExclamation, "Validation"
            txtScriptFileName.SetFocus
            Exit Function
        End If
    End If
    
    ValidateEntry = True

End Function




Public Function getTicketNo() As String
Dim cntAdj As Long

    cntAdj = Val(GetFieldValue("Select MAX(OA_ID) as SEQNO From other_adjustment"))
    cntAdj = cntAdj + 1
    
    getTicketNo = sysBranchCode & Year(Now()) & cntAdj

End Function


Sub Clearform()

    'frmHeader
    txtRequestBy.Text = ""
    txtSubjRequest.Text = ""
    txtRequestStatus.Text = ""
    txtDateRequest.Text = Format(Now, "mm/dd/yyyy")
    txtTicketNo.Text = ""
    txtControlNo.Text = ""
    txtDateFinal.Text = "__/__/____"
    txtScriptFileName.Text = ""
    'frmBody
    txtExplanation.Text = ""
    txtMISDefinition.Text = ""
    'frmFooter
    txtApprvBy.Text = ""
    txtMISIncharge.Text = ""
    txtAuditApprvBy.Text = ""
    
    ARID = 0
    POSCODE = ""
    EMPID = 0
    myMISDifinitionScript = ""
    myUpload = False
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
    For i = 0 To 5
        cmdTrans(i).Visible = False
    Next i

    If DataEntry Then
        cmdTrans(6).Visible = True
        cmdTrans(7).Visible = True
    Else
        For i = 0 To 5
            cmdTrans(i).Visible = True
            cmdTrans(6).Visible = False
            cmdTrans(7).Visible = False
        Next i
    End If
End Sub


Sub LockData(DataEntry As Boolean)
    DataEntry = Not DataEntry
    
    'frmHeader
    txtRequestBy.Locked = DataEntry
    txtSubjRequest.Locked = DataEntry
    txtRequestStatus.Locked = DataEntry
    txtDateRequest.Enabled = Not DataEntry
    txtTicketNo.Locked = DataEntry
    txtControlNo.Locked = DataEntry
    txtDateFinal.Enabled = Not DataEntry
    txtScriptFileName.Locked = DataEntry
    'frmBody
    txtExplanation.Locked = DataEntry
    txtMISDefinition.Locked = DataEntry
    'frmFooter
    txtApprvBy.Locked = DataEntry
    txtMISIncharge.Locked = DataEntry
    txtAuditApprvBy.Locked = DataEntry
    
    If txtRequestBy.Locked Then
        'frmHeader
        txtRequestBy.BackColor = txtHoldGrayBackColor.BackColor
        txtSubjRequest.BackColor = txtHoldGrayBackColor.BackColor
        txtRequestStatus.BackColor = txtHoldGrayBackColor.BackColor
        'txtDateRequest.BackColor = txtHoldGrayBackColor.BackColor
        txtTicketNo.BackColor = txtHoldGrayBackColor.BackColor
        txtControlNo.BackColor = txtHoldGrayBackColor.BackColor
        txtDateFinal.BackColor = txtHoldGrayBackColor.BackColor
        txtScriptFileName.BackColor = txtHoldGrayBackColor.BackColor
        'frmBody
        txtExplanation.BackColor = txtHoldGrayBackColor.BackColor
        txtMISDefinition.BackColor = txtHoldGrayBackColor.BackColor
        'frmFooter
        txtApprvBy.BackColor = txtHoldGrayBackColor.BackColor
        txtMISIncharge.BackColor = txtHoldGrayBackColor.BackColor
        txtAuditApprvBy.BackColor = txtHoldGrayBackColor.BackColor
    Else
        'frmHeader
        txtRequestBy.Locked = True
        txtSubjRequest.BackColor = txtHoldWhiteBackColor.BackColor
        txtRequestStatus.Locked = True
        txtDateRequest.Enabled = True
        txtTicketNo.Locked = True
        txtControlNo.Locked = True
        txtDateFinal.Enabled = False
        txtScriptFileName.Locked = True
        'frmBody
        txtExplanation.BackColor = txtHoldWhiteBackColor.BackColor
        txtMISDefinition.Locked = True
        'frmFooter
        txtApprvBy.BackColor = txtHoldWhiteBackColor.BackColor
        txtMISIncharge.Locked = True
        txtAuditApprvBy.Locked = True
    End If
End Sub

Function ReadFileToText(strFile As String)
    Dim objStream, strData
    Dim myControlStorage As String, myFileName As String
    
    Set objStream = CreateObject("ADODB.Stream")
    
    objStream.Charset = "utf-8"
    objStream.Open
    objStream.LoadFromFile (strFile)
    
    strData = objStream.ReadText()
    
    objStream.Close
    Set objStream = Nothing
    'InStr(SysBranch, "VAN") >= 1
    If InStr(strFile, "control_no.txt") >= 1 Then
        myControlStorage = strData
        myFileName = GetFileNameFromPath(strFile)
        Call ExecuteSQLScript(LYRcon, myControlStorage, myFileName)
    Else
        txtMISDefinition.Text = strData
        myMISDifinitionScript = strData
        txtScriptFileName.Text = GetFileNameFromPath(strFile)
    End If
    
End Function
Private Sub cmdScriptUpload_Click()
myFile = ""
If formAction = "PROCESS" Then
    With cdlgVan
        .filename = ""
        .CancelError = True
        '.Filter = "(*.sql)|*.sql"
        .InitDir = App.Path
        .ShowOpen
        myFile = .filename
    End With
    Call Extract(myFile)
End If
End Sub

Sub Extract(myFile As String)
    Dim RarIt As String
    Dim Source As String
    Dim Desti As String
    Dim WinRarPath As String
    Dim i As Integer
    
    WinRarPath = "C:\Program Files\WinRar\"
    Source = myFile
    If Dir("D:\mis_scriptfolder\") = "" Then
        MkDir "D:\mis_scriptfolder\"
    End If
    Desti = "D:\mis_scriptfolder\"
    RarIt = Shell(WinRarPath & "WinRar.exe e -plYrMIS@2021 " & Source & " " & Desti, vbNormalFocus)
    
    If Dir("D:\mis_scriptfolder\control_no.txt", vbHidden) <> "" Then
        Call ReadFileToText("D:\mis_scriptfolder\control_no.txt")
        Kill "D:\mis_scriptfolder\control_no.txt"
    Else
        MsgBox "control_no text not found!", vbInformation + vbOKOnly, "Validation"
        Exit Sub
    End If
    
    If Dir("D:\mis_scriptfolder\mis_script.txt", vbHidden) <> "" Then
        Call ReadFileToText("D:\mis_scriptfolder\mis_script.txt")
        Kill "D:\mis_scriptfolder\mis_script.txt"
    Else
        MsgBox "mis script text not found!", vbInformation + vbOKOnly, "Validation"
        Exit Sub
    End If
    RmDir ("D:\mis_scriptfolder")
    myUpload = True
    Call txtControlNo_LostFocus
End Sub

Private Sub ExecuteSQLScript(objConn As ADODB.Connection, sSQLScript As String, strFile As String)
Dim commands As Variant
Dim cmd As String
Dim i As Integer
Dim results As String

    ' Break the script into semi-colon
    ' delimited commands.
    commands = Split(sSQLScript, ";")

    ' Execute each command.
    On Error Resume Next
    For i = LBound(commands) To UBound(commands)
        ' Clean up the command.
        cmd = commands(i)
        cmd = Replace(cmd, vbCr, " ")
        cmd = Replace(cmd, vbLf, " ")
        cmd = Trim$(cmd)
        
        ' Execute only non-blank commands.
        If Len(cmd) > 0 Then
            'Execute the non-query command.
            If InStr(strFile, "control_no.txt") >= 1 Then
                If i = 0 Then
                    txtControlNo.Text = commands(i)
                ElseIf i = 1 Then
                    txtMISIncharge.Text = commands(i)
                ElseIf i = 2 Then
                    txtAuditApprvBy.Text = commands(i)
                End If
            ElseIf InStr(strFile, "mis_script.txt") >= 1 Then
                objConn.Execute commands(i), , adCmdText
            End If
        End If
    Next i

    On Error GoTo 0
End Sub

Private Sub cmdTrans_Click(Index As Integer)
On Error GoTo er
Dim myDefinition As String

    If Index = 0 Then 'add
        Call Clearform
        Call LockData(True)
        Call SetButtonVisibility(True)
        formAction = "add"
        txtRequestBy.Text = GetEmployeeName
        EMPID = UserID
        txtRequestStatus.Text = "PENDING"
        txtTicketNo.Text = getTicketNo
        txtSubjRequest.SetFocus
    ElseIf Index = 1 Then 'edit
        If txtRequestStatus.Text = "PROCESSING" Then
            MsgBox "This Transaction is already processing to HO", vbExclamation + vbOKOnly, "Edit"
            Exit Sub
        End If
        If txtRequestStatus.Text = "FINAL" Then
            MsgBox "This Transaction is already Final", vbExclamation + vbOKOnly, "Script"
            Exit Sub
        End If
        Call SetButtonVisibility(True)
        Call LockData(True)
        formAction = "edit"
    ElseIf Index = 2 Then 'process script
        If txtRequestStatus.Text = "FINAL" Then
            MsgBox "This Transaction is already Final", vbExclamation + vbOKOnly, "Script"
            Exit Sub
        End If
        Call SetButtonVisibility(True)
        'Call LockData(True)
        txtControlNo.Locked = False
        txtControlNo.BackColor = txtHoldWhiteBackColor.BackColor
        txtDateFinal.Enabled = True
        txtDateFinal.BackColor = txtHoldWhiteBackColor.BackColor
        txtScriptFileName.Locked = False
        txtScriptFileName.BackColor = txtHoldWhiteBackColor.BackColor
        txtAuditApprvBy.Locked = False
        txtAuditApprvBy.BackColor = txtHoldWhiteBackColor.BackColor
        txtMISIncharge.Locked = False
        txtMISIncharge.BackColor = txtHoldWhiteBackColor.BackColor
        formAction = "PROCESS"
    ElseIf Index = 3 Then 'view summary
        FrmSummary.mySrc = "ADJUSTMENT REQUEST"
        FrmSummary.Show
        Unload Me
    ElseIf Index = 4 Then 'print
        Dim empLastName As String
        Dim chkExists As String
        Dim empPosition As String
        Dim ApproverPosition As String
        
        If txtTicketNo.Text = "<AUTO>" Then
            MsgBox "No Records Found", vbExclamation + vbOKOnly, "PRINT"
            Exit Sub
        End If
    
        Qry = "call sp_mrcretails.sp_other_adjustment_dtls('" & sysBranchCode & "','" & txtTicketNo.Text & "')"
        frmPrint.QryPrn = Qry

        frmPrint.ParamValue1 = GetFieldValue("Select `VALUE` From settings Where DESCRIPTION = 'ADJUSTMENT REQUEST' and REMARKS = 'AUDIT DIVISION HEAD' ")
        frmPrint.ParamValue2 = GetFieldValue("Select `VALUE` From settings Where DESCRIPTION = 'ADJUSTMENT REQUEST' and REMARKS = 'MIS' ")
        
        ApproverPosition = GetFieldValue("Select REMARKS From settings Where UCASE(VALUE) = '" & txtApprvBy.Text & "' and DESCRIPTION like '%AR -%' ")
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
                frmPrint.ParamValue4 = "Branch Manager"
            ElseIf chkExists = "OIC" Then
                frmPrint.ParamValue4 = "Branch Operataion In-Chief"
            ElseIf chkExists = "AOM" Then
                frmPrint.ParamValue4 = "Area Operation Manager"
            End If
        Else
            empPosition = GetFieldValue("Select POSITION From employees Where EMPID = " & Val(EMPID) & " ")
            If empPosition <> "" Then
                frmPrint.ParamValue4 = empPosition
            End If
        End If
        
        frmPrint.MySource = "ADJUSTMENT REQUEST"
        frmPrint.Show
    ElseIf Index = 5 Then 'exit
        formAction = "EXIT"
        Call Clearform
        formAction = ""
        Unload Me
    ElseIf Index = 6 Then 'save
        If ValidateEntry() = False Then Exit Sub

        Ans = MsgBox("Do you want to save?", vbQuestion + vbYesNo, "Save")

        If Ans = vbNo Then Exit Sub

        If formAction = "add" Then
            DoEvents
            Qry = "call sp_mrcretails.sp_other_adjustment_create('" & sysBranchCode & "', '" & txtTicketNo.Text & "', '" & txtSubjRequest.Text & "', '" & txtRequestStatus.Text & "', '" & Format(txtDateRequest.Text, "YYYY-mm-dd") & "', '" & txtExplanation.Text & "', '" & txtApprvBy.Text & "', " & UserID & ")"
            If tmpRs.State = adStateOpen Then tmpRs.Close
            tmpRs.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If tmpRs.RecordCount > 0 Then
                If tmpRs!level = "Granted" Then
                    ARID = tmpRs!ARID
                    MsgBox tmpRs!message, vbInformation + vbOKOnly, "SAVE"
                Else
                    MsgBox tmpRs!message, vbInformation + vbOKOnly, "SAVE"
                    Exit Sub
                End If
            End If
        ElseIf formAction = "edit" Then
            DoEvents
            Qry = "call sp_mrcretails.sp_other_adjustment_update('" & sysBranchCode & "'," & Val(ARID) & ", '" & txtTicketNo.Text & "', '" & txtSubjRequest.Text & "', '" & txtRequestStatus.Text & "', '" & Format(txtDateRequest.Text, "YYYY-mm-dd") & "', '" & txtExplanation.Text & "', '" & txtApprvBy.Text & "', " & UserID & ")"
            If tmpRs.State = adStateOpen Then tmpRs.Close
            tmpRs.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If tmpRs.RecordCount > 0 Then
                If tmpRs!level = "Granted" Then
                    MsgBox tmpRs!message, vbInformation + vbOKOnly, "SAVE"
                Else
                    MsgBox tmpRs!message, vbInformation + vbOKOnly, "SAVE"
                    Exit Sub
                End If
            End If
        ElseIf formAction = "PROCESS" Then
            DoEvents
            myDefinition = Replace(myMISDifinitionScript, "'", "\'")
            myDefinition = Replace(myDefinition, "=", "\=")
            Qry = "call sp_mrcretails.sp_other_adjustment_processScript('" & sysBranchCode & "', '" & txtTicketNo.Text & "', '" & txtControlNo.Text & "', '" & txtScriptFileName.Text & "', '" & Format(txtDateFinal.Text, "YYYY-mm-dd") & "', '" & myDefinition & "', '" & Trim(txtMISIncharge.Text) & "', '" & Trim(txtAuditApprvBy.Text) & "', " & UserID & ")"
            If tmpRs.State = adStateOpen Then tmpRs.Close
            tmpRs.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If tmpRs.RecordCount > 0 Then
                If tmpRs!level = "Granted" Then
                    Call ExecuteSQLScript(LYRcon, myMISDifinitionScript, "mis_script.txt")
                    txtRequestStatus.Text = "FINAL"
                    MsgBox tmpRs!message, vbInformation + vbOKOnly, "SAVE"
                Else
                    MsgBox tmpRs!message, vbInformation + vbOKOnly, "SAVE"
                    tmpRs.Close
                    Exit Sub
                End If
                tmpRs.Close
            End If
        End If
        Call SetButtonVisibility(False)
        Call LockData(False)
        formAction = ""
    ElseIf Index = 7 Then 'cancel
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
    txtDateRequest.Text = Format$(Now, "mm/dd/yyyy")
    txtTicketNo.Text = "<AUTO>"
    Call CboDisplay
    '===================BUTTON POSITION===================
    If cmdTrans(0).Visible = False Then
        'Align Save Button
        cmdTrans(6).Left = cmdTrans(0).Left
        cmdTrans(6).Top = cmdTrans(0).Top
        'Align Cancel Button
        cmdTrans(7).Left = cmdTrans(1).Left
        cmdTrans(7).Top = cmdTrans(1).Top
    End If
End Sub

Private Sub txtApprvBy_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With txtApprvBy
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtApprvBy-Change()")
End Sub

Private Sub txtApprvBy_Click()
    POSCODE = arrPOSCODE(txtApprvBy.ListIndex, 0)
End Sub

Private Sub txtApprvBy_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub txtApprvBy_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtApprvBy, KeyAscii, ValidChars)
End Sub

Private Sub txtApprvBy_LostFocus()
Dim cbolistIndex As Integer

    cbolistIndex = modObjTools.getListIndex(txtApprvBy, txtApprvBy.Text)
    
    If cbolistIndex < 0 Then
        txtApprvBy.Text = ""
        txtSubjRequest.SetFocus
        MsgBox "The Selected Names were not found on the List!", vbInformation + vbOKOnly, "Validation"
    Else
        txtApprvBy.ListIndex = cbolistIndex
        Call txtApprvBy_Click
    End If
End Sub

Private Sub txtAuditApprvBy_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtAuditApprvBy, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        txtSubjRequest.SetFocus
    End If
End Sub

Private Sub txtControlNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(Chr(KeyAscii))
    Call ObjKeyTrapping.Validation(txtControlNo, KeyAscii, ValidCodeCharsWithSpace)
    If KeyAscii = 13 Then
        txtDateFinal.SetFocus
    End If
End Sub

Private Sub txtControlNo_LostFocus()
Dim chkExists As String, chkExistsID As Integer

chkExists = ""
chkExistsID = 0
If txtControlNo.Text <> "" Then
    ' query for branch other adjustment table
    chkExists = GetFieldValue("Select TICKETNO From other_adjustment Where CONTROLNO = '" & txtControlNo.Text & "'")
    If chkExists = "" Then
        If myUpload = True Then Exit Sub
        ' query for adjustment monitoring database
        chkExistsID = Val(GetFieldValue("Select script_adjustment_id From adjustment_monitoring.script_adjustment Where CONTROL_NO = '" & txtControlNo.Text & "' and STATUS ='FBE' "))
        If Val(chkExistsID) > 0 Then
            Qry = "call adjustment_monitoring.sp_script_adjustment_get_with_incharge(" & Val(chkExistsID) & "," & Val(UserID) & ",'" & txtTicketNo.Text & "')"
            With mainRs
                If .State = adStateOpen Then .Close
                .Open Qry, LYRcon, adOpenKeyset, adLockReadOnly
                If .RecordCount > 0 Then
                    .MoveFirst
                    If !level = "Granted" Then
                        txtScriptFileName.Text = "AUTO GENERATED"
                        txtMISDefinition.Text = !MIS_SCRIPT
                        myMISDifinitionScript = !MIS_SCRIPT
                        txtMISIncharge.Text = !mis_incharge
                        txtAuditApprvBy.Text = !AUDIT_INCHARGE
                    Else
                        MsgBox "THIS CONTROL NUMBER IS NOT APPLICABLE FOR THIS TICKETNO!", vbInformation + vbOKOnly, "Validation"
                        txtControlNo.Text = ""
                        txtMISDefinition.Text = ""
                        myMISDifinitionScript = ""
                        txtMISIncharge.Text = ""
                        txtAuditApprvBy.Text = ""
                        txtScriptFileName.Text = ""
                    End If
                End If
                .Close
            End With
        Else
            MsgBox "CONTROL NO. NOT FOUND!", vbInformation + vbOKOnly, "Validation"
            txtControlNo.Text = ""
            txtMISDefinition.Text = ""
            myMISDifinitionScript = ""
            txtMISIncharge.Text = ""
            txtAuditApprvBy.Text = ""
            txtScriptFileName.Text = ""
            Exit Sub
        End If
    Else
        MsgBox "THIS CONTROL NUMBER IS ALREADY EXISTS!", vbInformation + vbOKOnly, "Validation"
        txtControlNo.Text = ""
        txtMISDefinition.Text = ""
        myMISDifinitionScript = ""
        txtMISIncharge.Text = ""
        txtAuditApprvBy.Text = ""
        txtScriptFileName.Text = ""
        Exit Sub
    End If
End If
    myUpload = False
End Sub

Private Sub txtExplanation_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(Chr(KeyAscii))
    Call ObjKeyTrapping.Validation(txtExplanation, KeyAscii, ValidSQLChars)
    If KeyAscii = 13 Then
        txtMISDefinition.SetFocus
    End If
End Sub

Private Sub txtMISDefinition_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(Chr(KeyAscii))
    'Call ObjKeyTrapping.Validation(txtSubjRequest, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        txtApprvBy.SetFocus
    End If
End Sub

Private Sub txtMISIncharge_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtMISIncharge, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        txtAuditApprvBy.SetFocus
    End If
End Sub

Private Sub txtSubjRequest_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtSubjRequest, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        txtExplanation.SetFocus
    End If
End Sub
