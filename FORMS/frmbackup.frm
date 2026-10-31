VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmbackup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DB Utility"
   ClientHeight    =   2235
   ClientLeft      =   6255
   ClientTop       =   4800
   ClientWidth     =   3795
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   3795
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   1800
      Width           =   1695
      _ExtentX        =   2990
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
      Image           =   "frmbackup.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdbrow 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Select Location"
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
      Image           =   "frmbackup.frx":059A
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtloc 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   3495
   End
   Begin MSComDlg.CommonDialog cdgbLoc 
      Left            =   3480
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin lvButton.lvButtons_H cmdBackup 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "&Execute"
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
      Image           =   "frmbackup.frx":0B34
      cBack           =   -2147483633
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "BACKUP DATABASE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "BACKUP DATABASE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Location:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   660
   End
End
Attribute VB_Name = "frmbackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myLoc As String
Private Sub cmdBackup_Click()

Dim myLoc, mybats, myFile, myPath, rFile As String

 If txtLoc = "" Or Right$(Left$(txtLoc, 3), 1) <> "\" Then MsgBox ("Browse location.."), vbInformation: Exit Sub
    myLoc = txtLoc.Text
    mybats = "Backup.bat"
    myFile = Right(myLoc, InStr(1, StrReverse(myLoc), "\") - 1)
    myPath = Left(myLoc, Len(myLoc) - InStr(1, StrReverse(myLoc), "\")) & "\"
    rFile = myPath & "rubiulas" & Int(Format(Now, "mmddyyhhmm")) & ".SQL"
    
    Open mybats For Output As #1
    Print #1, "@echo off": Print #1, "echo Backup " & SysBranch & " Databases. Please wait....."
    If SysBranch = "MOTO-R REPUBLIC CORP." Then
        'Print #1, App.Path & "\mysqldump -h192.168.1.40 -uroot -plaborday SQLMTRTGM > " & myLoc
        Print #1, App.Path & "\mysqldump -h192.168.1.10 -uSERVER -plaborday SQLMTRTGM > " & myLoc
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - ULAS" Then
        Print #1, App.Path & "\mysqldump -h192.168.1.4 -uSERVER -plaborday SQLMTRULAS > " & myLoc
        Print #1, App.Path & "\mysqldump -h192.168.1.11 -userver -plaborday rubiulas > " & rFile
    ElseIf SysBranch = "MC PARTS DEPOT TORIL" Then
        Print #1, App.Path & "\mysqldump -h192.168.1.3 -uCASHIER -plaborday SQLMCPDEPOT > " & myLoc
        'Print #1, App.Path & "\mysqldump -uroot -plaborday SQLMCPDEPOT > " & myLoc
    ElseIf SysBranch = "LYR MOTOR SHOP AND SPAREPARTS DAVAO" Then
        Print #1, App.Path & "\mysqldump -uSERVER -plaborday SQLLYRSHOP > " & myLoc
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - VAN SELLING" Then
        Print #1, App.Path & "\mysqldump -h192.168.1.10 -uSERVER -plaborday SQLMTRVAN > " & myLoc
    ElseIf SysBranch = "MOTO-R REPUBLIC CORP. 2" Then
        Print #1, App.Path & "\mysqldump -uSERVER -plaborday SQLMRC2TGM > " & myLoc
    End If
    'Rar moto-r Ulas
    Print #1, App.Path & "\rar a -plaborday " & Left(txtLoc, Len(txtLoc) - 3) & "rar " & myFile
    'Rubi Ulas
    If SysBranch = "LYR MOTORPARTS CORP. - ULAS" Then Print #1, App.Path & "\rar a -plaborday " & Left(rFile, Len(rFile) - 3) & "rar " & "rubiulas" & Int(Format(Now, "mmddyyhhmm")) & ".SQL"
    Print #1, "del " & cdgbLoc.filename
    'Moto-R Ulas backup Rubi Ulas
    If SysBranch = "LYR MOTORPARTS CORP. - ULAS" Then Print #1, "del " & rFile
    Print #1, "echo Finish processing Backup"
    'Print #1, "pause"

    Print #1, "del " & mybats: Close #1: Shell myPath & mybats, vbNormalNoFocus
    
End Sub

Private Sub cmdbrow_Click()
    With cdgbLoc
        .Filter = "mySQL Script Files (*.SQL)|*.SQL"
        .CancelError = False
        .DialogTitle = "Save Backup"
        If SysBranch = "MOTO-R REPUBLIC CORP." Then
            .filename = "mtrtgm" & Int(Format(Now, "mmddyyhhmm"))
        ElseIf SysBranch = "LYR MOTORPARTS CORP. - ULAS" Then
            .filename = "mtrulas" & Int(Format(Now, "mmddyyhhmm"))
        ElseIf SysBranch = "MC PARTS DEPOT TORIL" Then
            .filename = "mcpDepot" & Int(Format(Now, "mmddyyhhmm"))
        ElseIf SysBranch = "LYR MOTOR SHOP AND SPAREPARTS DAVAO" Then
            .filename = "SHOP" & Int(Format(Now, "mmddyyhhmm"))
        ElseIf SysBranch = "LYR MOTORPARTS CORP. - VAN SELLING" Then
            .filename = "mtrvan" & Int(Format(Now, "mmddyyhhmm"))
        ElseIf SysBranch = "MOTO-R REPUBLIC CORP. 2" Then
            .filename = "mrc2tgm" & Int(Format(Now, "mmddyyhhmm"))
        End If
        .ShowOpen
        txtLoc.Text = .filename
    End With
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub

