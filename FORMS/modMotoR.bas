Attribute VB_Name = "modMotoR"
Global LYRcon       As New ADODB.Connection
Global EndPointCon       As New ADODB.Connection
Global MainCon      As New ADODB.Connection
Global DavaoWHCon   As New ADODB.Connection
Global OnlineCon    As New ADODB.Connection
Global BrchConMU    As New ADODB.Connection
Global BrchConYM    As New ADODB.Connection
Global BrchConMD    As New ADODB.Connection
Global BrchConVAN   As New ADODB.Connection
Global conUlas      As New ADODB.Connection
Global conPanabo    As New ADODB.Connection
Global conPol       As New ADODB.Connection
Global conToril     As New ADODB.Connection
Global conMRC2      As New ADODB.Connection
Global conVAN1      As New ADODB.Connection
Global conVAN2      As New ADODB.Connection
Global conVAN3      As New ADODB.Connection
Global conMRC1      As New ADODB.Connection
Global conYM        As New ADODB.Connection
Global conGenSan    As New ADODB.Connection
Global conStoTomas  As New ADODB.Connection
Global conSamal     As New ADODB.Connection
Global conMarbel    As New ADODB.Connection
Global rsAnyUse     As New ADODB.Recordset
Global conPartsCat As New ADODB.Connection

Global ConMWHLocal       As Boolean

Global UserID           As Integer
Global UserName         As String
Global FrDate           As Date
Global ToDate           As Date
Global DispBtn          As Boolean
Global UserType         As String
Global UserPosition     As String
Global ServerTime       As String
Global SysBranch        As String
Global sysBranchCode    As String
Global SysAddress       As String
Global SysTel           As String
Global SysBranchID      As Integer
Global SysBranchOIC     As String
Global SysBranchCategory    As String
Global SysDBNAME        As String
Global wCost            As Long
Global CurrDate         As String
Global LstDayMnth       As String
Global ServerDate       As String
Global COMMENTS As String
'********************* CODE UPDATED: 04/03/2017 ****************************
Public AllArrBankInfo() As String
Public AllArrBrand() As String
Public AllArrModel() As String
Public AllArrBranch() As String
Public ObjKeyTrapping As New clsKeyTrapping
Public AllArrCust() As String
'***************** END CODE UPDATED: 04/03/2017 ****************************

'report function declaration
Declare Function CreateFieldDefFile Lib "p2smon.dll" (lpUnk As Object, ByVal filename As String, ByVal bOverWriteExistingFile As Long) As Long

'variable for saving/updating pictures
Global Chunk() As Byte
Global lngLengh As Long
Global intChunks As Integer
Global intFragment As Integer
'Global Const ChunkSize = 1000 '16384
Global Const ChunkSize = 16384

Global Const lngDataFile = 1
'searching variables

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Global Const CB_FINDSTRING = &H14C
Declare Function CharUpper Lib "user32" Alias "CharUpperA" (ByVal lpsz As String) As String
'user variables

Global isEndpoint As Boolean

Sub Main()
    On Error GoTo LocalErr
   ' ConMWHLocal = True 'para sa mga branches na naay internet ang connection
    'ConMWHLocal = False 'para sa mga branches na dili internet ang connection
    isEndpoint = False
    
    LYRcon.CursorLocation = adUseClient
    LYRcon.ConnectionTimeout = 60
    
    'Access
    'LYRcon.Open "dsn=MTREPTGM;uid=;pwd=LCRJR;database=MTREPTGM"
    'mySQL
    
    
    LYRcon.Open "dsn=branch" 'ESTABLISH CONNECTION VIA ODBC BASED
    
    
    'Detect Command Arguement
'   Branch          = 0     192.168.1.11
'   HQ              = 1     192.168.30.183
'   COPR            = 2     192.168.30.29
'   BRANCHSERVER    = 3     localhost
'   OLDBRANCH       = 4     192.168.1.10
'   VMDev           = 5     10.10.10.1

    'Call modConnection.ParseCommandArgument("13&mrt&192.168.30.183")        'Remove Remarks here & setMDIBackGroud, Add FormHoldBackground

    'Establish Database Connection
   ' Call modConnection.DBMainConnection(LYRcon) 'COMMENT OUT IF ODBC BASED UNCOMMENT IF PARAMETER BASED
    
    'Set Branch Setting
    Call SBranch

    'Set MotoRep BackGround and Icon
    'Call modConnection.setMDIBackGround

    'Call SBranch
    
    MotoRep.Show
    LogIn.Show vbModal
    'frmInvCtrlBPO.Show
    'frmSplash.Show vbModal
   Exit Sub

LocalErr:
    isEndpoint = True
    EndPointCon.Open "dsn=endpoint"
    MotoRep.Show
    LogInWithEnpoint.Show vbModal
   
End Sub

Public Sub ConnectOnline()
    On Error GoTo LocalErr

    If OnlineCon.State = adStateOpen Then OnlineCon.Close
    OnlineCon.CursorLocation = adUseClient
    OnlineCon.ConnectionTimeout = 1800
    
    'OnlineCon.Open "dsn=sqlmtrtgmonline"
    'OnlineCon.Open "dsn=SQLMTRULASOnline"
    'OnlineCon.Open "dsn=SQLMCPDEPOTOnline"
    'OnlineCon.Open "dsn=sqlmrc2tgmonline"   'WHOLESALE
    'OnlineCon.Open "dsn=mrcpanaboOnline"
    'OnlineCon.Open "dsn=mrcgensanOnline"
    'OnlineCon.Open "dsn=mrcstotomasOnline"
    'OnlineCon.Open "dsn=SQLMTRVANOnline"
    'OnlineCon.Open "dsn=mrcvan2Online"
    'OnlineCon.Open "dsn=mrcvan3Online"
    'OnlineCon.Open "dsn=mrcsamalOnline"
    'OnlineCon.Open "dsn=mrcmarbelOnline"
    'OnlineCon.Open "dsn=mrc2retailonline"
    'OnlineCon.Open "dsn=mrcpolomolokonline"
    'OnlineCon.Open "dsn=mrcpiaggiotgmonline"
    'OnlineCon.Open "dsn=mrckapalongonline"
    'OnlineCon.Open "dsn=mrcmankilamonline"
    'OnlineCon.Open "dsn=mrcstotomasonline"
    'OnlineCon.Open "dsn=mrcmotorhubulasonline"
    'OnlineCon.Open "dsn=promotionalonline"
    OnlineCon.Open modConnection.getOnlineConnectionString()
    Exit Sub
LocalErr:
    MsgBox ("Error connecting to " & SysBranch & " online @ MAIN OFFICE SERVER...."), vbExclamation

End Sub

Public Sub ConnectMain()

    On Error GoTo LocalErr
    If MainCon.State = adStateOpen Then: MainCon.Close
    MainCon.CursorLocation = adUseClient
    MainCon.ConnectionTimeout = 1800
    MainCon.Open "dsn=directmrcwarehouse"           '"dsn=SQLMRC2TGM"
    'MainCon.Open modConnection.getWarehouseConnectionString()
    Exit Sub

LocalErr:
    MsgBox ("Error connecting to DIRECTMRCWAREHOUSE, make sure to configure odbc correctly...."), vbExclamation
    
End Sub

Public Sub ConnectDavaoWH()

    On Error GoTo LocalErr
    If DavaoWHCon.State = adStateOpen Then: DavaoWHCon.Close
    DavaoWHCon.CursorLocation = adUseClient
    DavaoWHCon.ConnectionTimeout = 1800
    'DavaoWHCon.Open "dsn=mrcwarehouseagdao"           '"dsn=SQLMRC2TGM"
    DavaoWHCon.Open modConnection.getDavaoWarehouseConnectionString()
    Exit Sub

LocalErr:
    MsgBox ("Error connecting to LYRMP WAREHOUSE @ AGDAO DAVAO CITY Database...."), vbExclamation
    
End Sub


Public Sub CenterFrm(ByVal Parentfrm As MDIForm, ByVal Childfrm As Form)
    Childfrm.Left = (Parentfrm.Width \ 2) - (Childfrm.Width \ 2)
    Childfrm.Top = (Parentfrm.Height \ 2) - (Childfrm.Height \ 2) - 600
End Sub

Public Function MyPaperSize() As Integer
    Dim MyPrn As Printer
    Dim i As Integer
    'Dim doctoprint As New System.Drawing.Printing.PrintDocument
    
    'MyPaperSize = MyPrn.PaperSize
    
   '         doctoprint.PrinterSettings.PrinterName = "YourPrinterName" '(ex. "Epson SQ-1170 ESC/P 2")
            For i = 0 To doctoprint.PrinterSettings.PaperSizes.Count - 1
                Dim rawKind As Integer
                If doctoprint.PrinterSettings.PaperSizes(i).PaperName = "ADEL 8.5X5.5" Then
                    rawKind = CInt(doctoprint.PrinterSettings.PaperSizes(i).GetType().GetField("ADEL 8.5X5.5", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic).GetValue(doctoprint.PrinterSettings.PaperSizes(i)))
                    MyPaperSize = rawKind
                    Exit Function
               End If
            Next
     MyPaperSize = 0
End Function

Public Sub PrintMyForm(ByRef MyVal() As String, FormName As String, ByVal MyFlx As MSFlexGrid)
    Dim rsPref As New ADODB.Recordset
    Dim mcLyr As CRAXDRT.Application
    Dim mcReport As CRAXDRT.Report

    Dim Qry As String, z As Integer
    Dim i As Integer, j As Integer, NxtRow As Integer, InsertHere As Integer
    Dim MyDesc() As String, myTop() As String, myLeft() As String, myLen() As String
    Dim myHeight() As String, myFsize() As String, myFbold() As String, myAlign() As String
    Dim mySubLeft() As String, mySubLen() As String, mySubBold() As String, mySubAlign() As String
    Dim txtObj As TextObject
    
    Set mcLyr = New CRAXDRT.Application
    Set mcReport = mcLyr.NewReport
    
    'Access
    'Qry = "SELECT PRINTPREF.* From PRINTPREF " & _
          "WHERE (((PRINTPREF.FORMNAME)='" & FormName & "'));"
    'mySQL
    Qry = "SELECT PRINTPREF.* From PRINTPREF " & _
          "WHERE PRINTPREF.FORMNAME='" & FormName & "';"
    
    rsPref.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsPref.RecordCount > 0 Then
        rsPref.MoveFirst
        'Assign data to variables
        MyDesc = Split(rsPref!Description, ",")
        myTop = Split(rsPref!Top, ",")
        myLeft = Split(rsPref!Left, ",")
        myLen = Split(rsPref!Length, ",")
        myHeight = Split(rsPref!Height, ",")
        myFsize = Split(rsPref!FontSize, ",")
        myFbold = Split(rsPref!FontBold, ",")
        myAlign = Split(rsPref!Alignment, ",")
        InsertHere = IIf(rsPref!INSERTINDEX <> 0, rsPref!INSERTINDEX, -1)
        'Start Printing here
        If FormName = "INVOICE" Or FormName = "REQUISITION" Then
            mcReport.PaperSize = crDefaultPaperSize 'crPaperEnvelopeB6 changed at 2024-09-10
        ElseIf FormName = "JOB ORDER" Then
            mcReport.PaperSize = crDefaultPaperSize
        Else
            mcReport.PaperSize = crDefaultPaperSize
        End If
        mcReport.TopMargin = 0
        mcReport.BottomMargin = 0
        mcReport.Sections(1).Height = 0
        mcReport.Sections(2).Height = 0
        With mcReport.Sections(3)
            .Height = rsPref!FORMHEIGHT
            For i = LBound(MyDesc) To UBound(MyDesc)
                If i = InsertHere Then
                    NxtRow = rsPref!LOOPSTARTTOP
                    mySubLeft = Split(rsPref!SUBLEFT, ",")
                    mySubLen = Split(rsPref!SUBLEN, ",")
                    mySubBold = Split(rsPref!SUBFBOLD, ",")
                    mySubAlign = Split(rsPref!SUBALIGN, ",")
                    If FormName = "Promissory Note" Then z = 10
                    If FormName = "RECEIPT-CR" Then
                        If MyFlx.Rows > 10 Then
                            For j = 1 To 5
                                For k = LBound(mySubLeft) To UBound(mySubLeft)
                                    Set txtObj = .AddTextObject(MyFlx.TextMatrix(j, k + 1), mySubLeft(k), NxtRow)
                                    txtObj.Width = Val(mySubLen(k)): txtObj.HorAlignment = Val(mySubAlign(k)): txtObj.Font.Name = "Arial": txtObj.Font.Size = Val(rsPref!LOOPFSIZE): txtObj.Height = Val(rsPref!LOOPHEIGHT)
                                Next k
                                If FormName = "Promissory Note" Then    'Control rows in Promi
                                    NxtRow = NxtRow + rsPref!LOOPINTERVAL - z
                                    z = IIf(z = 70, 10, z + 10)
                                Else
                                    NxtRow = NxtRow + rsPref!LOOPINTERVAL
                                End If
                            Next j
                        Else
                            For j = 1 To MyFlx.Rows - 1
                                For k = LBound(mySubLeft) To UBound(mySubLeft)
                                    Set txtObj = .AddTextObject(MyFlx.TextMatrix(j, k + 1), mySubLeft(k), NxtRow)
                                    txtObj.Width = Val(mySubLen(k)): txtObj.HorAlignment = Val(mySubAlign(k)): txtObj.Font.Name = "Arial": txtObj.Font.Size = Val(rsPref!LOOPFSIZE): txtObj.Height = Val(rsPref!LOOPHEIGHT)
                                Next k
                                If FormName = "Promissory Note" Then    'Control rows in Promi
                                    NxtRow = NxtRow + rsPref!LOOPINTERVAL - z
                                    z = IIf(z = 70, 10, z + 10)
                                Else
                                    NxtRow = NxtRow + rsPref!LOOPINTERVAL
                                End If
                            Next j
                        End If
                    Else
                        For j = 1 To MyFlx.Rows - 1
                            For k = LBound(mySubLeft) To UBound(mySubLeft)
                                Set txtObj = .AddTextObject(MyFlx.TextMatrix(j, k + 1), mySubLeft(k), NxtRow)
                                'if (MyFlx.TextMatrix(j, k + 1) <> "0")
                                txtObj.Width = Val(mySubLen(k)): txtObj.HorAlignment = Val(mySubAlign(k)): txtObj.Font.Name = "Arial": txtObj.Font.Size = Val(rsPref!LOOPFSIZE): txtObj.Height = Val(rsPref!LOOPHEIGHT) 'pc content
                            Next k
                            If FormName = "Promissory Note" Then    'Control rows in Promi
                                NxtRow = NxtRow + rsPref!LOOPINTERVAL - z
                                z = IIf(z = 70, 10, z + 10)
                            Else
                                NxtRow = NxtRow + rsPref!LOOPINTERVAL
                            End If
                        Next j
                    End If
                End If
                Set txtObj = .AddTextObject(MyVal(i), Val(myLeft(i)), Val(myTop(i)))
                txtObj.Width = Val(myLen(i)): txtObj.HorAlignment = Val(myAlign(i)): txtObj.Font.Name = "Arial": txtObj.Font.Size = Val(myFsize(i)): txtObj.Height = Val(myHeight(i)): txtObj.Font.Bold = Val(myFbold(i))
            Next i
        End With
        mcReport.Sections(4).Height = 0
        mcReport.Sections(5).Height = 0
        
        mcReport.PrintOut False, 1, , 1, 1
    Else
        MsgBox "No Printing Preference found!", vbCritical + vbOKOnly, "Unknown Form Name"
    End If
    rsPref.Close
End Sub



Public Sub PrintMyFormNew(ByRef MyVal() As String, FormName As String, ByVal MyFlx As MSHFlexGrid)
    Dim rsPref As New ADODB.Recordset
    Dim mcLyr As CRAXDRT.Application
    Dim mcReport As CRAXDRT.Report

    Dim Qry As String, z As Integer
    Dim i As Integer, j As Integer, NxtRow As Integer, InsertHere As Integer
    Dim MyDesc() As String, myTop() As String, myLeft() As String, myLen() As String
    Dim myHeight() As String, myFsize() As String, myFbold() As String, myAlign() As String
    Dim mySubLeft() As String, mySubLen() As String, mySubBold() As String, mySubAlign() As String
    Dim txtObj As TextObject
    
    Set mcLyr = New CRAXDRT.Application
    Set mcReport = mcLyr.NewReport
    
    'Access
    'Qry = "SELECT PRINTPREF.* From PRINTPREF " & _
          "WHERE (((PRINTPREF.FORMNAME)='" & FormName & "'));"
    'mySQL
    Qry = "SELECT PRINTPREF.* From PRINTPREF " & _
          "WHERE PRINTPREF.FORMNAME='" & FormName & "';"
    
    rsPref.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsPref.RecordCount > 0 Then
        rsPref.MoveFirst
        'Assign data to variables
        MyDesc = Split(rsPref!Description, ",")
        myTop = Split(rsPref!Top, ",")
        myLeft = Split(rsPref!Left, ",")
        myLen = Split(rsPref!Length, ",")
        myHeight = Split(rsPref!Height, ",")
        myFsize = Split(rsPref!FontSize, ",")
        myFbold = Split(rsPref!FontBold, ",")
        myAlign = Split(rsPref!Alignment, ",")
        InsertHere = IIf(rsPref!INSERTINDEX <> 0, rsPref!INSERTINDEX, -1)
        'Start Printing here
        If FormName = "INVOICE" Or FormName = "REQUISITION" Then
            mcReport.PaperSize = crDefaultPaperSize 'crPaperEnvelopeB6 changed at 2024-09-10
        ElseIf FormName = "JOB ORDER" Then
            mcReport.PaperSize = crDefaultPaperSize
        Else
            mcReport.PaperSize = crDefaultPaperSize
        End If
        mcReport.TopMargin = 0
        mcReport.BottomMargin = 0
        mcReport.Sections(1).Height = 0
        mcReport.Sections(2).Height = 0
        With mcReport.Sections(3)
            .Height = rsPref!FORMHEIGHT
            For i = LBound(MyDesc) To UBound(MyDesc)
                If i = InsertHere Then
                    NxtRow = rsPref!LOOPSTARTTOP
                    mySubLeft = Split(rsPref!SUBLEFT, ",")
                    mySubLen = Split(rsPref!SUBLEN, ",")
                    mySubBold = Split(rsPref!SUBFBOLD, ",")
                    mySubAlign = Split(rsPref!SUBALIGN, ",")
                    If FormName = "Promissory Note" Then z = 10
                    If FormName = "RECEIPT-OR" Then
                        If MyFlx.Rows > 5 Then
                            For j = 1 To 5
                                For k = LBound(mySubLeft) To UBound(mySubLeft)
                                    Set txtObj = .AddTextObject(MyFlx.TextMatrix(j, k + 1), mySubLeft(k), NxtRow)
                                    txtObj.Width = Val(mySubLen(k)): txtObj.HorAlignment = Val(mySubAlign(k)): txtObj.Font.Name = "Arial": txtObj.Font.Size = Val(rsPref!LOOPFSIZE): txtObj.Height = Val(rsPref!LOOPHEIGHT)
                                Next k
                                If FormName = "Promissory Note" Then    'Control rows in Promi
                                    NxtRow = NxtRow + rsPref!LOOPINTERVAL - z
                                    z = IIf(z = 70, 10, z + 10)
                                Else
                                    NxtRow = NxtRow + rsPref!LOOPINTERVAL
                                End If
                            Next j
                        Else
                            For j = 1 To MyFlx.Rows - 1
                                For k = LBound(mySubLeft) To UBound(mySubLeft)
                                    If k = 3 Then
                                        If (MyFlx.TextMatrix(j, k + 1) <> "0") Then
                                            Set txtObj = .AddTextObject("(" & MyFlx.TextMatrix(j, k + 1) & "% Discount)", mySubLeft(k), NxtRow)
                                            txtObj.Width = Val(mySubLen(k)): txtObj.HorAlignment = Val(mySubAlign(k)): txtObj.Font.Name = "Arial": txtObj.Font.Size = Val(rsPref!LOOPFSIZE): txtObj.Height = Val(rsPref!LOOPHEIGHT)
                                        Else
                                            Set txtObj = .AddTextObject("", mySubLeft(k), NxtRow)
                                            txtObj.Width = Val(mySubLen(k)): txtObj.HorAlignment = Val(mySubAlign(k)): txtObj.Font.Name = "Arial": txtObj.Font.Size = Val(rsPref!LOOPFSIZE): txtObj.Height = Val(rsPref!LOOPHEIGHT)
                                        End If
                                    Else
                                        Set txtObj = .AddTextObject(MyFlx.TextMatrix(j, k + 1), mySubLeft(k), NxtRow)
                                        txtObj.Width = Val(mySubLen(k)): txtObj.HorAlignment = Val(mySubAlign(k)): txtObj.Font.Name = "Arial": txtObj.Font.Size = Val(rsPref!LOOPFSIZE): txtObj.Height = Val(rsPref!LOOPHEIGHT)
                                    End If
                                Next k
                                If FormName = "Promissory Note" Then    'Control rows in Promi
                                    NxtRow = NxtRow + rsPref!LOOPINTERVAL - z
                                    z = IIf(z = 70, 10, z + 10)
                                Else
                                    NxtRow = NxtRow + rsPref!LOOPINTERVAL
                                End If
                            Next j
                        End If
                    

'                   ElseIf FormName = "RECEIPT" Then
'                        If MyFlx.Rows > 10 Then
'                            For j = 1 To 5
'                                For k = LBound(mySubLeft) To UBound(mySubLeft)
'                                    Set txtObj = .AddTextObject(MyFlx.TextMatrix(j, k + 1), mySubLeft(k), NxtRow)
'                                    txtObj.Width = Val(mySubLen(k)): txtObj.HorAlignment = Val(mySubAlign(k)): txtObj.Font.Name = "Arial": txtObj.Font.Size = Val(rsPref!LOOPFSIZE): txtObj.Height = Val(rsPref!LOOPHEIGHT)
'                                Next k
'                                If FormName = "Promissory Note" Then    'Control rows in Promi
'                                    NxtRow = NxtRow + rsPref!LOOPINTERVAL - z
'                                    z = IIf(z = 70, 10, z + 10)
'                                Else
'                                    NxtRow = NxtRow + rsPref!LOOPINTERVAL
'                                End If
'                            Next j
'                        Else
'                            For j = 1 To MyFlx.Rows - 1
'                                For k = LBound(mySubLeft) To UBound(mySubLeft)
'                                    Set txtObj = .AddTextObject(MyFlx.TextMatrix(j, k + 1), mySubLeft(k), NxtRow)
'                                    txtObj.Width = Val(mySubLen(k)): txtObj.HorAlignment = Val(mySubAlign(k)): txtObj.Font.Name = "Arial": txtObj.Font.Size = Val(rsPref!LOOPFSIZE): txtObj.Height = Val(rsPref!LOOPHEIGHT)
'                                Next k
'                                If FormName = "Promissory Note" Then    'Control rows in Promi
'                                    NxtRow = NxtRow + rsPref!LOOPINTERVAL - z
'                                    z = IIf(z = 70, 10, z + 10)
'                                Else
'                                    NxtRow = NxtRow + rsPref!LOOPINTERVAL
'                                End If
'                            Next j
'                        End If
                    Else
                    
                        For j = 1 To MyFlx.Rows - 1
                            For k = LBound(mySubLeft) To UBound(mySubLeft)
                                Set txtObj = .AddTextObject(MyFlx.TextMatrix(j, k + 1), mySubLeft(k), NxtRow)
                                txtObj.Width = Val(mySubLen(k)): txtObj.HorAlignment = Val(mySubAlign(k)): txtObj.Font.Name = "Arial": txtObj.Font.Size = Val(rsPref!LOOPFSIZE): txtObj.Height = Val(rsPref!LOOPHEIGHT) 'pc content
                            Next k
                            If FormName = "Promissory Note" Then    'Control rows in Promi
                                NxtRow = NxtRow + rsPref!LOOPINTERVAL - z
                                z = IIf(z = 70, 10, z + 10)
                            Else
                                NxtRow = NxtRow + rsPref!LOOPINTERVAL
                            End If
                        Next j
                    End If
                End If
                Set txtObj = .AddTextObject(MyVal(i), Val(myLeft(i)), Val(myTop(i)))
                txtObj.Width = Val(myLen(i)): txtObj.HorAlignment = Val(myAlign(i)): txtObj.Font.Name = "Arial": txtObj.Font.Size = Val(myFsize(i)): txtObj.Height = Val(myHeight(i)): txtObj.Font.Bold = Val(myFbold(i))
            Next i
        End With
        mcReport.Sections(4).Height = 0
        mcReport.Sections(5).Height = 0
        
        mcReport.PrintOut False, 1, , 1, 1
    Else
        MsgBox "No Printing Preference found!", vbCritical + vbOKOnly, "Unknown Form Name"
    End If
    rsPref.Close
End Sub

Function DTCpy(MyQry As String)
    Dim rsPrn As New ADODB.Recordset
    Dim myCnt As Integer, Tmp As String
    
    With rsPrn
        .Open MyQry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            myCnt = IIf(IsNull(.Fields(0)), 0, .Fields(0)) + 1
            .Fields(0) = myCnt
            .Update
        End If
        .Close
    End With
    If myCnt = 1 Then
        DTCpy = Format(Now, "mm/dd/yy hh:mm ampm")
    Else
        If myCnt = 2 Then
            Tmp = "nd"
        ElseIf myCnt = 3 Then
            Tmp = "rd"
        Else
            Tmp = "th"
        End If
        DTCpy = Format(Now, "mm/dd/yy hh:mm ampm") & " " & myCnt & Tmp & " copy"
    End If
End Function

Public Function RepAlign(aStr As String) As Integer
    Select Case aStr
        Case "crLeftAlign"
            RepAlign = 3
        Case "crHorCenterAlign"
            RepAlign = 1
        Case "crRightAlign"
            RepAlign = 4
    End Select
End Function

Public Function PureNum(nStr As String) As Single
    Dim Tmp As String
    Dim j As Integer
    
    Tmp = nStr
    If InStr(Tmp, "P") <> 0 Then
        Tmp = Right$(Tmp, Len(Tmp) - 1)
    End If
    If Tmp = "" Then
        PureNum = 0
        Exit Function
    End If
    Do While Not InStr(Tmp, ",") = 0
        j = InStr(Tmp, ",")
        Tmp = Left$(Tmp, j - 1) & Mid$(Tmp, j + 1, Len(Tmp) - j)
    Loop
    
    PureNum = Tmp
End Function

Public Function CtrlNumExist(Qry As String) As Boolean
    Dim rsExist As New ADODB.Recordset
    
    With rsExist
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            CtrlNumExist = True
        Else
            CtrlNumExist = False
        End If
        .Close
    End With
End Function

Public Function GetUserName() As String
    Dim rsEmp As New ADODB.Recordset
    Dim Qry As String
    
    '****Access
    'Qry = "SELECT [LASTNAME] & ', ' & [FIRSTNAME] & ' ' & Left$([MIDDLENAME],1) & '.' AS USERNAME, EMPLOYEES.EMPID " & _
          "From EMPLOYEES WHERE (((EMPLOYEES.EMPID)=" & UserID & "));"
    'mySQL
    Qry = "SELECT concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.'))) AS USERNAME " & _
          "From EMPLOYEES WHERE EMPLOYEES.EMPID=" & UserID & ";"
    
    With rsEmp
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetUserName = IIf(IsNull(.Fields(0)), "", .Fields(0))
        Else
            GetUserName = ""
        End If
        .Close
    End With
End Function

Public Function GetEmployeeName() As String
    Dim rsEmp As New ADODB.Recordset
    Dim Qry As String
    
    '****Access
    'Qry = "SELECT [LASTNAME] & ', ' & [FIRSTNAME] & ' ' & Left$([MIDDLENAME],1) & '.' AS USERNAME, EMPLOYEES.EMPID " & _
          "From EMPLOYEES WHERE (((EMPLOYEES.EMPID)=" & UserID & "));"
    'mySQL
    Qry = "SELECT concat(EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.')),' ',EMPLOYEES.LASTNAME) AS EMPLOYEENAME " & _
          "From EMPLOYEES WHERE EMPLOYEES.EMPID=" & UserID & ";"
    
    With rsEmp
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetEmployeeName = IIf(IsNull(.Fields(0)), "", .Fields(0))
        Else
            GetEmployeeName = ""
        End If
        .Close
    End With
End Function



Public Function GetSettings(MyRemarks As String) As String
    Dim rsSettings As New ADODB.Recordset
    Dim Qry As String
    '****Access
    'Qry = "SELECT SETTINGS.VALUE, SETTINGS.DESCRIPTION, SETTINGS.REMARKS From SETTINGS " & _
          "WHERE (((SETTINGS.DESCRIPTION)='EMP DEFAULT VALUE') AND ((SETTINGS.REMARKS)='" & MyRemarks & "'));"
    
    'mySQL
    Qry = "SELECT SETTINGS.VALUE, SETTINGS.DESCRIPTION, SETTINGS.REMARKS From SETTINGS " & _
          "WHERE SETTINGS.DESCRIPTION='EMP DEFAULT VALUE' AND SETTINGS.REMARKS='" & MyRemarks & "';"
    
    With rsSettings
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            GetSettings = IIf(IsNull(!Value), "", !Value)
        Else
            GetSettings = ""
        End If
        .Close
    End With
End Function

Public Function GetFieldValue(MyQ As String) As Variant
    Dim rsAll As New ADODB.Recordset
    
    With rsAll
        .Open MyQ, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetFieldValue = IIf(IsNull(.Fields(0)), "", .Fields(0))
        Else
            GetFieldValue = ""
        End If
        .Close
    End With
        
End Function


Public Function GetFieldValueMain(MyQ As String) As Variant
    Dim rsAll As New ADODB.Recordset
    
    With rsAll
        .Open MyQ, MainCon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetFieldValueMain = IIf(IsNull(.Fields(0)), "", .Fields(0))
        Else
            GetFieldValueMain = ""
        End If
        .Close
    End With
        
End Function

Public Function GetFieldValueDVOWH(MyQ As String) As Variant
    Dim rsAll As New ADODB.Recordset
    
    With rsAll
        .Open MyQ, DavaoWHCon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetFieldValueDVOWH = IIf(IsNull(.Fields(0)), "", .Fields(0))
        Else
            GetFieldValueDVOWH = ""
        End If
        .Close
    End With
        
End Function

Public Sub UserAccess()
    Dim rsUA As New ADODB.Recordset
    Dim Qry As String, Ndx As Integer, IDx As Integer
    Dim mnuX As Menu
    
    'Disable all submenu
    '****Access
    'Qry = "SELECT USERACCESS.* From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & "));"
    
    'mySQL
    Qry = "SELECT USERACCESS.* From UserAccess WHERE USERACCESS.EMPID=" & UserID & " ORDER BY MENUINDX DESC;"
    
    rsUA.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsUA.RecordCount > 0 Then
        rsUA.MoveFirst
        Do While Not rsUA.EOF
            Ndx = rsUA!MenuIndx
            IDx = rsUA!MenuID
            Select Case Ndx
                Case 8 To 12, 118
                    For Each mnuX In MotoRep.mSubPI
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 13 To 17
                    For Each mnuX In MotoRep.mSubSP
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 18 To 20, 97, 102, 114, 115
                    'If Ndx = 18 Then MotoRep.BtnMain(1).Enabled = True
                    For Each mnuX In MotoRep.subOrder
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 21 To 24, 108, 109, 110
                    For Each mnuX In MotoRep.subReceive
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 25 To 29
                    'If Ndx = 25 Then MotoRep.BtnMain(2).Enabled = True
                    For Each mnuX In MotoRep.subSales
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 30 To 34, 106, 107
                    'If Ndx = 30 Then MotoRep.BtnMain(0).Enabled = True
                    For Each mnuX In MotoRep.SPsubStocks
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 35 To 37, 113
                    For Each mnuX In MotoRep.SPsubADJ
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 38 To 43
                    For Each mnuX In MotoRep.SubSRVC
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 44 To 50, 116
                    For Each mnuX In MotoRep.mSubTR
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 51 To 56, 111, 126, 127
                    For Each mnuX In MotoRep.subAdj
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 57 To 60
                    For Each mnuX In MotoRep.subDisb
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 61 To 63, 98, 99
                    'If Ndx = 62 Then MotoRep.BtnMain(5).Enabled = True
                    For Each mnuX In MotoRep.subCashFlow
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 64 To 68, 105
                    'If Ndx = 67 Then MotoRep.BtnMain(3).Enabled = True
                    For Each mnuX In MotoRep.subAR
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 69 To 72, 119
                    For Each mnuX In MotoRep.mSubRep
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 73 To 75, 117
                    'If Ndx = 73 Then MotoRep.BtnMain(4).Enabled = True
                    For Each mnuX In MotoRep.subRptDailySales
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 76 To 81, 100
                    For Each mnuX In MotoRep.subRptSP
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 82 To 84
                    For Each mnuX In MotoRep.subRptSrv
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 85 To 89, 101, 112
                    For Each mnuX In MotoRep.SubRptMonthly
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 90 To 93
                    For Each mnuX In MotoRep.mSubTools
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 120 To 124
                    For Each mnuX In MotoRep.mnuAcctReport
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 94
                    MotoRep.AudTools(1).Enabled = True
                Case 95
                    MotoRep.mSetPrn.Enabled = True
                Case 96
                    MotoRep.mSetOpt.Enabled = True
            End Select
            rsUA.MoveNext
        Loop
    End If
End Sub

Public Sub userPrivelege()
'*************** CODE UPDATED 11/27/17 *****************
    Qry = "Select EMPID, MENUINDX " & _
            "From useraccess Where useraccess.EMPID = " & UserID & ";"
    If rsAnyUse.State <> 0 Then rsAnyUse.Close
    Set rsAnyUse = LYRcon.Execute(Qry)
    If rsAnyUse.RecordCount > 0 Then
        rsAnyUse.MoveFirst
            
            Do While Not rsAnyUse.EOF
                If rsAnyUse!MenuIndx = 25 Then
                    MotoRep.toolBarButton.Buttons(1).Enabled = True
                End If
                If rsAnyUse!MenuIndx = 21 Then
                    MotoRep.toolBarButton.Buttons(2).Enabled = True
                End If
                If rsAnyUse!MenuIndx = 108 Then
                    MotoRep.toolBarButton.Buttons(3).Enabled = True
                End If
                If rsAnyUse!MenuIndx = 26 Then
                    MotoRep.toolBarButton.Buttons(4).Enabled = True
                End If
                If rsAnyUse!MenuIndx = 35 Then
                    MotoRep.toolBarButton.Buttons(5).Enabled = True
                End If
                If rsAnyUse!MenuIndx = 112 Then
                    MotoRep.toolBarButton.Buttons(6).Enabled = True
                End If
                rsAnyUse.MoveNext
            Loop
'            If rsAnyUse!MOTOR <> 1 Then
'                MotoRep.mnuPI.Enabled = False
'                MotoRep.mnuSP.Enabled = False
'                MotoRep.mnuSrvc.Enabled = False
'                MotoRep.mnuTRN.Enabled = False
'                MotoRep.mnuReports.Enabled = False
'                MotoRep.mnuTools.Enabled = False
'            End If
'*********** END CODE UPDATED 11/27/17 *****************
    End If
End Sub

Public Function ComputerName() As String
    Dim Cn As String

    Cn = Space(25)
    GetComputerName Cn, 25
    
    ComputerName = Left$(RTrim(Cn), Len(RTrim(Cn)) - 1)
End Function

Public Sub SBranch()
    Dim rsBrn As New ADODB.Recordset
    Dim mBrc As String
   If isEndpoint Then
       mBrc = "Select `server`,`port`,`database`,`user`,`password` from endpoint where branchname='" & LogInWithEnpoint.cboBranch.Text & "' ORDER BY BRANCHNAME ASC;"
        With rsBrn
            .Open mBrc, EndPointCon, adOpenKeyset, adLockOptimistic
            If .RecordCount + 2 > 0 Then
                .MoveFirst
                connectionString = "Driver={MySQL ODBC 3.51 Driver};Server=" & !server & ";Port=" & !Port & ";Database=" & _
            !Database & ";User=" & !User & ";Password=" & !Password & ";Option=3;"
            
            End If
            .Close
        End With
        If connectionString = "" Then
            MsgBox ("No connection string was found on endpoint database"), vbExclamation
            Exit Sub
        End If
        
        Call modConnection.DBMainConnect(LYRcon)
        
        mBrc = "SELECT  Curdate() as ServerDate, date_format(curdate(), '%M %d, %Y') as serTime , branch.BRANCHNAME, branch.branchid, Branch.Address, branch.CONTACTNOS, branch.manager, branch.BRANCHCODE, branchcategory From Branch Where Branch.SYSTEMBRANCH = 1;"
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                ServerTime = !sertime
                SysBranchID = Val(!BRANCHID)
                SysBranch = !branchName
                MotoRep.Caption = !branchName
                SysAddress = !Address
                SysBranchOIC = !MANAGER
                sysBranchCode = !branchCode
                SysTel = IIf(IsNull(!CONTACTNOS), "", !CONTACTNOS)
                SysBranchCategory = !branchcategory
                ServerDate = !ServerDate
            End If
            .Close
        End With
   
   Else
            mBrc = "SELECT  Curdate() as ServerDate, date_format(curdate(), '%M %d, %Y') as serTime , branch.BRANCHNAME, branch.branchid, Branch.Address, branch.CONTACTNOS, branch.manager, branch.BRANCHCODE, branchcategory From Branch Where Branch.SYSTEMBRANCH = 1;"
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                ServerTime = !sertime
                SysBranchID = Val(!BRANCHID)
                SysBranch = !branchName
                 
                SysAddress = !Address
                SysBranchOIC = !MANAGER
                sysBranchCode = !branchCode
                SysTel = IIf(IsNull(!CONTACTNOS), "", !CONTACTNOS)
                SysBranchCategory = !branchcategory
                ServerDate = !ServerDate
            End If
            .Close
        End With
    End If
End Sub

Public Function GetFieldValueMU(MyQ As String) As Variant
    Dim rsBrch As New ADODB.Recordset
    
    With rsBrch
        .Open MyQ, conUlas, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetFieldValueMU = IIf(IsNull(.Fields(0)), "", .Fields(0))
        Else
            GetFieldValueMU = ""
        End If
        .Close
    End With
        
End Function

Public Function GetFieldValueVAN(MyQ As String) As Variant
    Dim rsBrch As New ADODB.Recordset

    With rsBrch
        .Open MyQ, conVAN1, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetFieldValueVAN = IIf(IsNull(.Fields(0)), "", .Fields(0))
        Else
            GetFieldValueVAN = ""
        End If
        .Close
    End With
        
End Function

Public Function GetFieldValueMRC2(MyQ As String) As Variant
    Dim rsBrch As New ADODB.Recordset

    With rsBrch
        .Open MyQ, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetFieldValueMRC2 = IIf(IsNull(.Fields(0)), "", .Fields(0))
        Else
            GetFieldValueMRC2 = ""
        End If
        .Close
    End With
        
End Function

Public Function GetFieldValueYM(MyQ As String) As Variant
    Dim rsBrch As New ADODB.Recordset
    
    With rsBrch
        .Open MyQ, conYM, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetFieldValueYM = IIf(IsNull(.Fields(0)), "", .Fields(0))
        Else
            GetFieldValueYM = ""
        End If
        .Close
    End With
        
End Function

Public Function GetFieldValueMD(MyQ As String) As Variant
    Dim rsBrch As New ADODB.Recordset
    
    With rsBrch
        .Open MyQ, conToril, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetFieldValueMD = IIf(IsNull(.Fields(0)), "", .Fields(0))
        Else
            GetFieldValueMD = ""
        End If
        .Close
    End With
        
End Function

Public Function GetSalesQtyBranch(MyID As Long, myCON As String)
    Dim rsQty As New ADODB.Recordset
    Dim MyQ As String
    Dim MySales As Long, MyRet As Long
    
    MyQ = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        If myCON = "VAN1" Then
            .Open MyQ, conVAN1, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "VAN2" Then
            .Open MyQ, conVAN2, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "MRC2" Then
            .Open MyQ, conMRC2, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "MRC1" Then
            .Open MyQ, conMRC1, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "ULAS" Then
            .Open MyQ, conUlas, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "TORIL" Then
            .Open MyQ, conToril, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "YAMAHA" Then
            .Open MyQ, conYM, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "PANABO" Then
            .Open MyQ, conPanabo, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "POLOMOLOK" Then
            .Open MyQ, conPol, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "GENSAN" Then
            .Open MyQ, conGenSan, adOpenKeyset, adLockOptimistic
        End If
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    MyQ = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValueBranch(MyQ, myCON))
    GetSalesQtyBranch = MySales - MyRet
End Function


Public Function GetSalesQtyULAS(MyID As Long)
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, conUlas, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValueBranch(Qry, "ULAS"))
    GetSalesQtyULAS = MySales - MyRet
End Function

Public Function GetSalesQtyBrchYM(MyID As Long)
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, BrchConYM, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValueYM(Qry))
    GetSalesQtyBrchYM = MySales - MyRet
End Function

Public Function GetSalesQtyBrchMD(MyID As Long)
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, BrchConMD, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValueMD(Qry))
    GetSalesQtyBrchMD = MySales - MyRet
End Function


Public Function GetSalesQtyBrchVAN(MyID As Long)
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, BrchConVAN, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValueVAN(Qry))
    GetSalesQtyBrchVAN = MySales - MyRet
End Function


Public Function GetSalesQtyBrchMRC2(MyID As Long)
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, BrnConMRC2, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValueMRC2(Qry))
    GetSalesQtyBrchMRC2 = MySales - MyRet
End Function


Public Function GetDueDate(InvDate As Date, Term As Integer) As Date
    Dim mDay As Integer, TmpDue As Date, D30nxtM As Date, D45nxtM As Date
    
    
    If Term < 30 Then
        GetDueDate = DateAdd("d", Term, InvDate)
        Exit Function
    End If
    mDay = Day(InvDate)
    If mDay <= 15 Then
        TmpDue = DateSerial(Year(DateAdd("M", 2, InvDate)), Month(DateAdd("M", 2, InvDate)), 1) '1st Day of InvDate+2 months
        D30nxtM = DateAdd("D", -1, TmpDue)
        Select Case Term
            Case 30 To 35
                GetDueDate = D30nxtM
            Case 36 To 45
                GetDueDate = DateSerial(Year(DateAdd("M", 2, InvDate)), Month(DateAdd("M", 2, InvDate)), 15)
            Case 46 To 60
                TmpDue = DateSerial(Year(DateAdd("M", 2, D30nxtM)), Month(DateAdd("M", 2, D30nxtM)), 1) '1st Day of InvDate+2 months
                GetDueDate = DateAdd("D", -1, TmpDue)
            Case 61 To 75
                GetDueDate = DateSerial(Year(DateAdd("M", 3, InvDate)), Month(DateAdd("M", 3, InvDate)), 15)
            Case 76 To 90
                TmpDue = DateSerial(Year(DateAdd("M", 3, D30nxtM)), Month(DateAdd("M", 3, D30nxtM)), 1) '1st Day of InvDate+3 months
                GetDueDate = DateAdd("D", -1, TmpDue)
            Case 91 To 105
                GetDueDate = DateSerial(Year(DateAdd("M", 4, InvDate)), Month(DateAdd("M", 4, InvDate)), 15)
            Case 106 To 120
                TmpDue = DateSerial(Year(DateAdd("M", 4, D30nxtM)), Month(DateAdd("M", 4, D30nxtM)), 1) '1st Day of InvDate+4 months
                GetDueDate = DateAdd("D", -1, TmpDue)
            Case 121 To 135
                GetDueDate = DateSerial(Year(DateAdd("M", 5, InvDate)), Month(DateAdd("M", 5, InvDate)), 15)
            Case 136 To 150
                TmpDue = DateSerial(Year(DateAdd("M", 5, D30nxtM)), Month(DateAdd("M", 5, D30nxtM)), 1) '1st Day of InvDate+5 months
                GetDueDate = DateAdd("D", -1, TmpDue)
            Case 151 To 165
                GetDueDate = DateSerial(Year(DateAdd("M", 6, InvDate)), Month(DateAdd("M", 6, InvDate)), 15)
            Case 166 To 180
                TmpDue = DateSerial(Year(DateAdd("M", 6, D30nxtM)), Month(DateAdd("M", 6, D30nxtM)), 1) '1st Day of InvDate+6 months
                GetDueDate = DateAdd("D", -1, TmpDue)
            Case Else
                GetDueDate = DateAdd("D", Term, InvDate)
        End Select
    Else
        D30nxtM = DateSerial(Year(DateAdd("M", 2, InvDate)), Month(DateAdd("M", 2, InvDate)), 15) '15th Day of InvDate+2 months
        TmpDue = DateSerial(Year(DateAdd("M", 3, InvDate)), Month(DateAdd("M", 3, InvDate)), 1) '1st Day of InvDate+3 months
        D45nxtM = DateAdd("D", -1, TmpDue)  '30th Day of InvDate+2 months
        Select Case Term
            Case 30
                GetDueDate = D30nxtM
            Case 36 To 45
                GetDueDate = D45nxtM
            Case 46 To 60
                GetDueDate = DateAdd("M", 1, D30nxtM)
            Case 61 To 75
                TmpDue = DateSerial(Year(DateAdd("M", 2, D45nxtM)), Month(DateAdd("M", 2, D45nxtM)), 1) '1st Day of D45nxtM+2 months
                GetDueDate = DateAdd("D", -1, TmpDue)
            Case 76 To 90
                GetDueDate = DateAdd("M", 2, D30nxtM)
            Case 91 To 105
                TmpDue = DateSerial(Year(DateAdd("M", 3, D45nxtM)), Month(DateAdd("M", 3, D45nxtM)), 1) '1st Day of D45nxtM+3 months
                GetDueDate = DateAdd("D", -1, TmpDue)
            Case 106 To 120
                GetDueDate = DateAdd("M", 3, D30nxtM)
            Case 121 To 135
                TmpDue = DateSerial(Year(DateAdd("M", 4, D45nxtM)), Month(DateAdd("M", 4, D45nxtM)), 1) '1st Day of D45nxtM+4 months
                GetDueDate = DateAdd("D", -1, TmpDue)
            Case 136 To 150
                GetDueDate = DateAdd("M", 4, D30nxtM)
            Case 151 To 165
                TmpDue = DateSerial(Year(DateAdd("M", 5, D45nxtM)), Month(DateAdd("M", 5, D45nxtM)), 1) '1st Day of D45nxtM+5 months
                GetDueDate = DateAdd("D", -1, TmpDue)
            Case 166 To 180
                GetDueDate = DateAdd("M", 5, D30nxtM)
            Case Else
                GetDueDate = DateAdd("D", Term, InvDate)
        End Select
    End If
End Function


Public Function GetFieldValueBranch(MyQ As String, myCON As String) As Variant
    Dim rsAll As New ADODB.Recordset
    
    With rsAll
        If myCON = "VAN1" Then
            .Open MyQ, conVAN1, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "VAN2" Then
            .Open MyQ, conVAN2, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "VAN3" Then
            .Open MyQ, conVAN3, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "MRC2" Then
            .Open MyQ, conMRC2, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "MRC1" Then
            .Open MyQ, conMRC1, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "ULAS" Then
            .Open MyQ, conUlas, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "TORIL" Then
            .Open MyQ, conToril, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "YAMAHA" Then
            .Open MyQ, conYM, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "PANABO" Then
            .Open MyQ, conPanabo, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "POLOMOLOK" Then
            .Open MyQ, conPol, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "GENSAN" Then
            .Open MyQ, conGenSan, adOpenKeyset, adLockOptimistic
        End If
        If .RecordCount > 0 Then
            .MoveFirst
            GetFieldValueBranch = IIf(IsNull(.Fields(0)), "", .Fields(0))
        Else
            GetFieldValueBranch = ""
        End If
        .Close
    End With
        
End Function

Public Function ConnectMRC2() As Boolean

    On Error GoTo LocalErr
    If conMRC2.State = adStateOpen Then: conMRC2.Close
    conMRC2.CursorLocation = adUseClient
    conMRC2.ConnectionTimeout = 60
    conMRC2.Open "dsn=SQLMRC2TGM"
    ConnectMRC2 = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to MRC2 Database...."), vbExclamation
    ConnectMRC2 = False
End Function

Public Function ConnectVan3() As Boolean

    On Error GoTo LocalErr
    If conVAN3.State = adStateOpen Then: conVAN3.Close
    conVAN3.CursorLocation = adUseClient
    conVAN3.ConnectionTimeout = 60
    conVAN3.Open "dsn=MRCVAN3"
    ConnectVan3 = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to VAN3 Database...."), vbExclamation
    ConnectVan3 = False
End Function


Public Function ConnectVAN2() As Boolean

    On Error GoTo LocalErr
    If conVAN2.State = adStateOpen Then: conVAN2.Close
    conVAN2.CursorLocation = adUseClient
    conVAN2.ConnectionTimeout = 60
    conVAN2.Open "dsn=MRCVAN2"
    ConnectVAN2 = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to VAN2 Database...."), vbExclamation
    ConnectVAN2 = False
End Function

Public Function ConnectVAN1() As Boolean

    On Error GoTo LocalErr
    If conVAN1.State = adStateOpen Then: conVAN1.Close
    conVAN1.CursorLocation = adUseClient
    conVAN1.ConnectionTimeout = 60
    conVAN1.Open "dsn=SQLMTRVAN"
    ConnectVAN1 = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to VAN1 Database...."), vbExclamation
    ConnectVAN1 = False
End Function

Public Function ConnectMRC1() As Boolean

    On Error GoTo LocalErr
    If conMRC1.State = adStateOpen Then: conMRC1.Close
    conMRC1.CursorLocation = adUseClient
    conMRC1.ConnectionTimeout = 60
    conMRC1.Open "dsn=SQLMTRTGM"
    ConnectMRC1 = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to MRC1 Database...."), vbExclamation
    ConnectMRC1 = False
End Function

Public Function ConnectMRCGensan() As Boolean

    On Error GoTo LocalErr
    If conGenSan.State = adStateOpen Then: conGenSan.Close
    conGenSan.CursorLocation = adUseClient
    conGenSan.ConnectionTimeout = 60
    conGenSan.Open "dsn=MRCGENSAN"
    ConnectMRCGensan = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to MRC GENSAN Database...."), vbExclamation
    ConnectMRCGensan = False
End Function


Public Function ConnectUlas() As Boolean

    On Error GoTo LocalErr
    If conUlas.State = adStateOpen Then: conUlas.Close
    conUlas.CursorLocation = adUseClient
    conUlas.ConnectionTimeout = 60
    conUlas.Open "dsn=SQLMTRULAS"
    ConnectUlas = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to MOTO-R ULAS Database...."), vbExclamation
    ConnectUlas = False
End Function

Public Function ConnectToril() As Boolean

    On Error GoTo LocalErr
    If conToril.State = adStateOpen Then: conToril.Close
    conToril.CursorLocation = adUseClient
    conToril.ConnectionTimeout = 60
    conToril.Open "dsn=SQLMCPDEPOT"
    ConnectToril = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to MCPARTS DEPOT Database...."), vbExclamation
    ConnectToril = False
End Function

Public Function ConnectPanabo() As Boolean

    On Error GoTo LocalErr
    If conPanabo.State = adStateOpen Then: conPanabo.Close
    conPanabo.CursorLocation = adUseClient
    conPanabo.ConnectionTimeout = 60
    conPanabo.Open "dsn=MRCPANABO"
    ConnectPanabo = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to MRC PANABO Database...."), vbExclamation
    ConnectPanabo = False
End Function

Public Function ConnectHondaPol() As Boolean

    On Error GoTo LocalErr
    If conPol.State = adStateOpen Then: conPol.Close
    conPol.CursorLocation = adUseClient
    conPol.ConnectionTimeout = 60
    conPol.Open "dsn=HONDAPOL"
    ConnectHondaPol = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to BINOS POLOMOLOK Database...."), vbExclamation
    ConnectHondaPol = False
End Function

Public Function ConnectYamaha() As Boolean

    On Error GoTo LocalErr
    If conYM.State = adStateOpen Then: conYM.Close
    conYM.CursorLocation = adUseClient
    conYM.ConnectionTimeout = 60
    conYM.Open "dsn=YAMAHA"
    ConnectYamaha = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to YAMAHA MOTORHUB Database...."), vbExclamation
    ConnectYamaha = False
End Function

Public Function ConnectStoTomas() As Boolean

    On Error GoTo LocalErr
    If conStoTomas.State = adStateOpen Then: conStoTomas.Close
    conStoTomas.CursorLocation = adUseClient
    conStoTomas.ConnectionTimeout = 60
    conStoTomas.Open "dsn=mrcStoTomas"
    ConnectStoTomas = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to BINOS STO. TOMAS Database...."), vbExclamation
    ConnectStoTomas = False
End Function

Public Function ConnectMarbel() As Boolean

    On Error GoTo LocalErr
    If conMarbel.State = adStateOpen Then: conMarbel.Close
    conMarbel.CursorLocation = adUseClient
    conMarbel.ConnectionTimeout = 60
    conMarbel.Open "dsn=mrcMarbel"
    ConnectMarbel = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to YAMAHA MARBEL Database...."), vbExclamation
    ConnectMarbel = False
End Function

Public Function ConnectSamal() As Boolean

    On Error GoTo LocalErr
    If conSamal.State = adStateOpen Then: conSamal.Close
    conSamal.CursorLocation = adUseClient
    conSamal.ConnectionTimeout = 60
    conSamal.Open "dsn=mrcSamal"
    ConnectSamal = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to BINOS Samal Database...."), vbExclamation
    ConnectSamal = False
End Function

'********************* CODE UPDATED: 03/21/2017 ****************************
Public Sub sp_brnchReceiving_getDetails(pBRANCHCODE As String, pPACKINGNO As String, pPMATRIXID As Integer, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_get_PartsmatrixID(
'  pBranchCode VarChar(4)
'  )

    DoEvents
    Qry = "Call sp_partsmatrix.sp_brnchReceiving_getDetails('" & sysBranchCode & "', '" & pBRANCHCODE & "', '" & pPACKINGNO & "', " & pPMATRIXID & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchReceiving_getDetails()")
End Sub
'***************** END CODE UPDATED: 03/21/2017 ****************************

'********************* CODE UPDATED: 03/23/2017 ****************************
Public Sub sp_brnchReceiving_showrecord(pBRANCHCODE As String, pREFERENCE As String, pPONO As String, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_get_PartsmatrixID(
'  pBranchCode VarChar(4)
'  )

    DoEvents
    Qry = "Call sp_partsmatrix.sp_brnchReceiving_showrecord('" & sysBranchCode & "', '" & pBRANCHCODE & "', '" & pREFERENCE & "', '" & pPONO & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchReceiving_showrecord()")
End Sub
'***************** END CODE UPDATED: 03/23/2017 ****************************

'********************* CODE UPDATED: 03/30/2017 ****************************
Public Sub sp_brnchEvaluate_PM(pDateFrom As String, pDateTo As String, pDCode As Integer, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_get_PartsmatrixID(
'  pBranchCode VarChar(4)
'  )

    DoEvents
    Qry = "Call sp_partsmatrix.sp_brnchEvaluate_PM('" & sysBranchCode & "', '" & pDateFrom & "', '" & pDateTo & "', " & pDCode & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchEvaluate_PM()")
End Sub
'***************** END CODE UPDATED: 03/30/2017 ****************************

'********************* CODE UPDATED: 03/30/2017 ****************************
Public Sub sp_brnchEvaluate_PMdetails(pPACKINGNO As String, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_get_PartsmatrixID(
'  pBranchCode VarChar(4)
'  )

    DoEvents
    Qry = "Call sp_partsmatrix.sp_brnchEvaluate_PMdetails('" & sysBranchCode & "', '" & pPACKINGNO & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchEvaluate_PMdetails()")
End Sub
'***************** END CODE UPDATED: 03/30/2017 ****************************

'********************* CODE UPDATED: 04/03/2017 ****************************
Public Sub get_Model(pBRAND As String, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_get_Model(
'  pBranchCode VarChar(4)
'  ,pBRAND  Varchar(30)
'  )

    DoEvents
    Qry = "call sp_partsmatrix.sp_brnchQty_getModel('" & sysBranchCode & "', '" & pBRAND & "')"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_get_Model()")
End Sub


Public Sub get_Brand(objRs As ADODB.Recordset)
On Error GoTo er

'    PROCEDURE sp_partsmatrix.sp_get_Brand(
'      pBRANCHCODE VarChar(4)
'  )

    DoEvents
   ' Qry = "Select Distinct(PMBRAND) from mrc2retail.partsmatrix INNER JOIN mrc2retail.partsmatrix_dtls ON partsmatrix.PMATRIXID = partsmatrix_dtls.PMATRIXID group by partsmatrix_dtls.PMATRIXID"
          
    Qry = "call sp_partsmatrix.sp_get_Brand('" & sysBranchCode & "')"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->get_Brand_Model()")
End Sub

Public Sub sp_brnchQty_update(pPMATRIXID As Integer, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_brnchQty_update(
'  pBranchCode VarChar(4)
'  ,pPMATRIXID int(11)
'  )

    DoEvents
    Qry = "Call sp_partsmatrix.sp_brnchQty_update('" & sysBranchCode & "', " & pPMATRIXID & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchQty_update()")

End Sub


Public Sub sp_brnchQty_getDetails(pPMATRIXID As Integer, pBOOLEAN As String, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_brnchQty_getDetails(
'  pBranchCode VarChar(4)
'  , pPMATRIXID Int(11)
'  , pBOOLEAN Varchar(6)
'  )

    DoEvents
    Qry = "Call sp_partsmatrix.sp_brnchQty_getDetails('" & sysBranchCode & "', " & pPMATRIXID & ",'" & pBOOLEAN & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchQty_getDetails()")

End Sub



Public Sub sp_brnchQty_getDetails_total(pPMATRIXID As Integer, pBOOLEAN As String, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_brnchQty_getDetails_total(
'    pBranchCode VarChar(4)
'  , pPMATRIXID Int(11)
'  )

    DoEvents
    Qry = "Call sp_partsmatrix.sp_brnchQty_getDetails_total('" & sysBranchCode & "', " & pPMATRIXID & ",'" & pBOOLEAN & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchQty_getDetails_total()")

End Sub

Public Sub sp_brnchQty_getBrand(objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_brnchQty_summary(
'  pBranchCode VarChar(4)
'  )

    DoEvents
    Qry = "Call sp_partsmatrix.sp_brnchQty_getBrand('" & sysBranchCode & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
  Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchQty_summary()")

End Sub

Public Sub sp_brnchQty_summary(objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_brnchQty_summary(
'  pBranchCode VarChar(4)
'  )sp_brnchQty_getBrand

    DoEvents
    Qry = "Call sp_partsmatrix.sp_brnchQty_summary('" & sysBranchCode & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
  Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchQty_summary()")

End Sub


Public Sub sp_brnchQty_summary_WBRAND(pPMBRAND As String, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_brnchQty_summary(
'  pBranchCode VarChar(4)
'  )

    DoEvents
    Qry = "Call sp_partsmatrix.sp_brnchQty_summary_wBRAND('" & sysBranchCode & "','" & pPMBRAND & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
  Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchQty_summary()")

End Sub

Public Sub sp_brnchQty_create(pPMATRIXID As Integer, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_brnchQty_update(
'  pBranchCode VarChar(4)
'  ,pPMATRIXID int(11)
'  )

    DoEvents
    Qry = "Call sp_partsmatrix.sp_brnchQty_create('" & sysBranchCode & "', " & pPMATRIXID & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchQty_create()")

End Sub
'***************** END CODE UPDATED: 04/03/2017 ****************************

'********************* CODE UPDATED: 04/17/2017 ****************************
Public Sub sp_UpdatePartsmatrix(objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_UpdatePartsmatrix(
'  pBranchCode VarChar(4)
'  )

    DoEvents
    Qry = "Call sp_partsmatrix.sp_UpdatePartsmatrix('" & sysBranchCode & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_UpdatePartsmatrix()")

End Sub

Public Sub sp_UpdatePartsmatrix_dtls(objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_UpdatePartsmatrix_dtls(
'  pBranchCode VarChar(4)
'  )

    DoEvents
    Qry = "Call sp_partsmatrix.sp_UpdatePartsmatrix_dtls('" & sysBranchCode & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_UpdatePartsmatrix_dtls()")

End Sub

Public Sub sp_insertNewPMpgroup(objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_insertNewPMpgroup(
'  pBranchCode VarChar(4)
'  )

    DoEvents
    Qry = "Call sp_partsmatrix.sp_insertNewPMpgroup_v1('" & sysBranchCode & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_insertNewPMpgroup()")

End Sub
'***************** END CODE UPDATED: 04/17/2017 ****************************

'********************* CODE UPDATED: 08/01/2017 ****************************
Public Sub sp_get_Branch(objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_get_Branch(
'    pBranchCode VarChar(4)
'  )
   
    DoEvents
    Qry = "Select br.BRANCHID, br.BRANCHNAME, br.BRANCHCODE , ifnull(br.ADDRESS,''), ifnull(br.CONTACTNOS, '') from branch as br Where branchcategory IN ('MRC','WHOLESALE','DEPOT')"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_get_Branch()")

End Sub
'****************** END CODE UPDATED: 08/01/2017 ***************************

'********************* CODE UPDATED: 12/17/2020 ****************************
Public Sub sp_get_Department(objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_get_Department(
'    pBranchCode VarChar(4)
'  )
   
    DoEvents
    Qry = "Select DESCRIPTION as DEPARTMENT, `VALUE` as DEPCODE from settings where REMARKS in ('SUPPORT GROUP','LYRMP ADMIN')"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_get_Department()")

End Sub
'****************** END CODE UPDATED: 12/17/2020 ***************************


Public Function ConnectPartsCatalog() As Boolean

    On Error GoTo LocalErr
    If conPartsCat.State = adStateOpen Then: conPartsCat.Close
    conPartsCat.CursorLocation = adUseClient
    conPartsCat.ConnectionTimeout = 60
    conPartsCat.Open "dsn=partscatalog"
    ConnectPartsCatalog = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to Parts Catalog Database...."), vbExclamation
    ConnectPartsCatalog = False
End Function
