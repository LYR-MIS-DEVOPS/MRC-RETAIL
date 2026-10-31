Attribute VB_Name = "modMotoR"
Global LYRcon As New ADODB.Connection

'report function declaration
Declare Function CreateFieldDefFile Lib "p2smon.dll" (lpUnk As Object, ByVal filename As String, ByVal bOverWriteExistingFile As Long) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'variable for saving/updating pictures
Global Chunk() As Byte
Global lngLengh As Long
Global intChunks As Integer
Global intFragment As Integer
Global Const ChunkSize = 1000 '16384
Global Const lngDataFile = 1
'searching variables

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Global Const CB_FINDSTRING = &H14C
Declare Function CharUpper Lib "user32" Alias "CharUpperA" (ByVal lpsz As String) As String

'user variables
Global UserID       As Integer
Global DispBtn      As Boolean
Global UserType     As String
Global ShowCost     As Boolean
Global SysBranch    As String
Global DateStart    As String
Global DateEnd      As String
Global DateNow      As String

Sub Main()
    
    On Error GoTo LocalErr
    LYRcon.CursorLocation = adUseClient
    
    
    'LYRcon.Open "dsn=SQLMTRTGM"
    'LYRcon.Open "dsn=mrcpanabo"
    'LYRcon.Open "dsn=SQLMTRULAS"
    'LYRcon.Open "dsn=SQLMRC2TGM"
    'LYRcon.Open "dsn=MRCVAN2"
    'LYRcon.Open "dsn=sqlmtrvan"
    'LYRcon.Open "dsn=MRCGENSAN"
    'LYRcon.Open "dsn=MRCPANABO"
    'LYRcon.Open "dsn=MRCSTOTOMAS"
    'LYRcon.Open "dsn=YAMAHA"
    'LYRcon.Open "dsn=mrcsamal"
    'LYRcon.Open "dsn=hondapol"
   LYRcon.Open "dsn=mrcwarehouse"
   'LYRcon.Open "dsn=mrcvan3"
   'LYRcon.Open "dsn=mrc2retail"
   'LYRcon.Open "DSN=SQLMCPDEPOT"
   'LYRcon.Open "dsn=mrcmarbel"
   'LYRcon.Open "dsn=MRCMOTORHUBTAGUM"
   'LYRcon.Open "dsn=MRCPOLOMOLOK"
   'LYRcon.Open "DSN=MRCCABANTIAN"
   
    
    'MotoRep.Show
   Call SBranch
   LogIn.Show vbModal
    'frmSplash.Show vbModal
   Exit Sub
LocalErr:
    MsgBox "Error connecting to server..." & vbCrLf & "Make sure you are connected to the Server before running the program.", vbExclamation + vbOKOnly, "Connection Failed"

End Sub

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

Public Sub CenterFrm(ByVal Parentfrm As MDIForm, ByVal Childfrm As Form)
    Childfrm.Left = (Parentfrm.Width \ 2) - (Childfrm.Width \ 2)
    Childfrm.Top = (Parentfrm.Height \ 2) - (Childfrm.Height \ 2) - 600
End Sub

Public Function MaxID(sql As String) As Integer
    Dim rsMax As New ADODB.Recordset
    
    rsMax.Open sql, LYRcon, adOpenKeyset, adLockOptimistic
    rsMax.MoveFirst
    MaxID = IIf(IsNull(rsMax.Fields(0)), 0, rsMax.Fields(0))
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
        InsertHere = IIf(rsPref!InsertIndex <> 0, rsPref!InsertIndex, -1)
        'Start Printing here
        mcReport.PaperSize = crPaperUser
        If FormName = "SwMoUB Front LSCP" Then
            mcReport.PaperOrientation = crLandscape
            mcReport.PaperSize = crPaperLegal
        End If
        If FormName = "Agreement" Or FormName = "Promissory Note" Or FormName = "SwMoUB Front PORT" Or FormName = "Chattel Mortgage" Then
            mcReport.PaperSize = crPaperLegal
        End If
        mcReport.TopMargin = 0
        mcReport.BottomMargin = 0
        mcReport.Sections(1).Height = 0
        mcReport.Sections(2).Height = 0
        With mcReport.Sections(3)
            .Height = rsPref!FormHeight
            For i = LBound(MyDesc) To UBound(MyDesc)
                If i = InsertHere Then
                    NxtRow = rsPref!LoopStartTop
                    mySubLeft = Split(rsPref!SubLeft, ",")
                    mySubLen = Split(rsPref!SubLen, ",")
                    mySubBold = Split(rsPref!SubFBold, ",")
                    mySubAlign = Split(rsPref!SubAlign, ",")
                    If FormName = "Promissory Note" Then z = 10
                    For j = 1 To MyFlx.Rows - 1
                        For k = LBound(mySubLeft) To UBound(mySubLeft)
                            Set txtObj = .AddTextObject(MyFlx.TextMatrix(j, k + 1), mySubLeft(k), NxtRow)
                            txtObj.Width = Val(mySubLen(k)): txtObj.HorAlignment = Val(mySubAlign(k)): txtObj.Font.Name = "Arial": txtObj.Font.Size = Val(rsPref!LoopFSize): txtObj.Height = Val(rsPref!LoopHeight)
                        Next k
                        If FormName = "Promissory Note" Then    'Control rows in Promi
                            NxtRow = NxtRow + rsPref!LoopInterval - z
                            z = IIf(z = 70, 10, z + 10)
                        Else
                            NxtRow = NxtRow + rsPref!LoopInterval
                        End If
                    Next j
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

Public Function Nth(myNum As Integer) As String
    Dim sNum As String
    
    Select Case myNum
        Case 1
            sNum = myNum & "st"
        Case 2
            sNum = myNum & "nd"
        Case 3
            sNum = myNum & "rd"
        Case Else
            sNum = myNum & "th"
    End Select
    Nth = sNum
End Function

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
    
    'Qry = "SELECT [LASTNAME] & ', ' & [FIRSTNAME] & ' ' & Left$([MIDDLENAME],1) & '.' AS USERNAME, EMPLOYEES.EMPID " & _
          "From EMPLOYEES WHERE (((EMPLOYEES.EMPID)=" & UserID & "));"
    'mySQL
    Qry = "SELECT CONCAT(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.'))) AS USERNAME, EMPLOYEES.EMPID " & _
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

'Ageing functions

Public Function GetMaxValue(MyAcct As String, Bill As Boolean) As Integer
    Dim rsMax As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT Max(LEDGER." & IIf(Bill, "BillNo", "SeqNo") & ") AS MaxOfNO From LEDGER " & _
          "WHERE (((LEDGER.ACCOUNTNO)='" & MyAcct & "'));"
    'mySQL
    Qry = "SELECT Max(count(LEDGER." & IIf(Bill, "BillNo", "SeqNo") & ")) AS MaxOfNO From LEDGER " & _
          "WHERE LEDGER.ACCOUNTNO='" & MyAcct & "';"
    
    With rsMax
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            GetMaxValue = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            GetMaxValue = 0
        End If
        .Close
    End With
End Function


Public Function LastArrAdv(MyAcct As String) As Single
    Dim rsLast As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT LEDGER.ACCOUNTNO, LEDGER.CRNTARRADV, LEDGER.SEQNO From LEDGER " & _
          "WHERE (((LEDGER.ACCOUNTNO)='" & MyAcct & "')) ORDER BY LEDGER.SEQNO DESC;"
    'mySQL
    Qry = "SELECT LEDGER.ACCOUNTNO, LEDGER.CRNTARRADV, LEDGER.SEQNO From LEDGER " & _
          "WHERE LEDGER.ACCOUNTNO='" & MyAcct & "' ORDER BY LEDGER.SEQNO DESC;"
    
    With rsLast
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            LastArrAdv = IIf(IsNull(!CrntArrAdv), 0, !CrntArrAdv)
        Else
            LastArrAdv = 0
        End If
        .Close
    End With
End Function

Public Function LastTotalCredit(MyAcct As String) As Single
    Dim rsTtlCredit As New ADODB.Recordset
    Dim Qry As String
    
    If MyAcct = "" Then Exit Function
    'Qry = "SELECT LEDGER.SEQNO, LEDGER.TTLCREDIT, LEDGER.DATEPAID From LEDGER " & _
          "WHERE (((LEDGER.ACCOUNTNO)='" & MyAcct & "') AND (Not (LEDGER.DATEPAID) Is Null)) " & _
          "ORDER BY LEDGER.SEQNO DESC;"
    'mySQL
    Qry = "SELECT LEDGER.SEQNO, LEDGER.TTLCREDIT, LEDGER.DATEPAID From LEDGER " & _
          "WHERE LEDGER.ACCOUNTNO='" & MyAcct & "' AND (Not (LEDGER.DATEPAID) Is Null) " & _
          "ORDER BY LEDGER.SEQNO DESC;"
    
    With rsTtlCredit
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            LastTotalCredit = IIf(IsNull(!TtlCredit), 0, !TtlCredit)
        Else
            LastTotalCredit = 0
        End If
        .Close
    End With
End Function


Public Function StandingBal(MyAcct As String) As Single
    Dim rsBal As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT LEDGER.SEQNO, LEDGER.OSBAL, LEDGER.DATEPAID From LEDGER " & _
          "WHERE (((LEDGER.ACCOUNTNO)='" & MyAcct & "') AND (Not (LEDGER.DATEPAID) Is Null)) " & _
          "ORDER BY LEDGER.SEQNO DESC;"
    'mySQL
    Qry = "SELECT LEDGER.SEQNO, LEDGER.OSBAL, LEDGER.DATEPAID From LEDGER " & _
          "WHERE LEDGER.ACCOUNTNO='" & MyAcct & "' AND (Not (LEDGER.DATEPAID) Is Null) " & _
          "ORDER BY LEDGER.SEQNO DESC;"
    
    With rsBal
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            StandingBal = IIf(IsNull(!OSBal), 0, !OSBal)
        Else
            StandingBal = 0
        End If
        .Close
    End With
End Function

Public Function DatePaid(MyAcct As String, AgeYr As Integer, AgeMonth As Integer) As String
    Dim rsDate As New ADODB.Recordset
    Dim Qry As String
    Dim dFrom As Date
    
    dFrom = DateSerial(AgeYr, AgeMonth, 7)
    'Qry = "SELECT LEDGER.DATEPAID, LEDGER.AMTPAID From LEDGER " & _
          "WHERE (((LEDGER.ACCOUNTNO)='" & MyAcct & "') AND ((LEDGER.DATEPAID)>=#" & dFrom & "# And (LEDGER.DATEPAID)<#" & DateAdd("m", 1, dFrom) & "#)) ORDER BY LEDGER.DATEPAID DESC;"
    'mySQL
    Qry = "SELECT LEDGER.DATEPAID, LEDGER.AMTPAID From LEDGER " & _
          "WHERE LEDGER.ACCOUNTNO='" & MyAcct & "' AND LEDGER.DATEPAID>='" & Format(dFrom, "yyyy-mm-dd") & "' And LEDGER.DATEPAID<'" & Format(DateAdd("m", 1, dFrom), "yyyy-mm-dd") & "' ORDER BY LEDGER.DATEPAID DESC;"
    
    With rsDate
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            DatePaid = IIf(IsNull(!DatePaid), "", Day(!DatePaid))
        Else
            DatePaid = ""
        End If
        .Close
    End With
End Function


Public Function AmtDownPaid(DP As Boolean, MyAcct As String, AgeYr As Integer, AgeMonth As Integer) As Single
    Dim rsPaid As New ADODB.Recordset
    Dim Qry As String, Tmp As String, dFrom As Date
    Dim dpSum As Single, MySum As Single
    
    dFrom = DateSerial(AgeYr, AgeMonth, 7)
    If DP Then
        'Qry = "SELECT LEDGER.ACCOUNTNO, LEDGER.AMTPAID, LEDGER.ARRADV, LEDGER.DATEPAID, LEDGER.REMARKS From LEDGER " & _
              "WHERE (((LEDGER.ACCOUNTNO)='" & MyAcct & "') AND ((LEDGER.DATEPAID)>=#" & dFrom & "# And (LEDGER.DATEPAID)<#" & DateAdd("m", 1, dFrom) & "#));"
        'mySQl
        Qry = "SELECT LEDGER.ACCOUNTNO, LEDGER.AMTPAID, LEDGER.ARRADV, LEDGER.DATEPAID, LEDGER.REMARKS From LEDGER " & _
              "WHERE (((LEDGER.ACCOUNTNO)='" & MyAcct & "') AND ((LEDGER.DATEPAID)>='" & Format(dFrom, "yyyy-mm-dd") & "' And (LEDGER.DATEPAID)<'" & Format(DateAdd("m", 1, dFrom), "yyyy-mm-dd") & "'));"
        
    Else
        'Qry = "SELECT LEDGER.ACCOUNTNO, LEDGER.AMTPAID, LEDGER.ARRADV, LEDGER.DATEPAID, LEDGER.REMARKS From LEDGER " & _
              "WHERE (((LEDGER.ACCOUNTNO)='" & MyAcct & "') AND ((LEDGER.DATEPAID)>=#" & dFrom & "# And (LEDGER.DATEPAID)<#" & DateAdd("m", 1, dFrom) & "#));"
        'mySQL
        Qry = "SELECT LEDGER.ACCOUNTNO, LEDGER.AMTPAID, LEDGER.ARRADV, LEDGER.DATEPAID, LEDGER.REMARKS From LEDGER " & _
              "WHERE (((LEDGER.ACCOUNTNO)='" & MyAcct & "') AND ((LEDGER.DATEPAID)>='" & Format(dFrom, "yyyy-mm-dd") & "' And (LEDGER.DATEPAID)<'" & Format(DateAdd("m", 1, dFrom), "yyyy-mm-dd") & "'));"
        
    End If
    With rsPaid
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySum = 0: dpSum = 0
            Do While Not .EOF
                Tmp = IIf(IsNull(!Remarks), "", !Remarks)
                If Tmp = "" Or Tmp = "DP" Then
                    Tmp = IIf(IsNull(!ArrAdv), "", !ArrAdv) 'Identify if DP
                    If UCase$(Tmp) <> "FULL DP" And UCase$(Tmp) <> "PRTL DP" Then
                        MySum = MySum + IIf(IsNull(!AmtPaid), 0, !AmtPaid)
                    Else
                        dpSum = dpSum + IIf(IsNull(!AmtPaid), 0, !AmtPaid)
                    End If
                End If
                .MoveNext
            Loop
        Else
            MySum = 0: dpSum = 0
        End If
        .Close
    End With
    If DP Then
        AmtDownPaid = dpSum
    Else
        AmtDownPaid = MySum
    End If
End Function

Public Function Rebate(MyAcct As String, AgeYr As Integer, AgeMonth As Integer) As Single
    Dim rsRebate As New ADODB.Recordset
    Dim Qry As String
    Dim dFrom As Date
    
    dFrom = DateSerial(AgeYr, AgeMonth, 7)
    'Qry = "SELECT Sum(LEDGER.REBATE) AS SumOfREBATE From LEDGER " & _
          "WHERE (((LEDGER.ACCOUNTNO)='" & MyAcct & "') AND ((LEDGER.DATEPAID)>=#" & dFrom & "# And (LEDGER.DATEPAID)<#" & DateAdd("m", 1, dFrom) & "#));"
    'mySQl
    Qry = "SELECT Sum(LEDGER.REBATE) AS SumOfREBATE From LEDGER " & _
          "WHERE (((LEDGER.ACCOUNTNO)='" & MyAcct & "') AND ((LEDGER.DATEPAID)>='" & Format(dFrom, "yyyy-mm-dd") & "' And (LEDGER.DATEPAID)<'" & Format(DateAdd("m", 1, dFrom), "yyyy-mm-dd") & "'));"
    
    With rsRebate
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Rebate = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            Rebate = 0
        End If
        .Close
    End With
End Function


Public Function GetSettings(MyRemarks As String) As String
    Dim rsSettings As New ADODB.Recordset
    Dim Qry As String
    
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


Public Function NumberOfAccts(MyCustID As Integer) As Integer
    Dim rsOrder As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT UNITORDERED.* From UNITORDERED " & _
          "WHERE (((UNITORDERED.CUSTOMERID)=" & MyCustID & ") AND ((UNITORDERED.ACCTSTATUS)<>'Closed' Or (UNITORDERED.ACCTSTATUS) Is Null));"
    'mySQL
    Qry = "SELECT UNITORDERED.* From UNITORDERED " & _
          "WHERE UNITORDERED.CUSTOMERID=" & MyCustID & " AND UNITORDERED.ACCTSTATUS<>'Closed' Or (UNITORDERED.ACCTSTATUS) Is Null;"
    
    With rsOrder
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        NumberOfAccts = .RecordCount
        .Close
    End With
End Function
Public Function ComputerName() As String
    Dim Cn As String

    Cn = Space(25)
    
    GetComputerName Cn, 25
    
    ComputerName = Left$(RTrim(Cn), Len(RTrim(Cn)) - 1)
End Function

Public Sub SBranch()
    Dim rsBrn As New ADODB.Recordset
    Dim mBrc As String
    
    mBrc = "SELECT branch.BRANCHNAME, Branch.Address, branch.CONTACTNOS From Branch Where Branch.SYSTEMBRANCH = 1"
    With rsBrn
        .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            SysBranch = !BRANCHNAME
            SysAddress = !ADDRESS
            SysTel = IIf(IsNull(!CONTACTNOS), "", !CONTACTNOS)
        End If
        .Close
    End With
    
End Sub
