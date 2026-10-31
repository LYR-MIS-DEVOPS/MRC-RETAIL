Attribute VB_Name = "modConnection"
Option Explicit

Public Enum ACTIVELOCATION
    
    Branch = 0
    HQ = 1
    COPR = 2
    BRANCHSERVER = 3
    OLDBRANCH = 4
    YAMAHA = 5
    BINOSPOL = 6
    HQ28 = 7
    MOTORHUB = 8
    GLOBE = 9
    PLDT = 10
    GANTERETAIL = 11
    DATACENTER2 = 12
    IPBASED = 13
    
End Enum

Private EndPointBranchID        As Integer

Private ParamBCHCode            As String
Private ParamLocation           As ACTIVELOCATION
Public ParamIPaddress As String '04/22/19 change to public
Public connectionString As String
Private MRCEndPoint             As Object

Private Function getBCHConnectionString() As String
On Error GoTo er
    
    Set MRCEndPoint = CreateObject("MRCEndPoint.clsConnection")
    
    'Public Function  MRCEndPoint.getBCHConnectionString(ActiveLoc As ACTIVELOCATION, BCHCODE As String) As String
    getBCHConnectionString = MRCEndPoint.getBCHConnectionString(ParamLocation, ParamBCHCode, ParamIPaddress)
    EndPointBranchID = MRCEndPoint.getBranchID()
    SysDBNAME = MRCEndPoint.getDBNAME()
    Set MRCEndPoint = Nothing
Exit Function
er:
    MsgBox "ActiveX Connection String Binding Error!", vbCritical, "Binding Error!"
    End
End Function

Public Function getOnlineConnectionString() As String
On Error GoTo er
    
    Set MRCEndPoint = CreateObject("MRCEndPoint.clsConnection")
    
    'Public Function getBCHOnlineConnectionString(ActiveLoc As ACTIVELOCATION, BCHCODE As String) As String
    If ConMWHLocal = False Then
        getOnlineConnectionString = MRCEndPoint.getBCHOnlineConnectionString(ACTIVELOCATION.HQ, sysBranchCode, ParamIPaddress)
    Else
        If ParamLocation = 9 Then
            getOnlineConnectionString = MRCEndPoint.getBCHOnlineConnectionString(ACTIVELOCATION.GLOBE, sysBranchCode, ParamIPaddress)
        ElseIf ParamLocation = 10 Then
            getOnlineConnectionString = MRCEndPoint.getBCHOnlineConnectionString(ACTIVELOCATION.PLDT, sysBranchCode, ParamIPaddress)
        ElseIf ParamLocation = 12 Then
            getOnlineConnectionString = MRCEndPoint.getBCHOnlineConnectionString(ACTIVELOCATION.DATACENTER2, sysBranchCode, ParamIPaddress)
        ElseIf ParamLocation = 13 Then
            getOnlineConnectionString = MRCEndPoint.getBCHOnlineConnectionString(ACTIVELOCATION.IPBASED, sysBranchCode, ParamIPaddress)
        Else
            getOnlineConnectionString = ""

        End If
    End If
    
    Set MRCEndPoint = Nothing
Exit Function
er:
    MsgBox "ActiveX Connection String Binding Error!", vbCritical, "Binding Error!"
    End
End Function

Public Function getWarehouseConnectionString() As String
On Error GoTo er
    
    Set MRCEndPoint = CreateObject("MRCEndPoint.clsConnection")
    
    getWarehouseConnectionString = MRCEndPoint.getWarehouseConnectionString(ParamLocation)
'    If ConMWHLocal = False Then
'        getWarehouseConnectionString = MRCEndPoint.getBCHConnectionString(ACTIVELOCATION.HQ, "MWH", ParamIPaddress)
'    Else
'        If ParamLocation = 9 Then 'globe
'            getWarehouseConnectionString = "Driver={MySQL ODBC 3.51 Driver};Server=192.168.254.10;Port=3306;Database=mrcwarehouse;User=root;Password=laborday;Option=3;"
'        ElseIf ParamLocation = 10 Then 'pldt
'            getWarehouseConnectionString = "Driver={MySQL ODBC 3.51 Driver};Server=192.168.0.10;Port=3306;Database=mrcwarehouse;User=root;Password=laborday;Option=3;"
'        ElseIf ParamLocation = 12 Then 'DATACENTER2
'            getWarehouseConnectionString = MRCEndPoint.getBCHConnectionString(ACTIVELOCATION.DATACENTER2, "MWHFSRVR", ParamIPaddress)
'        ElseIf ParamLocation = 13 Then 'IPBASED
'            getWarehouseConnectionString = MRCEndPoint.getBCHConnectionString(ACTIVELOCATION.IPBASED, "MWHFSRVR", ParamIPaddress)
'        Else
'            getWarehouseConnectionString = "Driver={MySQL ODBC 3.51 Driver};Server=192.168.30.29;Port=3306;Database=mrcwarehouse;User=root;Password=laborday;Option=3;"
'        End If
'    End If
'    Set MRCEndPoint = Nothing
Exit Function
er:
    MsgBox "ActiveX Connection String Binding Error!", vbCritical, "Binding Error!"
    End
End Function

Public Function getDavaoWarehouseConnectionString() As String
On Error GoTo er
    
    Set MRCEndPoint = CreateObject("MRCEndPoint.clsConnection")
    
    'Public Function  MRCEndPoint.getBCHConnectionString(ActiveLoc As ACTIVELOCATION, BCHCODE As String) As String
    If ConMWHLocal = False Then
        getDavaoWarehouseConnectionString = MRCEndPoint.getBCHConnectionString(ACTIVELOCATION.HQ, "MWA", ParamIPaddress)
    Else
        If ParamLocation = 9 Then 'globe
            getDavaoWarehouseConnectionString = "Driver={MySQL ODBC 3.51 Driver};Server=192.168.254.10;Port=3306;Database=mrcwarehouseagdao;User=root;Password=laborday;Option=3;"
        ElseIf ParamLocation = 10 Then 'pldt
            getDavaoWarehouseConnectionString = "Driver={MySQL ODBC 3.51 Driver};Server=192.168.0.10;Port=3306;Database=mrcwarehouseagdao;User=root;Password=laborday;Option=3;"
        ElseIf ParamLocation = 12 Then 'DATACENTER2
            getDavaoWarehouseConnectionString = MRCEndPoint.getBCHConnectionString(ACTIVELOCATION.DATACENTER2, "MWAFSRVR", ParamIPaddress)
        ElseIf ParamLocation = 13 Then 'IPBASED
            If sysBranchCode = "MLA" Then
                getDavaoWarehouseConnectionString = MRCEndPoint.getBCHConnectionString(ACTIVELOCATION.IPBASED, "MWALOCAL", ParamIPaddress)
            Else
                getDavaoWarehouseConnectionString = MRCEndPoint.getBCHConnectionString(ACTIVELOCATION.IPBASED, "MWAFSRVR", ParamIPaddress)
            End If
        Else
            getDavaoWarehouseConnectionString = "Driver={MySQL ODBC 3.51 Driver};Server=192.168.1.10;Port=3306;Database=mrcwarehouseagdao;User=root;Password=laborday;Option=3;"
        End If
    End If
    Set MRCEndPoint = Nothing
Exit Function
er:
    MsgBox "ActiveX Connection String Binding Error!", vbCritical, "Binding Error!"
    End
End Function

Public Sub setMDIBackGround()
On Error GoTo er
    'Set MDIMain BackGround
    If EndPointBranchID >= 0 Then
        'MotoRep.bckgrnd.Height = MotoRep.Height
        'MotoRep.bckgrnd.Picture = frmHoldBackground.picBackground(EndPointBranchID).Picture
        If EndPointBranchID <= 33 Then
            'MotoRep.imgBG.Picture = frmHoldBackground.picBackground(EndPointBranchID).Picture
            MotoRep.Picture = frmHoldBackground.picBackground(EndPointBranchID).Picture
            MotoRep.Icon = frmHoldBackground.picBackground(EndPointBranchID).MouseIcon
        End If
    Else
        MotoRep.Picture = frmHoldBackground.picDefault.Picture
        'MotoRep.imgBG.Picture = frmHoldBackground.picDefault.Picture
    End If

    DoEvents
    Call Unload(frmHoldBackground)
    Set frmHoldBackground = Nothing
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Binding Error!"
    End
End Sub


Public Sub DBMainConnection(objCon As ADODB.Connection)
On Error GoTo er
    If objCon.State <> 0 Then objCon.Close
    objCon.CursorLocation = adUseClient
    objCon.ConnectionTimeout = 1800
    objCon.connectionString = getBCHConnectionString()
    objCon.Open
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Database Connection Error!"
    End
End Sub


Public Sub DBMainConnect(objCon As ADODB.Connection)
On Error GoTo er

    
    If objCon.State <> 0 Then objCon.Close
    objCon.CursorLocation = adUseClient
    objCon.ConnectionTimeout = 1800
    
    objCon.connectionString = connectionString
    objCon.Open
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Database Connection Error!"
    End
End Sub

Public Sub ParseCommandArgument(Optional paramCommand As String = "")
On Error GoTo er
    Dim strCommand As String
    Dim strArr() As String
    'Command Argument Format
    'Ative Location & Branch Code
    'Separator = &
    
    strCommand = Command$ '0&MRT
    If Trim(strCommand) = "" Then
        strCommand = paramCommand
    End If
    
    strArr() = Split(strCommand, "&")
    ParamLocation = strArr(0)
    ParamBCHCode = strArr(1)
    
    If ParamLocation = "13" Then
        ParamIPaddress = strArr(2) '04/22/19
    End If
    If Trim(ParamLocation) = "" Or Trim(ParamBCHCode) = "" Then
        GoTo er
    End If
Exit Sub
er:
    MsgBox "INVALID COMMAND ARGUMENT!", vbCritical, App.Title
    End
End Sub

 
'************** CODE UPDATED: 04/19/17 *****************
Public Function TestConnection() As Boolean
On Error GoTo er
    TestConnection = False
    If MainCon.State = 0 Then
        If OnlineCon.State = 0 Then
            Exit Function
        End If
    End If
    TestConnection = True
Exit Function
er:
    TestConnection = False
    MsgBox Err.Description, vbCritical, "Unable to Connect to Server"
End Function
'********** END CODE UPDATED: 04/19/17 *****************














 
