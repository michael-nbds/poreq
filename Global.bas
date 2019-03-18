Attribute VB_Name = "Module1"
Attribute VB_Description = "globals"
Option Explicit
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName As String, ByVal KeyName As String, ByVal default As String, ByVal ReturnedString As String, ByVal MAXSIZE As Long, ByVal FileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName As String, ByVal KeyName As Any, ByVal NewString As String, ByVal FileName As String) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer$, nSize&) As Long
Declare Sub cCnvASCIItoEBCDIC Lib "mcstr-32.dll" (Txt As String)
Declare Sub cCnvEBCDICtoASCII Lib "mcstr-32.dll" (Txt As String)

Public Const LOCALE_SLANGUAGE As Long = &H2     'localized name of language
Public Const LOCALE_SSHORTDATE As Long = &H1F   'short date format string
Public Const LOCALE_SLONGDATE As Long = &H20    'long date format string
Public Const DATE_LONGDATE As Long = &H2
Public Const DATE_SHORTDATE As Long = &H1
Public Const HWND_BROADCAST As Long = &HFFFF&
Public Const WM_SETTINGCHANGE As Long = &H1A

Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long


Public Declare Function GetLocaleInfo Lib "kernel32" _
   Alias "GetLocaleInfoA" _
  (ByVal Locale As Long, _
   ByVal LCType As Long, _
   ByVal lpLCData As String, _
   ByVal cchData As Long) As Long

'---------------------
'DefInt A-Z
Public lCnct As String
Public Cnct As String
Public D_Cnct As String
Public G_SQL_Server As String
Public G_SQL_User As String
Public G_SQL_Pass As String
Public G_SQL_ODBC As String
Public G_SQL_GLOBAL As String
Public G_SQL_Provider As String
Public G_Local As Integer
Public G_Filename As String
Public MarsDataPath As String
Public preq_path As String
Public G_Sorders As Integer
Public MarsDataBaseName As String
Public G_PO_format As Integer
Public G_Vspace As Integer
Public G_I_VXref As Integer
Public G_Buy_Perform As Integer
Public G_VendCons As String
Public G_T_Weight As Double
Public G_pormemotxt As String
Public G_combo1 As String
Public G_pdate As String
Public G_Load_Comma As Integer
Public G_Cust As Integer
Public In_Perform As Integer
Public Label4 As String
Public rdate As String
Public G_Caw As Integer
Public G_Warehouse As String
Public POFname As String
Public G_TR_ConSwitch As Integer
Public G_PON As Double
Public G_Free As Integer

Public cnn As adodb.Connection              'server connection
Public lcnn As adodb.Connection             'local (MSDE) connection
Public MarsWorkspace As Workspace

Public IFields(200) As String

Global DB As Database
Global DefaultWorkspace As Workspace
Public Function GetUserLocaleInfo(ByVal dwLocaleID As Long, _
                                  ByVal dwLCType As Long) As String

   Dim sReturn As String
   Dim r As Long

  'call the function passing the Locale type
  'variable to retrieve the required size of
  'the string buffer needed
   r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
    
  'if successful..
   If r Then
    
     'pad the buffer with spaces
      sReturn = Space$(r)
       
     'and call again passing the buffer
      r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
     
     'if successful (r > 0)
      If r Then
      
        'r holds the size of the string
        'including the terminating null
         GetUserLocaleInfo = Left$(sReturn, r - 1)
      
      End If
   
   End If
    
End Function
Function check_po(ponum As String) As Integer
    Dim filnum As Integer
    Dim tstr$
    Dim ctbl As Recordset
    Dim oTbl As Recordset
    On Error Resume Next
    Err = 0
    DB.Execute "drop table pocheck"
    DB.Execute "create table pocheck (" & _
               " vendor   text(9) ," & _
               " item_num text(28)," & _
               " loc      text(5) ," & _
               " qty      double , " & _
               " match    text(1))  "
               
    DB.Execute " create index IndexIW on pocheck (item_num, loc)"

    filnum = FreeFile
    Err = 0
    Open preq_path & "\" & ponum & ".req" For Input As #filnum
    If Err = 53 Then
        DB.Execute "drop table pocheck"
        Close
        Exit Function
    End If
    check_po = 0
    Set ctbl = DB.OpenRecordset("pocheck")
    Do While Not EOF(filnum)
        Line Input #filnum, tstr$
        DoEvents
        If G_Load_Comma Then
            Call Parse_Fields(tstr$, 43)
            'If Trim(IFields(3)) = "FDB1050REC" Then
            '    Debug.Print
            'End If
                
            ctbl.AddNew
            ctbl(0) = IFields(4)
            ctbl(1) = IFields(3)
            ctbl(2) = IFields(5)
            ctbl(3) = Val(IFields(9))
            ctbl(4) = "N"
            ctbl.Update
        Else
            ctbl.AddNew
            ctbl(0) = Mid(tstr$, 40, 9)
            ctbl(1) = Mid(tstr$, 12, 28)
            ctbl(2) = Mid(tstr$, 49, 5)
            ctbl(3) = Val(Mid(tstr$, 100, 7))
            ctbl(4) = "N"
            ctbl.Update
        End If
    Loop
    ctbl.Close
    Dim ccnt As Long
    Dim ocnt As Long
    Set ctbl = DB.OpenRecordset("pocheck")
    ctbl.Index = "IndexIW"
    ctbl.MoveLast
    ccnt = ctbl.RecordCount
    Set oTbl = DB.OpenRecordset("select * from orders where ord_qty_im > 0")
    oTbl.MoveLast
    ocnt = oTbl.RecordCount
    oTbl.MoveFirst
    If oTbl.EOF Then
        ctbl.Close
        oTbl.Close
        Exit Function
    End If
    If ccnt <> ocnt Then
        ctbl.Close
        oTbl.Close
        Exit Function
    End If
    On Error GoTo 0
    Do While Not oTbl.EOF
        ctbl.Seek "=", oTbl!item_num_im, oTbl!whse_im
        If Not ctbl.NoMatch Then
            If Trim(ctbl(0)) = Trim(oTbl!vendor_im) And _
               Trim(ctbl(1)) = Trim(oTbl!item_num_im) And _
               Trim(ctbl(2)) = Trim(oTbl!whse_im) And _
               Val(Trim(ctbl(3))) = Val(Trim(oTbl!ord_qty_im)) Then
                    ctbl.Edit
                    ctbl(4) = "Y"
                    ctbl.Update
            End If
        End If
        oTbl.MoveNext
        DoEvents
    Loop
    ctbl.Close
    oTbl.Close
    Set ctbl = DB.OpenRecordset("select count(*) from pocheck where match='N'")
    If ctbl(0) = 0 Then
        check_po = -1
    Else
        check_po = 0
    End If
    Exit Function
End Function
Public Function Parse_Fields(pStr As String, num As Integer) As Integer
    Dim j%, k%, i%
    On Error GoTo bail_out
    For i% = 1 To num
        IFields(i%) = ""
    Next
    'Replace(pStr, Chr$(39), "'")
    pStr = Replace(pStr, "'", "`")
    For i% = 1 To num
        j% = InStr(pStr, ",")
        If j% = 0 Then
            If Left(pStr, 1) = Chr$(34) Then
                If Len(pStr) > 2 Then
                    pStr = Mid(pStr, 2, Len(pStr) - 2)
                End If
            End If
            IFields(i%) = pStr
            Exit For
        End If
        If j% - InStr(j% + 1, pStr, ",") = 0 Then
            IFields(i%) = Mid$(pStr, j% + 1)
            Exit For
        End If
        If InStr(j% + 1, pStr, ",") = 1 Then
            IFields(i%) = " "
            pStr = Mid$(pStr, j% + 1)
        Else
            If i% = 6 And InStr(1, pStr, Chr$(34)) > 0 Then
                'If Trim(IFields(3)) = "FDB1050REC" Then
                If Trim(IFields(3)) = "FDB2000RFB" Then
                    Debug.Print
                End If
                'If InStr(1, pStr, Chr$(34) & Chr$(34)) > 0 And InStr(1, pStr, ",") > InStr(1, pStr, Chr$(34) & Chr$(34)) Then
                If InStr(1, pStr, Chr$(34)) > 0 And InStr(1, pStr, ",") < InStr(2, pStr, Chr$(34)) Then
                    IFields(i%) = Mid$(pStr, 2, InStr(2, pStr, Chr$(34)) - 2)
                    IFields(i%) = Replace(IFields(i%), ",", " ")
                    j% = InStr(InStr(2, pStr, Chr$(34)), pStr, Chr$(34))
                    'IFields(i%) = Mid$(pStr, 2, j% - 1)
                    'IFields(i%) = Replace(pStr, Chr$(34), "`")
                ElseIf InStr(1, pStr, Chr$(34)) > 0 And InStr(1, pStr, ",") > InStr(2, pStr, Chr$(34)) Then
                    j% = InStr(InStr(1, pStr, ","), pStr, Chr$(34)) - 2
                    IFields(i%) = Mid$(pStr, 2, j% - 1)
                    IFields(i%) = Replace(IFields(i%), Chr$(34), "`")
                Else
                    j% = InStr(2, pStr, Chr$(34))
                    IFields(i%) = Mid$(pStr, 1, j% - 1)
                End If
                'IFields(i%) = Mid$(pStr, 1, j% - 1)
                pStr = Mid$(pStr, j% + 2)
            Else
                IFields(i%) = Mid$(pStr, 1, j% - 1)
                If Right(Trim(IFields(i%)), 1) = "-" Then
                    IFields(i%) = "-" & Left$(Trim(IFields(i%)), Len(IFields(i%)) - 1)
                End If
                pStr = Mid$(pStr, j% + 1)
            End If
        End If
                
        For k% = 1 To Len(IFields(i%))
            If Mid$(IFields(i%), k%, 1) = Chr$(34) Then
                Mid$(IFields(i%), k%, 1) = " "
            End If
        Next k%
        IFields(i%) = Trim(IFields(i%))
    Next
    Exit Function
bail_out:
    Parse_Fields = Err
End Function
Sub DB_Open()
    'Set DefaultWorkspace = Workspaces(0)
    'Set DB = DefaultWorkspace.OpenDatabase(App.Path & "\wmars.mdb")
End Sub
Function PercentPoint_FillRate(RateIN As Variant) As Double
    Dim diff
    Dim delta
    Select Case RateIN
        Case 50 To 59.999
            delta = 0.5
            diff = (RateIN - 50) / 10
            PercentPoint_FillRate = 0.1 + (delta * diff)
        Case 60 To 69.999
            delta = 0.3
            diff = (RateIN - 60) / 10
            PercentPoint_FillRate = 0.6 + (delta * diff)
        Case 70 To 79.999
            delta = 0.2
            diff = (RateIN - 70) / 10
            PercentPoint_FillRate = 0.9 + (delta * diff)
        Case 80 To 84.999
            delta = 0.2
            diff = (RateIN - 80) / 5
            PercentPoint_FillRate = 1.1 + (delta * diff)
        Case 85 To 89.999
            delta = 0.3
            diff = (RateIN - 85) / 5
            PercentPoint_FillRate = 1.3 + (delta * diff)
        Case 90 To 94.999
            delta = 0.4
            diff = (RateIN - 90) / 5
            PercentPoint_FillRate = 1.6 + (delta * diff)
        Case 95 To 96.999
            delta = 0.4
            diff = (RateIN - 95) / 2
            PercentPoint_FillRate = 2 + (delta * diff)
        Case 97 To 97.999
            delta = 0.2
            diff = (RateIN - 97) / 1
            PercentPoint_FillRate = 2.4 + (delta * diff)
        Case 98 To 98.999
            delta = 0.3
            diff = (RateIN - 98) / 1
            PercentPoint_FillRate = 2.6 + (delta * diff)
        Case 99 To 99.499
            delta = 0.3
            diff = (RateIN - 99) / 0.5
            PercentPoint_FillRate = 2.9 + (delta * diff)
        Case Else
            If RateIN <= 50 Then PercentPoint_FillRate = 0.1
            If RateIN >= 99.5 Then PercentPoint_FillRate = 3.2
    End Select
End Function

Public Function S_Check(st As String) As String
    If Trim(st) = "" Then
        S_Check = " "
    Else
        S_Check = Trim(st)
    End If
End Function

Sub SelectText(ctrIn As Control)
    ctrIn.SelStart = 0
    ctrIn.SelLength = Len(ctrIn.Text)
End Sub
Sub AppRunning()
    Dim sMsg As String
    If App.PrevInstance Then
        sMsg = App.EXEName & " already running! "
        MsgBox sMsg, 4096
        End
    End If
End Sub
Sub CenterMe(Frm As Form)
    Frm.Move Screen.Width / 2 - Frm.Width / 2, Screen.Height / 2 - Frm.Height / 2
End Sub
Sub Main()
    Screen.MousePointer = 11
    Dim AppName As String
    Dim inifilename$
    Dim KeyName As String
    Dim writeERR As Integer
    Dim tstr$
    Dim rval As Integer
    Dim ans%
    If App.PrevInstance Then
        End
    End If
    tstr$ = Command$
    If Trim(tstr$) = "" Then
        End
    End If
    'GoTo skip1
    GetEnvironment
    If Trim(G_Warehouse) = "" Then
        inifilename$ = App.Path & "\WMARS.INI"
        tstr$ = Format(Val(Command$) - 1, "00000000")
        rval = check_po(tstr$)
        If rval Then
            ans% = MsgBox("Po Req already exists, create another req?", vbYesNo)
            If ans% = vbNo Then
                On Error Resume Next
                AppName = "POREQQ"
                KeyName = "noswitch"
                writeERR = WriteINI(AppName, KeyName, "-1", inifilename$)
                DB.Close
                cnn.Close
                Screen.MousePointer = 0
                End
            Else
                AppName = "POREQQ"
                KeyName = "noswitch"
                writeERR = WriteINI(AppName, KeyName, " ", inifilename$)
            End If
        Else
            AppName = "POREQQ"
            KeyName = "noswitch"
            writeERR = WriteINI(AppName, KeyName, " ", inifilename$)
        End If
    End If
skip1:
    
    AppName = "SORDERS"
    KeyName = "value"
    G_Sorders = Val(ReadINI(AppName, KeyName, inifilename$))
    
    AppName = "CAW"
    KeyName = "value"
    G_Caw = Val(ReadINI(AppName, KeyName, inifilename$))
    
    POFname = Trim(Command$)
'    GetEnvironment
    MainFrm.Show
    MainFrm.Label1.Caption = "Please Wait"
    PO_Order
    Unload MainFrm
    
    On Error Resume Next
    DB.Close
    cnn.Close
    Screen.MousePointer = 0
    End
    
End Sub
Public Sub GetEnvironment()
    Dim AppName As String
    Dim inifilename$
    Dim KeyName As String
    Dim uTbl As Recordset
                
    Dim tbl As adodb.Recordset
    Dim Tblj As Recordset
    Dim UserName As String
    Dim rLen&
    Dim writeERR As Integer
    
    Call get_region_date
    If rdate = "M/dd/yy" Then
        rdate = "M/d/yyyy"
    End If
    
    inifilename$ = App.Path & "\WMARS.INI"
    AppName = "PONUM"
    KeyName = "value"
    G_PON = Val(ReadINI(AppName, KeyName, inifilename$))
    
    AppName = "POREQQ"
    KeyName = "vendor"
    G_VendCons = ReadINI(AppName, KeyName, inifilename$)
    KeyName = "pormemo"
    G_pormemotxt = ReadINI(AppName, KeyName, inifilename$)
    KeyName = "combo1"
    G_combo1 = ReadINI(AppName, KeyName, inifilename$)
    KeyName = "g_t_weight"
    G_T_Weight = Val(ReadINI(AppName, KeyName, inifilename$))
    KeyName = "g_warehouse"
    G_Warehouse = ReadINI(AppName, KeyName, inifilename$)
    KeyName = "pdate"
    G_pdate = ReadINI(AppName, KeyName, inifilename$)
    KeyName = "in_perform"
    In_Perform = Val(ReadINI(AppName, KeyName, inifilename$))
    KeyName = "label4"
    Label4 = ReadINI(AppName, KeyName, inifilename$)

    AppName = "CUSTOMER"
    KeyName = "badger"
    G_Cust = Val(ReadINI(AppName, KeyName, inifilename$))
    
    AppName = "BUYPERFORM"
    KeyName = "value"
    G_Buy_Perform = Val(ReadINI(AppName, KeyName, inifilename$))
        
    AppName = "ITEMXREF"
    KeyName = "use"
    G_I_VXref = Val(ReadINI(AppName, KeyName, inifilename$))
    
    AppName = "GFREE"
    KeyName = "value"
    G_Free = Val(ReadINI(AppName, KeyName, inifilename$))
    
    AppName = "Trans"
    KeyName = "trcon"
    G_TR_ConSwitch = Val(ReadINI(AppName, KeyName, inifilename$))
          
    AppName = "MARSUPDT_PATH"
    KeyName = "MarsDataPath"
    MarsDataPath = ReadINI(AppName, KeyName, inifilename$)
    KeyName = "poreq_asc"
    preq_path = ReadINI(AppName, KeyName, inifilename$)
    KeyName = "load_comma"
    G_Load_Comma = Val(ReadINI(AppName, KeyName, inifilename$))
        
    AppName = "VENDSPACE"
    KeyName = "value"
    G_Vspace = Val(ReadINI(AppName, KeyName, inifilename$))
    
    AppName = "SQL_Server"
    KeyName = "local"
    G_Local = Val(ReadINI(AppName, KeyName, inifilename$))
    
    lCnct = "Provider=Microsoft.Jet.OLEDB.4.0" & _
            ";Data Source=" & MarsDataPath & "\wmars.mdb" & _
            ";"
    On Error Resume Next
    Err = 0
    If Not G_Local Then
        Set lcnn = New adodb.Connection
        lcnn.ConnectionString = lCnct
        lcnn.Open
    End If
    AppName = "MARSUPDT_PATH"
    inifilename$ = App.Path & "\WMARS.INI"
    
    KeyName = "MarsDataPath"
    MarsDataPath = ReadINI(AppName, KeyName, inifilename$)
    
    If Trim(MarsDataBaseName) = "" Then
        AppName = "DataBase"
        KeyName = "name"
        MarsDataBaseName = ReadINI(AppName, KeyName, inifilename$)
        If Trim(MarsDataBaseName) = "" Then MarsDataBaseName = "wmars"
    End If
    
    If Trim(G_SQL_Server) = "" And Not G_Local Then
        AppName = "SQL_Server"
        KeyName = "sqlserv"
        G_SQL_Server = ReadINI(AppName, KeyName, inifilename$)
        KeyName = "sqlprovider"
        G_SQL_Provider = ReadINI(AppName, KeyName, inifilename$)
    End If
    On Error Resume Next
    Set MarsWorkspace = Workspaces(0)
    Set DB = MarsWorkspace.OpenDatabase(MarsDataPath & "\WMARS.MDB")
    
    If Not G_Local Then
        rLen& = 30
        UserName = Space$(30)
        writeERR = GetUserName(UserName, rLen&)
        G_SQL_User = Trim(Left$(UserName, rLen& - 1))

        Set uTbl = DB.OpenRecordset("select * from users where user_name = '" & Trim(G_SQL_User) & "'")
        If uTbl.EOF Then
            If G_SQL_Provider = "sqloledb" Then
                MsgBox "User not defined"
                End
            End If
        Else
            G_SQL_Pass = uTbl(1)
            Call cCnvEBCDICtoASCII(G_SQL_Pass)
        End If
        uTbl.Close
        
    End If
    If G_Local Then
        Cnct = lCnct
    Else
        Cnct = "Provider=sqloledb" & _
               ";Server=" & G_SQL_Server & _
               ";Initial Catalog=" & MarsDataBaseName & _
               ";User Id=" & G_SQL_User & _
               ";Password=" & G_SQL_Pass & _
               ";"
    End If
    Err = 0
    'If Not G_Local Then
        Set cnn = New adodb.Connection
        cnn.ConnectionString = Cnct
        cnn.Open
        If Err <> 0 Then
            MsgBox "error: " & Str$(Err) & "-" & Err.Description
        End If
        Set tbl = New adodb.Recordset
        tbl.ActiveConnection = cnn
        tbl.CursorLocation = adUseClient
    'End If
        
    If G_Local Then
        Set uTbl = DB.OpenRecordset("select * from system")
        G_PO_format = uTbl("po_format_s")
        uTbl.Close
    Else
        tbl.ActiveConnection = cnn
        tbl.CursorLocation = adUseClient
        tbl.Open "select * from system"
        G_PO_format = tbl("po_format_s")
        tbl.Close
        tbl.Open "select * from recalcparms"
        If Trim(tbl!d10_file_type) = "99" Then
            G_Load_Comma = True
        End If
        
    End If
    
End Sub
Function ReadINI(AppName, KeyName, filename1 As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(AppName, KeyName, " ", sRet, Len(sRet), filename1))
End Function
Function WriteINI(AppName, KeyName, NewString, FileName As String) As Long
    WriteINI = WritePrivateProfileString(AppName, CStr(KeyName), NewString, FileName)
End Function
Public Sub get_region_date()

   Dim LCID As Long
   Dim sNewFormat As String
   LCID = GetSystemDefaultLCID()
   rdate = GetUserLocaleInfo(LCID, LOCALE_SSHORTDATE)
End Sub
Public Sub Create_SORDERS()
    Dim tbl As Recordset
    On Error Resume Next
    'DB.Execute "drop table sorders"
    Err = 0
    Set tbl = DB.OpenRecordset("sorders")
    If Err <> 0 Then
        tbl.Close
        Err = 0
        On Error GoTo 0
        Dim r_txt As String
        r_txt = " "
        r_txt = r_txt & " CREATE TABLE SORDERS "
        r_txt = r_txt & " ( "
        r_txt = r_txt & "  item_num_im      char (28) , "
        r_txt = r_txt & "  item_desc_im     char (30) , "
        r_txt = r_txt & "  unit_meas_im     char (4)  , "
        r_txt = r_txt & "  whse_im          char (5)  , "
        r_txt = r_txt & "  abc_class_im     char (1)  , "
        r_txt = r_txt & "  qty_on_hand_im   double ,"
        r_txt = r_txt & "  qty_on_ord_im    double ,"
        r_txt = r_txt & "  qty_cmt_im       double ,"
        r_txt = r_txt & "  qty_back_im      double ,"
        r_txt = r_txt & "  min_ot_im        double ,"
        r_txt = r_txt & "  dump_flg_im      char (1)  , "
        r_txt = r_txt & "  item_date_im     Date , "
        r_txt = r_txt & "  price_brk_im     double ,"
        r_txt = r_txt & "  lot_size_im      double ,"
        r_txt = r_txt & "  min_ord_qty_im   double ,"
        r_txt = r_txt & "  unit_cost_im     double ,"
        r_txt = r_txt & "  brand_im         char (4)  , "
        r_txt = r_txt & "  brand_2_im       char (4)  , "
        r_txt = r_txt & "  s_vendor_im      char (9)  , "
        r_txt = r_txt & "  vendor_im        char (9)  , "
        r_txt = r_txt & "  h_im_1           double ,"
        r_txt = r_txt & "  h_im_2           double ,"
        r_txt = r_txt & "  h_im_3           double ,"
        r_txt = r_txt & "  h_im_4           double ,"
        r_txt = r_txt & "  h_im_5           double ,"
        r_txt = r_txt & "  h_im_6           double ,"
        r_txt = r_txt & "  h_im_7           double ,"
        r_txt = r_txt & "  h_im_8           double ,"
        r_txt = r_txt & "  h_im_9           double ,"
        r_txt = r_txt & "  h_im_10          double ,"
        r_txt = r_txt & "  h_im_11          double ,"
        r_txt = r_txt & "  h_im_12          double ,"
        r_txt = r_txt & "  s_ind_im         char (3)  , "
        r_txt = r_txt & "  mtd_im           double ,"
        r_txt = r_txt & "  month_data_im    char (2)  , "
        r_txt = r_txt & "  wght_im          double ,"
        r_txt = r_txt & "  lead_time_im     double ,"
        r_txt = r_txt & "  ord_freq_im      double ,"
        r_txt = r_txt & "  safety_stock_im  double ,"
        r_txt = r_txt & "  qty_comm_im      double ,"
        r_txt = r_txt & "  re_ord_low_im    double ,"
        r_txt = r_txt & "  re_ord_nor_im    double ,"
        r_txt = r_txt & "  re_ord_hi_im     double ,"
        r_txt = r_txt & "  re_ord_max_im    double ,"
        r_txt = r_txt & "  qty_avl_im       double ,"
        r_txt = r_txt & "  net_qty_avl_im   double ,"
        r_txt = r_txt & "  min_rop_im       double ,"
        r_txt = r_txt & "  max_rop_im       double ,"
        r_txt = r_txt & "  mon_avrg_im      double ,"
        r_txt = r_txt & "  safety_fac_im    double ,"
        r_txt = r_txt & "  ord_qty_im       double ,"
        r_txt = r_txt & "  de_season_avg_im double ,"
        r_txt = r_txt & "  season_avg_im    double ,"
        r_txt = r_txt & "  mean_abs_dev_im  double ,"
        r_txt = r_txt & "  lot_sens_im      double ,"
        r_txt = r_txt & "  ord_fill_im      double ,"
        r_txt = r_txt & "  fcast_im_1       double ,"
        r_txt = r_txt & "  fcast_im_2       double ,"
        r_txt = r_txt & "  fcast_im_3       double ,"
        r_txt = r_txt & "  fcast_im_4       double ,"
        r_txt = r_txt & "  fcast_im_5       double ,"
        r_txt = r_txt & "  fcast_im_6       double ,"
        r_txt = r_txt & "  fcast_im_7       double ,"
        r_txt = r_txt & "  fcast_im_8       double ,"
        r_txt = r_txt & "  fcast_im_9       double ,"
        r_txt = r_txt & "  fcast_im_10      double ,"
        r_txt = r_txt & "  fcast_im_11      double ,"
        r_txt = r_txt & "  fcast_im_12      double ,"
        r_txt = r_txt & "  stat_im          char (2)  , "
        r_txt = r_txt & "  mon_avg_im       double ,"
        r_txt = r_txt & "  mot_flg_im       char (1)  , "
        r_txt = r_txt & "  att_flg_im       char (1)  , "
        r_txt = r_txt & "  eoq_flg_im       char (1)  , "
        r_txt = r_txt & "  eoq_flg_p_im     char (1)  , "
        r_txt = r_txt & "  eoq_b_im         double ,"
        r_txt = r_txt & "  eoq_im           double ,"
        r_txt = r_txt & "  eoq_p_im         double ,"
        r_txt = r_txt & "  filt_flg_im      double ,"
        r_txt = r_txt & "  lock_im          char (1)  , "
        r_txt = r_txt & "  hl_im            char (1)  , "
        r_txt = r_txt & "  t_flg_im         char (1)  , "
        r_txt = r_txt & "  l_h_im_1         double ,"
        r_txt = r_txt & "  l_h_im_2         double ,"
        r_txt = r_txt & "  l_h_im_3         double ,"
        r_txt = r_txt & "  l_h_im_4         double ,"
        r_txt = r_txt & "  l_h_im_5         double ,"
        r_txt = r_txt & "  l_h_im_6         double ,"
        r_txt = r_txt & "  l_h_im_7         double ,"
        r_txt = r_txt & "  l_h_im_8         double ,"
        r_txt = r_txt & "  l_h_im_9         double ,"
        r_txt = r_txt & "  l_h_im_10        double ,"
        r_txt = r_txt & "  l_h_im_11        double ,"
        r_txt = r_txt & "  l_h_im_12        double ,"
        r_txt = r_txt & "  p_h_im_1         double ,"
        r_txt = r_txt & "  p_h_im_2         double ,"
        r_txt = r_txt & "  p_h_im_3         double ,"
        r_txt = r_txt & "  p_h_im_4         double ,"
        r_txt = r_txt & "  p_h_im_5         double ,"
        r_txt = r_txt & "  p_h_im_6         double ,"
        r_txt = r_txt & "  p_h_im_7         double ,"
        r_txt = r_txt & "  p_h_im_8         double ,"
        r_txt = r_txt & "  p_h_im_9         double ,"
        r_txt = r_txt & "  p_h_im_10        double ,"
        r_txt = r_txt & "  p_h_im_11        double ,"
        r_txt = r_txt & "  p_h_im_12        double ,"
        r_txt = r_txt & "  r_opt_im         double ,"
        r_txt = r_txt & "  r_monf_im        char (10) , "
        r_txt = r_txt & "  r_mont_im        char (10) , "
        r_txt = r_txt & "  r_cper_im        double ,"
        r_txt = r_txt & "  r_lper_im        double ,"
        r_txt = r_txt & "  r_pper_im        double ,"
        r_txt = r_txt & "  r_per_im         double ,"
        r_txt = r_txt & "  r_rot_im         double ,"
        r_txt = r_txt & "  lot_size_2_im    double, "
        r_txt = r_txt & "  lot_sens_2_im    double,  "
        r_txt = r_txt & "  commited_im    char(1),  "
        r_txt = r_txt & "  ponum_im    char(9)  "
        r_txt = r_txt & " )"
        DB.Execute r_txt
        DB.Execute " CREATE UNIQUE INDEX index_load ON SORDERS (item_num_im, whse_im,vendor_im,ponum_im)"
        DB.Execute " CREATE INDEX IndexIW ON SORDERS (item_num_im, whse_im)"
        DB.Execute " CREATE INDEX POKEY ON SORDERS (ponum_im) "
    End If
End Sub




Public Sub Create_GORDERS()

    On Error Resume Next
    DB.Execute "drop table gorders"
    On Error GoTo 0
    Dim r_txt As String
    r_txt = " "
    r_txt = r_txt & " CREATE TABLE GORDERS "
    r_txt = r_txt & " ( "
    r_txt = r_txt & "  item_num_im      char (28) , "
    r_txt = r_txt & "  item_desc_im     char (30) , "
    r_txt = r_txt & "  unit_meas_im     char (4)  , "
    r_txt = r_txt & "  whse_im          char (5)  , "
    r_txt = r_txt & "  abc_class_im     char (1)  , "
    r_txt = r_txt & "  qty_on_hand_im   double ,"
    r_txt = r_txt & "  qty_on_ord_im    double ,"
    r_txt = r_txt & "  qty_cmt_im       double ,"
    r_txt = r_txt & "  qty_back_im      double ,"
    r_txt = r_txt & "  min_ot_im        double ,"
    r_txt = r_txt & "  dump_flg_im      char (1)  , "
    r_txt = r_txt & "  item_date_im     Date , "
    r_txt = r_txt & "  price_brk_im     double ,"
    r_txt = r_txt & "  lot_size_im      double ,"
    r_txt = r_txt & "  min_ord_qty_im   double ,"
    r_txt = r_txt & "  unit_cost_im     double ,"
    r_txt = r_txt & "  brand_im         char (4)  , "
    r_txt = r_txt & "  brand_2_im       char (4)  , "
    r_txt = r_txt & "  s_vendor_im      char (9)  , "
    r_txt = r_txt & "  vendor_im        char (9)  , "
    r_txt = r_txt & "  h_im_1           double ,"
    r_txt = r_txt & "  h_im_2           double ,"
    r_txt = r_txt & "  h_im_3           double ,"
    r_txt = r_txt & "  h_im_4           double ,"
    r_txt = r_txt & "  h_im_5           double ,"
    r_txt = r_txt & "  h_im_6           double ,"
    r_txt = r_txt & "  h_im_7           double ,"
    r_txt = r_txt & "  h_im_8           double ,"
    r_txt = r_txt & "  h_im_9           double ,"
    r_txt = r_txt & "  h_im_10          double ,"
    r_txt = r_txt & "  h_im_11          double ,"
    r_txt = r_txt & "  h_im_12          double ,"
    r_txt = r_txt & "  s_ind_im         char (3)  , "
    r_txt = r_txt & "  mtd_im           double ,"
    r_txt = r_txt & "  month_data_im    char (2)  , "
    r_txt = r_txt & "  wght_im          double ,"
    r_txt = r_txt & "  lead_time_im     double ,"
    r_txt = r_txt & "  ord_freq_im      double ,"
    r_txt = r_txt & "  safety_stock_im  double ,"
    r_txt = r_txt & "  qty_comm_im      double ,"
    r_txt = r_txt & "  re_ord_low_im    double ,"
    r_txt = r_txt & "  re_ord_nor_im    double ,"
    r_txt = r_txt & "  re_ord_hi_im     double ,"
    r_txt = r_txt & "  re_ord_max_im    double ,"
    r_txt = r_txt & "  qty_avl_im       double ,"
    r_txt = r_txt & "  net_qty_avl_im   double ,"
    r_txt = r_txt & "  min_rop_im       double ,"
    r_txt = r_txt & "  max_rop_im       double ,"
    r_txt = r_txt & "  mon_avrg_im      double ,"
    r_txt = r_txt & "  safety_fac_im    double ,"
    r_txt = r_txt & "  ord_qty_im       double ,"
    r_txt = r_txt & "  de_season_avg_im double ,"
    r_txt = r_txt & "  season_avg_im    double ,"
    r_txt = r_txt & "  mean_abs_dev_im  double ,"
    r_txt = r_txt & "  lot_sens_im      double ,"
    r_txt = r_txt & "  ord_fill_im      double ,"
    r_txt = r_txt & "  fcast_im_1       double ,"
    r_txt = r_txt & "  fcast_im_2       double ,"
    r_txt = r_txt & "  fcast_im_3       double ,"
    r_txt = r_txt & "  fcast_im_4       double ,"
    r_txt = r_txt & "  fcast_im_5       double ,"
    r_txt = r_txt & "  fcast_im_6       double ,"
    r_txt = r_txt & "  fcast_im_7       double ,"
    r_txt = r_txt & "  fcast_im_8       double ,"
    r_txt = r_txt & "  fcast_im_9       double ,"
    r_txt = r_txt & "  fcast_im_10      double ,"
    r_txt = r_txt & "  fcast_im_11      double ,"
    r_txt = r_txt & "  fcast_im_12      double ,"
    r_txt = r_txt & "  stat_im          char (2)  , "
    r_txt = r_txt & "  mon_avg_im       double ,"
    r_txt = r_txt & "  mot_flg_im       char (1)  , "
    r_txt = r_txt & "  att_flg_im       char (1)  , "
    r_txt = r_txt & "  eoq_flg_im       char (1)  , "
    r_txt = r_txt & "  eoq_flg_p_im     char (1)  , "
    r_txt = r_txt & "  eoq_b_im         double ,"
    r_txt = r_txt & "  eoq_im           double ,"
    r_txt = r_txt & "  eoq_p_im         double ,"
    r_txt = r_txt & "  filt_flg_im      double ,"
    r_txt = r_txt & "  lock_im          char (1)  , "
    r_txt = r_txt & "  hl_im            char (1)  , "
    r_txt = r_txt & "  t_flg_im         char (1)  , "
    r_txt = r_txt & "  l_h_im_1         double ,"
    r_txt = r_txt & "  l_h_im_2         double ,"
    r_txt = r_txt & "  l_h_im_3         double ,"
    r_txt = r_txt & "  l_h_im_4         double ,"
    r_txt = r_txt & "  l_h_im_5         double ,"
    r_txt = r_txt & "  l_h_im_6         double ,"
    r_txt = r_txt & "  l_h_im_7         double ,"
    r_txt = r_txt & "  l_h_im_8         double ,"
    r_txt = r_txt & "  l_h_im_9         double ,"
    r_txt = r_txt & "  l_h_im_10        double ,"
    r_txt = r_txt & "  l_h_im_11        double ,"
    r_txt = r_txt & "  l_h_im_12        double ,"
    r_txt = r_txt & "  p_h_im_1         double ,"
    r_txt = r_txt & "  p_h_im_2         double ,"
    r_txt = r_txt & "  p_h_im_3         double ,"
    r_txt = r_txt & "  p_h_im_4         double ,"
    r_txt = r_txt & "  p_h_im_5         double ,"
    r_txt = r_txt & "  p_h_im_6         double ,"
    r_txt = r_txt & "  p_h_im_7         double ,"
    r_txt = r_txt & "  p_h_im_8         double ,"
    r_txt = r_txt & "  p_h_im_9         double ,"
    r_txt = r_txt & "  p_h_im_10        double ,"
    r_txt = r_txt & "  p_h_im_11        double ,"
    r_txt = r_txt & "  p_h_im_12        double ,"
    r_txt = r_txt & "  r_opt_im         double ,"
    r_txt = r_txt & "  r_monf_im        char (10) , "
    r_txt = r_txt & "  r_mont_im        char (10) , "
    r_txt = r_txt & "  r_cper_im        double ,"
    r_txt = r_txt & "  r_lper_im        double ,"
    r_txt = r_txt & "  r_pper_im        double ,"
    r_txt = r_txt & "  r_per_im         double ,"
    r_txt = r_txt & "  r_rot_im         double ,"
    r_txt = r_txt & "  lot_size_2_im    double, "
    r_txt = r_txt & "  lot_sens_2_im    double,  "
    r_txt = r_txt & "  commited_im    char(1),  "
    r_txt = r_txt & "  rep_item_im    char(28)  "
    r_txt = r_txt & " )"
    DB.Execute r_txt
    DB.Execute " CREATE UNIQUE INDEX index_load ON GORDERS (item_num_im, whse_im,vendor_im)"
    DB.Execute " CREATE INDEX IndexIW ON GORDERS (item_num_im, whse_im)"
    DB.Execute " CREATE INDEX index_profile ON GORDERS (vendor_im, whse_im)"
    DB.Execute " CREATE INDEX item_num_im ON GORDERS (item_num_im)"
    DB.Execute " CREATE INDEX PrimaryKey ON GORDERS (whse_im, item_num_im,vendor_im) "
    DB.Execute " CREATE INDEX whse_im ON GORDERS (whse_im)"
    DB.Execute " CREATE INDEX Vendor_Whse ON GORDERS (vendor_im, whse_im)"
    DB.Execute " CREATE INDEX Index_v ON GORDERS (vendor_im)"
    DB.Execute " CREATE INDEX Index_Calc ON GORDERS (dump_flg_im)"
End Sub


