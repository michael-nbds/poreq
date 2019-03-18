Attribute VB_Name = "Module2"
Public Sub PO_Order()
    Dim preq_path As String
    MainFrm.Label1.Caption = "P.O.Req. output in progress..."
    Dim sav_vendor_im As String
    Dim sav_whse As String
    Dim sav_g_pon As Double
    Dim o_qty As Double
    Dim o_wght As Double
    Dim c_val As Double
    Dim t_str As String
    Dim s_str As String
    Dim i%
    Dim tblc As Recordset
    Dim atblc As Recordset
    Dim bftbl As adodb.Recordset
    Dim oTbl As Recordset
    Dim mTbl As Recordset
    Dim gTbl As Recordset
    Dim sTbl As Recordset
    Dim cnt%
    Dim GG_TBL As Recordset
    Dim FTbl As Recordset
    Dim vTbl As Recordset
    Dim vtbls As Recordset
    Dim m_ponum(10) As Long
    Dim sm_ponum(10) As Long
    Dim t_memo As String
    Dim ctxt As String
    Dim btxt As String
    Dim w_loop As Integer
    Dim FGood As Integer
    Dim FGamt As Double
    Dim tvend As String
    Dim q$
    Dim filnum As Integer
    Create_SORDERS
    
    filnum = FreeFile
    On Error Resume Next
    Open App.Path & "\poreq.txt" For Input As #filnum
    If Err = 53 Then
        Close #filnum
        Open App.Path & "\poreq.txt" For Output As #filnum
        Print #filnum, "PO Num  Vendor   Item              Location  Date "
    Else
        Close #filnum
        Open App.Path & "\poreq.txt" For Append As #filnum
    End If
    Err = 0
    q$ = Chr$(34)
    
    Set bftbl = New adodb.Recordset
    bftbl.CursorType = adOpenStatic
    bftbl.LockType = adLockOptimistic
    bftbl.CursorLocation = adUseClient
    Err = 0
    Set tblc = DB.OpenRecordset("orders")
    If tblc.EOF Or Err <> 0 Then
        tblc.Close
        Set tblc = Nothing
        MsgBox "No items for Requisition"
        Exit Sub
    Else
        tblc.Close
    End If
    Set GG_TBL = DB.OpenRecordset("ovendor")
    GG_TBL.Index = "Index_v_w"
    Set vTbl = DB.OpenRecordset("vendor")
    vTbl.Index = "Index_v_w"
    If G_Warehouse <> "" Then
        On Error Resume Next
        DB.Execute "DROP TABLE MORDERS"
        DB.Execute "SELECT ORDERS.* INTO MORDERS FROM ORDERS ORDER BY item_num_im, whse_im"
        DB.Execute "delete from morders where 1=1"
        Set mTbl = DB.OpenRecordset("morders")
        
        Call Create_GORDERS
        Set oTbl = DB.OpenRecordset("select item_num_im,count(item_num_im) " & _
                                    "from orders where ord_qty_im > 0 " & _
                                    "group by item_num_im")
        If oTbl.EOF Then
            oTbl.Close
            GoTo skip1
        End If
        oTbl.MoveFirst
        On Error Resume Next
        'On Error GoTo 0
        'Print #filnum, "in g_whse loop"; oTbl(0); "^"; oTbl(1)
        Do While Not oTbl.EOF
            Print #filnum, "in g_whse loop"; oTbl(0); "^"; oTbl(1)
            If oTbl(1) > 1 Then
                DB.Execute "delete from gorders where 1=1"
                DB.Execute "insert into gorders select * from orders where " & _
                           "item_num_im = '" & oTbl(0) & "'"
                Set tblc = DB.OpenRecordset("select sum(ord_qty_im) from gorders")
                Set gTbl = DB.OpenRecordset("select * from orders where " & _
                                            "item_num_im = '" & oTbl(0) & "'")
                Print #filnum, "otbl loop"; gTbl(0); "^"; tblc(0)
                If Not gTbl.EOF Then
                    cnt% = 0
                    Do While Not gTbl.EOF
                        Print #filnum, "in morder add"; gTbl(0); "^"; tblc(0); "^"; cnt%
                        If cnt% = 0 Then
                            mTbl.AddNew
                            For i% = 0 To gTbl.Fields.Count
                                mTbl(i%) = gTbl(i%)
                            Next
                            mTbl!ord_qty_im = tblc(0)
                            mTbl.Update
                            gTbl.Delete
                        Else
                            gTbl.Delete
                        End If
                        cnt% = cnt% + 1
                        gTbl.MoveNext
                        DoEvents
                    Loop
                End If
            End If
            DoEvents
            oTbl.MoveNext
        Loop
        DB.Execute "insert into orders select * from morders "
        DB.Execute "delete from orders where ord_qty_im = 0"
        DB.Execute "update orders set whse_im = '" & G_Warehouse & "' where 1=1"
    End If
skip1:
    
    Set tblc = DB.OpenRecordset("totals")
    If Not tblc.EOF Then
        o_qty = tblc!t_qty
        o_wght = tblc!t_wght
    End If
    tblc.Close
    '------------------------------------- new logic added here
    w_loop = 0
    If G_T_Weight > 0 Then
        w_loop = o_wght / G_T_Weight
    End If
    Screen.MousePointer = 11
    On Error Resume Next
    If w_loop > 1 Then
        Call PO_Order_Mult
        Exit Sub
    End If
    '------------------------------------- end of new logic
    Screen.MousePointer = 11
    If Len(Trim(G_combo1)) = 0 Then G_combo1 = "ZZ - UNKNOWN"
    If Len(G_pormemotxt) = 0 Then G_pormemotxt = Space(1)
    Dim dn As Recordset
    
    POFname = Trim(POFname)
    Set sTbl = DB.OpenRecordset("sorders")
    
    '------------------------------------------------------------------
    ' added logic for saved orders table
    If G_Sorders Then
        Set dn = DB.OpenRecordset("select * from orders")
        If Not dn.EOF Then
            Do While Not dn.EOF
                sTbl.AddNew
                For i% = 0 To dn.Fields.Count - 1
                    sTbl(i%) = dn(i%)
                Next
                sTbl("ponum_im") = Format$(G_PON, "00000000")
                sTbl.Update
                dn.MoveNext
                DoEvents
            Loop
        End If
    End If
    '------------------------------------------------------------------
    dn.Close
    Dim FilePO As Integer
    Dim pos%
    pos% = 0
    If w_loop > 1 Then
        If G_Buy_Perform Then
            query = "SELECT * FROM MORDERS "
        Else
            query = "SELECT * FROM MORDERS WHERE ord_qty_im > 0 "
        End If
    Else
        If G_Buy_Perform Then
            query = "SELECT * FROM ORDERS "
        Else
            query = "SELECT * FROM ORDERS WHERE ord_qty_im > 0 "
        End If
    End If
    pos% = InStr(1, query, "ORDER BY")
    If pos% > 0 Then
        query = Left$(query, pos% - 1)
        pos% = 0
        pos% = InStr(1, query, "WHERE")
        If pos% = 0 Then
            query = query & " WHERE 1 = 1 "
        End If
        query = query & " AND ord_qty_im > 0 "
    Else
        If G_Buy_Perform Then
            query = "SELECT * FROM ORDERS "
        Else
            query = "SELECT * FROM ORDERS WHERE ord_qty_im > 0 "
        End If
    End If
    query = query & "ORDER BY  vendor_im, whse_im, item_num_im"
    inifilename$ = App.Path + "\WMARS.INI"
    AppName = "MARSUPDT_PATH"
    KeyName = "poreq_asc"
    preq_path = ReadINI(AppName, KeyName, inifilename$)
    
    DB.Execute "select    "
    FilePO = FreeFile
    Open (preq_path & "\" & POFname & ".req") For Output As #FilePO
    'Dim dn As Recordset
    Set dn = DB.OpenRecordset(query)
    If dn.EOF Then
        Close #FilePO
        dn.Close
        If Not G_TR_ConSwitch Then
            Call Dump_TPor
        End If
        Exit Sub
    End If
    rec_counter = 0
    PoMade = True
    On Error Resume Next
If G_PO_format = -1 Then
    If Not dn.EOF Then
       sav_vendor_im = dn!vendor_im
       Do While Not dn.EOF          'Old format
            If sav_vendor_im <> dn!vendor_im Then
                rec_counter = 0
                sav_vendor_im = dn!vendor_im
            End If
            LblStatus.Caption = "A-Working on: " & dn!item_num_im
            DoEvents
            rec_counter = rec_counter + 1
            Print #FilePO, Format$(G_PON, "0000000");
            Print #FilePO, Format$(rec_counter, "0000");
            Print #FilePO, Tab(12); Trim(dn!item_num_im);
            Print #FilePO, Tab(40); Trim(dn!vendor_im);
            Print #FilePO, Tab(49); Trim(dn!whse_im);
            Print #FilePO, Tab(54); Trim(dn!item_desc_im);
            Print #FilePO, Tab(84); Trim(dn!unit_meas_im);
            Print #FilePO, Tab(88); Format$(dn!unit_cost_im, "000000.00000");
            Print #FilePO, Tab(100); Format$(dn!ord_qty_im, "0000000")
            On Error Resume Next
            dn.MoveNext
            If dn.EOF Then Exit Do
        Loop
    End If
Else           'new MARS-95 format, long version
    If Not dn.EOF Then
        FGamt = 0
        FGood = False
        If Trim(G_VendCons) <> "" And Left(G_combo1, 2) = "77" Then    '---------------changed
            sav_vendor_im = G_VendCons
        Else
            sav_vendor_im = dn!vendor_im
        End If
        sav_whse = dn!whse_im
                
        Do While Not dn.EOF
            FGamt = 0
            FGood = False
            'If Trim(dn!item_num_im) = "AX005" Then
            '    Debug.Print
            'End If
            If G_Buy_Perform Then
                If dn!ord_qty_im = 0 Then
                    GoTo get_next_first
                End If
            End If
            If sav_vendor_im <> dn!vendor_im Then
                rec_counter = 0
                sav_vendor_im = dn!vendor_im
            End If
            rec_counter = rec_counter + 1
            'LblStatus.Caption = "A-Working on: " & dn!item_num_im
            Print #filnum, Format$(G_PON, "0000000"); " "; PO_string(dn!vendor_im, 9); PO_string(dn!item_num_im, 18); PO_string(dn!whse_im, 10); Now
            DoEvents
            If G_Load_Comma Then
                On Error GoTo 0
                q_txt = ""
                q_txt = q_txt + q$ + Format$(G_PON, "0000000") + q$ + ","
                q_txt = q_txt + q$ + Format$(rec_counter, "0000") + q$ + ","
                q_txt = q_txt + q$ + PO_string(Trim(dn!item_num_im), 28) + q$ + ","
                If G_Vspace Then
                    If Trim(G_VendCons) <> "" And Left(G_combo1, 2) = "77" Then
                        tvend = Replace(G_VendCons, "^", " ")
                        q_txt = q_txt + q$ + PO_string(Trim(tvend), 9) + q$ + ","
                        'q_txt = q_txt + PO_string(Trim(G_VendCons), 9)
                    Else
                        tvend = Replace(dn!vendor_im, "^", " ")
                        q_txt = q_txt + q$ + PO_string(Trim(tvend), 9) + q$ + ","
                        'q_txt = q_txt + PO_string(Trim(dn!vendor_im), 9)
                    End If
                Else
                    If Trim(G_VendCons) <> "" And Left(G_combo1, 2) = "77" Then
                        q_txt = q_txt + q$ + PO_string(Trim(G_VendCons), 9) + q$ + ","
                    Else
                        q_txt = q_txt + q$ + PO_string(Trim(dn!vendor_im), 9) + q$ + ","
                    End If
                End If
                'q_txt = q_txt + q$ + PO_string(Trim(dn!vendor_im), 9) + q$ + ","
                q_txt = q_txt & q$ + PO_string(Trim(dn!whse_im), 5) + q$ + ","
                q_txt = q_txt + q$ + PO_string(Trim(dn!item_desc_im), 30) + q$ + ","
                q_txt = q_txt + q$ + PO_string(Trim(dn!unit_meas_im), 4) + q$ + ","
                'add xdesc logic here
                Set atblc = DB.OpenRecordset("select * from xdesc where " & _
                                             "item = '" & dn!item_num_im & "' and " & _
                                             "whse = '" & dn!whse_im & "'")
                If atblc.EOF Then
                    q_txt = q_txt + Format$(dn!unit_cost_im, "000000.00000") + ","
                Else
                    s_str = atblc!descr
                    For i% = 1 To Len(s_str)
                        If Mid$(s_str, i%, 1) = Chr$(10) Or Mid$(s_str, i%, 1) = Chr$(13) Then
                            Mid$(s_str, i%, 1) = " "
                        End If
                    Next
                    If InStr(s_str, "ACT COST: $") = 0 Then
                        q_txt = q_txt + Format$(dn!unit_cost_im, "000000.00000") + ","
                    Else
                        i% = InStr(s_str, "ACT COST: $") + 11
                        If InStr(i%, Trim(s_str), " ") > 0 Or Right$(Trim(s_str), 1) <> "$" Then
                            If InStr(i%, Trim(s_str), " ") = 0 Then
                                If InStr(Mid$(s_str, i%, 15), "/") = 0 Then
                                    c_val = Val(Mid$(s_str, i%))
                                End If
                            Else
                                If InStr(Mid$(s_str, i%, 15), "/") = 0 Then
                                    c_val = Val(Mid$(s_str, i%, InStr(i%, Trim(s_str), " ") - 1))
                                End If
                            End If
                            If InStr(Mid$(s_str, i%, 15), "/") = 0 Then
                                q_txt = q_txt + Format$(c_val, "000000.00000") + ","
                            Else
                                i% = InStr(s_str, "C/W LB PRICE")
                                If i% = 0 Then
                                    q_txt = q_txt + Format$(dn!unit_cost_im, "000000.00000") + q$ + ","
                                Else
                                    i% = InStr(s_str, "C/W LB PRICE") + 12
                                    c_val = Val(Mid$(s_str, i%, 8))
                                    q_txt = q_txt + Format$(c_val, "000000.00000") + ","
                                End If
                            End If
                        End If
                    End If
                End If
                q_txt = q_txt & Format$(dn!ord_qty_im, "0000000") + ","
                q_txt = q_txt + q$ + dn!abc_class_im + q$ + ","
                q_txt = q_txt + Format$(dn!wght_im, "0000000.000") + ","
                q_txt = q_txt + q$ + Format$(dn!att_flg_im, "@") + q$ + ","
                If Trim(G_VendCons) <> "" And Left(G_combo1, 2) = "77" Then
                    If G_Local Then
                        vTbl.Seek "=", Trim(G_VendCons), dn!whse_im
                        If vTbl.NoMatch Then
                            Llt = 0
                            ord_fr = 0
                        Else
                            Select Case dn!abc_class_im
                                Case "A"
                                    Llt = vTbl!a_lead_time_v
                                    ord_fr = vTbl!a_ord_fr_v
                                Case "B"
                                    Llt = vTbl!b_lead_time_v
                                    ord_fr = vTbl!b_ord_fr_v
                                Case "C"
                                    Llt = vTbl!c_lead_time_v
                                    ord_fr = vTbl!c_ord_fr_v
                                Case Else
                                    Llt = vTbl!a_lead_time_v
                                    ord_fr = vTbl!a_ord_fr_v
                                    'Llt = 0
                                    'ord_fr = 0
                            End Select
                        End If
                    Else
                        Set vtbls = DB.OpenRecordset("select * from ovendor where " & _
                                                     "vendor_v='" & Trim(G_VendCons) & "' and " & _
                                                     "whse_v='" & Trim(dn("whse_im")) & "'")
                        If vtbls.EOF Then
                            Llt = 0
                            ord_fr = 0
                        Else
                            Select Case dn!abc_class_im
                                Case "A"
                                    Llt = vtbls!a_lead_time_v
                                    ord_fr = vtbls!a_ord_fr_v
                                Case "B"
                                    Llt = vtbls!b_lead_time_v
                                    ord_fr = vtbls!b_ord_fr_v
                                Case "C"
                                    Llt = vtbls!c_lead_time_v
                                    ord_fr = vtbls!c_ord_fr_v
                                Case Else
                                    Llt = vtbls!a_lead_time_v
                                    ord_fr = vtbls!a_ord_fr_v
                            End Select
                        End If
                    End If
                Else
                    GG_TBL.Seek "=", dn!vendor_im, dn!whse_im
                    If GG_TBL.NoMatch Then
                        Set vtbls = DB.OpenRecordset("select * from ovendor where " & _
                                                     "vendor_v='" & dn!vendor_im & "' and " & _
                                                     "whse_v='" & dn!whse_im & "'")
                        If Not vtbls.EOF Then
                            Select Case dn!abc_class_im
                                Case "A"
                                    Llt = vtbls!a_lead_time_v
                                    ord_fr = vtbls!a_ord_fr_v
                                Case "B"
                                    Llt = vtbls!b_lead_time_v
                                    ord_fr = vtbls!b_ord_fr_v
                                Case "C"
                                    Llt = vtbls!c_lead_time_v
                                    ord_fr = vtbls!c_ord_fr_v
                                Case Else
                                    Llt = vtbls!a_lead_time_v
                                    ord_fr = vtbls!a_ord_fr_v
                                    'Llt = 0
                                    'ord_fr = 0
                            End Select
                        End If
                        'Else
                        '    Llt = 0
                        '    ord_fr = 0
                        'End If
                    Else
                        Select Case dn!abc_class_im
                            Case "A"
                                Llt = GG_TBL!a_lead_time_v
                                ord_fr = GG_TBL!a_ord_fr_v
                            Case "B"
                                Llt = GG_TBL!b_lead_time_v
                                ord_fr = GG_TBL!b_ord_fr_v
                            Case "C"
                                Llt = GG_TBL!c_lead_time_v
                                ord_fr = GG_TBL!c_ord_fr_v
                            Case Else
                                Llt = GG_TBL!a_lead_time_v
                                ord_fr = GG_TBL!a_ord_fr_v
                                'Llt = 0
                                'ord_fr = 0
                        End Select
                    End If
                End If
'                GG_TBL.Seek "=", dn!vendor_im, dn!whse_im
'                If GG_TBL.NoMatch Then
'                    Llt = 0
'                    ord_fr = 0
'                Else
'                    Select Case dn!abc_class_im
'                        Case "A"
'                            Llt = GG_TBL!a_lead_time_v
'                            ord_fr = GG_TBL!a_ord_fr_v
'                        Case "B"
'                            Llt = GG_TBL!b_lead_time_v
'                            ord_fr = GG_TBL!b_ord_fr_v
'                        Case "C"
'                            Llt = GG_TBL!c_lead_time_v
'                            ord_fr = GG_TBL!c_ord_fr_v
'                        Case Else
'                            Llt = 0
'                            ord_fr = 0
'                    End Select
'                End If
                'q_txt = q_txt & PO_string(Llt, 6)
                'q_txt = q_txt & Format(Llt, "####.##") + ","
                'q_txt = q_txt & Format(dn!lead_time_im, "####.##") + ","
                If dn!abc_class_im = "X" Then
                    q_txt = q_txt & PO_string(Llt, 6) + ","
                Else
                    q_txt = q_txt & PO_string(dn!lead_time_im, 6) + ","
                End If
                'q_txt = q_txt & PO_string(Llt, 6)
                If Trim(G_VendCons) <> "" And Left(G_combo1, 2) = "77" Then
                    If G_Local Then
                        q_txt = q_txt & q$ + PO_string(Trim(vTbl!org_v), 30) + q$ + ","
                        q_txt = q_txt & q$ + PO_string(Trim(vTbl!address1_v), 30) + q$ + ","
                        q_txt = q_txt & q$ + PO_string(Trim(vTbl!address2_v), 30) + q$ + ","
                        q_txt = q_txt & q$ + PO_string(Trim(vTbl!city_v), 15) + q$ + ","
                        q_txt = q_txt & q$ + PO_string(Trim(vTbl!state_v), 2) + q$ + ","
                        q_txt = q_txt & q$ + PO_string(Trim(vTbl!zip_v), 10) + q$ + ","
                        q_txt = q_txt & q$ + PO_string(Trim(vTbl!phone_v), 15) + q$ + ","
                        q_txt = q_txt & q$ + PO_string(Trim(vTbl!fax_v), 15) + q$ + ","
                        q_txt = q_txt & q$ + PO_string(Trim(vTbl!attn_v), 30) + q$ + ","
                        q_txt = q_txt & q$ + PO_string(Trim(vTbl!buyer_v), 15) + q$ + ","
                    Else
                        q_txt = q_txt & q$ + PO_string(Trim(vtbls!org_v), 30) + q$ + ","
                        q_txt = q_txt & q$ + PO_string(Trim(vtbls!address1_v), 30) + q$ + ","
                        q_txt = q_txt & q$ + PO_string(Trim(vtbls!address2_v), 30) + q$ + ","
                        q_txt = q_txt & q$ + PO_string(Trim(vtbls!city_v), 15) + q$ + ","
                        q_txt = q_txt & q$ + PO_string(Trim(vtbls!state_v), 2) + q$ + ","
                        q_txt = q_txt & q$ + PO_string(Trim(vtbls!zip_v), 10) + q$ + ","
                        q_txt = q_txt & q$ + PO_string(Trim(vtbls!phone_v), 15) + q$ + ","
                        q_txt = q_txt & q$ + PO_string(Trim(vtbls!fax_v), 15) + q$ + ","
                        q_txt = q_txt & q$ + PO_string(Trim(vtbls!attn_v), 30) + q$ + ","
                        q_txt = q_txt & q$ + PO_string(Trim(vtbls!buyer_v), 15) + q$ + ","
                        vtbls.Close
                    End If
                Else
                    q_txt = q_txt & q$ + PO_string(Trim(GG_TBL!org_v), 30) + q$ + ","
                    q_txt = q_txt & q$ + PO_string(Trim(GG_TBL!address1_v), 30) + q$ + ","
                    q_txt = q_txt & q$ + PO_string(Trim(GG_TBL!address2_v), 30) + q$ + ","
                    q_txt = q_txt & q$ + PO_string(Trim(GG_TBL!city_v), 15) + q$ + ","
                    q_txt = q_txt & q$ + PO_string(Trim(GG_TBL!state_v), 2) + q$ + ","
                    q_txt = q_txt & q$ + PO_string(Trim(GG_TBL!zip_v), 10) + q$ + ","
                    q_txt = q_txt & q$ + PO_string(Trim(GG_TBL!phone_v), 15) + q$ + ","
                    q_txt = q_txt & q$ + PO_string(Trim(GG_TBL!fax_v), 15) + q$ + ","
                    q_txt = q_txt & q$ + PO_string(Trim(GG_TBL!attn_v), 30) + q$ + ","
                    q_txt = q_txt & q$ + PO_string(Trim(GG_TBL!buyer_v), 15) + q$ + ","
                End If
                q_txt = q_txt & q$ + PO_string(" ", 5) + q$ + ","
                q_txt = q_txt & q$ + PO_string(" ", 30) + q$ + ","
                q_txt = q_txt & q$ + PO_string(" ", 30) + q$ + ","
                q_txt = q_txt & q$ + PO_string(" ", 30) + q$ + ","
                If atblc.EOF Then
                    q_txt = q_txt & q$ + PO_string(" ", 15) + q$ + ","
                Else
                    If InStr(s_str, "COST: $") = 0 Then
                        q_txt = q_txt & q$ + PO_string(" ", 15) + q$ + ","
                    Else
                        i% = InStr(s_str, "COST: $") + 7
                        c_val = Val(Mid$(s_str, i%, InStr(i%, s_str, " ") - 1))
                        q_txt = q_txt + Format$(c_val, "000000000.00000") + ","
                    End If
                End If
                If atblc.EOF Then
                    q_txt = q_txt & q$ + "  " + q$ + ","
                    q_txt = q_txt & q$ + PO_string(" ", 10) + q$ + ","
                    q_txt = q_txt & q$ + PO_string(" ", 15) + q$ + ","
                    q_txt = q_txt & q$ + PO_string(" ", 15) + q$ + ","
                    q_txt = q_txt & q$ + PO_string(" ", 30) + q$ + ","
                Else
                    If InStr(s_str, "FREIGHT TYPE:") = 0 Then
                        q_txt = q_txt & q$ + "  " + q$ + ","
                    Else
                        i% = InStr(s_str, "FREIGHT TYPE:") + 13
                        t_str = Space$(2)
                        LSet t_str = Mid$(s_str, i%, InStr(i%, s_str, " ") - 1)
                        q_txt = q_txt & q$ + t_str + q$ + ","
                    End If
                    q_txt = q_txt & q$ + PO_string(" ", 10) + q$ + ","
                    q_txt = q_txt & q$ + PO_string(" ", 15) + q$ + ","
                    If InStr(atblc!descr, "FREIGHT AMOUNT: $") = 0 Then
                        q_txt = q_txt & q$ + PO_string(" ", 15) + q$ + ","
                    Else
                        i% = InStr(s_str, "FREIGHT AMOUNT: $") + 17
                        c_val = Val(Mid$(s_str, i%, InStr(i%, s_str, " ") - 1))
                        q_txt = q_txt + Format$(c_val, "000000000.00000") + ","
                    End If
                    If InStr(s_str, "EFFECTIVE DATE:") = 0 Then
                        q_txt = q_txt & q$ + PO_string(" ", 30) + q$ + ","
                    Else
                        i% = InStr(s_str, "EFFECTIVE DATE:") + 15
                        t_str = Space$(30)
                        LSet t_str = Mid$(s_str, i%, 10)    'InStr(i% + 1, s_str, " ") - 1)
                        q_txt = q_txt & q$ + t_str + q$ + ","
                    End If
                End If
                q_txt = q_txt & q$ + PO_string(" ", 15) + q$ + ","
                If InStr(G_combo1, "PROMISE DATE") > 0 Or Trim(G_pdate) <> "" Then
                    q_txt = q_txt & q$ + G_pdate + q$ + ","
                Else
                    'date_var = Format(Now, "mm/dd/yy")
                    date_var = Format(Now, rdate)
                    plus_days = Int(Llt + dn!ord_freq_im - ord_fr + 0.5)
                    'res_date = Format(Date + plus_days, "mm/dd/yy")
                    res_date = Format(Date + plus_days, rdate)
                    end_line = res_date
                    q_txt = q_txt & q$ + end_line + q$ + ","
                End If
                '++++++++++++++
                If Label4 = "True" Then
                    q_txt = q_txt & q$ + "02" + q$ + ","
                Else
                    q_txt = q_txt & q$ + Left$(G_combo1, 2) + q$ + ","
                End If
                Set oTbl = DB.OpenRecordset("select * from odesc where " & _
                                            "item = '" & dn!item_num_im & "' and " & _
                                            "whse = '" & dn!whse_im & "'")
                If Not oTbl.EOF Then
                    t_str = Space$(20)
                    LSet t_str = oTbl!descr
                    q_txt = q_txt & q$ + t_str + q$ + ","
                Else
                    q_txt = q_txt & q$ + PO_string(G_pormemotxt, 20) + q$ + ","
                End If
                oTbl.Close
                q_txt = q_txt + q$ + PO_string(Trim(dn!s_vendor_im), 9) + q$ + ","
                If IsNull(dn!brand_im) Then
                    q_txt = q_txt + q$ + Space$(4) + q$ + ","
                Else
                    q_txt = q_txt + q$ + PO_string(Trim(dn!brand_im), 4) + q$ + ","
                End If
                If IsNull(dn!brand_2_im) Then
                    q_txt = q_txt + q$ + Space$(4) + q$ + ","
                Else
                    q_txt = q_txt + q$ + PO_string(Trim(dn!brand_2_im), 4) + q$ + ","
                End If
                q_txt = q_txt + Format$(dn!price_brk_im, "0000000.000") + ","
            Else
free_at_last:
                'If Trim(dn!item_num_im) = "AX002" Then
                '    Debug.Print
                'End If
                q_txt = ""
                q_txt = q_txt + Format$(G_PON, "0000000")
                q_txt = q_txt + Format$(rec_counter, "0000")
                q_txt = q_txt + PO_string(Trim(dn!item_num_im), 28)
                If G_Vspace Then
                    If Trim(G_VendCons) <> "" And Left(G_combo1, 2) = "77" Then
                        tvend = Replace(G_VendCons, "^", " ")
                        q_txt = q_txt + PO_string(Trim(tvend), 9)
                        'q_txt = q_txt + PO_string(Trim(G_VendCons), 9)
                    Else
                        tvend = Replace(dn!vendor_im, "^", " ")
                        q_txt = q_txt + PO_string(Trim(tvend), 9)
                        'q_txt = q_txt + PO_string(Trim(dn!vendor_im), 9)
                    End If
                Else
                    If Trim(G_VendCons) <> "" And Left(G_combo1, 2) = "77" Then
                        q_txt = q_txt + PO_string(Trim(G_VendCons), 9)
                    Else
                        q_txt = q_txt + PO_string(Trim(dn!vendor_im), 9)
                    End If
                End If
                q_txt = q_txt & PO_string(Trim(dn!whse_im), 5)
                q_txt = q_txt + PO_string(Trim(dn!item_desc_im), 30)
                q_txt = q_txt + PO_string(Trim(dn!unit_meas_im), 4)
                'add xdesc logic here
                Set atblc = DB.OpenRecordset("select * from xdesc where " & _
                                             "item = '" & dn!item_num_im & "' and " & _
                                             "whse = '" & dn!whse_im & "'")
                If atblc.EOF Then
                If G_Free And FGamt > 0 Then
                        q_txt = q_txt + Format$(0, "000000.00000")
                    Else
                        q_txt = q_txt + Format$(dn!unit_cost_im, "000000.00000")
                    End If
                    'q_txt = q_txt + Format$(dn!unit_cost_im, "000000.00000")
                Else
                    If G_Free And FGamt > 0 Then
                        q_txt = q_txt + Format$(0, "000000.00000")
                    Else
                        s_str = atblc!descr
                        For i% = 1 To Len(s_str)
                            If Mid$(s_str, i%, 1) = Chr$(10) Or Mid$(s_str, i%, 1) = Chr$(13) Then
                                Mid$(s_str, i%, 1) = " "
                            End If
                        Next
                        If InStr(s_str, "ACT COST: $") = 0 Then
                            q_txt = q_txt + Format$(dn!unit_cost_im, "000000.00000")
                        Else
                            i% = InStr(s_str, "ACT COST: $") + 11
                            If InStr(i%, Trim(s_str), " ") > 0 Or Right$(Trim(s_str), 1) <> "$" Then
                                If InStr(i%, Trim(s_str), " ") = 0 Then
                                    If InStr(Mid$(s_str, i%, 15), "/") = 0 Then
                                        c_val = Val(Mid$(s_str, i%))
                                    End If
                                Else
                                    If InStr(Mid$(s_str, i%, 15), "/") = 0 Then
                                        c_val = Val(Mid$(s_str, i%, InStr(i%, Trim(s_str), " ") - 1))
                                    End If
                                End If
                                If InStr(Mid$(s_str, i%, 15), "/") = 0 Then
                                    q_txt = q_txt + Format$(c_val, "000000.00000")
                                Else
                                    i% = InStr(s_str, "C/W LB PRICE")
                                    If i% = 0 Then
                                        q_txt = q_txt + Format$(dn!unit_cost_im, "000000.00000")
                                    Else
                                        i% = InStr(s_str, "C/W LB PRICE") + 12
                                        c_val = Val(Mid$(s_str, i%, 8))
                                        q_txt = q_txt + Format$(c_val, "000000.00000")
                                    End If
                                End If
                            Else
                                q_txt = q_txt + Format$(dn!unit_cost_im, "000000.00000")
                            End If
                        End If
                    End If
                End If
                Set oTbl = DB.OpenRecordset("select * from odesc where " & _
                                            "item = '" & dn!item_num_im & "' and " & _
                                            "whse = '" & dn!whse_im & "'")
                If Not oTbl.EOF Then
                    'If Left$(Trim(oTbl!descr), 2) = "FG" And G_Free And FGamt > 0 Then
                    If Left$(Trim(oTbl!descr), 2) = "FG" And G_Free Then
                        If Left$(Trim(oTbl!descr), 2) = "FG" Then
                            FGood = True
                            FGamt = Val(Mid$(Trim(oTbl!descr), 3))
                            q_txt = q_txt & Format$(dn!ord_qty_im - FGamt, "0000000")
                            oTbl.Edit
                            oTbl!descr = "Free Good Line " & Str$(FGamt)
                            oTbl.Update
                        ElseIf Left$(Trim(oTbl!descr), 14) = "Free Good Line" Then
                            q_txt = q_txt & Format$(FGamt, "0000000")
                            FGood = False
                        Else
                            q_txt = q_txt & Format$(dn!ord_qty_im, "0000000")
                        End If
                    ElseIf Left$(Trim(oTbl!descr), 14) = "Free Good Line" Then
                        q_txt = q_txt & Format$(FGamt, "0000000")
                        FGood = False
                    Else
                        q_txt = q_txt & Format$(dn!ord_qty_im, "0000000")
                    End If
                Else
                    q_txt = q_txt & Format$(dn!ord_qty_im, "0000000")
                End If
                q_txt = q_txt + dn!abc_class_im
                q_txt = q_txt + Format$(dn!wght_im, "0000000.000")
                q_txt = q_txt + Format$(dn!att_flg_im, "@")
                If Trim(G_VendCons) <> "" And Left(G_combo1, 2) = "77" Then
                    If G_Local Then
                        vTbl.Seek "=", Trim(G_VendCons), dn!whse_im
                        If vTbl.NoMatch Then
                            Llt = 0
                            ord_fr = 0
                        Else
                            Select Case dn!abc_class_im
                                Case "A"
                                    Llt = vTbl!a_lead_time_v
                                    ord_fr = vTbl!a_ord_fr_v
                                Case "B"
                                    Llt = vTbl!b_lead_time_v
                                    ord_fr = vTbl!b_ord_fr_v
                                Case "C"
                                    Llt = vTbl!c_lead_time_v
                                    ord_fr = vTbl!c_ord_fr_v
                                Case Else
                                    Llt = vTbl!a_lead_time_v
                                    ord_fr = vTbl!a_ord_fr_v
                                    'Llt = 0
                                    'ord_fr = 0
                            End Select
                        End If
                    Else
                        Set vtbls = DB.OpenRecordset("select * from ovendor where " & _
                                                     "vendor_v='" & Trim(G_VendCons) & "' and " & _
                                                     "whse_v='" & Trim(dn("whse_im")) & "'")
                        If vtbls.EOF Then
                            Llt = 0
                            ord_fr = 0
                        Else
                            Select Case dn!abc_class_im
                                Case "A"
                                    Llt = vtbls!a_lead_time_v
                                    ord_fr = vtbls!a_ord_fr_v
                                Case "B"
                                    Llt = vtbls!b_lead_time_v
                                    ord_fr = vtbls!b_ord_fr_v
                                Case "C"
                                    Llt = vtbls!c_lead_time_v
                                    ord_fr = vtbls!c_ord_fr_v
                                Case Else
                                    Llt = vtbls!a_lead_time_v
                                    ord_fr = vtbls!a_ord_fr_v
                            End Select
                        End If
                    End If
                Else
                    GG_TBL.Seek "=", dn!vendor_im, dn!whse_im
                    If GG_TBL.NoMatch Then
                        Set vtbls = DB.OpenRecordset("select * from ovendor where " & _
                                                     "vendor_v='" & dn!vendor_im & "' and " & _
                                                     "whse_v='" & dn!whse_im & "'")
                        If Not vtbls.EOF Then
                            Select Case dn!abc_class_im
                                Case "A"
                                    Llt = vtbls!a_lead_time_v
                                    ord_fr = vtbls!a_ord_fr_v
                                Case "B"
                                    Llt = vtbls!b_lead_time_v
                                    ord_fr = vtbls!b_ord_fr_v
                                Case "C"
                                    Llt = vtbls!c_lead_time_v
                                    ord_fr = vtbls!c_ord_fr_v
                                Case Else
                                    Llt = vtbls!a_lead_time_v
                                    ord_fr = vtbls!a_ord_fr_v
                                    'Llt = 0
                                    'ord_fr = 0
                            End Select
                        End If
                        'Else
                        '    Llt = 0
                        '    ord_fr = 0
                        'End If
                    Else
                        Select Case dn!abc_class_im
                            Case "A"
                                Llt = GG_TBL!a_lead_time_v
                                ord_fr = GG_TBL!a_ord_fr_v
                            Case "B"
                                Llt = GG_TBL!b_lead_time_v
                                ord_fr = GG_TBL!b_ord_fr_v
                            Case "C"
                                Llt = GG_TBL!c_lead_time_v
                                ord_fr = GG_TBL!c_ord_fr_v
                            Case Else
                                Llt = GG_TBL!a_lead_time_v
                                ord_fr = GG_TBL!a_ord_fr_v
                                'Llt = 0
                                'ord_fr = 0
                        End Select
                    End If
                End If
                If dn!abc_class_im = "X" Then
                    q_txt = q_txt & PO_string(Llt, 6)
                Else
                    q_txt = q_txt & PO_string(dn!lead_time_im, 6)
                End If
                'q_txt = q_txt & PO_string(Llt, 6)
                If Trim(G_VendCons) <> "" And Left(G_combo1, 2) = "77" Then
                    If G_Local Then
                        q_txt = q_txt & PO_string(Trim(vTbl!org_v), 30)
                        q_txt = q_txt & PO_string(Trim(vTbl!address1_v), 30)
                        q_txt = q_txt & PO_string(Trim(vTbl!address2_v), 30)
                        q_txt = q_txt & PO_string(Trim(vTbl!city_v), 15)
                        q_txt = q_txt & PO_string(Trim(vTbl!state_v), 2)
                        q_txt = q_txt & PO_string(Trim(vTbl!zip_v), 10)
                        q_txt = q_txt & PO_string(Trim(vTbl!phone_v), 15)
                        q_txt = q_txt & PO_string(Trim(vTbl!fax_v), 15)
                        q_txt = q_txt & PO_string(Trim(vTbl!attn_v), 30)
                        q_txt = q_txt & PO_string(Trim(vTbl!buyer_v), 15)
                    Else
                        q_txt = q_txt & PO_string(Trim(vtbls!org_v), 30)
                        q_txt = q_txt & PO_string(Trim(vtbls!address1_v), 30)
                        q_txt = q_txt & PO_string(Trim(vtbls!address2_v), 30)
                        q_txt = q_txt & PO_string(Trim(vtbls!city_v), 15)
                        q_txt = q_txt & PO_string(Trim(vtbls!state_v), 2)
                        q_txt = q_txt & PO_string(Trim(vtbls!zip_v), 10)
                        q_txt = q_txt & PO_string(Trim(vtbls!phone_v), 15)
                        q_txt = q_txt & PO_string(Trim(vtbls!fax_v), 15)
                        q_txt = q_txt & PO_string(Trim(vtbls!attn_v), 30)
                        q_txt = q_txt & PO_string(Trim(vtbls!buyer_v), 15)
                        vtbls.Close
                    End If
                Else
                    If Not GG_TBL.NoMatch Then
                        q_txt = q_txt & PO_string(Trim(GG_TBL!org_v), 30)
                        q_txt = q_txt & PO_string(Trim(GG_TBL!address1_v), 30)
                        q_txt = q_txt & PO_string(Trim(GG_TBL!address2_v), 30)
                        q_txt = q_txt & PO_string(Trim(GG_TBL!city_v), 15)
                        q_txt = q_txt & PO_string(Trim(GG_TBL!state_v), 2)
                        q_txt = q_txt & PO_string(Trim(GG_TBL!zip_v), 10)
                        q_txt = q_txt & PO_string(Trim(GG_TBL!phone_v), 15)
                        q_txt = q_txt & PO_string(Trim(GG_TBL!fax_v), 15)
                        q_txt = q_txt & PO_string(Trim(GG_TBL!attn_v), 30)
                        q_txt = q_txt & PO_string(Trim(GG_TBL!buyer_v), 15)
                    Else
                        q_txt = q_txt & PO_string(Trim(vtbls!org_v), 30)
                        q_txt = q_txt & PO_string(Trim(vtbls!address1_v), 30)
                        q_txt = q_txt & PO_string(Trim(vtbls!address2_v), 30)
                        q_txt = q_txt & PO_string(Trim(vtbls!city_v), 15)
                        q_txt = q_txt & PO_string(Trim(vtbls!state_v), 2)
                        q_txt = q_txt & PO_string(Trim(vtbls!zip_v), 10)
                        q_txt = q_txt & PO_string(Trim(vtbls!phone_v), 15)
                        q_txt = q_txt & PO_string(Trim(vtbls!fax_v), 15)
                        q_txt = q_txt & PO_string(Trim(vtbls!attn_v), 30)
                        q_txt = q_txt & PO_string(Trim(vtbls!buyer_v), 15)
                    End If
                End If
                q_txt = q_txt & PO_string(" ", 5)
                Set oTbl = DB.OpenRecordset("select * from odesc where " & _
                                            "item = '" & dn!item_num_im & "' and " & _
                                            "whse = '" & dn!whse_im & "'")
                If Not oTbl.EOF Then
'                    If (Left$(Trim(oTbl!descr), 2) = "FG" Or Left$(Trim(oTbl!descr), 14) = "Free Good Line") And G_Free Then
'                        t_str = Space$(30)
'                    Else
                        t_str = Space$(30)
                        LSet t_str = oTbl!descr
                    'End If
                    q_txt = q_txt & t_str
                Else
                    q_txt = q_txt & PO_string(G_pormemotxt, 30)
                End If
                oTbl.Close
                'q_txt = q_txt & PO_string(" ", 30)
                q_txt = q_txt & PO_string(" ", 30)
                t_str = Space$(30)
                LSet t_str = Trim(Mid$(G_combo1, InStr(G_combo1, "-") + 1))
                q_txt = q_txt & t_str
                'q_txt = q_txt & PO_string(" ", 30)
                If atblc.EOF Then
                    q_txt = q_txt & PO_string(" ", 15)
                Else
                    If InStr(s_str, "COST: $") = 0 Then
                        q_txt = q_txt & PO_string(" ", 15)
                    Else
                        i% = InStr(s_str, "COST: $") + 7
                        c_val = Val(Mid$(s_str, i%, InStr(i%, s_str, " ") - 1))
                        q_txt = q_txt + Format$(c_val, "000000000.00000")
                    End If
                End If
                If atblc.EOF Then
                    q_txt = q_txt & "  "
                    q_txt = q_txt & PO_string(" ", 10)
                    q_txt = q_txt & PO_string(" ", 15)
                    q_txt = q_txt & PO_string(" ", 15)
                    q_txt = q_txt & PO_string(" ", 30)
                Else
                    If InStr(s_str, "FREIGHT TYPE:") = 0 Then
                        q_txt = q_txt & "  "
                    Else
                        i% = InStr(s_str, "FREIGHT TYPE:") + 13
                        t_str = Space$(2)
                        LSet t_str = Mid$(s_str, i%, InStr(i%, s_str, " ") - 1)
                        q_txt = q_txt & t_str
                    End If
                    q_txt = q_txt & PO_string(" ", 10)
                    q_txt = q_txt & PO_string(" ", 15)
                    If InStr(atblc!descr, "FREIGHT AMOUNT: $") = 0 Then
                        q_txt = q_txt & PO_string(" ", 15)
                    Else
                        i% = InStr(s_str, "FREIGHT AMOUNT: $") + 17
                        c_val = Val(Mid$(s_str, i%, InStr(i%, s_str, " ") - 1))
                        q_txt = q_txt + Format$(c_val, "000000000.00000")
                    End If
                    If InStr(s_str, "EFFECTIVE DATE:") = 0 Then
                        q_txt = q_txt & PO_string(" ", 30)
                    Else
                        i% = InStr(s_str, "EFFECTIVE DATE:") + 15
                        t_str = Space$(30)
                        LSet t_str = Mid$(s_str, i%, 10)    'InStr(i% + 1, s_str, " ") - 1)
                        q_txt = q_txt & t_str
                    End If
                End If
                q_txt = q_txt & PO_string(" ", 15)
                'If InStr(Combo1.Text, "PROMISE DATE") > 0 Then
'               MsgBox "hmmm..." & Pdate & "^" & In_Perform & "^" & GoalSeeked & "^" & Llt & "^"
                If Len(Trim(G_pdate)) > 0 Then
                    If Len(G_pdate) <> 8 Then
                        'q_txt = q_txt & Format(Pdate, rdate)
                        q_txt = q_txt & Format(G_pdate, "mm/dd/yy")
                    Else
                        q_txt = q_txt & Format(G_pdate, "mm/dd/yy") ' G_pdate
                    End If
                ElseIf In_Perform And Not GoalSeeked Then
                    plus_days = Int(dn!lead_time_im)
                    res_date = Format(Date + plus_days, "mm/dd/yy")
                    'res_date = Format(Date + plus_days, rdate)
                    end_line = res_date
                    q_txt = q_txt & end_line
                Else
                    date_var = Format(Now, "mm/dd/yy")
                    plus_days = Int(Llt)
                    res_date = Format(Date + plus_days, "mm/dd/yy")
                    'res_date = Format(Date + plus_days, rdate)
                    end_line = res_date
                    q_txt = q_txt & end_line
                End If
                If Label4 = "True" Then
                    q_txt = q_txt & "02"
                Else
                    q_txt = q_txt & Left$(G_combo1, 2)
                End If
                q_txt = q_txt & PO_string(G_pormemotxt, 20)
                q_txt = q_txt + PO_string(Trim(dn!s_vendor_im), 9)
                If IsNull(dn!brand_im) Then
                    q_txt = q_txt + Space$(4)
                Else
                    q_txt = q_txt + PO_string(Trim(dn!brand_im), 4)
                End If
                If IsNull(dn!brand_2_im) Then
                    q_txt = q_txt + Space$(4)
                Else
                    q_txt = q_txt + PO_string(Trim(dn!brand_2_im), 4)
                End If
                q_txt = q_txt + Format$(dn!price_brk_im, "0000000.000")
            End If
            Print #FilePO, q_txt
            atblc.Close
            'vtbls.Close
            If FGood And G_Free Then GoTo free_at_last
get_next_first:
            dn.MoveNext
            If dn.EOF Then Exit Do
        Loop
        End If
       End If
    dn.Close
    Close #FilePO

    Dim iTbl As Recordset
    
    Set gTbl = DB.OpenRecordset(query)
    If G_I_VXref Then
        ans% = MsgBox("Use Vendor item numbers on PO's?", vbYesNo)
        If ans% = vbYes Then
            Do While Not gTbl.EOF
                Set iTbl = DB.OpenRecordset("select * from itemxref where " & _
                                            "mitem_num = '" & Trim(gTbl!item_num_im) & "' and " & _
                                            "mitem_loc = '" & Trim(gTbl!whse_im) & "'")
                If Not iTbl.EOF Then
                    gTbl.Edit
                    gTbl!item_num_im = iTbl!vitem_num
                    gTbl!item_desc_im = iTbl!vitem_desc
                    gTbl.Update
                End If
                iTbl.Close
                gTbl.MoveNext
                DoEvents
            Loop
            gTbl.Close
            writeERR = WriteINI("ITEMXREF", "whichway", "V", inifilename$)
        End If
    End If
    
    Dim TblPO As Recordset
    Set TblPO = DB.OpenRecordset("por")
    Set dn = DB.OpenRecordset(query)
    Dim ctbl As Recordset
    cnt = 0
    'Set sTbl = DB.OpenRecordset("sorders")
    On Error Resume Next
    On Error GoTo gary_error
    If Not dn.EOF Then
            
        sav_vendor_im = dn!vendor_im
        sav_whse = dn!whse_im
        Do While Not dn.EOF
            On Error GoTo gary_error
            If G_Buy_Perform Then
                If dn!ord_qty_im = 0 And Trim(dn!r_monf_im) = "M" Then
                    GoSub buy_perform_logic
                    GoTo get_next
                End If
                If dn!ord_qty_im = 0 And Trim(dn!r_monf_im) <> "M" Then
                    GoTo get_next
                End If
            End If
            If sav_vendor_im <> dn!vendor_im Then
                rec_counter = 0
                sav_vendor_im = dn!vendor_im
            End If
            cnt = cnt + 1
            TblPO.AddNew

            MainFrm.Label1.Caption = "M-Working on: " & dn!item_num_im
            DoEvents
            
            TblPO!por = Format$(G_PON, "00000000")
'            '------------------------------------------------------------------
'            ' added logic for saved orders table
'            If G_Sorders Then
'                sTbl.AddNew
'                For i% = 0 To dn.Fields.Count - 1
'                    sTbl(i%) = dn(i%)
'                Next
'                sTbl("ponum_im") = Format$(G_PON, "00000000")
'                sTbl.Update
'            End If
'            '------------------------------------------------------------------
'
            If Label4 = "True" Then
                TblPO!blanked = "02 - BLANKET"
            Else
                TblPO!blanked = G_combo1
            End If
            TblPO!poline = Format$(cnt, "0000")
            If Trim(G_VendCons) <> "" And Left(G_combo1, 2) = "77" Then
                TblPO!vnd = S_Check(Trim(G_VendCons))
            Else
                TblPO!vnd = S_Check(Trim(dn!vendor_im))
            End If
            TblPO("itm") = S_Check(Trim(dn!item_num_im))
            TblPO("idesc") = S_Check(Trim(dn!item_desc_im))
            TblPO!um = S_Check(Trim(dn!unit_meas_im))
            TblPO("abc") = S_Check(dn!abc_class_im)
            TblPO("cost") = Format(dn!unit_cost_im, "#####0.0000")
            TblPO("weight") = Format(dn!wght_im, "0000000.000")
            TblPO("attn") = Format$(dn!att_flg_im, "@")
            If Trim(G_VendCons) <> "" And Left(G_combo1, 2) = "77" Then
                If G_Local Then
                    vTbl.Seek "=", G_VendCons, dn!whse_im
                    If vTbl.NoMatch Then
                        Set vtbls = DB.OpenRecordset("select * from ovendor where " & _
                                                     "vendor_v='" & dn!vendor_im & "' and " & _
                                                     "whse_v='" & dn!whse_im & "'")
                        If Not vtbls.EOF Then
                            Select Case dn!abc_class_im
                                Case "A"
                                    Llt = vtbls!a_lead_time_v
                                    ord_fr = vtbls!a_ord_fr_v
                                Case "B"
                                    Llt = vtbls!b_lead_time_v
                                    ord_fr = vtbls!b_ord_fr_v
                                Case "C"
                                    Llt = vtbls!c_lead_time_v
                                    ord_fr = vtbls!c_ord_fr_v
                                Case Else
                                    Llt = vtbls!a_lead_time_v
                                    ord_fr = vtbls!a_ord_fr_v
                                    'Llt = 0
                                    'ord_fr = 0
                            End Select
                        Else
                            Llt = 0
                            ord_fr = 0
                        End If
                        'Llt = 0
                        'ord_fr = 0
                    Else
                        Select Case dn!abc_class_im
                            Case "A"
                                Llt = vTbl!a_lead_time_v
                                ord_fr = vTbl!a_ord_fr_v
                            Case "B"
                                Llt = vTbl!b_lead_time_v
                                ord_fr = vTbl!b_ord_fr_v
                            Case "C"
                                Llt = vTbl!c_lead_time_v
                                ord_fr = vTbl!c_ord_fr_v
                            Case Else
                                Llt = vTbl!a_lead_time_v
                                ord_fr = vTbl!a_ord_fr_v
                                'Llt = 0
                                'ord_fr = 0
                        End Select
                    End If
                Else
                    Set vtbls = DB.OpenRecordset("select * from ovendor where " & _
                                                 "vendor_v='" & Trim(G_VendCons) & "' and " & _
                                                 "whse_v='" & Trim(dn("whse_im")) & "'")
                    If vtbls.EOF Then
                        Llt = 0
                        ord_fr = 0
                    Else
                        Select Case dn!abc_class_im
                            Case "A"
                                Llt = vtbls!a_lead_time_v
                                ord_fr = vtbls!a_ord_fr_v
                            Case "B"
                                Llt = vtbls!b_lead_time_v
                                ord_fr = vtbls!b_ord_fr_v
                            Case "C"
                                Llt = vtbls!c_lead_time_v
                                ord_fr = vtbls!c_ord_fr_v
                            Case Else
                                Llt = vtbls!a_lead_time_v
                                ord_fr = vtbls!a_ord_fr_v
                        End Select
                    End If
                End If
            Else
                GG_TBL.Seek "=", dn!vendor_im, dn!whse_im
                If GG_TBL.NoMatch Then
                    Set vtbls = DB.OpenRecordset("select * from ovendor where " & _
                                                 "vendor_v='" & dn!vendor_im & "' and " & _
                                                 "whse_v='" & dn!whse_im & "'")
                    If Not vtbls.EOF Then
                        Select Case dn!abc_class_im
                            Case "A"
                                Llt = vtbls!a_lead_time_v
                                ord_fr = vtbls!a_ord_fr_v
                            Case "B"
                                Llt = vtbls!b_lead_time_v
                                ord_fr = vtbls!b_ord_fr_v
                            Case "C"
                                Llt = vtbls!c_lead_time_v
                                ord_fr = vtbls!c_ord_fr_v
                            Case Else
                                Llt = vtbls!a_lead_time_v
                                ord_fr = vtbls!a_ord_fr_v
                                'Llt = 0
                                'ord_fr = 0
                        End Select
                    Else
                        Llt = 0
                        ord_fr = 0
                    End If
                'Else
                '    Llt = 0
                '    ord_fr = 0
                Else
                    Select Case dn!abc_class_im
                        Case "A"
                            Llt = GG_TBL!a_lead_time_v
                            ord_fr = GG_TBL!a_ord_fr_v
                        Case "B"
                            Llt = GG_TBL!b_lead_time_v
                            ord_fr = GG_TBL!b_ord_fr_v
                        Case "C"
                            Llt = GG_TBL!c_lead_time_v
                            ord_fr = GG_TBL!c_ord_fr_v
                        Case Else
                            Llt = GG_TBL!a_lead_time_v
                            ord_fr = GG_TBL!a_ord_fr_v
                            'Llt = 0
                            'ord_fr = 0
                    End Select
                End If
            End If
            TblPO("lead") = Llt
            TblPO("freq") = ord_fr
            TblPO!BFreq = dn!ord_freq_im
            TblPO!qorder = Format(dn!ord_qty_im, "######0")
            TblPO!whse_f = Trim(dn!whse_im)
            'On Error Resume Next
            If Trim(G_VendCons) <> "" And Left(G_combo1, 2) = "77" Then
                If G_Local Then
                    TblPO!name_f = S_Check(Trim(vTbl!org_v))
                    TblPO!address1_f = S_Check(Trim(vTbl!address1_v))
                    TblPO!address2_f = S_Check(Trim(vTbl!address2_v))
                    TblPO!city_f = S_Check(Trim(vTbl!city_v))
                    TblPO!state_f = S_Check(Trim(vTbl!state_v))
                    TblPO!zip_f = S_Check(Trim(vTbl!zip_v))
                    TblPO!phone_f = S_Check(Trim(vTbl!phone_v))
                    TblPO!fax_f = S_Check(Trim(vTbl!fax_v))
                    TblPO!attn_f = S_Check(Trim(vTbl!attn_v))
                    TblPO!buyer_f = S_Check(Trim(vTbl!buyer_v))
                Else
                    TblPO!name_f = S_Check(Trim(vtbls!org_v))
                    TblPO!address1_f = S_Check(Trim(vtbls!address1_v))
                    TblPO!address2_f = S_Check(Trim(vtbls!address2_v))
                    TblPO!city_f = S_Check(Trim(vtbls!city_v))
                    TblPO!state_f = S_Check(Trim(vtbls!state_v))
                    TblPO!zip_f = S_Check(Trim(vtbls!zip_v))
                    TblPO!phone_f = S_Check(Trim(vtbls!phone_v))
                    TblPO!fax_f = S_Check(Trim(vtbls!fax_v))
                    TblPO!attn_f = S_Check(Trim(vtbls!attn_v))
                    TblPO!buyer_f = S_Check(Trim(vtbls!buyer_v))
                End If
            Else
                On Error Resume Next
                If GG_TBL.NoMatch Then
                    TblPO!name_f = S_Check(Trim(vtbls!org_v))
                    TblPO!address1_f = S_Check(Trim(vtbls!address1_v))
                    TblPO!address2_f = S_Check(Trim(vtbls!address2_v))
                    TblPO!city_f = S_Check(Trim(vtbls!city_v))
                    TblPO!state_f = S_Check(Trim(vtbls!state_v))
                    TblPO!zip_f = S_Check(Trim(vtbls!zip_v))
                    TblPO!phone_f = S_Check(Trim(vtbls!phone_v))
                    TblPO!fax_f = S_Check(Trim(vtbls!fax_v))
                    TblPO!attn_f = S_Check(Trim(vtbls!attn_v))
                    TblPO!buyer_f = S_Check(Trim(vtbls!buyer_v))
                Else
                    TblPO!name_f = S_Check(Trim(GG_TBL!org_v))
                    TblPO!address1_f = S_Check(Trim(GG_TBL!address1_v))
                    TblPO!address2_f = S_Check(Trim(GG_TBL!address2_v))
                    TblPO!city_f = S_Check(Trim(GG_TBL!city_v))
                    TblPO!state_f = S_Check(Trim(GG_TBL!state_v))
                    TblPO!zip_f = S_Check(Trim(GG_TBL!zip_v))
                    TblPO!phone_f = S_Check(Trim(GG_TBL!phone_v))
                    TblPO!fax_f = S_Check(Trim(GG_TBL!fax_v))
                    TblPO!attn_f = S_Check(Trim(GG_TBL!attn_v))
                    TblPO!buyer_f = S_Check(Trim(GG_TBL!buyer_v))
                End If
                On Error GoTo gary_error
            End If
            
            TblPO!whse_t = Space(2)
            TblPO!name_t = Space(30)
            TblPO!address1_t = Space(30)
            TblPO!address2_t = Space(30)
            If Not G_Local Then
                Set atblc = DB.OpenRecordset("select * from txdesc where " & _
                                             "item = '" & dn!item_num_im & "' and " & _
                                             "whse = '" & dn!whse_im & "'")
            Else
                Set atblc = DB.OpenRecordset("select * from xdesc where " & _
                                             "item = '" & dn!item_num_im & "' and " & _
                                             "whse = '" & dn!whse_im & "'")
            End If
            If atblc.EOF Then
                TblPO!city_t = PO_string(" ", 15)
            Else
                If InStr(atblc!descr, "COST: $") = 0 Then
                    TblPO!city_t = PO_string(" ", 15)
                Else
                    i% = InStr(atblc!descr, "COST: $") + 7
                    c_val = Val(Mid$(atblc!descr, i%, InStr(i%, tblc!descr, " ") - 1))
                    TblPO!city_t = Format$(c_val, "000000000.00000")
                End If
            End If
            atblc.Close
            If InStr(G_combo1, "NDS FREIGHT") > 0 Then
                TblPO!state_t = FType
                TblPO!zip_t = FFactor
                TblPO!phone_t = FDollars
                If Trim(FType) = "W" Then
                    TblPO!fax_t = Format((o_wght / FFactor) * FDollars, "000000000000.00")
                Else
                    TblPO!fax_t = Format((o_qty / FFactor) * FDollars, "000000000000.00")
                End If
            Else
                TblPO!state_t = Space(2)
                TblPO!zip_t = Space(10)
                TblPO!phone_t = Space(15)
                'TblPO!fax_t = dn!hl_im
                TblPO!fax_t = Space(15)
                If G_Cust Then
                    TblPO!address1_t = Format(dn!qty_on_ord_im, "#######.00")
                    TblPO!address2_t = Format(dn!r_cper_im, "#######.00")
                    TblPO!phone_t = Format(dn!qty_on_hand_im, "#######.00")
                    TblPO!fax_t = dn!hl_im
                    '---------------------------- added fields 11/09/05
                    TblPO!zip_t = Format(dn!qty_back_im, "#######.00")
                    TblPO!attn_t = Format(dn!qty_comm_im, "#######.00")
                    TblPO!buyer_t = Format(dn!qty_avl_im, "#######.00")
                    '----------------------------- end of add
                Else
                    TblPO!attn_t = Space(30)
                    TblPO!buyer_t = Space(15)
                End If
            End If
            If G_Cust Then
                TblPO!address1_t = Format(dn!qty_on_ord_im, "#######.00")
                TblPO!address2_t = Format(dn!r_cper_im, "#######.00")
                TblPO!phone_t = Format(dn!qty_on_hand_im, "#######.00")
                TblPO!fax_t = dn!hl_im
                '---------------------------- added fields 11/09/05
                TblPO!zip_t = Format(dn!qty_back_im, "#######.00")
                TblPO!attn_t = Format(dn!qty_comm_im, "#######.00")
                TblPO!buyer_t = Format(dn!net_qty_avl_im, "#######.00")
                '----------------------------- end of add
            End If

            Print #filnum, "G_pdate=", G_pdate
            Print #filnum, "rdate="; rdate; "^"
            If Len(Trim(G_pdate)) > 0 Then
                If Len(Trim(rdate)) = 8 Then
                    Print #filnum, "tbl!dop - 1 ="; Format(G_pdate, Left(rdate, 4) & "yy")
                    TblPO!dop = Format(G_pdate, Left(rdate, 4) & "yy")
                Else
                    Print #filnum, "tbl!dop - 2 ="; Format(CVDate(G_pdate), Left(rdate, 4) & "yy")
                    TblPO!dop = Format(CVDate(G_pdate), Left(rdate, 4) & "yy")
                End If    'cdate
            ElseIf In_Perform And Not GoalSeeked Then
                plus_days = Int(dn!lead_time_im)
                'res_date = Format(Date + plus_days, "mm/dd/yy")
                If Len(Trim(rdate)) = 8 Then
                    Print #filnum, "tbl!dop - 3 ="; Format(Date + plus_days, Left(rdate, 4) & "yy")
                    res_date = Format(Date + plus_days, Left(rdate, 4) & "yy")
                Else
                    Print #filnum, "tbl!dop - 4 ="; Format(Date + plus_days, Left(rdate, 6) & "yy")
                    res_date = Format(Date + plus_days, Left(rdate, 6) & "yy")
                End If
                'res_date = Format(Date + plus_days, rdate)
                end_line = res_date
                TblPO!dop = end_line
            Else
                date_var = Format(Now, "mm/dd/yy")
                'plus_days = Int(Llt + dn!ord_freq_im - ord_fr + 0.5)
                plus_days = Int(Llt)
                'res_date = Format(Date + plus_days, "mm/dd/yy")
                If Len(Trim(rdate)) = 8 Then
                    Print #filnum, "tbl!dop - 5 ="; Format(Date + plus_days, Left(rdate, 4) & "yy")
                    res_date = Format(Date + plus_days, Left(rdate, 4) & "yy")
                Else
                    Print #filnum, "tbl!dop - 6 ="; Format(Date + plus_days, Left(rdate, 6) & "yy")
                    res_date = Format(Date + plus_days, Left(rdate, 6) & "yy")
                End If
                end_line = res_date
                PROMISE_DATE = end_date
                TblPO!dop = end_line
            End If
            'TblPO!doo = Format(Now, "mm/dd/yy")
            If Len(Trim(rdate)) = 8 Then
                Print #filnum, "tbl!dop - 7 ="; Format(Now, Left(rdate, 4) & "yy")
                TblPO!doo = Format(Now, Left(rdate, 4) & "yy")
            Else
                Print #filnum, "tbl!dop - 8 ="; Format(Now, Left(rdate, 6) & "yy")
                TblPO!doo = Format(Now, Left(rdate, 6) & "yy")
            End If
            Set oTbl = DB.OpenRecordset("select * from odesc where " & _
                                        "item = '" & dn!item_num_im & "' and " & _
                                        "whse = '" & dn!whse_im & "'")  ' and " & _
                                        '"vendor = '" & dn!vendor_im & "'")
            'On Error Resume Next
            If Not oTbl.EOF Then
                TblPO!memo1 = Trim(oTbl!descr)
            Else
                TblPO!memo1 = G_pormemotxt
            End If
            oTbl.Close
            'On Error Resume Next
            TblPO("s_vendor") = S_Check(Trim(dn!s_vendor_im))
            If IsNull(dn!brand_im) Then
                TblPO!group1 = Space$(4)
            Else
                TblPO!group1 = Trim(dn!brand_im)
            End If
            If IsNull(TblPO!group1) Or Trim(TblPO!group1) = "" Then
                TblPO!group1 = Space(4)
            End If
            If IsNull(dn!brand_2_im) Then
                TblPO!group2 = Space$(4)
            Else
                TblPO!group2 = Trim(dn!brand_2_im)
            End If
            If IsNull(TblPO!group2) Or Trim(TblPO!group2) = "" Then
                TblPO!group2 = Space(4)
            End If
            TblPO!volume = S_Check(Trim(dn!price_brk_im))
            If G_FOrders Then
                Dim tbdate As Date
                If OQfirst = 0 Then
                    tbdate = Date
                Else
                    tbdate = Date + (dn!ord_freq_im - ord_fr)
                End If
                Set FTbl = DB.OpenRecordset("select * from forders where " & _
                                            "item='" & dn!item_num_im & "' and " & _
                                            "loc='" & dn!whse_im & "' and (fused<>'Y' or " & _
                                            "isnull(fused)) order by fdate")
                TblPO!f_onord = dn!qty_on_hand_im
                TblPO!f_comm = dn!qty_comm_im
                TblPO!f_avail = dn!qty_avl_im
                TblPO!f_wsup = dn!mon_avrg_im
                TblPO!f_avg = dn!fcast_im_12
                lead = dn!lead_time_im
                If Not FTbl.EOF Then
                    Do While Not FTbl.EOF
                        If FTbl!fdate > Format(tbdate, "yyyymmdd") And FTbl!fdate < Format(tbdate + lead, "yyyymmdd") _
                           And (FTbl!fused <> "Y" Or IsNull(FTbl!fused)) Then
                            Exit Do
                        End If
                        FTbl.MoveNext
                    Loop
                    If Not FTbl.EOF Then
                        If FTbl!fdate > Format(tbdate, "yyyymmdd") Then
                            TblPO!f_date = FTbl!fdate
                            TblPO!f_qty = FTbl!fqty
                            TblPO!f_ordnum = FTbl!customer
                        Else
                            TblPO!f_date = ""
                            TblPO!f_qty = 0
                            TblPO!f_ordnum = ""
                        End If
                    Else
                        TblPO!f_date = ""
                        TblPO!f_qty = 0
                        TblPO!f_ordnum = ""
                    End If
                Else
                    TblPO!f_date = ""
                    TblPO!f_qty = 0
                    TblPO!f_ordnum = ""
                End If
            End If
            If G_Caw Then
                TblPO!f_wsup = dn!min_ot_im
            End If
            TblPO.Update
            GoSub buy_perform_logic
get_next:
            dn.MoveNext
            On Error Resume Next
            vtbls.Close
            If dn.EOF Then Exit Do
        Loop
    End If
    On Error Resume Next
    inifilename$ = App.Path + "\WMARS.INI"
    AppName = "PONUM"
    KeyName = "value"
    writeERR = WriteINI(AppName, KeyName, Str$(G_PON), inifilename$)
    dn.Close
    TblPO.Close
    If G_Warehouse <> "" Then
        inifilename$ = App.Path + "\WMARS.INI"
        AppName = "PONUM"
        KeyName = "value"
        writeERR = WriteINI(AppName, KeyName, Str$(G_PON + 1), inifilename$)
        'AppName = "POREQQ"
        'KeyName = "noswitch"
        'writeERR = WriteINI(AppName, KeyName, " ", inifilename$)
        DB.Execute "delete from orders where 1=1"
    End If
    writeERR = WriteINI(vbNull, vbNull, vbNull, inifilename$)
    
    If Not G_TR_ConSwitch Then
        Call Dump_TPor
    End If
    Close #filnum
    Exit Sub
buy_perform_logic:
    Dim g$
    If G_Buy_Perform Then
        If dn!re_ord_nor_im <> dn!ord_qty_im Then
        
            g$ = "select * from buy_perform where " & _
                 "buyer = '" & Trim(G_Buyer) & "' and " & _
                 "bdate = '" & Format(Now, "yyyymmdd") & "' and " & _
                 "item_num = '" & Trim(dn!item_num_im) & "' and " & _
                 "loc = '" & Trim(dn!whse_im) & "'"
            
            Err = 0
            bftbl.Open g$, Cnct
            If Err <> 0 Then
                MsgBox "Buy Perform Error1: " & Str$(Err) & "  " & Err.Description
            End If
            If bftbl.EOF Then
                bftbl.AddNew
            End If
            bftbl!Buyer = G_Buyer
            bftbl!bdate = Format(Date$, "yyyymmdd")
            bftbl!item_num = dn!item_num_im
            bftbl!Loc = dn!whse_im
            bftbl!Vendor = dn!vendor_im
            bftbl!o_qty = dn!re_ord_nor_im
            bftbl!n_qty = dn!ord_qty_im
            bftbl!reason_f = dn!r_monf_im
            Err = 0
            bftbl.Update
            If Err <> 0 Then
                MsgBox "Buy Perform Error2: " & Str$(Err) & "  " & Err.Description
            End If
            bftbl.Close
        End If
    End If
    Return
    
    TblPO.Close
    Return
    
gary_error:
    'MsgBox "Error: " & Str$(Err) & "  " & Err.Description
    Resume Next

End Sub


Public Sub Dump_TPor()
    Dim ttbl As Recordset
    Dim pTbl As Recordset
    Dim sTbl As Recordset
    Dim iiTbl As New adodb.Recordset
    Dim GG_TBL As New adodb.Recordset
    Dim lin As Long
    Dim slocf$
    Dim sloct$
    Dim sitem$
    Dim svend$
    Dim POFname As String
    Dim FilePO As Integer
    Dim preq_path As String
    Dim q$
    q$ = Chr$(34)
    
    iiTbl.CursorType = adOpenKeyset
    iiTbl.LockType = adLockOptimistic
    iiTbl.CursorLocation = adUseClient
    Set ttbl = DB.OpenRecordset("select * from tpor order by whse_f,whse_t")
    If ttbl.EOF Then
        ttbl.Close
        Exit Sub
    End If
    G_PON = G_PON + 1
    If G_PON = 9999999 Then G_PON = 1
    On Error Resume Next
    Err = 0
    inifilename$ = App.Path + "\WMARS.INI"
    AppName = "PONUM"
    KeyName = "value"
    writeERR = WriteINI(AppName, KeyName, Str$(G_PON), inifilename$)
    Dim Dnv As Recordset
    Set Dnv = DB.OpenRecordset("ovendor")
    Dnv.Index = "Index_v_w"
    
    POFname = Format(G_PON, "00000000")
    inifilename$ = App.Path & "\WMARS.INI"
    AppName = "MARSUPDT_PATH"
    KeyName = "poreq_asc"
    preq_path = ReadINI(AppName, KeyName, inifilename$)
    FilePO = FreeFile
    Open (preq_path & "\" & POFname & ".req") For Output As #FilePO
    
    slocf$ = Trim(ttbl!whse_f)
    sloct$ = Trim(ttbl!whse_t)
    sitem$ = Trim(ttbl("itm"))
    svend$ = ttbl("vnd")
    lin = 0
    Set pTbl = DB.OpenRecordset("por")
    Do While Not ttbl.EOF
        If sloct$ <> Trim(ttbl!whse_t) Or slocf$ <> Trim(ttbl!whse_f) Then
            Close FilePO
            G_PON = G_PON + 1
            If G_PON = 9999999 Then G_PON = 1
            On Error Resume Next
            inifilename$ = App.Path + "\WMARS.INI"
            AppName = "PONUM"
            KeyName = "value"
            writeERR = WriteINI(AppName, KeyName, Str$(G_PON), inifilename$)
            POFname = Format(G_PON, "00000000")
            FilePO = FreeFile
            Open (preq_path & "\" & POFname & ".req") For Output As #FilePO
            lin = 0
            sloct$ = Trim(ttbl!whse_t)
            slocf$ = Trim(ttbl!whse_f)
        End If
        lin = lin + 1
        iiTbl.ActiveConnection = cnn
        
        TR = ttbl!qorder
        
        
        If G_PO_format = -1 Then
            Print #FilePO, Format$(G_PON, "0000000");
            Print #FilePO, Format$(lin, "0000");
            Print #FilePO, Tab(12); Trim(ttbl("itm"));
            Print #FilePO, Tab(40); Trim(ttbl!whse_f);
            Print #FilePO, Tab(49); Trim(ttbl!whse_f);
            Print #FilePO, Tab(54); Trim(ttbl("idesc"));
            Print #FilePO, Tab(84); Trim(ttbl!um);
            Print #FilePO, Tab(88); Format$(ttbl("cost"), "000000.00000");
            Print #FilePO, Tab(100); Format$(ttbl!qorder, "0000000")
        Else
            If G_Load_Comma Then
                q_txt = ""
                q_txt = q_txt + q$ + Format$(G_PON, "0000000") + q$ + ","
                q_txt = q_txt + q$ + Format$(lin, "0000") + q$ + ","
                q_txt = q_txt + q$ + PO_string(Trim(ttbl!itm), 28) + q$ + ","
                q_txt = q_txt + q$ + PO_string(Trim(ttbl!vnd), 9) + q$ + ","
                q_txt = q_txt & q$ + PO_string(Trim(ttbl!whse_f), 5) + q$ + ","
                q_txt = q_txt + q$ + PO_string(Trim(ttbl!idesc), 30) + q$ + ","
                q_txt = q_txt + q$ + PO_string(Trim(ttbl!um), 4) + q$ + ","
                q_txt = q_txt + Format$(ttbl!cost, "000000.00000") + ","
                q_txt = q_txt & Format$(ttbl!qorder, "0000000") + ","
                q_txt = q_txt + q$ + ttbl!abc + q$ + ","
                q_txt = q_txt + Format$(ttbl!Weight, "0000000.000") + ","
                q_txt = q_txt + q$ + Format$(ttbl!attn, "@") + q$ + ","
                q_txt = q_txt & Format(ttbl!lead, "####.##") + ","
                q_txt = q_txt & q$ + PO_string(Trim(ttbl("name_f")), 30) + q$ + ","
                q_txt = q_txt & q$ + PO_string(Trim(ttbl("address1_f")), 30) + q$ + ","
                q_txt = q_txt & q$ + PO_string(Trim(ttbl("address2_f")), 30) + q$ + ","
                q_txt = q_txt & q$ + PO_string(Trim(ttbl("city_f")), 15) + q$ + ","
                q_txt = q_txt & q$ + PO_string(Trim(ttbl("state_f")), 2) + q$ + ","
                q_txt = q_txt & q$ + PO_string(Trim(ttbl("zip_f")), 10) + q$ + ","
                q_txt = q_txt & q$ + PO_string(Trim(ttbl("phone_f")), 15) + q$ + ","
                q_txt = q_txt & q$ + PO_string(Trim(ttbl("fax_f")), 15) + q$ + ","
                q_txt = q_txt & q$ + PO_string(Trim(ttbl("attn_f")), 30) + q$ + ","
                q_txt = q_txt & q$ + PO_string(Trim(ttbl("buyer_f")), 15) + q$ + ","
                q_txt = q_txt & q$ + PO_string(Trim(ttbl("whse_t")), 5) + q$ + ","
                q_txt = q_txt & q$ + PO_string(" ", 30) + q$ + ","
                q_txt = q_txt & q$ + PO_string(" ", 30) + q$ + ","
                dm = Space$(30)
                LSet dm = Trim(Mid$(G_combo1, InStr(G_combo1, "-") + 1))
                q_txt = q_txt & q$ + dm + q$ + ","
                'q_txt = q_txt & q$ + PO_string(" ", 30) + q$ + ","
                q_txt = q_txt & q$ + PO_string(" ", 15) + q$ + ","
                q_txt = q_txt & q$ + "  " + q$ + ","
                q_txt = q_txt & q$ + PO_string(" ", 10) + q$ + ","
                q_txt = q_txt & q$ + PO_string(" ", 15) + q$ + ","
                q_txt = q_txt & q$ + PO_string(" ", 15) + q$ + ","
                q_txt = q_txt & q$ + PO_string(" ", 30) + q$ + ","
                q_txt = q_txt & q$ + PO_string(" ", 15) + q$ + ","
                date_var = Format(Now, "mm/dd/yy")
                plus_days = Int(Llt + ttbl!freq - ttbl!freq + 0.5)
                res_date = Format(Date + plus_days, rdate)
                end_line = res_date
                q_txt = q_txt & q$ + end_line + q$ + ","
                '++++++++++++++
                q_txt = q_txt & q$ + ttbl("blanked") + q$ + ","
                q_txt = q_txt & q$ + PO_string("Transfer", 20) + q$ + ","
                q_txt = q_txt + q$ + PO_string(Trim(ttbl!s_vendor), 9) + q$ + ","
                If IsNull(ttbl!group1) Then
                    q_txt = q_txt + q$ + Space$(4) + q$ + ","
                Else
                    q_txt = q_txt + q$ + PO_string(Trim(ttbl!group1), 4) + q$ + ","
                End If
                If IsNull(ttbl!group2) Then
                    q_txt = q_txt + q$ + Space$(4) + q$ + ","
                Else
                    q_txt = q_txt + q$ + PO_string(Trim(ttbl!group2), 4) + q$ + ","
                End If
                q_txt = q_txt + Format$(ttbl!volume, "0000000.000") '+ ","
            Else
                q_txt = ""
                q_txt = q_txt + Format$(G_PON, "0000000")
                q_txt = q_txt + Format$(lin, "0000")
                q_txt = q_txt + PO_string(Trim(ttbl("itm")), 28)
                q_txt = q_txt + PO_string(Trim(ttbl("vnd")), 9)
                q_txt = q_txt & PO_string(Trim(ttbl("whse_f")), 5)
                q_txt = q_txt + PO_string(Trim(ttbl("idesc")), 30)
                q_txt = q_txt + PO_string(Trim(ttbl("um")), 4)
                'add xdesc logic here
                q_txt = q_txt + Format$(ttbl("cost"), "000000.00000")
                q_txt = q_txt & Format$(ttbl("qorder"), "0000000")
                q_txt = q_txt + ttbl("abc")
                q_txt = q_txt + Format$(ttbl("weight"), "0000000.000")
                q_txt = q_txt + Format$(ttbl("attn"), "@")
                q_txt = q_txt & PO_string(ttbl("lead"), 6)
                q_txt = q_txt & PO_string(Trim(ttbl("name_f")), 30)
                q_txt = q_txt & PO_string(Trim(ttbl("address1_f")), 30)
                q_txt = q_txt & PO_string(Trim(ttbl("address2_f")), 30)
                q_txt = q_txt & PO_string(Trim(ttbl("city_f")), 15)
                q_txt = q_txt & PO_string(Trim(ttbl("state_f")), 2)
                q_txt = q_txt & PO_string(Trim(ttbl("zip_f")), 10)
                q_txt = q_txt & PO_string(Trim(ttbl("phone_f")), 15)
                q_txt = q_txt & PO_string(Trim(ttbl("fax_f")), 15)
                q_txt = q_txt & PO_string(Trim(ttbl("attn_f")), 30)
                q_txt = q_txt & PO_string(Trim(ttbl("buyer_f")), 15)
                q_txt = q_txt & PO_string(Trim(ttbl("whse_t")), 5)
                q_txt = q_txt & PO_string(" ", 30)
                q_txt = q_txt & PO_string(" ", 30)
                dm = Space$(30)
                LSet dm = Trim(Mid$(G_combo1, InStr(G_combo1, "-") + 1))
                q_txt = q_txt & dm
                q_txt = q_txt & PO_string(" ", 15)
                q_txt = q_txt & "  "
                q_txt = q_txt & PO_string(" ", 10)
                q_txt = q_txt & PO_string(" ", 15)
                q_txt = q_txt & PO_string(" ", 15)
                q_txt = q_txt & PO_string(" ", 30)
                q_txt = q_txt & PO_string(" ", 15)
                date_var = Format(Now, "mm/dd/yy")
                plus_days = Int(Llt + ttbl("freq") - ttbl("freq") + 0.5)
                res_date = Format(Date + plus_days, "mm/dd/yy")
                end_line = res_date
                q_txt = q_txt & end_line
                q_txt = q_txt & Left$(ttbl("blanked"), 2)
                q_txt = q_txt & PO_string("Transfer", 20)
                ocTbl.Close
                q_txt = q_txt + PO_string(Trim(ttbl("s_vendor")), 9)
                If IsNull(ttbl!group1) Then
                    q_txt = q_txt + Space$(4)
                Else
                    q_txt = q_txt + PO_string(Trim(ttbl("group1")), 4)
                End If
                If IsNull(ttbl!group2) Then
                    q_txt = q_txt + Space$(4)
                Else
                    q_txt = q_txt + PO_string(Trim(ttbl("group2")), 4)
                End If
                q_txt = q_txt + Format$(ttbl("volume"), "0000000.000")
            End If
            Print #FilePO, q_txt
        End If
'        On Error GoTo bad_error
        pTbl.AddNew
        pTbl!por = Format$(G_PON, "00000000")
        pTbl!blanked = ttbl!blanked
        pTbl!poline = Format$(lin, "0000")
        For i% = 3 To 43
            Err = 0: Err.Description = ""   ' recordset
            If ttbl(i%).Type = dbText Then
                If Trim(ttbl(i%)) = "" Then
                    pTbl(i%) = " "
                Else
                    pTbl(i%) = Trim(ttbl(i%))
                End If
            ElseIf ttbl(i%).Type = dbDate Then
                pTbl(i%) = Format(ttbl(i%), "mm/dd/yy")
            Else
                pTbl(i%) = ttbl(i%)
            End If
        Next
        pTbl("vnd") = ttbl!whse_f
        pTbl.Update
        Err = 0: Err.Description = ""
        '-----------< Recalc Items >-------------
        Set Dnv = DB.OpenRecordset("ovendor")
        Dnv.Index = "Index_v_w"
        Dnv.Seek "=", ttbl!vnd, ttbl!whse_t
        pdas = Dnv(2): pdal = Dnv(3): pdao = Dnv(4)
        pdas = PercentPoint_FillRate(pdas)
        pdbs = Dnv(5): pdbl = Dnv(6): pdbo = Dnv(7)
        pdbs = PercentPoint_FillRate(pdbs)
        pdcs = Dnv(8): pdcl = Dnv(9): pdco = Dnv(10)
        pdcs = PercentPoint_FillRate(pdcs)
        Item = Trim(ttbl("itm"))
        WHSE = Trim(ttbl!whse_t)
        Vendor = Trim(Dnv(0))
        Call TrnFrm.RecalcItemTrn("items")
        Dnv.Seek "=", ttbl!vnd, ttbl!whse_f
        pdas = Dnv(2): pdal = Dnv(3): pdao = Dnv(4)
        pdas = PercentPoint_FillRate(pdas)
        pdbs = Dnv(5): pdbl = Dnv(6): pdbo = Dnv(7)
        pdbs = PercentPoint_FillRate(pdbs)
        pdcs = Dnv(8): pdcl = Dnv(9): pdco = Dnv(10)
        pdcs = PercentPoint_FillRate(pdcs)
        Item = Trim(ttbl("itm"))
        WHSE = Trim(ttbl!whse_f)
        Vendor = Trim(Dnv(0))
        Call TrnFrm.RecalcItemTrn("items")
        ttbl.MoveNext
        DoEvents
    Loop
    ttbl.Close
    Close FilePO
    'Call UpdateItem("items")
    Exit Sub
bad_error:
    MsgBox "Error: " & Str$(Err) & "  " & Err.Description
    Resume Next
    
End Sub


Function PO_string(InputString As Variant, OutputSize As Integer) As String
    Dim tstr$
    tstr$ = Space$(OutputSize)
    On Error Resume Next
    If Trim(InputString) <> "" Then
        For i% = 1 To OutputSize
            If Asc(Mid$(InputString, i%, 1)) <> 0 Or Not IsNull(Mid$(InputString, i%, 1)) Then
                Mid$(tstr$, i%, 1) = Mid$(InputString, i%, 1)
            End If
        Next
    End If
    PO_string = tstr$
    Exit Function
    StrSize = Len(Trim(InputString))
    On Error GoTo 0
    If IsNull(InputString) = 0 Then Stop
    RemSize = OutputSize - StrSize
    addstring = ""
    If RemSize <> 0 Then
        For i = 1 To RemSize
            addstring = addstring & " "
        Next i
    End If
    PO_string = InputString & addstring
End Function

Public Sub PO_Order_Mult()
    Dim preq_path As String
    Dim sav_vendor_im As String
    Dim o_qty As Double
    Dim o_wght As Double
    Dim c_val As Double
    Dim t_str As String
    Dim s_str As String
    Dim cnt%
    Dim oqty As Double
    Dim tblc As Recordset
    Dim oTbl As Recordset
    Dim gTbl As Recordset
    Dim mTbl As Recordset
    Dim ocTbl As Recordset
    Dim GG_TBL As Recordset
    Dim bftbl As adodb.Recordset
    
    Dim iTbl As New adodb.Recordset
    Dim tmp_item As String
    Dim tmp_item_desc As String
    Dim dm As String
    Dim tmp_weight As Double
    
    On Error Resume Next
    Set GG_TBL = DB.OpenRecordset("ovendor")
    GG_TBL.Index = "Index_v_w"
    
    Set bftbl = New adodb.Recordset
    bftbl.CursorType = adOpenStatic
    bftbl.LockType = adLockOptimistic
    bftbl.CursorLocation = adUseClient
    
    atblc.CursorLocation = adUseClient
    atblc.CursorType = adOpenForwardOnly
    atblc.ActiveConnection = cnn
    g_tbl.CursorLocation = adUseClient
    g_tbl.CursorType = adOpenForwardOnly
    g_tbl.ActiveConnection = cnn
    On Error Resume Next
skip1:
    Dim TblPO As Recordset
    Set tblc = DB.OpenRecordset("totals")
    If Not tblc.EOF Then
        o_qty = tblc!t_qty
        o_wght = tblc!t_wght
    End If
    tblc.Close
    If G_I_VXref Then
        ans% = MsgBox("Use Vendor item numbers on PO's?", vbYesNo)
    End If
    Screen.MousePointer = 11
    If G_T_Weight > 0 Then
        DB.Execute "DROP TABLE MORDERS"
        DB.Execute "SELECT ORDERS.* INTO MORDERS FROM ORDERS ORDER BY item_num_im, whse_im"
        DB.Execute "update orders set r_rot_im = ord_qty_im * wght_im "
    End If
    If Len(Trim(G_combo1)) = 0 Then G_combo1 = "ZZ - UNKNOWN"
    If Len(G_pormemotxt) = 0 Then G_pormemotxt = Space(1)
    
    POFname = Trim(POFname)
    
    Dim FilePO As Integer
    Dim pos%
    pos% = 0
        If G_Buy_Perform Then
        query = "SELECT * FROM MORDERS  "
    Else
        query = "SELECT * FROM MORDERS WHERE ord_qty_im > 0 "
    End If
    'query = "SELECT * FROM MORDERS WHERE ord_qty_im > 0 "
    query = query & "ORDER BY r_rot_im "
    
    inifilename$ = App.Path + "\WMARS.INI"
    AppName = "MARSUPDT_PATH"
    KeyName = "poreq_asc"
    preq_path = ReadINI(AppName, KeyName, inifilename$)
    FilePO = FreeFile
    Open (preq_path & "\" & POFname & Ext) For Output As #FilePO
    Dim dn As Recordset
    Set dn = DB.OpenRecordset(query)
    If dn.EOF Then
        Close #FilePO
        dn.Close
        If Not G_TR_ConSwitch Then
            Call Dump_TPor
        End If
        Exit Sub
    End If
    dn.MoveFirst
    rec_counter = 0
If G_PO_format = -1 Then                 '-------------- old short format
                                         '-------------- not used here
Else        'MARS-95 format              '-------------- new long format
    If Not dn.EOF Then
        sav_vendor_im = dn!vendor_im
        Do While Not dn.EOF
            If G_Buy_Perform Then
                If dn!ord_qty_im = 0 And r_monf_im = "M" Then
                    GoSub buy_perform_logic
                    GoTo get_next
                ElseIf dn!ord_qty_im = 0 And r_monf_im <> "M" Then
                    GoTo get_next
                End If
            End If
            If sav_vendor_im <> dn!vendor_im Then
                rec_counter = 0
                sav_vendor_im = dn!vendor_im
            End If
            If tmp_weight + (dn!ord_qty_im * dn!wght_im) > G_T_Weight Then
                Close #FilePO
                FilePO = FreeFile
                rec_counter = 0
                tmp_weight = 0
                G_PON = G_PON + 1
                If G_PON = 9999999 Then G_PON = 1
                inifilename$ = App.Path + "\WMARS.INI"
                AppName = "PONUM"
                KeyName = "value"
                writeERR = WriteINI(AppName, KeyName, Str$(G_PON), inifilename$)
                'DB.Execute "UPDATE SYSTEM SET PON_s = '" & G_PON & "' WHERE 1 = 1"
                POFname = Format(G_PON, "00000000")
                Open (preq_path & "\" & POFname & Ext) For Output As #FilePO
            End If
            tmp_weight = tmp_weight + (dn!ord_qty_im * dn!wght_im)
            rec_counter = rec_counter + 1
            q_txt = ""
            q_txt = q_txt + Format$(G_PON, "0000000")
            q_txt = q_txt + Format$(rec_counter, "0000")
            q_txt = q_txt + PO_string(Trim(dn!item_num_im), 28)
            q_txt = q_txt + PO_string(Trim(dn!vendor_im), 9)
            q_txt = q_txt & PO_string(Trim(dn!whse_im), 5)
            q_txt = q_txt + PO_string(Trim(dn!item_desc_im), 30)
            q_txt = q_txt + PO_string(Trim(dn!unit_meas_im), 4)
            'add xdesc logic here
            atblc.Open "select * from xdesc where " & _
                       "item='" & dn!item_num_im & "' and " & _
                       "whse='" & dn!whse_im & "' and " & _
                       "vendor='" & dn!vendor_im & "'"
            If atblc.EOF Then
                q_txt = q_txt + Format$(dn!unit_cost_im, "000000.00000")
            Else
                s_str = atblc!descr
                For i% = 1 To Len(s_str)
                    If Mid$(s_str, i%, 1) = Chr$(10) Or Mid$(s_str, i%, 1) = Chr$(13) Then
                        Mid$(s_str, i%, 1) = " "
                    End If
                Next
                If InStr(s_str, "ACT COST: $") = 0 Then
                    q_txt = q_txt + Format$(dn!unit_cost_im, "000000.00000")
                Else
                    i% = InStr(s_str, "ACT COST: $") + 11
                    If InStr(i%, Trim(s_str), " ") > 0 Or Right$(Trim(s_str), 1) <> "$" Then
                        If InStr(i%, Trim(s_str), " ") = 0 Then
                            If InStr(Mid$(s_str, i%, 15), "/") = 0 Then
                                c_val = Val(Mid$(s_str, i%))
                            End If
                        Else
                            If InStr(Mid$(s_str, i%, 15), "/") = 0 Then
                                c_val = Val(Mid$(s_str, i%, InStr(i%, Trim(s_str), " ") - 1))
                            End If
                        End If
                        If InStr(Mid$(s_str, i%, 15), "/") = 0 Then
                            q_txt = q_txt + Format$(c_val, "000000.00000")
                        Else
                            i% = InStr(s_str, "C/W LB PRICE")
                            If i% = 0 Then
                                q_txt = q_txt + Format$(dn!unit_cost_im, "000000.00000")
                            Else
                                i% = InStr(s_str, "C/W LB PRICE") + 12
                                c_val = Val(Mid$(s_str, i%, 8))
                                q_txt = q_txt + Format$(c_val, "000000.00000")
                            End If
                        End If
                    Else
                        q_txt = q_txt + Format$(dn!unit_cost_im, "000000.00000")
                    End If
                End If
            End If
            q_txt = q_txt & Format$(dn!ord_qty_im, "0000000")
            
            q_txt = q_txt + dn!abc_class_im
            q_txt = q_txt + Format$(dn!wght_im, "0000000.000")
            q_txt = q_txt + Format$(dn!att_flg_im, "@")
            GG_TBL.Seek "=", dn!vendor_im, dn!whse_im
            If GG_TBL.NoMatch Then
                Llt = 0
                ord_fr = 0
            Else
                Select Case dn!abc_class_im
                    Case "A"
                        Llt = GG_TBL!a_lead_time_v
                        ord_fr = GG_TBL!a_ord_fr_v
                    Case "B"
                        Llt = GG_TBL!b_lead_time_v
                        ord_fr = GG_TBL!b_ord_fr_v
                    Case "C"
                        Llt = GG_TBL!c_lead_time_v
                        ord_fr = GG_TBL!c_ord_fr_v
                    Case Else
                        Llt = 0
                        ord_fr = 0
                End Select
            End If
            'q_txt = q_txt & PO_string(Llt, 6)
            q_txt = q_txt & PO_string(dn!lead_time_im, 6)
            
            On Error Resume Next
            q_txt = q_txt & PO_string(Trim(GG_TBL!org_v), 30)
            q_txt = q_txt & PO_string(Trim(GG_TBL!address1_v), 30)
            q_txt = q_txt & PO_string(Trim(GG_TBL!address2_v), 30)
            q_txt = q_txt & PO_string(Trim(GG_TBL!city_v), 15)
            q_txt = q_txt & PO_string(Trim(GG_TBL!state_v), 2)
            q_txt = q_txt & PO_string(Trim(GG_TBL!zip_v), 10)
            q_txt = q_txt & PO_string(Trim(GG_TBL!phone_v), 15)
            q_txt = q_txt & PO_string(Trim(GG_TBL!fax_v), 15)
            q_txt = q_txt & PO_string(Trim(GG_TBL!attn_v), 30)
            q_txt = q_txt & PO_string(Trim(GG_TBL!buyer_v), 15)
            
            q_txt = q_txt & Space$(5)
            q_txt = q_txt & Space$(30)
            q_txt = q_txt & Space$(30)
            dm = Space$(30)
            LSet dm = Trim(Mid$(G_combo1, InStr(G_combo1, "-") + 1))
            q_txt = q_txt & dm
            GG_TBL.Close
            If atblc.EOF Then
                q_txt = q_txt & Space$(15)
            Else
                If InStr(s_str, "COST: $") = 0 Then
                    q_txt = q_txt & Space$(15)
                Else
                    i% = InStr(s_str, "COST: $") + 7
                    c_val = Val(Mid$(s_str, i%, InStr(i%, s_str, " ") - 1))
                    q_txt = q_txt + Format$(c_val, "000000000.00000")
                End If
            End If
            If atblc.EOF Then
                q_txt = q_txt & "  "
                q_txt = q_txt & Space$(10)
                q_txt = q_txt & Space$(15)
                q_txt = q_txt & Space$(15)
                q_txt = q_txt & Space$(30)
            Else
                If InStr(s_str, "FREIGHT TYPE:") = 0 Then
                    q_txt = q_txt & "  "
                Else
                    i% = InStr(s_str, "FREIGHT TYPE:") + 13
                    t_str = Space$(2)
                    LSet t_str = Mid$(s_str, i%, InStr(i%, s_str, " ") - 1)
                    q_txt = q_txt & t_str
                End If
                q_txt = q_txt & Space$(10)
                q_txt = q_txt & Space$(15)
                If InStr(atblc!descr, "FREIGHT AMOUNT: $") = 0 Then
                    q_txt = q_txt & Space$(15)
                Else
                    i% = InStr(s_str, "FREIGHT AMOUNT: $") + 17
                    c_val = Val(Mid$(s_str, i%, InStr(i%, s_str, " ") - 1))
                    q_txt = q_txt + Format$(c_val, "000000000.00000")
                End If
                If InStr(s_str, "EFFECTIVE DATE:") = 0 Then
                    q_txt = q_txt & Space$(30)
                Else
                    i% = InStr(s_str, "EFFECTIVE DATE:") + 15
                    t_str = Space$(30)
                    LSet t_str = Mid$(s_str, i%, 10)    'InStr(i% + 1, s_str, " ") - 1)
                    q_txt = q_txt & t_str
                End If
            End If
            atblc.Close
            q_txt = q_txt & Space$(15)
            If InStr(G_combo1, "PROMISE DATE") > 0 Then
                If Len(G_pdate) <> 8 Then
                    q_txt = q_txt & Format(CDate(G_pdate), "mm/dd/yy")
                Else
                    q_txt = q_txt & G_pdate
                End If
                If Err > 0 Then MsgBox "ascii Error: " & Err.Description & "-" & G_pdate & "-"
            Else
                date_var = Format(Now, "mm/dd/yy")
                plus_days = Int(Llt + dn!ord_freq_im - ord_fr + 0.5)
                res_date = Format(Date + plus_days, "mm/dd/yy")
                end_line = res_date
                q_txt = q_txt & end_line
            End If
            q_txt = q_txt & Left$(G_combo1, 2)
            Set ocTbl = DB.OpenRecordset("select * from odesc where " & _
                                         "item='" & dn!item_num_im & "' and " & _
                                         "whse='" & dn!whse_im & "' and " & _
                                         "vendor='" & dn!vendor_im & "'")
            If Not ocTbl.EOF Then
                dm = Space$(20)
                LSet dm = ocTbl!descr
                q_txt = q_txt & dm
            Else
                q_txt = q_txt & PO_string(G_pormemotxt, 20)
            End If
            ocTbl.Close
            q_txt = q_txt + PO_string(Trim(dn!s_vendor_im), 9)
            If IsNull(dn!brand_im) Then
                q_txt = q_txt + Space$(4)
            Else
                q_txt = q_txt + PO_string(Trim(dn!brand_im), 4)
            End If
            If IsNull(dn!brand_2_im) Then
                q_txt = q_txt + Space$(4)
            Else
                q_txt = q_txt + PO_string(Trim(dn!brand_2_im), 4)
            End If
            q_txt = q_txt + Format$(dn!price_brk_im, "0000000.000")
            
            Print #FilePO, q_txt
            
            If G_I_VXref Then
                If ans% = vbYes Then
                    iTbl.Open "select * from itemxref where " & _
                              "mitem_num = '" & dn!item_num_im & "' and " & _
                              "mitem_loc = '" & dn!whse_im & "'"
                    If Not iTbl.EOF Then
                        tmp_item = iTbl!vitem_num
                        tmp_item_desc = iTbl!vitem_desc
                    Else
                        tmp_item = dn!item_num_im
                        tmp_item_desc = dn!item_desc_im
                    End If
                    iTbl.Close
                End If
            End If
            
            TblPO.AddNew
            TblPO!por = Format$(G_PON, "00000000")
            
            TblPO!blanked = G_combo1
            TblPO!poline = Format$(rec_counter, "0000")
            TblPO!vnd = Trim(dn!vendor_im)
            If IsNull(TblPO!vnd) = True Or TblPO!vnd = "" Then TblPO!vnd = Space(9)
            TblPO("itm") = tmp_item                    'Trim(dn("item_num_im"))
            If IsNull(TblPO("itm")) = True Or TblPO("itm") = "" Then TblPO("itm") = Space(28)
            TblPO("idesc") = tmp_item_desc             'Trim(dn("item_desc_im"))
            If IsNull(TblPO("idesc")) = True Or TblPO("idesc") = "" Then TblPO("idesc") = Space(30)
            TblPO!um = Trim(dn!unit_meas_im)
            If IsNull(TblPO!um) = True Or TblPO!um = "" Then TblPO!um = Space(4)
            TblPO("abc") = dn!abc_class_im
            If IsNull(TblPO("abc")) = True Or TblPO("abc") = "" Then TblPO("abc") = Space(1)
            TblPO("cost") = Format(dn!unit_cost_im, "#####0.0000")
            TblPO("weight") = Format(dn!wght_im, "0000000.000")
            TblPO("attn") = Format$(dn!att_flg_im, "@")
            If IsNull(TblPO("attn")) = True Or TblPO("attn") = "" Then TblPO("attn") = Space(1)
            g_tbl.Open "select * from vendor where " & _
                       "vendor_v='" & Trim(dn!vendor_im) & "' and " & _
                       "whse_v='" & Trim(dn!whse_im) & "'"
            If g_tbl.EOF Then
                Llt = 0
                ord_fr = 0
            Else
                Select Case dn!abc_class_im
                    Case "A"
                        Llt = g_tbl!a_lead_time_v
                        ord_fr = g_tbl!a_ord_fr_v
                    Case "B"
                        Llt = g_tbl!b_lead_time_v
                        ord_fr = g_tbl!b_ord_fr_v
                    Case "C"
                        Llt = g_tbl!c_lead_time_v
                        ord_fr = g_tbl!c_ord_fr_v
                    Case Else
                        Llt = 0
                        ord_fr = 0
                End Select
            End If

            TblPO("lead") = Llt
            TblPO("freq") = ord_fr
            TblPO!BFreq = dn!ord_freq_im
            TblPO!qorder = Format(dn!ord_qty_im, "######0")
            TblPO!whse_f = Trim(dn!whse_im)
            On Error Resume Next
            TblPO!name_f = Trim(g_tbl!org_v)
            If IsNull(TblPO!name_f) = True Or TblPO!name_f = "" Then TblPO!name_f = Space(30)

            TblPO!address1_f = Trim(g_tbl!address1_v)
            If IsNull(TblPO!address1_f) = True Or TblPO!address1_f = "" Then TblPO!address1_f = Space(30)

            TblPO!address2_f = Trim(g_tbl!address2_v)
            If IsNull(TblPO!address2_f) = True Or TblPO!address2_f = "" Then TblPO!address2_f = Space(30)
                                        
            TblPO!city_f = Trim(g_tbl!city_v)
            If IsNull(TblPO!city_f) = True Or TblPO!city_f = "" Then TblPO!city_f = Space(15)

            TblPO!state_f = Trim(g_tbl!state_v)
            If IsNull(TblPO!state_f) = True Or TblPO!state_f = "" Then TblPO!state_f = Space(2)

            TblPO!zip_f = Trim(g_tbl!zip_v)
            If IsNull(TblPO!zip_f) = True Or TblPO!zip_f = "" Then TblPO!zip_f = Space(10)

            TblPO!phone_f = Trim(g_tbl!phone_v)
            If IsNull(TblPO!phone_f) = True Or TblPO!phone_f = "" Then TblPO!phone_f = Space(15)

            TblPO!fax_f = Trim(g_tbl!fax_v)
            If IsNull(TblPO!fax_f) = True Or TblPO!fax_f = "" Then TblPO!fax_f = Space(15)

            TblPO!attn_f = Trim(g_tbl!attn_v)
            If IsNull(TblPO!attn_f) = True Or TblPO!attn_f = "" Then TblPO!attn_f = Space(30)
            
            TblPO!buyer_f = Trim(g_tbl!buyer_v)
            If IsNull(TblPO!buyer_f) = True Or TblPO!buyer_f = "" Then TblPO!buyer_f = Space(15)
            TblPO!whse_t = Space(2)
            TblPO!name_t = Space(30)
            TblPO!address1_t = Space(30)
            TblPO!address2_t = Space(30)
            g_tbl.Close
            atblc.Open "select * from xdesc where item='" & dn!item_num_im & "' and " & _
                       "whse='" & dn!whse_im & "'"
            If atblc.EOF Then
                TblPO!city_t = Space$(15)
            Else
                If InStr(atblc!descr, "COST: $") = 0 Then
                    TblPO!city_t = Space$(15)
                Else
                    i% = InStr(atblc!descr, "COST: $") + 7
                    c_val = Val(Mid$(atblc!descr, i%, InStr(i%, atblc!descr, " ") - 1))
                    TblPO!city_t = Format$(c_val, "000000000.00000")
                End If
            End If
            If InStr(G_combo1, "NDS FREIGHT") > 0 Then
                TblPO!state_t = FType
                TblPO!zip_t = FFactor
                TblPO!phone_t = FDollars
                If Trim(FType) = "W" Then
                    TblPO!fax_t = Format((o_wght / FFactor) * FDollars, "000000000000.00")
                Else
                    TblPO!fax_t = Format((o_qty / FFactor) * FDollars, "000000000000.00")
                End If
            Else
                TblPO!state_t = Space(2)
                TblPO!zip_t = Space(10)
                TblPO!phone_t = Space(15)
                TblPO!fax_t = Space(15)
            End If
            TblPO!attn_t = Space(30)
            TblPO!buyer_t = Space(15)

            If InStr(G_combo1, "PROMISE DATE") > 0 Then
                TblPO!dop = Format(CVDate(G_pdate), "mm/dd/yy")
            Else
                date_var = Format(Now, "mm/dd/yy")
                plus_days = Int(Llt + dn!ord_freq_im - ord_fr + 0.5)
                res_date = Format(Date + plus_days, "mm/dd/yy")
                end_line = res_date
                PROMISE_DATE = end_date
                TblPO!dop = end_line
            End If
            TblPO!doo = Format(Now, "mm/dd/yy")
            Set ocTbl = DB.OpenRecordset("select * from odesc where " & _
                                         "item='" & dn!item_num_im & "' and " & _
                                         "whse='" & dn!whse_im & "' and " & _
                                         "vendor='" & dn!vendor_im & "'")
            If Not ocTbl.EOF Then
                dm = Space$(20)
                LSet dm = ocTbl!descr
                TblPO!memo1 = dm
            Else
                TblPO!memo1 = G_pormemotxt
            End If
            ocTbl.Close
            TblPO("s_vendor") = Trim(dn!s_vendor_im)
            If IsNull(TblPO("s_vendor")) = True Or TblPO("s_vendor") = "" Then TblPO("s_vendor") = Space(9)
            If IsNull(dn!brand_im) Then
                TblPO!group1 = Space$(4)
            Else
                TblPO!group1 = Trim(dn!brand_im)
            End If
            If IsNull(TblPO!group1) Or Trim(TblPO!group1) = "" Then TblPO!group1 = Space(4)
            If IsNull(dn!brand_2_im) Then
                TblPO!group2 = Space$(4)
            Else
                TblPO!group2 = Trim(dn!brand_2_im)
            End If
            If IsNull(TblPO!group2) Or Trim(TblPO!group2) = "" Then TblPO!group2 = Space(4)
            TblPO!volume = Trim(dn!price_brk_im)
            If IsNull(TblPO!volume) = True Or TblPO!volume = "" Then TblPO!volume = 0
            Err = 0
            On Error Resume Next
            TblPO.Update
            If Err > 0 Then
                MsgBox "error: " & Str$(Err) & "-" & Err.Description
            End If
            If G_Buy_Perform Then
                GoSub buy_perform_logic
            End If
            On Error Resume Next
            atblc.Close
get_next:
            dn.MoveNext
            If dn.EOF Then Exit Do
        Loop
    End If
End If
dn.Close
TblPO.Close
If G_Warehouse <> "" Then
    DB.Execute "delete from orders where 1=1"
End If
Exit Sub

buy_perform_logic:
    Dim g$
    If G_Buy_Perform Then
        If dn!re_ord_nor_im <> dn!ord_qty_im Then
        
            g$ = "select * from buy_perform where " & _
                 "buyer = '" & Trim(G_Buyer) & "' and " & _
                 "bdate = '" & Format(Now, "yyyymmdd") & "' and " & _
                 "item_num = '" & Trim(dn!item_num_im) & "' and " & _
                 "loc = '" & Trim(dn!whse_im) & "'"
            
            Err = 0
            bftbl.Open g$, Cnct
            If Err <> 0 Then
                MsgBox "Buy Perform Error1: " & Str$(Err) & "  " & Err.Description
            End If
            If bftbl.EOF Then
                bftbl.AddNew
            End If
            bftbl!Buyer = G_Buyer
            bftbl!bdate = Format(Date$, "yyyymmdd")
            bftbl!item_num = dn!item_num_im
            bftbl!Loc = dn!whse_im
            bftbl!Vendor = dn!vendor_im
            bftbl!o_qty = dn!re_ord_nor_im
            bftbl!n_qty = dn!ord_qty_im
            bftbl!reason_f = dn!r_monf_im
            Err = 0
            bftbl.Update
            If Err <> 0 Then
                MsgBox "Buy Perform Error2: " & Str$(Err) & "  " & Err.Description
            End If
            bftbl.Close
        End If
    End If
    Return



End Sub



