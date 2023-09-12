Dim DKID(500)
Dim DWX(500)
Dim DWY(500)
Dim arID(10)
Sub OnClick()
    UPDATEAREA
    dkpx "9530226"
    dkpx "9530225"
    YBQFZ
    'exit sub
    mdbName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb mdbName
    sql = "select FHDYBH from �˷�������Ԫ���Ա� inner join GeoAreaTB on GeoAreaTB.ID = �˷�������Ԫ���Ա�.ID where [GeoAreaTB].[Mark] Mod 2<>0 "
    GetSQLRecordAll mdbName,sql,arSQLRecord,iRecordCount
    If iRecordCount < 0 Then Exit Sub
    For i = 0 To iRecordCount - 1
        FHDYBH = arSQLRecord(i)
        '��ȡ�ڱ������
        sql1 = "select sum(YBMJ) from �ڱε�Ԫ���Ա� inner join GeoAreaTB on GeoAreaTB.ID = �ڱε�Ԫ���Ա�.ID where [GeoAreaTB].[Mark] Mod 2<>0 and FHDYBH = '" & FHDYBH & "'"
        GetSQLRecordAll mdbName,sql1,arYBRecrod,strYBCount
        If strYBCount = 1 Then
            sql2 = "update �˷�������Ԫ���Ա� set YBMJ = " & arYBRecrod(0) & " where FHDYBH = '" & FHDYBH & "'"
            SSProcess.ExecuteAccessSql mdbName,sql2
        End If
    Next
    SSProcess.CloseAccessMdb mdbName
    SSProcess.MapMethod "clearattrbuffer", "�˷�������Ԫ���Ա�"
    
    '������ϵ���ķǻ�����λ����
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==","9461023,9461043"
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    For i = 0 To geocount - 1
        id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        ZSXS = SSProcess.GetSelGeoValue(i, "[ZheSXS]")
        ZSXS = CDbl(ZSXS)
        If ZSXS <> 0.0 Then CWSL = CInt(geocount * ZSXS)
        SSProcess.SetObjectAttr id, "[CheWGS]", CWSL
    Next
    
    '������/�ǻ�����λ�������ܵ�������Ԫ��
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "=","9530226"
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    FJDCWCount = CWSL
    JDCWCount = 0
    WXCWCount = 0
    'CW_������ͣ��λ��Ϣ���Ա�
    For i = 0 To geocount - 1
        id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        ids = SSProcess.SearchInnerObjIDs(id, 2, "9461013,9461033,9461053", 0)
        If ids <> "" Then
            strList = Split(ids,",")
            For i1 = 0 To UBound(strList)
                code = SSProcess.GetObjectAttr (strList(i1), "SSObj_Code")
                If code = 9461013 Or code = 9461033 Then
                    CWLX = SSProcess.GetObjectAttr(strList(i1),"[CheWLX]")
                    ZSXS = CDbl(SSProcess.GetObjectAttr(strList(i1),"[ZSXS]"))
                    If CWLX = "���ͳ�λ" Then
                        JDCWCount = JDCWCount + ZSXS
                    Else
                        JDCWCount = JDCWCount + 1
                    End If
                ElseIf code = 9461053 Then
                    WXCWCount = WXCWCount + 1
                End If
            Next
        End If
        SSProcess.SetObjectAttr id, "[TCWS]", CInt(JDCWCount) + CInt(WXCWCount * 0.7)
        SSProcess.SetObjectAttr id, "[FJDCS]", FJDCWCount
    Next
End Sub


Function GetSQLRecordAll(ByRef mdbName, ByRef StrSqlStatement, ByRef arSQLRecord(), ByRef iRecordCount)
    If StrSqlStatement = "" Then
        MsgBox "��ѯ���Ϊ�գ�����ֹͣ��",48
    End If
    iRecordCount =  - 1
    'SQL���
    sql = StrSqlStatement
    '�򿪼�¼��
    SSProcess.OpenAccessRecordset mdbName, sql
    '��ȡ��¼����
    RecordCount = SSProcess.GetAccessRecordCount (mdbName, sql)
    If RecordCount > 0 Then
        iRecordCount = 0
        ReDim arSQLRecord(RecordCount)
        '����¼�α��Ƶ���һ��
        SSProcess.AccessMoveFirst mdbName, sql
        iRecordCount = 0
        '�����¼
        While SSProcess.AccessIsEOF (mdbName, sql) = 0
            fields = ""
            values = ""
            '��ȡ��ǰ��¼����
            SSProcess.GetAccessRecord mdbName, sql, fields, values
            arSQLRecord(iRecordCount) = values
            iRecordCount = iRecordCount + 1
            '�ƶ���¼�α�
            SSProcess.AccessMoveNext mdbName, sql
        WEnd
    End If
    '�رռ�¼��
    SSProcess.CloseAccessRecordset mdbName, sql
End Function
'�����˷���Ԫ���ڱ������
Function UPDATEAREA
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "=","9530226"
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    For I = 0 To GEOCOUNT - 1
        id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        AREA = SSProcess.GetSelGeoValue(i, "SSObj_AREA")
        If AREA <> "" Then AREA = FormatNumber(AREA,2, - 1,0,0)
        
        SSProcess.SetObjectAttr ID, "[JZMJ]", AREA
    Next
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "=","9530225"
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    For I = 0 To GEOCOUNT - 1
        id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        AREA = SSProcess.GetSelGeoValue(i, "SSObj_AREA")
        If AREA <> "" Then AREA = FormatNumber(AREA,2, - 1,0,0)
        MsgBox area
        SSProcess.SetObjectAttr  ID, "[YBMJ]", AREA
    Next
End Function

Function dkpx(CODE)
    SSProcess.ClearInputParameter
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Type", "==", "AREA"
    SSProcess.SetSelectCondition "SSObj_Code", "==", CODE
    SSProcess.SelectFilter
    gcount = SSProcess.GetSelGeoCount()
    If gcount > 0 Then
        For i = 0 To gcount - 1
            DKID(I) = SSProcess.GetSelGeoValue(i, "SSObj_ID")
            SSProcess.GetObjectFocusPoint DKID(I), DWX(I), DWY(I)
            If I = 0 Then
                XMIN = DWX(I)
                YMIN = DWY(I)
                XMAX = DWX(I)
                YMAX = DWY(I)
            Else
                If XMIN > DWX(I) Then XMIN = DWX(I)
                If YMIN > DWY(I) Then YMIN = DWY(I)
                If XMAX < DWX(I)Then XMAX = DWX(I)
                If YMAX < DWY(I) Then YMAX = DWY(I)
            End If
        Next
        '�Ƚ�x��ֵ y��ֵ
        xcz = xmax - xmin
        ycz = ymax - ymin
        
        If xcz > ycz Then
            '�����Ҹ�ֵ�ؿ��
            For J = 0 To gcount - 1
                '��8����x �Եؿ�����
                For k = j + 1 To gcount - 1
                    If DWX(k) < x(j) Then
                        a = DWX(j)
                        DWX(j) = DWX(k)
                        DWX(k) = a
                        b = DKID(j)
                        DKID(j) = DKID(k)
                        DKID(k) = b
                    End If
                Next
            Next
        Else
            '���ϵ��¸�ֵ�ؿ��
            For J = 0 To gcount - 1
                '��8����x �Եؿ�����
                For k = j + 1 To gcount - 1
                    If DWY(k) > DWY(j) Then
                        a = DWY(j)
                        DWY(j) = DWY(k)
                        DWY(k) = a
                        b = DKID(j)
                        DKID(j) = DKID(k)
                        DKID(k) = b
                    End If
                Next
            Next
        End If
        If CODE = "9530226" Then
            For J = 0 To gcount - 1
                SSProcess.SetObjectAttr DKID(j),"[FHDYBH]","������Ԫ���" & j + 1
            Next
        Else
            For J = 0 To gcount - 1
                SSProcess.SetObjectAttr DKID(j),"[YBDYBH]","�ڱε�Ԫ" & j + 1
            Next
        End If
    End If
End Function

Function YBQFZ
    SSProcess.ClearInputParameter
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Type", "==", "AREA"
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9530226"
    SSProcess.SelectFilter
    gcount = SSProcess.GetSelGeoCount()
    If gcount > 0 Then
        For i = 0 To gcount - 1
            ID = SSProcess.GetSelGeoValue(i, "SSObj_ID")
            FHDYH = SSProcess.GetSelGeoValue(i, "[FHDYBH]")
            IDS = SSProcess.SearchInnerObjIDs(ID,2,"9530225",0)
            If ids <> "" Then
                SSFunc.ScanString ids, ",", arID, idCount
                For k = 0 To idCount - 1
                    SSProcess.SetObjectAttr CInt(arID(k)), "[FHDYBH]", FHDYH
                Next
            End If
        Next
    End If
End Function