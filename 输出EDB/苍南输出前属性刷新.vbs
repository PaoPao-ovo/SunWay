

Sub OnClick()
    KeyStr = "���,��Ŀ����,��Ŀ��ַ,��Ƶ�λ,���赥λ,ί�е�λ,��ҵʱ��,���ʱ��,�����ϲ�ֵ,�߳����ϲ�ֵ"
    SSProcess.ClearInputParameter
    
    KeyArr = Split(KeyStr,",", - 1,1)
    
    For i = 0 To UBound(KeyArr) - 2
        SSProcess.AddInputParameter KeyArr(i) , SSProcess.ReadEpsIni("���߱�����Ϣ", KeyArr(i) ,"") , 0 , "" , ""
    Next 'i
    
    ShowBoolen = SSProcess.ShowInputParameterDlg ("���߱�����Ϣ¼��")
    
    For i = 0 To UBound(KeyArr)
        SSProcess.WriteEpsIni "���߱�����Ϣ", KeyArr(i) ,SSProcess.GetInputParameter(KeyArr(i))
    Next 'i

    GETNOTEZB
    
    SSProcess.OpenAccessMdb SSProcess.GetProjectFileName
    
    SqlStr = "Update ���¹��ߵ����Ա� SET XMBH = " & "'" & SSProcess.ReadEpsIni("���߱�����Ϣ", "���" ,"") & "'"
    SsProcess.ExecuteAccessSql SSProcess.GetProjectFileName,SqlStr
    
    SqlStr = "Update ���¹��������Ա� SET XMBH = " & "'" & SSProcess.ReadEpsIni("���߱�����Ϣ", "���" ,"") & "'"
    SsProcess.ExecuteAccessSql SSProcess.GetProjectFileName,SqlStr
    
    SSProcess.CloseAccessMdb SSProcess.GetProjectFileName
    
    SSProcess.MapMethod "clearattrbuffer",  "���¹��ߵ����Ա�"
    SSProcess.MapMethod "clearattrbuffer",  "���¹��������Ա�"
End Sub' OnClick

'ˢע������ֵ
Function GETNOTEZB
    projectName = SSProcess.GetProjectFileName
    sql = "Select ���¹��ߵ����Ա�.id,GXDDH From ���¹��ߵ����Ա� INNER JOIN GeoPOINTTB ON ���¹��ߵ����Ա�.ID = GeoPOINTTB.ID WHERE ([GeopointTB].[Mark] Mod 2)<>0"
    GetSQLRecordAll projectName,sql,arSQLRecord,iRecordCount
    If iRecordCount > 0 Then
        For i = 0 To iRecordCount - 1
            arTemp = Split(arSQLRecord(i), ",")
            gdid = arTemp(0)
            gxddh = arTemp(1)
            SSProcess.ClearSelection
            SSProcess.ClearSelectCondition
            SSProcess.SetSelectCondition "SSObj_FontString", "==", gxddh
            SSProcess.SelectFilter
            gdCount = SSProcess.GetSelnoteCount
            If gdCount > 0 Then
                X = Round(SSProcess.GetSelnoteValue(0,"SSObj_X"),3)
                Y = Round(SSProcess.GetSelnoteValue(0,"SSObj_Y"),3)
                SSProcess.SetObjectAttr gdid, "[MapNo_Y]", X
                SSProcess.SetObjectAttr gdid, "[MapNo_X]",Y
            End If
        Next
    End If
    'δ�ÿ���ˢ��
    sql = "Select ���¹��������Ա�.id,zks,yyks From ���¹��������Ա� INNER JOIN GeolineTB ON ���¹��������Ա�.ID = GeolineTB.ID WHERE ([GeolineTB].[Mark] Mod 2)<>0"
    GetSQLRecordAll projectName,sql,arSQLRecord,iRecordCount
    If iRecordCount > 0 Then
        SSProcess.OpenAccessMdb projectName
        For i = 0 To iRecordCount - 1
            arTemp = Split(arSQLRecord(i), ",")
            gxid = arTemp(0)
            zks = arTemp(1)
            yyks = arTemp(2)
            lash = IsNumeric (zks)
            If    lash = True Then
                If yyks = "" Then yyks = 0
                wyks = CDbl(zks) - CDbl(yyks)
                sql = "update  ���¹��������Ա� set wyks='" & wyks & "' where ���¹��������Ա�.id=" & gxid
                SSProcess.ExecuteAccessSql  projectName,sql
            End If
        Next
        SSProcess.MapMethod "clearattrbuffer",  "���¹��������Ա�"
        SSProcess.CloseAccessMdb mdbName
    End If
End Function

Function GetSQLRecordAll(ByRef mdbName, ByRef StrSqlStatement, ByRef arSQLRecord(), ByRef iRecordCount)
    SSProcess.OpenAccessMdb mdbName
    iRecordCount =  - 1
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
        '�����¼
        While SSProcess.AccessIsEOF (mdbName, sql) = 0
            fields = ""
            values = ""
            '��ȡ��ǰ��¼����
            SSProcess.GetAccessRecord mdbName, sql, fields, values
            arSQLRecord(iRecordCount) = values                                        '��ѯ��¼
            iRecordCount = iRecordCount + 1                                                    '��ѯ��¼��
            '�ƶ���¼�α�
            SSProcess.AccessMoveNext mdbName, sql
        WEnd
    End If
    '�رռ�¼��
    SSProcess.CloseAccessRecordset mdbName, sql
    SSProcess.CloseAccessMdb mdbName
End Function