'=============================================�������ֶ�����=========================================================

'���������Ա���
Dim TableArr(2)

TableArr(0) = "���۷��������Ա�"
TableArr(1) = "ʵ����������Ա�"

'=====================================================��鼯����=====================================================

'��鼯��Ŀ����
Dim strGroupName
strGroupName = "���߼��"

'��鼯������
Dim strCheckName
strCheckName = "��������"

'���ģ������
Dim CheckmodelName
CheckmodelName = "�Զ���ű������->��������"

'�������
Dim strDescription
strDescription = "������������չ����Ϊ��"

'==============================================��������========================================================

'��ں���
Sub OnClick()
    ClearCheckRecord()
     For i = 0 To 1
        PoiCheck TableArr(i)
    Next 'i
End Sub' OnClick

'���������Լ��
Function PoiCheck(tablename)
    projectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb projectName
    SqlString = "Select * From " & tablename & "  inner join GeoPointTB on " & tablename & ".ID = GeoPointTB.ID WHERE (GeoPointTB.Mark Mod 2)<>0"
    'MsgBox SqlString
    GetSQLRecordAll projectName,SqlString,arSQLRecord,iRecordCount
    For j = 0 To iRecordCount - 1
        RecordString = arSQLRecord(j)
        'MsgBox RecordString
        Recordarr = Split(RecordString,",", - 1,1)
            If Recordarr(1) = "*" Then
                id = Recordarr(0)
                x = SSProcess.GetObjectAttr(id,"SSObj_X")
                y = SSProcess.GetObjectAttr(id,"SSObj_Y")
                SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription, x, y,0,0,id, ""
            End If
            If Recordarr(1) = "" Then
                id = Recordarr(0)
                x = SSProcess.GetObjectAttr(id,"SSObj_X")
                y = SSProcess.GetObjectAttr(id,"SSObj_Y")
                SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription, x, y,0,0,id, ""
            End If
            If Recordarr(1) = "0" Then
                id = Recordarr(0)
                x = SSProcess.GetObjectAttr(id,"SSObj_X")
                y = SSProcess.GetObjectAttr(id,"SSObj_Y")
                SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription, x, y,0,0,id, ""
            End If
            If Recordarr(1) = Null Then
                id = Recordarr(0)
                x = SSProcess.GetObjectAttr(id,"SSObj_X")
                y = SSProcess.GetObjectAttr(id,"SSObj_Y")
                SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription, x, y,0,0,id, ""
            End If
        'MsgBox Recordarr(0)
    Next 'j
    SSProcess.CloseAccessMdb projectName
    SSProcess.ShowCheckOutput
    SSProcess.SaveCheckRecord
End Function' ZDX

'��ȡ���м�¼
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

'��ռ�鼯
Function ClearCheckRecord()
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
End Function' ClearCheckRecord