
'===================================================����������==================================================

'���������
Dim strGroupName

'���������
Dim strCheckName

'���ģ������
Dim CheckmodelName

'�������
Dim strDescription

'====================================================���=========================================================

'������
Sub OnClick()
    
    ClearCheckRecord
    
    FWCheck
    
    ShowCheckRecord

End Sub' OnClick

'===================================================��麯��=======================================================

'������;�빦������;�������ֵ�Ƿ�һ�£����д���
Function FWCheck()
    
    ' 1����Ҫ����ָ�����������Ϣ��(ZYJJZBMJHZB)�е�ÿ����YT�������磺סլ �����SCJZMJ��
    ' 2���滮��������GHGNQ�����еġ�YT�� = ��סլ�����������ֵ��
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "������;�빦������;�������ֵһ���Լ��"
    CheckmodelName = "�Զ���ű������->������;�빦������;�������ֵһ���Լ��"
    strDescription = "������;�빦������;�������ֵ��һ��"

    SqlStr = "Select DISTINCT ZYJJZBMJHZB.YT From ZYJJZBMJHZB Where ZYJJZBMJHZB.ID > 0"
    GetSQLRecordAll SqlStr,YTArr,YTCount
    
    For i = 0 To YTCount - 1
        
        SqlStr = "Select Sum(JG_�滮���������Ա�.JZMJ) From JG_�滮���������Ա� Inner Join GeoAreaTB On JG_�滮���������Ա�.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And JG_�滮���������Ա�.YT = '" & YTArr(i) & "'"
        GetSQLRecordAll SqlStr,SumAreaArr,SumCount
        SumArea = SumAreaArr(0)
        
        SqlStr = "Select ZYJJZBMJHZB.SCJZMJ From ZYJJZBMJHZB Where ZYJJZBMJHZB.ID > 0 And ZYJJZBMJHZB.YT = '" & YTArr(i) & "'"
        GetSQLRecordAll SqlStr,SCJZMJArr,SearchCount
        SCJZMJ = SCJZMJArr(0)
        
        If SumArea - SCJZMJ <> 0 Then
            SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
        End If
    Next 'i

End Function' FWCheck

'======================================================�����ຯ��====================================================

'��ջ�������м���¼
Function ClearCheckRecord()
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
End Function' ClearCheckRecord

'��ʾ���м���¼
Function ShowCheckRecord()
    SSProcess.ShowCheckOutput
    SSProcess.SaveCheckRecord
End Function' ShowCheckRecord

'��ȡ���м�¼
Function GetSQLRecordAll(ByVal StrSqlStatement, ByRef SQLRecord(), ByRef iRecordCount)
    ProJectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb ProJectName
    If StrSqlStatement = "" Then
        MsgBox "��ѯ���Ϊ�գ�����ֹͣ��",48
    End If
    iRecordCount =  - 1
    SSProcess.OpenAccessRecordset ProJectName, StrSqlStatement
    RecordCount = SSProcess.GetAccessRecordCount (ProJectName, StrSqlStatement)
    If RecordCount > 0 Then
        iRecordCount = 0
        ReDim SQLRecord(RecordCount)
        SSProcess.AccessMoveFirst ProJectName, StrSqlStatement
        iRecordCount = 0
        While SSProcess.AccessIsEOF (ProJectName, StrSqlStatement) = 0
            fields = ""
            values = ""
            SSProcess.GetAccessRecord ProJectName, StrSqlStatement, fields, values
            SQLRecord(iRecordCount) = values
            iRecordCount = iRecordCount + 1
            SSProcess.AccessMoveNext ProJectName, StrSqlStatement
        WEnd
    End If
    SSProcess.CloseAccessRecordset ProJectName, StrSqlStatement
    SSProcess.CloseAccessMdb ProJectName
End Function

'��������ת��
Function Transform(ByVal Values)
    If Values <> "" Then
        If IsNumeric(Values) = True Then
            Values = CDbl(Values)
        End If
    Else
        Values = 0
    End If
    Transform = Values
End Function'Transform