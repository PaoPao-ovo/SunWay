
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

    JzZmjCheck "ZD_�ڵػ�����Ϣ���Ա�"
    
    DxJzzMjCheck
    
    DsJzzMjCheck
    
    HDSCheck
    
    HDXCheck
    
    ShowCheckRecord
    
End Sub' OnClick

'===================================================��麯��=======================================================

'ʵ��潨����������
Function JzZmjCheck(ByVal TableName)
    
    ' 1 ����������ڵػ�����Ϣ��JZMJ����ZD_�ڵػ�����Ϣ���Ա�[JZZMJ]��
    ' 2 ���ϲ����ܼƣ����ݵ��ϵ��������������Ϣ��FWDSDXZMJHZXX���ֶΣ���YCDSZJZMJ�����ֶΡ�SCDSZJZMJ��
    ' 3 ���ϲ����ܼƣ����ݵ��ϵ��������������Ϣ��FWDSDXZMJHZXX���ֶΣ���YCDXZJZMJ�����ֶΡ�SCDXZJZMJ��
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "���ݻ�����Ϣ��������߼����"
    CheckmodelName = "�Զ���ű������->���ݻ�����Ϣ��������߼����"
    strDescription = TableName & "�ġ�JZZMJ����FWDSDXZMJHZXX��ġ�SCDSZJZMJ���͡�SCDXZJZMJ��֮�Ͳ����"
    
    '��ȡ�ܽ������ JZZMJ
    SqlStr = "Select " & TableName & ".ID,JZZMJ From " & TableName & " Inner Join GeoAreaTB On " & TableName & ".ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 "
    GetSQLRecordAll SqlStr,TotalAreaArr,SearchCount
    
    If SearchCount > 1 Then
        ZDArr = Split(TotalAreaArr(0),",", - 1,1)
        JZZMJ = Transform(ZDArr(1))
    Else
        JZZMJ = 0
        Dim ZDArr(0)
        ZDArr(0) =  - 1
    End If
    
    '��ȡ�ܵ��Ͻ������ SCDSZMJ
    SqlStr = "Select Sum(FWDSDXZMJHZXX.SCDSZJZMJ) From FWDSDXZMJHZXX WHERE FWDSDXZMJHZXX.ID > 0"
    GetSQLRecordAll SqlStr,SCDSArr,SearchCount
    
    If SearchCount > 0 Then
        SCDSZMJ = Transform(SCDSArr(0))
    Else
        SCDSZMJ = 0
    End If
    
    '��ȡ�ܵ��½������ SCDXZMJ
    SqlStr = "Select Sum(FWDSDXZMJHZXX.SCDXZJZMJ) From FWDSDXZMJHZXX WHERE FWDSDXZMJHZXX.ID > 0"
    GetSQLRecordAll SqlStr,SCDXArr,SearchCount
    
    If SearchCount > 0 Then
        SCDXZMJ = Transform(SCDXArr(0))
    Else
        SCDXZMJ = 0
    End If
    
    SumArea = SCDSZMJ + SCDXZMJ
    
    '����ж�
    If JZZMJ - SumArea <> 0 Then
        If ZDArr(0) <> - 1 Then
            SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(ZDArr(0),"SSObj_X"),SSProcess.GetObjectAttr(ZDArr(0),"SSObj_Y"),0,2,ZDArr(0),""
        Else
            SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
        End If
        
    End If
    
End Function' JzZmjCheck 

'ʵ����½�����������
Function DxJzzMjCheck()
    
    ' 1:���²����ܼ�: ���ݵ��ϵ��������������Ϣ��FWDSDXZMJHZXX���ֶΣ���YCDXZJZMJ�����ֶΡ�SCDXZJZMJ��
    ' 2:��������+�˷���λ�������������������Ϣ��FWLXMJHZXX����SCJZMJ����YCJZMJ�����ۼƺ͡����������ƣ��ռ�λ�á�KJWZ��Ϊ�����£���
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "���ݻ�����Ϣ��������߼����"
    CheckmodelName = "�Զ���ű������->���ݻ�����Ϣ��������߼����"
    strDescription = "ʵ������ܽ������������ֺ��˷��������֮�Ͳ���"
    
    '��ȡ��������� SCDXZMJ
    SqlStr = "Select Sum(FWDSDXZMJHZXX.SCDXZJZMJ) From FWDSDXZMJHZXX WHERE FWDSDXZMJHZXX.ID > 0"
    GetSQLRecordAll SqlStr,SCDXArr,SearchCount
    
    If SearchCount > 0 Then
        SCDXZMJ = Transform(SCDXArr(0))
    Else
        SCDXZMJ = 0
    End If
    
    '������������������˷�������� QTMJ
    SqlStr = "Select Sum(FWLXMJHZXX.SCJZMJ) From FWLXMJHZXX WHERE FWLXMJHZXX.ID > 0 And FWLXMJHZXX.KJWZ = '����' "
    GetSQLRecordAll SqlStr,QTArr,SearchCount
    
    If SearchCount > 0 Then
        QTMJ = Transform(QTArr(0))
    Else
        QTMJ = 0
    End If
    
    If SCDXZMJ - QTMJ <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' DxJzzMjCheck

'ʵ����Ͻ�����������
Function DsJzzMjCheck()
    
    ' 1�����ϲ����ܼƣ����ݵ��ϵ��������������Ϣ��FWDSDXZMJHZXX���ֶΣ���YCDSZJZMJ�����ֶΡ�SCDSZJZMJ��
    ' 2: ���ϻ����ͳ��: �����������������Ϣ��FWLXMJHZXX����SCJZMJ����YCJZMJ�����ۼƺ͡����������ƣ��ռ�λ�á�KJWZ��Ϊ�����ϣ���
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "���ݻ�����Ϣ��������߼����"
    CheckmodelName = "�Զ���ű������->���ݻ�����Ϣ��������߼����"
    strDescription = "ʵ������ܽ������������ֺ��˷��������֮�Ͳ���"
    
    '��ȡ��������� SCDSZMJ
    SqlStr = "Select Sum(FWDSDXZMJHZXX.SCDSZJZMJ) From FWDSDXZMJHZXX WHERE FWDSDXZMJHZXX.ID > 0"
    GetSQLRecordAll SqlStr,SCDXArr,SearchCount
    
    If SearchCount > 0 Then
        SCDSZMJ = Transform(SCDXArr(0))
    Else
        SCDSZMJ = 0
    End If
    
    '������������������˷�������� QTMJ
    SqlStr = "Select Sum(FWLXMJHZXX.SCJZMJ) From FWLXMJHZXX WHERE FWLXMJHZXX.ID > 0 And FWLXMJHZXX.KJWZ = '����' "
    GetSQLRecordAll SqlStr,QTArr,SearchCount
    
    If SearchCount > 0 Then
        QTMJ = Transform(QTArr(0))
    Else
        QTMJ = 0
    End If
    
    If SCDSZMJ - QTMJ <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' DsJzzMjCheck

'H����ϼ��
Function HDSCheck()
    
    ' 1���������ͻ���ֵ�������������������Ϣ��FWLXMJHZXX���С�FWLXMC���͡�SCJZMJ���͡�KJWZ��
    ' 2������H����ʵ�ʲ�����SJCS���������������ơ�FWLXMC����Ԥ�⽨�������YCJZMJ����ʵ�⽨�������SCJZMJ����ֵ���ۼӺ͡���˵�������յ��ϡ����·ֱ����жϣ�
    ' ����˵�����������������������Ϣ��FWLXMJHZXX���ġ�KJWZ��=���� �ҡ�FWLXMC��=��סլ���ġ�SCJZMJ����ֵ�Ƿ���ڻ���H���ġ�SJCS������0�ҡ�FWLXMC��=��סլ���ġ�SCJZMJ����ֵ���ۼӺ͡�
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "���ݻ�����Ϣ��������߼����"
    CheckmodelName = "�Զ���ű������->���ݻ�����Ϣ��������߼����"
    strDescription = "���������������ֵ�뻧��ͳ�����ֵ��һ��"
    
    '��ȡ���еķ����������� FWLXMCArr
    SqlStr = "Select DISTINCT FWLXMJHZXX.FWLXMC From FWLXMJHZXX Where FWLXMJHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,FWLXMCArr,FWLXMCCount
    
    If FWLXMCCount > 0 Then
        '��ȡ��Ӧ��ʵ����Ͻ������
        For CurrentCount = 0 To UBound(FWLXMCArr)
            If FWLXMCArr(CurrentCount) <> "" Then
                
                SqlStr = "Select Sum(FWLXMJHZXX.SCJZMJ) From FWLXMJHZXX Where FWLXMJHZXX.ID > 0 And FWLXMJHZXX.FWLXMC = " & "'" & FWLXMCArr(CurrentCount) & "' And " & "FWLXMJHZXX.KJWZ = '����' "
                GetSQLRecordAll SqlStr,SCJZMJArr,SearchCount
                
                If SearchCount > 0 Then
                    SCJZMJ = Transform(SCJZMJArr(0))
                Else
                    SCJZMJ = 0
                End If
                
                SqlStr = "Select Sum(H.SCJZMJ) From H Where H.ID > 0 And H.FWLXMC = " & "'" & FWLXMCArr(CurrentCount) & "' And " & "H.KJWZ = '����' And H.SJCS > 0 "
                
                GetSQLRecordAll SqlStr,HSCJZMJArr,SearchCount
                
                If SearchCount > 0 Then
                    HSCJZMJ = Transform(HSCJZMJArr(0))
                Else
                    HSCJZMJ = 0
                End If
                
                If SCJZMJ - HSCJZMJ <> 0 Then
                    SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
                End If
                
            End If
        Next 'CurrentCount
    End If
    
End Function' HDSCheck

'H����¼��
Function HDXCheck()
    
    ' 1���������ͻ���ֵ�������������������Ϣ��FWLXMJHZXX���С�FWLXMC���͡�SCJZMJ���͡�KJWZ��
    ' 2������H����ʵ�ʲ�����SJCS���������������ơ�FWLXMC����Ԥ�⽨�������YCJZMJ����ʵ�⽨�������SCJZMJ����ֵ���ۼӺ͡���˵�������յ��ϡ����·ֱ����жϣ�
    ' ����˵�����������������������Ϣ��FWLXMJHZXX���ġ�KJWZ��=���� �ҡ�FWLXMC��=��סլ���ġ�SCJZMJ����ֵ�Ƿ���ڻ���H���ġ�SJCS������0�ҡ�FWLXMC��=��סլ���ġ�SCJZMJ����ֵ���ۼӺ͡�
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "���ݻ�����Ϣ��������߼����"
    CheckmodelName = "�Զ���ű������->���ݻ�����Ϣ��������߼����"
    strDescription = "���������������ֵ�뻧��ͳ�����ֵ��һ��"
    
    '��ȡ���еķ����������� FWLXMCArr
    SqlStr = "Select DISTINCT FWLXMJHZXX.FWLXMC From FWLXMJHZXX Where FWLXMJHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,FWLXMCArr,FWLXMCCount
    
    If FWLXMCCount > 0 Then
        '��ȡ��Ӧ��ʵ����Ͻ������
        For CurrentCount = 0 To UBound(FWLXMCArr)
            If FWLXMCArr(CurrentCount) <> "" Then
                
                SqlStr = "Select Sum(FWLXMJHZXX.SCJZMJ) From FWLXMJHZXX Where FWLXMJHZXX.ID > 0 And FWLXMJHZXX.FWLXMC = " & "'" & FWLXMCArr(CurrentCount) & "' And " & "FWLXMJHZXX.KJWZ = '����' "
                GetSQLRecordAll SqlStr,SCJZMJArr,SearchCount
                
                If SearchCount > 0 Then
                    SCJZMJ = Transform(SCJZMJArr(0))
                Else
                    SCJZMJ = 0
                End If
                
                SqlStr = "Select Sum(H.SCJZMJ) From H Where H.ID > 0 And H.FWLXMC = " & "'" & FWLXMCArr(CurrentCount) & "' And " & "H.KJWZ = '����' And H.SJCS > 0 "
                GetSQLRecordAll SqlStr,HSCJZMJArr,SearchCount
                
                If SearchCount > 0 Then
                    HSCJZMJ = Transform(HSCJZMJArr(0))
                Else
                    HSCJZMJ = 0
                End If
                
                If SCJZMJ <> HSCJZMJ Then
                    SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
                End If
            End If
        Next 'CurrentCount
    End If
    
End Function' HDXCheck

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