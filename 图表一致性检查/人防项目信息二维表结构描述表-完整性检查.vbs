
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
    
    FildsEmptyCheck "RFPROJECTINFO","�����ṹ,סլ����,���Ͻ������(�O),����סլ�������(�O),���������������(�O),���ϲ���,����ƽʱ����,���½������(�O),���²���,������ͨ���,���վ������������,��ǽ�������(С��10��ʱ��д),��ƺ�߲� (�������߳�����ʱ��д),������,�����,�˷��������,�ڱ������,������Ԫ����","��Ϣ��"
    
    ShowCheckRecord
    
End Sub' OnClick

'===================================================��麯��=======================================================

'���ֶο�ֵ���
Function FildsEmptyCheck(ByVal TableName,ByVal FildsStr,ByVal TableType)
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = TableName & "��ֵ���"
    CheckmodelName = "�Զ���ű������->" & strCheckName
    
    If TableType = "��Ϣ��" Then
        
        FildsArr = Split(FildsStr,",", - 1,1)
        For i = 0 To UBound(FildsArr)
            
            SqlStr = "Select " & TableName & ".ID " & "From " & TableName & " Where " & TableName & ".Key = " & "'" & FildsArr(i) & "'"
            GetSQLRecordAll SqlStr,KeyArr,KeyCount
            If KeyCount < 1 Then
                strDescription = "��" & TableName & "���ġ�Key��Ϊ��" & FildsArr(i) & "��ȱʧ"
                SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,0,0,""
            Else
                SqlStr = "Select " & TableName & ".Value " & "From " & TableName & " Where Key = " & "'" & FildsArr(i) & "'"  & " And (" & "Value" & " = '' " & " Or " & "Value" & " = '*' Or " & "Value" & " IS NULL)"
                
                GetSQLRecordAll SqlStr,ValueArr,ValueCount
                
                If ValueCount > 0  Then
                    strDescription = "��" & TableName & "���ġ�Value��Ϊ" & "��" & FildsArr(i) & "��" & "Ϊ��ֵ"
                    SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,0,0,""
                End If
            End If
        Next 'i
    End If
    
End Function' FildsEmptyCheck

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