
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

    CheckFilds = "X,Y,GC" '����ֶ�

    AccuracyCheck "KZDZBCGXXB",CheckFilds,3

    ShowCheckRecord
    
End Sub' OnClick

'=====================================================��麯��======================================================

'С��λ�����ȼ��
Function AccuracyCheck(ByVal TableName,ByVal FildsStr,ByVal CheckBits) 'TableName = ����,FildsStr = ��ѯ���ֶ��ַ���,CheckBits = ���λ��
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "���Ƶ������С��λ�淶�Լ��"
    CheckmodelName = "�Զ���ű������->���Ƶ������С��λ�淶�Լ��"
    
    '��ѯ�ֶ�ֵ
    SqlStr = "Select " & TableName & "." & "DH," & FildsStr & " From " & TableName
    GetSQLRecordAll SqlStr,ValArr,SearchCount  'ValArr = [(ֵ1,ֵ2,ֵ3....)(ֵ1,ֵ2,ֵ3....)]
    
    '�ֶ���������
    FildsNameArr = Split(FildsStr,",", - 1,1)
    
    '�����ֶ�ֵ
    For i = 0 To SearchCount - 1
        CurrentValArr = Split(ValArr(i),",", - 1,1)
        For j = 1 To UBound(CurrentValArr)
            DecimalJudgment Transform(CurrentValArr(j)),CheckBits,ErrorBool
            If ErrorBool Then
                strDescription = "���Ƶ�����ɹ���" & TableName & "����DHΪ��" & CurrentValArr(0) & "����" & "��" & FildsNameArr(j - 1) & "���ֶ�" & "С��λ��������"
                SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,0,0,""
            End If
        Next 'j
    Next 'i
End Function' AccuracyCheck

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

'С��λ���ж�
Function DecimalJudgment(ByVal Num,ByVal CheckBits,ByRef ErrorBool) 'Num = �����,CheckBits = ���λ��,ErrorBool = �Ƿ����,���󷵻�True
    
    ErrorBool = False
    
    DecimalPointPoi = InStr(1,Num,".",1)
    
    If Num = "" Then
        ErrorBool = False
    ElseIf Num <> "" And DecimalPointPoi = 0 Then
        ErrorBool = False
    ElseIf Num <> "" And DecimalPointPoi > 0 Then
        DecimalLen = Len(Num) - DecimalPointPoi
        If DecimalLen < CheckBits Then
            ErrorBool = False
        Else
            ErrorBool = True
        End If
    Else
        ErrorBool = True
    End If
End Function' DecimalJudgment

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