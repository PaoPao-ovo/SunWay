
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
    
    RecordExist
    
    LHYSCheck
    
    ShowCheckRecord
    
End Sub' OnClick

'===================================================��麯��=======================================================


Function LHYSCheck()
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "111"
    CheckmodelName = "�Զ���ű������->111"
    strDescription = "222"
    
    SqlStr = "Select DISTINCT GH_�̻�Ҫ�����Ա�.ID_LDK From ����GH_�̻�Ҫ�����Ա�ͼ�����Ա� INNER JOIN GeoAreaTB ON GH_�̻�Ҫ�����Ա�.ID = GeoAreaTB.ID WHERE([GeoAreaTB].[Mark] Mod 2)<>0"
    
    GetSQLRecordAll SqlStr,ID_LDKArr,ID_LDKCount
    
    
    If ID_LDKCount > 0 Then
        For i = 0 To ID_LDKCount - 1
            SqlStr = "Select LHHF.�̵ؿ�ID From LHHF Where LHHF.�̵ؿ�ID = '" & ID_LDKArr(i) & "'"
            GetSQLRecordAll SqlStr,LHHFArr,LHHFCount
            If LHHFCount < 0 Then
                SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,0,0,""
            End If
        Next 'i
    End If
    
End Function' LHYSCheck

Function RecordExist()
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "111"
    CheckmodelName = "�Զ���ű������->111"
    strDescription = "�̻����ֱ�LHHF���ġ��̵ؿ�ID�������ڼ�¼"
    
    SqlStr = "Select DISTINCT LHHF.�̵ؿ�ID From LHHF"
    GetSQLRecordAll SqlStr,LHHFArr,LHHFCount
    
    If LHHFCount < 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,0,0,""
    End If

End Function' RecordExist

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