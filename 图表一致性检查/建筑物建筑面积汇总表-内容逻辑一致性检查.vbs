
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
    
    DDFWCheck
    
    ShowCheckRecord
    
End Sub' OnClick

'===================================================��麯��=======================================================

'������;�빦������;�������ֵ�Ƿ�һ�£���������
Function DDFWCheck()
    
    ' 1��ʵ��¥�����������Ϣ��SCLDMJHZXX�����С�LD��=��1#���ҡ�YT��=��סլ���ġ�JZMJ��
    ' 2���滮��������GHGNQ�����еġ�SSZRZ��=��1#���ҡ�YT��=��סլ���ġ�JZMJ����ֵ���ۼ�ֵ��
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "������;�빦������;�������ֵһ���Լ�飨��������"
    CheckmodelName = "�Զ���ű������->������;�빦������;�������ֵһ���Լ�飨��������"
    strDescription = "������;�빦������;�������ֵ��һ��"
    
    '���е�¥��
    SqlStr = "Select DISTINCT SCLDMJHZXX.LD From SCLDMJHZXX Where SCLDMJHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,LDArr,LDCount
    
    If LDCount > 0 Then
        For i = 0 To LDCount - 1
            
            SqlStr = "Select DISTINCT SCLDMJHZXX.YT From SCLDMJHZXX Where SCLDMJHZXX.ID > 0 And SCLDMJHZXX.LD = '" & LDArr(i) & "'"
            GetSQLRecordAll SqlStr,YTArr,YTCount
            
            For j = 0 To YTCount - 1
                
                SqlStr = "Select Sum(JG_�滮���������Ա�.JZMJ) From JG_�滮���������Ա� Inner Join GeoAreaTB On JG_�滮���������Ա�.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And JG_�滮���������Ա�.SSZRZ = '" & LDArr(i) & "' And JG_�滮���������Ա�.YT = '" & YTArr(j) & "'"
                GetSQLRecordAll SqlStr,SumAreaArr,SumCount
                SumArea = Transform(SumAreaArr(0))
                
                SqlStr = "Select SCLDMJHZXX.JZMJ From SCLDMJHZXX Where SCLDMJHZXX.LD = '" & LDArr(i) & "' And SCLDMJHZXX.YT = '" & YTArr(j) & "'"
                GetSQLRecordAll SqlStr,JZMJArr,SearchCount
                JZMJ = Transform(JZMJArr(0))
                
                If JZMJ - SumArea <> 0 Then
                    SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
                End If
                
            Next 'j
        Next 'i   
    End If
    
    
End Function' DDFWCheck

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