
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

    FildsEmptyCheck "GXSCHZXX","GXLB,GXZL,CGCLCD,TCCD,ZCD","��Ϣ��"

    ShowCheckRecord

End Sub' OnClick

'===================================================��麯��=======================================================

'���ֶο�ֵ���
Function FildsEmptyCheck(ByVal TableName,ByVal FildsStr,ByVal TableType)

    MdbName = SSProcess.GetProjectFileName

    SSProcess.OpenAccessMdb MdbName

    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = TableName & "��ֵ���"
    CheckmodelName = "�Զ���ű������->" & strCheckName
    
    If TableType = "��Ϣ��" Then
        
        FildsArr = Split(FildsStr,",", - 1,1)
        For i = 0 To UBound(FildsArr)
        
            '�ֶ�����,��������,�ֶδ�С,�ֶ�����,�ֶ����,�Ƿ�����ֶ�,�Ƿ�����Ϊ��,����ȽϷ�ʽ,�ֶα���,Դ�ֶ���,Դ����,�ֶι���,�ֶι�������,ȱʡֵ
            '��������Ϊ 7(Double),6(Float)
            '�ַ���Ϊ 10(Char & String)
            
            SSProcess.GetAccessFieldInfo1 MdbName,TableName,FildsArr(i),FieldsInfo
            
            FieldsInfoArr = Split(FieldsInfo,",", - 1,1)
            If FieldsInfoArr(1) = "10" Then

                SqlStr = "Select " & TableName & "." & FildsArr(i) & " From " & TableName & " Where " & FildsArr(i) & " = '' Or " & FildsArr(i) & " = '*' Or " & FildsArr(i) & " IS NULL "
                GetSQLRecordAll SqlStr,StringArr,StringEmptyCount

                If StringEmptyCount > 0 Then
                    strDescription = "��" & TableName & "��" & "��" & "��" & FildsArr(i) & "��" & "���ڿ�ֵ"
                    SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,0,0,""
                End If

            ElseIf FieldsInfoArr(1) = "7" Or FieldsInfoArr(1) = "6" Then

                SqlStr = "Select " & TableName & "." & FildsArr(i) & " From " & TableName & " Where " & FildsArr(i) & " IS NULL Or " & FildsArr(i) & " = '' "
                GetSQLRecordAll SqlStr,NumArr,NumEmptyCount

                If NumEmptyCount > 0 Then
                    strDescription = "��" & TableName & "��" & "��" & "��" & FildsArr(i) & "��" & "���ڿ�ֵ"
                    SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,0,0,""
                End If

            End If
        Next 'i
    End If

    SSProcess.CloseAccessMdb MdbName

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