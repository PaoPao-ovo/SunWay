
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
    
    FHDYGSCheck

    RFJZMJCheck
    
    YBQCheck
    
    ShowCheckRecord

End Sub' OnClick

'===================================================��麯��=======================================================

'������Ԫ�����������Ԫ��Χ�߸�����һ��
Function FHDYGSCheck()
    
    ' 1���˷���Ŀ��Ϣ��RFPROJECTINFO���еġ�FHDYGS����ֵ
    ' 2:�˷�������Ԫ��Χ�ߣ�RFFHDYFW��Ҫ�ظ�����
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "������Ԫ�����������Ԫ��Χ�߸���һ���Լ��"
    CheckmodelName = "�Զ���ű������->������Ԫ�����������Ԫ��Χ�߸���һ���Լ��"
    strDescription = "������Ԫ�����������Ԫ��Χ�߸�����һ��"

    '��ȡ������Ԫ���� FHDYGS
    SqlStr = "Select RFPROJECTINFO.Value From RFPROJECTINFO Where RFPROJECTINFO.ID > 0 And RFPROJECTINFO.Key = '������Ԫ����' "
    GetSQLRecordAll SqlStr,FHDYGSArr,FHDYGSCount

    If FHDYGSCount > 0 Then
        FHDYGS = Transform(FHDYGSArr(0))
    Else
        FHDYGS = 0
    End If
    
    '��ȡͼ�Ϸ�Χ�߸��� YSCount
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 9450013
    SSProcess.SelectFilter
    YSCount = SSProcess.GetSelGeoCount()
    
    If YSCount - FHDYGS <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If

End Function' FHDYGSCheck

'�˷�����������˷��������������ֵ�Ƿ�һ��
Function RFJZMJCheck()
    
    ' 1���˷���Ŀ��Ϣ��RFPROJECTINFO���еġ�RFJZMJ����ֵ
    ' 2:�˷���������RFGNQ���еġ�JZMJ�������л���ֵ
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "�˷�����������˷��������������ֵһ���Լ��"
    CheckmodelName = "�Զ���ű������->�˷�����������˷��������������ֵһ���Լ��"
    strDescription = "�˷�����������˷��������������ֵ��һ��"

    '�˷�������� RFJZMJ
    SqlStr = "Select RFPROJECTINFO.Value From RFPROJECTINFO Where RFPROJECTINFO.ID > 0 And RFPROJECTINFO.Key = '�˷��������' "
    GetSQLRecordAll SqlStr,RFJZMJArr,RFJZCount

    If RFJZCount > 0 Then
        RFJZMJ = Transform(RFJZMJArr(0))
    Else
        RFJZMJ = 0
    End If
    
    '�˷��������������ֵ SumArea
    SqlStr = "Select Sum(RF_�˷����������Ա�.JZMJ) From RF_�˷����������Ա� Inner Join GeoAreaTB On RF_�˷����������Ա�.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0"
    GetSQLRecordAll SqlStr,SumAreaArr,SumCount

    If SumCount > 0 Then
        SumArea = Transform(SumAreaArr(0))
    Else
        SumArea = 0
    End If
    
    If RFJZMJ - SumArea <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If

End Function' RFJZMJCheck

'�ڱ���������˷����������ڱ������������ֵ�Ƿ�һ��
Function YBQCheck()
    
    ' 1���˷���Ŀ��Ϣ��RFPROJECTINFO���еġ�YBQMJ����ֵ
    ' 2:�˷���������RFGNQ���еġ�YSDM��=��600301���ġ�JZMJ�������л���ֵ
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "�ڱ���������˷����������ڱ������������ֵһ���Լ��"
    CheckmodelName = "�Զ���ű������->�ڱ���������˷����������ڱ������������ֵһ���Լ��"
    strDescription = "�ڱ���������˷����������ڱ������������ֵ��һ��"

    '�ڱ������ YBQMJ
    SqlStr = "Select RFPROJECTINFO.Value From RFPROJECTINFO Where RFPROJECTINFO.ID > 0 And RFPROJECTINFO.Key = '�ڱ������' "
    GetSQLRecordAll SqlStr,YBQMJArr,YBQCount

    If YBQCount > 0 Then
        YBQMJ = Transform(YBQMJArr(0))
    Else
        YBQMJ = 0
    End If
    
    '�˷����������ڱ������������ֵ SumArea
    SqlStr = "Select Sum(RF_�˷����������Ա�.JZMJ) From RF_�˷����������Ա� Inner Join GeoAreaTB On RF_�˷����������Ա�.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And RF_�˷����������Ա�.YSDM = '" & "600301'"
    GetSQLRecordAll SqlStr,SumAreaArr,SumCount

    If SumCount > 0 Then
        SumArea = Transform(SumAreaArr(0))
    Else
        SumArea = 0
    End If
    
    If YBQMJ - SumArea <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
  
End Function' YBQCheck

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