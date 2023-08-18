
'=============================================��鼯����==============================================

'�������Ŀ����
Dim strGroupName
strGroupName = "���߼��"

'��鼯������
Dim strCheckName
strCheckName = "�ܾ����"

'���ģ������
Dim CheckmodelName
CheckmodelName = "�Զ���ű������->�ܾ����"

'�������
Dim strDescription
strDescription = "�ܾ�����"

'===========================================�������========================================================

'�����
Sub OnClick()
    
    ClearCheckRecord '���ԭ���ļ���¼
    
    GetErrorLines LineIds '������ߵ�ID
    
    AddRecords LineIds '��Ӽ���¼
    
End Sub' OnClick

'���ش������ID
Function GetErrorLines(ByRef LineIds)
    
    EorrorCount = 0
    
    ReDim LineIds(EorrorCount)
    
    
    SqlStr = "Select ���¹��������Ա�.ID,GXQDMS,GXZDMS,GJ From ���¹��������Ա� Inner Join GeoLineTB on ���¹��������Ա�.ID = GeoLineTB.ID Where (GeoLineTB.Mark Mod 2)<>0"
    
    GetSQLRecordAll SqlStr,LineArr,LineCount
    
    For i = 0 To LineCount - 1
        LineAttrArr = Split(LineArr(i),",", - 1,1) '0=ID,1=GXQDMS,2=GXZDMS,3=GJ
        If InStr(LineAttrArr(3),"*") <> 0 Then
            GJArr = Split(LineAttrArr(3),"*", - 1,1) '0=Length 1=Width
            Length = Transform(GJArr(0)) / 1000
            Width = Transform(GJArr(1)) / 1000
            If Length > Width Then
                CompareGJ = Length
            Else
                CompareGJ = Width
            End If
        ElseIf LineAttrArr(3) <> "" Then
            CompareGJ = Transform(LineAttrArr(3)) / 1000
        End If
        GXQDMS = Transform(LineAttrArr(1))
        GXZDMS = Transform(LineAttrArr(2))
        If CompareGJ > GXQDMS Or CompareGJ > GXZDMS Then
            LineIds(EorrorCount) = LineAttrArr(0)
            EorrorCount = EorrorCount + 1
            ReDim Preserve LineIds(EorrorCount)
        End If
    Next 'i
End Function' GetErrorLines

'��Ӽ���¼
Function AddRecords(ByVal LineIds())
    For i = 0 To UBound(LineIds) - 1
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(LineIds(i),"SSObj_X"),SSProcess.GetObjectAttr(LineIds(i),"SSObj_Y"),0,1,LineIds(i),""
    Next 'i
    SSProcess.ShowCheckOutput
    SSProcess.SaveCheckRecord
End Function' AddRecords

'================================================�����ຯ��===========================================

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

'��ռ�鼯
Function ClearCheckRecord()
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
End Function' ClearCheckRecord

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