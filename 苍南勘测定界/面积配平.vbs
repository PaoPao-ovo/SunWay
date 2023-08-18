
'===========================================�������========================================================

'�����
Sub OnClick()
    
    GetDKId DkIdArr '���еؿ��ID
    
    TrimIndex = "" '�޸ĵ�ͼ��ID
    
    TrimBool = False
    For i = 0 To UBound(DkIdArr)
        GetDiffInfo DkIdArr(i),DKMJ,TbTotalArea
        If TbTotalArea <> 0 And TbTotalArea <> DKMJ Then
            TrimBool = True
            DiffArea = Round(Transform(DKMJ) - TbTotalArea,4)
            AreaTrim DkIdArr(i),DiffArea,MathIndex
            If TrimIndex = "" Then
                TrimIndex = MathIndex
            Else
                TrimIndex = TrimIndex & "," & MathIndex
            End If
        End If
    Next 'i
    
    Ending TrimBool,TrimIndex
    
End Sub' OnClick

'�������еؿ��ID
Function GetDKId(ByRef DkIdArr())
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "504"
    SSProcess.SelectFilter
    DKCount = SSProcess.GetSelgeoCount()
    ReDim DkIdArr(DKCount - 1)
    For i = 0 To DKCount - 1
        DkIdArr(i) = SSProcess.GetSelGeoValue(i,"SSObj_ID")
    Next 'i
End Function' GetDKId

'��ȡ�������Ϣ
Function GetDiffInfo(ByVal DKId,ByRef DKMJ,ByRef TotalArea)
    
    DKH = SSProcess.GetObjectAttr(DKId,"[DKH]")
    DKMJ = Transform(SSProcess.GetObjectAttr(DKId,"[DKMJ]"))
    
    SqlStr = "Select SUM(����ͼ�����Ա�.TBMJ) From ����ͼ�����Ա� INNER JOIN GeoAreaTB ON ����ͼ�����Ա�.ID = GeoAreaTB.ID WHERE ([GeoAreaTB].[Mark] Mod 2)<>0 And ����ͼ�����Ա�.DKH= " & DKH
    GetSQLRecordAll SqlStr,TbMjArr,TbCount
    
    If TbCount > 0 Then
        TotalArea = Transform(TbMjArr(0))
    Else
        TotalArea = 0
    End If
    
End Function' GetDiffInfo

'�����ƽ
Function AreaTrim(ByVal DKId,ByVal DiffArea,ByRef MathIndex)
    ' IsBool = �ж��Ƿ���Ҫƽ��
    If DiffArea > 0 Then
        Attr = True
        IsBool = True
        SearchNum = 10000 * DiffArea 'SearchNum=��Ҫ�޸�ͼ�����������
    ElseIf DiffArea < 0 Then
        Attr = False
        SearchNum = Abs(10000 * DiffArea)
        IsBool = True
    Else
        IsBool = False
    End If
    If IsBool Then
        TrimTb Attr,DKId,SearchNum,MathIndex
    End If
End Function' AreaTrim

'�޸����ͼ��
Function TrimTb(ByVal Attr,ByVal DKId,ByVal SearchNum,ByRef MathIndex)
    
    TrimIndex = "" '�޸������ͼ��ID
    
    If Attr Then
        
        '������λ�� 1=С�����һλ
        SearchByte = 1
        
        '�����С�������ַ���
        FractionalPart = ""
        
        SearchCount = 0
        
        'С������SearchByteλ����4��
        Dim TempArr()
        ReDim TempArr(SearchCount)
        
        '1����ȡͼ�����С�����֣���ƴ���ַ�������
        DKH = SSProcess.GetObjectAttr(DKId,"[DKH]")
        SqlStr = "Select ����ͼ�����Ա�.ID From ����ͼ�����Ա� INNER JOIN GeoAreaTB ON ����ͼ�����Ա�.ID = GeoAreaTB.ID WHERE ([GeoAreaTB].[Mark] Mod 2)<>0 And ����ͼ�����Ա�.DKH= " & DKH
        GetSQLRecordAll SqlStr,TbArr,TbCount
        ReDim TbmjArr(TbCount - 1)
        For i = 0 To TbCount - 1
            Tbmj = SSProcess.GetObjectAttr(TbArr(i),"SSObj_Area")
            NumberArr = Split(Tbmj,".", - 1,1) 'NumberArr(1)С������
            If FractionalPart = "" Then
                FractionalPart = NumberArr(1)
            Else
                FractionalPart = FractionalPart & "," & NumberArr(1)
            End If
        Next 'i
        DecimalArr = Split(FractionalPart,",", - 1,1)
        
        '2��ѭ��������������1λС��4�Ĳ�������Ŀ������ѭ���жϵڶ�λ������ı�����TempArr��(��С��Ҫ��ȥ1)
        Do While SearchCount < SearchNum
            For i = 0 To UBound(DecimalArr)
                ByteNum = Transform(Mid(DecimalArr(i),SearchByte,1))
                If ByteNum < 4 Then
                    TempArr(SearchCount) = DecimalArr(i)
                    SearchCount = SearchCount + 1
                    ReDim Preserve TempArr(SearchCount)
                End If
            Next 'i
            
            '3���ض������鲢����֮ǰ������
            ReDim Preserve TempArr(UBound(TempArr) - 1)
            
            '4���������鲢������������飨�Ӵ�С��
            SortNum TempArr,ResultArr,SearchByte
            
            SearchByte = SearchByte + 1
        Loop
        
        '5���ж������λ��,�������ַ�����
        MathIndex = "" 'MathIndex = �����±��ַ���
        For i = 0 To UBound(ResultArr)
            For j = 0 To UBound(DecimalArr)
                If ResultArr(i) = DecimalArr(j) Then
                    If MathIndex = "" Then
                        MathIndex = j
                    Else
                        MathIndex = MathIndex & "," & j
                    End If
                End If
            Next 'j
        Next 'i
        
        '6�����������±��ҵ�ͼ�߲��������
        IndexArr = Split(MathIndex,",", - 1,1)
        For i = 0 To SearchNum - 1
            Tbmj = Transform(SSProcess.GetObjectAttr(TbArr(IndexArr(i)),"[TBMJ]"))
            SSProcess.SetObjectAttr TbArr(IndexArr(i)),"[TBMJ]",Tbmj + 0.0001
        Next 'i
        
        '7�������޸ĵ�ͼ�ߵ�ID
        MathIndex = "" '�ƿ�
        For i = 0 To SearchNum - 1
            If MathIndex = "" Then
                MathIndex = TbArr(IndexArr(i))
            Else
                MathIndex = MathIndex & "," & TbArr(IndexArr(i))
            End If
        Next 'i
        
    ElseIf Not Attr Then
        
        '������λ�� 1=С�����һλ
        SearchByte = 1
        
        '�����С�������ַ���
        FractionalPart = ""
        
        SearchCount = 0
        
        'С������SearchByteλ����4��
        Dim TempsArr()
        ReDim TempsArr(SearchCount)
        
        '1����ȡͼ�����С�����֣���ƴ���ַ�������
        DKH = SSProcess.GetObjectAttr(DKId,"[DKH]")
        SqlStr = "Select ����ͼ�����Ա�.ID From ����ͼ�����Ա� INNER JOIN GeoAreaTB ON ����ͼ�����Ա�.ID = GeoAreaTB.ID WHERE ([GeoAreaTB].[Mark] Mod 2)<>0 And ����ͼ�����Ա�.DKH= " & DKH
        GetSQLRecordAll SqlStr,TbArr,TbCount
        ReDim TempsArr(TbCount - 1)
        For i = 0 To TbCount - 1
            Tbmj = SSProcess.GetObjectAttr(TbArr(i),"SSObj_Area")
            NumberArr = Split(Tbmj,".", - 1,1) 'NumberArr(1)С������
            If FractionalPart = "" Then
                FractionalPart = NumberArr(1)
            Else
                FractionalPart = FractionalPart & "," & NumberArr(1)
            End If
        Next 'i
        DecimalArr = Split(FractionalPart,",", - 1,1)
        
        '2��ѭ��������������1λ����4�Ĳ�������Ŀ������ѭ���жϵ�2λ������ı�����TempsArr��(��С��Ҫ��ȥ1)
        Do While SearchCount < SearchNum
            For i = 0 To UBound(DecimalArr)
                ByteNum = Transform(Mid(DecimalArr(i),SearchByte,1))
                If ByteNum > 4 Then
                    TempsArr(SearchCount) = DecimalArr(i)
                    SearchCount = SearchCount + 1
                    ReDim Preserve TempsArr(SearchCount)
                End If
            Next 'i
            
            '3���ض������鲢����֮ǰ������
            ReDim Preserve TempsArr(UBound(TempsArr) - 1)
            
            '4���������鲢������������飨�Ӵ�С��
            SortNum TempsArr,ResultArr,SearchByte
            
            SearchByte = SearchByte + 1
        Loop
        
        '5���ж������λ��,�������ַ�����
        MathIndex = "" 'MathIndex = �����±��ַ���
        For i = 0 To UBound(ResultArr)
            For j = 0 To UBound(DecimalArr)
                If ResultArr(i) = DecimalArr(j) Then
                    If MathIndex = "" Then
                        MathIndex = j
                    Else
                        MathIndex = MathIndex & "," & j
                    End If
                End If
            Next 'j
        Next 'i
        
        '6�����������±��ҵ�ͼ�߲���ȥ���
        IndexArr = Split(MathIndex,",", - 1,1)
        For i = SearchNum - 1 To 0
            Tbmj = Transform(SSProcess.GetObjectAttr(TbArr(IndexArr(i)),"[TBMJ]"))
            SSProcess.SetObjectAttr TbArr(IndexArr(i)),"[TBMJ]",Tbmj - 0.0001
        Next 'i
        
        '7�������޸ĵ�ͼ�ߵ�ID
        MathIndex = "" '�ƿ�
        For i = 0 To SearchNum - 1
            If MathIndex = "" Then
                MathIndex = TbArr(IndexArr(i))
            Else
                MathIndex = MathIndex & "," & TbArr(IndexArr(i))
            End If
        Next 'i
    End If
End Function' TrimTb

'=================================================�����ຯ��=====================================================

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

'ѡ������Ӵ�С
Function SortNum(ByVal InputArr(),ByRef SortArr(),ByVal SearchByte)
    Size = UBound(InputArr)
    ReDim SortArr(Size)
    For i = 0 To Size - 1
        Max = i
        For j = i + 1 To Size
            If Transform(Mid(InputArr(j),SearchByte,1)) > Transform(Mid(InputArr(i),SearchByte,1)) Then
                Max = j
            End If
        Next 'j
        If Max <> i Then
            Temp = InputArr(i)
            InputArr(i) = InputArr(Max)
            InputArr(Max) = Temp
        End If
    Next 'i
    For i = 0 To Size
        SortArr(i) = InputArr(i)
    Next 'i
End Function' SortNum

'�����ʾ
Function Ending(ByVal TrimBool,ByVal TrimIndex)
    If TrimBool Then
        MsgBox  "ͼ��IDΪ��" & TrimIndex & " ��ͼ������޸ģ������ƽ"
    Else
        MsgBox "������ƽ"
    End If
End Function' Ending