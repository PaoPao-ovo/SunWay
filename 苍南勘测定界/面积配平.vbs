
'===========================================功能入口========================================================

'总入口
Sub OnClick()
    
    GetDKId DkIdArr '所有地块的ID
    
    TrimIndex = "" '修改的图斑ID
    
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

'返回所有地块的ID
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

'获取面积差信息
Function GetDiffInfo(ByVal DKId,ByRef DKMJ,ByRef TotalArea)
    
    DKH = SSProcess.GetObjectAttr(DKId,"[DKH]")
    DKMJ = Transform(SSProcess.GetObjectAttr(DKId,"[DKMJ]"))
    
    SqlStr = "Select SUM(地类图斑属性表.TBMJ) From 地类图斑属性表 INNER JOIN GeoAreaTB ON 地类图斑属性表.ID = GeoAreaTB.ID WHERE ([GeoAreaTB].[Mark] Mod 2)<>0 And 地类图斑属性表.DKH= " & DKH
    GetSQLRecordAll SqlStr,TbMjArr,TbCount
    
    If TbCount > 0 Then
        TotalArea = Transform(TbMjArr(0))
    Else
        TotalArea = 0
    End If
    
End Function' GetDiffInfo

'面积配平
Function AreaTrim(ByVal DKId,ByVal DiffArea,ByRef MathIndex)
    ' IsBool = 判断是否需要平差
    If DiffArea > 0 Then
        Attr = True
        IsBool = True
        SearchNum = 10000 * DiffArea 'SearchNum=需要修改图斑面积的数量
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

'修改面积图斑
Function TrimTb(ByVal Attr,ByVal DKId,ByVal SearchNum,ByRef MathIndex)
    
    TrimIndex = "" '修改面积的图斑ID
    
    If Attr Then
        
        '搜索的位数 1=小数后第一位
        SearchByte = 1
        
        '面积的小数部分字符串
        FractionalPart = ""
        
        SearchCount = 0
        
        '小数点后第SearchByte位大于4的
        Dim TempArr()
        ReDim TempArr(SearchCount)
        
        '1、获取图斑面积小数部分，并拼接字符串保存
        DKH = SSProcess.GetObjectAttr(DKId,"[DKH]")
        SqlStr = "Select 地类图斑属性表.ID From 地类图斑属性表 INNER JOIN GeoAreaTB ON 地类图斑属性表.ID = GeoAreaTB.ID WHERE ([GeoAreaTB].[Mark] Mod 2)<>0 And 地类图斑属性表.DKH= " & DKH
        GetSQLRecordAll SqlStr,TbArr,TbCount
        ReDim TbmjArr(TbCount - 1)
        For i = 0 To TbCount - 1
            Tbmj = SSProcess.GetObjectAttr(TbArr(i),"SSObj_Area")
            NumberArr = Split(Tbmj,".", - 1,1) 'NumberArr(1)小数部分
            If FractionalPart = "" Then
                FractionalPart = NumberArr(1)
            Else
                FractionalPart = FractionalPart & "," & NumberArr(1)
            End If
        Next 'i
        DecimalArr = Split(FractionalPart,",", - 1,1)
        
        '2、循环遍历，搜索第1位小于4的并排序，数目不足则循环判断第二位，满足的保存在TempArr中(大小需要减去1)
        Do While SearchCount < SearchNum
            For i = 0 To UBound(DecimalArr)
                ByteNum = Transform(Mid(DecimalArr(i),SearchByte,1))
                If ByteNum < 4 Then
                    TempArr(SearchCount) = DecimalArr(i)
                    SearchCount = SearchCount + 1
                    ReDim Preserve TempArr(SearchCount)
                End If
            Next 'i
            
            '3、重定义数组并保存之前的数据
            ReDim Preserve TempArr(UBound(TempArr) - 1)
            
            '4、排序数组并输出排序后的数组（从大到小）
            SortNum TempArr,ResultArr,SearchByte
            
            SearchByte = SearchByte + 1
        Loop
        
        '5、判断数组的位置,保存在字符串中
        MathIndex = "" 'MathIndex = 数组下标字符串
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
        
        '6、根据数组下标找到图斑并加上面积
        IndexArr = Split(MathIndex,",", - 1,1)
        For i = 0 To SearchNum - 1
            Tbmj = Transform(SSProcess.GetObjectAttr(TbArr(IndexArr(i)),"[TBMJ]"))
            SSProcess.SetObjectAttr TbArr(IndexArr(i)),"[TBMJ]",Tbmj + 0.0001
        Next 'i
        
        '7、返回修改的图斑的ID
        MathIndex = "" '制空
        For i = 0 To SearchNum - 1
            If MathIndex = "" Then
                MathIndex = TbArr(IndexArr(i))
            Else
                MathIndex = MathIndex & "," & TbArr(IndexArr(i))
            End If
        Next 'i
        
    ElseIf Not Attr Then
        
        '搜索的位数 1=小数后第一位
        SearchByte = 1
        
        '面积的小数部分字符串
        FractionalPart = ""
        
        SearchCount = 0
        
        '小数点后第SearchByte位大于4的
        Dim TempsArr()
        ReDim TempsArr(SearchCount)
        
        '1、获取图斑面积小数部分，并拼接字符串保存
        DKH = SSProcess.GetObjectAttr(DKId,"[DKH]")
        SqlStr = "Select 地类图斑属性表.ID From 地类图斑属性表 INNER JOIN GeoAreaTB ON 地类图斑属性表.ID = GeoAreaTB.ID WHERE ([GeoAreaTB].[Mark] Mod 2)<>0 And 地类图斑属性表.DKH= " & DKH
        GetSQLRecordAll SqlStr,TbArr,TbCount
        ReDim TempsArr(TbCount - 1)
        For i = 0 To TbCount - 1
            Tbmj = SSProcess.GetObjectAttr(TbArr(i),"SSObj_Area")
            NumberArr = Split(Tbmj,".", - 1,1) 'NumberArr(1)小数部分
            If FractionalPart = "" Then
                FractionalPart = NumberArr(1)
            Else
                FractionalPart = FractionalPart & "," & NumberArr(1)
            End If
        Next 'i
        DecimalArr = Split(FractionalPart,",", - 1,1)
        
        '2、循环遍历，搜索第1位大于4的并排序，数目不足则循环判断第2位，满足的保存在TempsArr中(大小需要减去1)
        Do While SearchCount < SearchNum
            For i = 0 To UBound(DecimalArr)
                ByteNum = Transform(Mid(DecimalArr(i),SearchByte,1))
                If ByteNum > 4 Then
                    TempsArr(SearchCount) = DecimalArr(i)
                    SearchCount = SearchCount + 1
                    ReDim Preserve TempsArr(SearchCount)
                End If
            Next 'i
            
            '3、重定义数组并保存之前的数据
            ReDim Preserve TempsArr(UBound(TempsArr) - 1)
            
            '4、排序数组并输出排序后的数组（从大到小）
            SortNum TempsArr,ResultArr,SearchByte
            
            SearchByte = SearchByte + 1
        Loop
        
        '5、判断数组的位置,保存在字符串中
        MathIndex = "" 'MathIndex = 数组下标字符串
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
        
        '6、根据数组下标找到图斑并减去面积
        IndexArr = Split(MathIndex,",", - 1,1)
        For i = SearchNum - 1 To 0
            Tbmj = Transform(SSProcess.GetObjectAttr(TbArr(IndexArr(i)),"[TBMJ]"))
            SSProcess.SetObjectAttr TbArr(IndexArr(i)),"[TBMJ]",Tbmj - 0.0001
        Next 'i
        
        '7、返回修改的图斑的ID
        MathIndex = "" '制空
        For i = 0 To SearchNum - 1
            If MathIndex = "" Then
                MathIndex = TbArr(IndexArr(i))
            Else
                MathIndex = MathIndex & "," & TbArr(IndexArr(i))
            End If
        Next 'i
    End If
End Function' TrimTb

'=================================================工具类函数=====================================================

'获取所有记录
Function GetSQLRecordAll(ByVal StrSqlStatement, ByRef SQLRecord(), ByRef iRecordCount)
    ProJectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb ProJectName
    If StrSqlStatement = "" Then
        MsgBox "查询语句为空，操作停止！",48
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

'数据类型转换
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

'选择排序从大到小
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

'完成提示
Function Ending(ByVal TrimBool,ByVal TrimIndex)
    If TrimBool Then
        MsgBox  "图斑ID为：" & TrimIndex & " 的图斑面积修改，完成配平"
    Else
        MsgBox "无需配平"
    End If
End Function' Ending