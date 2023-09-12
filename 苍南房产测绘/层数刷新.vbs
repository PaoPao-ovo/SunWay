' /*
'  * @Description: 请填写简介
'  * @Author: LHY
'  * @Date: 2023-09-06 16:16:03
'  * @LastEditors: LHY
'  * @LastEditTime: 2023-09-12 16:44:17
'  */

'入口
Sub Onclick()
    
    GetCH
    
    SumArea
    
    SetJZ
    
End Sub' Onclick

'==================================================业务函数=====================================================

'层号刷新
Function GetCH()
    
    '1、获取【ZRZ_LP_信息表】的【ZRZGUID】（ID>0）
    '2、根据【ZRZGUID】关联【FC_楼层信息属性表】并获取其【CH】字段值（有多个）
    '3、获取【CH】字段的最大值和最小值
    '4、将最大值填写在【ZRZ_LP_信息表】的【地上】和【地下】层数
    '5、填写总层数【ZRZ_LP_信息表】的【ZCS】
    
    '1、获取所有的【自然幢GUID】
    
    '最大层数
    Dim MaxCH
    
    '最小层数
    Dim MinCH
    
    SqlStr = "Select ZRZ_LP_信息表.ZRZGUID From ZRZ_LP_信息表 Where ZRZ_LP_信息表.ID > 0 "
    GetSQLRecordAll SqlStr,GUIDArr,GUIDCount
    
    '2、关联【FC_楼层信息属性表】获取【CH】
    
    For i = 0 To GUIDCount - 1
        
        '层高初始化
        MaxCH = 0
        MinCH = 0
        
        SqlStr = "Select FC_楼层信息属性表.CH From FC_楼层信息属性表 Inner Join GeoAreaTB On FC_楼层信息属性表.ID = GeoAreaTB.ID Where (GeoAreaTB.Mark Mod 2) <> 0 And FC_楼层信息属性表.ZRZGUID = " & GUIDArr(i)
        GetSQLRecordAll SqlStr,CHArr,CHCount
        
        For j = 0 To CHCount - 1
            
            TempMax = GetMaxCH(CHArr(j))
            TempMin =  - GetMinCH(CHArr(j))
            
            If TempMax > MaxCH Then
                MaxCH = TempMax
            End If
            
            If TempMin < MinCH Then
                MinCH = TempMin
            End If
            
        Next 'j
        
        DelPoint MaxCH,InterMax
        DelPoint MinCH,InterMin
        
        MaxCH = Transform(InterMax)
        MinCH = Transform(InterMin)
        
        ProJectName = SSProcess.GetProjectFileName
        
        SSProcess.OpenAccessMdb ProJectName
        
        SqlStr = "Update ZRZ_LP_信息表 Set DSCS = " & MaxCH & " Where ZRZ_LP_信息表.ZRZGUID = " & GUIDArr(i)
        SSProcess.ExecuteAccessSql ProJectName,SqlStr
        
        SqlStr = "Update ZRZ_LP_信息表 Set DXCS = " & MinCH & " Where ZRZ_LP_信息表.ZRZGUID = " & GUIDArr(i)
        SSProcess.ExecuteAccessSql ProJectName,SqlStr
        
        SqlStr = "Update ZRZ_LP_信息表 Set ZCS = " & Abs(MinCH) + MaxCH & " Where ZRZ_LP_信息表.ZRZGUID = " & GUIDArr(i)
        SSProcess.ExecuteAccessSql ProJectName,SqlStr
        
        SSProcess.CloseAccessMdb ProJectName
        
    Next 'i
    
End Function' GetCH

'获取当前最大层数
Function GetMaxCH(ByVal CHStr)
    
    GetMaxCH = 0
    
    If CHStr = "" Then
        GetMaxCH = 0
    Else
        If InStr(CHStr,"+") <> 0 Then
            NumArr = Split(CHStr,"+", - 1,1)
            For i = 0 To UBound(NumArr)
                If i = 0 Then
                    GetMaxCH = Transform(NumArr(i))
                Else
                    If Transform(NumArr(i)) > GetMaxCH Then
                        GetMaxCH = Transform(NumArr(i))
                    End If
                End If
            Next 'i
        Else
            If InStr(CHStr,"-") = 0 Then
                GetMaxCH = Transform(CHStr)
            Else
                GetMaxCH = 0
            End If
        End If
    End If
    
End Function' GetMaxCH

'获取当前最小层数
Function GetMinCH(ByVal CHStr)
    
    GetMinCH = 0
    
    If CHStr = "" Then
        GetMinCH = 0
    Else
        If InStr(CHStr,"-") <> 0 Then
            NumArr = Split(CHStr,"-", - 1,1)
            For i = 0 To UBound(NumArr)
                If i = 0 Then
                    GetMinCH = Transform(NumArr(i))
                Else
                    If Transform(NumArr(i)) > GetMinCH Then
                        GetMinCH = Transform(NumArr(i))
                    End If
                End If
            Next 'i
        Else
            GetMinCH = 0
        End If
    End If
    
End Function' GetMinCH

'面积汇总
Function SumArea()
    
    ' 1、获取【ZRZ_LP_信息表】的【ZRZGUID】（ID>0）
    ' 2、根据【ZRZGUID】关联【FC_LPB_户信息表】，计算面积和【JZMJ】
    ' 3、根据【ZRZGUID】关联【FC_面积块信息属性表】找到【QSXZ】为【共有不分摊】的【JZMJ】之和
    ' 4、将上述面积之和求和填写【ZRZ_LP_信息表】的【JZMJ】
    
    '1、获取所有的【自然幢GUID】
    
    SqlStr = "Select ZRZ_LP_信息表.ZRZGUID From ZRZ_LP_信息表 Where ZRZ_LP_信息表.ID > 0 "
    GetSQLRecordAll SqlStr,GUIDArr,GUIDCount
    
    '2、获取户面积之和
    For i = 0 To GUIDCount - 1
        
        SqlStr = "Select Sum(FC_LPB_户信息表.JZMJ) From FC_LPB_户信息表 Where FC_LPB_户信息表.ZRZGUID = " & GUIDArr(i)
        GetSQLRecordAll SqlStr,HSumArr,HSumCount
        HArea = HSumArr(0)
        
        ' SqlStr = "Select Sum(FC_面积块信息属性表.JZMJ) From FC_面积块信息属性表 Inner Join GeoAreaTB On FC_面积块信息属性表.ID = GeoAreaTB.ID Where (GeoAreaTB.Mark Mod 2) <> 0 And FC_面积块信息属性表.ZRZGUID = " & GUIDArr(i) & " And FC_面积块信息属性表.QSXZ = '共有不分摊"
        ' GetSQLRecordAll SqlStr,TbSumArr,TbSumCount
        ' TbArea = TbSumArr(0)
        
        ProJectName = SSProcess.GetProjectFileName
        
        SSProcess.OpenAccessMdb ProJectName
        
        SqlStr = "Update ZRZ_LP_信息表 Set JZMJ = " & HArea + TbArea & " Where ZRZ_LP_信息表.ZRZGUID = " & GUIDArr(i)
        
        SSProcess.ExecuteAccessSql ProJectName,SqlStr
        
        SSProcess.CloseAccessMdb ProJectName
        
    Next 'i
    
End Function' SumArea

'设置建筑形式
Function SetJZ()
    
    ' 低层1-3；多层4 - 6；高层7层以上
    ' 有地下室加地下室
    
    '结构名称
    Dim JZStr
    
    '低层出现个数
    Dim DC_Count
    
    '多层出现个数
    Dim MC_Count
    
    '高层出现个数
    Dim GC_Count
    
    '地下出现个数
    Dim DX_Count
    
    '数据初始化
    JZStr = ""
    
    DC_Count = 0
    
    MC_Count = 0
    
    GC_Count = 0
    
    DX_Count = 0
    
    '获取所有的自然幢GUID
    SqlStr = "Select ZRZ_LP_信息表.ZRZGUID From ZRZ_LP_信息表 Where ZRZ_LP_信息表.ID > 0"
    GetSQLRecordAll SqlStr,GUIDArr,GUIDCount
    
    For i = 0 To GUIDCount - 1
        
        '制空
        JZStr = ""
        
        SqlStr = "Select ZRZ_LP_信息表.ZCS From ZRZ_LP_信息表 Where ZRZ_LP_信息表.ZRZGUID = " & GUIDArr(i)
        GetSQLRecordAll SqlStr,ZCSArr,SearchCount
        
        If ZCSArr(0) <> "" Then
            ZCS = Transform(ZCSArr(0))
            If ZCS >= 1 And ZCS <= 3 Then
                JZStr = "低层"
                DC_Count = DC_Count + 1
            ElseIf ZCS >= 4 And ZCS <= 6 Then
                JZStr = "多层"
                MC_Count = MC_Count + 1
            ElseIf ZCS >= 7 Then
                JZStr = "高层"
                GC_Count = GC_Count + 1
            Else
                JZStr = ""
            End If
        End If
        
        SqlStr = "Select ZRZ_LP_信息表.DXCS From ZRZ_LP_信息表 Where ZRZ_LP_信息表.ZRZGUID = " & GUIDArr(i)
        GetSQLRecordAll SqlStr,DXCSArr,SearchCount
        If SearchCount > 0 Then
            JZStr = JZStr & "、地下室"
            DX_Count = DX_Count + 1
        Else
            JZStr = JZStr
        End If
        
        If JZStr <> "" Then
            SqlStr = "Update ZRZ_LP_信息表 Set JianZXS = '" & JZStr & "'" & " Where ZRZ_LP_信息表.ZRZGUID = " & GUIDArr(i)
            ProJectName = SSProcess.GetProjectFileName
            SSProcess.OpenAccessMdb ProJectName
            SSProcess.ExecuteAccessSql ProJectName,SqlStr
            SSProcess.CloseAccessMdb ProJectName
        End If
        
    Next 'i
    
    '设置宗地建筑形式
    
    '制空
    JZStr = ""
    
    If DC_Count > 0 Then
        If JZStr <> "" Then
            JZStr = JZStr & "、" & "低层"
        Else
            JZStr = "低层"
        End If
    End If
    
    If MC_Count > 0 Then
        If JZStr <> "" Then
            JZStr = JZStr & "、" & "多层"
        Else
            JZStr = "多层"
        End If
    End If
    
    If GC_Count > 0 Then
        If JZStr <> "" Then
            JZStr = JZStr & "、" & "高层"
        Else
            JZStr = "高层"
        End If
    End If
    
    If DX_Count > 0 Then
        If JZStr <> "" Then
            JZStr = JZStr & "、" & "地下室"
        Else
            JZStr = "地下室"
        End If
    End If
    
    SqlStr = "Update ZD_XM信息属性表 Set JianZXS = '" & JZStr & "'" & " Where ZD_XM信息属性表.ID > 0"
    ProJectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb ProJectName
    SSProcess.ExecuteAccessSql ProJectName,SqlStr
    SSProcess.CloseAccessMdb ProJectName
    
End Function' SetJZ

'==================================================工具类函数===================================================

'数值转换
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

'SQL查询，获取所有的记录
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

'去除小数点
Function DelPoint(ByVal Num,ByRef IntergetNum)
    
    If InStr(Num,".") <> 0 Then
        NumArr = Split(Num,".", - 1,1)
        IntergetNum = NumArr(0)
    Else
        IntergetNum = Num
    End If
End Function' DelPoint