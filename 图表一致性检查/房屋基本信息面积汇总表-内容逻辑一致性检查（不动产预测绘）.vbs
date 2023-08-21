
'===================================================检查参数定义==================================================

'检查组名称
Dim strGroupName

'检查项名称
Dim strCheckName

'检查模型名称
Dim CheckmodelName

'检查描述
Dim strDescription

'====================================================入口=========================================================

'检查入口
Sub OnClick()
    
    ClearCheckRecord

    JzZmjCheck "ZD_宗地基本信息属性表"
    
    DxJzzMjCheck

    DsJzzMjCheck

    HDSCheck

    HDXCheck
    
    ShowCheckRecord

End Sub' OnClick

'===================================================检查函数=======================================================

'预测绘建筑总面积检查
Function JzZmjCheck(ByVal TableName)
    
    ' 1 建筑面积：宗地基本信息表【JZMJ】（ZD_宗地基本信息属性表[JZZMJ]）
    ' 2 地上部分总计：房屋地上地下总面积汇总信息表（FWDSDXZMJHZXX）字段：【YCDSZJZMJ】或字段【SCDSZJZMJ】
    ' 3 地上部分总计：房屋地上地下总面积汇总信息表（FWDSDXZMJHZXX）字段：【YCDXZJZMJ】或字段【SCDXZJZMJ】
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "建筑总面积检查"
    CheckmodelName = "自定义脚本检查类->建筑总面积检查"
    strDescription = TableName & "的【JZZMJ】与FWDSDXZMJHZXX表的【YCDSZJZMJ】和【YCDXZJZMJ】之和不相等"

    '获取总建筑面积 JZZMJ
    SqlStr = "Select " & TableName & ".ID,JZZMJ From " & TableName & " Inner Join GeoAreaTB On " & TableName & ".ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 "
    GetSQLRecordAll SqlStr,TotalAreaArr,SearchCount
    
    If SearchCount = 1 Then
        ZDArr = Split(TotalAreaArr(0),",", - 1,1)
        JZZMJ = Transform(ZDArr(1))
    Else
        JZZMJ = 0
        Dim ZDArr(0)
        ZDArr(0) =  - 1
    End If
    
    '获取总地上建筑面积 YCDSZMJ
    SqlStr = "Select Sum(FWDSDXZMJHZXX.YCDSZJZMJ) From FWDSDXZMJHZXX WHERE FWDSDXZMJHZXX.ID > 0"
    GetSQLRecordAll SqlStr,YCDSArr,SearchCount
    YCDSZMJ = Transform(YCDSArr(0))
    
    '获取总地下建筑面积 YCDXZMJ
    SqlStr = "Select Sum(FWDSDXZMJHZXX.YCDXZJZMJ) From FWDSDXZMJHZXX WHERE FWDSDXZMJHZXX.ID > 0"
    GetSQLRecordAll SqlStr,YCDXArr,SearchCount
    YCDXZMJ = Transform(YCDXArr(0))
    
    SumArea = YCDSZMJ + YCDXZMJ
    
    '检查判断
    If JZZMJ - SumArea <> 0 Then
        If ZDArr(0) <> - 1 Then
            SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(ZDArr(0),"SSObj_X"),SSProcess.GetObjectAttr(ZDArr(0),"SSObj_Y"),0,2,ZDArr(0),""
        Else
            SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
        End If
        
    End If

End Function' JzZmjCheck 

'预测地下建筑总面积检查
Function DxJzzMjCheck()
    
    ' 1:地下部分总计: 房屋地上地下总面积汇总信息表（FWDSDXZMJHZXX）字段：【YCDXZJZMJ】或字段【SCDXZJZMJ】
    ' 2:其他部分+人防部位：房屋类型面积汇总信息表（FWLXMJHZXX）表【SCJZMJ】或【YCJZMJ】的累计和。（条件限制：空间位置【KJWZ】为：地下）。
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "房屋基本信息面积汇总逻辑检查"
    CheckmodelName = "自定义脚本检查类->房屋基本信息面积汇总逻辑检查"
    strDescription = "预测地下总建筑与其他部分和人防部分面积之和不等"

    '获取地下总面积 YCDXZMJ
    SqlStr = "Select Sum(FWDSDXZMJHZXX.YCDXZJZMJ) From FWDSDXZMJHZXX WHERE FWDSDXZMJHZXX.ID > 0"
    GetSQLRecordAll SqlStr,YCDXArr,SearchCount
    YCDXZMJ = Transform(YCDXArr(0))
    
    '地下其他部分面积和人防部分面积 QTMJ
    SqlStr = "Select Sum(FWLXMJHZXX.YCJZMJ) From FWLXMJHZXX WHERE FWLXMJHZXX.ID > 0 And FWLXMJHZXX.KJWZ = '地下' "
    GetSQLRecordAll SqlStr,QTArr,SearchCount
    QTMJ = Transform(QTArr(0))
    
    If YCDXZMJ - QTMJ <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If

End Function' DxJzzMjCheck

'预测地上建筑总面积检查
Function DsJzzMjCheck()
    
    ' 1：地上部分总计：房屋地上地下总面积汇总信息表（FWDSDXZMJHZXX）字段：【YCDSZJZMJ】或字段【SCDSZJZMJ】
    ' 2: 地上户面积统计: 房屋类型面积汇总信息表（FWLXMJHZXX）表【SCJZMJ】或【YCJZMJ】的累计和。（条件限制：空间位置【KJWZ】为：地上）。
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "房屋基本信息面积汇总逻辑检查"
    CheckmodelName = "自定义脚本检查类->房屋基本信息面积汇总逻辑检查"
    strDescription = "预测地下总建筑与其他部分和人防部分面积之和不等"

    '获取地下总面积 YCDSZMJ
    SqlStr = "Select Sum(FWDSDXZMJHZXX.YCDSZJZMJ) From FWDSDXZMJHZXX WHERE FWDSDXZMJHZXX.ID > 0"
    GetSQLRecordAll SqlStr,YCDXArr,SearchCount
    YCDSZMJ = Transform(YCDXArr(0))
    
    '地下其他部分面积和人防部分面积 QTMJ
    SqlStr = "Select Sum(FWLXMJHZXX.YCJZMJ) From FWLXMJHZXX WHERE FWLXMJHZXX.ID > 0 And FWLXMJHZXX.KJWZ = '地上' "
    GetSQLRecordAll SqlStr,QTArr,SearchCount
    QTMJ = Transform(QTArr(0))
    
    If YCDSZMJ - QTMJ <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If

End Function' DsJzzMjCheck

'H表地上检查
Function HDSCheck()
    
    ' 1：房屋类型汇总值：房屋类型面积汇总信息表（FWLXMJHZXX）中【FWLXMC】和【SCJZMJ】和【KJWZ】
    ' 2：户（H）：实际层数【SJCS】、房屋类型名称【FWLXMC】、预测建筑面积【YCJZMJ】、实测建筑面积【SCJZMJ】的值的累加和。（说明：按照地上、地下分别检查判断）
    ' 举例说明：当房屋类型面积汇总信息表（FWLXMJHZXX）的【KJWZ】=地上 且【FWLXMC】=”住宅”的【SCJZMJ】的值是否等于户（H）的【SJCS】大于0且【FWLXMC】=”住宅”的【SCJZMJ】的值的累加和。
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "房屋基本信息面积汇总逻辑检查"
    CheckmodelName = "自定义脚本检查类->房屋基本信息面积汇总逻辑检查"
    strDescription = "房屋类型面积汇总值与户表统计面积值不一致"

    '获取所有的房屋类型名称 FWLXMCArr
    SqlStr = "Select DISTINCT FWLXMJHZXX.FWLXMC From FWLXMJHZXX Where FWLXMJHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,FWLXMCArr,FWLXMCCount
    
    If FWLXMCCount > 1 Then
        '获取对应的预测地上建筑面积
        For CurrentCount = 0 To UBound(FWLXMCArr)
            If FWLXMCArr(CurrentCount) <> "" Then
                
                SqlStr = "Select Sum(FWLXMJHZXX.YCJZMJ) Where FWLXMJHZXX.ID > 0 And FWLXMJHZXX.FWLXMC = " & "'" & FWLXMCArr(CurrentCount) & "' And " & "FWLXMJHZXX.KJWZ = '地上' "
                GetSQLRecordAll SqlStr,YCJZMJArr,SearchCount
                YCJZMJ = Transform(YCJZMJArr(0))
                
                SqlStr = "Select Sum(H.YCJZMJ) Where H.ID > 0 And H.FWLXMC = " & "'" & FWLXMCArr(CurrentCount) & "' And " & "H.KJWZ = '地上' And H.SJCS > 0 "
                GetSQLRecordAll SqlStr,HYCJZMJArr,SearchCount
                HYCJZMJ = Transform(HYCJZMJArr(0))
                
                If YCJZMJ - HYCJZMJ <> 0 Then
                    SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
                End If
                
            End If
        Next 'CurrentCount
    End If

End Function' HDSCheck

'H表地下检查
Function HDXCheck()
    
    ' 1：房屋类型汇总值：房屋类型面积汇总信息表（FWLXMJHZXX）中【FWLXMC】和【SCJZMJ】和【KJWZ】
    ' 2：户（H）：实际层数【SJCS】、房屋类型名称【FWLXMC】、预测建筑面积【YCJZMJ】、实测建筑面积【SCJZMJ】的值的累加和。（说明：按照地上、地下分别检查判断）
    ' 举例说明：当房屋类型面积汇总信息表（FWLXMJHZXX）的【KJWZ】=地上 且【FWLXMC】=”住宅”的【SCJZMJ】的值是否等于户（H）的【SJCS】大于0且【FWLXMC】=”住宅”的【SCJZMJ】的值的累加和。
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "房屋基本信息面积汇总逻辑检查"
    CheckmodelName = "自定义脚本检查类->房屋基本信息面积汇总逻辑检查"
    strDescription = "房屋类型面积汇总值与户表统计面积值不一致"

    '获取所有的房屋类型名称 FWLXMCArr
    SqlStr = "Select DISTINCT FWLXMJHZXX.FWLXMC From FWLXMJHZXX Where FWLXMJHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,FWLXMCArr,FWLXMCCount
    
    If FWLXMCCount > 1 Then
        '获取对应的预测地下建筑面积
        For CurrentCount = 0 To UBound(FWLXMCArr)
            If FWLXMCArr(CurrentCount) <> "" Then
                
                SqlStr = "Select Sum(FWLXMJHZXX.YCJZMJ) Where FWLXMJHZXX.ID > 0 And FWLXMJHZXX.FWLXMC = " & "'" & FWLXMCArr(CurrentCount) & "' And " & "FWLXMJHZXX.KJWZ = '地下' "
                GetSQLRecordAll SqlStr,YCJZMJArr,SearchCount
                YCJZMJ = Transform(YCJZMJArr(0))
                
                SqlStr = "Select Sum(H.YCJZMJ) Where H.ID > 0 And H.FWLXMC = " & "'" & FWLXMCArr(CurrentCount) & "' And " & "H.KJWZ = '地下' And H.SJCS > 0 "
                GetSQLRecordAll SqlStr,HYCJZMJArr,SearchCount
                HYCJZMJ = Transform(HYCJZMJArr(0))
                
                If YCJZMJ - HYCJZMJ <> 0 Then
                    SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
                End If
            End If
        Next 'CurrentCount
    End If

End Function' HDXCheck

'======================================================工具类函数====================================================

'清空缓存的所有检查记录
Function ClearCheckRecord()
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
End Function' ClearCheckRecord

'显示所有检查记录
Function ShowCheckRecord()
    SSProcess.ShowCheckOutput
    SSProcess.SaveCheckRecord
End Function' ShowCheckRecord

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