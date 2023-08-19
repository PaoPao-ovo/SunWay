
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
    JzZmjCheck "ZD_宗地基本信息属性表"
End Sub' OnClick

'===================================================检查函数=======================================================

'实测绘建筑总面积检查
Function JzZmjCheck(ByVal TableName)
    
    ' 1 建筑面积：宗地基本信息表【JZMJ】（ZD_宗地基本信息属性表[JZZMJ]）
    ' 2 地上部分总计：房屋地上地下总面积汇总信息表（FWDSDXZMJHZXX）字段：【YCDSZJZMJ】或字段【SCDSZJZMJ】
    ' 3 地上部分总计：房屋地上地下总面积汇总信息表（FWDSDXZMJHZXX）字段：【YCDXZJZMJ】或字段【SCDXZJZMJ】
    
    '检查记录配置
    strGroupName = "房屋基本信息面积汇总逻辑检查"
    strCheckName = "建筑总面积检查"
    CheckmodelName = "自定义脚本检查类->建筑总面积检查"
    strDescription = TableName & "的【JZZMJ】与FWDSDXZMJHZXX表的【SCDSZJZMJ】和【SCDXZJZMJ】之和不相等"
    
    '获取总建筑面积 JZZMJ
    SqlStr = "Select " & TableName & ".ID,JZZMJ From " & TableName "Inner Join GeoAreaTB On" & TableName & ".ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 "
    GetSQLRecordAll SqlStr,TotalAreaArr,SearchCount
    If SearchCount = 1 Then
        ZDArr = Split(TotalAreaArr(0),",", - 1,1)
        JZZMJ = Transform(ZDArr(1))
    End If
    
    '获取总地上建筑面积 SCDSZMJ
    SqlStr = "Select Sum(FWDSDXZMJHZXX.SCDSZJZMJ) From FWDSDXZMJHZXX WHERE FWDSDXZMJHZXX.ID > 0"
    GetSQLRecordAll SqlStr,SCDSArr,SearchCount
    SCDSZMJ = SCDSArr(0)
    
    '获取总地下建筑面积 SCDXZMJ
    SqlStr = "Select Sum(FWDSDXZMJHZXX.SCDXZJZMJ) From FWDSDXZMJHZXX WHERE FWDSDXZMJHZXX.ID > 0"
    GetSQLRecordAll SqlStr,SCDXArr,SearchCount
    SCDXZMJ = SCDXArr(0)
    
    SumArea = SCDSZMJ + SCDXZMJ
    
    '检查判断
    If JZZMJ <> SumArea Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,ZDArr(0),""
    End If
    
End Function' JzZmjCheck 

'实测地下建筑总面积检查
Function DxJzzMjCheck()
    
    ' 1:地下部分总计: 房屋地上地下总面积汇总信息表（FWDSDXZMJHZXX）字段：【YCDXZJZMJ】或字段【SCDXZJZMJ】
    ' 2:其他部分+人防部位：房屋类型面积汇总信息表（FWLXMJHZXX）表【SCJZMJ】或【YCJZMJ】的累计和。（条件限制：空间位置【KJWZ】为：地下）。
    
    '检查记录配置
    strGroupName = "房屋基本信息面积汇总逻辑检查"
    strCheckName = "建筑地下总面积检查"
    CheckmodelName = "自定义脚本检查类->建筑地下总面积检查"
    strDescription = "实测地下总建筑与其他部分和人防部分面积之和不等"
    
    '获取地下总面积 SCDXZMJ
    SqlStr = "Select Sum(FWDSDXZMJHZXX.SCDXZJZMJ) From FWDSDXZMJHZXX WHERE FWDSDXZMJHZXX.ID > 0"
    GetSQLRecordAll SqlStr,SCDXArr,SearchCount
    SCDXZMJ = SCDXArr(0)
    
    '地下其他部分面积和人防部分面积 QTMJ
    SqlStr = "Select Sum(FWLXMJHZXX.SCJZMJ) From FWLXMJHZXX WHERE FWLXMJHZXX.ID > 0 And FWLXMJHZXX.KJWZ = '地下' "
    GetSQLRecordAll SqlStr,QTArr,SearchCount
    QTMJ = QTArr(0)
    
    If SCDXZMJ <> QTMJ Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' DxJzzMjCheck

'实测地上建筑总面积检查
Function DsJzzMjCheck()
    
    ' 1：地上部分总计：房屋地上地下总面积汇总信息表（FWDSDXZMJHZXX）字段：【YCDSZJZMJ】或字段【SCDSZJZMJ】
    ' 2: 地上户面积统计: 房屋类型面积汇总信息表（FWLXMJHZXX）表【SCJZMJ】或【YCJZMJ】的累计和。（条件限制：空间位置【KJWZ】为：地上）。
    
    '检查记录配置
    strGroupName = "房屋基本信息面积汇总逻辑检查"
    strCheckName = "建筑地下总面积检查"
    CheckmodelName = "自定义脚本检查类->建筑地下总面积检查"
    strDescription = "实测地下总建筑与其他部分和人防部分面积之和不等"
    
    '获取地下总面积 SCDSZMJ
    SqlStr = "Select Sum(FWDSDXZMJHZXX.SCDSZJZMJ) From FWDSDXZMJHZXX WHERE FWDSDXZMJHZXX.ID > 0"
    GetSQLRecordAll SqlStr,SCDXArr,SearchCount
    SCDSZMJ = SCDXArr(0)
    
    '地下其他部分面积和人防部分面积 QTMJ
    SqlStr = "Select Sum(FWLXMJHZXX.SCJZMJ) From FWLXMJHZXX WHERE FWLXMJHZXX.ID > 0 And FWLXMJHZXX.KJWZ = '地上' "
    GetSQLRecordAll SqlStr,QTArr,SearchCount
    QTMJ = QTArr(0)
    
    If SCDSZMJ <> QTMJ Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' DsJzzMjCheck

'======================================================工具类函数====================================================

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