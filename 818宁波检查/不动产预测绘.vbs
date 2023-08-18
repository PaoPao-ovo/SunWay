
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

'预测绘建筑总面积检查
Function JzZmjCheck(ByVal TableName)
    
    '检查记录配置
    strGroupName = "房屋基本信息面积汇总逻辑检查"
    strCheckName = "建筑总面积检查"
    CheckmodelName = "自定义脚本检查类->建筑总面积检查"
    
    '获取总建筑面积
    SqlStr = "Select " & TableName & ".ID,JZZMJ From " & TableName "Inner Join GeoAreaTB On" & TableName & ".ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 "
    GetSQLRecordAll SqlStr,TotalAreaArr,SearchCount
    If SearchCount = 1 Then
        ZDArr = Split(TotalAreaArr(0),",", - 1,1)
        JZZMJ = Transform(ZDArr(1))
    End If
    
    '获取总地上建筑面积
    SqlStr = "Select Sum(FWDSDXZMJHZXX.YCDSZJZMJ) From FWDSDXZMJHZXX WHERE FWDSDXZMJHZXX.ID > 0"
    GetSQLRecordAll SqlStr,YCDSArr,SearchCount
    YCDSZMJ = YCDSArr(0)
    
    '获取总地下建筑面积
    SqlStr = "Select Sum(FWDSDXZMJHZXX.YCDXZJZMJ) From FWDSDXZMJHZXX WHERE FWDSDXZMJHZXX.ID > 0"
    GetSQLRecordAll SqlStr,YCDXArr,SearchCount
    YCDXZMJ = YCDXArr(0)
    
    SumArea = YCDSZMJ + YCDXZMJ
    
    '检查判断
    If JZZMJ <> SumArea Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,ZDArr(0),""
    End If
    
End Function' JzZmjCheck 

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