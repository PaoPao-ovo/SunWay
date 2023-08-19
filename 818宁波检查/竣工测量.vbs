
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
    
End Sub' OnClick

'===================================================检查函数=======================================================

'建筑面积值与幢面积汇总值是否一致
Function ZhuangCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【JZMJ】
    ' 2:自然幢（ZRZ）表中【JZMJ】累计汇总。
    
    '检查记录配置
    strGroupName = "总体指标表面积逻辑检查"
    strCheckName = "建筑面积值与幢面积汇总值一致性检查"
    CheckmodelName = "自定义脚本检查类->建筑面积值与幢面积汇总值一致性检查"
    strDescription = "建筑面积值与幢面积汇总值不一致"
    
    '获取总建筑面积 JZMJ
    SqlStr = "Select Sum(JGSCHZXX.JZMJ) From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,JZMJArr,SearchCount
    JZMJ = JZMJArr(0)
    
    '获取自然幢总面积
    
End Function' ZhuangCheck


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