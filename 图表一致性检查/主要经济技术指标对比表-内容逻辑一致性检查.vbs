
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

    FWCheck
    
    ShowCheckRecord

End Sub' OnClick

'===================================================检查函数=======================================================

'房屋用途与功能区用途面积汇总值是否一致（所有幢）
Function FWCheck()
    
    ' 1：主要经济指标面积汇总信息表(ZYJJZBMJHZB)中的每个【YT】：例如：住宅 面积【SCJZMJ】
    ' 2：规划功能区（GHGNQ）表中的【YT】 = “住宅”的所有面积值。
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "房屋用途与功能区用途面积汇总值一致性检查"
    CheckmodelName = "自定义脚本检查类->房屋用途与功能区用途面积汇总值一致性检查"
    strDescription = "房屋用途与功能区用途面积汇总值不一致"

    SqlStr = "Select DISTINCT ZYJJZBMJHZB.YT From ZYJJZBMJHZB Where ZYJJZBMJHZB.ID > 0"
    GetSQLRecordAll SqlStr,YTArr,YTCount
    
    For i = 0 To YTCount - 1
        
        SqlStr = "Select Sum(JG_规划功能区属性表.JZMJ) From JG_规划功能区属性表 Inner Join GeoAreaTB On JG_规划功能区属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And JG_规划功能区属性表.YT = '" & YTArr(i) & "'"
        GetSQLRecordAll SqlStr,SumAreaArr,SumCount
        SumArea = SumAreaArr(0)
        
        SqlStr = "Select ZYJJZBMJHZB.SCJZMJ From ZYJJZBMJHZB Where ZYJJZBMJHZB.ID > 0 And ZYJJZBMJHZB.YT = '" & YTArr(i) & "'"
        GetSQLRecordAll SqlStr,SCJZMJArr,SearchCount
        SCJZMJ = SCJZMJArr(0)
        
        If SumArea - SCJZMJ <> 0 Then
            SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
        End If
    Next 'i

End Function' FWCheck

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