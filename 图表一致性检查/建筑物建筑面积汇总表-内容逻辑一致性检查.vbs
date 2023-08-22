
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
    
    DDFWCheck
    
    ShowCheckRecord
    
End Sub' OnClick

'===================================================检查函数=======================================================

'房屋用途与功能区用途面积汇总值是否一致（按单幢）
Function DDFWCheck()
    
    ' 1：实测楼栋面积汇总信息表（SCLDMJHZXX）表中【LD】=“1#”且【YT】=“住宅”的【JZMJ】
    ' 2：规划功能区（GHGNQ）表中的【SSZRZ】=“1#”且【YT】=“住宅”的【JZMJ】的值的累加值。
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "房屋用途与功能区用途面积汇总值一致性检查（按单幢）"
    CheckmodelName = "自定义脚本检查类->房屋用途与功能区用途面积汇总值一致性检查（按单幢）"
    strDescription = "房屋用途与功能区用途面积汇总值不一致"
    
    '所有的楼栋
    SqlStr = "Select DISTINCT SCLDMJHZXX.LD From SCLDMJHZXX Where SCLDMJHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,LDArr,LDCount
    
    If LDCount > 0 Then
        For i = 0 To LDCount - 1
            
            SqlStr = "Select DISTINCT SCLDMJHZXX.YT From SCLDMJHZXX Where SCLDMJHZXX.ID > 0 And SCLDMJHZXX.LD = '" & LDArr(i) & "'"
            GetSQLRecordAll SqlStr,YTArr,YTCount
            
            For j = 0 To YTCount - 1
                
                SqlStr = "Select Sum(JG_规划功能区属性表.JZMJ) From JG_规划功能区属性表 Inner Join GeoAreaTB On JG_规划功能区属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And JG_规划功能区属性表.SSZRZ = '" & LDArr(i) & "' And JG_规划功能区属性表.YT = '" & YTArr(j) & "'"
                GetSQLRecordAll SqlStr,SumAreaArr,SumCount
                SumArea = Transform(SumAreaArr(0))
                
                SqlStr = "Select SCLDMJHZXX.JZMJ From SCLDMJHZXX Where SCLDMJHZXX.LD = '" & LDArr(i) & "' And SCLDMJHZXX.YT = '" & YTArr(j) & "'"
                GetSQLRecordAll SqlStr,JZMJArr,SearchCount
                JZMJ = Transform(JZMJArr(0))
                
                If JZMJ - SumArea <> 0 Then
                    SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
                End If
                
            Next 'j
        Next 'i   
    End If
    
    
End Function' DDFWCheck

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