
'====================================================入口=========================================================

'检查入口
Sub OnClick()
    
End Sub' OnClick

'=====================================================检查函数======================================================

'小数位数精度检查
Function AccuracyCheck(ByVal TableName,ByVal TableType,ByVal FildsStr) 'TableName = 表名,TableType = 表类型,FildsStr = 查询的字段字符串
    
    SqlStr = "Select " & TableName & "." FildsStr & " From " & TableName
    GetSQLRecordAll SqlStr,ValArr,SearchCount  'ValArr = [(值1值2值3....),(值1值2值3....)]
    For i = 0 To SearchCount - 1
        Split(ValArr(i),",", - 1,1)
        
    Next 'i
    
End Function' AccuracyCheck




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

'小数位数判断
Function DecimalJudgment(ByVal Num,ByVal CheckBits) 'Num = 检查数,CheckBits = 检查位数
    
End Function' DecimalJudgment