
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
    CheckFilds = X,Y,GC '检查字段
    AccuracyCheck KZDZBCGXXB,CheckFilds,3
End Sub' OnClick

'=====================================================检查函数======================================================

'小数位数精度检查
Function AccuracyCheck(ByVal TableName,ByVal FildsStr,ByVal CheckBits) 'TableName = 表名,FildsStr = 查询的字段字符串,CheckBits = 检查位数
    
    '检查记录配置
    strGroupName = "属性精度检查"
    strCheckName = "控制点坐标小数位规范性检查"
    CheckmodelName = "自定义脚本检查类->控制点坐标小数位规范性检查"
    
    '查询字段值
    SqlStr = "Select " & TableName & "." & "objectid," & FildsStr & " From " & TableName & "Where " & TableName & ".ID > 0"
    GetSQLRecordAll SqlStr,ValArr,SearchCount  'ValArr = [(值1,值2,值3....)(值1,值2,值3....)]
    
    '字段名称数组
    FildsNameArr = Split(FildsStr,",", - 1,1)
    
    '遍历字段值
    For i = 0 To SearchCount - 1
        CurrentValArr = Split(ValArr(i),",", - 1,1)
        For j = 1 To UBound(CurrentValArr)
            DecimalJudgment CurrentValArr(j),CheckBits,ErrorBool
            If ErrorBool Then
                strDescription = TableName & "表，ObjectId为" & CurrentValArr(0) & "字段：" & FildsNameArr(j) & "小数位数大于三"
                SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,0,0,""
            End If
        Next 'j
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
Function DecimalJudgment(ByVal Num,ByVal CheckBits,ByVal ErrorBool) 'Num = 检查数,CheckBits = 检查位数,ErrorBool = 是否错误,错误返回True
    
    ErrorBool = False
    
    DecimalPointPoi = InStr(1,Num,".",1)
    
    If Num = "" Then
        ErrorBool = False
    ElseIf Num <> "" And DecimalPointPoi = 0 Then
        ErrorBool = False
    ElseIf Num <> "" And DecimalPointPoi > 0 Then
        DecimalLen = Len(Num) - DecimalPointPoi
        If DecimalLen < CheckBits Then
            ErrorBool = False
        Else
            ErrorBool = True
        End If
    Else
        ErrorBool = True
    End If
    
End Function' DecimalJudgment