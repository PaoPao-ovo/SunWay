
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

    FildsEmptyCheck "JZMJDBCYSM","CYYY","信息表"

    ShowCheckRecord

End Sub' OnClick

'===================================================检查函数=======================================================

'表字段空值检查
Function FildsEmptyCheck(ByVal TableName,ByVal FildsStr,ByVal TableType)

    MdbName = SSProcess.GetProjectFileName

    SSProcess.OpenAccessMdb MdbName

    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = TableName & "空值检查"
    CheckmodelName = "自定义脚本检查类->" & strCheckName
    
    If TableType = "信息表" Then
        
        FildsArr = Split(FildsStr,",", - 1,1)
        For i = 0 To UBound(FildsArr)
        
            '字段名称,数据类型,字段大小,字段属性,字段序号,是否必须字段,是否允许为空,排序比较方式,字段别名,源字段名,源表名,字段规则,字段规则内容,缺省值
            '数字类型为 7(Double),6(Float)
            '字符串为 10(Char & String)
            
            SSProcess.GetAccessFieldInfo1 MdbName,TableName,FildsArr(i),FieldsInfo
            
            FieldsInfoArr = Split(FieldsInfo,",", - 1,1)
            If FieldsInfoArr(1) = "10" Then

                SqlStr = "Select " & TableName & "." & FildsArr(i) & " From " & TableName & " Where " & FildsArr(i) & " = '' Or " & FildsArr(i) & " = '*' Or " & FildsArr(i) & " IS NULL "
                GetSQLRecordAll SqlStr,StringArr,StringEmptyCount

                If StringEmptyCount > 0 Then
                    strDescription = "【" & TableName & "】" & "的" & "【" & FildsArr(i) & "】" & "存在空值"
                    SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,0,0,""
                End If

            ElseIf FieldsInfoArr(1) = "7" Or FieldsInfoArr(1) = "6" Then

                SqlStr = "Select " & TableName & "." & FildsArr(i) & " From " & TableName & " Where " & FildsArr(i) & " IS NULL Or " & FildsArr(i) & " = '' "
                GetSQLRecordAll SqlStr,NumArr,NumEmptyCount

                If NumEmptyCount > 0 Then
                    strDescription = "【" & TableName & "】" & "的" & "【" & FildsArr(i) & "】" & "存在空值"
                    SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,0,0,""
                End If

            End If
        Next 'i
    End If

    SSProcess.CloseAccessMdb MdbName

End Function' FildsEmptyCheck

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