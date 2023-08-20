'表字段空值检查
Function FildsEmptyCheck(ByVal TableName,ByVal FildsStr,ByVal TableType)
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = TableName & "空值检查"
    CheckmodelName = "自定义脚本检查类->" & strCheckName
    
    If TableType = "信息表" Then
        MdbName = SSProcess.GetProjectFileName
        SSProcess.OpenAccessMdb MdbName
        FildsArr = Split(FildsStr,",", - 1,1)
        For i = 0 To UBound(FildsArr)
            '字段名称,数据类型,字段大小,字段属性,字段序号,是否必须字段,是否允许为空,排序比较方式,字段别名,源字段名,源表名,字段规则,字段规则内容,缺省值
            '数字类型为 7(Double),6(Float)
            '字符串为 10(Char & String)
            SSProcess.GetAccessFieldInfo1 MdbName,TableName,FildsArr(i),FieldsInfo
            SSProcess.CloseAccessMdb MdbName
            FieldsInfoArr = Split(FieldsInfo,",", - 1,1)
            If FieldsInfoArr(1) = "10" Then
                SqlStr = "Select " & TableName & "." & FildsArr(i) & " Where " & FildsArr(i) & " = '' Or " & FildsArr(i) & " = '*' Or " & FildsArr(i) & " = NULL "
                GetSQLRecordAll SqlStr,StringArr,StringEmptyCount
                If StringEmptyCount > 0 Then
                    strDescription = TableName & "的" & FildsArr(i) & "存在空值"
                    SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,0,0,""
                End If
            ElseIf FieldsInfoArr(1) = "7" Or FieldsInfoArr(1) = "6" Then
                SqlStr = "Select " & TableName & "." & FildsArr(i) & " Where " & FildsArr(i) & " = NULL "
                GetSQLRecordAll SqlStr,NumArr,NumEmptyCount
                If NumEmptyCount > 0 Then
                    strDescription = TableName & "的" & FildsArr(i) & "存在空值"
                    SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,0,0,""
                End If
            End If
        Next 'i
    End If
End Function' FildsEmptyCheck
