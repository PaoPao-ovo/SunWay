
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
    
    FildsEmptyCheck "PROJECTINFO","测绘单位地址,委托单位,项目名称,项目编号,项目地址,总用地面积(m?),总建筑面积(m?),地下建筑面积(m?),容积率,建筑基底面积(m?),建筑密度(%),绿地率(%),规划许可证编号,测量开始时间,测量完成时间,约定完成时间,测绘目的,项目类别,已有资料情况,控制测量,作业内容,质量控制,成果内容说明,地上建筑面积(m?),装配式建筑面积(m?),测绘单位,测绘单位资质等级,测绘资质证书编号,测绘单位电话,编制人员,审核人员,作业依据,实测住宅户数,规划住宅户数,单块绿地面积,集中绿地面积,登高场地个数","信息表"
    
    ShowCheckRecord
    
End Sub' OnClick

'===================================================检查函数=======================================================

'表字段空值检查
Function FildsEmptyCheck(ByVal TableName,ByVal FildsStr,ByVal TableType)
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = TableName & "空值检查"
    CheckmodelName = "自定义脚本检查类->" & strCheckName
    
    If TableType = "信息表" Then
        
        FildsArr = Split(FildsStr,",", - 1,1)
        For i = 0 To UBound(FildsArr)
            
            SqlStr = "Select " & TableName & ".ID " & "From " & TableName & " Where " & TableName & ".Key = " & "'" & FildsArr(i) & "'"
            GetSQLRecordAll SqlStr,KeyArr,KeyCount
            If KeyCount < 1 Then
                strDescription = "【" & TableName & "】的【Key】为【" & FildsArr(i) & "】缺失"
                SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,0,0,""
            Else
                SqlStr = "Select " & TableName & ".Value " & "From " & TableName & " Where Key = " & "'" & FildsArr(i) & "'"  & " And (" & "Value" & " = '' " & " Or " & "Value" & " = '*' Or " & "Value" & " IS NULL)"
                
                GetSQLRecordAll SqlStr,ValueArr,ValueCount
                
                If ValueCount > 0  Then
                    strDescription = "【" & TableName & "】的【Value】为" & "【" & FildsArr(i) & "】" & "为空值"
                    SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,0,0,""
                End If
            End If
        Next 'i
    End If
    
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