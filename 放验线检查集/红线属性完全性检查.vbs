'=============================================表名和字段配置=========================================================

'红线属性表名
TableName_HX = "放验线红线属性表"

'=====================================================检查集配置=====================================================

'检查集项目名称
Dim strGroupName
strGroupName = "放线输出报告前检查"

'检查集组名称
Dim strCheckName
strCheckName = "红线属性检查"

'检查模型名称
Dim CheckmodelName
CheckmodelName = "自定义脚本检查类->红线属性检查"

'检查描述
Dim strDescription
strDescription = "红线扩展属性不全"

'==============================================函数主体========================================================

'入口函数
Sub OnClick()
    ClearCheckRecord()
    HxCheck TableName_HX
End Sub' OnClick

'线属性检查
Function HxCheck(tablename)
    projectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb projectName
    SqlString = "Select * From " & tablename & "  inner join GeoAreaTB on " & tablename & ".ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0"
    'MsgBox SqlString
    GetSQLRecordAll projectName,SqlString,arSQLRecord,iRecordCount
    For j = 0 To iRecordCount - 1
        RecordString = arSQLRecord(j)
        Recordarr = Split(RecordString,",", - 1,1)
        For k = 1 To 10
            If Recordarr(k) = "*" Then
                id = Recordarr(0)
                x = SSProcess.GetObjectAttr(id,"SSObj_X")
                y = SSProcess.GetObjectAttr(id,"SSObj_Y")
                SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription, x, y,0,2,id, ""
                Exit For
            End If
            If Recordarr(k) = "" Then
                id = Recordarr(0)
                x = SSProcess.GetObjectAttr(id,"SSObj_X")
                y = SSProcess.GetObjectAttr(id,"SSObj_Y")
                SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription, x, y,0,2,id, ""
                Exit For
            End If
            If Recordarr(k) = "0" Then
                id = Recordarr(0)
                x = SSProcess.GetObjectAttr(id,"SSObj_X")
                y = SSProcess.GetObjectAttr(id,"SSObj_Y")
                SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription, x, y,0,2,id, ""
                Exit For
            End If
            If Recordarr(k) = Null Then
                id = Recordarr(0)
                x = SSProcess.GetObjectAttr(id,"SSObj_X")
                y = SSProcess.GetObjectAttr(id,"SSObj_Y")
                SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription, x, y,0,2,id, ""
                Exit For
            End If
        Next 'k
        'MsgBox Recordarr(0)
    Next 'j
    SSProcess.CloseAccessMdb projectName
    SSProcess.ShowCheckOutput
    SSProcess.SaveCheckRecord
End Function' ZDX

'获取所有记录
Function GetSQLRecordAll(ByRef mdbName, ByRef StrSqlStatement, ByRef arSQLRecord(), ByRef iRecordCount)
    If StrSqlStatement = "" Then
        MsgBox "查询语句为空，操作停止！",48
    End If
    iRecordCount =  - 1
    'SQL语句
    sql = StrSqlStatement
    '打开记录集
    SSProcess.OpenAccessRecordset mdbName, sql
    '获取记录总数
    RecordCount = SSProcess.GetAccessRecordCount (mdbName, sql)
    If RecordCount > 0 Then
        iRecordCount = 0
        ReDim arSQLRecord(RecordCount)
        '将记录游标移到第一行
        SSProcess.AccessMoveFirst mdbName, sql
        iRecordCount = 0
        '浏览记录
        While SSProcess.AccessIsEOF (mdbName, sql) = 0
            fields = ""
            values = ""
            '获取当前记录内容
            SSProcess.GetAccessRecord mdbName, sql, fields, values
            arSQLRecord(iRecordCount) = values
            iRecordCount = iRecordCount + 1
            '移动记录游标
            SSProcess.AccessMoveNext mdbName, sql
        WEnd
    End If
    '关闭记录集
    SSProcess.CloseAccessRecordset mdbName, sql
End Function

'清空检查集
Function ClearCheckRecord()
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
End Function' ClearCheckRecord