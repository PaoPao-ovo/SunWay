'=============================================表名和字段配置=========================================================

'放样点属性表名
Dim TableArr(2)

TableArr(0) = "理论放样点属性表"
TableArr(1) = "实测放样点属性表"

'=====================================================检查集配置=====================================================

'检查集项目名称
Dim strGroupName
strGroupName = "验线检查"

'检查集组名称
Dim strCheckName
strCheckName = "放样点检查"

'检查模型名称
Dim CheckmodelName
CheckmodelName = "自定义脚本检查类->放样点检查"

'检查描述
Dim strDescription
strDescription = "建筑物名称扩展属性为空"

'==============================================函数主体========================================================

'入口函数
Sub OnClick()
    ClearCheckRecord()
     For i = 0 To 1
        PoiCheck TableArr(i)
    Next 'i
End Sub' OnClick

'放样点属性检查
Function PoiCheck(tablename)
    projectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb projectName
    SqlString = "Select * From " & tablename & "  inner join GeoPointTB on " & tablename & ".ID = GeoPointTB.ID WHERE (GeoPointTB.Mark Mod 2)<>0"
    'MsgBox SqlString
    GetSQLRecordAll projectName,SqlString,arSQLRecord,iRecordCount
    For j = 0 To iRecordCount - 1
        RecordString = arSQLRecord(j)
        'MsgBox RecordString
        Recordarr = Split(RecordString,",", - 1,1)
            If Recordarr(1) = "*" Then
                id = Recordarr(0)
                x = SSProcess.GetObjectAttr(id,"SSObj_X")
                y = SSProcess.GetObjectAttr(id,"SSObj_Y")
                SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription, x, y,0,0,id, ""
            End If
            If Recordarr(1) = "" Then
                id = Recordarr(0)
                x = SSProcess.GetObjectAttr(id,"SSObj_X")
                y = SSProcess.GetObjectAttr(id,"SSObj_Y")
                SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription, x, y,0,0,id, ""
            End If
            If Recordarr(1) = "0" Then
                id = Recordarr(0)
                x = SSProcess.GetObjectAttr(id,"SSObj_X")
                y = SSProcess.GetObjectAttr(id,"SSObj_Y")
                SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription, x, y,0,0,id, ""
            End If
            If Recordarr(1) = Null Then
                id = Recordarr(0)
                x = SSProcess.GetObjectAttr(id,"SSObj_X")
                y = SSProcess.GetObjectAttr(id,"SSObj_Y")
                SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription, x, y,0,0,id, ""
            End If
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