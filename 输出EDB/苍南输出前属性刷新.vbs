

Sub OnClick()
    KeyStr = "编号,项目名称,项目地址,设计单位,建设单位,委托单位,外业时间,测绘时间,点最大较差值,高程最大较差值"
    SSProcess.ClearInputParameter
    
    KeyArr = Split(KeyStr,",", - 1,1)
    
    For i = 0 To UBound(KeyArr) - 2
        SSProcess.AddInputParameter KeyArr(i) , SSProcess.ReadEpsIni("管线报告信息", KeyArr(i) ,"") , 0 , "" , ""
    Next 'i
    
    ShowBoolen = SSProcess.ShowInputParameterDlg ("管线报告信息录入")
    
    For i = 0 To UBound(KeyArr)
        SSProcess.WriteEpsIni "管线报告信息", KeyArr(i) ,SSProcess.GetInputParameter(KeyArr(i))
    Next 'i

    GETNOTEZB
    
    SSProcess.OpenAccessMdb SSProcess.GetProjectFileName
    
    SqlStr = "Update 地下管线点属性表 SET XMBH = " & "'" & SSProcess.ReadEpsIni("管线报告信息", "编号" ,"") & "'"
    SsProcess.ExecuteAccessSql SSProcess.GetProjectFileName,SqlStr
    
    SqlStr = "Update 地下管线线属性表 SET XMBH = " & "'" & SSProcess.ReadEpsIni("管线报告信息", "编号" ,"") & "'"
    SsProcess.ExecuteAccessSql SSProcess.GetProjectFileName,SqlStr
    
    SSProcess.CloseAccessMdb SSProcess.GetProjectFileName
    
    SSProcess.MapMethod "clearattrbuffer",  "地下管线点属性表"
    SSProcess.MapMethod "clearattrbuffer",  "地下管线线属性表"
End Sub' OnClick

'刷注记坐标值
Function GETNOTEZB
    projectName = SSProcess.GetProjectFileName
    sql = "Select 地下管线点属性表.id,GXDDH From 地下管线点属性表 INNER JOIN GeoPOINTTB ON 地下管线点属性表.ID = GeoPOINTTB.ID WHERE ([GeopointTB].[Mark] Mod 2)<>0"
    GetSQLRecordAll projectName,sql,arSQLRecord,iRecordCount
    If iRecordCount > 0 Then
        For i = 0 To iRecordCount - 1
            arTemp = Split(arSQLRecord(i), ",")
            gdid = arTemp(0)
            gxddh = arTemp(1)
            SSProcess.ClearSelection
            SSProcess.ClearSelectCondition
            SSProcess.SetSelectCondition "SSObj_FontString", "==", gxddh
            SSProcess.SelectFilter
            gdCount = SSProcess.GetSelnoteCount
            If gdCount > 0 Then
                X = Round(SSProcess.GetSelnoteValue(0,"SSObj_X"),3)
                Y = Round(SSProcess.GetSelnoteValue(0,"SSObj_Y"),3)
                SSProcess.SetObjectAttr gdid, "[MapNo_Y]", X
                SSProcess.SetObjectAttr gdid, "[MapNo_X]",Y
            End If
        Next
    End If
    '未用孔数刷新
    sql = "Select 地下管线线属性表.id,zks,yyks From 地下管线线属性表 INNER JOIN GeolineTB ON 地下管线线属性表.ID = GeolineTB.ID WHERE ([GeolineTB].[Mark] Mod 2)<>0"
    GetSQLRecordAll projectName,sql,arSQLRecord,iRecordCount
    If iRecordCount > 0 Then
        SSProcess.OpenAccessMdb projectName
        For i = 0 To iRecordCount - 1
            arTemp = Split(arSQLRecord(i), ",")
            gxid = arTemp(0)
            zks = arTemp(1)
            yyks = arTemp(2)
            lash = IsNumeric (zks)
            If    lash = True Then
                If yyks = "" Then yyks = 0
                wyks = CDbl(zks) - CDbl(yyks)
                sql = "update  地下管线线属性表 set wyks='" & wyks & "' where 地下管线线属性表.id=" & gxid
                SSProcess.ExecuteAccessSql  projectName,sql
            End If
        Next
        SSProcess.MapMethod "clearattrbuffer",  "地下管线线属性表"
        SSProcess.CloseAccessMdb mdbName
    End If
End Function

Function GetSQLRecordAll(ByRef mdbName, ByRef StrSqlStatement, ByRef arSQLRecord(), ByRef iRecordCount)
    SSProcess.OpenAccessMdb mdbName
    iRecordCount =  - 1
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
        '浏览记录
        While SSProcess.AccessIsEOF (mdbName, sql) = 0
            fields = ""
            values = ""
            '获取当前记录内容
            SSProcess.GetAccessRecord mdbName, sql, fields, values
            arSQLRecord(iRecordCount) = values                                        '查询记录
            iRecordCount = iRecordCount + 1                                                    '查询记录数
            '移动记录游标
            SSProcess.AccessMoveNext mdbName, sql
        WEnd
    End If
    '关闭记录集
    SSProcess.CloseAccessRecordset mdbName, sql
    SSProcess.CloseAccessMdb mdbName
End Function