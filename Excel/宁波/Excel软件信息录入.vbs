
'Excel变量
Dim xlApp,xlFile,xlsheet

'用地红线GUID 和 宗地代码
Dim YDHXGUID
ZDCode = "9410001"
'建设工程规划许可证GUID
Dim JSGHXKZGUID

'幢的ID
Dim DTid(10000000)
Sub OnClick()
    
    '选取宗地
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    'SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
    SSProcess.SetSelectCondition "SSObj_Code", "=", ZDCode
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount
    If geoCount = 0 Then
        GetZDID = 0
        Exit Sub
    ElseIf geoCount = 1 Then
        ZDID = SSProcess.GetSelGeoValue (0, "SSObj_ID")
        YDHXGUID = SSProcess.GetSelGeoValue (0, "[YDHXGUID]")
        If YDHXGUID = "{00000000-0000-0000-0000-000000000000}"  Then
            YDHXGUID = GenNewGUID
            SSProcess.SetObjectAttr ZDID, "[YDHXGUID]", YDHXGUID
        End If
    Else
        MsgBox "图上有多个地!"
        Exit Sub
    End If
    aa = MsgBox("将覆盖已有数据，是否导入信息？",4 + 64)'是6 否7
    If aa = 7 Then  Exit Sub
    
    '打开Excel表格
    ExcelFile = SSProcess.SelectFileName(1,"选择excel文件",0,"EXCEL Files(*.xlsx)|*.xlsx|EXCEL Files(*.xls)|*.xls|All Files (*.*)|*.*||")
    If ExcelFile = "" Then Exit Sub
    Set xlApp = CreateObject("Excel.Application")
    Set xlFile = xlApp.Workbooks.Open(ExcelFile)
    
    GZLTJ()
    RJPZ()
    RYXX()
    XMKZCLCG()
    YQSB()
    xlApp.quit
End Sub

'属性表名称
Table_GZLTJ = "INFO_GZLTJ"
Table_RJPZ = "INFO_RJPZ"
Table_RYXX = "INFO_RYXX"
Table_XMKZCLCG = "INFO_XMKZCLCG"
Table_YQSB = "INFO_YQSB"

'工作统计量信息添加
Function GZLTJ()
    EmptyGZLTJInfo()
    Set xlsheet = xlFile.Worksheets("工作量统计表")
    xlsheet.Activate
    gztjlxx = ""
    excelhs = xlApp.ActiveSheet.UsedRange.Rows.Count
    For i = 2 To excelhs
        If xlApp.Cells(i,1) <> "" Then
            For k = 1 To 5
                str = xlApp.Cells(i,k)
                If gztjlxx = ""  Then
                    gztjlxx = str
                ElseIf k = 1 Then
                    gztjlxx = gztjlxx & str
                ElseIf k = 5 Then
                    gztjlxx = gztjlxx & "," & str & ";"
                Else
                    gztjlxx = gztjlxx & "," & str
                End If
            Next
        End If
    Next
    Infile = "YDHXGUID,序号,工作内容,工作量,工作量单位,备注"
    Sql = "select " & Infile & " from " & Table_GZLTJ & " where ID>0"
    Dim arr(100000),Count,Info
    SSFunc.ScanString gztjlxx, ";", arr, Count
    For y = 0 To Count - 2
        Values = YDHXGUID & "," & arr(y)
        InsertInfo Sql,Infile,Values
    Next
End Function

'软件配置信息添加
Function RJPZ()
    EmptyRJPZInfo()
    Set xlsheet = xlFile.Worksheets("软件配置表")
    xlsheet.Activate
    rjpzxx = ""
    excelhs = xlApp.ActiveSheet.UsedRange.Rows.Count
    For i = 2 To excelhs
        If xlApp.Cells(i,1) <> "" Then
            For k = 1 To 3
                str = xlApp.Cells(i,k)
                If rjpzxx = ""  Then
                    rjpzxx = str
                ElseIf k = 1 Then
                    rjpzxx = rjpzxx & str
                ElseIf k = 3 Then
                    rjpzxx = rjpzxx & "," & str & ";"
                Else
                    rjpzxx = rjpzxx & "," & str
                End If
            Next
        End If
    Next
    
    Infile = "YDHXGUID,序号,软件名称,软件用途"
    Sql = "select " & Infile & " from " & Table_RJPZ & " where ID>0"
    Dim arr(100000),Count,Info
    SSFunc.ScanString rjpzxx, ";", arr, Count
    For y = 0 To Count - 2
        Values = YDHXGUID & "," & arr(y)
        InsertInfo Sql,Infile,Values
    Next
End Function

'人员信息表添加
Function RYXX()
    EmptyRYXXInfo()
    Set xlsheet = xlFile.Worksheets("人员信息表")
    xlsheet.Activate
    ryxxxx = ""
    excelhs = xlApp.ActiveSheet.UsedRange.Rows.Count
    For i = 2 To excelhs
        If xlApp.Cells(i,1) <> "" Then
            For k = 1 To 6
                str = xlApp.Cells(i,k)
                If ryxxxx = ""  Then
                    ryxxxx = str
                ElseIf k = 1 Then
                    ryxxxx = ryxxxx & str
                ElseIf k = 6 Then
                    ryxxxx = ryxxxx & "," & str & ";"
                Else
                    ryxxxx = ryxxxx & "," & str
                End If
            Next
        End If
    Next
    
    Infile = "YDHXGUID,序号,姓名,职称或职业资格,上岗证书编号或职业资格证书号,主要工作职责,备注"
    Sql = "select " & Infile & " from " & Table_RYXX & " where ID>0"
    Dim arr(100000),Count,Info
    SSFunc.ScanString ryxxxx, ";", arr, Count
    For y = 0 To Count - 2
        
        Values = YDHXGUID & "," & arr(y)
        InsertInfo Sql,Infile,Values
    Next
End Function

'人员信息表添加
Function RYXX()
    EmptyRYXXInfo()
    Set xlsheet = xlFile.Worksheets("人员信息表")
    xlsheet.Activate
    ryxxxx = ""
    excelhs = xlApp.ActiveSheet.UsedRange.Rows.Count
    For i = 2 To excelhs
        If xlApp.Cells(i,1) <> "" Then
            For k = 1 To 6
                str = xlApp.Cells(i,k)
                If ryxxxx = ""  Then
                    ryxxxx = str
                ElseIf k = 1 Then
                    ryxxxx = ryxxxx & str
                ElseIf k = 6 Then
                    ryxxxx = ryxxxx & "," & str & ";"
                Else
                    ryxxxx = ryxxxx & "," & str
                End If
            Next
        End If
    Next
    
    Infile = "YDHXGUID,序号,姓名,职称或职业资格,上岗证书编号或职业资格证书号,主要工作职责,备注"
    Sql = "select " & Infile & " from " & Table_RYXX & " where ID>0"
    Dim arr(100000),Count,Info
    SSFunc.ScanString ryxxxx, ";", arr, Count
    For y = 0 To Count - 2
        FeatureGUID = GenNewGUID
        Values = YDHXGUID & "," & arr(y)
        InsertInfo Sql,Infile,Values
    Next
End Function

'项目控制测量成果添加
Function XMKZCLCG()
    EmptyXMKZCLCGInfo()
    Set xlsheet = xlFile.Worksheets("项目控制测量成果")
    xlsheet.Activate
    xmkzclcgxx = ""
    excelhs = xlApp.ActiveSheet.UsedRange.Rows.Count
    For i = 2 To excelhs
        If xlApp.Cells(i,1) <> "" Then
            For k = 1 To 6
                str = xlApp.Cells(i,k)
                If xmkzclcgxx = ""  Then
                    xmkzclcgxx = str
                ElseIf k = 1 Then
                    xmkzclcgxx = xmkzclcgxx & str
                ElseIf k = 6 Then
                    xmkzclcgxx = xmkzclcgxx & "," & str & ";"
                Else
                    xmkzclcgxx = xmkzclcgxx & "," & str
                End If
            Next
        End If
    Next
    
    Infile = "YDHXGUID,备注,高程H,平面坐标Y,平面坐标X,等级,点号"
    Sql = "select " & Infile & " from " & Table_XMKZCLCG & " where ID>0"
    Dim arr(100000),Count,Info
    SSFunc.ScanString xmkzclcgxx, ";", arr, Count
    For y = 0 To Count - 2
        FeatureGUID = GenNewGUID
        Values = YDHXGUID & "," & arr(y)
        InsertInfo Sql,Infile,Values
    Next
End Function

'仪器设备添加
Function YQSB()
    EmptyYQSBInfo()
    Set xlsheet = xlFile.Worksheets("仪器设备")
    xlsheet.Activate
    yqsbxx = ""
    excelhs = xlApp.ActiveSheet.UsedRange.Rows.Count
    For i = 2 To excelhs
        If xlApp.Cells(i,1) <> "" Then
            For k = 1 To 6
                str = xlApp.Cells(i,k)
                If yqsbxx = ""  Then
                    yqsbxx = str
                ElseIf k = 1 Then
                    yqsbxx = yqsbxx & str
                ElseIf k = 6 Then
                    yqsbxx = yqsbxx & "," & str & ";"
                Else
                    yqsbxx = yqsbxx & "," & str
                End If
            Next
        End If
    Next
    
    Infile = "YDHXGUID,序号,仪器名称,品牌型号,仪器编号,等级精度,仪器检定有效性"
    Sql = "select " & Infile & " from " & Table_YQSB & " where ID>0"
    Dim arr(100000),Count,Info
    SSFunc.ScanString yqsbxx, ";", arr, Count
    For y = 0 To Count - 2
        FeatureGUID = GenNewGUID
        Values = YDHXGUID & "," & arr(y)
        InsertInfo Sql,Infile,Values
    Next
End Function

'修改表信息
Function inAttr(sql,infile,invalues)
    ProjectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb ProjectName
    SSProcess.OpenAccessRecordset ProjectName, sql
    rscount = SSProcess.GetAccessRecordCount (ProjectName, sql)
    If rscount > 0 Then
        SSProcess.AccessMoveFirst ProjectName, sql
        While (SSProcess.AccessIsEOF (ProjectName, sql ) = False)
            SSProcess.ModifyAccessRecord  ProjectName, sql, infile , invalues'输出到mdb表中
            SSProcess.AccessMoveNext ProjectName, sql
        WEnd
    End If
    SSProcess.CloseAccessRecordset ProjectName, sql
    SSProcess.CloseAccessMdb ProjectName
End Function

'获取最新的FeatureGUID
Function GenNewGUID()
    Set TypeLib = CreateObject("Scriptlet.TypeLib")
    GenNewGUID = Left(TypeLib.Guid,38)
    Set TypeLib = Nothing
End Function

'********插入新纪录
Function InsertRecord( tableName, fields, values)
    sqlString = "insert into " & tableName & " (" & fields & ") values (" & values & ")"
    InsertRecord = SSProcess.ExecuteSql (sqlString)
End Function



'清空许可证表信息
Function EmptyRYXXInfo()
    sql = "SELECT * FROM INFO_RYXX where INFO_RYXX.ID > " & "0" & ";"
    mdbName = SSProcess.GetProjectFileName
    
    SSProcess.OpenAccessMdb mdbName
    SSProcess.OpenAccessRecordset mdbName, sql
    
    While  SSProcess.AccessIsEOF (mdbName, sql) = False
        SSProcess.DelAccessRecord mdbName, sql
    WEnd
    SSProcess.CloseAccessRecordset mdbName, sql
    SSProcess.CloseAccessMdb mdbName
End Function

'插入记录
Function InsertInfo(sql,Infile,Values)
    mdbName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb mdbName
    SSProcess.OpenAccessRecordset mdbName, sql
    recordc = SSProcess.GetAccessRecordCount(mdbName, sql)
    SSProcess.AddAccessRecord mdbName,sql,Infile,Values
    SSProcess.CloseAccessRecordset mdbName, sql
    SSProcess.CloseAccessMdb mdbName
End Function

'清空软件配置表数据
Function EmptyRJPZInfo()
    sql = "SELECT * FROM INFO_RJPZ where INFO_RJPZ.ID > " & "0" & ";"
    mdbName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb mdbName
    SSProcess.OpenAccessRecordset mdbName, sql
    
    While  SSProcess.AccessIsEOF (mdbName, sql) = False
        SSProcess.DelAccessRecord mdbName, sql
    WEnd
    SSProcess.CloseAccessRecordset mdbName, sql '关库
    SSProcess.CloseAccessMdb mdbName
End Function

'清空工作量统计表信息
Function EmptyGZLTJInfo()
    sql = "SELECT * FROM INFO_GZLTJ where INFO_GZLTJ.ID > " & "0" & ";"
    mdbName = SSProcess.GetProjectFileName
    
    SSProcess.OpenAccessMdb mdbName
    SSProcess.OpenAccessRecordset mdbName, sql  '打开数据库
    
    While  SSProcess.AccessIsEOF (mdbName, sql) = False
        SSProcess.DelAccessRecord mdbName, sql
    WEnd
    
    SSProcess.CloseAccessRecordset mdbName, sql '关库
    SSProcess.CloseAccessMdb mdbName
End Function

'清空项目控制测量成果
Function EmptyXMKZCLCGInfo()
    sql = "SELECT * FROM INFO_XMKZCLCG where INFO_XMKZCLCG.ID > " & "0" & ";"
    mdbName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb mdbName
    SSProcess.OpenAccessRecordset mdbName, sql
    
    While  SSProcess.AccessIsEOF (mdbName, sql) = False
        SSProcess.DelAccessRecord mdbName, sql
    WEnd
    SSProcess.CloseAccessRecordset mdbName, sql '关库
    SSProcess.CloseAccessMdb mdbName
End Function

'清空仪器设备
Function EmptyYQSBInfo()
    sql = "SELECT * FROM INFO_YQSB where INFO_YQSB.ID > " & "0" & ";"
    mdbName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb mdbName
    SSProcess.OpenAccessRecordset mdbName, sql
    
    While  SSProcess.AccessIsEOF (mdbName, sql) = False
        SSProcess.DelAccessRecord mdbName, sql
    WEnd
    SSProcess.CloseAccessRecordset mdbName, sql '关库
    SSProcess.CloseAccessMdb mdbName
End Function