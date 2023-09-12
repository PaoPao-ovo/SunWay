
' 1、将模板复制到当前项目路径下
' 2、通过SQL查询地下管线点属性表的【物探点号】与DEAFULT图层【0】点的点名相同的点

' [管线报告信息]
' 编号 = ""
' 项目名称 = ""
' 项目地址 = ""
' 设计单位 = ""
' 建设单位 = ""
' 委托单位 = ""
' 外业时间 = ""
' 测绘时间 = ""
' 点最大较差值 = ""
' 高程最大较差值 = ""

'========================================================Excel操作对象和文件路径操作对象======================================================

'路径操作对象
Dim FileSysObj
Set FileSysObj = CreateObject("Scripting.FileSystemObject")

'Excel操作对象
Dim ExcelObj
Set ExcelObj = CreateObject("Excel.Application")

'============================================================检查记录配置====================================================================

'检查集项目名称
Dim strGroupName
strGroupName = "管线精度检查"

'检查集组名称
Dim strCheckName
strCheckName = "坐标精度检查"

'检查模型名称
Dim CheckmodelName
CheckmodelName = "自定义脚本检查类->坐标精度检查"

'检查描述
Dim strDescription
strDescription = "坐标精度超标"

'=============================================================功能入口=======================================================================

Sub OnClick()
    
    AllVisible
    
    FileSysObj.CopyFile  SSProcess.GetSysPathName (7) & "输出模板\" & "测量精度调查表模板.xlsx",SSProcess.GetSysPathName(5) & "测量精度调查表.xlsx"
    
    OpenExcel SSProcess.GetSysPathName(5) & "测量精度调查表.xlsx",ExcleFile
    
    GetPoiNameZero ZeorCount,ZeroArr,ZeroStr
    
    GetGXDDH GXDDHArr,DhLxCount,ZeroStr
    
    InsertExcle 3,EndRow,ZeorCount,ZeroArr,GXDDHArr,DhLxCount
    
    InsertDiff 3,EndRow
    
    GetResultInfo ZeorCount,3,EndRow,OverPoiCount,OverHeightCount,AveragePoi,AverageHeight,MiddlePoi,MiddleHei,OverPercent,OverPoiName,OverHeiName,MaxLen,MaxHei
    
    GXEWB MaxLen,MaxHei

    InsertResult EndRow + 1,ZeorCount,OverPoiCount,OverHeightCount,AveragePoi,AverageHeight,MiddlePoi,MiddleHei,OverPercent
    
    DelSelCol 6
    
    CloseExcel ExcleFile
    
    ClearCheckRecord
    
    AddRecord OverPoiName,OverHeiName
    
End Sub' OnClick

'==========================================================Excel填值===================================================================

'填写统计结果
Function InsertResult(ByVal ResultRow,ByVal ZeroCount,ByVal OverPoiCount,ByVal OverHeightCount,ByVal AveragePoi,ByVal AverageHeight,ByVal MiddlePoi,ByVal MiddleHei,ByVal OverPercent)
    If ZeroCount < 15 Then
        If OverPoiCount > 0 And OverHeightCount > 0 Then
            ResultString = "统计结果：检查点数：" & ZeroCount & "个，点位超差的点数" & OverPoiCount & "个，高程超差的个数" & OverHeightCount & "个，超差点位占总点数百分比" & OverPercent & "%，点位误差平均值：" & AveragePoi & "（M），高程误差平均值：" & AverageHeight & "（M）"
            ExcelObj.Cells(ResultRow,1) = ResultString
        ElseIf OverPoiCount > 0 And OverHeightCount = 0 Then
            ResultString = "统计结果：检查点数：" & ZeroCount & "个，点位超差的点数" & OverPoiCount & "个，超差点位占总点数百分比" & OverPercent & "%，点位误差平均值：" & AveragePoi & "（M），高程误差平均值：" & AverageHeight & "（M）"
            ExcelObj.Cells(ResultRow,1) = ResultString
        ElseIf OverPoiCount = 0 And OverHeightCount > 0 Then
            ResultString = "统计结果：检查点数：" & ZeroCount & "个" & "高程超差的个数" & OverHeightCount & "个，超差点位占总点数百分比" & OverPercent & "%，点位误差平均值：" & AveragePoi & "（M），高程误差平均值：" & AverageHeight & "（M）"
            ExcelObj.Cells(ResultRow,1) = ResultString
        Else
            ResultString = "统计结果：检查点数：" & ZeroCount & "个" & "点位误差平均值：" & AveragePoi & "（M），高程误差平均值：" & AverageHeight & "（M）"
            ExcelObj.Cells(ResultRow,1) = ResultString
        End If
    ElseIf ZeroCount >= 15 Then
        If OverPoiCount > 0 And OverHeightCount > 0 Then
            ResultString = "统计结果：检查点数：" & ZeroCount & "个，点位超差的点数" & OverPoiCount & "个，高程超差的个数" & OverHeightCount & "个，超差点位占总点数百分比" & OverPercent & "%，点位误差中误差：" & MiddlePoi & "（M），高程误差中误差：" & MiddleHei & "（M）"
            ExcelObj.Cells(ResultRow,1) = ResultString
        ElseIf OverPoiCount > 0 And OverHeightCount = 0 Then
            ResultString = "统计结果：检查点数：" & ZeroCount & "个，点位超差的点数" & OverPoiCount & "个，超差点位占总点数百分比" & OverPercent & "%，点位误差中误差：" & MiddlePoi & "（M），高程误差中误差：" & MiddleHei & "（M）"
            ExcelObj.Cells(ResultRow,1) = ResultString
        ElseIf OverPoiCount = 0 And OverHeightCount > 0 Then
            ResultString = "统计结果：检查点数：" & ZeroCount & "个，高程超差的个数" & OverHeightCount & "个，超差点位占总点数百分比" & OverPercent & "%，点位误差中误差：" & MiddlePoi & "（M），高程误差中误差：" & MiddleHei & "（M）"
            ExcelObj.Cells(ResultRow,1) = ResultString
        Else
            ResultString = "统计结果：检查点数：" & ZeroCount & "个" & "点位误差中误差：" & MiddlePoi & "（M），高程误差中误差：" & MiddleHei & "（M）"
            ExcelObj.Cells(ResultRow,1) = ResultString
        End If
    End If
End Function' InsertResult

'获取结果信息
Function GetResultInfo(ByVal ZeroCount,ByVal StartRow,ByVal EndRow,ByRef OverPoiCount,ByRef OverHeightCount,ByRef AveragePoi,ByRef AverageHeight,ByRef MiddlePoi,ByRef MiddleHei,ByRef OverPercent,ByRef OverPoiName,ByRef OverHeiName,ByRef MaxLen,ByRef MaxHei)
    OverPoiCount = 0
    OverHeightCount = 0
    OverNum = 0
    AveragePoi = 0.00
    AverageHeight = 0.00
    SquarePoi = 0
    SquareHei = 0
    MiddlePoi = 0.00
    MiddleHei = 0.00
    OverPoiName = ""
    OverHeiName = ""
    MaxLen = 0.00
    MaxHei = 0.00
    For i = StartRow To EndRow
        If Transform(ExcelObj.Cells(i,10)) > 0.05 Then
            OverPoiCount = OverPoiCount + 1
            If OverPoiName = "" Then
                OverPoiName = "'" & ExcelObj.Cells(i,2) & "'"
            Else
                OverPoiName = OverPoiName & "," & "'" & ExcelObj.Cells(i,2) & "'"
            End If
        End If
        If Transform(ExcelObj.Cells(i,10)) > MaxLen Then
            MaxLen = Transform(ExcelObj.Cells(i,10))
        End If
        If Transform(ExcelObj.Cells(i,11)) > MaxHei Then
            MaxHei = Transform(ExcelObj.Cells(i,10))
        End If
        If Transform(ExcelObj.Cells(i,11)) > 0.03 Then
            OverHeightCount = OverHeightCount + 1
            If OverHeiName = "" Then
                OverHeiName = "'" & ExcelObj.Cells(i,2) & "'"
            Else
                OverHeiName = OverHeiName & "," & "'" & ExcelObj.Cells(i,2) & "'"
            End If
        End If
        TotalPoi = TotalPoi + Transform(ExcelObj.Cells(i,10))
        TotalHei = TotalHei + Transform(ExcelObj.Cells(i,11))
        SquarePoi = SquarePoi + Transform(ExcelObj.Cells(i,10)) ^ 2
        SquareHei = SquareHei + Transform(ExcelObj.Cells(i,11)) ^ 2
        If Transform(ExcelObj.Cells(i,10)) > 0.05 Or Transform(ExcelObj.Cells(i,11)) > 0.03 Then
            OverNum = OverNum + 1
        End If
    Next 'i
    OverPercent = Round((OverNum / ZeroCount),2) * 100
    AveragePoi = Round(TotalPoi / ZeroCount,3)
    AverageHeight = Round(TotalHei / ZeroCount,3)
    MiddlePoi = Round(Sqr(SquarePoi / ZeroCount),3)
    MiddleHei = Round(Sqr(SquareHei / ZeroCount),3)
End Function' GetResultInfo

'获取长度差
Function GetDiffS(x1,y1,x2,y2)
    GetDiffS = Round(Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2),3)
End Function' GetDiffS

'填写高程差和距离差
Function InsertDiff(ByVal StartRow,ByVal EndRow)
    For i = StartRow To EndRow
        ExcelObj.Cells(i,10) = GetDiffS(Transform(ExcelObj.Cells(i,3)),Transform(ExcelObj.Cells(i,4)),Transform(ExcelObj.Cells(i,7)),Transform(ExcelObj.Cells(i,8)))
        ExcelObj.Cells(i,11) = Round(Abs(Transform(ExcelObj.Cells(i,5)) - Transform(ExcelObj.Cells(i,9))),3)
    Next 'i
End Function' InsertDiff

'填写Excel
Function InsertExcle(ByVal StartRow,ByRef EndRow,ByVal ZeroCount,ByVal ZeroArr(),ByVal GXDDHArr(),ByVal DhLxCount)
    EndRow = ZeroCount + 2
    If EndRow <= 5  Then
        For i = StartRow To EndRow
            ExcelObj.Cells(i,1) = i - 2
            ExcelObj.Cells(i,6) = ZeroArr(i - StartRow,0)
            InsertZeroXYZ i,i - StartRow,ZeroArr
            InsertGxPoint StartRow,EndRow,GXDDHArr,DhLxCount
        Next 'i
    ElseIf EndRow > 5 Then
        InsertRows StartRow,ZeroCount - 3
        For i = StartRow To EndRow
            ExcelObj.Cells(i,1) = i - 2
            ExcelObj.Cells(i,6) = ZeroArr(i - StartRow,0)
            InsertZeroXYZ i,i - StartRow,ZeroArr
            InsertGxPoint StartRow,EndRow,GXDDHArr,DhLxCount
        Next 'i
    End If
End Function' InsertExcle

'填写管线点值
Function InsertGxPoint(ByVal StartRow,ByVal EndRow,ByVal GXDDHArr,ByVal DhLxCount)
    For i = 0 To DhLxCount
        For j = StartRow To EndRow
            If ExcelObj.Cells(j,6) = SSProcess.GetObjectAttr(GXDDHArr(i),"[WTDH]") Then
                ExcelObj.Cells(j,2) = SSProcess.GetObjectAttr(GXDDHArr(i),"[WTDH]")
                ExcelObj.Cells(j,3) = Round(Transform(SSProcess.GetObjectAttr(GXDDHArr(i),"SSObj_Y")),3)
                ExcelObj.Cells(j,4) = Round(Transform(SSProcess.GetObjectAttr(GXDDHArr(i),"SSObj_X")),3)
                ExcelObj.Cells(j,5) = Round(Transform(SSProcess.GetObjectAttr(GXDDHArr(i),"SSObj_Z")),3)
            End If
        Next 'j
    Next 'i
End Function' InsertGxPoint

'填写0点的XYZ值到Excel表格中
Function InsertZeroXYZ(ByVal InsertRow,ByVal Index,ByVal ZeroArr())
    ExcelObj.Cells(InsertRow,7) = Round(Transform(SSProcess.GetObjectAttr(ZeroArr(Index,1),"SSObj_Y")),3)
    ExcelObj.Cells(InsertRow,8) = Round(Transform(SSProcess.GetObjectAttr(ZeroArr(Index,1),"SSObj_X")),3)
    ExcelObj.Cells(InsertRow,9) = Round(Transform(SSProcess.GetObjectAttr(ZeroArr(Index,1),"SSObj_Z")),3)
End Function' InsertZeroXYZ

'获取所有的0点的点名和ID
Function GetPoiNameZero(ByRef ZeorCount,ByRef ZeroArr(),ByRef ZeroStr)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "0"
    SSProcess.SetSelectCondition "SSObj_PointName", "<>", ""
    SSProcess.SelectFilter
    ZeorCount = SSProcess.GetSelGeoCount
    ReDim  ZeroArr(ZeorCount - 1,1)
    For i = 0 To ZeorCount - 1
        ZeroArr(i,0) = SSProcess.GetSelGeoValue(i,"SSObj_PointName")
        ZeroArr(i,1) = SSProcess.GetSelGeoValue(i,"SSObj_ID")
        If ZeroStr = "" Then
            ZeroStr = "'" & SSProcess.GetSelGeoValue(i,"SSObj_PointName") & "'"
        Else
            ZeroStr = ZeroStr & "," & "'" & SSProcess.GetSelGeoValue(i,"SSObj_PointName") & "'"
        End If
    Next 'i
End Function' GetPoiName_0

Function Transform(ByVal Values)
    If Values <> "" Then
        Values = CDbl(Values)
    Else
        Values = 0
        Exit Function
    End If
    Transform = Values
End Function'Transform

'获取所有满足条件的物探点号的ID
Function GetGXDDH(ByRef GXDDHArr(),ByRef DhLxCount,ByVal ZeoStr)
    SqlString = "Select 地下管线点属性表.ID From 地下管线点属性表 Inner Join GeoPointTB on 地下管线点属性表.ID = GeoPointTB.ID WHERE (GeoPointTB.Mark Mod 2)<>0 And 地下管线点属性表.WTDH In " & "(" & ZeoStr & ")"
    GetSQLRecordAll SqlString,GXDDHArr,DhLxCount
End Function' GetGXDDH

'添加检查记录
Function AddRecord(ByVal OverPoiName,ByVal OverHeiName)
    If OverPoiName <> "" Then
        SqlString = "Select 地下管线点属性表.ID From 地下管线点属性表 Inner Join GeoPointTB on 地下管线点属性表.ID = GeoPointTB.ID WHERE (GeoPointTB.Mark Mod 2)<>0 And 地下管线点属性表.WTDH In " & "(" & OverPoiName & ")"
        GetSQLRecordAll SqlString,OverPoiArr,OverPoiCount
        For i = 0 To OverPoiCount - 1
            SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(OverPoiArr(i),"SSObj_X"),SSProcess.GetObjectAttr(OverPoiArr(i),"SSObj_Y"),0,0,OverPoiArr(i),""
        Next 'i
    End If
    strCheckName = "高程精度检查"
    CheckmodelName = "自定义脚本检查类->高程精度检查"
    strDescription = "高程精度超标"
    If OverHeiName <> "" Then
        SqlString = "Select 地下管线点属性表.ID From 地下管线点属性表 Inner Join GeoPointTB on 地下管线点属性表.ID = GeoPointTB.ID WHERE (GeoPointTB.Mark Mod 2)<>0 And 地下管线点属性表.WTDH In " & "(" & OverHeiName & ")"
        GetSQLRecordAll SqlString,OverHeiArr,OverHeiCount
        For i = 0 To OverHeiCount - 1
            SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(OverHeiArr(i),"SSObj_X"),SSProcess.GetObjectAttr(OverHeiArr(i),"SSObj_Y"),0,0,OverHeiArr(i),""
        Next 'i
    End If
    SSProcess.ShowCheckOutput
    SSProcess.SaveCheckRecord
End Function' AddRecord

'清空检查集
Function ClearCheckRecord()
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
End Function' ClearCheckRecord

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

'插入指定行数
Function InsertRows(ByVal StartRow,ByVal InsertCount)
    For i = 0 To InsertCount - 1
        ExcelObj.ActiveSheet.Rows(StartRow).Insert
    Next 'i
End Function' InsertRows

'打开Excel表
Function OpenExcel(ByVal FilePath,ByRef ExcleFile)
    ExcelObj.Application.Visible = True
    Set ExcleFile = ExcelObj.WorkBooks.Open(FilePath)
    Set ExcelSheet = ExcleFile.WorkSheets("精度调查表")
    ExcelSheet.Activate
End Function

'打开所有图层
Function AllVisible()
    count = SSProcess.GetLayerCount
    For i = 0 To count - 1
        layername = SSProcess.GetLayerName (i)
        SSProcess.SetLayerStatus layername, 1, 1
    Next
    SSProcess.RefreshView
End Function

'保存关闭Excel表格
Function CloseExcel(ByVal ExcleFile)
    ExcleFile.Save
    ExcelObj.Quit
End Function' CloseExcel

'删除指定的列
Function DelSelCol(ByVal ColNum)
    ExcelObj.ActiveSheet.Columns(ColNum).Delete
End Function' DelSelCol

'刷新二维表
Function GXEWB(DWZDJC,GCZDJC)
    mdbName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb mdbName
    sql = "update  管线项目信息表 set DWZDJC = " & DWZDJC & "where 管线项目信息表.ID= 1"
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  管线项目信息表 set GCZDJC = " & GCZDJC & "where 管线项目信息表.ID= 1"
    SSProcess.ExecuteAccessSql  mdbName,sql
    SSProcess.CloseAccessMdb mdbName
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
End Function