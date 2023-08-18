'========================================================Excel操作对象和文件路径操作对象======================================================

'路径操作对象
Dim FileSysObj
Set FileSysObj = CreateObject("Scripting.FileSystemObject")

'Excel操作对象
Dim ExcelObj
Set ExcelObj = CreateObject("Excel.Application")

'============================================================检查记录配置====================================================================

'埋深阈值
Dim Threshold
Threshold = 5

'检查集项目名称
Dim strGroupName
strGroupName = "管线精度检查"

'检查集组名称
Dim strCheckName
strCheckName = "埋深精度检查"

'检查模型名称
Dim CheckmodelName
CheckmodelName = "自定义脚本检查类->埋深精度检查"

'检查描述
Dim strDescription
strDescription = "埋深精度超标"


'=============================================================功能入口=======================================================================

Sub OnClick()
    
    AllVisible
    
    JcrInfo DateTime,PersonName
    
    ExcelFilePath = SSProcess.SelectFileName(1,"选择Excel文件",0,"EXCEL Files(*.xlsx)|*.xlsx|EXCEL Files(*.xls)|*.xls|All Files (*.*)|*.*||")
    
    If  ExcelFilePath = "" Then
        MsgBox "未选择文件，已退出"
        Exit Sub
    End If
    
    FileName = Right(ExcelFilePath,Len(ExcelFilePath) - InStrRev(ExcelFilePath,"\"))
    FileSysObj.CopyFile  ExcelFilePath,SSProcess.GetSysPathName(5) & "测量精度\" & FileName
    
    ClearCheckRecord
    
    OpenExcel SSProcess.GetSysPathName(5) & FileName,ExcelFile
    
    SetTableHeader
    
    InsertExcel ExcelFile,4,GetEndRow - 1,ErrorIds
    
    GetMaxDeep ExcelFile,4,GetEndRow - 1
    
    SetFooter DateTime,PersonName,GetEndRow - 1
    
    CloseExcel ExcelFile
    
    AddRecord ErrorIds
    
    Ending
    
End Sub' OnClick

'=============================================================报表填值==============================================================

'输入检查人弹框
Function JcrInfo(ByRef DateTime,ByRef PersonName)
    SSProcess.ClearInputParameter
    SSProcess.AddInputParameter "检查人" , "" , 0 , "" , ""
    SSProcess.AddInputParameter "日期" , GetNowTime , 0 , "" , ""
    Result = SSProcess.ShowInputParameterDlg ("信息录入")
    If Result = 1 Then
        DateTime = SSProcess.GetInputParameter ("日期")
        PersonName = SSProcess.GetInputParameter ("检查人")
    End If
End Function' JcrInfo

'获取当前系统时间
Function GetNowTime()
    GetNowTime = CStr(FormatDateTime(Now(),1))
End Function' GetNowTime

'设置表名
Function SetTableHeader()
    SqlStr = "Select XMMC From 管线项目信息表 Where 管线项目信息表.ID = 1"
    GetSQLRecordAll SqlStr,Xmmc,Count
    ExcelObj.Cells(2,2) = Xmmc(0)
End Function' SetTableHeader

'获取最大行数
Function GetEndRow()
    GetEndRow = 4
    Poisition = InStr(ExcelObj.Cells(GetEndRow,1),"最大较差为：")
    Do While ExcelObj.Cells(GetEndRow,1) <> ""
        Poisition = InStr(ExcelObj.Cells(GetEndRow,1),"最大较差为：")
        'MsgBox Poisition
        If Poisition = 0 Then
            GetEndRow = GetEndRow + 1
        Else
            Exit Function
        End If
    Loop
End Function' GetEndRow

Function InsertExcel(ByVal ExcelFile,ByVal StartRow,ByVal EndRow,ByRef ErrorIds)
    ErrorIds = ""
    For i = StartRow To EndRow
        SqlStr = "Select 地下管线线属性表.ID,地下管线线属性表.GXQDMS,地下管线线属性表.GXQDDH From 地下管线线属性表 inner join GeoLineTB on 地下管线线属性表.ID = GeoLineTB.ID WHERE (GeoLineTB.Mark Mod 2)<>0 And GXQDDH = " & AddDenote(ExcelObj.Cells(i,1)) & " And GXZDDH = " & AddDenote(ExcelObj.Cells(i,2))
        GetSQLRecordAll SqlStr,DeepArr,DeepCount
        If DeepCount >= 1 Then
            TempArr = Split(DeepArr(0),",", - 1,1)
            SqlStr = "Select 地下管线点属性表.FSW From 地下管线点属性表 inner join GeoPointTB on 地下管线点属性表.ID = GeoPointTB.ID WHERE (GeoPointTB.Mark Mod 2)<>0 And 地下管线点属性表.WTDH = " & "'" & TempArr(2) & "'"
            GetSQLRecordAll SqlStr,PoiFWSArr,Count
            ExcelObj.Cells(i,3) = Round(Transform(TempArr(1)) * 100,2)
            Diff = Abs(Round(Transform(TempArr(1)) * 100 - Transform(ExcelObj.Cells(i,4)),2))
            ExcelObj.Cells(i,5) = Diff
            If PoiFWSArr(0) = "" Then
                If Diff > 0.15 * Transform(TempArr(1)) * 100 Then
                    ExcelObj.Cells(i,6) = "埋深超限"
                    If ErrorIds = "" Then
                        ErrorIds = TempArr(0)
                    Else
                        ErrorIds = ErrorIds & "," & TempArr(0)
                    End If
                End If
            ElseIf PoiFWSArr(0) = "*" Then
                If Diff > 0.15 * Transform(TempArr(1)) * 100 Then
                    ExcelObj.Cells(i,6) = "埋深超限"
                    If ErrorIds = "" Then
                        ErrorIds = TempArr(0)
                    Else
                        ErrorIds = ErrorIds & "," & TempArr(0)
                    End If
                End If
            ElseIf PoiFWSArr(0) = Null Then
                If Diff > 0.15 * Transform(TempArr(1)) * 100 Then
                    ExcelObj.Cells(i,6) = "埋深超限"
                    If ErrorIds = "" Then
                        ErrorIds = TempArr(0)
                    Else
                        ErrorIds = ErrorIds & "," & TempArr(0)
                    End If
                End If
            Else
                If Diff > Threshold Then
                    ExcelObj.Cells(i,6) = "埋深超限"
                    If ErrorIds = "" Then
                        ErrorIds = TempArr(0)
                    Else
                        ErrorIds = ErrorIds & "," & TempArr(0)
                    End If
                End If
            End If
        End If
    Next 'i
End Function' InsertExcel

'获取最大深度差
Function GetMaxDeep(ByVal ExcelFile,ByVal StartRow,ByVal EndRow)
    GetMaxDeep = 0
    For i = StartRow To EndRow
        If GetMaxDeep < Transform(ExcelObj.Cells(i,5)) Then
            GetMaxDeep = Transform(ExcelObj.Cells(i,5))
        End If
    Next 'i
    ExcelObj.Cells(EndRow + 1,1) = "最大较差为：" & GetMaxDeep & "CM"
    TIANzdjc "MSZDJC",GetMaxDeep
End Function' GetMaxDeep

'填写检查人和日期
Function SetFooter(ByVal DateTime,ByVal PersonName,ByVal EndRow)
    ExcelObj.Cells(EndRow + 2,2) = PersonName
    ExcelObj.Cells(EndRow + 2,6) = DateTime
End Function' SetFooter

'添加检查记录
Function AddRecord(ByVal ErrorIds)
    If ErrorIds <> "" Then
        ErrorArr = Split(ErrorIds,",", - 1,1)
        For i = 0 To UBound(ErrorArr)
            SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(ErrorArr(i),"SSObj_X"),SSProcess.GetObjectAttr(ErrorArr(i),"SSObj_Y"),0,1,ErrorArr(i),""
        Next 'i
    End If
    SSProcess.ShowCheckOutput
    SSProcess.SaveCheckRecord
End Function' AddRecord

'埋深检查最大值填入信息表
Function TIANzdjc (zdmc,zdjc)
    mdbName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb mdbName
    sql = "update  管线项目信息表 set " & zdmc & " = " & zdjc & " where 管线项目信息表.ID= 1"
    SSProcess.ExecuteAccessSql  mdbName,sql
    SSProcess.CloseAccessMdb mdbName
End Function

'==========================================================工具类函数====================================================================

'打开所有图层
Function AllVisible()
    count = SSProcess.GetLayerCount
    For i = 0 To count - 1
        layername = SSProcess.GetLayerName (i)
        SSProcess.SetLayerStatus layername, 1, 1
    Next
    SSProcess.RefreshView
End Function

'数据类型转换
Function Transform(ByVal Values)
    If Values <> "" Then
        Values = CDbl(Values)
    Else
        Values = 0
        Exit Function
    End If
    Transform = Values
End Function'Transform

'打开Excel表
Function OpenExcel(ByVal FilePath,ByRef ExcleFile)
    ExcelObj.Application.Visible = False
    Set ExcleFile = ExcelObj.WorkBooks.Open(FilePath)
    Set ExcelSheet = ExcleFile.WorkSheets(1)
    ExcelSheet.Activate
End Function

'获取Excel名称
Function GetExcelName(ByVal ExcelFilePath)
    ExcelFilePathArr = Split(ExcelFilePath,"\", - 1,1)
    GetExcelName = ExcelFilePathArr(UBound(ExcelFilePathArr))
End Function' GetExcelName

'保存关闭Excel表格
Function CloseExcel(ByVal ExcelFile)
    ExcelFile.Save
    ExcelFile.Close
    ExcelObj.Quit
End Function' CloseExcel

'获取所有记录
Function GetSQLRecordAll(ByVal StrSqlStatement, ByRef SQLRecord(), ByRef iRecordCount)
    SSProcess.OpenAccessMdb SSProcess.GetProjectFileName
    If StrSqlStatement = "" Then
        MsgBox "查询语句为空，操作停止！",48
    End If
    iRecordCount =  - 1
    SSProcess.OpenAccessRecordset SSProcess.GetProjectFileName, StrSqlStatement
    RecordCount = SSProcess.GetAccessRecordCount (SSProcess.GetProjectFileName, StrSqlStatement)
    If RecordCount > 0 Then
        iRecordCount = 0
        ReDim SQLRecord(RecordCount)
        SSProcess.AccessMoveFirst SSProcess.GetProjectFileName, StrSqlStatement
        iRecordCount = 0
        While SSProcess.AccessIsEOF (SSProcess.GetProjectFileName, StrSqlStatement) = 0
            fields = ""
            values = ""
            SSProcess.GetAccessRecord SSProcess.GetProjectFileName, StrSqlStatement, fields, values
            SQLRecord(iRecordCount) = values
            iRecordCount = iRecordCount + 1
            SSProcess.AccessMoveNext SSProcess.GetProjectFileName, StrSqlStatement
        WEnd
    End If
    SSProcess.CloseAccessRecordset SSProcess.GetProjectFileName, StrSqlStatement
    SSProcess.CloseAccessMdb SSProcess.GetProjectFileName
End Function

'添加单引号
Function AddDenote(ByVal Value)
    AddDenote = "'" & Value & "'"
End Function' AddDenote

'清空检查集
Function ClearCheckRecord()
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
End Function' ClearCheckRecord

'结束提示
Function Ending()
    MsgBox "输出完成"
End Function' Ending