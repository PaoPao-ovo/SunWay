
'=============================================检查集配置==============================================

'检查组项目名称
Dim strGroupName
strGroupName = "管线检查"

'检查集项名称
Dim strCheckName
strCheckName = "管径检查"

'检查模型名称
Dim CheckmodelName
CheckmodelName = "自定义脚本检查类->管径检查"

'检查描述
Dim strDescription
strDescription = "管径超高"

'===========================================功能入口========================================================

'总入口
Sub OnClick()
    
    ClearCheckRecord '清空原来的检查记录
    
    GetErrorLines LineIds '错误管线的ID
    
    AddRecords LineIds '添加检查记录
    
End Sub' OnClick

'返回错误管线ID
Function GetErrorLines(ByRef LineIds)
    
    EorrorCount = 0
    
    ReDim LineIds(EorrorCount)
    
    
    SqlStr = "Select 地下管线线属性表.ID,GXQDMS,GXZDMS,GJ From 地下管线线属性表 Inner Join GeoLineTB on 地下管线线属性表.ID = GeoLineTB.ID Where (GeoLineTB.Mark Mod 2)<>0"
    
    GetSQLRecordAll SqlStr,LineArr,LineCount
    
    For i = 0 To LineCount - 1
        LineAttrArr = Split(LineArr(i),",", - 1,1) '0=ID,1=GXQDMS,2=GXZDMS,3=GJ
        If InStr(LineAttrArr(3),"*") <> 0 Then
            GJArr = Split(LineAttrArr(3),"*", - 1,1) '0=Length 1=Width
            Length = Transform(GJArr(0)) / 1000
            Width = Transform(GJArr(1)) / 1000
            If Length > Width Then
                CompareGJ = Length
            Else
                CompareGJ = Width
            End If
        ElseIf LineAttrArr(3) <> "" Then
            CompareGJ = Transform(LineAttrArr(3)) / 1000
        End If
        GXQDMS = Transform(LineAttrArr(1))
        GXZDMS = Transform(LineAttrArr(2))
        If CompareGJ > GXQDMS Or CompareGJ > GXZDMS Then
            LineIds(EorrorCount) = LineAttrArr(0)
            EorrorCount = EorrorCount + 1
            ReDim Preserve LineIds(EorrorCount)
        End If
    Next 'i
End Function' GetErrorLines

'添加检查记录
Function AddRecords(ByVal LineIds())
    For i = 0 To UBound(LineIds) - 1
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(LineIds(i),"SSObj_X"),SSProcess.GetObjectAttr(LineIds(i),"SSObj_Y"),0,1,LineIds(i),""
    Next 'i
    SSProcess.ShowCheckOutput
    SSProcess.SaveCheckRecord
End Function' AddRecords

'================================================工具类函数===========================================

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