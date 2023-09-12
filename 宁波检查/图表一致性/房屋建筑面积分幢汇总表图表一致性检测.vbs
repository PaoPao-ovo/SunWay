
'==========================================================检查配置============================================================

'检查集项目名称
Dim strGroupName
strGroupName = "房屋建筑面积分幢汇总表"

'检查集组名称
Dim strCheckName
strCheckName = "图表一致性检查"

'检查模型名称
Dim CheckmodelName
CheckmodelName = "自定义脚本检查类->图表一致性检查"

'检查描述
Dim strDescription
strDescription = "房屋建筑面积分幢汇总表,户室房屋建筑面积分层汇总表,自然幢中【BSM】在户【ZRZBSM】中找不到相同的值"

'================================================================检查表配置======================================================

'自然幢属性表
Dim FxTable
FxTable = "FC_自然幢信息属性表"

'放线实测点属性表
Dim RealTable
RealTable = "H"

'=============================================================功能入口=======================================================================

Sub OnClick()
    
    AddRecordInner
    
    
End Sub' OnClick

'=============================================================点号字段判断并添加检查记录================================================

'添加检查记录入口
Function AddRecordInner()
    ClearCheckRecord
    FxPoiInfo FxDhArr,DhCount
    ConfirmScPoi FxDhArr,DhCount
    ShowCheckRecord
End Function' AddRecordInner

'获取放线桩点的点号
Function FxPoiInfo(ByRef FxDhArr(),ByRef DhCount)
    SqlStr = "Select " & FxTable & ".BSM," & FxTable & ".ID" & " From " & FxTable & " Inner Join GeoAreaTB on " & FxTable & ".ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0 And " & FxTable & ".BSM Is Not Null "
    GetSQLRecordAll SqlStr,FxDhArr,DhCount
End Function' FxPoiInfo

'判断实测点是否存在
Function ConfirmScPoi(ByVal FxDhArr(),ByVal DhCount)
    For i = 0 To DhCount - 1
        FxArr = Split(FxDhArr(i),",", - 1,1)
        If FxArr(0) <> "" Then
            SqlStr = "Select " & RealTable & ".ID From " & RealTable & " WHERE " & RealTable & ".ZRZBSM = " & FxArr(0) 
            GetSQLRecordAll SqlStr,ScDhArr,ScCount
            If ScCount <= 0 Then
                SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(FxArr(1),"SSObj_X"),SSProcess.GetObjectAttr(FxArr(1),"SSObj_Y"),0,2,FxArr(1),""
            End If
        End If
    Next 'i
End Function' ConfirmScPoi

'==============================================================工具函数==========================================================

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
End Function' GetSQLRecordAll

'清空检查集
Function ClearCheckRecord()
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
End Function' ClearCheckRecord

'显示检查记录
Function ShowCheckRecord()
    SSProcess.ShowCheckOutput
    SSProcess.SaveCheckRecord
End Function' ShowCheckRecord
