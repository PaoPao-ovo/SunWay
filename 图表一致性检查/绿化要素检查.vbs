
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
    
    RecordExist
    
    LHYSCheck
    
    ShowCheckRecord
    
End Sub' OnClick

'===================================================检查函数=======================================================


Function LHYSCheck()
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "111"
    CheckmodelName = "自定义脚本检查类->111"
    strDescription = "222"
    
    SqlStr = "Select DISTINCT GH_绿化要素属性表.ID_LDK From 地类GH_绿化要素属性表图斑属性表 INNER JOIN GeoAreaTB ON GH_绿化要素属性表.ID = GeoAreaTB.ID WHERE([GeoAreaTB].[Mark] Mod 2)<>0"
    
    GetSQLRecordAll SqlStr,ID_LDKArr,ID_LDKCount
    
    
    If ID_LDKCount > 0 Then
        For i = 0 To ID_LDKCount - 1
            SqlStr = "Select LHHF.绿地块ID From LHHF Where LHHF.绿地块ID = '" & ID_LDKArr(i) & "'"
            GetSQLRecordAll SqlStr,LHHFArr,LHHFCount
            If LHHFCount < 0 Then
                SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,0,0,""
            End If
        Next 'i
    End If
    
End Function' LHYSCheck

Function RecordExist()
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "111"
    CheckmodelName = "自定义脚本检查类->111"
    strDescription = "绿化划分表【LHHF】的【绿地块ID】不存在记录"
    
    SqlStr = "Select DISTINCT LHHF.绿地块ID From LHHF"
    GetSQLRecordAll SqlStr,LHHFArr,LHHFCount
    
    If LHHFCount < 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,0,0,""
    End If

End Function' RecordExist

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