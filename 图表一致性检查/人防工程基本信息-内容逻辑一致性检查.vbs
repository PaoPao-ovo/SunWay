
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

    FHDYGSCheck

    RFJZMJCheck
    
    YBQCheck
    
    ShowCheckRecord

End Sub' OnClick

'===================================================检查函数=======================================================

'防护单元个数与防护单元范围线个数否一致
Function FHDYGSCheck()
    
    ' 1：人防项目信息表（RFPROJECTINFO）中的【FHDYGS】的值
    ' 2:人防防护单元范围线（RFFHDYFW）要素个数。
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "防护单元个数与防护单元范围线个数一致性检查"
    CheckmodelName = "自定义脚本检查类->防护单元个数与防护单元范围线个数一致性检查"
    strDescription = "防护单元个数与防护单元范围线个数不一致"

    '获取防护单元个数 FHDYGS
    SqlStr = "Select RFPROJECTINFO.Value From RFPROJECTINFO Where RFPROJECTINFO.ID > 0 And RFPROJECTINFO.Key = '防护单元个数' "
    GetSQLRecordAll SqlStr,FHDYGSArr,FHDYGSCount

    If FHDYGSCount > 0 Then
        FHDYGS = Transform(FHDYGSArr(0))
    Else
        FHDYGS = 0
    End If
    
    '获取图上范围线个数 YSCount
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 9450013
    SSProcess.SelectFilter
    YSCount = SSProcess.GetSelGeoCount()
    
    If YSCount - FHDYGS <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If

End Function' FHDYGSCheck

'人防建筑面积与人防功能区面积汇总值是否一致
Function RFJZMJCheck()
    
    ' 1：人防项目信息表（RFPROJECTINFO）中的【RFJZMJ】的值
    ' 2:人防功能区（RFGNQ）中的【JZMJ】的所有汇总值
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "人防建筑面积与人防功能区面积汇总值一致性检查"
    CheckmodelName = "自定义脚本检查类->人防建筑面积与人防功能区面积汇总值一致性检查"
    strDescription = "人防建筑面积与人防功能区面积汇总值不一致"

    '人防建筑面积 RFJZMJ
    SqlStr = "Select RFPROJECTINFO.Value From RFPROJECTINFO Where RFPROJECTINFO.ID > 0 And RFPROJECTINFO.Key = '人防建筑面积' "
    GetSQLRecordAll SqlStr,RFJZMJArr,RFJZCount

    If RFJZCount > 0 Then
        RFJZMJ = Transform(RFJZMJArr(0))
    Else
        RFJZMJ = 0
    End If
    
    '人防功能区面积汇总值 SumArea
    SqlStr = "Select Sum(RF_人防功能区属性表.JZMJ) From RF_人防功能区属性表 Inner Join GeoAreaTB On RF_人防功能区属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0"
    GetSQLRecordAll SqlStr,SumAreaArr,SumCount

    If SumCount > 0 Then
        SumArea = Transform(SumAreaArr(0))
    Else
        SumArea = 0
    End If
    
    If RFJZMJ - SumArea <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If

End Function' RFJZMJCheck

'掩蔽区面积与人防功能区（掩蔽区）面积汇总值是否一致
Function YBQCheck()
    
    ' 1：人防项目信息表（RFPROJECTINFO）中的【YBQMJ】的值
    ' 2:人防功能区（RFGNQ）中的【YSDM】=“600301”的【JZMJ】的所有汇总值
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "掩蔽区面积与人防功能区（掩蔽区）面积汇总值一致性检查"
    CheckmodelName = "自定义脚本检查类->掩蔽区面积与人防功能区（掩蔽区）面积汇总值一致性检查"
    strDescription = "掩蔽区面积与人防功能区（掩蔽区）面积汇总值不一致"

    '掩蔽区面积 YBQMJ
    SqlStr = "Select RFPROJECTINFO.Value From RFPROJECTINFO Where RFPROJECTINFO.ID > 0 And RFPROJECTINFO.Key = '掩蔽区面积' "
    GetSQLRecordAll SqlStr,YBQMJArr,YBQCount

    If YBQCount > 0 Then
        YBQMJ = Transform(YBQMJArr(0))
    Else
        YBQMJ = 0
    End If
    
    '人防功能区（掩蔽区）面积汇总值 SumArea
    SqlStr = "Select Sum(RF_人防功能区属性表.JZMJ) From RF_人防功能区属性表 Inner Join GeoAreaTB On RF_人防功能区属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And RF_人防功能区属性表.YSDM = '" & "600301'"
    GetSQLRecordAll SqlStr,SumAreaArr,SumCount

    If SumCount > 0 Then
        SumArea = Transform(SumAreaArr(0))
    Else
        SumArea = 0
    End If
    
    If YBQMJ - SumArea <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
  
End Function' YBQCheck

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