
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
    
End Sub' OnClick

'===================================================检查函数=======================================================

'建筑面积值与幢面积汇总值是否一致
Function ZhuangCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【JZMJ】
    ' 2:自然幢（JG_自然幢属性表）表中【JZMJ】累计汇总。
    
    '检查记录配置
    strGroupName = "总体指标表面积逻辑检查"
    strCheckName = "建筑面积值与幢面积汇总值一致性检查"
    CheckmodelName = "自定义脚本检查类->建筑面积值与幢面积汇总值一致性检查"
    strDescription = "建筑面积值与幢面积汇总值不一致"
    
    '获取总建筑面积 JZMJ
    SqlStr = "Select Sum(JGSCHZXX.JZMJ) From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,JZMJArr,SearchCount
    JZMJ = JZMJArr(0)
    
    '获取自然幢总面积 SumArea
    SqlStr = "Select Sum(JG_自然幢属性表.JZMJ) From JG_自然幢属性表 Inner Join GeoAreaTB On JG_自然幢属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0"
    GetSQLRecordAll SqlStr,SumAreaArr,SumCount
    SumArea = SumAreaArr(0)
    
    If JZMJ <> SumArea Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
End Function' ZhuangCheck

'建筑基地面积与基地面汇总值是否一致
Function BasementCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【JDMJ】
    ' 2: 基底_面(JD_POLYGON)属性表中的【JDMJ】的所有记录的累加和
    
    '检查记录配置
    strGroupName = "总体指标表面积逻辑检查"
    strCheckName = "建筑基地面积与基地面汇总值一致性检查"
    CheckmodelName = "自定义脚本检查类->建筑基地面积与基地面汇总值一致性检查"
    strDescription = "建筑基地面积与基地面汇总值不一致"
    
    '获取总面积 JDMJ
    SqlStr = "Select Sum(JGSCHZXX.JDMJ) From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,JDMJArr,SearchCount
    JDMJ = JDMJArr(0)
    
    '获取基地面积之和 SumArea
    SqlStr = "Select Sum(JD_POLYGON.JDMJ) From JD_POLYGON Inner Join GeoAreaTB On JG_自然幢属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And JD_POLYGON.ID > 0"
    GetSQLRecordAll SqlStr,SumAreaArr,SumCount
    SumArea = SumAreaArr(0)
    
    If JDMJ <> SumArea Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
End Function' BasementCheck

'绿地面积与绿地范围线面积汇总值是否一致性
Function LvAreaCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【LDMJ】
    ' 2:绿化要素属性表(LHYS)中【LHMJ】的所有记录的累加和
    
    '检查记录配置
    strGroupName = "总体指标表面积逻辑检查"
    strCheckName = "绿地面积与绿地范围线面积汇总值一致性检查"
    CheckmodelName = "自定义脚本检查类->绿地面积与绿地范围线面积汇总值一致性检查"
    strDescription = "绿地面积与绿地范围线面积汇总值不一致"
    
    '绿地总面积 LDMJ
    SqlStr = "Select Sum(JGSCHZXX.LDMJ) From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,LDMJArr,SearchCount
    LDMJ = LDMJArr(0)
    
    '绿化要素面积之和 SumLhArea
    SqlStr = "Select Sum(GH_绿化要素属性表.LHMJ) From GH_绿化要素属性表 Inner Join GeoAreaTB On GH_绿化要素属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And GH_绿化要素属性表.ID > 0"
    GetSQLRecordAll SqlStr,LHMJArr,LHCount
    SumLhArea = LHMJArr(0)
    
    If LDMJ <> SumLhArea Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
End Function' LvAreaCheck

'建筑密度与基地面积除用地面积的值是否一致
Function ConstractDensityCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【JZMD】
    ' 2：规划实测汇总信息表(JGSCHZXX)表中【JDMJ】/【YDMJ】
    
    '检查记录配置
    strGroupName = "总体指标表面积逻辑检查"
    strCheckName = "建筑密度与基地面积除用地面积一致性检查"
    CheckmodelName = "自定义脚本检查类->建筑密度与基地面积除用地面积一致性检查"
    strDescription = "建筑密度与基地面积除用地面积不一致"
    
    '获取建筑密度 JZMD
    SqlStr = "Select JGSCHZXX.JZMD From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,JZMDArr,SearchCount
    JZMD = JZMDArr(0)
    
    '获取基底面积 JDMJ
    SqlStr = "Select JGSCHZXX.JDMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,JDMJArr,SearchCount
    JDMJ = JDMJArr(0)
    
    '获取用地面积 YDMJ
    SqlStr = "Select JGSCHZXX.YDMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,YDMJArr,SearchCount
    YDMJ = YDMJArr(0)
    
    '计算密度 Density
    Density = JDMJ / YDMJ
    
    If JZMD <> Density Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
End Function' ConstractDensityCheck

'绿化率值与绿地面积除以用地面积值是否一致
Function LHPercrntCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【LVL】
    ' 2：规划实测汇总信息表(JGSCHZXX)表中【LDMJ】/【YDMJ】
    
    '检查记录配置
    strGroupName = "总体指标表面积逻辑检查"
    strCheckName = "绿化率值与绿地面积除以用地面积一致性检查"
    CheckmodelName = "自定义脚本检查类->绿化率值与绿地面积除以用地面积一致性检查"
    strDescription = "绿化率值与绿地面积除以用地面积不一致"
    
    '获取绿化率 LVL
    SqlStr = "Select JGSCHZXX.LVL From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,LVLArr,SearchCount
    LVL = LVLArr(0)
    
    '获取绿地面积 LDMJ
    SqlStr = "Select JGSCHZXX.LDMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,LDMJArr,SearchCount
    LDMJ = LDMJArr(0)
    
    '获取用地面积 YDMJ
    SqlStr = "Select JGSCHZXX.YDMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,YDMJArr,SearchCount
    YDMJ = YDMJArr(0)
    
    '实际密度 RealDensity
    RealDensity = LDMJ / YDMJ
    
    If RealDensity <> LVL Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' LHPercrntCheck

'地上机动车位个数与地上停车位个数是否一致
Function DSJDCCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【DSJDCWGS】
    ' 2：室外车位属性表（SWCW）表中【CWLX】<> “非机动车位” ，按照【ZSXS】值进行统计汇总（面积*折算系数算出个数，汇总）
    
    '检查记录配置
    strGroupName = "总体指标表面积逻辑检查"
    strCheckName = "地上机动车位个数与地上停车位个数一致性检查"
    CheckmodelName = "自定义脚本检查类->地上机动车位个数与地上停车位个数一致性检查"
    strDescription = "地上机动车位个数与地上停车位个数不一致"
    
    '获取地上机动车车位个数 DSJDCWGS
    SqlStr = "Select JGSCHZXX.DSJDCWGS From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,DSJDCWGSArr,SearchCount
    DSJDCWGS = DSJDCWGSArr(0)
    
    '获取室外机动车个数 SWCWGS
    SqlStr = "Select GH_室外车位属性表.ID From GH_室外车位属性表 Inner Join GeoAreaTB On GH_室外车位属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And GH_室外车位属性表.CWLX <> '非机动车位' "
    GetSQLRecordAll SqlStr,IDArr,IDCount
    
    For i = 0 To IDCount - 1
        ZSXS = Transform(SSProcess.GetObjectAttr(IDArr(i),"[ZSXS]"))
        Area = Transform(SSProcess.GetObjectAttr(IDArr(i),"[CWMJ]"))
        SWCWGS = SWCWGS + Round(Area * ZSXS)
    Next 'i
    
    If DSJDCWGS <> SWCWGS Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
End Function' DSJDCCheck

'地下机动车位个数与地下停车位个数是否一致
Function DXJDCCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【DXJDCWGS】
    ' 2：室内车位属性表（SNCW）表中【CWLX】 <> “非机动车位“ ，按照【ZSXS】值进行汇总（面积 * 折算系数算出个数，汇总）
    
    '检查记录配置
    strGroupName = "总体指标表面积逻辑检查"
    strCheckName = "地下机动车位个数与地下停车位个数一致性检查"
    CheckmodelName = "自定义脚本检查类->地下机动车位个数与地下停车位个数一致性检查"
    strDescription = "地下机动车位个数与地下停车位个数不一致"
    
    '获取地下机动车车位个数 DXJDCWGS
    SqlStr = "Select JGSCHZXX.DXJDCWGS From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,DXJDCWGSArr,SearchCount
    DXJDCWGS = DXJDCWGSArr(0)
    
    '获取室外机动车个数 SNCWGS
    SqlStr = "Select GH_室内车位属性表.ID From GH_室内车位属性表 Inner Join GeoAreaTB On GH_室内车位属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And GH_室内车位属性表.CWLX <> '非机动车位' "
    GetSQLRecordAll SqlStr,IDArr,IDCount
    
    For i = 0 To IDCount - 1
        ZSXS = Transform(SSProcess.GetObjectAttr(IDArr(i),"[ZSXS]"))
        Area = Transform(SSProcess.GetObjectAttr(IDArr(i),"[CWMJ]"))
        SNCWGS = SNCWGS + Round(Area * ZSXS)
    Next 'i
    
    If DXJDCWGS <> SNCWGS Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
End Function' DXJDCCheck

'地上非机动车位个数与地上非机动车位个数是否一致
Function Name(Parameter)
    
End Function ' Name

'======================================================工具类函数====================================================

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