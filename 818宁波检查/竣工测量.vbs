
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
        Area = Transform(SSProcess.GetObjectAttr(IDArr(i),"[MJ]"))
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
        Area = Transform(SSProcess.GetObjectAttr(IDArr(i),"[MJ]"))
        SNCWGS = SNCWGS + Round(Area * ZSXS)
    Next 'i
    
    If DXJDCWGS <> SNCWGS Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' DXJDCCheck

'地上非机动车位个数与地上非机动车位个数是否一致
Function DSFJDCWCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【DSFJDCWGS】
    ' 2：室外车位属性表（SWCW）表中【CWLX】=“非机动车位“ ，按照【ZSXS】值进行统计汇总（面积*折算系数算出个数，汇总）
    
    '检查记录配置
    strGroupName = "总体指标表面积逻辑检查"
    strCheckName = "地上非机动车位个数与地上非机动车位个数一致性检查"
    CheckmodelName = "自定义脚本检查类->地上非机动车位个数与地上非机动车位个数一致性检查"
    strDescription = "地上非机动车位个数与地上非机动车位个数不一致"
    
    '获取地下机动车车位个数 DSFJDCWGS
    SqlStr = "Select JGSCHZXX.DXJDCWGS From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,DSFJDCWGSArr,SearchCount
    DSFJDCWGS = DSFJDCWGSArr(0)
    
    '获取室外车位个数 SWCWGS
    SqlStr = "Select GH_室外车位属性表.ID From GH_室外车位属性表 Inner Join GeoAreaTB On GH_室外车位属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And GH_室外车位属性表.CWLX = '非机动车位' "
    GetSQLRecordAll SqlStr,IDArr,IDCount
    
    For i = 0 To IDCount - 1
        ZSXS = Transform(SSProcess.GetObjectAttr(IDArr(i),"[ZSXS]"))
        Area = Transform(SSProcess.GetObjectAttr(IDArr(i),"[MJ]"))
        SWCWGS = SWCWGS + Round(Area * ZSXS)
    Next 'i
    
    If DSFJDCWGS <> SWCWGS Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
End Function' DSFJDCWCheck

'地上非机动车位核实数量检查
Function DSFJDCHES()
    
    ' 1：室外车位属性表（SWCW）表中【CWLX】=“非机动车位“ ，面积【MJ】*折算系数【ZSXS】是否等于车位个数【CWGS】
    
    '检查记录配置
    strGroupName = "总体指标表面积逻辑检查"
    strCheckName = "地上非机动车位核实数量检查"
    CheckmodelName = "自定义脚本检查类->地上非机动车位核实数量检查"
    strDescription = "地上非机动车位核实数量不一致"
    
    SqlStr = "Select GH_室外车位属性表.ID From GH_室外车位属性表 Inner Join GeoAreaTB On GH_室外车位属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And GH_室外车位属性表.CWLX = '非机动车位' "
    GetSQLRecordAll SqlStr,IDArr,IDCount
    
    For i = 0 To IDCount - 1
        ZSXS = Transform(SSProcess.GetObjectAttr(IDArr(i),"[ZSXS]"))
        Area = Transform(SSProcess.GetObjectAttr(IDArr(i),"[MJ]"))
        CWGS = Transform(SSProcess.GetObjectAttr(IDArr(i),"[CWGS]"))
        If Round(Area * ZSXS) <> CWGS Then
            SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(IDArr(i),"SSObj_X"),SSProcess.GetObjectAttr(IDArr(i),"SSObj_Y"),0,2,IDArr(i),""
        End If
    Next 'i
    
End Function' DSFJDCHES

'地下非机动车位个数与地下非机动车位个是否一致
Function DXFJDCWCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【DXFJDCWGS】
    ' 2：室内车位属性表（SNCW）表中【CWLX】=“非机动车位“ ，按照【ZSXS】值进行汇总（面积*折算系数算出个数，汇总）
    
    '检查记录配置
    strGroupName = "总体指标表面积逻辑检查"
    strCheckName = "地下非机动车位个数与地下非机动车位个数一致性检查"
    CheckmodelName = "自定义脚本检查类->地下非机动车位个数与地下非机动车位个数一致性检查"
    strDescription = "地下非机动车位个数与地下非机动车位个数不一致"
    
    '获取地下机动车车位个数 DXFJDCWGS
    SqlStr = "Select JGSCHZXX.DXJDCWGS From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,DXFJDCWGSArr,SearchCount
    DXFJDCWGS = DXFJDCWGSArr(0)
    
    '获取室外车位个数 SNCWGS
    SqlStr = "Select GH_室内车位属性表.ID From GH_室内车位属性表 Inner Join GeoAreaTB On GH_室内车位属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And GH_室内车位属性表.CWLX = '非机动车位' "
    GetSQLRecordAll SqlStr,IDArr,IDCount
    
    For i = 0 To IDCount - 1
        ZSXS = Transform(SSProcess.GetObjectAttr(IDArr(i),"[ZSXS]"))
        Area = Transform(SSProcess.GetObjectAttr(IDArr(i),"[MJ]"))
        SNCWGS = SNCWGS + Round(Area * ZSXS)
    Next 'i
    
    If DXFJDCWGS <> SNCWGS Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
End Function' DXFJDCWCheck

'地下非机动车位核实数量检查
Function DXFJDCHES()
    
    ' 1：室内车位属性表（SNCW）表中【CWLX】=“非机动车位“ ，面积【MJ】*折算系数【ZSXS】是否等于车位个数【CWGS】
    
    '检查记录配置
    strGroupName = "总体指标表面积逻辑检查"
    strCheckName = "地下非机动车位核实数量检查"
    CheckmodelName = "自定义脚本检查类->地下非机动车位核实数量检查"
    strDescription = "地下非机动车位核实数量不一致"
    
    SqlStr = "Select GH_室内车位属性表.ID From GH_室内车位属性表 Inner Join GeoAreaTB On GH_室内车位属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And GH_室内车位属性表.CWLX = '非机动车位' "
    GetSQLRecordAll SqlStr,IDArr,IDCount
    
    For i = 0 To IDCount - 1
        ZSXS = Transform(SSProcess.GetObjectAttr(IDArr(i),"[ZSXS]"))
        Area = Transform(SSProcess.GetObjectAttr(IDArr(i),"[MJ]"))
        CWGS = Transform(SSProcess.GetObjectAttr(IDArr(i),"[CWGS]"))
        If Round(Area * ZSXS) <> CWGS Then
            SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(IDArr(i),"SSObj_X"),SSProcess.GetObjectAttr(IDArr(i),"SSObj_Y"),0,2,IDArr(i),""
        End If
    Next 'i
    
End Function' DXFJDCHES

'绿地总面积是否等于集中绿地面积+单块绿地面积面积和
Function LvDAreaCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【LDZMJ】=【JZLDMJ】+【DKLDMJ】
    
    '检查记录配置
    strGroupName = "总体指标表面积逻辑检查"
    strCheckName = "绿地总面积检查"
    CheckmodelName = "自定义脚本检查类->绿地总面积检查"
    strDescription = "绿地总面积与集中绿地和单块绿地面积之和不一致"
    
    '获取绿地总面积 LDZMJ
    SqlStr = "Select JGSCHZXX.LDZMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,LDZMJArr,LDCount
    LDZMJ = LDZMJArr(0)
    
    '获取集中绿地和单块绿地面积之和 SumArea
    SqlStr = "Select JGSCHZXX.JZLDMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,JZLDMJArr,JZLDCount
    JZLDMJ = Transform(JZLDMJArr(0))
    
    SqlStr = "Select JGSCHZXX.DKLDMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,DKLDMJArr,DKLDCount
    DKLDMJ = Transform(DKLDMJArr(0))
    
    SumArea = JZLDMJ + DKLDMJ
    
    If LDZMJ <> SumArea Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
End Function' LvDAreaCheck

'单块绿地面积与单块绿地范围面面积汇总值是否一致
Function DKLVCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【DKLDMJ】
    ' 2：绿化划分信息表（LHHF）其中的【MC】=单块绿地，并通过【ID_LDK】绿地块ID与绿化要素属性表（LHYS）中的【ID_LDK】取【LHMJ】的汇总值
    
    '检查记录配置
    strGroupName = "总体指标表面积逻辑检查"
    strCheckName = "单块绿地面积与单块绿地范围面面积汇总值一致性检查"
    CheckmodelName = "自定义脚本检查类->单块绿地面积与单块绿地范围面面积汇总值一致性检查"
    strDescription = "单块绿地面积与单块绿地范围面面积汇总值不一致"
    
    '单块绿地总面积 DKLDMJ
    SqlStr = "Select JGSCHZXX.DKLDMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,DKLDMJArr,JZLDCount
    DKLDMJ = Transform(DKLDMJArr(0))
    
    '汇总绿化面积 SumArea
    SqlStr = "Select LHHF.ID_LDK From LHHF Where LHHF.MC = '单块绿地' "
    GetSQLRecordAll SqlStr,IDArr,IDCount
    
    For i = 0 To IDCount - 1
        SumArea = SumArea + Transform(SSProcess.GetObjectAttr(IDArr(i),"[LHMJ]"))
    Next 'i
    
    If DKLDMJ <> SumArea Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' DKLVCheck

'集中绿地面积与集中绿地范围面面积汇总值是否一致
Function JZLDCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【JZLDMJ】
    ' 2：绿化划分信息表（LHHF）其中的【MC】=集中绿地，并通过【ID_LDK】绿地块ID与绿化要素属性表（LHYS）中的【ID_LDK】取【LHMJ】的汇总值
    
    '检查记录配置
    strGroupName = "总体指标表面积逻辑检查"
    strCheckName = "集中绿地面积与集中绿地范围面面积汇总值一致性检查"
    CheckmodelName = "自定义脚本检查类->集中绿地面积与集中绿地范围面面积汇总值一致性检查"
    strDescription = "集中绿地面积与集中绿地范围面面积汇总值不一致"
    
    '集中绿地面积 JZLDMJ
    SqlStr = "Select JGSCHZXX.JZLDMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,JZLDMJArr,JZLDCount
    JZLDMJ = Transform(JZLDMJArr(0))
    
    '汇总绿化面积 SumArea
    SqlStr = "Select LHHF.ID_LDK From LHHF Where LHHF.MC = '集中绿地' "
    GetSQLRecordAll SqlStr,IDArr,IDCount
    
    For i = 0 To IDCount - 1
        SumArea = SumArea + Transform(SSProcess.GetObjectAttr(IDArr(i),"[LHMJ]"))
    Next 'i
    
    If DKLDMJ <> SumArea Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' JZLDCheck

'登高场地个数与登高场地面个数是否一致
Function DGCDCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【RFZMJ】
    ' 2：人防功能区属性表（RFGNQ）中【JZMJ】值累加和
    
    '检查记录配置
    strGroupName = "总体指标表面积逻辑检查"
    strCheckName = "登高场地个数与登高场地面个数一致性检查"
    CheckmodelName = "自定义脚本检查类->登高场地个数与登高场地面个数一致性检查"
    strDescription = "登高场地个数与登高场地面个数不一致"
    
    
End Function' DGCDCheck

'人防总面积与人防功能区面积汇总值是否一致
Function RFMJCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【RFZMJ】
    ' 2：人防功能区属性表（RFGNQ）中【JZMJ】值累加和
    
    '检查记录配置
    strGroupName = "总体指标表面积逻辑检查"
    strCheckName = "人防总面积与人防功能区面积汇总值一致性检查"
    CheckmodelName = "自定义脚本检查类->人防总面积与人防功能区面积汇总值一致性检查"
    strDescription = "人防总面积与人防功能区面积汇总值不一致"
    
    '获取人防总面积 RFZMJ
    SqlStr = "Select JGSCHZXX.RFZMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,RFZMJArr,JZLDCount
    RFZMJ = Transform(RFZMJArr(0))
    
    '汇总人防面积 SumArea
    SqlStr = "Select Sum(RF_人防功能区属性表.JZMJ) From RF_人防功能区属性表 Inner Join GeoAreaTB On RF_人防功能区属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 "
    GetSQLRecordAll SqlStr,SumAreaArr,SumCount
    SumArea = Transform(SumAreaArr(0))
    
    If RFZMJ <> SumArea Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
End Function' RFMJCheck

'房屋用途与功能区用途面积汇总值是否一致（所有幢）
Function FWCheck()
    
    ' 1：主要经济指标面积汇总信息表(SCZYJJZBXXB)中的每个【YT】：例如：住宅 面积【SCJZMJ】
    ' 2：规划功能区（GHGNQ）表中的【YT】 = “住宅”的所有面积值。
    
    '检查记录配置
    strGroupName = "主要经济技术指标检查"
    strCheckName = "房屋用途与功能区用途面积汇总值一致性检查"
    CheckmodelName = "自定义脚本检查类->房屋用途与功能区用途面积汇总值一致性检查"
    strDescription = "房屋用途与功能区用途面积汇总值不一致"
    
    SqlStr = "Select DISTINCT SCZYJJZBXXB.YT From SCZYJJZBXXB Where SCZYJJZBXXB.ID > 0"
    GetSQLRecordAll SqlStr,YTArr,YTCount
    
    For i = 0 To YTCount - 1
        
        SqlStr = "Select Sum(JG_规划功能区属性表.JZMJ) Form JG_规划功能区属性表 Inner Join GeoAreaTB On JG_规划功能区属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And JG_规划功能区属性表.YT = '" & YTArr(i) & "'"
        GetSQLRecordAll SqlStr,SumAreaArr,SumCount
        SumArea = SumAreaArr(0)
        
        SqlStr = "Select SCZYJJZBXXB.SCJZMJ From SCZYJJZBXXB Where SCZYJJZBXXB.ID > 0 And SCZYJJZBXXB.YT = '" & YTArr(i) & "'"
        GetSQLRecordAll SqlStr,SCJZMJArr,SearchCount
        SCJZMJ = SCJZMJArr(0)
        
        If SumArea <> SCJZMJ Then
            SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
        End If
    Next 'i
End Function' FWCheck

'房屋用途与功能区用途面积汇总值是否一致（按单幢）
Function DDFWCheck()
    
    ' 1：实测楼栋面积汇总信息表（SCLDMJHZXX）表中【LD】=“1#”且【YT】=“住宅”的【JZMJ】
    ' 2：规划功能区（GHGNQ）表中的【SSZRZ】=“1#”且【YT】=“住宅”的【JZMJ】的值的累加值。
    
    '检查记录配置
    strGroupName = "建筑物建筑面积汇检查"
    strCheckName = "房屋用途与功能区用途面积汇总值一致性检查（按单幢）"
    CheckmodelName = "自定义脚本检查类->房屋用途与功能区用途面积汇总值一致性检查（按单幢）"
    strDescription = "房屋用途与功能区用途面积汇总值不一致"
    
    '所有的楼栋
    SqlStr = "Select DISTINCT SCLDMJHZXX.LD From SCLDMJHZXX Where SCLDMJHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,LDArr,LDCount
    
    For i = 0 To LDCount - 1
        SqlStr = "Select DISTINCT SCLDMJHZXX.YT From SCLDMJHZXX Where SCLDMJHZXX.ID > 0 And SCLDMJHZXX.LD = '" & LDArr(i) & "'"
        GetSQLRecordAll SqlStr,YTArr,YTCount
        For j = 0 To YTCount - 1
            SqlStr = "Select Sum(JG_规划功能区属性表.JZMJ) Form JG_规划功能区属性表 Inner Join GeoAreaTB On JG_规划功能区属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And JG_规划功能区属性表.SSZRZ = '" & LDArr(i) & "' And JG_规划功能区属性表.YT = '" & YTArr(j) & "'"
            GetSQLRecordAll SqlStr,SumAreaArr,SumCount
            SumArea = SumAreaArr(0)
            
            SqlStr = "Select SCLDMJHZXX.JZMJ Where SCLDMJHZXX.LD = '" & LDArr(i) & "' And SCLDMJHZXX.YT = '" & YTArr(j) & "'"
            GetSQLRecordAll SqlStr,JZMJArr,SearchCount
            JZMJ = JZMJArr(0)
            
            If JZMJ <> SumArea Then
                SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
            End If
        Next 'j
    Next 'i
    
End Function' DDFWCheck

'防护单元个数与防护单元范围线个数否一致
Function FHDYGSCheck()
    
    ' 1：人防项目信息表（RFPROJECTINFO）中的【FHDYGS】的值
    ' 2:人防防护单元范围线（RFFHDYFW）要素个数。
    
    '检查记录配置
    strGroupName = "人防工程基本信息检查"
    strCheckName = "防护单元个数与防护单元范围线个数一致性检查"
    CheckmodelName = "自定义脚本检查类->防护单元个数与防护单元范围线个数一致性检查"
    strDescription = "防护单元个数与防护单元范围线个数不一致"
    
    '获取防护单元个数 FHDYGS
    SqlStr = "Select RFPROJECTINFO.FHDYGS From RFPROJECTINFO Where RFPROJECTINFO.ID > 0 "
    GetSQLRecordAll SqlStr,FHDYGSArr,FHDYGSCount
    FHDYGS = FHDYGSArr(0)
    
    '获取图上范围线个数 YSCount
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 9450013
    SSProcess.SelectFilter
    YSCount = SSProcess.GetSelGeoCount()
    
    If YSCount <> FHDYGS Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
End Function' FHDYGSCheck

'人防建筑面积与人防功能区面积汇总值是否一致
Function RFJZMJCheck()
    
    ' 1：人防项目信息表（RFPROJECTINFO）中的【RFJZMJ】的值
    ' 2:人防功能区（RFGNQ）中的【JZMJ】的所有汇总值
    
    '检查记录配置
    strGroupName = "人防工程基本信息检查"
    strCheckName = "人防建筑面积与人防功能区面积汇总值一致性检查"
    CheckmodelName = "自定义脚本检查类->人防建筑面积与人防功能区面积汇总值一致性检查"
    strDescription = "人防建筑面积与人防功能区面积汇总值不一致"
    
    '人防建筑面积 RFJZMJ
    SqlStr = "Select RFPROJECTINFO.RFJZMJ From RFPROJECTINFO Where RFPROJECTINFO.ID > 0 "
    GetSQLRecordAll SqlStr,RFJZMJArr,RFJZCount
    RFJZMJ = RFJZMJArr(0)
    
    '人防功能区面积汇总值 SumArea
    SqlStr = "Select Sum(RF_人防功能区属性表.JZMJ) Form RF_人防功能区属性表 Inner Join GeoAreaTB On RF_人防功能区属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0"
    GetSQLRecordAll SqlStr,SumAreaArr,SumCount
    SumArea = SumAreaArr(0)
    
    If RFJZMJ <> SumArea Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
End Function' RFJZMJCheck

'掩蔽区面积与人防功能区（掩蔽区）面积汇总值是否一致
Function YBQCheck()
    
    ' 1：人防项目信息表（RFPROJECTINFO）中的【YBQMJ】的值
    ' 2:人防功能区（RFGNQ）中的【YSDM】=“600301”的【JZMJ】的所有汇总值
    
    '检查记录配置
    strGroupName = "人防工程基本信息检查"
    strCheckName = "掩蔽区面积与人防功能区（掩蔽区）面积汇总值一致性检查"
    CheckmodelName = "自定义脚本检查类->掩蔽区面积与人防功能区（掩蔽区）面积汇总值一致性检查"
    strDescription = "掩蔽区面积与人防功能区（掩蔽区）面积汇总值不一致"
    
    '掩蔽区面积 YBQMJ
    SqlStr = "Select RFPROJECTINFO.YBQMJ From RFPROJECTINFO Where RFPROJECTINFO.ID > 0 "
    GetSQLRecordAll SqlStr,YBQMJArr,YBQCount
    YBQMJ = YBQMJArr(0)
    
    '人防功能区（掩蔽区）面积汇总值 SumArea
    SqlStr = "Select Sum(RF_人防功能区属性表.JZMJ) Form RF_人防功能区属性表 Inner Join GeoAreaTB On RF_人防功能区属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And RF_人防功能区属性表.YSDM = '" & "600301'"
    GetSQLRecordAll SqlStr,SumAreaArr,SumCount
    SumArea = SumAreaArr(0)
    
    If YBQMJ <> SumArea Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
End Function' YBQCheck

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