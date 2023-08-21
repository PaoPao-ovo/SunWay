
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
    
    ClearCheckRecord
    
    ZhuangCheck
    
    BasementCheck
    
    LvAreaCheck
    
    ConstractDensityCheck
    
    LHPercrntCheck
    
    DSJDCCheck
    
    DXJDCCheck
    
    DSFJDCWCheck
    
    DSFJDCHES
    
    DXFJDCWCheck
    
    DXFJDCHES
    
    LvDAreaCheck
    
    DKLVCheck
    
    JZLDCheck
    
    DGCDCheck
    
    RFMJCheck
    
    ShowCheckRecord
    
End Sub' OnClick

'===================================================检查函数=======================================================

'建筑面积值与幢面积汇总值是否一致
Function ZhuangCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【JZMJ】
    ' 2:自然幢（JG_自然幢属性表）表中【JZMJ】累计汇总。
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "总体指标表面积逻辑检查"
    CheckmodelName = "自定义脚本检查类->总体指标表面积逻辑检查"
    strDescription = "建筑面积值与幢面积汇总值不一致"
    
    '获取总建筑面积 JZMJ
    SqlStr = "Select Sum(JGSCHZXX.JZMJ) From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,JZMJArr,SearchCount
    
    If SearchCount > 0 Then
        JZMJ = Transform(JZMJArr(0))
    Else
        JZMJ = 0
    End If
    
    
    '获取自然幢总面积 SumArea
    SqlStr = "Select Sum(FC_自然幢信息属性表.JZMJ) From FC_自然幢信息属性表 Inner Join GeoAreaTB On FC_自然幢信息属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0"
    GetSQLRecordAll SqlStr,SumAreaArr,SumCount
    
    If SumCount > 0 Then
        SumArea = Transform(SumAreaArr(0))
    Else
        SumArea = 0
    End If
    
    If JZMJ - SumArea <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' ZhuangCheck

'建筑基地面积与基地面汇总值是否一致
Function BasementCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【JZJDMJ】
    ' 2: 基底_面(JG_建筑物基底面属性表)属性表中的【JDMJ】的所有记录的累加和
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "建筑基地面积与基地面汇总值一致性检查"
    CheckmodelName = "自定义脚本检查类->建筑基地面积与基地面汇总值一致性检查"
    strDescription = "建筑基地面积与基地面汇总值不一致"
    
    '获取总面积 JDMJ
    SqlStr = "Select Sum(JGSCHZXX.JZJDMJ) From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,JDMJArr,SearchCount
    
    If SearchCount > 0 Then
        JDMJ = Transform(JDMJArr(0))
    Else
        JDMJ = 0
    End If
    
    '获取基地面积之和 SumArea
    SqlStr = "Select Sum(JG_建筑物基底面属性表.JDMJ) From JG_建筑物基底面属性表 Inner Join GeoAreaTB On JG_建筑物基底面属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And JG_建筑物基底面属性表.ID > 0"
    GetSQLRecordAll SqlStr,SumAreaArr,SumCount
    
    If SumCount > 0 Then
        SumArea = Transform(SumAreaArr(0))
    Else
        SumArea = 0
    End If
    
    If JDMJ - SumArea <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' BasementCheck

'绿地面积与绿地范围线面积汇总值是否一致性
Function LvAreaCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【LDMJ】
    ' 2:绿化要素属性表(LHYS)中【LHMJ】的所有记录的累加和
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "绿地面积与绿地范围线面积汇总值一致性检查"
    CheckmodelName = "自定义脚本检查类->绿地面积与绿地范围线面积汇总值一致性检查"
    strDescription = "绿地面积与绿地范围线面积汇总值不一致"
    
    '绿地总面积 LDMJ
    SqlStr = "Select Sum(JGSCHZXX.LDMJ) From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,LDMJArr,SearchCount
    
    If SearchCount > 0 Then
        LDMJ = Transform(LDMJArr(0))
    Else
        LDMJ = 0
    End If
    
    '绿化要素面积之和 SumLhArea
    SqlStr = "Select Sum(GH_绿化要素属性表.LHMJ) From GH_绿化要素属性表 Inner Join GeoAreaTB On GH_绿化要素属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And GH_绿化要素属性表.ID > 0"
    GetSQLRecordAll SqlStr,LHMJArr,LHCount
    
    If LHCount > 0 Then
        SumLhArea = Transform(LHMJArr(0))
    Else
        SumLhArea = 0
    End If
    
    If LDMJ - SumLhArea <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' LvAreaCheck

'建筑密度与基地面积除用地面积的值是否一致
Function ConstractDensityCheck()
    
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【JZMD】
    ' 2：规划实测汇总信息表(JGSCHZXX)表中【JZJDMJ】/【YDMJ】
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "建筑密度与基地面积除用地面积一致性检查"
    CheckmodelName = "自定义脚本检查类->建筑密度与基地面积除用地面积一致性检查"
    strDescription = "建筑密度与基地面积除用地面积不一致"
    
    '获取建筑密度 JZMD
    SqlStr = "Select JGSCHZXX.JZMD From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,JZMDArr,SearchCount
    
    If SearchCount > 0 Then
        JZMD = Transform(JZMDArr(0))
    Else
        JZMD = 0
    End If
    
    
    '获取基底面积 JDMJ
    SqlStr = "Select JGSCHZXX.JZJDMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,JDMJArr,SearchCount
    
    If SearchCount > 0 Then
        JDMJ = Transform(JDMJArr(0))
    Else
        JDMJ = 0
    End If
    
    '获取用地面积 YDMJ
    SqlStr = "Select JGSCHZXX.YDMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,YDMJArr,SearchCount
    
    If SearchCount > 0 Then
        YDMJ = Transform(YDMJArr(0))
    Else
        YDMJ = 0
    End If
    
    '计算密度 Density
    If YDMJ <> 0 Then
        Density = (JDMJ / YDMJ) * 100
    Else
        MsgBox "基底面积为空或零"
        Exit Function
        Density = 100
    End If
    
    If JZMD - Density <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' ConstractDensityCheck

'绿化率值与绿地面积除以用地面积值是否一致
Function LHPercrntCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【LVL】
    ' 2：规划实测汇总信息表(JGSCHZXX)表中【LDMJ】/【YDMJ】
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "绿化率值与绿地面积除以用地面积一致性检查"
    CheckmodelName = "自定义脚本检查类->绿化率值与绿地面积除以用地面积一致性检查"
    strDescription = "绿化率值与绿地面积除以用地面积不一致"
    
    '获取绿化率 LVL
    SqlStr = "Select JGSCHZXX.LVL From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,LVLArr,SearchCount
    
    If SearchCount > 0 Then
        LVL = Transform(LVLArr(0))
    Else
        LVL = 0
    End If
    
    
    '获取绿地面积 LDMJ
    SqlStr = "Select JGSCHZXX.LDMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,LDMJArr,SearchCount
    
    If SearchCount > 0 Then
        LDMJ = Transform(LDMJArr(0))
    Else
        LDMJ = 0
    End If
    
    '获取用地面积 YDMJ
    SqlStr = "Select JGSCHZXX.YDMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,YDMJArr,SearchCount
    
    
    If SearchCount > 0 Then
        YDMJ = Transform(YDMJArr(0))
    Else
        YDMJ = 0
    End If
    
    '实际密度 RealDensity
    If YDMJ <> 0 Then
        RealDensity = (LDMJ / YDMJ) * 100
    Else
        MsgBox "用地面积为空或零"
        Exit Function
        RealDensity = 100
    End If
    
    If RealDensity - LVL <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' LHPercrntCheck

'地上机动车位个数与地上停车位个数是否一致
Function DSJDCCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【DSJDCWGS】
    ' 2：室外车位属性表（SWCW）表中【CWLX】<> “非机动车位” ，按照【ZSXS】值进行统计汇总（面积*折算系数算出个数，汇总）
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "地上机动车位个数与地上停车位个数一致性检查"
    CheckmodelName = "自定义脚本检查类->地上机动车位个数与地上停车位个数一致性检查"
    strDescription = "地上机动车位个数与地上停车位个数不一致"
    
    '获取地上机动车车位个数 DSJDCWGS
    SqlStr = "Select JGSCHZXX.DSJDCWGS From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,DSJDCWGSArr,SearchCount
    
    If SearchCount > 0 Then
        DSJDCWGS = Transform(DSJDCWGSArr(0))
    Else
        DSJDCWGS = 0
    End If
    
    
    '获取室外机动车个数 SWCWGS
    SqlStr = "Select GH_室外车位属性表.ID From GH_室外车位属性表 Inner Join GeoAreaTB On GH_室外车位属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And GH_室外车位属性表.CWLX <> '非机动车位' "
    GetSQLRecordAll SqlStr,IDArr,IDCount
    
    If IDCount > 0 Then
        For i = 0 To IDCount - 1
            ZSXS = Transform(SSProcess.GetObjectAttr(IDArr(i),"[ZSXS]"))
            Area = Transform(SSProcess.GetObjectAttr(IDArr(i),"[MJ]"))
            SWCWGS = SWCWGS + Round(Area * ZSXS)
        Next 'i
    Else
        SWCWGS = 0
    End If
    
    If DSJDCWGS - SWCWGS <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' DSJDCCheck

'地下机动车位个数与地下停车位个数是否一致
Function DXJDCCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【DXJDCWGS】
    ' 2：室内车位属性表（SNCW）表中【CWLX】 <> “非机动车位“ ，按照【ZSXS】值进行汇总（面积 * 折算系数算出个数，汇总）
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "地下机动车位个数与地下停车位个数一致性检查"
    CheckmodelName = "自定义脚本检查类->地下机动车位个数与地下停车位个数一致性检查"
    strDescription = "地下机动车位个数与地下停车位个数不一致"
    
    '获取地下机动车车位个数 DXJDCWGS
    SqlStr = "Select JGSCHZXX.DXJDCWGS From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,DXJDCWGSArr,SearchCount
    
    If SearchCount > 0 Then
        DXJDCWGS = Transform(DXJDCWGSArr(0))
    Else
        DXJDCWGS = 0
    End If
    
    '获取室外机动车个数 SNCWGS
    SqlStr = "Select GH_室内车位属性表.ID From GH_室内车位属性表 Inner Join GeoAreaTB On GH_室内车位属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And GH_室内车位属性表.CWLX <> '非机动车位' "
    GetSQLRecordAll SqlStr,IDArr,IDCount
    
    If IDCount > 0 Then
        For i = 0 To IDCount - 1
            ZSXS = Transform(SSProcess.GetObjectAttr(IDArr(i),"[ZSXS]"))
            Area = Transform(SSProcess.GetObjectAttr(IDArr(i),"[MJ]"))
            SNCWGS = SNCWGS + Round(Area * ZSXS)
        Next 'i
    Else
        SNCWGS = 0
    End If
    
    If DXJDCWGS - SNCWGS <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' DXJDCCheck

'地上非机动车位个数与地上非机动车位个数是否一致
Function DSFJDCWCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【DSFJDCWGS】
    ' 2：室外车位属性表（SWCW）表中【CWLX】=“非机动车位“ ，按照【ZSXS】值进行统计汇总（面积*折算系数算出个数，汇总）
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "地上非机动车位个数与地上非机动车位个数一致性检查"
    CheckmodelName = "自定义脚本检查类->地上非机动车位个数与地上非机动车位个数一致性检查"
    strDescription = "地上非机动车位个数与地上非机动车位个数不一致"
    
    '获取地下机动车车位个数 DSFJDCWGS
    SqlStr = "Select JGSCHZXX.DXJDCWGS From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,DSFJDCWGSArr,SearchCount
    
    If SearchCount > 0 Then
        DSFJDCWGS = Transform(DSFJDCWGSArr(0))
    Else
        DSFJDCWGS = 0
    End If
    
    '获取室外车位个数 SWCWGS
    SqlStr = "Select GH_室外车位属性表.ID From GH_室外车位属性表 Inner Join GeoAreaTB On GH_室外车位属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And GH_室外车位属性表.CWLX = '非机动车位' "
    GetSQLRecordAll SqlStr,IDArr,IDCount
    
    If IDCount > 0 Then
        For i = 0 To IDCount - 1
            ZSXS = Transform(SSProcess.GetObjectAttr(IDArr(i),"[ZSXS]"))
            Area = Transform(SSProcess.GetObjectAttr(IDArr(i),"[MJ]"))
            SWCWGS = SWCWGS + Round(Area * ZSXS)
        Next 'i
    Else
        SWCWGS = 0
    End If
    
    If DSFJDCWGS - SWCWGS <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' DSFJDCWCheck

'地上非机动车位核实数量检查
Function DSFJDCHES()
    
    ' 1：室外车位属性表（SWCW）表中【CWLX】=“非机动车位“ ，面积【MJ】*折算系数【ZSXS】是否等于车位个数【CWGS】
    
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "地上非机动车位核实数量检查"
    CheckmodelName = "自定义脚本检查类->地上非机动车位核实数量检查"
    strDescription = "地上非机动车位核实数量不一致"
    
    SqlStr = "Select GH_室外车位属性表.ID From GH_室外车位属性表 Inner Join GeoAreaTB On GH_室外车位属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And GH_室外车位属性表.CWLX = '非机动车位' "
    GetSQLRecordAll SqlStr,IDArr,IDCount
    
    For i = 0 To IDCount - 1
        ZSXS = Transform(SSProcess.GetObjectAttr(IDArr(i),"[ZSXS]"))
        Area = Transform(SSProcess.GetObjectAttr(IDArr(i),"[MJ]"))
        CWGS = Transform(SSProcess.GetObjectAttr(IDArr(i),"[CWGS]"))
        If Round(Area * ZSXS) - CWGS <> 0  Then
            SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(IDArr(i),"SSObj_X"),SSProcess.GetObjectAttr(IDArr(i),"SSObj_Y"),0,2,IDArr(i),""
        End If
    Next 'i
    
End Function' DSFJDCHES

'地下非机动车位个数与地下非机动车位个是否一致
Function DXFJDCWCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【DXFJDCWGS】
    ' 2：室内车位属性表（SNCW）表中【CWLX】=“非机动车位“ ，按照【ZSXS】值进行汇总（面积*折算系数算出个数，汇总）
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "地下非机动车位个数与地下非机动车位个数一致性检查"
    CheckmodelName = "自定义脚本检查类->地下非机动车位个数与地下非机动车位个数一致性检查"
    strDescription = "地下非机动车位个数与地下非机动车位个数不一致"
    
    '获取地下机动车车位个数 DXFJDCWGS
    SqlStr = "Select JGSCHZXX.DXJDCWGS From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,DXFJDCWGSArr,SearchCount
    
    If SearchCount > 0 Then
        DXFJDCWGS = Transform(DXFJDCWGSArr(0))
    Else
        DXFJDCWGS = 0
    End If
    
    '获取室外车位个数 SNCWGS
    SqlStr = "Select GH_室内车位属性表.ID From GH_室内车位属性表 Inner Join GeoAreaTB On GH_室内车位属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And GH_室内车位属性表.CWLX = '非机动车位' "
    GetSQLRecordAll SqlStr,IDArr,IDCount
    
    If IDCount > 0 Then
        For i = 0 To IDCount - 1
            ZSXS = Transform(SSProcess.GetObjectAttr(IDArr(i),"[ZSXS]"))
            Area = Transform(SSProcess.GetObjectAttr(IDArr(i),"[MJ]"))
            SNCWGS = SNCWGS + Round(Area * ZSXS)
        Next 'i
    Else
        SNCWGS = 0
    End If
    
    If DXFJDCWGS - SNCWGS <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' DXFJDCWCheck

'地下非机动车位核实数量检查
Function DXFJDCHES()
    
    ' 1：室内车位属性表（SNCW）表中【CWLX】=“非机动车位“ ，面积【MJ】*折算系数【ZSXS】是否等于车位个数【CWGS】
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "地下非机动车位核实数量检查"
    CheckmodelName = "自定义脚本检查类->地下非机动车位核实数量检查"
    strDescription = "地下非机动车位核实数量不一致"
    
    SqlStr = "Select GH_室内车位属性表.ID From GH_室内车位属性表 Inner Join GeoAreaTB On GH_室内车位属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And GH_室内车位属性表.CWLX = '非机动车位' "
    GetSQLRecordAll SqlStr,IDArr,IDCount
    
    If IDCount > 0 Then
        For i = 0 To IDCount - 1
            ZSXS = Transform(SSProcess.GetObjectAttr(IDArr(i),"[ZSXS]"))
            Area = Transform(SSProcess.GetObjectAttr(IDArr(i),"[MJ]"))
            CWGS = Transform(SSProcess.GetObjectAttr(IDArr(i),"[CWGS]"))
            If Round(Area * ZSXS) - CWGS <> 0  Then
                SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(IDArr(i),"SSObj_X"),SSProcess.GetObjectAttr(IDArr(i),"SSObj_Y"),0,2,IDArr(i),""
            End If
        Next 'i
    End If
    
End Function' DXFJDCHES

'绿地总面积是否等于集中绿地面积+单块绿地面积面积和
Function LvDAreaCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【LDZMJ】=【JZLDMJ】+【DKLDMJ】
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "绿地总面积检查"
    CheckmodelName = "自定义脚本检查类->绿地总面积检查"
    strDescription = "绿地总面积与集中绿地和单块绿地面积之和不一致"
    
    '获取绿地总面积 LDZMJ
    SqlStr = "Select JGSCHZXX.LDZMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,LDZMJArr,LDCount
    
    If LDCount > 0 Then
        LDZMJ = Transform(LDZMJArr(0))
    Else
        LDZMJ = 0
    End If
    
    '获取集中绿地和单块绿地面积之和 SumArea
    SqlStr = "Select JGSCHZXX.JZLDMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,JZLDMJArr,JZLDCount
    
    If JZLDCount > 0 Then
        JZLDMJ = Transform(JZLDMJArr(0))
    Else
        JZLDMJ = 0
    End If
    
    SqlStr = "Select JGSCHZXX.DKLDMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,DKLDMJArr,DKLDCount
    
    If DKLDCount > 0 Then
        DKLDMJ = Transform(DKLDMJArr(0))
    Else
        DKLDMJ = 0
    End If
    
    SumArea = JZLDMJ + DKLDMJ
    
    If LDZMJ - SumArea <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' LvDAreaCheck

'单块绿地面积与单块绿地范围面面积汇总值是否一致
Function DKLVCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【DKLDMJ】
    ' 2：绿化划分信息表（LHHF）其中的【MC】=单块绿地，并通过【ID_LDK】绿地块ID与绿化要素属性表（LHYS）中的【ID_LDK】取【LHMJ】的汇总值
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "单块绿地面积与单块绿地范围面面积汇总值一致性检查"
    CheckmodelName = "自定义脚本检查类->单块绿地面积与单块绿地范围面面积汇总值一致性检查"
    strDescription = "单块绿地面积与单块绿地范围面面积汇总值不一致"
    
    '单块绿地总面积 DKLDMJ
    SqlStr = "Select JGSCHZXX.DKLDMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,DKLDMJArr,JZLDCount
    
    If JZLDCount > 0 Then
        DKLDMJ = Transform(DKLDMJArr(0))
    Else
        DKLDMJ = 0
    End If
    
    '汇总绿化面积 SumArea
    SqlStr = "Select LHHF.ID_LDK From LHHF Where LHHF.MC = '单块绿地' "
    GetSQLRecordAll SqlStr,IDArr,IDCount
    
    If IDCount > 0 Then
        For i = 0 To IDCount - 1
            SumArea = SumArea + Transform(SSProcess.GetObjectAttr(IDArr(i),"[LHMJ]"))
        Next 'i
    Else
        SumArea = 0
    End If
    
    
    If DKLDMJ - SumArea <> 0  Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' DKLVCheck

'集中绿地面积与集中绿地范围面面积汇总值是否一致
Function JZLDCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【JZLDMJ】
    ' 2：绿化划分信息表（LHHF）其中的【MC】=集中绿地，并通过【ID_LDK】绿地块ID与绿化要素属性表（LHYS）中的【ID_LDK】取【LHMJ】的汇总值
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "集中绿地面积与集中绿地范围面面积汇总值一致性检查"
    CheckmodelName = "自定义脚本检查类->集中绿地面积与集中绿地范围面面积汇总值一致性检查"
    strDescription = "集中绿地面积与集中绿地范围面面积汇总值不一致"
    
    '集中绿地面积 JZLDMJ
    SqlStr = "Select JGSCHZXX.JZLDMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,JZLDMJArr,JZLDCount
    
    If JZLDCount > 0 Then
        JZLDMJ = Transform(JZLDMJArr(0))
    Else
        JZLDMJ = 0
    End If
    
    '汇总绿化面积 SumArea
    SqlStr = "Select LHHF.ID_LDK From LHHF Where LHHF.MC = '集中绿地' "
    GetSQLRecordAll SqlStr,IDArr,IDCount
    
    If IDCount > 0 Then
        For i = 0 To IDCount - 1
            SumArea = SumArea + Transform(SSProcess.GetObjectAttr(IDArr(i),"[LHMJ]"))
        Next 'i
    Else
        SumArea = 0
    End If
    
    If DKLDMJ - SumArea <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' JZLDCheck

'登高场地个数与登高场地面个数是否一致
Function DGCDCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【RFZMJ】
    ' 2：人防功能区属性表（RFGNQ）中【JZMJ】值累加和
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "登高场地个数与登高场地面个数一致性检查"
    CheckmodelName = "自定义脚本检查类->登高场地个数与登高场地面个数一致性检查"
    strDescription = "登高场地个数与登高场地面个数不一致"
    
End Function' DGCDCheck

'人防总面积与人防功能区面积汇总值是否一致
Function RFMJCheck()
    
    ' 1：规划实测汇总信息表(JGSCHZXX)表中【RFZMJ】
    ' 2：人防功能区属性表（RFGNQ）中【JZMJ】值累加和
    
    '检查记录配置
    strGroupName = "图表一致性检查"
    strCheckName = "人防总面积与人防功能区面积汇总值一致性检查"
    CheckmodelName = "自定义脚本检查类->人防总面积与人防功能区面积汇总值一致性检查"
    strDescription = "人防总面积与人防功能区面积汇总值不一致"
    
    '获取人防总面积 RFZMJ
    SqlStr = "Select JGSCHZXX.RFZMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,RFZMJArr,JZLDCount
    
    If JZLDCount > 0 Then
        RFZMJ = Transform(RFZMJArr(0))
    Else
        RFZMJ = 0
    End If
    
    
    '汇总人防面积 SumArea
    SqlStr = "Select Sum(RF_人防功能区属性表.JZMJ) From RF_人防功能区属性表 Inner Join GeoAreaTB On RF_人防功能区属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 "
    GetSQLRecordAll SqlStr,SumAreaArr,SumCount
    
    If SumCount > 0 Then
        SumArea = Transform(SumAreaArr(0))
    Else
        SumArea = 0
    End If
    
    If RFZMJ - SumArea <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' RFMJCheck

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