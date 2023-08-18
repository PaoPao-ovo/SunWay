'储存已存在属性信息
Dim Info(10000,2)

'入口函数
Sub OnClick()
    GetFxInfo()
    SetNewInfo()
End Sub

'获取放验线属性表的原有的信息
Function GetFxInfo()
    
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9130224"
    SSProcess.SelectFilter
    
    hxid = SSProcess.GetSelGeoValue(0,"SSObj_ID")
    xmmc = SSProcess.GetSelGeoValue (0,"[XiangMMC]")
    xmdz = SSProcess.GetSelGeoValue (0,"[XiangMDZ]")
    jsdw = SSProcess.GetSelGeoValue (0,"[JianSDW]")
    wtdw = SSProcess.GetSelGeoValue (0,"[WeiTDW]")
    chdw = SSProcess.GetSelGeoValue (0,"[CeHDW]")
    fxsj = SSProcess.GetSelGeoValue (0,"[FXDATE]")
    fxxmsj = SSProcess.GetSelGeoValue (0,"[FXXMDATE]")
    spsj = SSProcess.GetSelGeoValue (0,"[ShenPDATE]")
    xmfzr = SSProcess.GetSelGeoValue (0,"[XiangMFZR]")
    bgbz = SSProcess.GetSelGeoValue (0,"[BaoGBZ]")
    xmbh = SSProcess.GetSelGeoValue (0,"[XiangMBH]")
    jsgcghxkzh = SSProcess.GetSelGeoValue (0,"[GuiHXKZH]")
    sjdw = SSProcess.GetSelGeoValue (0,"[SheJDW]")
    zzs = SSProcess.GetSelGeoValue (0,"[ZongZS]")
    psr = SSProcess.GetSelGeoValue (0,"[PaiSR]")
    zpmtgcbh = SSProcess.GetSelGeoValue (0,"[ZongPMJTBH]")
    yxsj = SSProcess.GetSelGeoValue (0,"[YXDATE]")
    yxxmsj = SSProcess.GetSelGeoValue (0,"[YXXMDATE]")
    
    Info(0,0) = xmmc
    Info(1,0) = xmdz
    Info(2,0) = jsdw
    Info(3,0) = wtdw
    Info(4,0) = chdw
    Info(5,0) = fxsj
    Info(6,0) = fxxmsj
    Info(7,0) = spsj
    Info(8,0) = xmfzr
    Info(9,0) = bgbz
    Info(10,0) = xmbh
    Info(11,0) = jsgcghxkzh
    Info(12,0) = sjdw
    Info(13,0) = zzs
    Info(14,0) = psr
    Info(15,0) = zpmtgcbh
    Info(16,0) = yxsj
    Info(17,0) = yxxmsj
    Info(18,0) = hxid
    
    Info(0,1) = "项目名称"
    Info(1,1) = "项目地址"
    Info(2,1) = "建设单位"
    Info(3,1) = "委托单位"
    Info(4,1) = "测绘单位"
    Info(5,1) = "放线时间"
    Info(6,1) = "放线项目时间"
    Info(7,1) = "审批时间"
    Info(8,1) = "项目负责人"
    Info(9,1) = "报告编制"
    Info(10,1) = "项目编号"
    Info(11,1) = "建设工程规划许可证号"
    Info(12,1) = "设计单位"
    Info(13,1) = "总幢数"
    Info(14,1) = "拍摄人"
    Info(15,1) = "总平面图工程编号"
    Info(16,1) = "验线时间"
    Info(17,1) = "验线项目时间"
    
End Function' GetFxInfo

'配置新的放验线属性表信息
Function SetNewInfo()
    SSProcess.ClearInputParameter
    For i = 0 To 17
        If Info(i,0) <> "" Then
            SSProcess.AddInputParameter Info(i,1) , Info(i,0) , 0 , "" , ""
        Else
            SSProcess.AddInputParameter Info(i,1) , "" , 0 , "" , ""
        End If
    Next 'i
    result = SSProcess.ShowInputParameterDlg ("信息录入")
    
    xmmc = SSProcess.GetInputParameter ("项目名称")
    xmdz = SSProcess.GetInputParameter ("项目地址")
    jsdw = SSProcess.GetInputParameter ("建设单位")
    wtdw = SSProcess.GetInputParameter ("委托单位")
    chdw = SSProcess.GetInputParameter ("测绘单位")
    fxsj = SSProcess.GetInputParameter ("放线时间")
    fxxmsj = SSProcess.GetInputParameter ("放线项目时间")
    spsj = SSProcess.GetInputParameter ("审批时间")
    xmfzr = SSProcess.GetInputParameter ("项目负责人")
    bgbz = SSProcess.GetInputParameter ("报告编制")
    xmbh = SSProcess.GetInputParameter ("项目编号")
    jsgcghxkzh = SSProcess.GetInputParameter ("建设工程规划许可证号")
    sjdw = SSProcess.GetInputParameter ("设计单位")
    zzs = SSProcess.GetInputParameter ("总幢数")
    psr = SSProcess.GetInputParameter ("拍摄人")
    zpmtgcbh = SSProcess.GetInputParameter ("总平面图工程编号")
    yxsj = SSProcess.GetInputParameter ("验线时间")
    yxxmsj = SSProcess.GetInputParameter ("验线项目时间")
    
    objID = Info(18,0)

    SSProcess.SetObjectAttr objID, "[XiangMMC]", xmmc
    SSProcess.SetObjectAttr objID, "[XiangMDZ]", xmdz
    SSProcess.SetObjectAttr objID, "[JianSDW]", jsdw
    SSProcess.SetObjectAttr objID, "[WeiTDW]", wtdw
    SSProcess.SetObjectAttr objID, "[CeHDW]", chdw
    SSProcess.SetObjectAttr objID, "[FXDATE]", fxsj
    SSProcess.SetObjectAttr objID, "[FXXMDATE]", fxxmsj
    SSProcess.SetObjectAttr objID, "[ShenPDATE]", spsj
    SSProcess.SetObjectAttr objID, "[XiangMFZR]", xmfzr
    SSProcess.SetObjectAttr objID, "[BaoGBZ]", bgbz
    SSProcess.SetObjectAttr objID, "[XiangMBH]", xmbh
    SSProcess.SetObjectAttr objID, "[GuiHXKZH]", jsgcghxkzh
    SSProcess.SetObjectAttr objID, "[SheJDW]", sjdw
    SSProcess.SetObjectAttr objID, "[ZongZS]", zzs
    SSProcess.SetObjectAttr objID, "[PaiSR]", psr
    SSProcess.SetObjectAttr objID, "[ZongPMJTBH]", zpmtgcbh
    SSProcess.SetObjectAttr objID, "[YXDATE]", yxsj
    SSProcess.SetObjectAttr objID, "[YXXMDATE]", yxxmsj
    
End Function' SetNewInfo