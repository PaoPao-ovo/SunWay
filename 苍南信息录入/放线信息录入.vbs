Function InfoPro()
    SSProcess.UpdateScriptDlgParameter 1
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9130224"
    SSProcess.SelectFilter
    hxcount = SSProcess.GetSelgeoCount
    If hxcount > 1 Then
        MsgBox "存在多个红线，请确认数据"
        Exit Function
    ElseIf hxcount = 1 Then
        SSProcess.ClearInputParameter
        SSProcess.AddInputParameter "项目名称" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "项目地址" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "建设单位" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "委托单位" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "测绘单位" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "放线时间" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "放线项目时间" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "审批时间" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "项目负责人" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "报告编制" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "项目编号" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "建设工程规划许可证号" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "设计单位" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "总幢数" , zzs , 0 , "" , ""
        SSProcess.AddInputParameter "拍摄人" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "总平面图工程编号" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "验线时间" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "验线项目时间" , "" , 0 , "" , ""
        
        result = SSProcess.ShowInputParameterDlg ("信息录入")
        
        If result = 1 Then
            zzs = ZZScount()
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
            
            objID = SSProcess.GetSelGeoValue(0, "SSObj_ID")
            'MsgBox objID
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
        End If
    Else
        MsgBox "不存在范围线"
        Exit Function
    End If
    SSProcess.RefreshView
End Function

Function ZZScount() '获取当前地物内的总幢数
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9130224"
    SSProcess.SelectFilter
    hxid = SSProcess.GetSelGeoValue(0,"SSObj_ID")
    'MsgBox hxid
    idString = SSProcess.SearchInnerObjIDs(hxid,2,"9310013",1)
    
    ZSArr = Split(idString,",", - 1,1)
    Count = UBound(ZSArr) + 1
    ZZScount = Count
End Function' ZZScount

Sub OnClick()
    '添加代码
    '窗口配置
    InfoPro()
End Sub