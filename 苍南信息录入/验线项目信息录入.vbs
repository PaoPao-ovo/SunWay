'�����Ѵ���������Ϣ
Dim Info(10000,2)

'��ں���
Sub OnClick()
    GetFxInfo()
    SetNewInfo()
End Sub

'��ȡ���������Ա��ԭ�е���Ϣ
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
    
    Info(0,1) = "��Ŀ����"
    Info(1,1) = "��Ŀ��ַ"
    Info(2,1) = "���赥λ"
    Info(3,1) = "ί�е�λ"
    Info(4,1) = "��浥λ"
    Info(5,1) = "����ʱ��"
    Info(6,1) = "������Ŀʱ��"
    Info(7,1) = "����ʱ��"
    Info(8,1) = "��Ŀ������"
    Info(9,1) = "�������"
    Info(10,1) = "��Ŀ���"
    Info(11,1) = "���蹤�̹滮���֤��"
    Info(12,1) = "��Ƶ�λ"
    Info(13,1) = "�ܴ���"
    Info(14,1) = "������"
    Info(15,1) = "��ƽ��ͼ���̱��"
    Info(16,1) = "����ʱ��"
    Info(17,1) = "������Ŀʱ��"
    
End Function' GetFxInfo

'�����µķ��������Ա���Ϣ
Function SetNewInfo()
    SSProcess.ClearInputParameter
    For i = 0 To 17
        If Info(i,0) <> "" Then
            SSProcess.AddInputParameter Info(i,1) , Info(i,0) , 0 , "" , ""
        Else
            SSProcess.AddInputParameter Info(i,1) , "" , 0 , "" , ""
        End If
    Next 'i
    result = SSProcess.ShowInputParameterDlg ("��Ϣ¼��")
    
    xmmc = SSProcess.GetInputParameter ("��Ŀ����")
    xmdz = SSProcess.GetInputParameter ("��Ŀ��ַ")
    jsdw = SSProcess.GetInputParameter ("���赥λ")
    wtdw = SSProcess.GetInputParameter ("ί�е�λ")
    chdw = SSProcess.GetInputParameter ("��浥λ")
    fxsj = SSProcess.GetInputParameter ("����ʱ��")
    fxxmsj = SSProcess.GetInputParameter ("������Ŀʱ��")
    spsj = SSProcess.GetInputParameter ("����ʱ��")
    xmfzr = SSProcess.GetInputParameter ("��Ŀ������")
    bgbz = SSProcess.GetInputParameter ("�������")
    xmbh = SSProcess.GetInputParameter ("��Ŀ���")
    jsgcghxkzh = SSProcess.GetInputParameter ("���蹤�̹滮���֤��")
    sjdw = SSProcess.GetInputParameter ("��Ƶ�λ")
    zzs = SSProcess.GetInputParameter ("�ܴ���")
    psr = SSProcess.GetInputParameter ("������")
    zpmtgcbh = SSProcess.GetInputParameter ("��ƽ��ͼ���̱��")
    yxsj = SSProcess.GetInputParameter ("����ʱ��")
    yxxmsj = SSProcess.GetInputParameter ("������Ŀʱ��")
    
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