Function InfoPro()
    SSProcess.UpdateScriptDlgParameter 1
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9130224"
    SSProcess.SelectFilter
    hxcount = SSProcess.GetSelgeoCount
    If hxcount > 1 Then
        MsgBox "���ڶ�����ߣ���ȷ������"
        Exit Function
    ElseIf hxcount = 1 Then
        SSProcess.ClearInputParameter
        SSProcess.AddInputParameter "��Ŀ����" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "��Ŀ��ַ" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "���赥λ" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "ί�е�λ" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "��浥λ" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "����ʱ��" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "������Ŀʱ��" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "����ʱ��" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "��Ŀ������" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "�������" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "��Ŀ���" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "���蹤�̹滮���֤��" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "��Ƶ�λ" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "�ܴ���" , zzs , 0 , "" , ""
        SSProcess.AddInputParameter "������" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "��ƽ��ͼ���̱��" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "����ʱ��" , "" , 0 , "" , ""
        SSProcess.AddInputParameter "������Ŀʱ��" , "" , 0 , "" , ""
        
        result = SSProcess.ShowInputParameterDlg ("��Ϣ¼��")
        
        If result = 1 Then
            zzs = ZZScount()
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
        MsgBox "�����ڷ�Χ��"
        Exit Function
    End If
    SSProcess.RefreshView
End Function

Function ZZScount() '��ȡ��ǰ�����ڵ��ܴ���
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
    '��Ӵ���
    '��������
    InfoPro()
End Sub