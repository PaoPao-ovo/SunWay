'docȫ�ֶ���
Dim g_docObj

'·����������
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

'�洢��������
Dim HXarr(1000,2)

'������������
Dim RowBS
RowBS = 0

'��2�����������
Dim Row1arr(1000,2)

'��3������������
Dim Row2arr(1000,2)

'����
Dim Tablecount
Tablecount = 0

'�����ϵ��
Dim DisPoi(1000)

'�������ߺ�
Dim DisLine(1000)

'�����Ͽ��Ƶ��ʶ
Dim DifKzPoi(1000)

'�������������߱��
Dim DifZFL(1000)

'����Ŀ��Ƶ����
Dim KzPoiCount
KzPoiCount = 0

'��ں���
Sub OnClick()
    allvisible()
    strTempFileName = "���߲�������ģ��.doc"
    strTempFilePath = SSProcess.GetSysPathName (7) & "���ģ��\" & strTempFileName
    Set g_docObj = CreateObject ("asposewordscom.asposewordshelper")
    If  TypeName (g_docObj) = "AsposeWordsHelper" Then
        g_docObj.CreateDocumentByTemplate strTempFilePath
    Else
        MsgBox "����ע��Aspose.Word���"
        Exit Sub
    End If
    
    pathName = GetFilePath()
    
    proname = GetFileName()
    
    ReplaceValue()
    
    SetKZD()
    
    SetInfoTable()
    
    CopyTable()
    
    SetPosition()
    ZFL()
    Set4Line()
    SetResultTable()
    'InsertPhoto()
    
    strFileSavePath = pathName & proname
    'MsgBox strFileSavePath
    g_docObj.SaveEx  strFileSavePath
    
End Sub

'//��ȡ�ɹ�Ŀ¼·��
Function  GetFilePath()
    filePath = SSProcess.GetSysPathName(5)
    filePath = filePath & "3�ɹ�" & "\"
    ' filePath = filePath & "\"
    GetFilePath = filePath
End Function

'//��ȡ�ɹ���������
Function  GetFileName()
    proname = GetProName()
    GetFileName = proname & "���߲�������.doc"
End Function


'��ȡ��ǰ���ߵ���Ŀ����
Function GetProName()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9130224"
    SSProcess.SelectFilter
    hxcount = SSProcess.GetSelgeoCount
    If hxcount = 1 Then xmmc = SSProcess.GetSelGeoValue (0,"[XiangMMC]")
    GetProName = xmmc
End Function' GetProName

'�ַ��滻 
Function ReplaceValue()
    hxid = SSProcess.GetSelGeoValue(0,"SSObj_ID")
    xmmc = SSProcess.GetSelGeoValue (0,"[XiangMMC]")
    xmdz = SSProcess.GetSelGeoValue (0,"[XiangMDZ]")
    jsdw = SSProcess.GetSelGeoValue (0,"[JianSDW]")
    wtdw = SSProcess.GetSelGeoValue (0,"[WeiTDW]")
    chdw = SSProcess.GetSelGeoValue (0,"[CeHDW]")
    ' fxsj = SSProcess.GetSelGeoValue (0,"[FXDATE]")
    ' fxxmsj = SSProcess.GetSelGeoValue (0,"[FXXMDATE]")
    ' spsj = SSProcess.GetSelGeoValue (0,"[ShenPDATE]")
    xmfzr = SSProcess.GetSelGeoValue (0,"[XiangMFZR]")
    bgbz = SSProcess.GetSelGeoValue (0,"[BaoGBZ]")
    ' xmbh = SSProcess.GetSelGeoValue (0,"[XiangMBH]")
    jsgcghxkzh = SSProcess.GetSelGeoValue (0,"[GuiHXKZH]")
    sjdw = SSProcess.GetSelGeoValue (0,"[SheJDW]")
    zzs = SSProcess.GetSelGeoValue (0,"[ZongZS]")
    psr = SSProcess.GetSelGeoValue (0,"[PaiSR]")
    ' zpmtgcbh = SSProcess.GetSelGeoValue (0,"[ZongPMJTBH]")
    yxsj = SSProcess.GetSelGeoValue (0,"[YXDATE]")
    yxxmsj = SSProcess.GetSelGeoValue (0,"[YXXMDATE]")
    
    HXarr(0,0) = xmmc
    HXarr(1,0) = xmdz
    HXarr(2,0) = sjdw
    HXarr(3,0) = jsdw
    HXarr(4,0) = wtdw
    HXarr(5,0) = chdw
    HXarr(6,0) = yxsj
    HXarr(7,0) = yxxmsj
    HXarr(8,0) = xmfzr
    HXarr(9,0) = bgbz
    HXarr(10,0) = jsgcghxkzh
    HXarr(11,0) = zzs
    HXarr(12,0) = psr
    HXarr(13,0) = hxid
    
    strFields = "XiangMMC,XiangMDZ,SheJDW,JianSDW,WeiTDW,CeHDW,YXDATE,YXXMDATE,XiangMFZR,BaoGBZ,GuiHXKZH,ZongZS,PaiSR"
    strarr = Split(strFields,",", - 1,1)
    
    For i = 0 To UBound(strarr)
        g_docObj.Replace "{" & strarr(i) & "}",HXarr(i,0),0
    Next 'i
    
End Function

'���Ʊ�
Function CopyTable()
    zzs = HXarr(11,0) '��Χ���ڵĽ������ܸ���
    zzs = transform(zzs)
    i = 1
    bulidname = GetBuildingName()
    'MsgBox bulidname
    bulidarr = Split(bulidname,",", - 1,1)
    text = bulidarr(0) & "���蹤�̹滮���߳ɹ���"
    g_docObj.SetCellText 4,0,0,text,True,False
    While i <= zzs - 1
        g_docObj.CloneTable 4,2,0,False
        text = bulidarr(i) & "���蹤�̹滮���߳ɹ���"
        g_docObj.SetCellText 4 + i,0,0,text,True,False
        i = i + 1
    WEnd
    Tablecount = zzs
    Tablecount = CInt(Tablecount)
End Function' CopyTable

'��ȡ��ǰ�Ľ���������
Function GetBuildingName()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9130511" 'ʵ��
    SSProcess.SelectFilter
    poicount = SSProcess.GetSelgeoCount
    poistring = ""
    For i = 0 To poicount - 1
        poiname = SSProcess.GetSelGeoValue(i,"[JianZWMC]")
        If poistring = "" Then
            poistring = poiname
        ElseIf Replace(poistring,poiname,"") = poistring Then
            poistring = poistring & "," & poiname
        End If
    Next 'i
    GetBuildingName = poistring
End Function' GetBuildingName

'�������Ƶ㲢��ֵ
Function SetKZD()
    Dim LLarr(1000,4)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9130211,9130212,1103021,1102021"
    SSProcess.SelectFilter
    llcount = SSProcess.GetSelgeoCount
    'MsgBox llcount
    Dim row
    row = 4
    Dim poiname
    poiname = ""
    'MsgBox poicount
    For i = 0 To llcount - 1
        x = SSProcess.GetSelGeoValue(i,"SSObj_X")
        y = SSProcess.GetSelGeoValue(i,"SSObj_Y")
        z = SSProcess.GetSelGeoValue(i,"SSObj_Z")
        name = SSProcess.GetSelGeoValue(i,"SSObj_PointName")
        
        x = FormatNumber(transform(x),3)
        y = FormatNumber(transform(y),3)
        z = FormatNumber(transform(z),3)
        
        LLarr(i,0) = x
        LLarr(i,1) = y
        LLarr(i,2) = z
        LLarr(i,3) = name
    Next 'i
    
    For j = 0 To llcount - 1
        If poiname = "" Then
            poiname = LLarr(j,3)
        Else
            poiname = LLarr(j,3) & "," & poiname
        End If
    Next 'j
    
    Dim SCarr(1000,4)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9130311,9130312,9130512,9130412"
    SSProcess.SetSelectCondition "SSObj_PointName", "==",poiname
    SSProcess.SelectFilter
    poicount = SSProcess.GetSelgeoCount
    KzPoiCount = poicount
    'MsgBox poicount
    For i = 0 To poicount - 1
        x = SSProcess.GetSelGeoValue(i,"SSObj_X")
        y = SSProcess.GetSelGeoValue(i,"SSObj_Y")
        z = SSProcess.GetSelGeoValue(i,"SSObj_Z")
        name = SSProcess.GetSelGeoValue(i,"SSObj_PointName")
        
        x = FormatNumber(transform(x),3)
        y = FormatNumber(transform(y),3)
        z = FormatNumber(transform(z),3)
        
        SCarr(i,0) = x
        SCarr(i,1) = y
        SCarr(i,2) = z
        SCarr(i,3) = name
        'MsgBox SCarr(i,3) 
    Next 'i
    Dim k
    k = 0
    count = 0
    If poicount > 3 Then
        g_docObj.CloneTableRow 4, 4, 1,poicount - 3, False
        For j = 0 To llcount - 1
            For i = 0 To poicount - 1
                If LLarr(j,3) = SCarr(i,3) Then
                    'MsgBox LLarr(j,1)
                    Diffxy = GetLengthDiff(LLarr(j,0),LLarr(j,1),SCarr(i,0),SCarr(i,1)) * 1000
                    Diffh = Abs(LLarr(j,2) - SCarr(i,2)) * 1000
                    Diffh = FormatNumber(Diffh,3)
                    g_docObj.SetCellText 4,row + k,0,LLarr(j,1),True,False
                    g_docObj.SetCellText 4,row + k,1,LLarr(j,0),True,False
                    g_docObj.SetCellText 4,row + k,2,LLarr(j,2),True,False
                    g_docObj.SetCellText 4,row + k,3,SCarr(i,1),True,False
                    g_docObj.SetCellText 4,row + k,4,SCarr(i,0),True,False
                    g_docObj.SetCellText 4,row + k,5,SCarr(i,2),True,False
                    g_docObj.SetCellText 4,row + k,6,Diffxy,True,False
                    g_docObj.SetCellText 4,row + k,7,Diffh,True,False
                    g_docObj.SetCellText 4,row + k,8,50,True,False
                    g_docObj.SetCellText 4,row + k,9,30,True,False
                    'MsgBox count
                    If Diffxy > 50 Then
                        g_docObj.SetCellText 4,row + k,10,"������",True,False
                        DifKzPoi(count) = SCarr(i,3)
                    Else
                        g_docObj.SetCellText 4,row + k,10,"����",True,False
                    End If
                    If Diffh > 30 Then
                        g_docObj.SetCellText 4,row + k,11,"������",True,False
                        DifKzPoi(count) = SCarr(i,3)
                    Else
                        g_docObj.SetCellText 4,row + k,11,"����",True,False
                    End If
                    k = k + 1
                    count = count + 1
                End If
            Next
        Next
        RowBS = poicount - 3
    Else
        For j = 0 To llcount - 1
            For i = 0 To poicount - 1
                If LLarr(j,3) = SCarr(i,3) Then
                    'MsgBox LLarr(j,1)
                    Diffxy = GetLengthDiff(LLarr(j,0),LLarr(j,1),SCarr(i,0),SCarr(i,1)) * 1000
                    Diffh = Abs(LLarr(j,2) - SCarr(i,2)) * 1000
                    Diffh = FormatNumber(Diffh,3)
                    g_docObj.SetCellText 4,row + k,0,LLarr(j,1),True,False
                    g_docObj.SetCellText 4,row + k,1,LLarr(j,0),True,False
                    g_docObj.SetCellText 4,row + k,2,LLarr(j,2),True,False
                    g_docObj.SetCellText 4,row + k,3,SCarr(i,1),True,False
                    g_docObj.SetCellText 4,row + k,4,SCarr(i,0),True,False
                    g_docObj.SetCellText 4,row + k,5,SCarr(i,2),True,False
                    g_docObj.SetCellText 4,row + k,6,Diffxy,True,False
                    g_docObj.SetCellText 4,row + k,7,Diffh,True,False
                    g_docObj.SetCellText 4,row + k,8,50,True,False
                    g_docObj.SetCellText 4,row + k,9,30,True,False
                    If Diffxy > 50 Then
                        g_docObj.SetCellText 4,row + k,10,"������",True,False
                        DifKzPoi(count) = SCarr(i,3)
                    Else
                        g_docObj.SetCellText 4,row + k,10,"����",True,False
                    End If
                    If Diffh > 30 Then
                        g_docObj.SetCellText 4,row + k,11,"������",True,False
                        DifKzPoi(count) = SCarr(i,3)
                    Else
                        g_docObj.SetCellText 4,row + k,11,"����",True,False
                    End If
                    k = k + 1
                    count = count + 1
                End If
            Next
        Next
    End If
    
End Function' SetKZD

'�����������߽ϲ�
Function SetPosition()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9316710"
    SSProcess.SelectFilter
    DiffLineCount = SSProcess.GetSelgeoCount()
    Jzwname = ""
    For i = 0 To DiffLineCount - 1
        id = SSProcess.GetSelGeoValue(i,"SSObj_ID")
        name = SSProcess.GetSelGeoValue(i,"[JianZWMC]")
        If Jzwname = "" Then
            Jzwname = name
        ElseIf Replace(Jzwname,name,"") = Jzwname Then
            Jzwname = Jzwname & "," & name
        End If
    Next 'i
    'MsgBox Jzwname
    namearr = Split(Jzwname,",", - 1,1)
    count = 0
    For i = 0 To UBound(namearr)
        SelDiffLine "9316710",namearr(i)
        'MsgBox namearr(i)
        SelCount = SSProcess.GetSelgeoCount()
        'MsgBox SelCount
        ReDim Pcarr(SelCount,6)
        For k = 0 To SelCount - 1
            llx = SSProcess.GetSelGeoValue(k,"[llzbx]")
            lly = SSProcess.GetSelGeoValue(k,"[llzby]")
            scx = SSProcess.GetSelGeoValue(k,"[sczbx]")
            scy = SSProcess.GetSelGeoValue(k,"[sczby]")
            pc = SSProcess.GetSelGeoValue(k,"[pcjl]")
            dh = SSProcess.GetSelGeoValue(k,"[dh]")
            Pcarr(k,0) = llx
            Pcarr(k,1) = lly
            Pcarr(k,2) = scx
            Pcarr(k,3) = scy
            Pcarr(k,4) = pc
            Pcarr(k,5) = dh
        Next 'k
        
        For j = 4 To Tablecount + 3
            TitleName = g_docObj.GetCellText(j,0,0,False)
            'MsgBox TitleName
            Title = Replace(TitleName,"���蹤�̹滮���߳ɹ���","")
            totallen = Len(Title)
            Title = Left(Title,totallen - 1)
            'MsgBox namearr(i)
            If namearr(i) = Title Then Tableindex = j
        Next 'j
        'MsgBox Tableindex
        If SelCount <= 4 Then
            For m = 0 To SelCount - 1
                g_docObj.SetCellText Tableindex,10 + RowBS + m,0,Pcarr(m,0),True,False
                g_docObj.SetCellText Tableindex,10 + RowBS + m,1,Pcarr(m,1),True,False
                g_docObj.SetCellText Tableindex,10 + RowBS + m,2,Pcarr(m,2),True,False
                g_docObj.SetCellText Tableindex,10 + RowBS + m,3,Pcarr(m,3),True,False
                g_docObj.SetCellText Tableindex,10 + RowBS + m,4,Pcarr(m,4),True,False
                g_docObj.SetCellText Tableindex,10 + RowBS + m,5,50,True,False
                If Pcarr(m,4) * 1000 > 50 Then
                    g_docObj.SetCellText Tableindex,10 + RowBS + m,6,"������",True,False
                    If DisPoi(count) = "" Then
                        DisPoi(count) = Pcarr(m,5)
                    Else
                        DisPoi(count) = DisPoi(count) & "," & Pcarr(m,5)
                    End If
                Else
                    g_docObj.SetCellText Tableindex,10 + RowBS + m,6,"����",True,False
                End If
            Next 'm
        End If
        
        If SelCount > 4 Then
            g_docObj.CloneTableRow Tableindex, 11, 1,SelCount - 4, False
            Row1arr(i,0) = Tableindex
            Row1arr(i,1) = SelCount - 4
            'MsgBox Row1arr(i,0)
            'MsgBox Row1arr(i,1)
            For m = 0 To SelCount - 1
                g_docObj.SetCellText Tableindex,10 + RowBS + m,0,Pcarr(m,0),True,False
                g_docObj.SetCellText Tableindex,10 + RowBS + m,1,Pcarr(m,1),True,False
                g_docObj.SetCellText Tableindex,10 + RowBS + m,2,Pcarr(m,2),True,False
                g_docObj.SetCellText Tableindex,10 + RowBS + m,3,Pcarr(m,3),True,False
                g_docObj.SetCellText Tableindex,10 + RowBS + m,4,Pcarr(m,4),True,False
                g_docObj.SetCellText Tableindex,10 + RowBS + m,5,50,True,False
                If Pcarr(m,4) * 1000 > 50 Then
                    g_docObj.SetCellText Tableindex,10 + RowBS + m,6,"������",True,False
                    If DisPoi(count) = "" Then
                        DisPoi(count) = Pcarr(m,5)
                    Else
                        DisPoi(count) = DisPoi(count) & "," & Pcarr(m,5)
                    End If
                Else
                    g_docObj.SetCellText Tableindex,10 + RowBS + m,6,"����",True,False
                End If
            Next
            
        End If
        count = count + 1
    Next 'i
End Function' SetPosition

'�����������߲���ֵ
Function ZFL()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9310013"
    SSProcess.SelectFilter
    JzwCount = SSProcess.GetSelgeoCount()
    count = 0
    'MsgBox RowBS+15
    For i = 0 To JzwCount - 1
        id = SSProcess.GetSelGeoValue(i,"SSObj_ID")
        jzwname = SSProcess.GetSelGeoValue(i,"[JianZWMC]")
        ids = SSProcess.SearchInnerObjIDs(id,0,"9130611",0)
        'MsgBox ids
        idsarr = Split(ids,",", - 1,1)
        ZFLcount = UBound(idsarr)
        
        'MsgBox ZFLcount
        '��ȡ��������
        For j = 4 To Tablecount + 3
            TitleName = g_docObj.GetCellText(j,0,0,False)
            'MsgBox TitleName
            Title = Replace(TitleName,"���蹤�̹滮���߳ɹ���","")
            totallen = Len(Title)
            Title = Left(Title,totallen - 1)
            'MsgBox linestrarr(i)
            If jzwname = Title Then Tableindex = j
        Next
        
        If ids <> "" Then
            For k = 0 To ZFLcount
                If ZFLcount <= 1 Then
                    sjgc = SSProcess.GetObjectAttr(idsarr(k),"[SheJGC]")
                    yxgc = SSProcess.GetObjectAttr(idsarr(k),"[YanXGC]")
                    jzwmc = SSProcess.GetObjectAttr(idsarr(k),"[JianZWMC]")
                    sjgc = Round(sjgc,3)
                    yxgc = Round(yxgc,3)
                    Diffh = Abs(sjgc - yxgc)
                    Diffh = FormatNumber(Diffh,3)
                    For m = 0 To Tablecount - 1
                        If Row1arr(m,0) = Tableindex Then
                            g_docObj.SetCellText Tableindex,16 + RowBS + k + Row1arr(m,1),0,sjgc,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + k + Row1arr(m,1),1,yxgc,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + k + Row1arr(m,1),2,Diffh,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + k + Row1arr(m,1),3,30,True,False
                            If Diffh * 1000 > 30 Then
                                g_docObj.SetCellText Tableindex,16 + RowBS + k + Row1arr(m,1),4,"������",True,False
                                If DifZFL(count) = "" Then
                                    DifZFL(count) = jzwname
                                Else
                                    DifZFL(count) = DifZFL(count) & "," & jzwname
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,16 + RowBS + k + Row1arr(m,1),4,"����",True,False
                            End If
                        Else
                            g_docObj.SetCellText Tableindex,16 + RowBS + k,0,sjgc,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + k,1,yxgc,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + k,2,Diffh,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + k,3,30,True,False
                            If Diffh * 1000 > 30 Then
                                g_docObj.SetCellText Tableindex,16 + RowBS + k,4,"������",True,False
                                If DifZFL(count) = "" Then
                                    DifZFL(count) = jzwname
                                Else
                                    DifZFL(count) = DifZFL(count) & "," & jzwname
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,16 + RowBS + k,4,"����",True,False
                            End If
                        End If
                    Next
                End If
                
                If ZFLcount >= 2 Then
                    g_docObj.CloneTableRow Tableindex, 16 + RowBS + k + Row1arr(k,1), 1,ZFLcount - 1 , False '��������
                    sjgc = SSProcess.GetObjectAttr(idsarr(k),"[SheJGC]")
                    yxgc = SSProcess.GetObjectAttr(idsarr(k),"[YanXGC]")
                    sjgc = Round(sjgc,3)
                    yxgc = Round(yxgc,3)
                    Diffh = Abs(sjgc - yxgc)
                    Diffh = FormatNumber(Diffh,3)
                    For m = 0 To Tablecount - 1
                        If Row1arr(m,0) = Tableindex Then
                            g_docObj.SetCellText Tableindex,16 + RowBS + k + Row1arr(m,1),0,sjgc,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + k + Row1arr(m,1),1,yxgc,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + k + Row1arr(m,1),2,Diffh,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + k + Row1arr(m,1),3,30,True,False
                            If Diffh * 1000 > 30 Then
                                g_docObj.SetCellText Tableindex,16 + RowBS + k + Row1arr(m,1),4,"������",True,False
                                If DifZFL(count) = "" Then
                                    DifZFL(count) = jzwname
                                Else
                                    DifZFL(count) = DifZFL(count) & "," & jzwname
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,16 + RowBS + k + Row1arr(m,1),4,"����",True,False
                            End If
                            Row2arr(k,0) = Tableindex
                            Row2arr(k,1) = ZFLcount - 1 + Row1arr(k,1)
                        Else
                            g_docObj.SetCellText Tableindex,16 + RowBS + k,0,sjgc,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + k,1,yxgc,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + k,2,Diffh,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + k,3,30,True,False
                            If Diffh * 1000 > 30 Then
                                g_docObj.SetCellText Tableindex,16 + RowBS + k,4,"������",True,False
                                If DifZFL(count) = "" Then
                                    DifZFL(count) = jzwname
                                Else
                                    DifZFL(count) = DifZFL(count) & "," & jzwname
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,16 + RowBS + k,4,"����",True,False
                            End If
                            Row2arr(k,0) = Tableindex
                            Row2arr(k,1) = ZFLcount - 1
                        End If
                    Next 'm
                    'MsgBox Row2arr(k,1)
                End If
            Next 'k
        End If
        'MsgBox JzwBS
        If ids = ""  Then
            SSProcess.ClearSelection
            SSProcess.ClearSelectCondition
            SSProcess.SetSelectCondition "SSObj_Code", "==", "9130224"
            SSProcess.SelectFilter
            For m = 0 To 0
                id = SSProcess.GetSelGeoValue(0,"SSObj_ID")
                ids = SSProcess.SearchInnerObjIDs(id,0,"9130611",0)
                idsarr = Split(ids,",", - 1,1)
                ZFLcount = UBound(idsarr)
                If ids <> "" Then
                    sjgc = SSProcess.GetObjectAttr(idsarr(0),"[SheJGC]")
                    yxgc = SSProcess.GetObjectAttr(idsarr(0),"[YanXGC]")
                    sjgc = Round(sjgc,3)
                    yxgc = Round(yxgc,3)
                    Diffh = Abs(sjgc - yxgc)
                    Diffh = FormatNumber(Diffh,3)
                    For n = 0 To Tablecount - 1
                        If Row1arr(n,0) = Tableindex Then
                            'MsgBox Row1arr
                            g_docObj.SetCellText Tableindex,16 + RowBS + Row1arr(n,1),0,sjgc,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + Row1arr(n,1),1,yxgc,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + Row1arr(n,1),2,Diffh,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + Row1arr(n,1),3,30,True,False
                            If Diffh * 1000 > 30 Then
                                g_docObj.SetCellText Tableindex,16 + RowBS + Row1arr(n,1),4,"������",True,False
                                If DifZFL(count) = "" Then
                                    DifZFL(count) = Tableindex
                                Else
                                    DifZFL(count) = DifZFL(count) & "," & Tableindex
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,16 + RowBS + Row1arr(n,1),4,"����",True,False
                            End If
                            
                        Else
                            g_docObj.SetCellText Tableindex,16 + RowBS ,0,sjgc,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS ,1,yxgc,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS ,2,Diffh,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS ,3,30,True,False
                            If Diffh * 1000 > 30 Then
                                g_docObj.SetCellText Tableindex,16 + RowBS ,4,"������",True,False
                                If DifZFL(count) = "" Then
                                    DifZFL(count) = Tableindex
                                Else
                                    DifZFL(count) = DifZFL(count) & "," & Tableindex
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,16 + RowBS ,4,"����",True,False
                            End If
                        End If
                    Next 'n
                End If
            Next
        End If
        count = count + 1
    Next
End Function' ZFL


'���������߳�
Function Set4Line()
    count = 0
    For i = 4 To Tablecount + 3
        TitleName = g_docObj.GetCellText(i,0,0,False)
        Title = Replace(TitleName,"���蹤�̹滮���߳ɹ���","")
        totallen = Len(Title)
        Title = Left(Title,totallen - 1)
        'MsgBox Title
        SelYxBc Title
        SelCount = SSProcess.GetSelgeoCount()
        'MsgBox SelCount
        Tableindex = i
        If SelCount <= 4 Then
            'MsgBox SelCount
            'For k = 0 To Tablecount - 1
            k = i - 4
            'MsgBox k
            If Row1arr(k,0) = Tableindex And Row2arr(k,0) = Tableindex  Then
                For l = 0 To SelCount - 1
                    bh = SSProcess.GetSelGeoValue(l,"[BH]")
                    fxbc = SSProcess.GetSelGeoValue(l,"[FXBC]")
                    yxbc = SSProcess.GetSelGeoValue(l,"[YXBC]")
                    bcjc = SSProcess.GetSelGeoValue(l,"[BCJC]")
                    bcjc = transform(bcjc)
                    tj = bcjc * 1000
                    temp = l Mod 2
                    hs = Int(l / 2) + 20 + Row1arr(k,1) + Row2arr(k,1) + RowBS
                    'MsgBox hs
                    If temp = 0 Then
                        g_docObj.SetCellText Tableindex,hs,0,bh,True,False
                        g_docObj.SetCellText Tableindex,hs,1,fxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,2,yxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,3,bcjc,True,False
                        If tj > 50 Then
                            content = g_docObj.GetCellText(Tableindex,hs,8,False)
                            contentlen = Len(content)
                            content = Left(content,contentlen - 1)
                            If content = "" Or content = "����" Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            End If
                        End If
                    Else
                        g_docObj.SetCellText Tableindex,hs,4,bh,True,False
                        g_docObj.SetCellText Tableindex,hs,5,fxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,6,yxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,7,bcjc,True,False
                        If tj > 50 Then
                            content = g_docObj.GetCellText(Tableindex,hs,8,False)
                            contentlen = Len(content)
                            content = Left(content,contentlen - 1)
                            If content = "" Or content = "����" Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            End If
                        End If
                    End If
                    tj1 = g_docObj.GetCellText(Tableindex,hs,8,False)
                    tjlen1 = Len(tj1)
                    tj1 = Left(tj1,tjlen1 - 1)
                    If tj1 = "" Then
                        g_docObj.SetCellText Tableindex,hs,8,"����",True,False
                    End If
                Next 'l
            ElseIf Row1arr(k,0) = Tableindex And Row2arr(k,0) <> Tableindex Then
                For l = 0 To SelCount - 1
                    bh = SSProcess.GetSelGeoValue(l,"[BH]")
                    fxbc = SSProcess.GetSelGeoValue(l,"[FXBC]")
                    yxbc = SSProcess.GetSelGeoValue(l,"[YXBC]")
                    bcjc = SSProcess.GetSelGeoValue(l,"[BCJC]")
                    bcjc = transform(bcjc)
                    tj = bcjc * 1000
                    temp = l Mod 2
                    hs = Int(l / 2) + 20 + Row1arr(k,1) + RowBS
                    If temp = 0 Then
                        g_docObj.SetCellText Tableindex,hs,0,bh,True,False
                        g_docObj.SetCellText Tableindex,hs,1,fxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,2,yxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,3,bcjc,True,False
                        If tj > 50 Then
                            content = g_docObj.GetCellText(Tableindex,hs,8,False)
                            contentlen = Len(content)
                            content = Left(content,contentlen - 1)
                            If content = "" Or content = "����" Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            End If
                        End If
                    Else
                        g_docObj.SetCellText Tableindex,hs,4,bh,True,False
                        g_docObj.SetCellText Tableindex,hs,5,fxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,6,yxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,7,bcjc,True,False
                        If tj > 50 Then
                            content = g_docObj.GetCellText(Tableindex,hs,8,False)
                            contentlen = Len(content)
                            content = Left(content,contentlen - 1)
                            'MsgBox content
                            If content = "" Or content = "����" Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            End If
                        End If
                    End If
                    
                    tj1 = g_docObj.GetCellText(Tableindex,hs,8,False)
                    tjlen1 = Len(tj1)
                    tj1 = Left(tj1,tjlen1 - 1)
                    If tj1 = "" Then
                        g_docObj.SetCellText Tableindex,hs,8,"����",True,False
                    End If
                Next 'l
            ElseIf Row1arr(k,0) <> Tableindex And Row2arr(k,0) = Tableindex Then
                For l = 0 To SelCount - 1
                    bh = SSProcess.GetSelGeoValue(l,"[BH]")
                    fxbc = SSProcess.GetSelGeoValue(l,"[FXBC]")
                    yxbc = SSProcess.GetSelGeoValue(l,"[YXBC]")
                    bcjc = SSProcess.GetSelGeoValue(l,"[BCJC]")
                    bcjc = transform(bcjc)
                    tj = bcjc * 1000
                    temp = l Mod 2
                    hs = Int(l / 2) + 20 + Row2arr(k,1) + RowBS
                    If temp = 0 Then
                        g_docObj.SetCellText Tableindex,hs,0,bh,True,False
                        g_docObj.SetCellText Tableindex,hs,1,fxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,2,yxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,3,bcjc,True,False
                        If tj > 50 Then
                            content = g_docObj.GetCellText(Tableindex,hs,8,False)
                            contentlen = Len(content)
                            content = Left(content,contentlen - 1)
                            If content = "" Or content = "����"  Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            End If
                        End If
                    Else
                        g_docObj.SetCellText Tableindex,hs,4,bh,True,False
                        g_docObj.SetCellText Tableindex,hs,5,fxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,6,yxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,7,bcjc,True,False
                        If tj > 50 Then
                            content = g_docObj.GetCellText(Tableindex,hs,8,False)
                            contentlen = Len(content)
                            content = Left(content,contentlen - 1)
                            If content = "" Or content = "����" Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            End If
                        End If
                    End If
                    tj1 = g_docObj.GetCellText(Tableindex,hs,8,False)
                    tjlen1 = Len(tj1)
                    tj1 = Left(tj1,tjlen1 - 1)
                    If tj1 = "" Then
                        g_docObj.SetCellText Tableindex,hs,8,"����",True,False
                    End If
                Next 'l
            Else
                For l = 0 To SelCount - 1
                    bh = SSProcess.GetSelGeoValue(l,"[BH]")
                    fxbc = SSProcess.GetSelGeoValue(l,"[FXBC]")
                    yxbc = SSProcess.GetSelGeoValue(l,"[YXBC]")
                    bcjc = SSProcess.GetSelGeoValue(l,"[BCJC]")
                    bcjc = transform(bcjc)
                    tj = bcjc * 1000
                    temp = l Mod 2
                    hs = Int(l / 2) + 20 + RowBS
                    'MsgBox hs
                    If temp = 0 Then
                        g_docObj.SetCellText Tableindex,hs,0,bh,True,False
                        g_docObj.SetCellText Tableindex,hs,1,fxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,2,yxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,3,bcjc,True,False
                        If tj > 50 Then
                            content = g_docObj.GetCellText(Tableindex,hs,8,False)
                            contentlen = Len(content)
                            content = Left(content,contentlen - 1)
                            'MsgBox content
                            If content = "" Or content = "����" Then
                                'MsgBox bh
                                g_docObj.SetCellText Tableindex,hs,8,bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            End If
                        End If
                    Else
                        g_docObj.SetCellText Tableindex,hs,4,bh,True,False
                        g_docObj.SetCellText Tableindex,hs,5,fxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,6,yxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,7,bcjc,True,False
                        If tj > 50 Then
                            content = g_docObj.GetCellText(Tableindex,hs,8,False)
                            contentlen = Len(content)
                            content = Left(content,contentlen - 1)
                            'MsgBox content
                            If content = "" Or content = "����" Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            End If
                        End If
                    End If
                    tj1 = g_docObj.GetCellText(Tableindex,hs,8,False)
                    tjlen1 = Len(tj1)
                    tj1 = Left(tj1,tjlen1 - 1)
                    'MsgBox tj
                    If tj1 = "" Then
                        g_docObj.SetCellText Tableindex,hs,8,"����",True,False
                    End If
                Next 'l
            End If
            ' Next 'k  
        End If
        
        If SelCount > 4 Then
            k = i - 4
            'For k = 0 To Tablecount - 1
            If Row1arr(k,0) = Tableindex And Row2arr(k,0) = Tableindex  Then
                g_docObj.CloneTableRow Tableindex, 20 + RowBS + Row1arr(k,1) + Row2arr(k,1), 1, Round((SelCount - 4) / 2), False '��������
                For l = 0 To SelCount - 1
                    bh = SSProcess.GetSelGeoValue(l,"[BH]")
                    fxbc = SSProcess.GetSelGeoValue(l,"[FXBC]")
                    yxbc = SSProcess.GetSelGeoValue(l,"[YXBC]")
                    bcjc = SSProcess.GetSelGeoValue(l,"[BCJC]")
                    bcjc = transform(bcjc)
                    tj = bcjc * 1000
                    temp = l Mod 2
                    hs = Int(l / 2) + 20 + Row1arr(k,1) + Row2arr(k,1) + RowBS
                    If temp = 0 Then
                        g_docObj.SetCellText Tableindex,hs,0,bh,True,False
                        g_docObj.SetCellText Tableindex,hs,1,fxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,2,yxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,3,bcjc,True,False
                        If tj > 50 Then
                            content = g_docObj.GetCellText(Tableindex,hs,8,False)
                            contentlen = Len(content)
                            content = Left(content,contentlen - 1)
                            If content = "" Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "������",True,False
                                If DisLine(count) = "" Or content = "����" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            End If
                        End If
                    Else
                        g_docObj.SetCellText Tableindex,hs,4,bh,True,False
                        g_docObj.SetCellText Tableindex,hs,5,fxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,6,yxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,7,bcjc,True,False
                        If tj > 50 Then
                            content = g_docObj.GetCellText(Tableindex,hs,8,False)
                            contentlen = Len(content)
                            content = Left(content,contentlen - 1)
                            If content = "" Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "������",True,False
                                If DisLine(count) = "" Or content = "����"  Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            End If
                        End If
                    End If
                    
                    tj1 = g_docObj.GetCellText(Tableindex,hs,8,False)
                    tjlen1 = Len(tj1)
                    tj1 = Left(tj1,tjlen1 - 1)
                    If tj1 = "" Then
                        g_docObj.SetCellText Tableindex,hs,8,"����",True,False
                    End If
                Next 'l
            ElseIf Row1arr(k,0) = Tableindex And Row2arr(k,0) <> Tableindex Then
                g_docObj.CloneTableRow Tableindex, 20 + RowBS + Row1arr(k,1) , 1, Round((SelCount - 4) / 2), False '��������
                For l = 0 To SelCount - 1
                    bh = SSProcess.GetSelGeoValue(l,"[BH]")
                    fxbc = SSProcess.GetSelGeoValue(l,"[FXBC]")
                    yxbc = SSProcess.GetSelGeoValue(l,"[YXBC]")
                    bcjc = SSProcess.GetSelGeoValue(l,"[BCJC]")
                    bcjc = transform(bcjc)
                    tj = bcjc * 1000
                    temp = l Mod 2
                    hs = Int(l / 2) + 20 + Row1arr(k,1) + RowBS
                    If temp = 0 Then
                        g_docObj.SetCellText Tableindex,hs,0,bh,True,False
                        g_docObj.SetCellText Tableindex,hs,1,fxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,2,yxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,3,bcjc,True,False
                        If tj > 50 Then
                            content = g_docObj.GetCellText(Tableindex,hs,8,False)
                            contentlen = Len(content)
                            content = Left(content,contentlen - 1)
                            If content = "" Or content = "����" Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            End If
                        End If
                    Else
                        g_docObj.SetCellText Tableindex,hs,4,bh,True,False
                        g_docObj.SetCellText Tableindex,hs,5,fxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,6,yxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,7,bcjc,True,False
                        If tj > 50 Then
                            content = g_docObj.GetCellText(Tableindex,hs,8,False)
                            contentlen = Len(content)
                            content = Left(content,contentlen - 1)
                            If content = "" Or content = "����"  Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            End If
                        End If
                    End If
                    
                    tj1 = g_docObj.GetCellText(Tableindex,hs,8,False)
                    tjlen1 = Len(tj1)
                    tj1 = Left(tj1,tjlen1 - 1)
                    If tj1 = "" Then
                        g_docObj.SetCellText Tableindex,hs,8,"����",True,False
                    End If
                Next 'l
            ElseIf Row1arr(k,0) <> Tableindex And Row2arr(k,0) = Tableindex Then
                g_docObj.CloneTableRow Tableindex, 20 + RowBS + Row2arr(k,1) , 1, Round((SelCount - 4) / 2), False '��������
                For l = 0 To SelCount - 1
                    bh = SSProcess.GetSelGeoValue(l,"[BH]")
                    fxbc = SSProcess.GetSelGeoValue(l,"[FXBC]")
                    yxbc = SSProcess.GetSelGeoValue(l,"[YXBC]")
                    bcjc = SSProcess.GetSelGeoValue(l,"[BCJC]")
                    bcjc = transform(bcjc)
                    tj = bcjc * 1000
                    temp = l Mod 2
                    hs = Int(l / 2) + 20 + Row2arr(k,1) + RowBS
                    If temp = 0 Then
                        g_docObj.SetCellText Tableindex,hs,0,bh,True,False
                        g_docObj.SetCellText Tableindex,hs,1,fxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,2,yxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,3,bcjc,True,False
                        If tj > 50 Then
                            content = g_docObj.GetCellText(Tableindex,hs,8,False)
                            contentlen = Len(content)
                            content = Left(content,contentlen - 1)
                            If content = "" Or content = "����"  Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            End If
                        End If
                    Else
                        g_docObj.SetCellText Tableindex,hs,4,bh,True,False
                        g_docObj.SetCellText Tableindex,hs,5,fxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,6,yxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,7,bcjc,True,False
                        If tj > 50 Then
                            content = g_docObj.GetCellText(Tableindex,hs,8,False)
                            contentlen = Len(content)
                            content = Left(content,contentlen - 1)
                            If content = "" Or content = "����"  Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            End If
                        End If
                    End If
                    
                    tj1 = g_docObj.GetCellText(Tableindex,hs,8,False)
                    tjlen1 = Len(tj1)
                    tj1 = Left(tj1,tjlen1 - 1)
                    If tj = "" Then
                        g_docObj.SetCellText Tableindex,hs,8,"����",True,False
                    End If
                Next 'l
            Else
                g_docObj.CloneTableRow Tableindex, 20 + RowBS , 1, Round((SelCount - 4) / 2), False '��������
                For l = 0 To SelCount - 1
                    bh = SSProcess.GetSelGeoValue(l,"[BH]")
                    fxbc = SSProcess.GetSelGeoValue(l,"[FXBC]")
                    yxbc = SSProcess.GetSelGeoValue(l,"[YXBC]")
                    bcjc = SSProcess.GetSelGeoValue(l,"[BCJC]")
                    bcjc = transform(bcjc)
                    tj = bcjc * 1000
                    temp = l Mod 2
                    hs = Int(l / 2) + 20 + RowBS
                    If temp = 0 Then
                        g_docObj.SetCellText Tableindex,hs,0,bh,True,False
                        g_docObj.SetCellText Tableindex,hs,1,fxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,2,yxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,3,bcjc,True,False
                        If tj > 50 Then
                            content = g_docObj.GetCellText(Tableindex,hs,8,False)
                            contentlen = Len(content)
                            content = Left(content,contentlen - 1)
                            If content = "" Or content = "����"  Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            End If
                        End If
                    Else
                        g_docObj.SetCellText Tableindex,hs,4,bh,True,False
                        g_docObj.SetCellText Tableindex,hs,5,fxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,6,yxbc,True,False
                        g_docObj.SetCellText Tableindex,hs,7,bcjc,True,False
                        If tj > 50 Then
                            content = g_docObj.GetCellText(Tableindex,hs,8,False)
                            contentlen = Len(content)
                            content = Left(content,contentlen - 1)
                            If content = "" Or content = "����" Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "������",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            End If
                        End If
                    End If
                    tj1 = g_docObj.GetCellText(Tableindex,hs,8,False)
                    tjlen1 = Len(tj1)
                    tj1 = Left(tj1,tjlen1 - 1)
                    If tj1 = "" Then
                        g_docObj.SetCellText Tableindex,hs,8,"����",True,False
                    End If
                Next 'l
            End If
            'Next 'k
        End If
        count = count + 1
    Next 'i
End Function' Set4Line

'�˲������
Function SetInfoTable()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9130224"
    SSProcess.SelectFilter
    hxcount = SSProcess.GetSelgeoCount
    If hxcount = 1 Then
        id = SSProcess.GetSelGeoValue(0,"SSObj_ID ")
        ids1 = SSProcess.SearchInnerObjIDs(id,0,"9130511", 0)
        idarr1 = Split(ids1,",", - 1,1)
        ScFyCount = UBound(idarr1) + 1
        
        ids2 = SSProcess.SearchInnerObjIDs(id,0,"9130611", 0)
        idarr2 = Split(ids2,",", - 1,1)
        ZFCount = UBound(idarr2) + 1
        
        ids3 = SSProcess.SearchInnerObjIDs(id,1,"9310053", 0)
        idarr3 = Split(ids3,",", - 1,1)
        YXCount = UBound(idarr3) + 1
        
        g_docObj.Replace "{" & "GETPO" & "}",ScFyCount,0
        g_docObj.Replace "{" & "GETGC" & "}",ZFCount,0
        g_docObj.Replace "{" & "GETBC" & "}",YXCount,0
    End If
End Function' SetInfoTable

'�滮���߲�������
Function SetResultTable()
    poiname = ""
    For i = 0 To KzPoiCount - 1
        If DifKzPoi(i) <> "" Then
            If poiname = "" Then
                poiname = DifKzPoi(i)
            Else
                poiname = poiname & "," & DifKzPoi(i)
            End If
        End If
    Next 'i
    'MsgBox poiname
    TotalStr = ""
    
    If poiname <> "" Then
        str = "��ʵ�⣬��������" & poiname & "���Ƶ㲻���㾫��Ҫ��"
        If TotalStr = "" Then
            TotalStr = str
        Else
            TotalStr = str & Chr(13) & TotalStr
        End If
    End If
    
    For i = 0 To Tablecount - 1
        TitleName = g_docObj.GetCellText(i + 4,0,0,False)
        'MsgBox TitleName
        Title = Replace(TitleName,"���蹤�̹滮���߳ɹ���","")
        totallen = Len(Title)
        Title = Left(Title,totallen - 1)
        If DisPoi(i) <> "" And DisLine(i) <> ""  Then
            str = "��ʵ�⣬��������" & Title & "�����" & DisPoi(i) & "�����޲Χ����������" & DisLine(i) & "�����޲Χ��"
            If TotalStr = "" Then
                TotalStr = str
            Else
                TotalStr = str & Chr(13) & TotalStr
            End If
        ElseIf DisPoi(i) = "" And DisLine(i) <> "" Then
            str = "��ʵ�⣬��������" & Title & "����������" & DisLine(i) & "�����޲Χ��"
            If TotalStr = "" Then
                TotalStr = str
            Else
                TotalStr = str & Chr(13) & TotalStr
            End If
        ElseIf DisPoi(i) <> "" And DisLine(i) = "" Then
            str = "��ʵ�⣬��������" & Title & "�����" & DisPoi(i) & "�����޲Χ��"
            If TotalStr = "" Then
                TotalStr = str
            Else
                TotalStr = str & Chr(13) & TotalStr
            End If
        End If
        
        If DifZFL(i) <> "" Then
            TotalJzwnameArr = Split(DifZFL(i),",", - 1,1)
            str = "��ʵ�⣬��������" & TotalJzwnameArr(0) & "�������߲����㾫��Ҫ��"
            If TotalStr = "" Then
                TotalStr = str
            Else
                TotalStr = str & Chr(13) & TotalStr
            End If
        End If
    Next 'i
    
    If TotalStr = "" Then
        TotalStr = "�����Ϻ˲���ֳ����߲������������ؽ��蹤�̷��߲������桷�������ݡ�ע�����ݼ����������ʽ���Ϲ涨�����߲���������Ƶ㡢�����㣨���㣩���㾫��Ҫ�󣬡������ؽ��蹤�̷��߲������桷�е��������ꡢ�߳���������ϵ��滮���һ�£����߷��Ϲ滮Ҫ��"
    End If
    g_docObj.Replace "{" & "TEXT" & "}",TotalStr,0
End Function' SetResultTable

'������Ƭ
Function InsertPhoto()
    Dim f1,fc,f
    filePath = SSProcess.GetSysPathName(5)
    filePath = filePath & "4Ӱ��"
    'MsgBox filePath
    Set f = fso.GetFolder(filePath)
    Set fc = f.Files
    s = ""
    For Each f1 In fc
        If s = "" Then
            s = f1.name
        Else
            s = s & "," & f1.name
        End If
    Next
    sarr = Split(s,",", - 1,1)
    gs = UBound(sarr)
    count = Tablecount + 4
    'MsgBox TypeName(count)
    For i = 0 To gs
        If gs <= 3 Then
            row = Int(i / 2) + 1
            col = i Mod 2
            'MsgBox filePath & "\" & sarr(i)
            If col = 0 Then
                'g_docObj.SetCellText Tableindex,hs,8,"����",True,False
                g_docObj.SetCellImageEx count, row, 0, - 1, filePath & "\" & sarr(i) ,1,1
            Else
                g_docObj.SetCellImageEx count, row + 1, 1, - 1, filePath & "\" & sarr(i),1,1
            End If
        End If
    Next 'i
End Function' InsertPhoto

'��������ת��
Function transform(content)
    If content <> "" Then
        content = CDbl(content)
    Else
        MsgBox "��������"
        Exit Function
    End If
    transform = content
End Function

'���Ƹ����ߣ�ʵ�⣩
Function MakeLine1(x1,y1,x2,y2,jzwname)
    SSProcess.CreateNewObj 1
    SSProcess.SetNewObjValue "SSObj_Code", "1"
    SSProcess.SetNewObjValue "[Note]", jzwname
    'MsgBox x1 & "," & y1 & ";" & x2 & "," & y2
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjPoint x2, y2, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

'���Ƹ����ߣ����ۣ�
Function MakeLine2(x1,y1,x2,y2,jzwname)
    SSProcess.CreateNewObj 1
    SSProcess.SetNewObjValue "SSObj_Code", "2"
    SSProcess.SetNewObjValue "[Note]", jzwname
    'MsgBox x1 & "," & y1 & ";" & x2 & "," & y2
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjPoint x2, y2, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

'ɾ����
Function DelLine()
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "1,2"
    SSProcess.SelectFilter
    SSProcess.DeleteSelectionObj
End Function' DelLine

'ѡ������
Function SelLine1(coed)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", coed
    SSProcess.SelectFilter
End Function' SelLine

'ѡ������
Function SelLine(coed,note)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", coed
    SSProcess.SetSelectCondition "[Note]", "==", note
    SSProcess.SelectFilter
End Function' SelLine

'ѡ��ƫ���
Function SelDiffLine(coed,name)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", coed
    SSProcess.SetSelectCondition "[JianZWMC]", "==", name
    SSProcess.SelectFilter
End Function' SelDiffLine

'ѡ�����߱߳�
Function SelYxBc(name)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9310053"
    SSProcess.SetSelectCondition "[JianZWMC]", "==", name
    SSProcess.SelectFilter
End Function' SelYxBc

'��������
Function GetLengthDiff(x1,y1,x2,y2)
    diff = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    diff = Round(diff,3)
    GetLengthDiff = diff
End Function' GetLengthDiff

'��ͼ��
Function allvisible()
    count = SSProcess.GetLayerCount
    For i = 0 To count - 1
        layername = SSProcess.GetLayerName (i)
        SSProcess.SetLayerStatus layername, 1, 1
    Next
    SSProcess.RefreshView
End Function