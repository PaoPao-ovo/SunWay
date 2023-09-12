'doc全局对象
Dim g_docObj

'路径操作对象
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

'存储红线属性
Dim HXarr(1000,2)

'起算点插入行数
Dim RowBS
RowBS = 0

'表2各表插入行数
Dim Row1arr(1000,2)

'表3各表插入的行数
Dim Row2arr(1000,2)

'表数
Dim Tablecount
Tablecount = 0

'不符合点号
Dim DisPoi(1000)

'不符合线号
Dim DisLine(1000)

'不符合控制点标识
Dim DifKzPoi(1000)

'不符合正负零标高表号
Dim DifZFL(1000)

'满足的控制点个数
Dim KzPoiCount
KzPoiCount = 0

'入口函数
Sub OnClick()
    allvisible()
    strTempFileName = "验线测量报告模板.doc"
    strTempFilePath = SSProcess.GetSysPathName (7) & "输出模板\" & strTempFileName
    Set g_docObj = CreateObject ("asposewordscom.asposewordshelper")
    If  TypeName (g_docObj) = "AsposeWordsHelper" Then
        g_docObj.CreateDocumentByTemplate strTempFilePath
    Else
        MsgBox "请先注册Aspose.Word插件"
        Exit Sub
    End If
    
    pathName = GetFilePath()
    
    proname = GetFileName()
    
    g_docObj.CreateDocumentByTemplate  strTempFilePath

    ReplaceValue()

    SetKZD()
    CopyTable()

    'SetPosition()
    ZFL()
    Set4Line()

    SetInfoTable()
    SetResultTable()
    'InsertPhoto()
    
    strFileSavePath = pathName & proname
    'MsgBox strFileSavePath
    g_docObj.SaveEx  strFileSavePath
		msgbox "输出成功"
    
End Sub

'//获取成果目录路径
Function  GetFilePath()
    filePath = SSProcess.GetSysPathName(5)
    filePath = filePath & "3成果" & "\"
    ' filePath = filePath & "\"
    GetFilePath = filePath
End Function

'//获取成果报告名称
Function  GetFileName()
    proname = GetProName()
    GetFileName = proname & "验线测量报告.doc"
End Function


'获取当前红线的项目名称
Function GetProName()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9130224"
    SSProcess.SelectFilter
    hxcount = SSProcess.GetSelgeoCount
    If hxcount = 1 Then xmmc = SSProcess.GetSelGeoValue (0,"[XiangMMC]")
    GetProName = xmmc
End Function' GetProName

'字符替换 
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
    
	currentDate = Date()
	targetDate = DateSerial(Year(currentDate), 11, 10)

	If currentDate < targetDate Then
		 g_docObj.Replace "{有效年}",Year(Now),0
	Else
		 g_docObj.Replace "{有效年}",Year(Now)+1,0
	End If
End Function

'复制表
Function CopyTable()
    zzs = HXarr(11,0) '范围线内的建筑物总个数
	 if zzs = "" then exit function
    zzs = transform(zzs)
    'MsgBox zzs
    i = 1
    bulidname = GetBuildingName()
    'MsgBox bulidname
    bulidarr = Split(bulidname,",", - 1,1)
    text = bulidarr(0) & "建设工程规划验线成果表"
    g_docObj.SetCellText 3,0,0,text,True,False
    While i <= zzs - 1
        g_docObj.CloneTable 3,0,0,False 
        text = bulidarr(i) & "建设工程规划验线成果表"
        'MsgBox text
        g_docObj.SetCellText 3,0,0,text,True,False
        i = i + 1
    WEnd
    Tablecount = zzs
    Tablecount = CInt(Tablecount)
End Function' CopyTable

'获取当前的建筑物名称
Function GetBuildingName()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9310013" '实测改为建筑物轴线
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

'遍历控制点并填值
Function SetKZD()
    Dim LLarr(1000,4)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9130211,9130212" '理论控制点
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

        x = FormatNumber(transform(x),3,,,0)  '保留3位小数，前面不加“，”
        y = FormatNumber(transform(y),3,,,0)
        z = FormatNumber(transform(z),3,,,0)

        
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
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9130311,9130312,1102021,1103021"
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
        
        x = FormatNumber(transform(x),3,,,0)
        y = FormatNumber(transform(y),3,,,0)
        z = FormatNumber(transform(z),3,,,0)
        
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

                    Diffh = FormatNumber(Diffh,0)
                    g_docObj.SetCellText 3,row + k,0,LLarr(j,1),True,False
                    g_docObj.SetCellText 3,row + k,1,LLarr(j,0),True,False
                    g_docObj.SetCellText 3,row + k,2,LLarr(j,2),True,False
                    g_docObj.SetCellText 3,row + k,3,SCarr(i,1),True,False
                    g_docObj.SetCellText 3,row + k,4,SCarr(i,0),True,False
                    g_docObj.SetCellText 3,row + k,5,SCarr(i,2),True,False
                    g_docObj.SetCellText 3,row + k,6,Diffxy,True,False
                    g_docObj.SetCellText 3,row + k,7,Diffh,True,False
                    g_docObj.SetCellText 3,row + k,8,50,True,False
                    g_docObj.SetCellText 3,row + k,9,30,True,False
                    'MsgBox count
                    If Diffxy > 50 Then
                        g_docObj.SetCellText 3,row + k,10,"不符合",True,False
                        DifKzPoi(count) = SCarr(i,3)
                    Else
                        g_docObj.SetCellText 3,row + k,10,"符合",True,False
                    End If
                    If Diffh > 30 Then
                        g_docObj.SetCellText 3,row + k,11,"不符合",True,False
                        DifKzPoi(count) = SCarr(i,3)
                    Else
                        g_docObj.SetCellText 3,row + k,11,"符合",True,False
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
                    Diffh = FormatNumber(Diffh,0)
                    g_docObj.SetCellText 3,row + k,0,LLarr(j,1),True,False
                    g_docObj.SetCellText 3,row + k,1,LLarr(j,0),True,False
                    g_docObj.SetCellText 3,row + k,2,LLarr(j,2),True,False
                    g_docObj.SetCellText 3,row + k,3,SCarr(i,1),True,False
                    g_docObj.SetCellText 3,row + k,4,SCarr(i,0),True,False
                    g_docObj.SetCellText 3,row + k,5,SCarr(i,2),True,False
                    g_docObj.SetCellText 3,row + k,6,Diffxy,True,False
                    g_docObj.SetCellText 3,row + k,7,Diffh,True,False
                    g_docObj.SetCellText 3,row + k,8,50,True,False
                    g_docObj.SetCellText 3,row + k,9,30,True,False
                    If Diffxy > 50 Then
                        g_docObj.SetCellText 3,row + k,10,"不符合",True,False
                        DifKzPoi(count) = SCarr(i,3)
                    Else
                        g_docObj.SetCellText 3,row + k,10,"符合",True,False
                    End If
                    If Diffh > 30 Then
                        g_docObj.SetCellText 3,row + k,11,"不符合",True,False
                        DifKzPoi(count) = SCarr(i,3)
                    Else
                        g_docObj.SetCellText 3,row + k,11,"符合",True,False
                    End If
                    k = k + 1
                    count = count + 1
                End If
            Next
        Next
    End If
  
End Function' SetKZD

'设置坐标验线较差
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

    namearr = Split(Jzwname,",", - 1,1)
    count = 0
    For i = 0 To UBound(namearr)
        SelDiffLine "9316710",namearr(i)
        SelCount = SSProcess.GetSelgeoCount()
        ReDim Pcarr(SelCount,6)
        For k = 0 To SelCount - 1
            llx = SSProcess.GetSelGeoValue(k,"[llzbx]")
            lly = SSProcess.GetSelGeoValue(k,"[llzby]")
            scx = SSProcess.GetSelGeoValue(k,"[sczbx]")
            scy = SSProcess.GetSelGeoValue(k,"[sczby]")
            pc = SSProcess.GetSelGeoValue(k,"[pcjl]")
            dh = SSProcess.GetSelGeoValue(k,"[dh]")
'msgbox llx
				llx = formatnumber(llx,3,,,0)
				lly = formatnumber(lly,3,,,0)
				scx = formatnumber(scx,3,,,0)
				scy = formatnumber(scy,3,,,0)
				pc = formatnumber(pc, 3,,,0)
'msgbox llx
				if pc >0 and  pc <1 then pc = formatnumber(pc, 3, -1) '修改显示小数点前的0
            Pcarr(k,0) = lly
            Pcarr(k,1) = llx
            Pcarr(k,2) = scy
            Pcarr(k,3) = scx
            Pcarr(k,4) = pc
            Pcarr(k,5) = dh
        Next 'k
        
        For j = 3 To Tablecount + 3
            TitleName = g_docObj.GetCellText(j,0,0,False)
            'MsgBox TitleName
            Title = Replace(TitleName,"建设工程规划验线成果表","")
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
                    g_docObj.SetCellText Tableindex,10 + RowBS + m,6,"不符合",True,False
                    If DisPoi(count) = "" Then
                        DisPoi(count) = Pcarr(m,5)
                    Else
                        DisPoi(count) = DisPoi(count) & "," & Pcarr(m,5)
                    End If
                Else
                    g_docObj.SetCellText Tableindex,10 + RowBS + m,6,"符合",True,False
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
                    g_docObj.SetCellText Tableindex,10 + RowBS + m,6,"不符合",True,False
                    If DisPoi(count) = "" Then
                        DisPoi(count) = Pcarr(m,5)
                    Else
                        DisPoi(count) = DisPoi(count) & "," & Pcarr(m,5)
                    End If
                Else
                    g_docObj.SetCellText Tableindex,10 + RowBS + m,6,"符合",True,False
                End If
            Next
            
        End If
        count = count + 1
'msgbox count
    Next 'i
End Function' SetPosition

'搜索正负零标高并填值
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
			SSProcess.GetObjectFocusPoint geoid, foc_x, foc_y  '获取放样建筑的焦点坐标
        jzwname = SSProcess.GetSelGeoValue(i,"[JianZWMC]")
        On Error Resume Next
				ids = SSProcess.SearchInnerObjIDs(id,0,"9130611",0)
			On Error GoTo 0
        'MsgBox ids
        idsarr = Split(ids,",", - 1,1)
        ZFLcount = UBound(idsarr)
        
        'MsgBox ZFLcount
        '获取表索引号
        For j = 3 To Tablecount + 3
            TitleName = g_docObj.GetCellText(j,0,0,False)
            'MsgBox TitleName
            Title = Replace(TitleName,"建设工程规划验线成果表","")
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
							'MSGBOX  jzwmc
                    sjgc = formatnumber(sjgc,3)
                    yxgc = formatnumber(yxgc,3)
                    Diffh = Abs(sjgc - yxgc)
                    Diffh = FormatNumber(Diffh,3,-1)
                    For m = 0 To Tablecount - 1
                        If Row1arr(m,0) = Tableindex Then
                            g_docObj.SetCellText Tableindex,16 + RowBS + k + Row1arr(m,1),0,sjgc,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + k + Row1arr(m,1),1,yxgc,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + k + Row1arr(m,1),2,Diffh,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + k + Row1arr(m,1),3,30,True,False
                            If Diffh * 1000 > 30 Then
                                g_docObj.SetCellText Tableindex,16 + RowBS + k + Row1arr(m,1),4,"不符合",True,False
                                If DifZFL(count) = "" Then
                                    DifZFL(count) = jzwname
                                Else
                                    DifZFL(count) = DifZFL(count) & "," & jzwname
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,16 + RowBS + k + Row1arr(m,1),4,"符合",True,False
                            End If
                        Else
                            g_docObj.SetCellText Tableindex,16 + RowBS + k,0,sjgc,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + k,1,yxgc,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + k,2,Diffh,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + k,3,30,True,False
                            If Diffh * 1000 > 30 Then
                                g_docObj.SetCellText Tableindex,16 + RowBS + k,4,"不符合",True,False
                                If DifZFL(count) = "" Then
                                    DifZFL(count) = jzwname
                                Else
                                    DifZFL(count) = DifZFL(count) & "," & jzwname
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,16 + RowBS + k,4,"符合",True,False
                            End If
                        End If
                    Next
                End If
                
                If ZFLcount >= 2 Then
                    g_docObj.CloneTableRow Tableindex, 16 + RowBS + k + Row1arr(k,1), 1,ZFLcount - 1 , False '插入列数
                    sjgc = SSProcess.GetObjectAttr(idsarr(k),"[SheJGC]")
                    yxgc = SSProcess.GetObjectAttr(idsarr(k),"[YanXGC]")
                    sjgc = formatnumber(sjgc,3)
                    yxgc = formatnumber(yxgc,3)
                    Diffh = Abs(sjgc - yxgc)
                    Diffh = FormatNumber(Diffh,3,-1)
                    For m = 0 To Tablecount - 1
                        If Row1arr(m,0) = Tableindex Then
                            g_docObj.SetCellText Tableindex,16 + RowBS + k + Row1arr(m,1),0,sjgc,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + k + Row1arr(m,1),1,yxgc,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + k + Row1arr(m,1),2,Diffh,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + k + Row1arr(m,1),3,30,True,False
                            If Diffh * 1000 > 30 Then
                                g_docObj.SetCellText Tableindex,16 + RowBS + k + Row1arr(m,1),4,"不符合",True,False
                                If DifZFL(count) = "" Then
                                    DifZFL(count) = jzwname
                                Else
                                    DifZFL(count) = DifZFL(count) & "," & jzwname
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,16 + RowBS + k + Row1arr(m,1),4,"符合",True,False
                            End If
                            Row2arr(k,0) = Tableindex
                            Row2arr(k,1) = ZFLcount - 1 + Row1arr(k,1)
                        Else
                            g_docObj.SetCellText Tableindex,16 + RowBS + k,0,sjgc,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + k,1,yxgc,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + k,2,Diffh,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + k,3,30,True,False
                            If Diffh * 1000 > 30 Then
                                g_docObj.SetCellText Tableindex,16 + RowBS + k,4,"不符合",True,False
                                If DifZFL(count) = "" Then
                                    DifZFL(count) = jzwname
                                Else
                                    DifZFL(count) = DifZFL(count) & "," & jzwname
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,16 + RowBS + k,4,"符合",True,False
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
						dim min_distance,min_id
						min_distance =9999
						for zfl_id=0 to ZFLcount

								zfl_x = SSProcess.GetObjectAttr (zfl_id, "SSObj_X")
								zfl_y = SSProcess.GetObjectAttr (zfl_id, "SSObj_Y")

							 distance = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
							 if distance < min_distance then
								min_distance = distance
								min_id = ZFLcount
							end if
						next

                    sjgc = SSProcess.GetObjectAttr(idsarr(min_id),"[SheJGC]")
                    yxgc = SSProcess.GetObjectAttr(idsarr(min_id),"[YanXGC]")
                    sjgc = formatnumber(sjgc,3)
                    yxgc = formatnumber(yxgc,3)
                    Diffh = Abs(sjgc - yxgc)
                    Diffh = FormatNumber(Diffh,3,-1)
                    For n = 0 To Tablecount - 1
                        If Row1arr(n,0) = Tableindex Then
                            'MsgBox Row1arr
                            g_docObj.SetCellText Tableindex,16 + RowBS + Row1arr(n,1),0,sjgc,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + Row1arr(n,1),1,yxgc,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + Row1arr(n,1),2,Diffh,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS + Row1arr(n,1),3,30,True,False
                            If Diffh * 1000 > 30 Then
                                g_docObj.SetCellText Tableindex,16 + RowBS + Row1arr(n,1),4,"不符合",True,False
                                If DifZFL(count) = "" Then
                                    DifZFL(count) = Tableindex
                                Else
                                    DifZFL(count) = DifZFL(count) & "," & Tableindex
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,16 + RowBS + Row1arr(n,1),4,"符合",True,False
                            End If
                            
                        Else
                            g_docObj.SetCellText Tableindex,16 + RowBS ,0,sjgc,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS ,1,yxgc,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS ,2,Diffh,True,False
                            g_docObj.SetCellText Tableindex,16 + RowBS ,3,30,True,False
                            If Diffh * 1000 > 30 Then
                                g_docObj.SetCellText Tableindex,16 + RowBS ,4,"不符合",True,False
                                If DifZFL(count) = "" Then
                                    DifZFL(count) = Tableindex
                                Else
                                    DifZFL(count) = DifZFL(count) & "," & Tableindex
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,16 + RowBS ,4,"符合",True,False
                            End If
                        End If
                    Next 'n
                End If
            Next
        End If
        count = count + 1
    Next
End Function' ZFL


'设置四至边长
Function Set4Line()
    count = 0

    For i = 3 To Tablecount + 3

        TitleName = g_docObj.GetCellText(i,0,0,False)
        Title = Replace(TitleName,"建设工程规划验线成果表","")
        totallen = Len(Title)
        Title = Left(Title,totallen - 1)
        
        SelYxBc Title
        SelCount = SSProcess.GetSelgeoCount()
        'MsgBox SelCount
        Tableindex = i
        If SelCount <= 4 Then
            k = i - 3
            If Row1arr(k,0) = Tableindex And Row2arr(k,0) = Tableindex  Then
                For l = 0 To SelCount - 1
                    bh = SSProcess.GetSelGeoValue(l,"[BH]")
                    fxbc = SSProcess.GetSelGeoValue(l,"[FXBC]")
                    yxbc = SSProcess.GetSelGeoValue(l,"[YXBC]")
                    bcjc = SSProcess.GetSelGeoValue(l,"[BCJC]")
                    bcjc = transform(bcjc)
						if bcjc >=0 and  bcjc <1 then bcjc = formatnumber(bcjc,2, -1) '修改显示小数点前的0
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
                            If content = "" Or content = "符合" Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "不符合",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "不符合",True,False
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
                            If content = "" Or content = "符合" Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "不符合",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "不符合",True,False
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
                        g_docObj.SetCellText Tableindex,hs,8,"符合",True,False
                    End If
                Next 'l
            ElseIf Row1arr(k,0) = Tableindex And Row2arr(k,0) <> Tableindex Then
                For l = 0 To SelCount - 1
                    bh = SSProcess.GetSelGeoValue(l,"[BH]")
                    fxbc = SSProcess.GetSelGeoValue(l,"[FXBC]")
                    yxbc = SSProcess.GetSelGeoValue(l,"[YXBC]")
                    bcjc = SSProcess.GetSelGeoValue(l,"[BCJC]")
                    bcjc = transform(bcjc)
						if bcjc >=0 and  bcjc <1 then bcjc = formatnumber(bcjc,2, -1) '修改显示小数点前的0
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
                            If content = "" Or content = "符合" Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "不符合",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "不符合",True,False
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
                            If content = "" Or content = "符合" Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "不符合",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "不符合",True,False
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
                        g_docObj.SetCellText Tableindex,hs,8,"符合",True,False
                    End If
                Next 'l
            ElseIf Row1arr(k,0) <> Tableindex And Row2arr(k,0) = Tableindex Then
                For l = 0 To SelCount - 1
                    bh = SSProcess.GetSelGeoValue(l,"[BH]")
                    fxbc = SSProcess.GetSelGeoValue(l,"[FXBC]")
                    yxbc = SSProcess.GetSelGeoValue(l,"[YXBC]")
                    bcjc = SSProcess.GetSelGeoValue(l,"[BCJC]")
                    bcjc = transform(bcjc)
							if bcjc >=0 and  bcjc <1 then bcjc = formatnumber(bcjc,2, -1) '修改显示小数点前的0
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
                            If content = "" Or content = "符合"  Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "不符合",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "不符合",True,False
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
                            If content = "" Or content = "符合" Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "不符合",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "不符合",True,False
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
                        g_docObj.SetCellText Tableindex,hs,8,"符合",True,False
                    End If
                Next 'l
            Else
                For l = 0 To SelCount - 1
                    bh = SSProcess.GetSelGeoValue(l,"[BH]")
                    fxbc = SSProcess.GetSelGeoValue(l,"[FXBC]")
                    yxbc = SSProcess.GetSelGeoValue(l,"[YXBC]")
                    bcjc = SSProcess.GetSelGeoValue(l,"[BCJC]")
                    bcjc = transform(bcjc)
							if bcjc >=0 and  bcjc <1 then bcjc = formatnumber(bcjc,2, -1) '修改显示小数点前的0
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
                            If content = "" Or content = "符合" Then
                                'MsgBox bh
                                g_docObj.SetCellText Tableindex,hs,8,bh & "不符合",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "不符合",True,False
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
                            If content = "" Or content = "符合" Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "不符合",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "不符合",True,False
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
                        g_docObj.SetCellText Tableindex,hs,8,"符合",True,False
                    End If
                Next 'l
            End If
            ' Next 'k  
        End If
        
        If SelCount > 4 Then
            k = i - 4 '原为-4
            If Row1arr(k,0) = Tableindex And Row2arr(k,0) = Tableindex  Then
                g_docObj.CloneTableRow Tableindex, 20 + RowBS + Row1arr(k,1) + Row2arr(k,1), 1, Round((SelCount - 4) / 2), False '插入行数
                For l = 0 To SelCount - 1
                    bh = SSProcess.GetSelGeoValue(l,"[BH]")
                    fxbc = SSProcess.GetSelGeoValue(l,"[FXBC]")
                    yxbc = SSProcess.GetSelGeoValue(l,"[YXBC]")
                    bcjc = SSProcess.GetSelGeoValue(l,"[BCJC]")
                    bcjc = transform(bcjc)
							if bcjc >=0 and  bcjc <1 then bcjc = formatnumber(bcjc,2, -1) '修改显示小数点前的0
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
                                g_docObj.SetCellText Tableindex,hs,8,bh & "不符合",True,False
                                If DisLine(count) = "" Or content = "符合" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "不符合",True,False
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
                                g_docObj.SetCellText Tableindex,hs,8,bh & "不符合",True,False
                                If DisLine(count) = "" Or content = "符合"  Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "不符合",True,False
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
                        g_docObj.SetCellText Tableindex,hs,8,"符合",True,False
                    End If
                Next 'l
            ElseIf Row1arr(k,0) = Tableindex And Row2arr(k,0) <> Tableindex Then
                g_docObj.CloneTableRow Tableindex, 20 + RowBS + Row1arr(k,1) , 1, Round((SelCount - 4) / 2), False '插入行数
                For l = 0 To SelCount - 1
                    bh = SSProcess.GetSelGeoValue(l,"[BH]")
                    fxbc = SSProcess.GetSelGeoValue(l,"[FXBC]")
                    yxbc = SSProcess.GetSelGeoValue(l,"[YXBC]")
                    bcjc = SSProcess.GetSelGeoValue(l,"[BCJC]")
                    bcjc = transform(bcjc)
                   if bcjc >=0 and  bcjc <1 then bcjc = formatnumber(bcjc,2, -1) '修改显示小数点前的0
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
                            If content = "" Or content = "符合" Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "不符合",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "不符合",True,False
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
                            If content = "" Or content = "符合"  Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "不符合",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "不符合",True,False
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
                        g_docObj.SetCellText Tableindex,hs,8,"符合",True,False
                    End If
                Next 'l
            ElseIf Row1arr(k,0) <> Tableindex And Row2arr(k,0) = Tableindex Then
                g_docObj.CloneTableRow Tableindex, 20 + RowBS + Row2arr(k,1) , 1, Round((SelCount - 4) / 2), False '插入行数
                For l = 0 To SelCount - 1
                    bh = SSProcess.GetSelGeoValue(l,"[BH]")
                    fxbc = SSProcess.GetSelGeoValue(l,"[FXBC]")
                    yxbc = SSProcess.GetSelGeoValue(l,"[YXBC]")
                    bcjc = SSProcess.GetSelGeoValue(l,"[BCJC]")
                    bcjc = transform(bcjc)
if bcjc >=0 and  bcjc <1 then bcjc = formatnumber(bcjc,2, -1) '修改显示小数点前的0
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
                            If content = "" Or content = "符合"  Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "不符合",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "不符合",True,False
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
                            If content = "" Or content = "符合"  Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "不符合",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "不符合",True,False
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
                        g_docObj.SetCellText Tableindex,hs,8,"符合",True,False
                    End If
                Next 'l
            Else
                g_docObj.CloneTableRow Tableindex, 20 + RowBS , 1, Round((SelCount - 4) / 2), False '插入行数
                For l = 0 To SelCount - 1
                    bh = SSProcess.GetSelGeoValue(l,"[BH]")
                    fxbc = SSProcess.GetSelGeoValue(l,"[FXBC]")
                    yxbc = SSProcess.GetSelGeoValue(l,"[YXBC]")
                    bcjc = SSProcess.GetSelGeoValue(l,"[BCJC]")
                    bcjc = transform(bcjc)
							if bcjc >=0 and  bcjc <1 then bcjc = formatnumber(bcjc,2, -1) '修改显示小数点前的0
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
                            If content = "" Or content = "符合"  Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "不符合",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "不符合",True,False
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
                            If content = "" Or content = "符合" Then
                                g_docObj.SetCellText Tableindex,hs,8,bh & "不符合",True,False
                                If DisLine(count) = "" Then
                                    DisLine(count) = bh
                                Else
                                    DisLine(count) = DisLine(count) & "," & bh
                                End If
                            Else
                                g_docObj.SetCellText Tableindex,hs,8,content & bh & "不符合",True,False
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
                        g_docObj.SetCellText Tableindex,hs,8,"符合",True,False
                    End If
                Next 'l
            End If
            'Next 'k
        End If
        count = count + 1
    Next 'i
End Function' Set4Line

'核查情况表
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

'规划验线测量结论
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
        str = "经实测，本次验线" & poiname & "控制点不满足精度要求。"
        If TotalStr = "" Then
            TotalStr = str
        Else
            TotalStr = str & Chr(13) & TotalStr
        End If
    End If
    
    For i = 0 To Tablecount - 1
        TitleName = g_docObj.GetCellText(i + 3,0,0,False)
        'MsgBox TitleName
        Title = Replace(TitleName,"建设工程规划验线成果表","")
        totallen = Len(Title)
        Title = Left(Title,totallen - 1)
        If DisPoi(i) <> "" And DisLine(i) <> ""  Then
            str = "经实测，本次验线" & Title & "，点号" & DisPoi(i) & "超出限差范围，四至距离" & DisLine(i) & "超出限差范围。"
            If TotalStr = "" Then
                TotalStr = str
            Else
                TotalStr = str & Chr(13) & TotalStr
            End If
        ElseIf DisPoi(i) = "" And DisLine(i) <> "" Then
            str = "经实测，本次验线" & Title & "，四至距离" & DisLine(i) & "超出限差范围。"
            If TotalStr = "" Then
                TotalStr = str
            Else
                TotalStr = str & Chr(13) & TotalStr
            End If
        ElseIf DisPoi(i) <> "" And DisLine(i) = "" Then
            str = "经实测，本次验线" & Title & "，点号" & DisPoi(i) & "超出限差范围。"
            If TotalStr = "" Then
                TotalStr = str
            Else
                TotalStr = str & Chr(13) & TotalStr
            End If
        End If
        
        If DifZFL(i) <> "" Then
            TotalJzwnameArr = Split(DifZFL(i),",", - 1,1)
            str = "经实测，本次验线" & TotalJzwnameArr(0) & "正负零标高不满足精度要求。"
            If TotalStr = "" Then
                TotalStr = str
            Else
                TotalStr = str & Chr(13) & TotalStr
            End If
        End If
    Next 'i
    
    If TotalStr = "" Then
        TotalStr = "经资料核查和现场验线测量，《苍南县建设工程放线测量报告》表述内容、注记数据及技术报告格式符合规定，放线测量起算控制点、条件点（验测点）满足精度要求，《苍南县建设工程放线测量报告》中的条件坐标、边长、四至关系与规划许可一致，放线符合规划要求。"
    End If
    g_docObj.Replace "{" & "TEXT" & "}",TotalStr,0
End Function' SetResultTable

'插入照片
Function InsertPhoto()
    Dim f1,fc,f
    filePath = SSProcess.GetSysPathName(5)
    filePath = filePath & "4影像"
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
    If gs <= 3 Then
			For i = 0 To gs
					row = Int(i / 2) + 1
					col = i Mod 2
					'MsgBox filePath & "\" & sarr(i)
					If col = 0 Then
						 'g_docObj.SetCellText Tableindex,hs,8,"符合",True,False
						 g_docObj.SetCellImageEx count, row, 0, - 1, filePath & "\" & sarr(i) ,250,250,false
					Else
						 g_docObj.SetCellImageEx count, row , 1, - 1, filePath & "\" & sarr(i),250,250,false
					End If
			Next 'i
		else
			for i=0 to 3
					row = Int(i / 2) + 1
					col = i Mod 2
					'MsgBox filePath & "\" & sarr(i)
					If col = 0 Then
						 'g_docObj.SetCellText Tableindex,hs,8,"符合",True,False
						 g_docObj.SetCellImageEx count, row, 0, - 1, filePath & "\" & sarr(i) ,250,250,false
					Else
						 g_docObj.SetCellImageEx count, row , 1, - 1, filePath & "\" & sarr(i),250,250,false
					End If
			next
    End If
End Function' InsertPhoto


'数据类型转换
Function transform(content)
    If content <> "" Then
        content = CDbl(content)
    Else
        MsgBox "数据有误"
        Exit Function
    End If
    transform = content
End Function

'绘制辅助线（实测）
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

'绘制辅助线（理论）
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

'删除线
Function DelLine()
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "1,2"
    SSProcess.SelectFilter
    SSProcess.DeleteSelectionObj
End Function' DelLine

'选择辅助线
Function SelLine1(coed)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", coed
    SSProcess.SelectFilter
End Function' SelLine

'选择辅助线
Function SelLine(coed,note)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", coed
    SSProcess.SetSelectCondition "[Note]", "==", note
    SSProcess.SelectFilter
End Function' SelLine

'选择偏差方向
Function SelDiffLine(coed,name)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", coed
    SSProcess.SetSelectCondition "[JianZWMC]", "==", name
    SSProcess.SelectFilter
End Function' SelDiffLine

'选择验线边长
Function SelYxBc(name)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9310053"
    SSProcess.SetSelectCondition "[JianZWMC]", "==", name
    SSProcess.SelectFilter
End Function' SelYxBc

'计算距离差
Function GetLengthDiff(x1,y1,x2,y2)
    diff = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    diff = Round(diff,3)
    GetLengthDiff = diff
End Function' GetLengthDiff

'打开图层
Function allvisible()
    count = SSProcess.GetLayerCount
    For i = 0 To count - 1
        layername = SSProcess.GetLayerName (i)
        SSProcess.SetLayerStatus layername, 1, 1
    Next
    SSProcess.RefreshView
End Function