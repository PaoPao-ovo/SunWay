Dim g_docObj
selectStr = "放线报告(有GPS),放线报告(无GPS)"
Sub OnClick()
    res = AddInputParameter( selectStr, ExportDocType)
    If res = 0  Then Exit Sub
    strTempFileName = ExportDocType & ".doc"
    strTempFilePath = SSProcess.GetSysPathName (7) & "\输出模板\" & strTempFileName
    Set g_docObj = CreateObject ("asposewordscom.asposewordshelper")
    If  TypeName (g_docObj) = "AsposeWordsHelper" Then
        g_docObj.CreateDocumentByTemplate strTempFilePath
    Else
        MsgBox "请先注册Aspose.Word插件"
        Exit Sub
    End If
    pathName = GetFilePath
    InitDB()
    '字符替换
    ReplaceValue
    If InStr(ExportDocType,"有") > 0 Then
        '有GPS
        'GPS-RTK校正检测记录表 
        OutGPSTable 2
        '测站记录表 
        OutStationTable 3
        '放样坐标表 
        OutFYTable 4
        '放样点抽检结果 
        OutFYCheckTable 5
    Else
        '无GPS
        '测站记录表 
        OutStationTable 2
        '放样坐标表 
        OutFYTable 3
        '放样点抽检结果 
        OutFYCheckTable 4
        '控制点平面计算表
        OutControlCountTable 5
        '控制点成果表
        OutControlResultTable 6
        '控制点边长检查表 
        OutControlLengthTable 7
        '段落
        OutPara()
    End If
    
    ReleaseDB()
    strFileSavePath = pathName & "3成果\" & strTempFileName
    g_docObj.SaveEx  strFileSavePath
    Set g_docObj = Nothing
    MsgBox "输出完成"
End Sub

'//字符替换 
Function ReplaceValue
    values = "XiangMBH,XiangMMC,XiangMDZ,JianSDW,WeiTDW,CeHDW,FXDATE,FXXMDATE,ShenPDATE,XiangMFZR,BaoGBZ"
    valuesList = Split(values,",")
    For i = 0 To UBound(valuesList)
        strFieldValue = ""
        strField = valuesList(i)
        listCount = GetProjectTableList ("放验线红线属性表",strField," 放验线红线属性表.ID>0 ","SpatialData","2",list,fieldCount)
        If listCount = 1 Then strFieldValue = list(0,0)
        g_docObj.Replace "{" & strField & "}",strFieldValue,0
    Next
    
    values = "SheJGC"
    valuesList = Split(values,",")
    For i = 0 To UBound(valuesList)
        strFieldValue = ""
        strField = valuesList(i)
        listCount = GetProjectTableList ("正负零标高属性表",strField," 正负零标高属性表.ID>0 ","SpatialData","0",list,fieldCount)
        If listCount = 1 Then strFieldValue = list(0,0)
        g_docObj.Replace "{" & strField & "}",strFieldValue,0
    Next
    
    g_docObj.Replace "{年月日}",Year(Now) & "年" & Month(Now) & "月" & Day(Now) & "日",0
End Function


'获取是否存在GPS检测点,暂时弃用
Function IsExistGPS()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9130215"
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    If geocount > 0 Then IsExistGPS = True Else IsExistGPS = False
End Function

'//选择成果报告
Function AddInputParameter(ByVal selectStr,ByRef ExportDocType)
    res = 1
    title = "输出成果报告"
    selectStrList = Split(selectStr,",")
    If UBound(selectStrList) =  - 1 Then  res = 0
    Exit Function
    If ExportDocType = "" Then  ExportDocType = selectStrList(0)
    SSProcess.ClearInputParameter
    SSProcess.AddInputParameter "放线成果报告", ExportDocType,0,selectStr, "请选择检查类型"
    res = SSProcess.ShowInputParameterDlg (title)
    ExportDocType = SSProcess.GetInputParameter ("放线成果报告" )
    SSProcess.WriteEpsIni title,"放线成果报告",ExportDocType
    AddInputParameter = res
End Function

'***********************************************************报表输出函数***********************************************************

' RTK校正检测记录表 
Function OutGPSTable(ByVal tableIndex)
    g_docObj.MoveToTable tableIndex,False
    ReDim CellList(CellCount)
    CellCount = 0
    '表格行列初始化
    iniRow = 1
    strGPSPointName = ""
    '选择集初始化
    GPSCode = "9130215"
    ControlCode = "9130211,9130212,1102021,1103021"
    '表格处理
    GPSCount = GetFeatureCount( GPSCode, geocount)
    If GPSCount < 0 Then Exit Function
    copyCount = GPSCount * 3 - 1
    '复制行
    g_docObj.CloneTableRow tableIndex,  iniRow, 1,copyCount, False
    For i = 0 To GPSCount - 1
        objID = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        strPointName = SSProcess.GetSelGeoValue(i, "SSObj_PointName")
        'msgbox strPointName
        SSProcess.GetObjectPoint objID, 0, x, y, z, pointtype, name
        'GPS检测点搜索的理论控制点的范围暂定0.1米，后面可能会调整
        ids = SSProcess.SearchNearObjIDs(x, y, 0.1, 0, ControlCode, objID )
        If ids <> "" Then
            
            strControlID = Split(ids,",")
            If UBound(strControlID) = 0 Then
                
                '获取理论控制点的xyz
                SSProcess.GetObjectPoint strControlID(0), 0, x1, y1, z1, pointtype, name
                '整理单元格值
                x = Round(x,3)
                y = Round(y,3)
                z = Round(z,3)
                x1 = Round(x1,3)
                y1 = Round(y1,3)
                z1 = Round(z1,3)
                strChange = Round(Sqr((x1 - x) * (x1 - x) + (y1 - y) * (y1 - y)),3)
                strBZ = ""      '备注不获取值，暂时当作合并单元格标识
                ''单元格数组
                GetValueGPSList CellList,CellCount, strPointName, "X", y, y1, strChange, "{合并备注}"
                GetValueGPSList CellList,CellCount, "", "Y", x, x1, "", ""
                GetValueGPSList CellList,CellCount, "", "Z", z, z1, "{合并删除}", ""
            End If
        End If
        '获取理论控制点点名字符串
        If strGPSPointName = "" Then
            strGPSPointName = strPointName
        Else
            strGPSPointName = strGPSPointName & "、" & strPointName
        End If
    Next 'i
    '填充单元格
    startRow = 1
    strPointChange = ""
    For i = 0 To CellCount - 1
        CellValueList = Split(CellList(i),"||")
        startCol = 0
        '组织点位较差字符串
        If cellValueList(0) <> "" Then
            If strPointChange = "" Then
                strPointChange = cellValueList(4)
            Else
                strPointChange = strPointChange & "," & cellValueList(4)
            End If
        End If
        '填充单元格
        For j = 0 To UBound(CellValueList)
            g_docObj.SetCellText tableIndex,startRow,startCol,cellValueList(j),True,False
            startCol = startCol + 1
        Next
        startRow = startRow + 1
    Next
    '合并单元格
    MergeColValue tableIndex, cellCount, 1, 0
    MergeColValue tableIndex, cellCount, 1, 4
    MergeColValue tableIndex, cellCount, 1, 5
    g_docObj.DeleteRow tableIndex,cellCount + 1,False
    '清除标识
    g_docObj.Replace "{合并删除}","",0
    g_docObj.Replace "{合并备注}","",0
    '获取点位较差最大值
    strPointChangeList = Split(strPointChange,",")
    strMaxChange = ""
    If UBound(strPointChangeList) > 0 Then
        For i = 0 To UBound(strPointChangeList) - 1
            If strPointChangeList(i) > strPointChangeList(i + 1) Then
                strMaxChange = strPointChangeList(i)
            Else
                strMaxChange = strPointChangeList(i + 1)
            End If
        Next
    Else
        strMaxChange = strPointChangeList(0)
    End If
    If strMaxChange <> "" Then strMaxChange = CDbl(strMaxChange * 100)
    g_docObj.Replace "{GPS最大点位较差}",strMaxChange,0
    If strMaxChange < 5.0 Then g_docObj.Replace "{GPS规范要求}","小于5.0cm，符合",0 Else g_docObj.Replace "{GPS规范要求}","大于5.0cm，不符合",0
    '替换段落理论控制点
    g_docObj.Replace "{理论控制点个数}",GPSCount,0
    g_docObj.Replace "{理论控制点点名}",strGPSPointName,0
End Function

'测站记录表
Function OutStationTable(ByVal tableIndex)
    g_docObj.MoveToTable tableIndex,False
    ReDim CellList(CellCount)
    CellCount = 0
    '初始化
    iniRow = 0
    strChange = ""
    '表格处理
    strCZTable = "支点线属性表"
    strFXtable = "方向线属性表"
    strJCtable = "检查线属性表"
    
    strCZField = "CeZDH"
    strFXField = "FangXDH,FangXZ,ShuiPJL"
    strJCField = "JianCDH,FangXZ,ShuiPJL,XZuoBCZ,YZuoBCZ"
    CZlistCount = GetProjectTableList (strFXtable,"distinct CeZDH",strFXtable & ".ID>0 and CeZDH<>'*' ","SpatialData","1",CZlist,fieldCount)
    For i = 0 To CZlistCount - 1
        strCeZDH = CZlist(i,0)
        str = ""
        '获取方向点号
        FXtion = strFXtable & ".ID>0 and " & strFXtable & ".CeZDH = '" & strCeZDH & "'"
        FXlistCount = GetProjectTableList (strFXtable,strFXField,FXtion,"SpatialData","1",FXlist,fieldCount)
        For i1 = 0 To FXlistCount - 1
            strFXDH = FXlist(i1,0)
            strFXZ = FXlist(i1,1)
            strSPJL = FXlist(i1,2)
            strFXDHList = GetString( strFXDH, "," , str)
        Next
        '测站点行
        GetValueCZList  CellList,CellCount, "", "测站点", strCeZDH, strFXDHList, "", "",""
        '方向点标题
        GetValueCZList  CellList,CellCount, "方向点||方向值||水平距离||X坐标差值||Y坐标差值", "", "", "", "", "",""
        For i1 = 0 To FXlistCount - 1
            strFXDH = FXlist(i1,0)
            strFXZ = FXlist(i1,1)
            strSPJL = FXlist(i1,2)
            '方向点行
            GetValueCZList  CellList,CellCount, "", "方向点", strFXDH, strFXZ, strSPJL, "",""
        Next
        '检查点标题
        GetValueCZList  CellList,CellCount, "检查点||方向值||水平距离||X坐标差值||Y坐标差值", "", "", "", "", "",""
        JCtion = strJCtable & ".ID>0 and " & strJCtable & ".CeZDH = '" & strCeZDH & "'"
        JClistCount = GetProjectTableList (strJCtable,strJCField,JCtion,"SpatialData","1",JClist,fieldCount)
        For i1 = 0 To JClistCount - 1
            strJCDH = JClist(i1,0)
            strFXZ = JClist(i1,1)
            strSPJL = JClist(i1,2)
            strX = JClist(i1,3)
            strY = JClist(i1,4)
            '检查点行
            GetValueCZList  CellList,CellCount, "", "检查点", strJCDH, strFXZ, strSPJL, strX,strY
            '获取点位较差字符串
            strXY = Round(Sqr(strX * strX + strY * strY),3)
            If strChange = "" Then
                strChange = strXY
            Else
                strChange = strChange & "," & strXY
            End If
        Next
    Next
    '复制行
    copyCount = CellCount - 1
    g_docObj.CloneTableRow tableIndex,  iniRow, 1,copyCount, False
    '填充单元格
    startRow = 0
    For i = 0 To CellCount - 1
        CellValueList = Split(CellList(i),"||")
        startCol = 0
        For j = 0 To UBound(CellValueList)
            g_docObj.SetCellText tableIndex,startRow,startCol,cellValueList(j),True,False
            startCol = startCol + 1
        Next
        startRow = startRow + 1
    Next
    '获取最大点位较差
    strChangeList = Split(strChange,",")
    If UBound(strChangeList) > 0 Then
        For i = 0 To UBound(strChangeList) - 1
            If strChangeList(i) > strChangeList(i + 1) Then
                strMaxChange = strChangeList(i)
            Else
                strMaxChange = strChangeList(i + 1)
            End If
        Next
    Else
        strMaxChange = strChangeList(0)
    End If
    If strMaxChange <> "" Then strMaxChange = CDbl(strMaxChange * 100)
    g_docObj.Replace "{测站最大点位较差}",strMaxChange,0
    If strMaxChange < 5.0 Then g_docObj.Replace "{测站规范要求}","小于5.0cm，符合",0  Else   g_docObj.Replace "{测站规范要求}","大于5.0cm，不符合",0
    
End Function


'放样坐标表
Function OutFYTable(ByVal tableIndex)
    g_docObj.MoveToTable tableIndex,False
    ReDim CellList(CellCount)
    CellCount = 0
    iniRow = 1
    CopyCount = 0
    '表格处理
    geocount = GetFeatureCount( "9310013", geocount)
    For i = 0 To geocount - 1
        objID = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        JianZWMC = SSProcess.GetSelGeoValue(i, "[JianZWMC]")
        pointcount = SSProcess.GetSelGeoPointCount(i)
        '复制的行数
        CopyCount = CopyCount + pointcount - 1
        For i1 = 0 To pointcount - 2
            SSProcess.GetObjectPoint objID, i1, x0, y0, z0, pointtype, name
            '获取下一个角点坐标
            SSProcess.GetObjectPoint objID, i1 + 1, x1, y1, z1, pointtype, name
            x0 = Round(x0,3)
            y0 = Round(y0,3)
            z0 = Round(z0,3)
            x1 = Round(x1,3)
            y1 = Round(y1,3)
            z1 = Round(z1,3)
            '获取边长值
            strChange = Round(Sqr((x1 - x0) * (x1 - x0) + (y1 - y0) * (y1 - y0)),3)
            '获取下一个角点的点名
            ids = SSProcess.SearchNearObjIDs(x1, y1, 0.1, 0, "9130411", objID )
            If ids <> "" Then
                strControlID = Split(ids,",")
                For i2 = 0 To UBound(strControlID)
                    '获取属性
                    LiLunPointName = SSProcess.GetObjectAttr(strControlID(i2), "SSObj_PointName")
                    LiLunJianZWMC = SSProcess.GetObjectAttr(strControlID(i2), "[JianZWMC]")
                    If LiLunJianZWMC = JianZWMC Then
                        LiLunPointName1 = LiLunPointName
                    End If
                Next
            End If
            
            '获取当前角点属性
            ids = SSProcess.SearchNearObjIDs(x0, y0, 0.1, 0, "9130411", objID )
            If ids <> "" Then
                strControlID = Split(ids,",")
                For i2 = 0 To UBound(strControlID)
                    '获取属性
                    LiLunPointName = SSProcess.GetObjectAttr(strControlID(i2), "SSObj_PointName")
                    LiLunJianZWMC = SSProcess.GetObjectAttr(strControlID(i2), "[JianZWMC]")
                    strPointChange = LiLunPointName & "-" & LiLunPointName1
                    If LiLunJianZWMC = JianZWMC Then
                        If i1 = 0 Then
                            CellValue = JianZWMC & "||" & LiLunPointName & "||" & y0 & "||" & x0 & "||" & strPointChange & "||" & strChange
                        Else
                            CellValue = "" & "||" & LiLunPointName & "||" & y0 & "||" & x0 & "||" & strPointChange & "||" & strChange
                        End If
                        ReDim Preserve CellList(CellCount)
                        CellList(CellCount) = CellValue
                        CellCount = CellCount + 1
                    End If
                Next
            End If
        Next
    Next
    '复制行
    g_docObj.CloneTableRow tableIndex,  iniRow, 1,CopyCount - 1, False
    '填充单元格
    startRow = 1
    For i = 0 To CellCount - 1
        CellValueList = Split(CellList(i),"||")
        startCol = 0
        For j = 0 To UBound(CellValueList)
            g_docObj.SetCellText tableIndex,startRow,startCol,cellValueList(j),True,False
            startCol = startCol + 1
        Next
        startRow = startRow + 1
    Next
    '合并单元格
    MergeColValue tableIndex, cellCount, 1, 0
    g_docObj.DeleteRow tableIndex,cellCount + 1,False
End Function

'放样点抽检结果
Function OutFYCheckTable(ByVal tableIndex)
    g_docObj.MoveToTable tableIndex,False
    ReDim CellList(CellCount)
    CellCount = 0
    iniRow = 1
    '表格处理
    geocount = GetFeatureCount("9130511", geocount)
    For i = 0 To geocount - 1
        '实测放样点点名和坐标
        objID = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        PointName = SSProcess.GetSelGeoValue(i, "SSObj_PointName")
        SSProcess.GetObjectPoint objID, 0, x1, y1, z1, pointtype, name
        x1 = Round(x1,3)
        y1 = Round(y1,3)
        z1 = Round(z1,3)
        '空间搜索理论放样点
        ids = SSProcess.SearchNearObjIDs(x1, y1, 0.1, 0, "9130411", objID )
        If ids <> "" Then
            strControlID = Split(ids,",")
            For i1 = 0 To UBound(strControlID)
                LiLunPointName = SSProcess.GetObjectAttr(strControlID(i1), "SSObj_PointName")
                SSProcess.GetObjectPoint strControlID(i1), 0, x0, y0, z0, pointtype, name
                x0 = Round(x0,3)
                y0 = Round(y0,3)
                z0 = Round(z0,3)
                '获取同名理论放样点的坐标，边长
                If LiLunPointName = PointName Then
                    strChange = Round(Sqr((x1 - x0) * (x1 - x0) + (y1 - y0) * (y1 - y0)),3)
                    CellValue = PointName & "||" & y0 & "||" & x0 & "||" & y1 & "||" & x1 & "||" & strChange
                    ReDim Preserve CellList(CellCount)
                    CellList(CellCount) = CellValue
                    CellCount = CellCount + 1
                End If
            Next
        End If
    Next
    '复制行
    copyCount = geocount - 1
    g_docObj.CloneTableRow tableIndex,  iniRow, 1,copyCount, False
    '填充单元格
    startRow = 1
    strPointChange = ""
    For i = 0 To CellCount - 1
        CellValueList = Split(CellList(i),"||")
        startCol = 0
        '组织点位较差字符串
        strPointChange = GetString( cellValueList(5), "," , strPointChange)
        '填充单元格
        For j = 0 To UBound(CellValueList)
            g_docObj.SetCellText tableIndex,startRow,startCol,cellValueList(j),True,False
            startCol = startCol + 1
        Next
        startRow = startRow + 1
    Next
    '获取点位较差最大值
    strPointChangeList = Split(strPointChange,",")
    strMaxChange = ""
    For i = 0 To UBound(strPointChangeList) - 1
        If strPointChangeList(i) > strPointChangeList(i + 1) Then
            strMaxChange = strPointChangeList(i)
        Else
            strMaxChange = strPointChangeList(i + 1)
        End If
    Next
    If strMaxChange <> "" Then strMaxChange = CDbl(strMaxChange * 100)
    g_docObj.Replace "{放样点最大点位较差}",strMaxChange,0
    If strMaxChange < 5.0 Then g_docObj.Replace "{放样点规范要求}","小于5.0cm，符合",0  Else   g_docObj.Replace "{放样点规范要求}","大于5.0cm，不符合",0
End Function


'控制点平面计算表
Function OutControlCountTable(ByVal tableIndex)
    g_docObj.MoveToTable tableIndex,False
    ReDim CellList(CellCount)
    CellCount = 0
    iniRow = 1
    strPointName = ""
    xhCount = 1
    '表格处理
    geocount = GetFeatureCount("1130211", geocount)
    For i = 0 To geocount - 1
        objID = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        PointName = SSProcess.GetSelGeoValue(i, "SSObj_PointName")
        '点名去重       
        If strPointName = "" Then
            strPointName = "," & PointName & ","
        ElseIf InStr(strPointName,"," & PointName & ",") = 0 Then
            strPointName = strPointName & PointName & ","
        End If
    Next
    strPointName = Mid(strPointName,2,Len(strPointName) - 2)
    strPointNameList = Split(strPointName,",")
    For i = 0 To UBound(strPointNameList)
        ControlCount = GetFeatureCount("1130211", geocount)
        strControlID = ""
        For i1 = 0 To ControlCount - 1
            objID = SSProcess.GetSelGeoValue(i1, "SSObj_ID")
            PointName = SSProcess.GetSelGeoValue(i1, "SSObj_PointName")
            If PointName = strPointNameList(i) Then
                strControlID = GetString( objID, "," , strControlID)
            End If
        Next
        strControlIDList = Split(strControlID,",")
        If UBound(strControlIDList) = 1 Then
            id0 = strControlIDList(0)
            id1 = strControlIDList(1)
            SSProcess.GetObjectPoint id0, 0, x0, y0, z0, pointtype, name
            SSProcess.GetObjectPoint id1, 0, x1, y1, z1, pointtype, name
            x0 = Round(x0,3)
            y0 = Round(y0,3)
            z0 = Round(z0,3)
            x1 = Round(x1,3)
            y1 = Round(y1,3)
            z1 = Round(z1,3)
            strChange = Round(Sqr((x1 - x0) * (x1 - x0) + (y1 - y0) * (y1 - y0)),3)
            '获取同名实测控制点个数
            SCControlCount = GetFeatureCount("9130311,9130312,9130412,9130512", geocount)
            For i1 = 0 To SCControlCount - 1
                objID = SSProcess.GetSelGeoValue(i1, "SSObj_ID")
                PointName = SSProcess.GetSelGeoValue(i1, "SSObj_PointName")
                If PointName = strPointNameList(i) Then
                    SSProcess.GetObjectPoint objID, 0, x, y, z, pointtype, name
                    x = Round(x,3)
                    y = Round(y,3)
                    z0 = Round(z,3)
                End If
            Next
            '获取单元格数组
            CellValue = xhCount & "||" & strPointNameList(i) & "||" & y0 & "||" & y1 & "||" & strChange & "||" & y
            ReDim Preserve CellList(CellCount)
            CellList(CellCount) = CellValue
            CellCount = CellCount + 1
            xhCount = xhCount + 1
            CellValue = "" & "||" & "" & "||" & x0 & "||" & x1 & "||" & "" & "||" & x
            ReDim Preserve CellList(CellCount)
            CellList(CellCount) = CellValue
            CellCount = CellCount + 1
        End If
    Next
    '复制行
    copyCount = (UBound(strPointNameList) + 1) * 2 - 1
    g_docObj.CloneTableRow tableIndex,  iniRow, 1,copyCount, False
    '填充单元格
    startRow = 1
    For i = 0 To CellCount - 1
        CellValueList = Split(CellList(i),"||")
        startCol = 0
        '填充单元格
        For j = 0 To UBound(CellValueList)
            g_docObj.SetCellText tableIndex,startRow,startCol,cellValueList(j),True,False
            startCol = startCol + 1
        Next
        startRow = startRow + 1
    Next
    '合并单元格
    MergeColValue tableIndex, cellCount, 1, 0
    MergeColValue tableIndex, cellCount, 1, 1
    MergeColValue tableIndex, cellCount, 1, 4
    g_docObj.DeleteRow tableIndex,cellCount + 1,False
End Function

'控制点成果表
Function OutControlResultTable(ByVal tableIndex)
    g_docObj.MoveToTable tableIndex,False
    ReDim CellList(CellCount)
    CellCount = 0
    iniRow = 1
    '获取理论控制点
    geocount = GetFeatureCount( "1103021,1102021,9130211,9130212", geocount)
    For i = 0 To geocount - 1
        objID = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        PointName = SSProcess.GetSelGeoValue(i, "SSObj_PointName")
        SSProcess.GetObjectPoint objID, 0, x, y, z, pointtype, name
        x = Round(x,3)
        y = Round(y,3)
        z = Round(z,3)
        strBZ = ""
        CellValue = PointName & "||" & x & "||" & y & "||" & z & "||" & strBZ
        ReDim Preserve CellList(CellCount)
        CellList(CellCount) = CellValue
        CellCount = CellCount + 1
    Next
    '复制行
    copyCount = geocount - 1
    g_docObj.CloneTableRow tableIndex,  iniRow, 1,copyCount, False
    '填充单元格
    startRow = 1
    For i = 0 To CellCount - 1
        CellValueList = Split(CellList(i),"||")
        startCol = 0
        '填充单元格
        For j = 0 To UBound(CellValueList)
            g_docObj.SetCellText tableIndex,startRow,startCol,cellValueList(j),True,False
            startCol = startCol + 1
        Next
        startRow = startRow + 1
    Next
End Function

'控制点边长检查表
Function OutControlLengthTable(ByVal tableIndex)
    g_docObj.MoveToTable tableIndex,False
    ReDim CellList(CellCount)
    CellCount = 0
    iniRow = 1
    xhCount = 1
    '表格处理
    strJCtable = "控制点检查线属性表"
    strJCField = "CeZDH,JianCDH,YZBC,JCBC,BCJC"
    JClistCount = GetProjectTableList (strJCtable,strJCField,strJCtable & ".ID>0 and CeZDH<>'*' ","SpatialData","1",JClist,fieldCount)
    For i = 0 To JClistCount - 1
        CeZDH = JClist(i,0)
        JianCDH = JClist(i,1)
        YZBC = JClist(i,2)
        JCBC = JClist(i,3)
        BCJC = JClist(i,4)
        strDH = CeZDH & "-" & JianCDH
        BCJC = Round(BCJC,3)
        CellValue = xhCount & "||" & strDH & "||" & YZBC & "||" & JCBC & "||" & BCJC
        ReDim Preserve CellList(CellCount)
        CellList(CellCount) = CellValue
        CellCount = CellCount + 1
        xhCount = xhCount + 1
    Next
    '复制行
    copyCount = JClistCount - 1
    g_docObj.CloneTableRow tableIndex,  iniRow, 1,copyCount, False
    '填充单元格
    startRow = 1
    For i = 0 To CellCount - 1
        CellValueList = Split(CellList(i),"||")
        startCol = 0
        '填充单元格
        For j = 0 To UBound(CellValueList)
            g_docObj.SetCellText tableIndex,startRow,startCol,cellValueList(j),True,False
            startCol = startCol + 1
        Next
        startRow = startRow + 1
    Next
End Function

'无GPS段落
Function OutPara()
    ControlCode = "9130212,9130211,1102021,1103021"
    ControlCount = GetFeatureCount( ControlCode, geocount)
    strControlPointName = ""
    For i = 0 To ControlCount - 1
        objID = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        strPointName = SSProcess.GetSelGeoValue(i, "SSObj_PointName")
        strControlPointName = GetString(strPointName, "," , strControlPointName)
    Next
    g_docObj.Replace "{理论控制点点名}",strControlPointName,0
    
    
    strJCtable = "控制点检查线属性表"
    strJCField = "CeZDH,JianCDH,YZBC,JCBC,BCJC"
    JClistCount = GetProjectTableList (strJCtable,"max(BCJC)",strJCtable & ".ID>0 and CeZDH<>'*' ","SpatialData","1",JClist,fieldCount)
    If JClistCount = 1 Then strMaxChange = Round(JClist(0,0),3)
    If strMaxChange <> "" Then strMaxChange = CDbl(strMaxChange * 100)
    g_docObj.Replace "{最大边长较差}",strMaxChange,0
    If strMaxChange < 5.0 Then g_docObj.Replace "{边长规范要求}","小于5.0cm，符合",0 Else g_docObj.Replace "{边长规范要求}","大于5.0cm，不符合",0
End Function


'*****************************表格辅助整理函数*******************************
' 通过选择集，获取要素数量，注意后面清空选择集操作，在开头运行
Function GetFeatureCount(ByVal Code,ByRef geocount)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code","==",Code
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    GetFeatureCount = geocount
End Function

'整理出字符串
Function GetString(ByVal value,ByVal splitMark , str)
    If str = "" Then
        str = value
    Else
        str = str & splitMark & value
    End If
    GetString = str
End Function' Name

'合并列
Function MergeColValue(ByVal tableIndex,ByVal cellCount,ByVal startRow,ByVal startCol)
    allxhValue = ""
    For i = 0 To cellCount
        xhValue = g_docObj.GetCellText( tableIndex, startRow + i, startCol,False)
        xhValue = Replace(xhValue,"","")
        If allxhValue = "" Then
            allxhValue = xhValue
        Else
            allxhValue = allxhValue & "||" & xhValue
        End If
    Next
    allxhValueList = Split(allxhValue,"||")
    For i = 0 To UBound(allxhValueList)
        If i > 0 And allxhValueList(i) <> "" Then
            ReDim Preserve MergeList(MergeCount)
            MergeList(MergeCount) = startRow & "||" & startRow + MergeRow - 1
            MergeCount = MergeCount + 1
            startRow = startRow + MergeRow
            MergeRow = 0
        End If
        MergeRow = MergeRow + 1
    Next
    For i = 0 To MergeCount - 1
        MergeListValue = Split(MergeList(i),"||")
        g_docObj.MergeCell tableIndex,  MergeListValue(0),  startCol,  MergeListValue(1), startCol,False
    Next
End Function

'***********************获取单元格属性函数**********************************
'GPS检测记录表
Function GetValueGPSList(CellList,CellCount,ByVal strPointName,ByVal strType,ByVal strTypeValue,ByVal strSCValue,ByVal strChange,ByVal strBZ)
    cellValue = ""
    value = strPointName & "||" & strType & "||" & strTypeValue & "||" & strSCValue & "||" & strChange & "||" & strBZ
    cellValue = value
    ReDim Preserve CellList(CellCount)
    CellList(CellCount) = cellValue
    CellCount = CellCount + 1
End Function

'测站记录表
Function GetValueCZList(CellList,CellCount,ByVal strtitle,ByVal strPointType,ByVal strDH,ByVal strFXZ,ByVal strSPJL,ByVal strX,ByVal strY)
    If strPointType = "测站点" Then
        CellValue = strPointType & "||" & strDH & "||" & "" & "||" & "方向点" & "||" & strFXZ
        ReDim Preserve CellList(CellCount)
        CellList(CellCount) = CellValue
        CellCount = CellCount + 1
    ElseIf strPointType = "方向点" Then
        CellValue = strDH & "||" & strFXZ & "||" & strSPJL
        ReDim Preserve CellList(CellCount)
        CellList(CellCount) = CellValue
        CellCount = CellCount + 1
    ElseIf strPointType = "检查点" Then
        CellValue = strDH & "||" & strFXZ & "||" & strSPJL & "||" & strX & "||" & strY
        ReDim Preserve CellList(CellCount)
        CellList(CellCount) = CellValue
        CellCount = CellCount + 1
    Else
        CellValue = strtitle
        ReDim Preserve CellList(CellCount)
        CellList(CellCount) = CellValue
        CellCount = CellCount + 1
    End If
End Function


'***********************************************************数据库操作函数***********************************************************
'//strTableName 表
'//strFields 字段
'//strAddCondition 条件 
'//strTableType "AttributeData（纯属性表） ,SpatialData（地物属性表）" 
'//strGeoType 地物类型 点、线、面、注记(0点，1线，2面，3注记)
'//rs 表记录二维数组rs(行,列)
'//fieldCount 字段个数
'//返回值 ：sql查询表记录个数
Function GetProjectTableList(ByVal strTableName,ByVal strFields,ByVal strAddCondition,ByVal strTableType,ByVal strGeoType,ByRef rs(),ByRef fieldCount)
    GetProjectTableList = 0
    values = ""
    rsCount = 0
    fieldCount = 0
    If strTableName = "" Or strFields = "" Then Exit Function
    '设置地物类型
    If strGeoType = "0" Then
        GeoType = "GeoPointTB"
    ElseIf strGeoType = "1" Then
        GeoType = "GeoLineTB"
    ElseIf strGeoType = "2" Then
        GeoType = "GeoAreaTB"
    ElseIf strGeoType = "3" Then
        GeoType = "MarkNoteTB"
    Else
        GeoType = "GeoAreaTB"
    End If
    If strTableType = "SpatialData" Then
        strCondition = " (" & GeoType & ".Mark Mod 2)<>0"
        If strAddCondition <> "" Then      strCondition = " (" & GeoType & ".Mark Mod 2)<>0 and " & strAddCondition & ""
        sql = "select  " & strFields & " from " & strTableName & "  INNER JOIN " & GeoType & " ON " & strTableName & ".ID = " & GeoType & ".ID WHERE " & strCondition & ""
    Else
        If strAddCondition <> "" Then
            strCondition = strAddCondition
            sql = "select  " & strFields & " from " & strTableName & "  WHERE  " & strCondition & ""
        Else
            sql = "select  " & strFields & " from " & strTableName & ""
        End If
    End If
    
    'if instr(sql,"scpcjzmj")>0 then  addloginfo sql
    '获取当前工程edb表记录
    AccessName = SSProcess.GetProjectFileName
    '判断表是否存在
    'if  IsTableExits(AccessName,strTableName)=false then exit function 
    'set adoConnection=createobject("adodb.connection")
    'strcon="DBQ="& AccessName &";DRIVER={Microsoft Access Driver (*.mdb)};"  
    'adoConnection.Open strcon
    Set adoRs = CreateObject("ADODB.recordset")
    count = 0
    adoRs.cursorLocation = 3
    adoRs.cursorType = 3
    'msgbox sql
    adoRs.open sql,adoConnection,3,3
    rcdCount = adoRs.RecordCount
    fieldCount = adoRs.Fields.Count
    ReDim rs(rcdCount,fieldCount)
    'erase rs
    While adoRs.Eof = False
        nowValues = ""
        For i = 0 To fieldCount - 1
            value = adoRs(i)
            If IsNull(value) Then value = ""
            value = Replace(value,",","，")
            rs(rsCount,i) = value
        Next
        rsCount = rsCount + 1
        adoRs.MoveNext
    WEnd
    adoRs.Close
    Set adoRs = Nothing
    'adoConnection.Close
    'Set adoConnection = Nothing
    GetProjectTableList = rsCount
End Function

Function GetSQLRecordAll(ByRef mdbName, ByRef StrSqlStatement, ByRef arSQLRecord(), ByRef iRecordCount)
    If StrSqlStatement = "" Then
        MsgBox "查询语句为空，操作停止！",48
    End If
    iRecordCount =  - 1
    'SQL语句
    sql = StrSqlStatement
    '打开记录集
    SSProcess.OpenAccessRecordset mdbName, sql
    '获取记录总数
    RecordCount = SSProcess.GetAccessRecordCount (mdbName, sql)
    If RecordCount > 0 Then
        iRecordCount = 0
        ReDim arSQLRecord(RecordCount)
        '将记录游标移到第一行
        SSProcess.AccessMoveFirst mdbName, sql
        iRecordCount = 0
        '浏览记录
        While SSProcess.AccessIsEOF (mdbName, sql) = 0
            fields = ""
            values = ""
            '获取当前记录内容
            SSProcess.GetAccessRecord mdbName, sql, fields, values
            arSQLRecord(iRecordCount) = values                                        '查询记录
            iRecordCount = iRecordCount + 1                                                    '查询记录数
            '移动记录游标
            SSProcess.AccessMoveNext mdbName, sql
        WEnd
    End If
    '关闭记录集
    SSProcess.CloseAccessRecordset mdbName, sql
End Function

'//开库
Dim  adoConnection
Function InitDB()
    accessName = SSProcess.GetProjectFileName
    Set adoConnection = CreateObject("adodb.connection")
    strcon = "DBQ=" & accessName & ";DRIVER={Microsoft Access Driver (*.mdb)};"
    adoConnection.Open strcon
End Function

'//关库
Function ReleaseDB()
    adoConnection.Close
    Set adoConnection = Nothing
End Function

'改路径
'//获取成果目录路径
Function  GetFilePath
    projectFileName = SSProcess.GetSysPathName (5)
    filePath = Replace(projectFileName,".edb","")
    filePath = filePath & "\"
    CreateFolder filePath
    GetFilePath = filePath
End Function

'//递归创建多级目录
Function CreateFolder(path)
    Set fso = CreateObject("scripting.filesystemobject")
    If fso.FolderExists(path) Then
        Exit Function
    End If
    If Not fso.FolderExists(fso.GetParentFolderName(path)) Then
        CreateFolder fso.GetParentFolderName(path)
    End If
    fso.CreateFolder(path)
    Set fso = Nothing
End Function