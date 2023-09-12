Dim g_docObj
selectStr = "???????(??GPS),???????(??GPS)"
Sub OnClick()
    res = AddInputParameter( selectStr, ExportDocType)
    If res = 0  Then Exit Sub
    strTempFileName = ExportDocType & ".doc"
    strTempFilePath = SSProcess.GetSysPathName (7) & "\??????\" & strTempFileName
    Set g_docObj = CreateObject ("asposewordscom.asposewordshelper")
    If  TypeName (g_docObj) = "AsposeWordsHelper" Then
        g_docObj.CreateDocumentByTemplate strTempFilePath
    Else
        MsgBox "???????Aspose.Word???"
        Exit Sub
    End If
    pathName = GetFilePath
    InitDB()
    '????ÿI
    ReplaceValue
    If InStr(ExportDocType,"??") > 0 Then
        '??GPS
        'GPS-RTK§µ????????? 
        OutGPSTable 2
        '???????? 
        OutStationTable 3
        '????????? 
        OutFYTable 4
        '?????????? 
        OutFYCheckTable 5
    Else
        '??GPS
        '???????? 
        OutStationTable 2
        '????????? 
        OutFYTable 3
        '?????????? 
        OutFYCheckTable 4
        '????????????
        OutControlCountTable 5
        '?????????
        OutControlResultTable 6
        '??????????? 
        OutControlLengthTable 7
        '????
        OutPara()
    End If
    
    ReleaseDB()
    strFileSavePath = pathName & "3???\" & strTempFileName
    g_docObj.SaveEx  strFileSavePath
    Set g_docObj = Nothing
    MsgBox "??????"
End Sub

'//????ÿI 
Function ReplaceValue
    values = "XiangMBH,XiangMMC,XiangMDZ,JianSDW,WeiTDW,CeHDW,FXDATE,FXXMDATE,ShenPDATE,XiangMFZR,BaoGBZ"
    valuesList = Split(values,",")
    For i = 0 To UBound(valuesList)
        strFieldValue = ""
        strField = valuesList(i)
        listCount = GetProjectTableList ("??????????????",strField," ??????????????.ID>0 ","SpatialData","2",list,fieldCount)
        If listCount = 1 Then strFieldValue = list(0,0)
        g_docObj.Replace "{" & strField & "}",strFieldValue,0
    Next
    
    values = "SheJGC"
    valuesList = Split(values,",")
    For i = 0 To UBound(valuesList)
        strFieldValue = ""
        strField = valuesList(i)
        listCount = GetProjectTableList ("?????????????",strField," ?????????????.ID>0 ","SpatialData","0",list,fieldCount)
        If listCount = 1 Then strFieldValue = list(0,0)
        g_docObj.Replace "{" & strField & "}",strFieldValue,0
    Next
    
    g_docObj.Replace "{??????}",Year(Now) & "??" & Month(Now) & "??" & Day(Now) & "??",0
End Function


'?????????GPS????,???????
Function IsExistGPS()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9130215"
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    If geocount > 0 Then IsExistGPS = True Else IsExistGPS = False
End Function

'//?????????
Function AddInputParameter(ByVal selectStr,ByRef ExportDocType)
    res = 1
    title = "??????????"
    selectStrList = Split(selectStr,",")
    If UBound(selectStrList) =  - 1 Then  res = 0
    Exit Function
    If ExportDocType = "" Then  ExportDocType = selectStrList(0)
    SSProcess.ClearInputParameter
    SSProcess.AddInputParameter "??????????", ExportDocType,0,selectStr, "???????????"
    res = SSProcess.ShowInputParameterDlg (title)
    ExportDocType = SSProcess.GetInputParameter ("??????????" )
    SSProcess.WriteEpsIni title,"??????????",ExportDocType
    AddInputParameter = res
End Function

'***********************************************************???????????***********************************************************

' RTK§µ????????? 
Function OutGPSTable(ByVal tableIndex)
    g_docObj.MoveToTable tableIndex,False
    ReDim CellList(CellCount)
    CellCount = 0
    '???????§Ô????
    iniRow = 1
    strGPSPointName = ""
    '???????
    GPSCode = "9130215"
    ControlCode = "9130211,9130212,1102021,1103021"
    '??????
    GPSCount = GetFeatureCount( GPSCode, geocount)
    If GPSCount < 0 Then Exit Function
    copyCount = GPSCount * 3 - 1
    '??????
    g_docObj.CloneTableRow tableIndex,  iniRow, 1,copyCount, False
    For i = 0 To GPSCount - 1
        objID = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        strPointName = SSProcess.GetSelGeoValue(i, "SSObj_PointName")
        'msgbox strPointName
        SSProcess.GetObjectPoint objID, 0, x, y, z, pointtype, name
        'GPS????????????????????¦¶???0.1??????????????
        ids = SSProcess.SearchNearObjIDs(x, y, 0.1, 0, ControlCode, objID )
        If ids <> "" Then
            
            strControlID = Split(ids,",")
            If UBound(strControlID) = 0 Then
                
                '????????????xyz
                SSProcess.GetObjectPoint strControlID(0), 0, x1, y1, z1, pointtype, name
                '??????????
                x = Round(x,3)
                y = Round(y,3)
                z = Round(z,3)
                x1 = Round(x1,3)
                y1 = Round(y1,3)
                z1 = Round(z1,3)
                strChange = Round(Sqr((x1 - x) * (x1 - x) + (y1 - y) * (y1 - y)),3)
                strBZ = ""      '????????????????????????????
                ''?????????
                GetValueGPSList CellList,CellCount, strPointName, "X", y, y1, strChange, "{??????}"
                GetValueGPSList CellList,CellCount, "", "Y", x, x1, "", ""
                GetValueGPSList CellList,CellCount, "", "Z", z, z1, "{??????}", ""
            End If
        End If
        '???????????????????
        If strGPSPointName = "" Then
            strGPSPointName = strPointName
        Else
            strGPSPointName = strGPSPointName & "??" & strPointName
        End If
    Next 'i
    '??????
    startRow = 1
    strPointChange = ""
    For i = 0 To CellCount - 1
        CellValueList = Split(CellList(i),"||")
        startCol = 0
        '?????¦Ë????????
        If cellValueList(0) <> "" Then
            If strPointChange = "" Then
                strPointChange = cellValueList(4)
            Else
                strPointChange = strPointChange & "," & cellValueList(4)
            End If
        End If
        '??????
        For j = 0 To UBound(CellValueList)
            g_docObj.SetCellText tableIndex,startRow,startCol,cellValueList(j),True,False
            startCol = startCol + 1
        Next
        startRow = startRow + 1
    Next
    '????????
    MergeColValue tableIndex, cellCount, 1, 0
    MergeColValue tableIndex, cellCount, 1, 4
    MergeColValue tableIndex, cellCount, 1, 5
    g_docObj.DeleteRow tableIndex,cellCount + 1,False
    '??????
    g_docObj.Replace "{??????}","",0
    g_docObj.Replace "{??????}","",0
    '?????¦Ë???????
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
    g_docObj.Replace "{GPS????¦Ë???}",strMaxChange,0
    If strMaxChange < 5.0 Then g_docObj.Replace "{GPS?ÿZ???}","§³??5.0cm??????",0 Else g_docObj.Replace "{GPS?ÿZ???}","????5.0cm????????",0
    '?ÿI????????????
    g_docObj.Replace "{???????????}",GPSCount,0
    g_docObj.Replace "{???????????}",strGPSPointName,0
End Function

'????????
Function OutStationTable(ByVal tableIndex)
    g_docObj.MoveToTable tableIndex,False
    ReDim CellList(CellCount)
    CellCount = 0
    '?????
    iniRow = 0
    strChange = ""
    '??????
    strCZTable = "??????????"
    strFXtable = "???????????"
    strJCtable = "??????????"
    
    strCZField = "CeZDH"
    strFXField = "FangXDH,FangXZ,ShuiPJL"
    strJCField = "JianCDH,FangXZ,ShuiPJL,XZuoBCZ,YZuoBCZ"
    CZlistCount = GetProjectTableList (strFXtable,"distinct CeZDH",strFXtable & ".ID>0 and CeZDH<>'*' ","SpatialData","1",CZlist,fieldCount)
    For i = 0 To CZlistCount - 1
        strCeZDH = CZlist(i,0)
        str = ""
        '?????????
        FXtion = strFXtable & ".ID>0 and " & strFXtable & ".CeZDH = '" & strCeZDH & "'"
        FXlistCount = GetProjectTableList (strFXtable,strFXField,FXtion,"SpatialData","1",FXlist,fieldCount)
        For i1 = 0 To FXlistCount - 1
            strFXDH = FXlist(i1,0)
            strFXZ = FXlist(i1,1)
            strSPJL = FXlist(i1,2)
            strFXDHList = GetString( strFXDH, "," , str)
        Next
        '???????
        GetValueCZList  CellList,CellCount, "", "?????", strCeZDH, strFXDHList, "", "",""
        '????????
        GetValueCZList  CellList,CellCount, "?????||?????||??????||X??????||Y??????", "", "", "", "", "",""
        For i1 = 0 To FXlistCount - 1
            strFXDH = FXlist(i1,0)
            strFXZ = FXlist(i1,1)
            strSPJL = FXlist(i1,2)
            '???????
            GetValueCZList  CellList,CellCount, "", "?????", strFXDH, strFXZ, strSPJL, "",""
        Next
        '???????
        GetValueCZList  CellList,CellCount, "????||?????||??????||X??????||Y??????", "", "", "", "", "",""
        JCtion = strJCtable & ".ID>0 and " & strJCtable & ".CeZDH = '" & strCeZDH & "'"
        JClistCount = GetProjectTableList (strJCtable,strJCField,JCtion,"SpatialData","1",JClist,fieldCount)
        For i1 = 0 To JClistCount - 1
            strJCDH = JClist(i1,0)
            strFXZ = JClist(i1,1)
            strSPJL = JClist(i1,2)
            strX = JClist(i1,3)
            strY = JClist(i1,4)
            '??????
            GetValueCZList  CellList,CellCount, "", "????", strJCDH, strFXZ, strSPJL, strX,strY
            '?????¦Ë????????
            strXY = Round(Sqr(strX * strX + strY * strY),3)
            If strChange = "" Then
                strChange = strXY
            Else
                strChange = strChange & "," & strXY
            End If
        Next
    Next
    '??????
    copyCount = CellCount - 1
    g_docObj.CloneTableRow tableIndex,  iniRow, 1,copyCount, False
    '??????
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
    '???????¦Ë???
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
    g_docObj.Replace "{???????¦Ë???}",strMaxChange,0
    If strMaxChange < 5.0 Then g_docObj.Replace "{????ÿZ???}","§³??5.0cm??????",0  Else   g_docObj.Replace "{????ÿZ???}","????5.0cm????????",0
    
End Function


'?????????
Function OutFYTable(ByVal tableIndex)
    g_docObj.MoveToTable tableIndex,False
    ReDim CellList(CellCount)
    CellCount = 0
    iniRow = 1
    CopyCount = 0
    '??????
    geocount = GetFeatureCount( "9310013", geocount)
    For i = 0 To geocount - 1
        objID = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        JianZWMC = SSProcess.GetSelGeoValue(i, "[JianZWMC]")
        pointcount = SSProcess.GetSelGeoPointCount(i)
        '?????????
        CopyCount = CopyCount + pointcount - 1
        For i1 = 0 To pointcount - 2
            SSProcess.GetObjectPoint objID, i1, x0, y0, z0, pointtype, name
            '???????????????
            SSProcess.GetObjectPoint objID, i1 + 1, x1, y1, z1, pointtype, name
            x0 = Round(x0,3)
            y0 = Round(y0,3)
            z0 = Round(z0,3)
            x1 = Round(x1,3)
            y1 = Round(y1,3)
            z1 = Round(z1,3)
            '???????
            strChange = Round(Sqr((x1 - x0) * (x1 - x0) + (y1 - y0) * (y1 - y0)),3)
            '???????????????
            ids = SSProcess.SearchNearObjIDs(x1, y1, 0.1, 0, "9130411", objID )
            If ids <> "" Then
                strControlID = Split(ids,",")
                For i2 = 0 To UBound(strControlID)
                    '???????
                    LiLunPointName = SSProcess.GetObjectAttr(strControlID(i2), "SSObj_PointName")
                    LiLunJianZWMC = SSProcess.GetObjectAttr(strControlID(i2), "[JianZWMC]")
                    If LiLunJianZWMC = JianZWMC Then
                        LiLunPointName1 = LiLunPointName
                    End If
                Next
            End If
            
            '?????????????
            ids = SSProcess.SearchNearObjIDs(x0, y0, 0.1, 0, "9130411", objID )
            If ids <> "" Then
                strControlID = Split(ids,",")
                For i2 = 0 To UBound(strControlID)
                    '???????
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
    '??????
    g_docObj.CloneTableRow tableIndex,  iniRow, 1,CopyCount - 1, False
    '??????
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
    '????????
    MergeColValue tableIndex, cellCount, 1, 0
    g_docObj.DeleteRow tableIndex,cellCount + 1,False
End Function

'??????????
Function OutFYCheckTable(ByVal tableIndex)
    g_docObj.MoveToTable tableIndex,False
    ReDim CellList(CellCount)
    CellCount = 0
    iniRow = 1
    '??????
    geocount = GetFeatureCount("9130511", geocount)
    For i = 0 To geocount - 1
        '?????????????????
        objID = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        PointName = SSProcess.GetSelGeoValue(i, "SSObj_PointName")
        SSProcess.GetObjectPoint objID, 0, x1, y1, z1, pointtype, name
        x1 = Round(x1,3)
        y1 = Round(y1,3)
        z1 = Round(z1,3)
        '????????????????
        ids = SSProcess.SearchNearObjIDs(x1, y1, 0.1, 0, "9130411", objID )
        If ids <> "" Then
            strControlID = Split(ids,",")
            For i1 = 0 To UBound(strControlID)
                LiLunPointName = SSProcess.GetObjectAttr(strControlID(i1), "SSObj_PointName")
                SSProcess.GetObjectPoint strControlID(i1), 0, x0, y0, z0, pointtype, name
                x0 = Round(x0,3)
                y0 = Round(y0,3)
                z0 = Round(z0,3)
                '???????????????????????
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
    '??????
    copyCount = geocount - 1
    g_docObj.CloneTableRow tableIndex,  iniRow, 1,copyCount, False
    '??????
    startRow = 1
    strPointChange = ""
    For i = 0 To CellCount - 1
        CellValueList = Split(CellList(i),"||")
        startCol = 0
        '?????¦Ë????????
        strPointChange = GetString( cellValueList(5), "," , strPointChange)
        '??????
        For j = 0 To UBound(CellValueList)
            g_docObj.SetCellText tableIndex,startRow,startCol,cellValueList(j),True,False
            startCol = startCol + 1
        Next
        startRow = startRow + 1
    Next
    '?????¦Ë???????
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
    g_docObj.Replace "{??????????¦Ë???}",strMaxChange,0
    If strMaxChange < 5.0 Then g_docObj.Replace "{??????ÿZ???}","§³??5.0cm??????",0  Else   g_docObj.Replace "{??????ÿZ???}","????5.0cm????????",0
End Function


'????????????
Function OutControlCountTable(ByVal tableIndex)
    g_docObj.MoveToTable tableIndex,False
    ReDim CellList(CellCount)
    CellCount = 0
    iniRow = 1
    strPointName = ""
    xhCount = 1
    '??????
    geocount = GetFeatureCount("1130211", geocount)
    For i = 0 To geocount - 1
        objID = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        PointName = SSProcess.GetSelGeoValue(i, "SSObj_PointName")
        '???????       
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
            '????????????????
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
            '????????????
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
    '??????
    copyCount = (UBound(strPointNameList) + 1) * 2 - 1
    g_docObj.CloneTableRow tableIndex,  iniRow, 1,copyCount, False
    '??????
    startRow = 1
    For i = 0 To CellCount - 1
        CellValueList = Split(CellList(i),"||")
        startCol = 0
        '??????
        For j = 0 To UBound(CellValueList)
            g_docObj.SetCellText tableIndex,startRow,startCol,cellValueList(j),True,False
            startCol = startCol + 1
        Next
        startRow = startRow + 1
    Next
    '????????
    MergeColValue tableIndex, cellCount, 1, 0
    MergeColValue tableIndex, cellCount, 1, 1
    MergeColValue tableIndex, cellCount, 1, 4
    g_docObj.DeleteRow tableIndex,cellCount + 1,False
End Function

'?????????
Function OutControlResultTable(ByVal tableIndex)
    g_docObj.MoveToTable tableIndex,False
    ReDim CellList(CellCount)
    CellCount = 0
    iniRow = 1
    '???????????
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
    '??????
    copyCount = geocount - 1
    g_docObj.CloneTableRow tableIndex,  iniRow, 1,copyCount, False
    '??????
    startRow = 1
    For i = 0 To CellCount - 1
        CellValueList = Split(CellList(i),"||")
        startCol = 0
        '??????
        For j = 0 To UBound(CellValueList)
            g_docObj.SetCellText tableIndex,startRow,startCol,cellValueList(j),True,False
            startCol = startCol + 1
        Next
        startRow = startRow + 1
    Next
End Function

'???????????
Function OutControlLengthTable(ByVal tableIndex)
    g_docObj.MoveToTable tableIndex,False
    ReDim CellList(CellCount)
    CellCount = 0
    iniRow = 1
    xhCount = 1
    '??????
    strJCtable = "??????????????"
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
    '??????
    copyCount = JClistCount - 1
    g_docObj.CloneTableRow tableIndex,  iniRow, 1,copyCount, False
    '??????
    startRow = 1
    For i = 0 To CellCount - 1
        CellValueList = Split(CellList(i),"||")
        startCol = 0
        '??????
        For j = 0 To UBound(CellValueList)
            g_docObj.SetCellText tableIndex,startRow,startCol,cellValueList(j),True,False
            startCol = startCol + 1
        Next
        startRow = startRow + 1
    Next
End Function

'??GPS????
Function OutPara()
    ControlCode = "9130212,9130211,1102021,1103021"
    ControlCount = GetFeatureCount( ControlCode, geocount)
    strControlPointName = ""
    For i = 0 To ControlCount - 1
        objID = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        strPointName = SSProcess.GetSelGeoValue(i, "SSObj_PointName")
        strControlPointName = GetString(strPointName, "," , strControlPointName)
    Next
    g_docObj.Replace "{???????????}",strControlPointName,0
    
    
    strJCtable = "??????????????"
    strJCField = "CeZDH,JianCDH,YZBC,JCBC,BCJC"
    JClistCount = GetProjectTableList (strJCtable,"max(BCJC)",strJCtable & ".ID>0 and CeZDH<>'*' ","SpatialData","1",JClist,fieldCount)
    If JClistCount = 1 Then strMaxChange = Round(JClist(0,0),3)
    If strMaxChange <> "" Then strMaxChange = CDbl(strMaxChange * 100)
    g_docObj.Replace "{????????}",strMaxChange,0
    If strMaxChange < 5.0 Then g_docObj.Replace "{????ÿZ???}","§³??5.0cm??????",0 Else g_docObj.Replace "{????ÿZ???}","????5.0cm????????",0
End Function


'*****************************??????????????*******************************
' ????????????????????????????????????????????
Function GetFeatureCount(ByVal Code,ByRef geocount)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code","==",Code
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    GetFeatureCount = geocount
End Function

'???????????
Function GetString(ByVal value,ByVal splitMark , str)
    If str = "" Then
        str = value
    Else
        str = str & splitMark & value
    End If
    GetString = str
End Function' Name

'?????
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

'***********************???????????????**********************************
'GPS???????
Function GetValueGPSList(CellList,CellCount,ByVal strPointName,ByVal strType,ByVal strTypeValue,ByVal strSCValue,ByVal strChange,ByVal strBZ)
    cellValue = ""
    value = strPointName & "||" & strType & "||" & strTypeValue & "||" & strSCValue & "||" & strChange & "||" & strBZ
    cellValue = value
    ReDim Preserve CellList(CellCount)
    CellList(CellCount) = cellValue
    CellCount = CellCount + 1
End Function

'????????
Function GetValueCZList(CellList,CellCount,ByVal strtitle,ByVal strPointType,ByVal strDH,ByVal strFXZ,ByVal strSPJL,ByVal strX,ByVal strY)
    If strPointType = "?????" Then
        CellValue = strPointType & "||" & strDH & "||" & "" & "||" & "?????" & "||" & strFXZ
        ReDim Preserve CellList(CellCount)
        CellList(CellCount) = CellValue
        CellCount = CellCount + 1
    ElseIf strPointType = "?????" Then
        CellValue = strDH & "||" & strFXZ & "||" & strSPJL
        ReDim Preserve CellList(CellCount)
        CellList(CellCount) = CellValue
        CellCount = CellCount + 1
    ElseIf strPointType = "????" Then
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


'***********************************************************????????????***********************************************************
'//strTableName ??
'//strFields ???
'//strAddCondition ???? 
'//strTableType "AttributeData??????????? ,SpatialData?????????????" 
'//strGeoType ???????? ??????èå???(0??1???2?—¨3???)
'//rs ????????????rs(??,??)
'//fieldCount ??¦È???
'//????? ??sql????????????
Function GetProjectTableList(ByVal strTableName,ByVal strFields,ByVal strAddCondition,ByVal strTableType,ByVal strGeoType,ByRef rs(),ByRef fieldCount)
    GetProjectTableList = 0
    values = ""
    rsCount = 0
    fieldCount = 0
    If strTableName = "" Or strFields = "" Then Exit Function
    '???????????
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
    '??????????edb?????
    AccessName = SSProcess.GetProjectFileName
    '?§Ø????????
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
            value = Replace(value,",","??")
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
        MsgBox "??????????????????",48
    End If
    iRecordCount =  - 1
    'SQL???
    sql = StrSqlStatement
    '??????
    SSProcess.OpenAccessRecordset mdbName, sql
    '??????????
    RecordCount = SSProcess.GetAccessRecordCount (mdbName, sql)
    If RecordCount > 0 Then
        iRecordCount = 0
        ReDim arSQLRecord(RecordCount)
        '??????¦Á?????????
        SSProcess.AccessMoveFirst mdbName, sql
        iRecordCount = 0
        '??????
        While SSProcess.AccessIsEOF (mdbName, sql) = 0
            fields = ""
            values = ""
            '?????????????
            SSProcess.GetAccessRecord mdbName, sql, fields, values
            arSQLRecord(iRecordCount) = values                                        '??????
            iRecordCount = iRecordCount + 1                                                    '????????
            '???????¦Á?
            SSProcess.AccessMoveNext mdbName, sql
        WEnd
    End If
    '???????
    SSProcess.CloseAccessRecordset mdbName, sql
End Function

'//????
Dim  adoConnection
Function InitDB()
    accessName = SSProcess.GetProjectFileName
    Set adoConnection = CreateObject("adodb.connection")
    strcon = "DBQ=" & accessName & ";DRIVER={Microsoft Access Driver (*.mdb)};"
    adoConnection.Open strcon
End Function

'//???
Function ReleaseDB()
    adoConnection.Close
    Set adoConnection = Nothing
End Function

'??¡¤??
'//????????¡¤??
Function  GetFilePath
    projectFileName = SSProcess.GetSysPathName (5)
    filePath = Replace(projectFileName,".edb","")
    filePath = filePath & "\"
    CreateFolder filePath
    GetFilePath = filePath
End Function

'//??ø’??????
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