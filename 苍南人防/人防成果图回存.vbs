strfields = "FHDYBH,TCWS,FJDCS"
Sub OnClick()
    chewei
    strFHDYValues = ""
    geocount = GetFeatureCount( "9530226", geocount)
    For i = 0 To geocount - 1
        id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        GetFieldValues id, strfields, strValues, fieldCount
        str = ""
        For i1 = 0 To fieldCount - 1
            str = GetString( strValues(i1), "," , str)
        Next
        strFHDYValues = GetString( str, "||" , strFHDYValues)
    Next
    SSProcess.CloseDatabase()
    
    mdbName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb mdbName
    
    strFHDYValuesList = Split(strFHDYValues,"||")
    strfieldsList = Split(strfields,",")
    For i = 0 To UBound(strFHDYValuesList)
        strValuesList = Split(strFHDYValuesList(i),",")
        For i1 = 1 To UBound(strValuesList)
            sql = "update ??????????????? set " & strfieldsList(i1) & " = " & strValuesList(i1) & " where " & strfieldsList(0) & " = '" & strValuesList(0) & "'"
            
            SSProcess.ExecuteAccessSql mdbName,sql
        Next
    Next
    SSProcess.CloseAccessMdb mdbName
    SSProcess.MapMethod "clearattrbuffer", "???????????????"
End Sub


'?????????
Function GetFeatureCount(ByVal Code,ByRef geocount)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code","==",Code
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    GetFeatureCount = geocount
End Function

'??????????????
Function GetFieldValues(ByVal id,ByVal fields,ByRef strValues(),ByRef fieldCount)
    ReDim strValues(fieldCount)
    fieldCount = 0
    fieldsList = Split(fields,",")
    For i = 0 To UBound(fieldsList)
        values = SSProcess.GetObjectAttr (id, "[" & fieldsList(i) & "]")
        ReDim Preserve strValues(fieldCount)
        strValues(fieldCount) = values
        fieldCount = fieldCount + 1
    Next
End Function

'???????????
Function GetString(ByVal value,ByVal splitMark , str)
    If str = "" Then
        str = value
    Else
        str = str & splitMark & value
    End If
    GetString = str
End Function

Function  chewei
    SSProcess.ExecuteSDLFunction "ssedit,objreset",0
    '?????????????????¦Ë????
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==","9461023,9461043"
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    For i = 0 To geocount - 1
        id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        ZSXS = SSProcess.GetSelGeoValue(i, "[ZheSXS]")
        ZSXS = CDbl(ZSXS)
        If ZSXS <> 0.0 Then CWSL = CInt(geocount * ZSXS)
        SSProcess.SetObjectAttr id, "[CheWGS]", CWSL
    Next
    
    '??????/???????¦Ë??????????????????
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "=","9530226"
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    FJDCWCount = CWSL
    JDCWCount = 0
    WXCWCount = 0
    'CW_?????????¦Ë????????
    For i = 0 To geocount - 1
        id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        ids = SSProcess.SearchInnerObjIDs(id, 2, "9461013,9461033,9461053", 0)
        If ids <> "" Then
            strList = Split(ids,",")
            For i1 = 0 To UBound(strList)
                code = SSProcess.GetObjectAttr (strList(i1), "SSObj_Code")
                If code = 9461013 Or code = 9461033 Then
                    CWLX = SSProcess.GetObjectAttr(strList(i1),"[CheWLX]")
                    ZSXS = CDbl(SSProcess.GetObjectAttr(strList(i1),"[ZSXS]"))
                    If CWLX = "?????¦Ë" Then
                        JDCWCount = JDCWCount + ZSXS
                    Else
                        JDCWCount = JDCWCount + 1
                    End If
                ElseIf code = 9461053 Then
                    WXCWCount = WXCWCount + 1
                End If
            Next
        End If
        SSProcess.SetObjectAttr id, "[TCWS]", CInt(JDCWCount) + CInt(WXCWCount * 0.7)
        SSProcess.SetObjectAttr id, "[FJDCS]", FJDCWCount
        RFCWSL = JDCWCount + WXCWCount + FJDCWCount
    Next
    
    
    '?????¦Ë????
    
    geocount = SSProcess.GetSelGeoCount()
    
    For i = 0 To geocount - 1
        FJDCWCount = 0
        JDCWCount = 0
        WXCWCount = 0
        id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        ids = SSProcess.SearchInnerObjIDs(id, 2, "9461023,9461043,9461013,9461033,9461053", 0)
        'msgbox geocount
        If ids <> "" Then
            strList = Split(ids,",")
            For i1 = 0 To UBound(strList)
                code = SSProcess.GetObjectAttr (strList(i1), "SSObj_Code")
                '???????¦Ë????
                If code = 9461023 Or code = 9461043 Then
                    cwgs = SSProcess.GetObjectAttr (strList(i1), "[CheWGS]")
                    If cwgs <> "" Then cwgs = CInt(cwgs)
                    FJDCWCount = FJDCWCount + cwgs
                    '??????¦Ë????
                ElseIf code = 9461013 Or code = 9461033 Then
                    'msgbox ""
                    JDCWCount = JDCWCount + 1
                ElseIf code = 9461053 Then
                    cwgs = SSProcess.GetObjectAttr (strList(i1), "[CheWGS]")
                    If cwgs <> "" Then cwgs = CInt(cwgs)
                    WXCWCount = WXCWCount + cwgs
                End If
            Next
        End If
        dxCWSL = JDCWCount + WXCWCount + FJDCWCount
        DXWXCWCount = WXCWCount
    Next
    HUAZHUJI  dxCWSL,RFCWSL,DXWXCWCount
    fileName = SSProcess.GetSysPathName(5)
    pathName = fileName & "???????????¦Ë????????.dwg"
    SZDWT  id,pathName
End Function



Function SZDWT(TKID,fileName)
    SSProcess.ClearDataXParameter
    SSProcess.SetDataXParameter "DataType", "1"
    SSProcess.SetDataXParameter "Version", "2004"
    SSProcess.SetDataXParameter "FeatureCodeTBName", "FeatureCodeTB_kcad"
    SSProcess.SetDataXParameter "SymbolScriptTBName", "SymbolScriptTB_cad"
    SSProcess.SetDataXParameter "NoteTemplateTBName", "NoteTemplateTB_cad"
    SSProcess.SetDataXParameter "ExportPathName",fileName
    SSProcess.SetDataXParameter "DataBoundMode", "0"'0(????????)?? 1(???????)?? 2(??????)?? 3(??????)??4(????????????)?? 5(???ID??????)?? 6(???????)
    SSProcess.SetDataXParameter "DataBoundID", TKID
    'SSProcess.SetDataXParameter "ZoomInOutDataBound", "0.0001"  '?????????¦¶?????????????¦Ë??????-0.0001??
    SSProcess.SetDataXParameter "ExportLayerCount", "0" '??????????????????0?????????????????
    SSProcess.SetDataXParameter "ZeroLineWidth", "0" '???AutoCAD???????0?????????§³????????????????????????0??
    SSProcess.SetDataXParameter "AcadColorMethod", "0" '???DWG?????¡Â?? 0 ?????????? 1??RGB??????
    SSProcess.SetDataXParameter "ColorUseStatus", "1"       '??????????0??????????Ú…??????????1?????????Ú…????????
    SSProcess.SetDataXParameter "ExplodeObjColorStatus", "0"      '??????????????????0?????????????Ú…??????? 1?????????????????
    SSProcess.SetDataXParameter "FontHeightScale", "0.8"
    SSProcess.SetDataXParameter "FontWidthScale", "0.8"
    'SSProcess.SetDataXParameter "FontWidthScale", "FontClass_1190001=0.8,FontClass_1190002=0.8,FontClass_1990001=0.8,FontClass_1990002=0.8,FontClass_1990011=0.8,FontClass_1990012=0.8,FontClass_1990013=0.8,FontClass_1990014=0.8,FontClass_1990015=0.8,FontClass_1990016=0.8,FontClass_1990017=0.8,FontClass_1990018=0.8,FontClass_1990019=0.8,FontClass_1990020=0.8,FontClass_1990021=0.8,FontClass_1990022=0.8,FontClass_1990023=0.8,FontClass_1990031=0.8,FontClass_1990032=0.8,FontClass_1990033=0.8,FontClass_1990034=0.8,FontClass_1990035=0.8,FontClass_1990036=0.8,FontClass_1990037=0.8,FontClass_1990038=0.8,FontClass_1990039=0.8,FontClass_1990040=0.8,FontClass_1990041=0.8,FontClass_1990042=0.8,FontClass_1990043=0.8,FontClass_2190001=0.8,FontClass_2190002=0.8,FontClass_2190003=0.8,FontClass_2190004=0.8,FontClass_2190005=0.8,FontClass_2190006=0.8,FontClass_2290001=0.8,FontClass_2390001=0.8,FontClass_2490001=0.8,FontClass_2590001=0.8,FontClass_2690001=0.8,FontClass_2790001=0.8,FontClass_2990001=0.8,FontClass_3190001=0.8,FontClass_3190002=0.8,FontClass_3190003=0.8,FontClass_3190004=0.8,FontClass_3190005=0.8,FontClass_3190006=0.8,FontClass_3190011=0.8,FontClass_3190012=0.8,FontClass_3190013=0.8,FontClass_3190014=0.8,FontClass_3190015=0.8,FontClass_3290001=0.8,FontClass_3390001=0.8,FontClass_3490001=0.8,FontClass_3590001=0.8,FontClass_3690001=0.8,FontClass_3790001=0.8,FontClass_3890001=0.8,FontClass_3990041=0.8,FontClass_3990042=0.8,FontClass_3990043=0.8,FontClass_3990044=0.8,FontClass_3990045=0.8,FontClass_3990046=0.8,FontClass_3990047=0.8,FontClass_3990048=0.8,FontClass_3990049=0.8,FontClass_3990050=0.8,FontClass_3990051=0.8,FontClass_3990052=0.8,FontClass_3990053=0.8,FontClass_3990054=0.8,FontClass_3990055=0.8,FontClass_3990056=0.8,FontClass_3990057=0.8,FontClass_3990058=0.8,FontClass_3990059=0.8,FontClass_3990060=0.8,FontClass_4190001=0.8,FontClass_4290001=0.8,FontClass_4290002=0.8,FontClass_4290003=0.8,FontClass_4390001=0.8,FontClass_4390002=0.8,FontClass_4390003=0.8,FontClass_4390004=0.8,FontClass_4490001=0.8,FontClass_4590001=0.8,FontClass_4590002=0.8,FontClass_4590003=0.8,FontClass_4690001=0.8,FontClass_4790001=0.8,FontClass_4890001=0.8,FontClass_4990001=0.8,FontClass_4990011=0.8,FontClass_4990012=0.8,FontClass_4990013=0.8,FontClass_4990014=0.8,FontClass_5190001=0.8,FontClass_5290001=0.8,FontClass_5390001=0.8,FontClass_5490001=0.8,FontClass_5990001=0.8,FontClass_6390001=0.8,FontClass_6490001=0.8,FontClass_6590001=0.8,FontClass_6690001=0.8,FontClass_6790001=0.8,FontClass_6790002=0.8,FontClass_6790003=0.8,FontClass_6790004=0.8,FontClass_6790005=0.8,FontClass_6990001=0.8,FontClass_7190001=0.8,FontClass_7290001=0.8,FontClass_7290002=0.8,FontClass_7390001=0.8,FontClass_7490001=0.8,FontClass_7490002=0.8,FontClass_7590001=0.8,FontClass_7590002=0.8,FontClass_7590003=0.8,FontClass_7690001=0.8,FontClass_7690002=0.8,FontClass_7990001=0.8,FontClass_7990002=0.8,FontClass_7990003=0.8,FontClass_7990004=0.8,FontClass_8190001=0.8,FontClass_8290001=0.8,FontClass_8390001=0.8,FontClass_8990001=0.8,FontClass_Z0001=0.7,FontClass_Z0002=1,FontClass_Z0003=1,FontClass_Z0004=1,FontClass_Z0005=1,FontClass_Z0006=0.8,FontClass_Z0007=0.8,FontClass_Z0008=0.8,FontClass_Z0009=1.33,FontClass_Z0010=1,FontClass_Z0011=1,FontClass_Z0012=1,FontClass_Z0013=1,FontClass_Z0212=0.5,FontClass_Z0213=0.8" '???????????????FontClass_?????=?????,????????§Õ?????,??????????????,?§Ø????????,???????(?? FontClass_0=0.6,FontClass_1=0.7),?????????¦¶0-1??
    'SSProcess.SetDataXParameter "FontSizeUseStatus","0"               '?????§³????? 0 ?????????????????????????? 1 ????????????????????
    SSProcess.SetDataXParameter "OthersExportMode", "3"'???AutoCAD?????????????????? 0??????????? 1????????§Ö?????? 2????????§Ö????????3???¨®?0??
    SSProcess.SetDataXParameter "OthersExportToZFactor", "0"       '???AutoCAD?????????????????Z????????? 0??????????? 1???????
    SSProcess.SetDataXParameter "SymbolExplodeMode", "1"   '??????????? 0???????????? 1???????????Ú…??????? 2????????????
    SSProcess.SetDataXParameter "LayerUseStatus", "0"     '??????????????????0??????????Ú…???????????1?????????Ú…?????????
    SSProcess.SetDataXParameter "ExplodeObjLayerStatus", "1"  '??????????????????0?????????????Ú…??????? 1??????????????????
    SSProcess.SetDataXParameter "LineExportMode", "1" '???AutoCAD????????????????????? 0 ????????????????????3DPolyline?????????2DPolyline??????? 1??????2DPolyline??????? 2?? ????3DPolyline????? 3?? ????Polyline?????
    SSProcess.SetDataXParameter "LineWidthUseStatus", "0"  '??????????0??????????Ú…??????????1?????????Ú…????????
    SSProcess.SetDataXParameter "GotoPointsMode", "0"                     '???????????????? 0 ????????????? 1 ??????????????? 2 ????????????????
    SSProcess.SetDataXParameter "AcadLineWidthMode", "1" 'Acad???????????0 ????? 1 ???
    SSProcess.SetDataXParameter "AcadLineScaleMode", "1"                'Acad???????????????0 ??????????????? 1 ?????1???
    SSProcess.SetDataXParameter "AcadLineWeightMode","1"               'Acad????????????0 ??????? 1 ??? 2 ??? 3 ???????
    SSProcess.SetDataXParameter "AcadBlockUseColorMode", "1"        'Acad???????????¡Â????0 ??? 1 ??? 2 ????????
    SSProcess.SetDataXParameter "AcadLinetypeGenerateMode", "1" '???AutoCAD?????????????????????¨¢? 0??????? 1???????
    SSProcess.SetDataXParameter "ExplodeObjMakeGroup ", "0"       'AutoCAD????????????????????????? 0???????ï‚?? 1?? ???ï…?????FeatureCodeTB???§Ö?ExtraInfo=1 
    SSProcess.SetDataXParameter "AcadUsePersonalBlockScaleCodes ", "1=7601023"       'AcadUsePersonalBlockScaleCodes ??????????????????????1?? ????1=????1,????2;????2=????1,????2???2?? ???? (?¡Â????????§Ò????????????????) 
    dwt_path = SSProcess.GetSysPathName (0 ) & "\Acadlin\acad.dwt"
    SSProcess.SetDataXParameter "AcadDwtFileName", dwt_path
    
    startIndex = 0
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"DEFAULT"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????¦¶??"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"POI"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??¡¤??"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??¡¤??"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??¡¤??"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??¡¤????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??¡¤???????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??¨°??"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??¨°??"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??¨°??"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????¦¶"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"GPS????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????????¦¶"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?œý????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?œý??????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?œý??????¦¶??"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????¦¶??"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????????¦¶??"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?œý¦¶???"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????¦¶??"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?œý???????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"¦Ë??????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????????????????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????????????????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????????¦Ë"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????????¦Ë"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???¦Ë???????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????¦¶??"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"TERP"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"GTFA"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"GTFL"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"?????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"????????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???¦Ë??¦¶??"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"KZ_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"KZ_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"SX_G"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"SX_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"SX_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"SX_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"JMD_G"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"JMD_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"JMD_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"JMD_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"JT_G"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"JT_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"JT_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"JT_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"GX_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"GX_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"GX_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"JJ_G"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"JJ_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"JJ_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"JJ_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"DM_G"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"DM_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"DM_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"DM_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ZB_G"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ZB_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ZB_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ZB_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"QT"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"TK"
    
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"??????????"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"???????"
    
    
    
    startIndex = 0
    
    SSProcess.SetDataXParameter "LayerRelationCount","2000"
    'SSProcess.SetDataXParameter "LayerRelation"&CStr(AddOne(startIndex)),"DEFAULT:0:0:0:0:0"
    SSProcess.ExportData
End Function

Function AddOne(ByRef startIndex)
    startIndex = startIndex + 1
    AddOne = startIndex
End Function


Function HUAZHUJI(DWCWSL,RFCWSL,WXCESL)
    '???????
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", "?????¦Ë????"
    SSProcess.SelectFilter
    SSProcess.DeleteSelectionObj()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 9530229
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    If geoCount > 0 Then
        For i = 0 To geoCount - 1
            TKID = SSProcess.GetSelGeoValue( i, "SSObj_ID" )
            SSProcess.GetObjectPoint TKID, 2, x0, y0, z, pointtype, name
            makeNote  x0 - 50, y0 - 20,RGB(255,0,0),437,437,"?????¦Ë???" & DWCWSL & "??????????????????¦Ë" & RFCWSL & "????","????"
            makeNote  x0 - 50, y0 - 24,RGB(255,0,0),437,437,"?????????????¦Ë" & WXCESL & "??????0.7????","????"
        Next
    End If
End Function

Function makeNote(x, y, color, width, height, fontString,ztmc)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontClass", "0"
    SSProcess.SetNewObjValue "SSObj_FontInterval", "80"
    SSProcess.SetNewObjValue "SSObj_FontString", fontString
    SSProcess.SetNewObjValue "SSObj_Color", color
    'SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_FontName", ztmc
    SSProcess.SetNewObjValue "SSObj_LayerName", "?????¦Ë????"
    SSProcess.SetNewObjValue "SSObj_FontInterval", "20"
    ' SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "1"
    SSProcess.SetNewObjValue "SSObj_FontWidth",width
    SSProcess.SetNewObjValue "SSObj_FontHeight", height
    SSProcess.AddNewObjPoint x, y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function