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
            sql = "update 人防防护单元属性表 set " & strfieldsList(i1) & " = " & strValuesList(i1) & " where " & strfieldsList(0) & " = '" & strValuesList(0) & "'"
            
            SSProcess.ExecuteAccessSql mdbName,sql
        Next
    Next
    SSProcess.CloseAccessMdb mdbName
    SSProcess.MapMethod "clearattrbuffer", "人防防护单元属性表"
End Sub


'获取要素个数
Function GetFeatureCount(ByVal Code,ByRef geocount)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code","==",Code
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    GetFeatureCount = geocount
End Function

'获取字段属性数组
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

'整理出字符串
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
    '计算有系数的非机动车位个数
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
    
    '将机动/非机动车位个数汇总到防护单元上
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "=","9530226"
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    FJDCWCount = CWSL
    JDCWCount = 0
    WXCWCount = 0
    'CW_机动车停车位信息属性表
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
                    If CWLX = "大型车位" Then
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
    
    
    '图廓车位数量
    
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
                '非机动车位个数
                If code = 9461023 Or code = 9461043 Then
                    cwgs = SSProcess.GetObjectAttr (strList(i1), "[CheWGS]")
                    If cwgs <> "" Then cwgs = CInt(cwgs)
                    FJDCWCount = FJDCWCount + cwgs
                    '机动车位个数
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
    pathName = fileName & "地下人防区车位测量成果图.dwg"
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
    SSProcess.SetDataXParameter "DataBoundMode", "0"'0(所有数据)， 1(选择集数据)， 2(当前图幅)， 3(缓冲区)，4(指定编码闭合地物)， 5(指定ID闭合地物)， 6(所有图幅)
    SSProcess.SetDataXParameter "DataBoundID", TKID
    'SSProcess.SetDataXParameter "ZoomInOutDataBound", "0.0001"  '数据输出范围缩放量，米为单位，缺省值为-0.0001。
    SSProcess.SetDataXParameter "ExportLayerCount", "0" '输出图层个数。如果等于0，则只输出当前打开的图层。
    SSProcess.SetDataXParameter "ZeroLineWidth", "0" '输出AutoCAD数据时，0线宽分界值，小于或等于该值的线宽，输出时均设为0。
    SSProcess.SetDataXParameter "AcadColorMethod", "0" '输出DWG颜色使用方式 0 （颜色号）、 1（RGB颜色值）
    SSProcess.SetDataXParameter "ColorUseStatus", "1"       '颜色使用状态。0（按编码表设定颜色输出）、1（按地物设定颜色输出）
    SSProcess.SetDataXParameter "ExplodeObjColorStatus", "0"      '内嵌符号颜色输出方式。0（按符号描述设定输出）、 1（与主地物同色输出）
    SSProcess.SetDataXParameter "FontHeightScale", "0.8"
    SSProcess.SetDataXParameter "FontWidthScale", "0.8"
    'SSProcess.SetDataXParameter "FontWidthScale", "FontClass_1190001=0.8,FontClass_1190002=0.8,FontClass_1990001=0.8,FontClass_1990002=0.8,FontClass_1990011=0.8,FontClass_1990012=0.8,FontClass_1990013=0.8,FontClass_1990014=0.8,FontClass_1990015=0.8,FontClass_1990016=0.8,FontClass_1990017=0.8,FontClass_1990018=0.8,FontClass_1990019=0.8,FontClass_1990020=0.8,FontClass_1990021=0.8,FontClass_1990022=0.8,FontClass_1990023=0.8,FontClass_1990031=0.8,FontClass_1990032=0.8,FontClass_1990033=0.8,FontClass_1990034=0.8,FontClass_1990035=0.8,FontClass_1990036=0.8,FontClass_1990037=0.8,FontClass_1990038=0.8,FontClass_1990039=0.8,FontClass_1990040=0.8,FontClass_1990041=0.8,FontClass_1990042=0.8,FontClass_1990043=0.8,FontClass_2190001=0.8,FontClass_2190002=0.8,FontClass_2190003=0.8,FontClass_2190004=0.8,FontClass_2190005=0.8,FontClass_2190006=0.8,FontClass_2290001=0.8,FontClass_2390001=0.8,FontClass_2490001=0.8,FontClass_2590001=0.8,FontClass_2690001=0.8,FontClass_2790001=0.8,FontClass_2990001=0.8,FontClass_3190001=0.8,FontClass_3190002=0.8,FontClass_3190003=0.8,FontClass_3190004=0.8,FontClass_3190005=0.8,FontClass_3190006=0.8,FontClass_3190011=0.8,FontClass_3190012=0.8,FontClass_3190013=0.8,FontClass_3190014=0.8,FontClass_3190015=0.8,FontClass_3290001=0.8,FontClass_3390001=0.8,FontClass_3490001=0.8,FontClass_3590001=0.8,FontClass_3690001=0.8,FontClass_3790001=0.8,FontClass_3890001=0.8,FontClass_3990041=0.8,FontClass_3990042=0.8,FontClass_3990043=0.8,FontClass_3990044=0.8,FontClass_3990045=0.8,FontClass_3990046=0.8,FontClass_3990047=0.8,FontClass_3990048=0.8,FontClass_3990049=0.8,FontClass_3990050=0.8,FontClass_3990051=0.8,FontClass_3990052=0.8,FontClass_3990053=0.8,FontClass_3990054=0.8,FontClass_3990055=0.8,FontClass_3990056=0.8,FontClass_3990057=0.8,FontClass_3990058=0.8,FontClass_3990059=0.8,FontClass_3990060=0.8,FontClass_4190001=0.8,FontClass_4290001=0.8,FontClass_4290002=0.8,FontClass_4290003=0.8,FontClass_4390001=0.8,FontClass_4390002=0.8,FontClass_4390003=0.8,FontClass_4390004=0.8,FontClass_4490001=0.8,FontClass_4590001=0.8,FontClass_4590002=0.8,FontClass_4590003=0.8,FontClass_4690001=0.8,FontClass_4790001=0.8,FontClass_4890001=0.8,FontClass_4990001=0.8,FontClass_4990011=0.8,FontClass_4990012=0.8,FontClass_4990013=0.8,FontClass_4990014=0.8,FontClass_5190001=0.8,FontClass_5290001=0.8,FontClass_5390001=0.8,FontClass_5490001=0.8,FontClass_5990001=0.8,FontClass_6390001=0.8,FontClass_6490001=0.8,FontClass_6590001=0.8,FontClass_6690001=0.8,FontClass_6790001=0.8,FontClass_6790002=0.8,FontClass_6790003=0.8,FontClass_6790004=0.8,FontClass_6790005=0.8,FontClass_6990001=0.8,FontClass_7190001=0.8,FontClass_7290001=0.8,FontClass_7290002=0.8,FontClass_7390001=0.8,FontClass_7490001=0.8,FontClass_7490002=0.8,FontClass_7590001=0.8,FontClass_7590002=0.8,FontClass_7590003=0.8,FontClass_7690001=0.8,FontClass_7690002=0.8,FontClass_7990001=0.8,FontClass_7990002=0.8,FontClass_7990003=0.8,FontClass_7990004=0.8,FontClass_8190001=0.8,FontClass_8290001=0.8,FontClass_8390001=0.8,FontClass_8990001=0.8,FontClass_Z0001=0.7,FontClass_Z0002=1,FontClass_Z0003=1,FontClass_Z0004=1,FontClass_Z0005=1,FontClass_Z0006=0.8,FontClass_Z0007=0.8,FontClass_Z0008=0.8,FontClass_Z0009=1.33,FontClass_Z0010=1,FontClass_Z0011=1,FontClass_Z0012=1,FontClass_Z0013=1,FontClass_Z0212=0.5,FontClass_Z0213=0.8" '输出注记字宽缩放比，FontClass_分类号=缩放比,如果直接填写缩放比,则默认为全局缩放比,有多个分类号时,用逗号分隔(如 FontClass_0=0.6,FontClass_1=0.7),缩放比取值范围0-1。
    'SSProcess.SetDataXParameter "FontSizeUseStatus","0"               '字体大小使用状态 0 （按注记分类表设置字高宽输出）、 1 （按注记设置字高宽输出）
    SSProcess.SetDataXParameter "OthersExportMode", "3"'输出AutoCAD数据时，厚度输出方式。 0（地物编码）、 1（编码表中的厚度）、 2（编码表中的别名）、3（置成0）
    SSProcess.SetDataXParameter "OthersExportToZFactor", "0"       '输出AutoCAD数据时，厚度输出到块Z比例方式。 0（不输出）、 1（输出）
    SSProcess.SetDataXParameter "SymbolExplodeMode", "1"   '符号打散方式。 0（自动打散）、 1（根据编码表设定打散）、 2（全部不打散）
    SSProcess.SetDataXParameter "LayerUseStatus", "0"     '数据输出层名使用状态。0（按编码表设定层名输出）、1（按地物设定层名输出）
    SSProcess.SetDataXParameter "ExplodeObjLayerStatus", "1"  '内嵌符号图层输出方式。0（按符号描述设定输出）、 1（与主地物同层输出）
    SSProcess.SetDataXParameter "LineExportMode", "1" '输出AutoCAD数据时，多义线输出方式， 0 （缺省方式，带不同高程时按3DPolyline输出，其余按2DPolyline输出）、 1（强制按2DPolyline输出）、 2（ 强制按3DPolyline输出） 3（ 强制按Polyline输出）
    SSProcess.SetDataXParameter "LineWidthUseStatus", "0"  '线宽使用状态。0（按编码表设定线宽输出）、1（按地物设定线宽输出）
    SSProcess.SetDataXParameter "GotoPointsMode", "0"                     '输出图形折线化方式。 0 （不折线化）、 1 （只折线化曲线）、 2 （所有图形折线化）
    SSProcess.SetDataXParameter "AcadLineWidthMode", "1" 'Acad线宽输出方式。0 不输出 1 输出
    SSProcess.SetDataXParameter "AcadLineScaleMode", "1"                'Acad线型比例输出方式。0 与比例尺成正比输出 1 总是按1输出
    SSProcess.SetDataXParameter "AcadLineWeightMode","1"               'Acad线重输出方式。0 地物线宽 1 随层 2 随块 3 随线定义
    SSProcess.SetDataXParameter "AcadBlockUseColorMode", "1"        'Acad图块输出颜色使用方式。0 随层 1 随块 2 随块内实体
    SSProcess.SetDataXParameter "AcadLinetypeGenerateMode", "1" '输出AutoCAD数据时，线型生成是否启用。 0（禁用） 1（启用）
    SSProcess.SetDataXParameter "ExplodeObjMakeGroup ", "0"       'AutoCAD数据时，打散对象编组输出方式。 0（不编组）、 1（ 编组，同时要求FeatureCodeTB表中的ExtraInfo=1 
    SSProcess.SetDataXParameter "AcadUsePersonalBlockScaleCodes ", "1=7601023"       'AcadUsePersonalBlockScaleCodes 指定使用特殊块比例的编码。格式1： 比例1=编码1,编码2;比例2=编码1,编码2格式2： 比例 (该方式指定所有编码均使用指定的块比例) 
    dwt_path = SSProcess.GetSysPathName (0 ) & "\Acadlin\acad.dwt"
    SSProcess.SetDataXParameter "AcadDwtFileName", dwt_path
    
    startIndex = 0
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"DEFAULT"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"叠加分析过渡面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"标注层"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"三维测图"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"图廓层"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"征地"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"乡镇属性点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"村属性点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"征地界址点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"境界"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"地类图斑"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"勘测图廓层"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"测量控制点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"求算控制点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"控制点检查线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"数学基础"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"图幅范围面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"水系网线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"水系面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"水系线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"水系点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"水系附属设施线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"水系附属设施点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"水系附属设施面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"海洋线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"海洋面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"海洋点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"门址"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"楼址"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"POI"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"居民地点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"居民地附属设施线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"居民地面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"居民地线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"房屋外轮廓面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"房产辅助线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"居民地附属设施点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"居民地附属设施面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"部件"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"铁路线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"交通附属设施面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"交通附属设施点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"交通附属设施线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"道路面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"道路线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"道路网线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"交通附属设施"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"道路交叉口面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"管线线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"管线点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"管线面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"境界线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"境界点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"境界面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"其他境界面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"其他境界线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"其他境界点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"等高线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"高程点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"地貌点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"地貌面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"地貌线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"植被与土质线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"植被与土质点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"植被与土质面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"勘测标注信息"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"三调"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"勘测村界图例"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"勘测图例"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"出图范围"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"征地注记"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"地类图斑图例"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"征地村注记"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"图廓"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"征地界址线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"行政区"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"地籍区"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"地籍子区"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"理论控制点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"所有权宗地"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"GPS检测点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"理论测站点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"实测测站点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"支点线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"使用权宗地"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"放验线用地范围"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"宗地界址点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"检查线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"宗地界址线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"方向线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"实测控制点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"宗海"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"宗海界址点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"宗海界址线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"理论放样点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"实测放样点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"正负零标高"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"宗地图廓层"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"勘测图廓"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"面积注记"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"构筑物"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"自然幢"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"点状定着物"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"线状定着物"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"面状定着物"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"楼层"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"面积块"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"户"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"房间"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"外墙线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"内墙线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"分户中墙线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"分间中墙线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"门线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"窗户线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"房产部件点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"房产部件线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"房产部件面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"房产图廓层"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"构件信息"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"规划建筑轴线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"偏差方向"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"规划控制线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"规划建筑物范围线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"建筑物范围线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"建筑物基底范围线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"建筑白膜"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"规划围墙线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"竣工标高信息"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"竣工标注信息"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"出图范围线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"规划测量成果图图廓"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"位移图图廓"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"建筑占地面积计算略图图廓"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"建筑高度及层高测量略图图廓"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"消防登高面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"机动车停车位"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"非机动车停车位"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"停车位分布图图廓"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"绿地范围线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"绿地竣工图图廓"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"问题信息"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"待更新区域"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"工作区域"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"更新区域"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"院落街区面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"TERP"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"GTFA"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"GTFL"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"测量控制线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"地方坐标原点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"地类图斑辅助线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"国家不一致图斑"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"不一致图斑"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"地方不一致图斑"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"权属区域"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"征地项目面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"宗地"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"放样建筑"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"建筑物外轮廓"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"立面图辅助线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"竣工平面图图廓"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"竣工对比图图廓"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"土地核验图图廓"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"停车位范围线"
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
    
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"人防防护单元"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"非人防区"
    
    
    
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
    '添加代码
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", "地下车位图图例"
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
            makeNote  x0 - 50, y0 - 20,RGB(255,0,0),437,437,"地下车位合计" & DWCWSL & "个，其中人防区域内车位" & RFCWSL & "个。","黑体"
            makeNote  x0 - 50, y0 - 24,RGB(255,0,0),437,437,"注：地下室微型车位" & WXCESL & "个，按0.7折算","黑体"
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
    SSProcess.SetNewObjValue "SSObj_LayerName", "地下车位图图例"
    SSProcess.SetNewObjValue "SSObj_FontInterval", "20"
    ' SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "1"
    SSProcess.SetNewObjValue "SSObj_FontWidth",width
    SSProcess.SetNewObjValue "SSObj_FontHeight", height
    SSProcess.AddNewObjPoint x, y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function