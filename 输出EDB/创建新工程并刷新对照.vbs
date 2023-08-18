
Set FileSystemObject = CreateObject("Scripting.FileSystemObject")

Sub OnClick()
    
    CurrentProPath = Mid(SSProcess.GetProjectFileName,1,Len(SSProcess.GetProjectFileName) - 4) & "副本" & ".edb"
    Set FormerFileObj = FileSystemObject.GetFile(SSProcess.GetProjectFileName)
    FormerFileObj.Copy CurrentProPath
    SSProcess.OpenDatabase   CurrentProPath
    ToChinese
    changeLayerCode
    Dim filenames(1000),filecount,strs(1000),count
    PrjPathName = SSProcess.GetSysPathName (2)
    propaths = ""
    GetAllFiles PrjPathName, "prj", filecount, filenames
    If filecount > 0 Then
        For i = 0 To filecount - 1
            If propaths = "" Then
                propaths = Replace(filenames(i),PrjPathName,"")
            Else
                propaths = propaths & "," & Replace(filenames(i),PrjPathName,"")
            End If
        Next
    End If
    SSProcess.ClearInputParameter
    SSProcess.AddInputParameter "选择投影文件","" ,0, propaths, "选择投影文件,没有不选。需要把投影文件放到eps环境下的comm目录下。"
    'SSProcess.AddInputParameter "普查区域名称" , "",0,"", ""
    ret = SSProcess.ShowInputParameterDlg ("参数设置")
    
    If ret = 0 Then
        Exit Sub
    End If
    If ret = 1 Then
        SSProcess.UpdateScriptDlgParameter 1
        'PCQYMC  = SSProcess.GetInputParameter ("普查区域名称")
    End If
    
    SXZQMC = SSProcess.ReadEpsIni("市行政区域名称", "LastAttr" ,"")
    XXZQMC = SSProcess.ReadEpsIni("县行政区域名称", "LastAttr" ,"")
    PCQYMC = SXZQMC & XXZQMC
    
    SSProcess.UpdateScriptDlgParameter 1
    SystemFileName = SSProcess.GetInputParameter("选择投影文件")
    
    'pathName = SSProcess.SelectPathName( )
    SSProcess.ExecuteSDLFunction "$SDL.SSProject.Display.RedrawExtend", 0
    
    mdbName = SSProcess.GetProjectFileName()
    
    strOutputPath = Replace(mdbName, ".edb", "\")
    
    fileName = "温标" & ".mdb"
    
    If FolderExist(strOutputPath) = False Then CreateFolder strOutputPath
    
    fileName = strOutputPath & "\" & fileName
    
    '输出GDB
    ExportGDB fileName,SystemFileName,PrjPathName
    
    MsgBox "OK"
    
End Sub

SSProcess.AccessIsEOF mdbName, sql
Function AddOne( ByRef startIndex )
    startIndex = startIndex + 1
    AddOne = startIndex
End Function

Function ToChinese()
    SqlStr = "Select 地下管线线属性表.ID,地下管线线属性表.FSFS From 地下管线线属性表 Inner Join GeoLineTB on 地下管线线属性表.ID = GeoLineTB.ID WHERE (GeoLineTB.Mark Mod 2)<>0"
    GetSQLRecordAll SqlStr,LineArr,LineCount
    For i = 0 To LineCount - 1
        SingleLineArr = Split(LineArr(i),",", - 1,1)
        If IsNumeric(SingleLineArr(1)) Then
            Select Case SingleLineArr(1)
                '属性对照：0,1,2,3,4,5,6,7,8,9,10,11,12
                '直埋,管埋,管块,管沟,架空,地面,上架,小通道,综合管廊（沟）,人防,井内连线,顶管,水下
                Case "0"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","直埋"
                Case "1"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","管埋"
                Case "2"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","管块"
                Case "3"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","管沟"
                Case "4"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","架空"
                Case "5"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","地面"
                Case "6"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","上架"
                Case "7"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","小通道"
                Case "8"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","综合管廊（沟）"
                Case "9"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","人防"
                Case "10"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","井内连线"
                Case "11"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","顶管"
                Case "12"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","水下"
            End Select
        End If
    Next 'i
    
    SqlStr = "Select 地下管线线属性表.ID,地下管线线属性表.SJYL From 地下管线线属性表 Inner Join GeoLineTB on 地下管线线属性表.ID = GeoLineTB.ID WHERE (GeoLineTB.Mark Mod 2)<>0"
    GetSQLRecordAll SqlStr,LineArr,LineCount
    For i = 0 To LineCount - 1
        SingleLineArr = Split(LineArr(i),",", - 1,1)
        If IsNumeric(SingleLineArr(1)) Then
            '属性对照：0,1,2,3,4,5,6,7,8
            '高压,高压A级,高压B级,次高压A级,次高压B级,中压,中压A级,中压B级,低压
            Select Case SingleLineArr(1)
                Case "0"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","高压"
                Case "1"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","高压A级"
                Case "2"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","高压B级"
                Case "3"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","次高压A级"
                Case "4"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","次高压B级"
                Case "5"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","中压"
                Case "6"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","中压A级"
                Case "7"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","中压B级"
                Case "8"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","低压"
            End Select
        End If
    Next 'i
    
    SqlStr = "Select 地下管线线属性表.ID,地下管线线属性表.GC From 地下管线线属性表 Inner Join GeoLineTB on 地下管线线属性表.ID = GeoLineTB.ID WHERE (GeoLineTB.Mark Mod 2)<>0"
    GetSQLRecordAll SqlStr,LineArr,LineCount
    For i = 0 To LineCount - 1
        SingleLineArr = Split(LineArr(i),",", - 1,1)
        If IsNumeric(SingleLineArr(1)) Then
            Select Case SingleLineArr(1)
                '属性对照：0,1,2,3,4,5,6,7,8,9,10,11,12
                '焊接钢管,无缝钢管,灰口铸铁管,球墨铸铁管,混凝土管,玻璃钢管,PVC,PE管,铜,光纤,钢筋铝绞线,砖石,其他
                Case "0"
                SSProcess.SetObjectAttr SingleLineArr(0),"[GC]","焊接钢管"
                Case "1"
                SSProcess.SetObjectAttr SingleLineArr(0),"[GC]","无缝钢管"
                Case "2"
                SSProcess.SetObjectAttr SingleLineArr(0),"[GC]","灰口铸铁管"
                Case "3"
                SSProcess.SetObjectAttr SingleLineArr(0),"[GC]","球墨铸铁管"
                Case "4"
                SSProcess.SetObjectAttr SingleLineArr(0),"[GC]","混凝土管"
                Case "5"
                SSProcess.SetObjectAttr SingleLineArr(0),"[GC]","玻璃钢管"
                Case "6"
                SSProcess.SetObjectAttr SingleLineArr(0),"[GC]","PVC"
                Case "7"
                SSProcess.SetObjectAttr SingleLineArr(0),"[GC]","PE管"
                Case "8"
                SSProcess.SetObjectAttr SingleLineArr(0),"[GC]","铜"
                Case "9"
                SSProcess.SetObjectAttr SingleLineArr(0),"[GC]","光纤"
                Case "10"
                SSProcess.SetObjectAttr SingleLineArr(0),"[GC]","钢筋铝绞线"
                Case "11"
                SSProcess.SetObjectAttr SingleLineArr(0),"[GC]","砖石"
                Case "12"
                SSProcess.SetObjectAttr SingleLineArr(0),"[GC]","其他"
            End Select
        End If
    Next 'i
End Function' ToChinese

'获取所有记录
Function GetSQLRecordAll(ByVal StrSqlStatement, ByRef SQLRecord(), ByRef iRecordCount)
    SSProcess.OpenAccessMdb SSProcess.GetProjectFileName
    If StrSqlStatement = "" Then
        MsgBox "查询语句为空，操作停止！",48
    End If
    iRecordCount =  - 1
    SSProcess.OpenAccessRecordset SSProcess.GetProjectFileName, StrSqlStatement
    RecordCount = SSProcess.GetAccessRecordCount (SSProcess.GetProjectFileName, StrSqlStatement)
    If RecordCount > 0 Then
        iRecordCount = 0
        ReDim SQLRecord(RecordCount)
        SSProcess.AccessMoveFirst SSProcess.GetProjectFileName, StrSqlStatement
        iRecordCount = 0
        While SSProcess.AccessIsEOF (SSProcess.GetProjectFileName, StrSqlStatement) = 0
            fields = ""
            values = ""
            SSProcess.GetAccessRecord SSProcess.GetProjectFileName, StrSqlStatement, fields, values
            SQLRecord(iRecordCount) = values
            iRecordCount = iRecordCount + 1
            SSProcess.AccessMoveNext SSProcess.GetProjectFileName, StrSqlStatement
        WEnd
    End If
    SSProcess.CloseAccessRecordset SSProcess.GetProjectFileName, StrSqlStatement
    SSProcess.CloseAccessMdb SSProcess.GetProjectFileName
End Function


Function break
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Type", "==", "POINT,LINE,AREA"
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    For i = 0 To geocount - 1
        '写基本信息
        sid = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        '符号打散方式。 0（自动打散）、 1（根据编码表设定打散）
        SSProcess.ExplodeObj sid, 1, 1, "AfterExplodeObj"
    Next
    
End Function

Function AfterExplodeObj()
    '取打散地物ID
    SSParameter.GetParameterINT "AfterExplodeObj", "ExplodeObjID", "0", objID
    objType = SSProcess.GetObjectAttr (objID, "SSObj_Type" )
    geoCode = SSProcess.GetObjectAttr (objID, "SSObj_Code" )
    layername = SSProcess.GetObjectAttr (objID, "SSObj_LayerName" )
    If layername = "GXFSLN" Then
        DATASOURCE = SSProcess.GetObjectAttr (objID, "[DATASOURCE]" )
        GXRQ = SSProcess.GetObjectAttr (objID, "[GXRQ]" )
        CHDW = SSProcess.GetObjectAttr (objID, "[CHDW]" )
        DMGC = SSProcess.GetObjectAttr (objID, "[DMGC]" )
    End If
    geoCount = SSProcess.GetSelGeoCount
    For i = 0 To geoCount - 1
        geoID = SSProcess.GetSelGeoValue (i, "SSObj_ID" )
        SSProcess.SetObjectAttr geoID, "[DATASOURCE]", DATASOURCE
        SSProcess.SetObjectAttr geoID, "[GXRQ]", GXRQ
        SSProcess.SetObjectAttr geoID, "[CHDW]", CHDW
        SSProcess.SetObjectAttr geoID, "[DMGC]", DMGC
    Next
End Function

Function CreateFolder(ByVal strFolderPath)
    Dim FSO,OutProject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    strDir = Left(strFolderPath,3)
    strFolderPath = Right( strFolderPath, Len(strFolderPath) - 3 )
    mulFolderPath = Split(strFolderPath,"\")
    nCount = UBound(mulFolderPath)
    OutFolderPath = strDir
    For i = 0 To nCount - 1
        OutFolderPath = OutFolderPath & "\" & mulFolderPath(i)
        If FSO.FolderExists(OutFolderPath) = False Then
            Set OutProject = FSO.CreateFolder(OutFolderPath)
            Set OutProject = Nothing
        End If
    Next
    Set FSO = Nothing
End Function




Function ExportGDB(fileName,SystemFileName,PrjPathName)
    
    '-----------清空转换参数--------------
    SSProcess.ClearDataXParameter
    '-----------设置基本转换参数------------------
    SSProcess.SetDataXParameter "DataType", "22"
    SSProcess.SetDataXParameter "FeatureCodeTBName", "FeatureCodeTB_500"
    SSProcess.SetDataXParameter "SymbolScriptTBName", "SymbolScriptTB_500"
    SSProcess.SetDataXParameter "NoteTemplateTBName", "NoteTemplateTB_500"
    SSProcess.SetDataXParameter "ExportPathName", fileName
    SSProcess.SetDataXParameter "DataBoundMode", "0"
    SSProcess.SetDataXParameter "SymbolExplodeMode", "1"
    SSProcess.SetDataXParameter "AddSystemFieldMode", "0"
    SSProcess.SetDataXParameter "LayerUseStatus", "0"
    SSProcess.SetDataXParameter "EXCHANGE_PDB_ExportNoteMode", "0"
    SSProcess.SetDataXParameter "EXCHANGE_PDB_ExportEmptyLayer", "1"
    'SSProcess.SetDataXParameter "EXCHANGE_PDB_ExportShapeMode", "1"
    SSProcess.SetDataXParameter "EXCHANGE_PDB_ExportShapeMode", "1"
    SSProcess.SetDataXParameter "FormatAttrValueStatus","1"
    'SSProcess.SetDataXParameter"EXCHANGE_PDB_ExportShortDate","1"
    SSProcess.SetDataXParameter "EXCHANGE_PDB_PrjFile", PrjPathName & SystemFileName
    
    SSProcess.SetDataXParameter"EXCHANGE_PDB_SpatialRF_MinX","190672.27475"
    SSProcess.SetDataXParameter"EXCHANGE_PDB_SpatialRF_MinY","1001661.35705"
    SSProcess.SetDataXParameter"EXCHANGE_PDB_SpatialRF_MaxX","808944.90075"
    SSProcess.SetDataXParameter"EXCHANGE_PDB_SpatialRF_MaxY","6219933.98305"
    
    startIndex = 0
    SSProcess.SetDataXParameter "ExportLayerCount", "100"
    ' SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "GXSSPT"
    ' SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "GXSSLN"
    ' SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "GXSSPY"
    ' SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "GXYJPT"
    ' SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "GXFSLN"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "JS"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "ZS"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "PS"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "YS"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "WS"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "DL"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "GD"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "LD"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "DC"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "XH"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "TX"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "DX"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "YD"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "LT"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "JX"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "JK"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "EX"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "DS"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "BZ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "RQ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "MQ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "TR"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "YH"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "RL"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "RS"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "ZQ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "BM"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "CD"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "CP"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "CT"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "CH"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "CY"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "CQ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "CS"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "QT"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "FQ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "XF"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "FS"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "SY"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "GS"
    
    
    
    startIndex = 0
    SSProcess.SetDataXParameter "LayerRelationCount", "100"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JS:JSPOINT,给水管线（点）:JSLINE,给水管线:,::"
    ' SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JY:POINT,（点）:LINE,（线）:,::"
    ' SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "ZS:POINT,（点）:LINE,（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "PS:PSPOINT,排水管线（点）:PSLINE,排水管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YS:YSPOINT,雨水管线（点）:YSLINE,雨水管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "WS:WSPOINT,污水管线（点）:WSLINE,污水管线（线）:,::"
    ' SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "HS:HSPOINT,（点）:HSLINE,（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DL:DLPOINT,电力管线（点）:DLLINE,电力管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "GD:GDPOINT,供电管线（点）:GDLINE,供电管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "LD:LDPOINT,路灯管线（点）:LDLINE,路灯管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DC:DCPOINT,电车管线（点）:DCLINE,电车管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "XH:XHPOINT,交通信号管线（点）:XHLINE,交通信号管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "TX:TXPOINT,综合管线（点）:TXLINE,综合管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DX:DXPOINT,电信管线（点）:DXLINE,电信管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YD:YDPOINT,移动管线（点）:YDLINE,移动管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "LT:LTPOINT,联通管线（点）:LTLINE,联通管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JX:JXPOINT,军用管线（点）:JXLINE,军用管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JK:JKPOINT,监控管线（点）:JKLINE,监控管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "EX:EXPOINT,电力通讯管线（点）:EXLINE,电力通讯管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DS:DSPOINT,广播电视管线（点）:DSLINE,广播电视管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "BZ:BZPOINT,保密专用管线（点）:BZLINE,电保密专用管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RQ:RQPOINT,燃气管线（点）:RQLINE,燃气管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "MQ:MQPOINT,煤气管线（点）:MQLINE,煤气管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "TR:TRPOINT,天然气管线（点）:TRLINE,天然气管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YH:YHPOINT,液化气管线（点）:YHLINE,液化气管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RL:RLPOINT,热力管线（点）:RLLINE,热力管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RS:RSPOINT,热水管线（点）:RSLINE,热水管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "ZQ:ZQPOINT,蒸汽管线（点）:ZQLINE,蒸汽管线（线）:,::"
    ' SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "GY:POINT,（点）:LINE,（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "BM:BMPOINT,不明管线（点）:BMLINE,不明管线（线）:,::"
    ' SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "ZH:POINT,（点）:LINE,（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CD:CDPOINT,长输输电管线（点）:CDLINE,长输输电管线（线）:,::"
    ' SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CP:POINT,（点）:LINE,（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CT:CTPOINT,长输通信（点）:CTLINE,长输通信（线）:,::"
    ' SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CH:POINT,（点）:LINE,（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CY:CYPOINT,油主管道管线（点）:CYLINE,油主管道管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CQ:CQPOINT,天然气主管道管线（点）:CQLINE,天然气主管道管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CS:CSPOINT,水主管道管线（点）:CSLINE,水主管道管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "QT:QTPOINT,其他主管道管线（点）:QTLINE,其他主管道管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "FQ:FQPOINT,废弃管线（点）:FQLINE,废弃管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "XF:XFPOINT,消防水管线（点）:XFLINE,消防水管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "FS:FSPOINT,生活废水管线（点）:FSLINE,生活废水管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "SY:SYPOINT,石油管线（点）:SYLINE,石油管线（线）:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "GS:GSPOINT,工业废水管线（点）:GSLINE,工业废水管线（线）:,::"
    
    startIndex = 0
    SSProcess.SetDataXParameter "TableFieldDefCount", "100000"
    
    
    
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,FCODE,FCODE,要素分类代码,byname,,dbText:1,10,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,PCDYBH,PCDYBH,普查单元编号,,,dbText:1,11,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,FEATUREID,FEATUREID,设施编码,,,dbText,14,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,SSWZ,SSWZ,设施位置,,,dbText,128,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,GXDDH,GXDDH,管线点点号,,,dbText:1,17,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,TZ,TZ,特征,,,dbText,20,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,FSW,FSW,附属物,,,dbText,16,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,DMGC,DMGC,地面高程,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,ORIENTATION,ORIENTATION,方向,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,PXJW,PXJW,偏心井位,,,dbText,20,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,DATASOURCE,DATASOURCE,数据源,,,dbText:1,30,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,CHDW,CHDW,测绘单位,,,dbText:1,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,GXRQ,GXRQ,更新日期,,,dbText:1,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,BZ,BZ,备注,,,dbText,255,0"
    
    
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,FCODE,FCODE,要素分类代码,byname,,dbText:1,10,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,PCDYBH,PCDYBH,普查单元编号,,,dbText:1,11,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,FEATUREID,FEATUREID,设施编码,,,dbText:1,14,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SSMC,SSMC,设施名称,,,dbText,64,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SSWZ,SSWZ,设施位置,,,dbText,128,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,ZFZGBM,ZFZGBM,政府主管部门,,,dbText:1,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,YGDW,YGDW,运管单位,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,QSDW,QSDW,权属单位,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,TXJYDW,TXJYDW,特许经营单位,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,JSDW,JSDW,建设单位,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SJDW,SJDW,设计单位,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,KCDW,KCDW,勘察单位,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SGDW,SGDW,施工单位,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,JCNY,JCNY,建成年月,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,KSSYNY,KSSYNY,开始使用年月,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SJBCSJ,SJBCSJ,设计报出时间,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GXDBH,GXDBH,管线段编号,,,dbText:1,35,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GXQDDH,GXQDDH,管线起点点号,,,dbText:1,17,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GXZDDH,GXZDDH,管线终点点号,,,dbText:1,17,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GXQDMS,GXQDMS,管线起点埋深,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GXZDMS,GXZDMS,管线终点埋深,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GXQDDMGC,GXQDDMGC,管线起点地面高程,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GXZDDMGC,GXZDDMGC,管线终点地面高程,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GXQDGDGC,GXQDGDGC,管线起点管道高程,,,dbDouble:1,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GXZDGDGC,GXZDGDGC,管线终点管道高程,,,dbDouble:1,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GC,GC,管材,,,dbText:1,16,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SSJZ,SSJZ,输送介质,,,dbText,16,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SJWD,SJWD,设计温度,,,dbText,16,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GDBWCL,GDBWCL,管道保温材料,,,dbText,32,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GBHD,GBHD,管壁厚度,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GWYL,GWYL,管网压力,,,dbinteger,1,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SJYL,SJYL,设计压力,,,dbText,16,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,DYZ,DYZ,电压值,,,dbText,16,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,LL,LL,流量,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,LX,LX,流向,,,dbText,16,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,FSFS,FSFS,敷设方式,,,dbText:1,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GJ,GJ,管径（DN）,,,dbText,16,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,DMCC,DMCC,断面尺寸,,,dbText,16,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GDGCWZ,GDGCWZ,管道高程位置,,,dbinteger:1,1,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,XLTS,XLTS,线缆条数,,,dbLong,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,ZKS,ZKS,总孔数,,,dbLong,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,YYKS,YYKS,已用孔数,,,dbLong,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,KJ,KJ,孔径,,,dbText,16,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GXZT,GXZT,管线状态,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GDJKXS,GDJKXS,管道接口形式,,,dbText,16,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GXAFXS,GXAFXS,管线安放形式,,,dbText,16,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SFMYSZX,SFMYSZX,是否埋有示踪线,,,dbInteger,1,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,DJQK,DJQK,地基情况,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,JCXS,JCXS,基础形式,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SJDXSW,SJDXSW,设计地下水位,,,dbText,20,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,XKDXSW,XKDXSW,现况地下水位,,,dbText,20,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,DXSSFYFSX,DXSSFYFSX,地下水是否有腐蚀性,,,dbInteger,1,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SGFS,SGFS,施工方式,,,dbText,32,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SJSYNX,SJSYNX,设计使用年限,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,JGSJAQDJ,JGSJAQDJ,结构设计安全等级,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,KZSFLD,KZSFLD,抗震设防烈度,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,KZSFLB,KZSFLB,抗震设防类别,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,DMHZSJBZ,DMHZSJBZ,地面活载设计标准,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SFCYDZDLD,SFCYDZDLD,是否处于地震断裂带,,,dbInteger,1,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SFCZBLDZ,SFCZBLDZ,是否存在不良地质,,,dbInteger,1,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SFCYQBSCZ,SFCYQBSCZ,是否处于浅部砂层中,,,dbInteger,1,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,MZGDWGJC,MZGDWGJC,明装管道外观检查,,,dbText,32,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,DATASOURCE,DATASOURCE,数据源,,,dbText:1,30,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GXRQ,GXRQ,更新日期,,,dbText:1,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,CHDW,CHDW,测绘单位,,,dbText:1,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,BZ,BZ,备注,,,dbText,255,0"
    ' 'SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,FJ,FJ,附件,,,dbText,255,0"
    
    
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,FCODE,FCODE,要素分类代码,byname,,dbText:1,10,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,PCDYBH,PCDYBH,普查单元编号,,,dbText:1,11,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,FEATUREID,FEATUREID,设施编码,,,dbText:1,14,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,SSWZ,SSWZ,设施位置,,,dbText,128,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,SSMC,SSMC,设施名称,,,dbText,64,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,ZFZGBM,ZFZGBM,政府主管部门,,,dbText:1,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,YGDW,YGDW,运管单位,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,QSDW,QSDW,权属单位,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,TXJYDW,TXJYDW,特许经营单位,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,JSDW,JSDW,建设单位,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,SJDW,SJDW,设计单位,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,KCDW,KCDW,勘察单位,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,SGDW,SGDW,施工单位,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,JCNY,JCNY,建成年月,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,KSSYNY,KSSYNY,开始使用年月,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,SJBCSJ,SJBCSJ,设计报出时间,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,SJSYNX,SJSYNX,设计使用年限,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,QDDMGC,QDDMGC,起点地面高程,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,ZDDMGC,ZDDMGC,终点地面高程,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,QDJGDBDMGC,QDJGDBDMGC,起点结构顶板顶面高程,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,ZDJGDBDMGC,ZDJGDBDMGC,终点结构顶板顶面高程,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,QDDBFTHD,QDDBFTHD,起点顶板覆土厚度,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,ZDDBFTHD,ZDDBFTHD,终点顶板覆土厚度,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,GLCSSL,GLCSSL,管廊舱室数量,,,dbText:1,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,GLNYXGXZL,GLNYXGXZL,管廊内运行管线种类,,,dbText,32,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,DJQK,DJQK,地基情况,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,JGXS,JGXS,结构形式,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,ZJFW,ZJFW,注浆范围,,,dbText,32,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,SJDXSW,SJDXSW,设计地下水位,,,dbText,20,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,XKDXSW,XKDXSW,现况地下水位,,,dbText,20,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,DXSSFYFSX,DXSSFYFSX,地下水是否有腐蚀性,,,dbInteger,1,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,SGFS,SGFS,施工方式,,,dbText,32,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,JGSJAQDJ,JGSJAQDJ,结构设计安全等级,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,KZSFLD,KZSFLD,抗震设防烈度,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,KZSFLB,KZSFLB,抗震设防类别,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,DMHZSJBZ,DMHZSJBZ,地面活载设计标准,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,SFCYDZDLD,SFCYDZDLD,是否处于地震断裂带,,,dbInteger,1,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,SFCZBLDZ,SFCZBLDZ,是否存在不良地质,,,dbInteger,1,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,SFCYQBSCZ,SFCYQBSCZ,是否处于浅部砂层中,,,dbInteger,1,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,DATASOURCE,DATASOURCE,数据源,,,dbText:1,30,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,GXRQ,GXRQ,更新日期,,,dbText:1,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,CHDW,CHDW,测绘单位,,,dbText:1,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,BZ,BZ,备注,,,dbText,255,0"
    ' 'SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,FJ,FJ,附件,,,dbText,255,0"
    
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,FCODE,FCODE,要素分类代码,byname,,dbText:1,10,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,PCDYBH,PCDYBH,普查单元编号,,,dbText:1,11,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,FEATUREID,FEATUREID,设施编码,,,dbText:1,14,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,SSWZ,SSWZ,设施位置,,,dbText,128,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,ZFZGBM,ZFZGBM,政府主管部门,,,dbText:1,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,YGDW,YGDW,运管单位,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,QSDW,QSDW,权属单位,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,TXJYDW,TXJYDW,特许经营单位,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,JSDW,JSDW,建设单位,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,SJDW,SJDW,设计单位,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,KCDW,KCDW,勘察单位,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,SGDW,SGDW,施工单位,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,GXDDH,GXDDH,管线点点号,,,dbText:1,17,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,TZ,TZ,特征,,,dbText,20,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,FSW,FSW,附属物,,,dbText,16,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,DMGC,DMGC,地面高程,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,ORIENTATION,ORIENTATION,方向,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,JGXZ,JGXZ,井盖形状,,,dbText:1,20,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,JGZJHDMCC,JGZJHDMCC,井盖直径或断面尺寸,,,dbText,20,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,JGCZ,JGCZ,井盖材质,,,dbText:1,20,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,JGXS,JGXS,结构形式,,,dbText,20,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,JBS,JBS,井脖深,,,dbDouble,12,2"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,JS,JS,井深,,,dbDouble,12,2"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,JBCC,JBCC,井脖直径或断面尺寸,,,dbText,20,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,JCC,JCC,井直径或断面尺寸,,,dbText,20,20"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,SFYAQW,SFYAQW,是否有安全网,,,dbInteger,1,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,DATASOURCE,DATASOURCE,数据源,,,dbText:1,30,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,GXRQ,GXRQ,更新日期,,,dbText:1,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,CHDW,CHDW,测绘单位,,,dbText:1,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,BZ,BZ,备注,,,dbText,255,0"
    
    
    
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXFSLN,1,FCODE,FCODE,要素分类代码,byname,,dbText:1,10,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXFSLN,1,PCDYBH,PCDYBH,普查单元编号,,,dbText:1,11,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXFSLN,1,DATASOURCE,DATASOURCE,数据源,,,dbText:1,30,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXFSLN,1,GXRQ,GXRQ,更新日期,,,dbText:1,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXFSLN,1,CHDW,CHDW,测绘单位,,,dbText:1,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXFSLN,1,BZ,BZ,备注,,,dbText,255,0"
    
    
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXFSLN,2,FCODE,FCODE,要素分类代码,byname,,dbText:1,10,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXFSLN,2,PCDYBH,PCDYBH,普查单元编号,,,dbText:1,11,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXFSLN,2,DATASOURCE,DATASOURCE,数据源,,,dbText:1,30,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXFSLN,2,GXRQ,GXRQ,更新日期,,,dbText:1,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXFSLN,2,CHDW,CHDW,测绘单位,,,dbText:1,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXFSLN,2,BZ,BZ,备注,,,dbText,255,0"
    
    
    
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,XMMC,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,WTDH,Exp_No,物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,WTTSDH,Map_NO,图上点号,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,DMGC,Surf_H,地面高程,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,TZ,Feature,特征,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,FSW,Subsid,附属物,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,PXJW,Offset,偏心井位,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,ORIENTATION,Angle,以X轴方向为0度,取逆时针方向,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,MapNumber,MapNum,图幅号,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,X,X,纵坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,Y,Y,横坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,MapNo_X,MapNo_X,位移后的图上点号位置X坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,MapNo_Y,MapNo_Y,位移后的图上点号位置y坐标,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,JCNY,Mdata,建设年代,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,QSDW,B_Code,权属单位代码中文字段,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,SSWZ,Road,路名代码中文字段,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,XMMC,PID,工程号,,,dbText,20,0"
    
    '参数顺序：层名,类型,EPS字段名,客户字段名,客户字段别名,系统字段名,缺省值,字段类型,字段长度,小数位
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,ZKS,Hole_Num,总孔数,,,dbLong,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,GXQDDH,S_point,起点物探点号,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,GXZDDH,E_point,连接方向,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,GXQDMS,S_Deep,起点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,GXZDMS,E_Deep,终点埋深,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,GC,Material,材质,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,DMCC,D_S,管径或管块,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,JCNY,Mdata,建设年代,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,XLTS,Cab_Count,电缆条数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,DYZ,Voltage,电压值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,ZKS,Hole_Num,总孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,LX,FlowDirect,排水流向,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,GXLX,P_Type,管线类型,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,BZ,Memo,备注,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,FSFS,D_Type,埋设类型,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,QSDW,B_Code,权属单位代码,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,SJYL,Pressure,压力值,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,WYKS,Hole_Used,未用孔数,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,D_Dia,D_Dia,套管尺寸,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,SSWZ,Road,路名代码,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,XMBH,PID,工程号,,,dbText,20,0"
    
    SSProcess.ExportData
    
End Function

Function ScanString(ByVal str, ByVal sep, ByRef strs(), ByRef count)
    Dim sepidx1, sepidx2, strtemp
    count = 0
    sepidx1 = 1
    sepidx2 = InStr(sepidx1 , str, sep, 1)
    While (sepidx2 > 0)
        strs(count) = Mid( str, sepidx1, sepidx2 - sepidx1)
        sepidx1 = sepidx2 + 1
        sepidx2 = InStr(sepidx1, str, sep, 1)
        count = count + 1
    WEnd
    strs(count) = Mid( str, sepidx1, Len(str) + 1 - sepidx1)
    count = count + 1
End Function


Function AngleFz
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Type", "==", "POINT"
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    If geocount > 0 Then
        For i = 0 To geocount - 1
            id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
            code = SSProcess.GetSelGeoValue(i, "SSObj_Code")
            FX = SSProcess.GetSelGeoValue(i, "[ORIENTATION]")
            angle = SSProcess.GetObjectAttr(id, "SSObj_Angle")
            angle = SSProcess.AdjustAngle(angle)
            Pi = 3.1415926
            angle1 = SSProcess.RadianToDms(angle)
            If FX = "" Then
                If angle1 < 90 And angle1 > 0 Then
                    angle = SSProcess.RadianToDms  (Pi / 2 - angle)
                    angle = Left(angle,InStr(angle,".") + 3)
                    SSProcess.SetObjectAttr id, "[ORIENTATION]", angle
                End If
                If angle1 > 90 Or angle1 = 90 Then
                    angle = SSProcess.RadianToDms  ((5 * Pi / 2) - angle)
                    angle = Left(angle,InStr(angle,".") + 3)
                    SSProcess.SetObjectAttr id, "[ORIENTATION]", angle
                End If
            ElseIf FX <> "" And angle <> FX Then
                If angle1 < 90 And angle1 > 0 Then
                    angle = SSProcess.RadianToDms  (Pi / 2 - angle)
                    angle = Left(angle,InStr(angle,".") + 3)
                    SSProcess.SetObjectAttr id, "[ORIENTATION]", angle
                End If
                If angle1 > 90 Or angle1 = 90 Then
                    angle = SSProcess.RadianToDms  ((5 * Pi / 2) - angle)
                    angle = Left(angle,InStr(angle,".") + 3)
                    SSProcess.SetObjectAttr id, "[ORIENTATION]", angle
                End If
            End If
        Next
    End If
    
End Function

Function DateRq
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Type", "==", "Point,Line,AREA"
    SSProcess.SelectFilter
    count = SSProcess.GetSelGeoCount
    If count > 0 Then
        For i = 0 To count - 1
            id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
            GXSJ = SSProcess.GetObjectAttr(id,"[GXRQ]")
            JCNY = SSProcess.GetObjectAttr(id,"[JCNY]")
            KSSYNY = SSProcess.GetObjectAttr(id,"[KSSYNY]")
            SJBCSJ = SSProcess.GetObjectAttr(id,"[SJBCSJ]")
            TCRQ = SSProcess.GetObjectAttr(id,"[TCRQ]")
            ZJYCDZXHGKJSJ = SSProcess.GetObjectAttr(id,"[ZJYCDZXHGKJSJ]")
            ZSSJ = SSProcess.GetObjectAttr(id,"[ZSSJ]")
            PCRQ = SSProcess.GetObjectAttr(id,"[PCRQ]")
            If GXSJ <> "" Then
                b = InStr(GXSJ," ")
                If b = 0 Then  SSProcess.SetObjectAttr id, "[GXSJ]", CDate(GXSJ)
                If b <> 0 Then  SSProcess.SetObjectAttr id, "[GXSJ]", CDate(Left(GXSJ,b - 1))
            End If
            If JCNY <> "" Then
                b = InStr(JCNY," ")
                If b = 0 Then  SSProcess.SetObjectAttr id, "[JCNY]", CDate(JCNY)
                If b <> 0 Then  SSProcess.SetObjectAttr id, "[JCNY]", CDate(Left(JCNY,b - 1))
            End If
            If KSSYNY <> "" Then
                b = InStr(KSSYNY," ")
                If b = 0 Then  SSProcess.SetObjectAttr id, "[KSSYNY]", CDate(KSSYNY)
                If b <> 0 Then  SSProcess.SetObjectAttr id, "[KSSYNY]", CDate(Left(KSSYNY,b - 1))
            End If
            If SJBCSJ <> "" Then
                b = InStr(SJBCSJ," ")
                If b = 0 Then  SSProcess.SetObjectAttr id, "[SJBCSJ]", CDate(SJBCSJ)
                If b <> 0 Then  SSProcess.SetObjectAttr id, "[SJBCSJ]", CDate(Left(SJBCSJ,b - 1))
            End If
            If TCRQ <> "" Then
                b = InStr(TCRQ," ")
                If b = 0 Then  SSProcess.SetObjectAttr id, "[TCRQ]", CDate(TCRQ)
                If b <> 0 Then  SSProcess.SetObjectAttr id, "[TCRQ]", CDate(Left(TCRQ,b - 1))
            End If
            If ZJYCDZXHGKJSJ <> "" Then
                b = InStr(ZJYCDZXHGKJSJ," ")
                If b = 0 Then  SSProcess.SetObjectAttr id, "[ZJYCDZXHGKJSJ]", CDate(ZJYCDZXHGKJSJ)
                If b <> 0 Then  SSProcess.SetObjectAttr id, "[ZJYCDZXHGKJSJ]", CDate(Left(ZJYCDZXHGKJSJ,b - 1))
            End If
            If ZSSJ <> "" Then
                b = InStr(ZSSJ," ")
                If b = 0 Then  SSProcess.SetObjectAttr id, "[ZSSJ]", CDate(ZSSJ)
                If b <> 0 Then  SSProcess.SetObjectAttr id, "[ZSSJ]", CDate(Left(ZSSJ,b - 1))
            End If
            If PCRQ <> "" Then
                b = InStr(PCRQ," ")
                If b = 0 Then  SSProcess.SetObjectAttr id, "[PCRQ]", CDate(PCRQ)
                If b <> 0 Then  SSProcess.SetObjectAttr id, "[PCRQ]", CDate(Left(PCRQ,b - 1))
            End If
        Next
    End If
End Function

Function SmallNumber
    mdbName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb mdbName
    sql = "update  地下管线点属性表 set  DMGC=Round(DMGC,3),ORIENTATION=Round(ORIENTATION,3),JBS=Round(JBS,2),ORIENTATION=Round(JS,2)"
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  地下管线线属性表 set  GXZDDMGC = round(GXZDDMGC,3),GXQDMS=Round(GXQDMS,3) ,GXZDMS=Round(GXZDMS,3) ,GXQDDMGC=Round(GXQDDMGC ,3),GXQDGDGC=Round(GXQDGDGC ,3),GXZDGDGC=Round(GXZDGDGC ,3),LL=Round(LL ,3)"
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  地下管线面属性表 set  QDDMGC=Round(QDDMGC,3),ZDDMGC=Round(ZDDMGC,3),QDJGDBDMGC=Round(QDJGDBDMGC,3),ZDJGDBDMGC=Round(ZDJGDBDMGC,3),QDDBFTHD=Round(QDDBFTHD ,3),ZDDBFTHD=Round(ZDDBFTHD ,3)"
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  SZDC_地下通道点属性表 set ORIENTATION=Round(ORIENTATION,3) "
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  SZDC_地下通道面属性表 set  QDDMGC=Round(QDDMGC,3),ZDDMGC=Round(ZDDMGC,3),QDJGDBDMGC=Round(QDJGDBDMGC,3),ZDJGDBDMGC=Round(ZDJGDBDMGC,3),QDDBFTHD=Round(QDDBFTHD ,3),ZDDBFTHD=Round(ZDDBFTHD ,3),CG=Round(CG,3),JZMJ=Round(JZMJ,3),SZCCG=Round(SZCCG,3),SZCJZMJ=Round(SZCJZMJ,3),LFKDA=Round(LFKDA,3),LFKDA1=Round(LFKDA1,3) ,LFKDA2=Round(LFKDA2,3) ,LF=Round(LF,3),PS=Round(PS,3),BJYCJ=Round(BJYCJ,3)"
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  SZDC_地下停车场面属性表 set  CG=Round(CG,3),SZCCG=Round(SZCCG,3),SZCJZMJ=Round(SZCJZMJ,3),DMGC=Round(DMGC,3),SJJZMJ=Round(SJJZMJ, 3),JGDBDMGC=Round(JGDBDMGC ,3),DBFTHD=Round(DBFTHD,3) "
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  SZDC_人防工程点属性表 set ORIENTATION=Round(ORIENTATION,3) "
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  SZDC_人防工程面属性表 set  CG=Round(CG,3),SZCCG=Round(SZCCG,3),JZMJ=Round(JZMJ,3),YJMJ=Round(YJMJ,3),YBMJ=Round(YBMJ,3),SZCYBMJ=Round(SZCYBMJ, 3),JGDBDMGC=Round(JGDBDMGC ,3),DBFTHD=Round(DBFTHD,3) ,DMGC=Round(DMGC,3),SZCJZMJ=Round(SZCJZMJ,3)"
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  SZDC_地下河道点属性表 set  ORIENTATION=Round(ORIENTATION,3)"
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  SZDC_地下河道面属性表 set  QDDMGC=Round(QDDMGC,3),ZDDMGC=Round(ZDDMGC,3)"
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  SZDC_废弃工程点属性表 set  DMGC=Round(DMGC,3),ORIENTATION=Round(ORIENTATION,3),JBS=Round(JBS,2),ORIENTATION=Round(JS,2)"
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  SZDC_废弃工程线属性表 set  GXQDMS=Round(GXQDMS,3) ,GXZDMS=Round(GXZDMS,3) ,GXQDDMGC=Round(GXQDDMGC ,3),GXQDGDGC=Round(GXQDGDGC ,3),GXZDGDGC=Round(GXZDGDGC ,3)"
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  SZDC_废弃工程面属性表 set  DMGC=Round(DMGC,3),JGDBDMGC=Round(JGDBDMGC ,3),DBFTHD=Round(DBFTHD,3) "
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  SZDC_其他地下空间设施点属性表 set ORIENTATION=Round(ORIENTATION,3) "
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  SZDC_其他地下空间设施面属性表 set  DMGC=Round(DMGC,3),JGDBDMGC=Round(JGDBDMGC ,3),DBFTHD=Round(DBFTHD,3),CG=Round(CG,3),SZCCG=Round(SZCCG,3),SZCJZMJ=Round(SZCJZMJ,3)"
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  SZDC_普查单元信息属性表 set  XLF=Round(XLF,3),WLF=Round(WLF,3),TP=Round(TP,3),CZBL=Round(CZBL,3),MM=Round(MM,3),GL=Round(GL,3),SS=Round(SS,3),CX=Round(CX,3),KC=Round(KC,3),FJ=Round(FJ,3),DB=Round(DB,3),KD=Round(KD,3),BDTK=Round(BDTK,3),PSui=Round(PSui,3),CT=Round(CT,3),PSun=Round(PSun,3),GQRQ=Round(GQRQ,3),TFLSH=Round(TFLSH,3),JCJXC=Round(JCJXC,3),JBLMSH=Round(JBLMSH,3)"
    SSProcess.ExecuteAccessSql  mdbName,sql
    SSProcess.CloseAccessMdb mdbName
End Function


'///获得文件夹下所有指定文件
Function GetAllFiles(ByRef pathname, ByRef fileExt, ByRef filecount, ByRef filenames())
    Dim fso, folder, file, files, subfolder,folder0, fcount
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(pathname)
    Set files = folder.Files
    '查找文件
    For Each file In files
        extname = fso.GetExtensionName(file.name)
        If UCase(extname) = UCase(fileExt) Then
            filenames(filecount) = pathname & file.name
            filecount = filecount + 1
        End If
    Next
End Function

#include "Function_beforeExportProcessFunc.vbs"

'********====<判断文件夹是否存在>==========&&&&&&&&&&
Function FolderExist(FolderName)
    FolderExist = True
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.folderExists(FolderName)  Then
        FolderExist = False
    End If
    Set fso = Nothing
End Function

Function CreateFolder(ByVal strFolderPath)
    Set Fso = CreateObject("Scripting.FileSystemObject")
    strDir = Left(strFolderPath,3)
    strFolderPath = Right( strFolderPath, Len(strFolderPath) - 3 )
    mulFolderPath = Split(strFolderPath,"\")
    nCount = UBound(mulFolderPath)
    strDirPath = strDir
    For i = 0 To nCount - 1
        strDirPath = strDirPath & "\" & mulFolderPath(i)
        If Fso.FolderExists(strDirPath) = False Then
            Fso.CreateFolder(strDirPath)
        End If
    Next
    Set Fso = Nothing
End Function