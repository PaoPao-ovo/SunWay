
Dim g_docObj
XMZD = "BH,XMMC,XMDZ,SJDW,JSDW,WTDW,CHDW,WYSJ,CHSJ"
GXLayerName = "CD,CT,CY,CQ,CS,QT,BM,FQ,DL,GD,LD,DC,XH,TX,DX,YD,LT,JX,JK,EX,DS,BZ,JS,XF,PS,YS,WS,FS,RQ,MQ,TR,YH,RL,RS,ZQ,SY,GS"

GXPointField0 = "TSDH,WTDH"
GXLineField0 = "GXQDDH^GXZDDH"
GXPointField1 = "TZ,FSW,Round(QDX,3),Round(QDY,3),DMGC"
GXLineField1 = "GC,GJ^DMCC,GXQDMS,GXZDMS,GXQDGDGC,GXZDGDGC,DYZ,SJYL,ZKS/YYKS,XLTS,LX,QSDW,FSFS,JCNY,BZ"
Sub OnClick()
    Set g_docObj = CreateObject ("AsposeCellsCom.AsposeCellsHelper")
    If  TypeName (g_docObj) <> "AsposeCellsHelper" Then
        MsgBox "请先注册Aspose.Excel插件"
        Exit Sub
    End If
    InitDB()
    str = GXAddInputParameter(filename,ExportMark,frameCount)
    If str = True Then
        pathName = SSProcess.GetSysPathName(5) & "成果表\"
        If pathName <> "" Then
            If ExportMark = "" Then
                '项目输出
                ExportMap pathName, filename
            Else
                '图幅输出
                ExportFrame pathName, filename, frameCount
            End If
        Else
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    ReleaseDB()
    MsgBox "输出完成"
End Sub


Function GXAddInputParameter(ByRef filename,ByRef ExportMark,ByRef frameCount)
    str = True
    SSProcess.ClearInputParameter
    SSProcess.AddInputParameter "管线报表输出方式", "按图幅输出", 0, "按图幅输出,按项目输出", ""
    result = SSProcess.ShowInputParameterDlg ("管线报表输出方式")
    If result = 1 Then
        res = SSProcess.GetInputParameter ("管线报表输出方式")
        If res = "按图幅输出" Then
            filename = SSProcess.GetSysPathName (7) & "\" & "管线报表模板（图幅）.xlsx"
            SSProcess.ClearInputParameter
            SSProcess.AddInputParameter "图幅输出输出方式", "单个图幅输出", 0, "单个图幅输出,按范围线分幅输出,全图分幅输出", ""
            result1 = SSProcess.ShowInputParameterDlg ("图幅输出输出方式")
            If result1 = 1 Then
                res1 = SSProcess.GetInputParameter ("图幅输出输出方式")
                SSProcess.ClearSelection
                SSProcess.ClearSelectCondition
                If res1 = "单个图幅输出" Then
                    frameID = SSProcess.GetCurMapFrame()
                    SSProcess.CreateMapFrameByRegionID  frameID
                    frameCount = SSProcess.GetMapFrameCount()
                ElseIf res1 = "按范围线分幅输出" Then
                    SSProcess.CreateMapFrameByRegion 1
                    frameCount = SSProcess.GetMapFrameCount()
                ElseIf res1 = "全图分幅输出" Then
                    SSProcess.CreateMapFrame
                    frameCount = SSProcess.GetMapFrameCount()
                End If
                ExportMark = 1
            Else
                Exit Function
                str = False
            End If
        Else
            filename = SSProcess.GetSysPathName (7) & "\" & "管线报表模板（项目）.xlsx"
            ExportMark = ""
        End If
    Else
        Exit Function
        str = False
    End If
    GXAddInputParameter = str
End Function

'项目输出
Function ExportMap(ByVal pathName,ByVal filename)
    g_docObj.CreateDocumentByTemplate filename
    '表头
    'XMMC=SSProcess.ReadEpsIni("管线报告信息", "项目名称" ,"")
    'CHDW=SSProcess.ReadEpsIni("管线报告信息", "测绘单位" ,""):RQDate=year(Date())&"年"&month(Date())&"月"
    GetXMXX GXXMXX
    XMMC = GXXMXX(1)
    CHDW = GXXMXX(6)
    RQDate = Year(Date()) & "年" & Month(Date()) & "月"
    '输出Excel路径
    exlfilePathName = pathName & XMMC & ".xlsx"
    '获取全图内的管线图层
    strLayer = GetMapGXLayer(strLayer)
    strLayerList = Split(strLayer,",")
    '复制excelsheet
    For i1 = 0 To UBound(strLayerList)
        g_docObj.CopySheet "Sheet1",strLayerList(i1)
    Next
    
    '按照图层顺序填sheet值
    For i1 = 0 To UBound(strLayerList)
        g_docObj.SetActiveSheet strLayerList(i1)
        
        g_docObj.SetCellValueEx 0,0,"工程名称：" & XMMC
        
        ids = GetGXLayerID( strLayerList(i1))
        idsList = Split(ids,",")
        ReDim CellList(Count)
        Count = 0
        For i2 = 0 To UBound(idsList)
            value = ""
            CellValue = ""
            layer = SSProcess.GetObjectAttr( idsList(i2), "SSObj_LayerName")
            If layer = strLayerList(i1) Then
                value = SSProcess.GetObjectAttr( idsList(i2), "[WTDH]")
                WTDH = value
                CellValue = WTDH
                
                '管线点填值判断标识
                rMarkCount = 0
                '连接点号
                ZDDHCount = GetProjectTableList( "地下管线线属性表", "GXZDDH,地下管线线属性表.ID", "GXQDDH='" & WTDH & "'", "SpatialData", "1", ZDDHList, fieldCount)
                If ZDDHCount > 0 Then
                    For i3 = 0 To ZDDHCount - 1
                        If rMarkCount > 0 Then CellValue = ""
                        '获取管点为管线起点的属性数组
                        GetCellValueList i3,ZDDHList,"GXQDDH",WTDH, CellValue, CellList, Count,rMarkCount
                    Next
                End If
                QDDHCount = GetProjectTableList( "地下管线线属性表", "GXQDDH,地下管线线属性表.ID", "GXZDDH='" & WTDH & "'", "SpatialData", "1", QDDHList, fieldCount)
                If QDDHCount > 0 Then
                    For i3 = 0 To QDDHCount - 1
                        If rMarkCount > 0 Then CellValue = ""
                        '获取管点为管线终点的属性数组
                        GetCellValueList i3,QDDHList,"GXZDDH",WTDH, CellValue, CellList, Count,rMarkCount
                    Next
                End If
            End If
        Next
        
        '复制行
        g_docObj.CopySheetRows 3,1,Count - 1
        
        '填值
        startRow = 3
        For i2 = 0 To Count - 1
            CellValueList = Split(CellList(i2),",")
            For i3 = 0 To UBound(CellValueList)
                CellValue = CellValueList(i3)
                g_docObj.SetCellValueEx startRow,i3,CellValue
            Next
            startRow = startRow + 1
        Next
        
        '填写表尾
        ' g_docObj.SetCellValueEx Count + 3,0,"作业单位：" & CHDW
        ' g_docObj.SetCellValueEx Count + 3,10,"日期：" & RQDate
        
        '删除第一行
        g_docObj.DeleteSheetRows 0,1
        
        HeaderStr = "工程名称：" & XMMC
        g_docObj.PageSetup2 0,HeaderStr,"作业单位：" & CHDW & Space(30) & " 制表者：苏世景"
        g_docObj.PageSetup2 1,"","校核者：林培兵" & Space(30) & "日期：" & RQDate
    Next
    
    '删除sheet1
    g_docObj.DeleteSheet "Sheet1"
    g_docObj.SaveEx exlfilePathName, 0
    
End Function


'图幅输出
Function ExportFrame(ByVal pathName,ByVal filename,ByVal frameCount)
    For i = 0 To frameCount - 1
        g_docObj.CreateDocumentByTemplate filename
        
        SSProcess.GetMapFrameCenterPoint i, x, y
        SSProcess.SetCurMapFrame x, y, 0, ""
        frameID = SSProcess.GetCurMapFrame()
        mapNumber = SSProcess.GetCurMapFrameNumber()
        '输出Excel路径
        exlfilePathName = pathName & mapNumber & ".xlsx"
        
        ids = SSProcess.SearchInPolyObjIDs(frameID, 0, "", 1, 1, 1)
        idsList = Split(ids,",")
        '获取图幅内的管线图层
        strLayer = GetFrameGXLayer( ids, strLayer)
        strLayerList = Split(strLayer,",")
        '复制excelsheet
        For i1 = 0 To UBound(strLayerList)
            g_docObj.CopySheet "Sheet1",strLayerList(i1)
        Next
        '获取项目信息
        GetXMXX GXXMXX
        XMMC = GXXMXX(1)
        CHDW = GXXMXX(6)
        '按照图层顺序填sheet值
        For i1 = 0 To UBound(strLayerList)
            g_docObj.SetActiveSheet strLayerList(i1)
            '表头
            ''未填写
            'XMMC=SSProcess.ReadEpsIni("管线报告信息", "项目名称" ,""):XMBH=SSProcess.ReadEpsIni("管线报告信息", "编号" ,"")
            'CHDW=SSProcess.ReadEpsIni("管线报告信息", "测绘单位" ,""):RQDate=year(Date())&"年"&month(Date())&"月"
            
            g_docObj.SetCellValueEx 0,0,"工程名称：" & XMMC
            g_docObj.SetCellValueEx 0,8,"工程编号：" & XMBH
            g_docObj.SetCellValueEx 0,17,"图幅号：" & mapNumber
            
            ReDim CellList(Count)
            Count = 0
            For i2 = 0 To UBound(idsList)
                value = ""
                CellValue = ""
                layer = SSProcess.GetObjectAttr( idsList(i2), "SSObj_LayerName")
                If layer = strLayerList(i1) Then
                    Point0List = Split(GXPointField0,",")
                    str = ""
                    For i3 = 0 To UBound(Point0List)
                        value = SSProcess.GetObjectAttr( idsList(i2), "[" & Point0List(i3) & "]")
                        If i3 = 1 Then WTDH = value
                        CellValue = GetValueString( i3, value,str)
                    Next
                    '管线点填值判断标识
                    rMarkCount = 0
                    '连接点号
                    ZDDHCount = GetProjectTableList( "地下管线线属性表", "GXZDDH,地下管线线属性表.ID", "GXQDDH='" & WTDH & "'", "SpatialData", "1", ZDDHList, fieldCount)
                    If ZDDHCount > 0 Then
                        For i3 = 0 To ZDDHCount - 1
                            If rMarkCount > 0 Then CellValue = ","
                            '获取管点为管线起点的属性数组
                            GetCellValueList i3,ZDDHList,"GXQDDH",WTDH, CellValue, CellList, Count,rMarkCount
                        Next
                    End If
                    QDDHCount = GetProjectTableList( "地下管线线属性表", "GXQDDH,地下管线线属性表.ID", "GXZDDH='" & WTDH & "'", "SpatialData", "1", QDDHList, fieldCount)
                    If QDDHCount > 0 Then
                        For i3 = 0 To QDDHCount - 1
                            If rMarkCount > 0 Then CellValue = ","
                            '获取管点为管线终点的属性数组
                            GetCellValueList i3,QDDHList,"GXZDDH",WTDH, CellValue, CellList, Count,rMarkCount
                        Next
                    End If
                End If
            Next
            
            '复制行
            g_docObj.CopySheetRows 3,1,Count - 1
            
            '填值
            startRow = 3
            For i2 = 0 To Count - 1
                CellValueList = Split(CellList(i2),",")
                For i3 = 0 To UBound(CellValueList)
                    CellValue = CellValueList(i3)
                    g_docObj.SetCellValueEx startRow,i3,CellValue
                Next
                startRow = startRow + 1
            Next
            
            '填写表尾
            ' g_docObj.SetCellValueEx Count + 3,0,"作业单位：" & CHDW
            ' g_docObj.SetCellValueEx Count + 3,10,"日期：" & RQDate
            
            '删除第一行
            g_docObj.DeleteSheetRows 0,1
            
            HeaderStr = "工程名称：" & XMMC
            g_docObj.PageSetup2 0,HeaderStr,"作业单位：" & CHDW & Space(30) & " 制表者：苏世景"
            g_docObj.PageSetup2 1,"","校核者：林培兵" & Space(30) & "日期：" & RQDate
            
        Next
        
        
        '删除sheet1
        g_docObj.DeleteSheet "Sheet1"
        g_docObj.SaveEx exlfilePathName, 0
    Next
    SSProcess.FreeMapFrame()
End Function

'获取管线属性数组
Function GetCellValueList(ByVal i,ByVal DHList,ByVal GXLineDH,ByVal WTDH,ByVal CellValue,ByRef CellList(),ByRef Count,ByRef rMarkCount)
    GXCellDH = DHList(i,0)
    GXID = DHList(i,1)
    CellValue = CellValue & "," & GXCellDH
    PonitValueCount = GetProjectTableList( "地下管线点属性表", GXPointField1, "WTDH='" & WTDH & "'", "SpatialData", "0", GXDList, fieldCount)
    If PonitValueCount = 1 Then
        If rMarkCount = 0 Then
            For i4 = 0 To fieldCount - 1
                value1 = GXDList(0,i4)
                CellValue = GetValueString( i3, value1,CellValue)
            Next
        Else
            CellValue = CellValue & ",,,,,"
        End If
    End If
    
    '管线属性，^为或者，/为并列
    GXLineField1List = Split(GXLineField1,",")
    For i3 = 0 To UBound(GXLineField1List)
        If InStr(GXLineField1List(i3),"^") > 0 Then
            GJList = Split(GXLineField1List(i3),"^")
            GJvalue = SSProcess.GetObjectAttr( GXID, "[GJ]")
            DMCCValue = SSProcess.GetObjectAttr( GXID, "[DMCC]")
            '管径和断面尺寸空值判断填哪个值
            If GJvalue <> "" And DMCCValue = "" Then
                GXLineFields = GJList(0)
            ElseIf  GJvalue = "" And DMCCValue <> "" Then
                GXLineFields = GJList(1)
            End If
        ElseIf InStr(GXLineField1List(i3),"/") > 0 Then
            KSList = Split(GXLineField1List(i3),"/")
            GXLineFields = KSList(0) & "," & KSList(1)
        Else
            GXLineFields = GXLineField1List(i3)
        End If
        If InStr(GXLineFields,",") > 0 Then
            GXLineCount = GetProjectTableList( "地下管线线属性表", GXLineFields, "" & GXLineDH & "='" & WTDH & "' and 地下管线线属性表.ID=" & GXID, "SpatialData", "1", ValueList, fieldCount)
            ZKS = ValueList(0,0)
            YYKS = ValueList(0,1)
            '组合总孔数和已用孔数
            If ZKS <> "" And YYKS = "" Then
                value1 = ZKS
            ElseIf ZKS <> "" And YYKS <> "" Then
                value1 = ZKS & "/" & YYKS
            ElseIf ZKS = "" And YYKS = "" Then
                value1 = ""
            End If
            CellValue = GetValueString( i3, value1,CellValue)
        Else
            GXLineCount = GetProjectTableList( "地下管线线属性表", GXLineFields, "" & GXLineDH & "='" & WTDH & "' and 地下管线线属性表.ID=" & GXID, "SpatialData", "1", ValueList, fieldCount)
            value1 = ValueList(0,0)
            CellValue = GetValueString( i3, value1,CellValue)
        End If
    Next
    CellValue = Replace(CellValue,"*","")
    ReDim Preserve CellList(Count)
    CellList(Count) = CellValue
    Count = Count + 1
    rMarkCount = rMarkCount + 1
End Function

'获取全图的管线点图层
Function GetMapGXLayer(ByRef strLayer)
    strLayer = ""
    GXLXCount = GetProjectTableList( "地下管线线属性表", "distinct(GXLX)", "", "SpatialData", "1", ValueList, fieldCount)
    For i = 0 To GXLXCount - 1
        GXLX = ValueList(i,0)
        strLayer = GetString( GXLX, "," , strLayer)
    Next
    GetMapGXLayer = strLayer
End Function

'选择集获取指定图层的管点id
Function GetGXLayerID(ByVal strLayer)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName","==",strLayer
    SSProcess.SetSelectCondition "SSObj_Type","==","POINT"
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    str = ""
    For i = 0 To geocount - 1
        id = SSProcess.GetSelGeoValue( i, "SSObj_ID")
        ids = GetString(id, "," , str)
    Next
    GetGXLayerID = ids
End Function


'获取图幅内的管线图层
Function GetFrameGXLayer(ByVal ids,ByRef strLayer)
    idlist = Split(ids,",")
    strLayer = ""
    For i = 0 To UBound(idlist)
        layer = SSProcess.GetObjectAttr( idlist(i), "SSObj_LayerName")
        If InStr(GXLayerName,layer) > 0  Then
            If InStr(strLayer,layer) = 0 Then
                strLayer = GetString( layer, "," , strLayer)
            End If
        End If
    Next
    GetFrameGXLayer = strLayer
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

'整理字段值字符串
Function GetValueString(ByVal index,ByVal value,ByRef str)
    If str = "" Then
        If index = 0 Then
            str = value
        Else
            str = str & "," & value
        End If
    Else
        str = str & "," & value
    End If
    GetValueString = str
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


Function GetXMXX(ByRef XMXXSZ())
    
    mdbName = SSProcess.GetProjectFileName
    sql = "Select 管线项目信息表." & XMZD & " From 管线项目信息表  WHERE 管线项目信息表.id=1"
    GetSQLRecordAll mdbName,sql,arSQLRecord,iRecordCount
    For i = 0 To iRecordCount - 1
        XMXXSZ = Split(arSQLRecord(i), ",")
    Next
End Function


Function GetSQLRecordAll(ByRef mdbName, ByRef StrSqlStatement, ByRef arSQLRecord(), ByRef iRecordCount)
    SSProcess.OpenAccessMdb mdbName
    iRecordCount =  - 1
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
        '浏览记录
        While SSProcess.AccessIsEOF (mdbName, sql) = 0
            fields = ""
            values = ""
            '获取当前记录内容
            SSProcess.GetAccessRecord mdbName, sql, fields, values
            arSQLRecord(iRecordCount) = values
            iRecordCount = iRecordCount + 1
            '移动记录游标
            SSProcess.AccessMoveNext mdbName, sql
        WEnd
    End If
    '关闭记录集
    SSProcess.CloseAccessRecordset mdbName, sql
    SSProcess.CloseAccessMdb mdbName
End Function