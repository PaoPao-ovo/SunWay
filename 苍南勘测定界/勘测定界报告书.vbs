
'========================================================文件路径操作对象================================================================

'路径操作对象
Dim FileSysObj
Set FileSysObj = CreateObject("Scripting.FileSystemObject")

'复制后路径字符串
Dim CopyPathStr
CopyPathStr = ""

'Word操作对象
Dim g_docObj

'==============================================================功能入口==================================================================

Sub OnClick()
    
    strTempFileName = "苍南县宗地位置图模板.docx"
    strTempFilePath = SSProcess.GetSysPathName (7) & "\功能模板\" & strTempFileName
    
    Set g_docObj = CreateObject ("asposewordscom.asposewordshelper")
    
    If TypeName (g_docObj) = "AsposeWordsHelper" Then
        g_docObj.CreateDocumentByTemplate strTempFilePath
    Else
        MsgBox "请先注册Aspose.Word插件"
        Exit Sub
    End If
    
    pathName = GetFilePath
    
    InitDB()
    
    ReplaceCell ZhuJi '替换单元格文字
    
    ReleaseDB()
    
    strWordFileName = Replace(strTempFileName, "模板", "")
    
    strFileSavePath = pathName & "报告成果\" & strWordFileName
    
    g_docObj.SaveEx strFileSavePath
    
    OpenProject strFileSavePath,pathName,strWordFileName,ZhuJi
    
    Set TempFile = FileSysObj.GetFile(strFileSavePath)
    
    TempFile.Delete
    
    MsgBox "输出完成"
    
End Sub

Function OpenProject(ByVal strFileSavePath,ByVal pathName,ByVal strWordFileName,ByVal ZhuJi)
    
    '工程名称
    EdbNameStr = ""
    
    '复制征地范围面（504）到粘贴板
    CloneArea
    
    '选择文件的路径(多选之间以","进行分隔)
    FilePathStr = SSProcess.SelectFileName(1,"选择文件",1,"EDB Files (*.edb)|*.edb|All Files (*.*)|*.*||")
    
    FilePathArr = Split(FilePathStr,",", - 1,1)
    
    '复制工程
    For i = 0 To UBound(FilePathArr)
        Set EdbFile = FileSysObj.GetFile(FilePathArr(i))
        EdbFile.Copy SSProcess.GetSysPathName(5) & "报告成果\" & EdbFile.Name
        If CopyPathStr = "" Then
            CopyPathStr = SSProcess.GetSysPathName(5) & "报告成果\" & EdbFile.Name
        Else
            CopyPathStr = CopyPathStr & "," & SSProcess.GetSysPathName(5) & "报告成果\" & EdbFile.Name
        End If
        If EdbNameStr = "" Then
            EdbNameStr = EdbFile.Name
        Else
            EdbNameStr = EdbNameStr & "," & EdbFile.Name
        End If
    Next 'i
    
    EdbNameArr = Split(EdbNameStr,",", - 1,1)
    
    For i = 0 To UBound(EdbNameArr)
        CreatPath = pathName & "报告成果\" & Replace(EdbNameArr(i),".edb","") & strWordFileName
        g_docObj.CreateDocumentByTemplate strFileSavePath
        ReadDT Replace(EdbNameArr(i),".edb","")
        g_docObj.SaveEx CreatPath
    Next 'i 
    
    EdbScale = SSProcess.GetMapScale()
    
    '打开选择的工程并粘贴
    CopyPathArr = Split(CopyPathStr,",", - 1,1)
    
    For i = 0 To UBound(CopyPathArr)
        
        SSProcess.OpenDatabase CopyPathArr(i)
        SSProcess.SetMapScale(EdbScale)
        SSProcess.AddClipBoardObjToMap 0,0
        CopyArr = Split(CopyPathArr(i),"\", - 1,1)
        EdbName = Replace(CopyArr(UBound(CopyArr)),".edb","")
        g_docObj.OpenDocument pathName & "报告成果\" & EdbName & strWordFileName
        
        NotePosition X,Y
        
        CJCTFWX MinX,MinY,MaxX,MaxY
        
        ZhuJiArr = Split(ZhuJi,";", - 1,1)
        
        For j = 0 To UBound(ZhuJiArr)
            DrawNote ZhuJiArr(j),X,Y + j * 6,500 / EdbScale
        Next 'j
        
        DrawCompass EdbScale,MinX,MinY,MaxX,MaxY
        
        InsterPicture MinX,MinY,MaxX,MaxY
        
        g_docObj.SaveEx pathName & "报告成果\" & EdbName & strWordFileName
        'SSProcess.CloseDatabase
    Next 'i
    
End Function' OpenProject()

'绘制图例和指南针
Function DrawCompass(ByVal EdbScale,ByRef MinX,ByRef MinY,ByRef MaxX,ByRef MaxY)
    If EdbScale = 500 Then
        makePoint MaxX - 10,MaxY - 10,"9120066",RGB(255,255,255),polygonID  '绘制指南针
        makePoint MinX + 10,MinY + 10,"9120046",RGB(255,255,255),polygonID  '绘制图例
    ElseIf EdbScale = 5000 Then
        makePoint MaxX,MaxY,"9120066",RGB(255,255,255),polygonID  '绘制指南针
        makePoint MinX,MinY,"9120047",RGB(255,255,255),polygonID  '绘制图例
    End If
    'makeArea MinX,MinY,MaxX,MinY,MaxX,MaxY,MinX,MaxY,2,RGB(255,255,255)
End Function' DrawNote

'返回注记位置
Function NotePosition(ByRef X,ByRef Y)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 504
    SSProcess.SelectFilter
    SelCount = SSProcess.GetSelGeoCount
    ID = SSProcess.GetSelGeoValue(0,"SSObj_ID")
    SSProcess.GetObjectFocusPoint ID,X,Y
End Function' NotePosition

'选择要素复制到粘贴板
Function CloneArea()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 504
    SSProcess.SelectFilter
    SSProcess.SelectionObjToClipBoard
End Function' CloneArea

'插入图片
Function InsterPicture(ByVal MinX,ByVal MinY,ByVal MaxX,ByVal MaxY)
    
    Path = SSProcess.GetSysPathName(7) & "Pictures\"
    StrBmpFile = Path & "RFT" & i & ".wmf"
    Dpi = 300

    SSFunc.DrawToImage MinX,MinY,MaxX,MaxY,"100" & "X" & "100",Dpi,StrBmpFile
    Rotation = 0
    
    Width = 100 * 4.28
    Height = 100 * 4.28
    
    g_docObj.MoveToCell TableIndex,5,0,0
    
    g_docObj.InsertImage StrBmpFile,Width,Height,Rotation
    
End Function' InsterPicture

'获取征地范围坐标    
Function CJCTFWX(ByRef MinX,ByRef MinY,ByRef MaxX,ByRef MaxY)

    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_code", "==", "504"
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount
    dh = 0
    For i = 0 To geocount - 1
        pointcount = SSProcess.GetSelGeoPointCount(i)
        objID = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        Dim x, y, z, pointtype, name
        For j = 0 To pointcount - 1
            dh = dh + 1
            SSProcess.GetObjectPoint objID, j, x, y, z, pointtype, name
            If dh <> 1 Then
                If  x > maxx Then  maxx = x
                If  x < minx Then  minx = x
                If  y > maxy Then  maxy = y
                If  y < miny Then  miny = y
            Else
                maxx = x
                minx = x
                maxy = y
                miny = y
            End If
        Next
    Next
    
    '最小框坐标
    MinX = MinX - 10
    MaxX = MaxX + 10
    MinY = MinY - 10
    MaxY = MaxY + 10
    
End Function

Function makeArea(x1,y1,x2,y2,x3,y3,x4,y4,code,color)
    
    SSProcess.CreateNewObj 2
    SSProcess.SetNewObjValue "SSObj_Code", code
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjPoint x2, y2, 0, 0, ""
    SSProcess.AddNewObjPoint x3, y3, 0, 0, ""
    SSProcess.AddNewObjPoint x4,y4, 0, 0, ""
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
    
End Function

Function makePoint(x,y,code,color,polygonID)
    SSProcess.CreateNewObj 0
    SSProcess.SetNewObjValue "SSObj_Code", code
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.AddNewObjPoint x, y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

Function ReadDT(ByVal YearTime)
    Tablename = "苍南县土地利用现状图【" & YearTime & "】宗地区位图"
    g_docObj.Replace "{Tablename}",Tablename,0
End Function

Function ReplaceCell(ByRef Hr)
    SuoZXZ = WriteFormHX  '从红线读取属性填写模板，并返回所属乡镇
    Res_value = WriteFormDLTB '从图斑读取属性填写模板，并返回土地权属和地图要使用的注记
    
    ZhuJi = Split(Res_value, "||")(0) '地图上的注记
    ZhuJi = Replace(ZhuJi, " ", "、")
    ZhuJi = Right(ZhuJi, Len(ZhuJi) - 1)
    ZjArr = Split(ZhuJi,"、", - 1,1)
    Hr = ""
    For i = 0 To UBound(ZjArr)
        If Hr = "" Then
            Hr = ZjArr(i)
        Else
            If i > 0 And i Mod 3 = 0 Then
                Hr = Hr & ";" & ZjArr(i)
            Else
                Hr = Hr & "、" & ZjArr(i)
            End If
        End If
    Next 'i
    TuDQS = SuoZXZ & Split(Res_value, "||")(1) '土地权属
    
    g_docObj.Replace "{TuDQS}",TuDQS,0
    
    MapScale = SSProcess.GetMapScale
    MapScale = "1:" & MapScale
    g_docObj.Replace "{MapScale}",MapScale,0
    
End Function

Function DrawNote(ByVal BZStr,ByVal X,ByVal Y,ByVal NoteScale)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontString", BZStr
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "5"
    SSProcess.SetNewObjValue "SSObj_FontWidth", 1000 * NoteScale
    SSProcess.SetNewObjValue "SSObj_Color", "RGB(0,255,191)"
    SSProcess.SetNewObjValue "SSObj_FontHeight", 1000 * NoteScale
    SSProcess.SetNewObjValue "SSObj_FontDirection", 0
    SSProcess.AddNewObjPoint X,Y,0,0,""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
    SSProcess.RefreshView
End Function' DrawNote

Function WriteFormDLTB

    '根据地类图斑输出土地权属、土地类型、{TuDQS}、{TuDLX} 
    TuDLX = ""
    ZhuJI = ""
    TuDQS = ""
    mdbName = SSProcess.GetProjectFileName
    sql = "Select DISTINCT 地类图斑属性表.dlmc From 地类图斑属性表 INNER JOIN GeoAreaTB ON 地类图斑属性表.ID = GeoAreaTB.ID WHERE([GeoAreaTB].[Mark] Mod 2)<>0"
    GetSQLRecordAll mdbName,sql,arSQLRecord,iRecordCount
    For i = 0 To iRecordCount - 1
        dlmc = arSQLRecord (i)
        TuDLX = TuDLX & " " & dlmc
        
        sql1 = "Select sum (地类图斑属性表.tbmj) From 地类图斑属性表 INNER JOIN GeoAreaTB ON 地类图斑属性表.ID = GeoAreaTB.ID WHERE  ([GeoAreaTB].[Mark] Mod 2)<>0 and 地类图斑属性表.dlmc = '" & dlmc & "'"
        GetSQLRecordAll mdbName,sql1,arSQLRecord1,iRecordCount1
        For j = 0 To iRecordCount1 - 1
            message = dlmc & arSQLRecord1(j) & "公顷"
            ZhuJI = ZhuJI & " " & message
        Next
    Next
    
    sql2 = "Select DISTINCT 地类图斑属性表.qsdw From 地类图斑属性表 INNER JOIN GeoAreaTB ON 地类图斑属性表.ID = GeoAreaTB.ID WHERE([GeoAreaTB].[Mark] Mod 2)<>0"
    GetSQLRecordAll mdbName,sql2,arSQLRecord2,iRecordCount2
    For k = 0 To iRecordCount2 - 1
        qsdw = arSQLRecord2 (k)
        TuDQS = TuDQS & "、" & qsdw
        
        TuDQS = Right(TuDQS, Len(TuDQS) - 1)
        
    Next
    TuDLX = Replace(TuDLX, " ", "、")
    
    TuDLX = Right(TuDLX, Len(TuDLX) - 1)
    
    g_docObj.Replace "{TuDLX}",TuDLX,0
    WriteFormDLTB = ZhuJI & "||" & TuDQS
    
End Function


Function WriteFormHX
    '根据红线属性中的用地单位,项目名称,地块面积 {XMMC},{YDDW},{DKMJ}填模板中的对应字段
    ' 将所在乡镇,{SuoZXZ}返回
    values = "XMMC,YDDW,DKMJ"
    valuesList = Split(values,",")
    SqlStr = "Select 征地属性表.DKMJ,YDDW,XMMC From 征地属性表 Inner Join GeoAreaTB on 征地属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0"
    ProJectName = SSProcess.GetProjectFileName
    GetSQLRecordAll ProJectName,SqlStr,InfoArr,DKCount
    Mj = 0
    For i = 0 To DKCount - 1
        TotalArr = Split(InfoArr(i),",", - 1,1)
        If Mj = 0 Then
            Mj = Transform(TotalArr(0))
        Else
            Mj = Mj + Transform(TotalArr(0))
        End If
    Next 'i
    First = Split(InfoArr(0),",", - 1,1)
    ValStr = First(1) & "," & First(2) & "," & Mj
    strFieldValue = Split(ValStr,",", - 1,1)
    For i = 0 To UBound(valuesList)
        'strFieldValue = ""
        strField = valuesList(i)
        'listCount = GetProjectTableList ("征地属性表",strField," 征地属性表.ID>0 ","SpatialData","2",list,fieldCount)
        'If listCount = 1 Then strFieldValue = list(0,0)
        g_docObj.Replace "{" & strField & "}",strFieldValue(i),0
    Next
    
    listCount = GetProjectTableList ("征地属性表","SuoZXZ"," 征地属性表.ID>0 ","SpatialData","2",list,fieldCount)
    If listCount = 1 Then
        SuoZXZ = list(0,0)
    End If
    WriteFormHX = SuoZXZ
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
            arSQLRecord(iRecordCount) = values    '查询记录
            iRecordCount = iRecordCount + 1        '查询记录数
            '移动记录游标
            SSProcess.AccessMoveNext mdbName, sql
        WEnd
    End If
    '关闭记录集
    SSProcess.CloseAccessRecordset mdbName, sql
    SSProcess.CloseAccessMdb mdbName
End Function


'获取成果目录路径
Function  GetFilePath
    projectFileName = SSProcess.GetSysPathName (5)
    GetFilePath = projectFileName
End Function

'***********************数据库操作函数*********************
'//开库
Dim  adoConnection
Function InitDB()
    accessName = SSProcess.GetProjectFileName
    Set adoConnection = CreateObject("adodb.connection")
    strcon = "DBQ=" & accessName & ";DRIVER={Microsoft Access Driver (*.mdb)};"
    adoConnection.Open strcon
End Function


'关库
Function ReleaseDB()
    adoConnection.Close
    Set adoConnection = Nothing
End Function


'递归创建多级目录
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


'SQL查询字段
Function GetProjectTableList(ByVal strTableName,ByVal strFields,ByVal strAddCondition,ByVal strTableType,ByVal strGeoType,ByRef rs(),ByRef fieldCount)
    'strTableName 表
    'strFields 字段
    'strAddCondition 条件 
    'strTableType AttributeData(纯属性表) ,SpatialData(地物属性表)
    'strGeoType 地物类型 点、线、面、注记(0点,1线,2面,3注记)
    'rs 表记录二维数组rs(行,列)
    'fieldCount 字段个数
    '返回值 :sql查询表记录个数
    
    
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
    
    '获取当前工程edb表记录
    AccessName = SSProcess.GetProjectFileName
    '判断表是否存在
    Set adoRs = CreateObject("ADODB.recordset")
    count = 0
    adoRs.cursorLocation = 3
    adoRs.cursorType = 3
    
    adoRs.open sql,adoConnection,3,3
    rcdCount = adoRs.RecordCount
    fieldCount = adoRs.Fields.Count
    ReDim rs(rcdCount,fieldCount)
    
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
    
    GetProjectTableList = rsCount
End Function

'数据类型转换
Function Transform(ByVal Values)
    If Values <> "" Then
        If IsNumeric(Values) = True Then
            Values = CDbl(Values)
        End If
    Else
        Values = 0
    End If
    Transform = Values
End Function'Transform