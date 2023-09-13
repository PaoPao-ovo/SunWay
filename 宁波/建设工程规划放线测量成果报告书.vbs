Dim g_docObj'doc全局变量
Dim fso, f1
Set fso = CreateObject("Scripting.FileSystemObject")

Sub OnClick()
    strTempFileName = "建设工程规划放线测量成果报告书.doc"
    strTempFilePath = SSProcess.GetSysPathName (7) & strTempFileName
    Set g_docObj = CreateObject ("asposewordscom.asposewordshelper")
    If  TypeName (g_docObj) = "AsposeWordsHelper" Then
        g_docObj.CreateDocumentByTemplate strTempFilePath
    Else
        MsgBox "请先注册Aspose.Word插件"
        Exit Sub
    End If
    
    pathName = GetFilePath'SSProcess.SelectPathName()
    g_docObj.CreateDocumentByTemplate  strTempFilePath
    fwnr = SSProcess.ReadEpsIni("签章GUID", "fwnr" ,"")
    
    InitDB()
    '字符替换
    ReplaceValue
    '插入 实地规划放线平面图
    'OutputTable5 7
    OutputBook "建设工程实地放线平面图","建设工程实地放线平面图","9310093",9
    '建设工程规划放线验线记录表
    OutMap 7, "建设工程规划放线验线记录表"
    '输出 建设工程规划放线周边关系校核表
    OutputTable4 6
    '输出 建设工程规划放线边长校核表
    OutputTable3 5
    '输出 建设工程规划放线坐标校核表
    OutputTable2 4
    '输出 建设工程规划放线条件坐标表
    OutputTable1 3
    '控制点坐标表
    OutputTable02 2
    '输出 测绘项目技术人员
    OutputTable6 1
    '插入签章与水印
    InsertSignature
    ReleaseDB()
    '签章
    If fwnr = "" Then
        fwnr = fwnrGUid()
        SSProcess.WriteEpsIni "签章GUID", "fwnr" ,fwnr
    Else
        fwnr = fwnr
    End If
    
    strFileSavePath = pathName & strTempFileName
    g_docObj.SaveEx  strFileSavePath
    bRes = ProtectDoc(strFileSavePath,True,fwnr)
    Set g_docObj = Nothing
    MsgBox "输出完成"
End Sub

'//签章guid
Function fwnrGUid()
    Set TypeLib = CreateObject("Scriptlet.TypeLib")
    fwnrGUid = TypeLib.Guid
    fwnrGUid = Replace(fwnrGUid,"-","")
    fwnrGUid = Replace(fwnrGUid,"{","")
    fwnrGUid = Replace(fwnrGUid,"}","")
    fwnrGUid = Left(fwnrGUid,10)
    Set TypeLib = Nothing
End Function

'//加密解密docx文件
'//strFilePath 需要加密的doc文件路径
'//isProtectDoc true 加密 false解密
'//password 密码
Function ProtectDoc(ByVal strFilePath, ByVal  isProtectDoc,ByVal password)
    bRes = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FileExists(strFilePath)) = True  Then
        Set pDocObj = CreateObject ("asposewordscom.asposewordshelper")
        If  TypeName (pDocObj) = "AsposeWordsHelper" Then
            pDocObj.OpenDocument strFilePath
        Else
            bRes = False
            Exit Function
        End If
        
        If isProtectDoc = True Then
            str = pDocObj.ProtectDoc (password)
        Else
            str = pDocObj.UnProtectDoc (password)
        End If
        pDocObj.SaveEx  strFilePath
        Set pDocObj = Nothing
    End If
    Set fso = Nothing
    If InStr(str,"成功") Then bRes = True
    ProtectDoc = bRes
End Function

Function InsertSignature
    folderPath = SSProcess.GetSysPathName (0) & "\签章\"
    names = "水印"
    nameList = Split(names,",")
    For i = 0 To UBound(nameList)
        name = nameList(i)
        imageFile = folderPath & name & ".png"
        If name = "水印" Then
            If IsFileExists(imageFile) = True Then    g_docObj.SetImgWatermark imageFile, 400, 400,0
        Else
            g_docObj.MoveToBookmark name
            If IsFileExists(imageFile) = True Then    g_docObj.InsertImageEx imageFile,  0, 250, 0, 390, 150, 150,3, 0
        End If
    Next
End Function

'//判断文件是否存在
Function IsFileExists(filespec)
    IsFileExists = False
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FileExists(filespec)) = True Then
        IsFileExists = True
    End If
    Set fso = Nothing
End Function

'//获取成果目录路径
Function  GetFilePath
    projectFileName = SSProcess.GetProjectFileName()
    filePath = Replace(projectFileName,".edb","")
    filePath = filePath & "\"
    CreateFolder filePath
    GetFilePath = filePath
End Function

'//获取成果目录路径
Function  GetFilePath
    projectFileName = SSProcess.GetProjectFileName()
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


'//字符替换 
Function ReplaceValue
    values = "项目编号,规划许可证编号,项目名称,项目地址,委托单位,测绘单位,测绘资质证书编号,测量开始时间,测量完成时间,作业依据,已有资料情况,建设单位,不动产权证编号"
    valuesList = Split(values,",")
    For i = 0 To UBound(valuesList)
        strFieldValue = ""
        strField = valuesList(i)
        listCount = GetProjectTableList ("projectinfo","value","key='" & strField & "'","","",list,fieldCount)
        If listCount = 1 Then strFieldValue = list(0,0)
        If strField = "作业依据" Or  strField = "已有资料情况" Then
            chrlist = Split(strFieldValue,Chr(10))
            str = ""
            For i1 = 0 To UBound(chrlist)
                If chrlist(i1) <> "" Then
                    If str = "" Then
                        str = chrlist(i1)
                    Else
                        str = str & Chr(10) & chrlist(i1)
                    End If
                End If
            Next
            g_docObj.MoveToBookmark strField
            g_docObj.Write(str)
        Else
            g_docObj.Replace "{" & strField & "}",strFieldValue,0
        End If
    Next
    
    strFieldValue = ""
    strField = "仪器名称"
    listCount = GetProjectTableList ("INFO_YQSB",strField," ID>0 ","","",list,fieldCount)
    For i = 0 To listCount - 1
        name = list(i,0)
        If name <> "" Then If strFieldValue = "" Then strFieldValue = name Else strFieldValue = strFieldValue & "," & name
    Next
    g_docObj.Replace "{" & strField & "}",strFieldValue,0
    
    g_docObj.Replace "{年月日}",Year(Now) & "年" & Month(Now) & "月" & Day(Now) & "日",0
End Function


'//输出 建设工程规划放线条件坐标表
Function OutputTable1(ByVal tableIndex)
    g_docObj.MoveToTable tableIndex,False
    
    writeRowCount = 24
    copyCount = 0
    code = "9310001"
    tableType = 0
    strTableName = SSProcess.GetCodeAttrTableName(code,tableType)
    listCount = GetProjectTableList (strTableName,strTableName & ".id,dh",strTableName & ".id is not null order by dh asc","SpatialData",tableType,list,fieldCount)
    ReDim cellList(listCount)
    For i = 0 To listCount - 1
        objid = list(i,0)
        dh = list(i,1)
        x = SSProcess.GetObjectAttr (objid,"[x]")
        y = SSProcess.GetObjectAttr (objid,"[y]")
        x = GetFormatNumber(x,3)
        y = GetFormatNumber(y,3)
        cellValue = dh & "||" & y & "||" & x
        cellList(cellCount) = cellValue
        If i > 0 And i Mod writeRowCount * 2 = 0 Then copyCount = copyCount + 1
        cellCount = cellCount + 1
    Next
    
    '根据数据个数复制表格
    For i = 0 To copyCount - 1
        g_docObj.CloneTable  tableIndex, 1,0,False
    Next
    
    '数组按桩号从小到大冒泡排序
    For i = 0 To cellCount - 1
        For j = 0 To cellCount - 1 - 1
            cellValue = cellList(j)
            cellValueList = Split(cellValue,"||")
            cellValue1 = cellList(j + 1)
            cellValueList1 = Split(cellValue1,"||")
            num1 = cellValueList(0)
            num2 = cellValueList1(0)
            If IsNumeric(num1) = True And IsNumeric(num2) = True Then
                If CDbl(num1) > CDbl(num2) Then
                    temp = cellList(j)
                    cellList(j) = cellList(j + 1)
                    cellList(j + 1) = temp
                End If
            End If
        Next
    Next
    
    '填充表格单元格
    iniRow = 3
    iniCol = 0
    startRow = iniRow
    startCol = iniCol
    colIndex = 0
    For i = 0 To listCount - 1
        cellValue = cellList(i)
        cellValueList = Split(cellValue,"||")
        dh = cellValueList(0)
        x = cellValueList(1)
        y = cellValueList(2)
        '根据数组个数动态计算表、行、列索引
        If (i Mod writeRowCount = 0)  And i > 0 Then       colIndex = colIndex + 1
        startRow = iniRow
        If colIndex Mod 2 = 0 Then  startCol = iniCol  Else  startCol = iniCol + 3
        If i > 0 And i Mod writeRowCount * 2 = 0 Then tableIndex = tableIndex + 1
        startRow = iniRow
        
        g_docObj.SetCellText tableIndex,startRow,startCol,dh,True,False
        g_docObj.SetCellText tableIndex,startRow,startCol + 1,x,True,False
        g_docObj.SetCellText tableIndex,startRow,startCol + 2,y,True,False
        startRow = startRow + 1
    Next
End Function


'//输出 建设工程规划放线坐标校核表
Function OutputTable2(ByVal tableIndex)
    g_docObj.MoveToTable tableIndex,False
    
    code = "9310001"
    tableType = 0
    strTableName = SSProcess.GetCodeAttrTableName(code,tableType)
    listCount = GetProjectTableList (strTableName,strTableName & ".id,dh",strTableName & ".id is not null order by dh asc","SpatialData",tableType,list,fieldCount)
    ReDim cellList(listCount)
    cellCount = 0
    For i = 0 To listCount - 1
        objid = list(i,0)
        dh = list(i,1)
        x = SSProcess.GetObjectAttr (objid,"[x]")
        y = SSProcess.GetObjectAttr (objid,"[y]")
        
        '用 放线桩点 找 实测点
        code0 = "9310031"
        tableType0 = 0
        x0 = 0
        y0 = 0
        dist = 0
        strTableName0 = SSProcess.GetCodeAttrTableName(code0,tableType0)
        listCount0 = GetProjectTableList (strTableName0,strTableName0 & ".id,dh","dh='" & dh & "'","SpatialData",tableType0,list0,fieldCount)
        If listCount0 > 0 Then
            objid0 = list0(0,0)
            dh0 = list0(0,1)
            x0 = SSProcess.GetObjectAttr (objid0,"[x]")
            y0 = SSProcess.GetObjectAttr (objid0,"[y]")
            
            x = GetFormatNumber(x,3)
            x0 = GetFormatNumber(x0,3)
            y = GetFormatNumber(y,3)
            y0 = GetFormatNumber(y0,3)
            dist = Sqr((x - x0) * (x - x0) + (y - y0) * (y - y0))
            dist = GetFormatNumber(dist * 100,1)
            cellValue = dh & "||" & y & "||" & x & "||" & dh0 & "||" & y0 & "||" & x0 & "||" & dist
            cellList(cellCount) = cellValue
            cellCount = cellCount + 1
        End If
    Next
    
    writeRowCount = 25
    copyCount = 0
    For i = 0 To cellCount - 1
        If i > 0 And i Mod writeRowCount = 0 Then copyCount = copyCount + 1
    Next
    
    '数组按桩号从小到大冒泡排序
    For i = 0 To cellCount - 1
        For j = 0 To cellCount - 1 - 1
            cellValue = cellList(j)
            cellValueList = Split(cellValue,"||")
            cellValue1 = cellList(j + 1)
            cellValueList1 = Split(cellValue1,"||")
            num1 = cellValueList(0)
            num2 = cellValueList1(0)
            If IsNumeric(num1) = True And IsNumeric(num2) = True Then
                If CDbl(num1) > CDbl(num2) Then
                    temp = cellList(j)
                    cellList(j) = cellList(j + 1)
                    cellList(j + 1) = temp
                End If
            End If
        Next
    Next
    
    iniRow = 3
    iniCol = 0
    startRow = iniRow
    startCol = iniCol
    '根据数据个数复制表格
    For i = 0 To copyCount - 1
        g_docObj.CloneTable  tableIndex, 1,0,False
    Next
    
    '填充表格单元格
    For i = 0 To cellCount - 1
        cellValue = cellList(i)
        cellValueList = Split(cellValue,"||")
        If i > 0 And i Mod writeRowCount = 0 Then tableIndex = tableIndex + 1
        startRow = iniRow
        
        If  UBound(cellValueList) = 6 Then
            For  j = 0 To UBound(cellValueList)
                g_docObj.SetCellText tableIndex,startRow,startCol + j,cellValueList(j),True,False
            Next
        End If
        startRow = startRow + 1
    Next
End Function


'//输出 建设工程规划放线边长校核表
Function OutputTable3(ByVal tableIndex)
    
    g_docObj.MoveToTable tableIndex,False
    
    code = "9310092"
    tableType = 1
    strTableName = SSProcess.GetCodeAttrTableName(code,tableType)
    SetGhfxAttr code,strTableName,tableType
    
    listCount = GetProjectTableList (strTableName,strTableName & ".id,qsdzh,zzdzh,tjbc,slbc,bz",strTableName & ".id is not null and slbc>0 ","SpatialData",tableType,list,fieldCount)
    ReDim cellList(listCount)
    cellCount = 0
    bcNumCount = 0
    For i = 0 To listCount - 1
        objid = list(i,0)
        qsdzh = list(i,1)
        zzdzh = list(i,2)
        tjbc = list(i,3)
        slbc = list(i,4)
        bz = Replace(list(i,5),"*","")
        tjbc = GetFormatNumber(tjbc,3)
        slbc = GetFormatNumber(slbc,3)
        subNum = GetFormatNumber((CDbl(tjbc) - CDbl(slbc)),3)
        If slbc < 50 Then subLimit = 0.020 Else subLimit = 0.025
        subLimit = GetFormatNumber(subLimit,3)
        If tjbc <= 50 Then bcNumCount = bcNumCount + 1
        cellValue = qsdzh & "||" & zzdzh & "||" & tjbc & "||" & slbc & "||" & subNum & "||" & subLimit & "||" & bz
        cellList(cellCount) = cellValue
        cellCount = cellCount + 1
    Next
    
    writeRowCount = 13
    copyCount = 0
    For i = 0 To cellCount - 1
        If i > 0 And i Mod writeRowCount = 0 Then copyCount = copyCount + 1
    Next
    
    '根据数据个数复制表格
    For i = 0 To copyCount - 1
        g_docObj.CloneTable  tableIndex, 1,0,False
    Next
    
    '保留最后的总结
    For i = 0 To copyCount - 2
        g_docObj.DeleteRow tableIndex + i,14,False
        g_docObj.DeleteRow tableIndex + i,14,False
        g_docObj.CloneTableRow tableIndex + i,13,1,2,False
    Next 'i
    
    '数组按桩号从小到大冒泡排序
    For i = 0 To cellCount - 1
        For j = 0 To cellCount - 1 - 1
            cellValue = cellList(j)
            cellValueList = Split(cellValue,"||")
            cellValue1 = cellList(j + 1)
            cellValueList1 = Split(cellValue1,"||")
            num1 = cellValueList(0)
            num2 = cellValueList1(0)
            If IsNumeric(num1) = True And IsNumeric(num2) = True Then
                If CDbl(num1) > CDbl(num2) Then
                    temp = cellList(j)
                    cellList(j) = cellList(j + 1)
                    cellList(j + 1) = temp
                End If
            End If
        Next
    Next
    
    iniRow = 2
    iniCol = 1
    startRow = iniRow
    startCol = iniCol

    '填充表格单元格
    For i = 0 To cellCount - 1
        cellValue = cellList(i)
        cellValueList = Split(cellValue,"||")
        If i > 0 And i Mod writeRowCount = 0 Then tableIndex = tableIndex + 1
        startRow = iniRow
        
        If  UBound(cellValueList) = 6 Then
            g_docObj.SetCellText tableIndex,startRow,0,i + 1,True,False
            For  j = 0 To UBound(cellValueList)
                g_docObj.SetCellText tableIndex,startRow,startCol + j,cellValueList(j),True,False
            Next
        End If
        startRow = startRow + 1
    Next
    
    '计算并筛选符合条件的较差值
    subCount = 0
    ReDim  subList(subCount)
    For i = 0 To cellCount - 1
        cellValue = cellList(i)
        cellValueList = Split(cellValue,"||")
        tjbc = cellValueList(3)
        subNum = cellValueList(4)
        If CDbl(tjbc) <= 50 Then
            ReDim Preserve subList(subCount)
            subList(subCount) = Abs(subNum)
            subCount = subCount + 1
        End If
    Next

    '较差值冒泡排序
    For i = 0 To subCount - 1
        For j = 0 To subCount - 1 - 1
            If IsNumeric(subList(j)) = True And IsNumeric(subList(j + 1)) = True Then
                If CDbl(subList(j)) > (CDbl(subList(j + 1))) Then
                    temp = subList(j)
                    subList(j) = subList(j + 1)
                    subList(j + 1) = temp
                End If
            End If
        Next
    Next

    '获取最小,最大,平均较差值
    minSubNum = ""
    maxSubNum = ""
    averageSubNum = ""
    If subCount > 0 Then
        sumSubNum = 0
        For i = 0 To subCount - 1
            sumSubNum = sumSubNum + CDbl(subList(i))
        Next
        
        minSubNum = GetFormatNumber(subList(0),3)
        maxSubNum = GetFormatNumber(subList(subCount - 1),3)
        averageSubNum = CDbl(sumSubNum) / subCount
        averageSubNum = GetFormatNumber(averageSubNum,3)
    End If

    g_docObj.Replace "{最大较差}",maxSubNum,0
    g_docObj.Replace "{最小较差}",minSubNum,0
    g_docObj.Replace "{平均较差}",averageSubNum,0
    g_docObj.Replace "{小于50}",bcNumCount,0
    
End Function


'//设置规划放线起点、终点桩号属性
Function SetGhfxAttr(ByVal code,ByVal strTableName,ByVal tableType)
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", code
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    If geocount > 0 Then
        For i = 0 To geocount - 1
            objID = SSProcess.GetSelGeoValue(i, "SSObj_ID")
            pt = SSProcess.GetObjectAttr (objID, "SSObj_PointCount")
            If pt > 1 Then
                SSProcess.GetObjectPoint objID, 0, x0, y0, z0, pointtype0, name0
                SSProcess.GetObjectPoint objID, pt - 1, x1, y1, z1, pointtype1, name1
                
                ids0 = SSProcess.SearchNearObjIDs(x0, y0, 0.01, 0, "9310001", 0 )
                ids1 = SSProcess.SearchNearObjIDs(x1, y1, 0.01, 0, "9310001", 0 )
                
                If InStr(ids0,",") = 0 And ids0 <> "" Then  zh0 = SSProcess.GetObjectAttr (ids0, "[DH]")  Else zh0 = ""
                If InStr(ids1,",") = 0 And ids1 <> "" Then  zh1 = SSProcess.GetObjectAttr (ids1, "[DH]") Else zh1 = ""
                ModifyTableInfo strTableName, "qsdzh,zzdzh", zh0 & "," & zh1, strTableName & ".id =" & objID & "", "SpatialData",tableType
            End If
        Next
    End If
End Function


'//输出 建设工程规划放线周边关系校核表
Function OutputTable4(ByVal tableIndex)
    g_docObj.MoveToTable tableIndex,False
    
    code = "9310082"
    tableType = 1
    strTableName = SSProcess.GetCodeAttrTableName(code,tableType)
    listCount = GetProjectTableList (strTableName,strTableName & ".id,bz,sjjl,scjl",strTableName & ".id is not null ","SpatialData",tableType,list,fieldCount)
    ReDim cellList(listCount)
    cellCount = 0
    For i = 0 To listCount - 1
        objid = list(i,0)
        bz = list(i,1)
        sjjl = list(i,2)
        scjl = list(i,3)
        sjjl = GetFormatNumber(sjjl,2)
        scjl = GetFormatNumber(scjl,2)
        subNum = GetFormatNumber(CDbl(sjjl) - CDbl(scjl),2)
        
        cellValue = bz & "||" & sjjl & "||" & scjl & "||" & subNum
        cellList(cellCount) = cellValue
        cellCount = cellCount + 1
    Next
    
    writeRowCount = 14
    copyCount = 0
    For i = 0 To cellCount - 1
        If i > 0 And i Mod writeRowCount = 0 Then copyCount = copyCount + 1
    Next
    
    '根据数据个数复制表格
    For i = 0 To copyCount - 1
        g_docObj.CloneTable  tableIndex, 1,0,False
    Next
    
    '数组按桩号从小到大冒泡排序
    For i = 0 To cellCount - 1
        For j = 0 To cellCount - 1 - 1
            cellValue = cellList(j)
            cellValueList = Split(cellValue,"||")
            cellValue1 = cellList(j + 1)
            cellValueList1 = Split(cellValue1,"||")
            If InStr(cellValueList(0),"号") > 0 And InStr(cellValueList1(0),"号") > 0 Then
                If cellValueList(0) <> "" And cellValueList(0) <> "*" Then num1 = Mid(cellValueList(0),1,InStr(cellValueList(0),"号") - 1)
                If cellValueList1(0) <> ""  And cellValueList1(0) <> "*" Then num2 = Mid(cellValueList1(0),1,InStr(cellValueList1(0),"号") - 1)
                If IsNumeric(num1) = True And IsNumeric(num2) = True Then
                    If CDbl(num1) > CDbl(num2) Then
                        temp = cellList(j)
                        cellList(j) = cellList(j + 1)
                        cellList(j + 1) = temp
                    End If
                End If
            End If
        Next
    Next
    
    iniRow = 2
    iniCol = 1
    startRow = iniRow
    startCol = iniCol
    
    '填充表格单元格
    For i = 0 To cellCount - 1
        cellValue = cellList(i)
        cellValueList = Split(cellValue,"||")
        If i > 0 And i Mod writeRowCount = 0 Then tableIndex = tableIndex + 1
        startRow = iniRow
        
        If  UBound(cellValueList) = 3 Then
            g_docObj.SetCellText tableIndex,startRow,0,i + 1,True,False
            For  j = 0 To UBound(cellValueList)
                g_docObj.SetCellText tableIndex,startRow,startCol + j,cellValueList(j),True,False
            Next
        End If
        startRow = startRow + 1
    Next
End Function


'//插入 实地规划放线平面图
Function OutputTable5(ByVal tableIndex)
    g_docObj.MoveToTable tableIndex,False
    
    '查找成果图edb文件
    accessName = SSProcess.GetProjectFileName
    filePath = Replace(accessName,".edb","") & "\"
    Dim edbList(10000)
    listCount = 0
    GetAllFiles filePath,"edb",listCount,edbList
    fileName = "放线平面图"
    For i = 0 To listCount - 1
        edbPath = edbList(i)
        If InStr(edbPath,fileName) > 0 And InStr(fileName,"bak") = 0 Then
            Exit For
        End If
    Next
    If FileExists(edbPath) = False Then Exit Function
    '打开edb文件并按图廓范围打印wmf
    bRes = SSProcess.OpenDatabase (edbPath)
    If bRes = 1 Then
        PrintImage "9310093",fileName
        SSProcess.CloseDatabase()
        'set f1 = fso.GetFile( edbPath )
        'f1.Delete 
    End If
    'Set f1 = Nothing
    'Set fso = Nothing
    '查找对应wmf文件并插入word报告中
    Dim imageList(10000)
    listCount = 0
    filePath = SSProcess.GetSysPathName (4)
    GetAllFiles filePath,"WMF",listCount,imageList
    For i = 0 To listCount - 1
        imageFile = imageList(i)
        Exit For
    Next
    If FileExists( imageFile) = True Then   g_docObj.SetCellImageEx2 tableIndex,  0, 0, 0,  imageFile, 0, 0, False
End Function


Function OutputTable02(ByVal tableIndex)
    g_docObj.MoveToTable tableIndex,False
    startRow = 3
    listCount = GetProjectTableList ("KZDZBCGXXB","DH,Y,X,GC,GXSJ,BZ"," ID>0","","",List,fieldCount)
    If listCount > 1 Then g_docObj.CloneTableRow tableIndex,3,listCount - 1,1,False
    For i = 0 To listCount - 1
        startCol = 0
        DH = List(i,0)
        Y = List(i,1)
        X = List(i,2)
        GC = List(i,3)
        GXSJ = List(i,4)
        BZ = List(i,5)
        If Y <> "" Then Y = GetFormatNumber(Y,3)
        If X <> "" Then X = GetFormatNumber(X,3)
        If GC <> "" Then GC = GetFormatNumber(GC,3)
        
        If GXSJ <> "" Then GXSJ = Year(GXSJ) & "年" & Month(GXSJ) & "月"
        g_docObj.SetCellText tableIndex,startRow,0,DH,True,False
        g_docObj.SetCellText tableIndex,startRow,1,X,True,False
        g_docObj.SetCellText tableIndex,startRow,2,Y,True,False
        g_docObj.SetCellText tableIndex,startRow,3,GC,True,False
        g_docObj.SetCellText tableIndex,startRow,4,GXSJ,True,False
        g_docObj.SetCellText tableIndex,startRow,5,BZ,True,False
        
        startRow = startRow + 1
    Next
End Function

'//输出 测绘项目技术人员
Function OutputTable6(ByVal tableIndex)
    g_docObj.MoveToTable tableIndex,False
    '获取人员信息表单元格
    cellCount = 0
    ReDim cellList(cellCount)
    strField = "姓名,职称或职业资格,上岗证书编号或职业资格证书号,主要工作职责"
    listCount = GetProjectTableList ("info_RYXX",strField," ID>0 ","","",list,fieldCount)
    For i = 0 To listCount - 1
        cellValue = ""
        For j = 0 To fieldCount - 1
            value = list(i,j)
            If j = 0 Then  cellValue = value Else cellValue = cellValue & "||" & value
        Next
        cellValue = i + 1 & "||" & cellValue
        ReDim Preserve cellList(cellCount)
        cellList(cellCount) = cellValue
        cellCount = cellCount + 1
    Next
    
    '填充人员信息表单元格
    iniRow = 1
    iniCol = 0
    startRow = iniRow
    startCol = iniCol
    If cellCount > 1 Then   g_docObj.CloneTableRow tableIndex, iniRow, 1,cellCount - 1, False
    For i = 0 To cellCount - 1
        startCol = iniCol
        cellValue = cellList(i)
        cellValueList = Split(cellValue,"||")
        For j = 0 To UBound(cellValueList)
            g_docObj.SetCellText tableIndex,startRow,startCol,cellValueList(j),True,False
            startCol = startCol + 1
        Next
        startRow = startRow + 1
    Next
End Function


'//数字进位
Function GetFormatNumber(ByVal number,ByVal numberDigit)
    If IsNumeric(numberDigit) = False Then numberDigit = 2
    If IsNumeric(number) = False Then number = 0
    number = FormatNumber(Round(number + 0.00000001,numberDigit),numberDigit, - 1,0,0)
    GetFormatNumber = (number)
End Function


'//判断文件是否存在
Function FileExists(ByVal strSrcFilePath)
    res = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FileExists(strSrcFilePath)) = True  Then    res = True
    Set fso = Nothing
    FileExists = res
End Function



'//根据图廓编码打印
Function PrintImage(ByVal tkCode,ByVal imageName,ByRef ShapeHeight,ByRef ShapeWidth,ByRef daYZZ)
    'DeleteAllImage
    outputTitle = "成果图打印输出"
    projectFileName = SSProcess.GetProjectFileName
    filePath = SSProcess.GetSysPathName (4)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    'SSProcess.SetSelectCondition "SSObj_Type", "==", "LINE,AREA" 
    SSProcess.SetSelectCondition "SSObj_Code", "==", tkCode
    If InStr(msgInfo ,"地下室") > 0 Then           SSProcess.SetSelectCondition "[JianZWMC]", "like", "地下室"    Else   SSProcess.SetSelectCondition "[JianZWMC]", "not like", "地下室"
    SSProcess.SelectFilter
    count = SSProcess.GetSelGeoCount
    For i = 0 To count - 1
        objID = SSProcess.GetSelGeoValue(i,"SSObj_ID")
        projectName = SSProcess.GetSelGeoValue(i,"[JianZWMC]")
        If projectName = "" Then     projectName = SSProcess.GetSelGeoValue(i,"[XiangMMC]")
        scale = SSProcess.GetSelGeoValue(i,"[DaYBL]")
        leftDist = SSProcess.GetSelGeoValue(i,"[ZuoBJ]")
        upDist = SSProcess.GetSelGeoValue(i,"[ShangBJ]")
        daYZZ = SSProcess.GetSelGeoValue(i,"[DaYZZ]")
        If IsNumeric(scale) = False Then scale = 500
        If IsNumeric(leftDist) = False Then leftDist = 0
        If IsNumeric(upDist) = False Then upDist = 0
        If leftDist = 0 Then leftDist = 10
        If upDist = 0 Then upDist = 10
        height = SSProcess.GetSelGeoValue(i,"[ZhiK]")
        width = SSProcess.GetSelGeoValue(i,"[ZhiG]")
        H = 0
        W = 0
        'if isnumeric(width)=false or isnumeric(height)=false then 
        If InStr(daYZZ,"A4纵向") > 0 Then
            BaseHeith = 70
            BaseWidth = 70
            width = 210
            height = 297
            H = 24.5
            W = 18.8
        ElseIf InStr(daYZZ,"A4横向") > 0  Then
            BaseHeith = 105
            BaseWidth = 148.5
            width = 297
            height = 210
            H = 17.1
            W = 25.6
            ShapeWidth = 26.345 * W
            ShapeHeight = 26.345 * H
        ElseIf InStr(daYZZ,"A3纵向") > 0 Then
            BaseHeith = 210
            BaseWidth = 148.5
            width = 297
            height = 420
            H = 37.2
            W = 26.3
        ElseIf InStr(daYZZ,"A3横向") > 0  Then
            BaseHeith = 148.5
            BaseWidth = 210
            width = 420
            height = 297
            H = 24.9
            W = 35.2
        ElseIf InStr(daYZZ,"A2纵向") > 0  Then
            width = 420
            height = 594
        ElseIf InStr(daYZZ,"A2横向") > 0 Then
            width = 594
            height = 420
        ElseIf InStr(daYZZ,"A1纵向") > 0  Then
            width = 594
            height = 841
        ElseIf InStr(daYZZ,"A1横向") > 0 Then
            width = 841
            height = 594
        Else
            width = 297
            height = 210
            H = 16.2
            W = 22.9
        End If
        'end if
        If H = 0 Then H = 24.9
        If W = 0 Then W = 17.6
        ShapeHeight = 28.345 * H
        ShapeWidth = 28.345 * W
        xDist = 1
        yDist = 0.4
        SSProcess.GetObjectPoint objID,0,x0,y0,z0,ptype0,name0
        SSProcess.GetObjectPoint objID,1,x1,y1,z1,ptype1,name1
        SSProcess.GetObjectPoint objID,2,x2,y2,z2,ptype2,name2
        
        minX = x0 - 2 * Sqr((x0 - x1) ^ 2 + (y0 - y1) ^ 2) / BaseWidth
        minY = y0 - 4 * Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2) / BaseHeith
        maxX = x2 + 2 * Sqr((x0 - x1) ^ 2 + (y0 - y1) ^ 2) / BaseWidth
        maxY = y2 + 4 * Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2) / BaseHeith
        dpi = 300
        
        
        If count = 1 Then
            imagePath = filePath & projectName & imageName & ".bmp"
            SSProcess.WriteEpsIni outputTitle, imageName ,imagePath
        Else
            imagePath = filePath & projectName & imageName & i + 1 & ".bmp"
            SSProcess.WriteEpsIni outputTitle, imageName & i + 1 ,imagePath
        End If
        SSFunc.DrawToImage minX - 10, minY - 5, maxX + 10, maxY + 10, width & "X" & height, 400, imagePath
        
    Next
End Function


'//打印前先删除旧数据
Function DeleteAllImage
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = SSProcess.GetSysPathName (4)
    Dim filenames(10000)
    GetAllFiles filePath,"bmp",filecount,filenames
    For i = 0 To filecount - 1
        projectName = filenames(i)
        fso.DeleteFile projectName
    Next
    Set fso = Nothing
End Function


'//获取所有文件
Function GetAllFiles(ByRef pathname, ByRef fileExt, ByRef filecount, ByRef filenames())
    Dim fso, folder, file, files, subfolder,folder0, fcount
    Set fso = CreateObject("Scripting.FileSystemObject")
    If  fso.FolderExists(pathname) Then
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
        '查找子目录
        Set subfolder = folder.SubFolders
        For Each folder0 In subfolder
            GetAllFiles pathname & folder0.name & "\", fileExt, filecount, filenames
        Next
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
    
    adoRs.open sql  ,adoConnection,3,3
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


'//条件修改表格信息
Function ModifyTableInfo(ByVal strTableName,ByVal strFields,ByVal strValues,ByVal strAddCondition,ByVal strTableType,ByVal strGeoType)
    '判断表是否存在
    mdbName = SSProcess.GetProjectFileName
    'if  IsTableExits(mdbName,strTableName)=false then exit function 
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
    
    Set adoRs = CreateObject("ADODB.recordset")
    adoRs.Open sql,adoConnection,3,3
    rcount = adoRs.RecordCount
    'if instr(sql,3779)>0 then  addloginfo rcount&"=="&sql
    strFieldsList = Split(strFields,",")
    strValuesList = Split(strValues,",")
    If rcount > 0 Then
        For i = 0 To UBound(strFieldsList)
            'addloginfo strFieldsList(i)&"=="&strValuesList(i)
            adoRs("" & strFieldsList(i) & "") = strValuesList(i)
            adoRs.UpdateBatch
        Next
    End If
    adoRs.Close
    Set adoRs = Nothing
End Function

Function OutputBook(ByVal bookmark,ByVal fileName,ByVal code,ByVal SectionIndex)
    g_docObj.MoveToBookmark    bookmark
    
    '查找成果图edb文件
    accessName = SSProcess.GetProjectFileName
    filePath = Replace(accessName,".edb","") & "\"
    Dim edbList(10000)
    listCount = 0
    GetAllFiles filePath,"edb",listCount,edbList
    outEdbPath = ""
    For i = 0 To listCount - 1
        edbPath = edbList(i)
        If InStr(edbPath,fileName) > 0 And InStr(fileName,"bak") = 0 Then
            outEdbPath = edbPath
            Exit For
        End If
    Next
    If FileExists(outEdbPath) = False Then Exit Function
    DeleteAllImage
    '打开edb文件并按图廓范围打印wmf
    bRes = SSProcess.OpenDatabase (outEdbPath)
    If bRes = 1 Then
        PrintImage code,fileName, ShapeHeight, ShapeWidth,daYZZ
        SSProcess.CloseDatabase()
    End If
    '查找对应wmf文件并插入word报告中
    Dim imageList(10000)
    listCount = 0
    filePath = SSProcess.GetSysPathName (4)
    GetAllFiles filePath,"bmp",listCount1,imageList
    insertImageFile = ""
    For i = 0 To listCount1 - 1
        imageFile = imageList(i)
        name = GetFileName(imageFile)
        extensionName = GetFileExtensionName(imageFile)
        name = Replace(name,"." & extensionName,"")
        nameNumber = Replace(name,fileName,"")
        If InStr(name,fileName) > 0 Then
            insertImageFile = imageFile
            If FileExists( insertImageFile) = True Then
                'RES = g_docObj.InsertImage (insertImageFile,ShapeWidth,ShapeHeight,0)
                If daYZZ = "A4横向" Then
                    paperSize = 1
                    orientation = 2
                    pageWidth =  - 1
                    pageHeight =  - 1
                    H = 17.1
                    W = 24.2
                    width = 26.345 * W
                    height = 26.345 * H
                    '设置纸张的大小
                    leftMargin = 20'毫米
                    rightMargin = 20
                    topMargin = 7
                    bottomMargin = 7
                ElseIf daYZZ = "A4纵向" Then
                    paperSize = 1
                    orientation = 1
                    pageWidth =  - 1
                    pageHeight =  - 1
                    '设置宽高
                    H = 26.8
                    W = 21.8
                    width = 20.245 * W
                    height = 10.345 * H
                    '设置纸张的大小
                    leftMargin = 10'毫米
                    rightMargin = 10
                    topMargin = 10
                    bottomMargin = 10
                ElseIf daYZZ = "A3纵向" Then
                    paperSize = 0
                    orientation = 1
                    pageWidth =  - 1
                    pageHeight =  - 1
                    H = 37.2
                    W = 26.3
                    width = 28.345 * W
                    height = 28.345 * H
                    '设置纸张的大小
                    leftMargin = 10'毫米
                    rightMargin = 10
                    topMargin = 10
                    bottomMargin = 10
                ElseIf daYZZ = "A3横向" Then
                    paperSize = 0
                    orientation = 2
                    pageWidth =  - 1
                    pageHeight =  - 1
                    H = 25.8
                    W = 36.5
                    width = 28.345 * W
                    height = 28.345 * H
                    '设置纸张的大小
                    leftMargin = 10'毫米
                    rightMargin = 10
                    topMargin = 10
                    bottomMargin = 10
                ElseIf daYZZ = "500*500" Then
                    paperSize = 1
                    orientation = 1
                    pageWidth = 500
                    pageHeight = 500
                    '设置宽高
                    H = 45.04
                    W = 45.01
                    width = 30.245 * W
                    height = 28.345 * H
                    '设置纸张的大小
                    leftMargin = 10'毫米50
                    rightMargin = 10
                    topMargin = 10
                    bottomMargin = 10
                End If
                If SectionIndex <> "" Then
                    g_docObj.SectionPageSetup SectionIndex, paperSize, orientation, pageWidth, pageHeight, leftMargin, rightMargin, topMargin, bottomMargin
                End If
                '水平相对位置模式（wrapType非0时起作用） Margin = 0, Page = 1, Column = 2, Default = 2, Character = 3, LeftMargin = 4, RightMargin = 5, InsideMargin = 6, OutsideMargin = 7
                horzPos = 0
                left0 = 0
                '垂直位置相对模式（wrapType非0时起作用） Margin = 0,  TableDefault = 0,  Page = 1,  Paragraph = 2, TextFrameDefault = 2,  Line = 3,  TopMargin = 4,   BottomMargin = 5,  InsideMargin = 6,  OutsideMargin = 7
                vertPos = 0
                top0 = 3
                '图像环绕方式 Inline = 0 嵌入,    TopBottom = 1 上下,   Square = 2 四周,   None = 3 浮于文字上方,    Tight = 4 紧密,  Through = 5 穿越
                wrapType = 0
                '旋转角度
                rotation = 0
                g_docObj.InsertImageEx insertImageFile, horzPos, left0, vertPos, top0, ShapeWidth,ShapeHeight,  wrapType, rotation
            End If
        End If
    Next
    'if FileExists( insertImageFile)  =true then   g_docObj.SetCellImageEx2 tableIndex,  0, 0, 0,  insertImageFile, 0, 0, false
    
    
End Function

Function OutMap(ByVal tableIndex,ByVal MapName)
    mdbName = SSProcess.GetSysPathName (5)
    filePath = Replace(mdbName,".edb","") & "\" & MapName & "\"
    Dim imageList(10000)
    listCount = 0
    GetAllFiles filePath,"jpg",listCount,imageList
    If listCount > 1 Then
        For p = 1 To listCount - 1
            g_docObj. CloneTable tableIndex,1,0,False
        Next
    End If
    tableNum = tableIndex
    For i = 0 To listCount - 1
        imageFile = imageList(i)
        name = GetFileName(imageFile)
        extensionName = GetFileExtensionName(imageFile)
        name = Replace(name,"." & extensionName,"")
        'if instr(name,fileName)>0 then 
        insertImageFile = imageFile
        If FileExists( insertImageFile) = True Then
            g_docObj.SetCellImageEx tableNum, 0, 0, 0,  insertImageFile, 210 * 2, 297 * 2, False
            tableNum = tableNum + 1
        End If
        'end if 
    Next
    
End Function

'//获取文件名
Function GetFileName(ByVal strSrcFilePath)
    GetFileName = ""
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FileExists(strSrcFilePath)) = True  Then
        Set f = fso.getfile(strSrcFilePath)
        GetFileName = fso.GetFileName(f) '获取不含路径的文件名称,这就是输出
    End If
    Set f = Nothing
    Set fso = Nothing
End Function

'//获取文件后缀名
Function GetFileExtensionName(ByVal strSrcFilePath)
    GetFileExtensionName = ""
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FileExists(strSrcFilePath)) = True  Then
        Set f = fso.getfile(strSrcFilePath)
        GetFileExtensionName = fso.GetExtensionName(f)
    End If
    Set f = Nothing
    Set fso = Nothing
End Function