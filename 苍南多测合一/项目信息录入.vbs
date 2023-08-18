'ControlIDS：控件ID数组形式，GeoFields：对应地物的字段名数组形式，DefaultValues：默认值数组形式，AlternativeValues：备选项内容数组形式，MemoryValues：是否记忆上次传输的值数组形式，ControlCount ：控件个数
Dim ConIDS(300),GFields(300),DefValues(300),AltValues(300), MemValues(300),ConCount,ZDID
Dim dlgHandle, dlgHandle1, dlgHandle2, g_scriptHandle

mdbName = SSProcess.GetProjectFileName
'记录正在要填写的宗地ID
ZDID = ""
ZDCode = "9130223"
Sub OnInitFreeScript(scriptHandle)
    mapHandle = SSProject.GetActiveMap
    mapType = SSProject.GetMapInfo(mapHandle, "MapType")
    If mapType <> 2 Then
        MsgBox "本功能只支持在地形图窗口执行！"
        Exit Sub
    End If
    
    g_scriptHandle = scriptHandle
    dlgHandle = SSProcess.CreateFreeScriptDlg(scriptHandle, 1)
    Rstate = GetZDID()  '获取要填属性的宗地ID,
    If Rstate = 0 Then
        MsgBox "图上无宗地面，无需使用该功能X"
        SSProcess.CloseScriptDlg
        Exit Sub
    End If
    Dim strs(100)
    IniDlgParameter()  '获取对话框Ini文本设置的参数
    mode = 1'0 有模 1 无模
    title = "项目信息录入"
    dlgTemplateName = "项目信息录入主对话框"
    dlgWidth = 750
    dlgHeight = 610
    colCount = 0
    titleWidth = 0
    valueWidth = 0
    For i = 0 To ConCount - 1
        If MemValues(i) = "1" Then
            'DefValues(i) = SSProcess.ReadEpsIni ("ZhuHaiInfoPut", "CLInfo_" & ConIDS(i) , "" )
        End If
        If DefValues(i) = "Date" Then DefValues(i) = Date
        If ZDID <> "" Then
            SSFunc.ScanString GFields(i),",",strs,scount
            Valuestr = ""
            For j = 0 To scount - 1
                If j = 0 Then Valuestr = SSProcess.GetObjectAttr (ZDID,"[" & strs(j) & "]")
                If j > 0 And Valuestr <> "" Then Valuestr = Valuestr & "、" & SSProcess.GetObjectAttr (ZDID,"[" & strs(j) & "]")
            Next
            If Valuestr = "" Then
                SSProcess.AddInputParameter_ex dlgHandle,ConIDS(i),DefValues(i),0,AltValues(i),""
            Else
                SSProcess.AddInputParameter_ex dlgHandle,ConIDS(i),Valuestr,0,AltValues(i),""   '获取当前已有值填充
            End If
        Else
            SSProcess.AddInputParameter_ex dlgHandle,ConIDS(i),DefValues(i),0,AltValues(i),""
        End If
    Next
    SSProcess.AddInputParameter_ex dlgHandle, "{标签}", "项目信息", 0, "项目信息,测绘责任人信息", ""
    SSProcess.ShowFreeScriptDlg dlgHandle, title, dlgTemplateName, dlgWidth, dlgHeight, colCount, titleWidth, valueWidth, dockMode
    OnTabCtrlSelChange "", 0, "{标签}", "项目信息"
    OnTabCtrlSelChange "", 0, "{标签}", "测绘责任人信息"
    OnTabCtrlSelChange "", 0, "{标签}", "项目信息"
End Sub

'获取要编辑属性的宗地ID，优先取选择集的宗地，若无选择，全图找宗地，找到唯一个宗地，提取ZDID，找到多个宗地，则提醒人工选择
Function GetZDID()
    '判断选择集是否有宗地
    GetZDID = 1
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.UpdateSysSelection 0
    geoCount = SSProcess.GetSelGeoCount
    LSID = ""
    ZDCount = 0
    For i = 0 To geoCount - 1
        geoCode = SSProcess.GetSelGeoValue (i, "SSObj_Code")
        If  InStr("," & ZDCode & ",","," & geoCode & ",") > 0 Then
            LSID = SSProcess.GetSelGeoValue (i, "SSObj_ID")
            ZDCount = ZDCount + 1
        End If
    Next
    If ZDCount = 1 Then
        ZDID = LSID
    Else
        SSProcess.ClearSelection
        SSProcess.ClearSelectCondition
        SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
        SSProcess.SetSelectCondition "SSObj_Code", "=", ZDCode
        SSProcess.SelectFilter
        geoCount = SSProcess.GetSelGeoCount
        If geoCount = 0 Then
            GetZDID = 0
        ElseIf geoCount = 1 Then
            ZDID = SSProcess.GetSelGeoValue (0, "SSObj_ID")
        Else
            MsgBox "图上有多个宗地，且当前选择集无选择宗地，请在弹出对话框后，选择宗地再填写属性X"
        End If
    End If
End Function

Sub OnExitFreeScript()
    '添加代码
End Sub

Sub OnOK()
    '添加代码
End Sub

'编辑项失去焦点
'Function OnEditKillFocus( tableName, objectID, fieldName, fieldValue )
'添加代码
'End Function

Sub OnCancel()
    If dlgHandle1 <> 0 Then SSProcess.CloseChildFreeScriptDlg dlgHandle1
    If dlgHandle2 <> 0 Then SSProcess.CloseChildFreeScriptDlg dlgHandle2
End Sub

'标签控件选择项改变
Function OnTabCtrlSelChange( tableName, objectID, fieldName, fieldValue )
    '添加代码
    Dim strs(100), mapnumberInfo(100000)
    If fieldName = "{标签}" Then
        '先隐藏对话框
        If dlgHandle1 <> 0 Then SSProcess.ShowScriptDlgWindow dlgHandle1, 0
        If dlgHandle2 <> 0 Then SSProcess.ShowScriptDlgWindow dlgHandle2, 0
        If fieldValue = "项目信息" Then
            If dlgHandle1 = 0 Then
                dlgHandle1 = SSProcess.CreateFreeScriptDlg(g_scriptHandle, 1)
                dlgTemplateName = "项目信息录入"
                dockCtrlID = "{停靠控件}"
                SSProcess.DockScriptDlg dlgHandle1, dlgTemplateName, dlgHandle, dockCtrlID
                For i = 0 To ConCount - 1
                    If MemValues(i) = "1" Then
                        'DefValues(i) = SSProcess.ReadEpsIni ("ZhuHaiInfoPut", "CLInfo_" & ConIDS(i) , "" )
                    End If
                    If DefValues(i) = "Date" Then DefValues(i) = Date
                    If ZDID <> "" Then
                        SSFunc.ScanString GFields(i),",",strs,scount
                        Valuestr = ""
                        For j = 0 To scount - 1
                            If j = 0 Then Valuestr = SSProcess.GetObjectAttr (ZDID,"[" & strs(j) & "]")
                            If j > 0 And Valuestr <> "" Then Valuestr = Valuestr & "、" & SSProcess.GetObjectAttr (ZDID,"[" & strs(j) & "]")
                        Next
                        
                        If ConIDS(i) = "[JianZXS]" Then  SSProcess.SetScriptDlgCellOptions_ex dlgHandle1, "[JianZXS]",AltValues(i)
                        If ConIDS(i) = "[JianZJG]" Then  SSProcess.SetScriptDlgCellOptions_ex dlgHandle1, "[JianZJG]",AltValues(i)
                        '数据字典
                        If Valuestr = ""  Then
                            '图幅号
                            SSProcess.SetScriptDlgCellValue_ex dlgHandle1,ConIDS(i),DefValues(i)
                        Else
                            SSProcess.SetScriptDlgCellValue_ex dlgHandle1,ConIDS(i),Valuestr      '获取当前已有值填充
                        End If
                    Else
                        SSProcess.SetScriptDlgCellValue_ex dlgHandle1,ConIDS(i),DefValues(i)
                    End If
                Next
            Else
                SSProcess.ShowScriptDlgWindow dlgHandle1, 1
            End If
        End If
        If fieldValue = "测绘责任人信息" Then
            If dlgHandle2 = 0 Then
                dlgHandle2 = SSProcess.CreateFreeScriptDlg(g_scriptHandle, 1)
                dlgTemplateName = "项目信息录入-责任人"
                dockCtrlID = "{停靠控件}"
                SSProcess.DockScriptDlg dlgHandle2, dlgTemplateName, dlgHandle, dockCtrlID
                For i = 0 To ConCount - 1
                    If MemValues(i) = "1" Then
                        'DefValues(i) = SSProcess.ReadEpsIni ("ZhuHaiInfoPut", "CLInfo_" & ConIDS(i) , "" )
                    End If
                    If DefValues(i) = "Date" Then DefValues(i) = Date
                    If ZDID <> "" Then
                        SSFunc.ScanString GFields(i),",",strs,scount
                        Valuestr = ""
                        For j = 0 To scount - 1
                            If j = 0 Then Valuestr = SSProcess.GetObjectAttr (ZDID,"[" & strs(j) & "]")
                            If j > 0 And Valuestr <> "" Then Valuestr = Valuestr & "、" & SSProcess.GetObjectAttr (ZDID,"[" & strs(j) & "]")
                        Next
                        
                        If Valuestr = "" Then
                            SSProcess.SetScriptDlgCellValue_ex dlgHandle2,ConIDS(i),DefValues(i)
                        Else
                            SSProcess.SetScriptDlgCellValue_ex dlgHandle2,ConIDS(i),Valuestr      '获取当前已有值填充
                        End If
                    Else
                        SSProcess.SetScriptDlgCellValue_ex dlgHandle2,ConIDS(i),DefValues(i)
                    End If
                Next
            Else
                SSProcess.ShowScriptDlgWindow dlgHandle2, 1
            End If
        End If
    End If
End Function


Function GetXMFZR(strs_,strs1_)
    Dim  strs(1000)
    returnStr = ""
    str = ""
    fileName = SSProcess.GetSysPathName(7) & "测绘责任人.txt"
    Dim fso, ts, chLine
    Set fso = CreateObject("Scripting.FileSystemObject")'主要用于创建组件、应用对象或脚本对象的实例。
    Set ts = fso.OpenTextFile(fileName, 1)' 1、以只读方式打开文件，不能写这个文件；2、以写方式打开文件；8、打开文件并从文件末尾开始写。
    Do While Not ts.AtEndOfStream'如果位于文件末，则返回 True；否则返回 False。
        chLine = ts.ReadLine'读取一整行
        chLine = Trim(chLine)
        If str = ""  Then str = chLine
    Else str = str & "," & chLine
    Loop
    SSProcess.SaveBufferObjToDatabase
    ts.Close
    
    'str = "赵卫民、3300300039,李新疆、3310300635,金立枢、3300300032"
    SSFunc.ScanString str,",",strs,strs1count
    ResVal_Dlg = SSFunc.SelectListAttr("选择列表", "待选数据列表", "选中数据列表", strs, strs1count)
    strs_ = ""
    strs1_ = ""
    If strs1count = 0 Then
        Exit Function
    End If
    If ResVal_Dlg = 1 Then
        If strs1count = 1 Then
            arinfo = Split(strs(0), "、")
            strs_ = arinfo(0)
            strs1_ = arinfo(1)
        Else
            MsgBox "请选择！"
            Exit Function
        End If
    End If
End Function

'所有按钮的实现在这里
Function OnButtonClick( tableName, objectID, fieldName, fieldValue )
    '添加代码
    If fieldName = "[XMFZR]" Then
        GetXMFZR  strs_,strs1 _
        If strs_ <> ""  Then SSProcess.SetScriptDlgCellValue_ex dlgHandle2 ,"[XiangMFZR]", strs _
        If strs1_ <> "" Then SSProcess.SetScriptDlgCellValue_ex dlgHandle2 ,"[XiangMFZRZSH]",strs1 _
    End If
    
    If fieldName = "[CLY]" Then
        GetXMFZR  strs_,strs1 _
        If strs_ <> ""  Then SSProcess.SetScriptDlgCellValue_ex dlgHandle2 ,"[CeLY]", strs _
        If strs1_ <> "" Then SSProcess.SetScriptDlgCellValue_ex dlgHandle2 ,"[CeLYZSH]",strs1 _
    End If
    If fieldName = "[ZTY]" Then
        GetXMFZR  strs_,strs1 _
        If strs_ <> ""  Then SSProcess.SetScriptDlgCellValue_ex dlgHandle2 ,"[ZhiTY]", strs _
        If strs1_ <> "" Then SSProcess.SetScriptDlgCellValue_ex dlgHandle2 ,"[ZhiTYZSH]",strs1 _
    End If
    If fieldName = "[JCY]" Then
        GetXMFZR  strs_,strs1 _
        If strs_ <> ""  Then SSProcess.SetScriptDlgCellValue_ex dlgHandle2 ,"[JianCY]", strs _
        If strs1_ <> "" Then SSProcess.SetScriptDlgCellValue_ex dlgHandle2 ,"[JianCYZSH]",strs1 _
    End If
    If fieldName = "[SHY]" Then
        GetXMFZR  strs_,strs1 _
        If strs_ <> ""  Then SSProcess.SetScriptDlgCellValue_ex dlgHandle2 ,"[ShenHR]", strs _
        If strs1_ <> "" Then SSProcess.SetScriptDlgCellValue_ex dlgHandle2 ,"[ShenHRZSH]",strs1 _
    End If
    
    If fieldName = "[CHZRR]"  Then
        
        fileName = SSProcess.GetSysPathName(7) & "测绘责任人.txt"
        Set oShell = CreateObject ("Wscript.shell")
        oShell.run   fileName
        Set oshell = Nothing
    End If
    
    Dim strs(10),strs1(10)
    mark = 1
    If fieldName = "[SAVE]" Then'点击保存按钮
        lhcybh = SSProcess.GetScriptDlgCellValue_ex (dlgHandle,"[HeTBH]")
        xmmc = SSProcess.GetScriptDlgCellValue_ex (dlgHandle,"[XiangMMC]")
        ywlx = SSProcess.GetScriptDlgCellValue_ex (dlgHandle,"[YeWLX]")
        If lhcybh = ""  Or  xmmc = ""  Or    ywlx = ""     Then mark = 0
        MsgBox "请填写项目名称、合同号、业务类型！"
        MsgBox "保存失败！"
        Exit Function
        For i = 0 To ConCount - 1
            lhcybh = SSProcess.GetScriptDlgCellValue_ex (dlgHandle,"[HeTBH]")
            xmmc = SSProcess.GetScriptDlgCellValue_ex (dlgHandle,"[XiangMMC]")
            If lhcybh = ""  And xmmc = ""  Then mark = 0
            MsgBox "请填写项目名称和合同号！"
            Exit For
            If i < 5 Then value = SSProcess.GetScriptDlgCellValue_ex (dlgHandle,ConIDS(i))
            If i > 4 And i < 21 Then value = SSProcess.GetScriptDlgCellValue_ex (dlgHandle1,ConIDS(i))
            If i > 20 Then value = SSProcess.GetScriptDlgCellValue_ex (dlgHandle2,ConIDS(i))
            If ConIDS(i) = "[ZhiTRQ]"   Or  ConIDS(i) = "[CeLRQ]"  Or  ConIDS(i) = "[KSSCSJ]"  Or  ConIDS(i) = "[JSSCSJ]"  Or  ConIDS(i) = "[JianCRQ]" Or  ConIDS(i) = "[ShenHRQ]"  Then value = FormatDateTime(value,1)
            SSFunc.ScanString GFields(i),",",strs,scount                'GFields(i)  对应地物字段
            SSFunc.ScanString value,"、",strs1,scount1
            If scount = 1 Then
                SSProcess.SetObjectAttr ZDID,"[" & GFields(i) & "]",value
            Else
                If scount1 > 1 Then
                    For j = 0 To scount - 1
                        SSProcess.SetObjectAttr ZDID,"[" & strs(j) & "]",strs1(j)
                    Next
                End If
            End If
        Next
        Dim zdarRecordList(),zdRecordListCount,zrzarRecordList(),zrzRecordListCount
        SSProcess.OpenAccessMdb mdbName
        Dim LCarRecordList(),LCRecordListCount
        sql = "Select ZDGUID,CeLY,CeLRQ,ZhiTY,ZhiTRQ,JianCY,JianCRQ,ShenHR,ShenHRQ,DiaoCR,DiaoCRQ From ZD_宗地基本信息属性表 INNER JOIN GeoAreaTB ON ZD_宗地基本信息属性表.ID=GeoAreaTB.ID WHERE (GeoAreaTB.Mark mod 2)<>0 And ZD_宗地基本信息属性表.ID = " & ZDID
        GetSQLRecordAll mdbName,sql,zdarRecordList,zdRecordListCount
        
        If zdRecordListCount = 1  Then
            artempzd = Split(zdarRecordList(0),",")
            sql = "select FC_自然幢信息属性表.ID from FC_自然幢信息属性表 inner join GeoAreaTB ON FC_自然幢信息属性表.ID = GeoAreaTB.ID Where (GeoAreaTB.Mark mod 2)<>0 and FC_自然幢信息属性表.ZDGUID =" & artempzd(0)
            GetSQLRecordAll mdbName,sql,zrzarRecordList,zrzRecordListCount
            
            For j = 0 To zrzRecordListCount - 1
                
                SSProcess.SetObjectAttr zrzarRecordList(j),"[CeLY],[CeLRQ],[ZhiTY],[HuiTRQ],[JianCY],[JianCRQ],[ShenHR],[ShenHRQ],[DiaoCR],[DiaoCRQ]",artempzd(1) & "," & artempzd(2) & "," & artempzd(3) & "," & artempzd(4) & "," & artempzd(5) & "," & artempzd(6) & "," & artempzd(7) & "," & artempzd(8) & "," & artempzd(9) & "," & artempzd(10)
                
            Next
        End If
        SSProcess.CloseAccessMdb mdbName
        
        KSSCSJ = SSProcess.GetScriptDlgCellValue_ex (dlgHandle1,"[KSSCSJ]")
        scsj = Left(KSSCSJ,4)
        SSProcess.SetObjectAttr ZDID,"[ShiCNF]",ZDID
        
        MsgBox "保存成功！"
    End If
    
    If fieldName = "[CLOSE]" Then'点击关闭按钮
        SSProcess.CloseFreeScriptDlg dlgHandle
        Exit Function
    End If
End Function

'获取指定sql语句下的  搜索结果
Function GetSQLRecordAll(ByRef mdbName, ByRef StrSqlStatement, ByRef arSQLRecord(), ByRef iRecordCount)
    SSProcess.OpenAccessMDB mdbName
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
            arSQLRecord(iRecordCount) = values                                        '查询记录
            iRecordCount = iRecordCount + 1                                                    '查询记录数
            '移动记录游标
            SSProcess.AccessMoveNext mdbName, sql
        WEnd
    End If
    '关闭记录集
    SSProcess.CloseAccessRecordset mdbName, sql
    SSProcess.CloseAccessMDB mdbName
End Function
'选择下拉框改变
'Function OnComboBoxSelChange( tableName, objectID, fieldName, fieldValue )
'添加代码
'End Function

'窗口大小发生改变
Function OnSize( dHandle, cx, cy )
    If dHandle = dlgHandle Then
        ctrlSize = SSProcess.GetScriptDlgCellRect_ex (dlgHandle,"{编辑框}")
        Dim vArray(4), nCount
        SSFunc.ScanString ctrlSize, ",", vArray, nCount
        rectleft = CLng(vArray(0))
        recttop = CLng(vArray(1))
        rectright = CLng(vArray(2))
        rectbottom = CLng(vArray(3))
        rectright = cx - 10
        rectbottom = cy - 10
        SSProcess.SetScriptDlgCellRect_ex dlgHandle, "{编辑框}", rectLeft, rectTop, rectRight, rectBottom
    End If
End Function


Function IniDlgParameter()
    Ininame = "BDC_项目信息录入-1.ini"
    ReadIniInfo Ininame, ConIDS,GFields,DefValues,AltValues, MemValues,ConCount
End Function


'解析DLG对应的ini设置记录
'入参，Ininame：ini文件的名称
'入参返回：ControlIDS：控件ID数组形式，GeoFields：对应地物的字段名数组形式，DefaultValues：默认值数组形式，AlternativeValues：备选项内容数组形式，MemoryValues：是否记忆上次传输的值数组形式，ControlCount ：控件个数
Function ReadIniInfo(ByVal Ininame,ByRef  ControlIDS(),ByRef GeoFields(),ByRef DefaultValues(),ByRef AlternativeValues(),ByRef MemoryValues(),ByRef ControlCount)
    ControlCount = 0
    
    TemplateFileName = SSProcess.GetTemplateFileName
    Ininame_ = Left(TemplateFileName,Len(TemplateFileName) - 4) & "\" & Ininame
    MsgBox Ininame _
    Dim strs(50)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(Ininame_, 1, False)
    While (f.atEndOfStream = False)
        strLine = f.ReadLine()
        If strLine <> "" And Len(strLine) > 4 Then
            If Left(strLine,2) <> "//" Then
                SSFunc.ScanString strLine,"|",strs,scount
                If scount = 5 Then
                    ControlIDS(ControlCount) = strs(0)
                    GeoFields(ControlCount) = strs(1)
                    DefaultValues(ControlCount) = strs(2)
                    AlternativeValues(ControlCount) = strs(3)
                    MemoryValues(ControlCount) = strs(4)
                    ControlCount = ControlCount + 1
                End If
            Else
                ReadIniInfo = strLine
            End If
        End If
    WEnd
    f.Close()
    Set f = Nothing
    Set fso = Nothing
End Function



