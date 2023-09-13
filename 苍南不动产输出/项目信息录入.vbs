'=======================全局变量定义====================
Dim ConIDS(600),GFields(600),DefValues(600),AltValues(600), MemValues(600),ConCount
ZDCode = "9130223"
ZRZCode = "9210123"

CURZDID = ""
CURZRZID = ""
'记录一个宗地下的所有自然幢信息和逻辑幢信息
'ZRZInfos二维数组，记录一个宗地下的所有自然幢的{ID,ZRZH,CHZT,ZL,XQMC,GHZcsPZ,GHDxcsPZ,GHGHDxcsPZPZ,LJZHLB}9个属性，数组定义认为一个宗地下不会超过100个自然幢
Dim ZRZInfos(100,27) ,ZRZCOUNT
ZRZFields = Array("ID","ZRZH","ZRZSXH","CHZT","JZWMC","FWJG","JGRQ","XMMC","ZL","QiuH","ChanB","FWXZ","SFZYYT","GHYT","FWYT","ZCS","DSCS","DXCS","ZZDMJ","DQTGS","NQTGS","XQTGS","BQTGS","LCFZXX","GuiHXKZBH","GHJZWMC","JZWMCGUID")
ZRZFieldTypes = Array(0,1,1,0,1,0,1,1,1,1,0,0,0,0,0,0,0,0,0,1,1,1,1,1,1,1,1)  '0表示数值型字段，1表示文本型字段
ZRZFieldcount = 27
ZRZFieldstr = "ID,ZRZH,ZRZSXH,CHZT,JZWMC,FWJG,JGRQ,XMMC,ZL,QiuH,ChanB,FWXZ,SFZYYT,GHYT,FWYT,ZCS,DSCS,DXCS,ZZDMJ,DQTGS,NQTGS,XQTGS,BQTGS,LCFZXX,GuiHXKZBH,GHJZWMC,JZWMCGUID"
ZRZConIDS = Array("","[ZRZH]","[ZRZSXH]","[CHZT]","[JZWMC]","[FWJG]","[JGRQ]","[XMMC]","[ZL]","[QiuH]","[ChanB]","[FWXZ]","[SFZYYT]","[GHYT]","[FWYT]","[ZCS]","[ZRZDSCS]","[ZRZDXCS]","[ZZDMJ]","[QD]","[QN]","[QX]","[QB]","[LCFZXX]","[GuiHXKZBH]","[GHJZWMC]","[JZWMCGUID]")

mdbName = SSProcess.GetProjectFileName

Dim dlgHandle
'获取指定宗地下的所有自然幢和逻辑幢信息

'解析DLG对应的ini设置记录
'入参，Ininame：ini文件的名称
'入参返回：ControlIDS：控件ID数组形式，GeoFields：对应地物的字段名数组形式，DefaultValues：默认值数组形式，AlternativeValues：备选项内容数组形式，MemoryValues：是否记忆上次传输的值数组形式，ControlCount ：控件个数
Function ReadIniInfo(ByVal Ininame,ByRef  ControlIDS(),ByRef GeoFields(),ByRef DefaultValues(),ByRef AlternativeValues(),ByRef MemoryValues(),ByRef ControlCount)
    ControlCount = 0
    TemplateFileName = SSProcess.GetTemplateFileName
    Ininame_ = Left(TemplateFileName,Len(TemplateFileName) - 4) & "\" & Ininame
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

Function IniDlgParameter()
    Ininame = "FC_自然幢与逻辑信息编辑-1.ini"
    ReadIniInfo Ininame, ConIDS,GFields,DefValues,AltValues, MemValues,ConCount
End Function


Function GetAllZRZAndLJZInfo()
    If CURZDID <> "" Then
        Dim strs(1000)
        erase ZRZInfos
        ZRZCOUNT = 0
        SetZDandZRZGUID()
        ZDGUID = SSProcess.GetObjectAttr(CURZDID,"[ZDGUID]")
        SSProcess.OpenAccessMdb mdbName
        sql = "select FC_自然幢信息属性表." & ZRZFieldstr & " from FC_自然幢信息属性表 inner join GeoAreaTB ON FC_自然幢信息属性表.ID = GeoAreaTB.ID Where (GeoAreaTB.Mark mod 2)<>0 and FC_自然幢信息属性表.ZDGUID = '" & ZDGUID & "' order by FC_自然幢信息属性表.ZRZSXH"
        SSProcess.OpenAccessRecordset mdbName, sql
        While SSProcess.AccessIsEOF (mdbName, sql) = False
            SSProcess.GetAccessRecord mdbName, sql, fields, values
            SSFunc.ScanString values,",",strs,count
            For i = 0 To ZRZFieldcount - 1
                ZRZInfos(ZRZCOUNT,i) = strs(i)
            Next
            ZRZCOUNT = ZRZCOUNT + 1
            SSProcess.AccessMoveNext mdbName, sql
        WEnd
        SSProcess.CloseAccessRecordset mdbName, sql
        SSProcess.CloseAccessMdb mdbName
        
    End If
End Function
'设置一个宗地和宗地下的自然幢的ZDGUID和ZRZGUID
Function SetZDandZRZGUID()
    SSProcess.MapMethod "clearattrbuffer", "FC_自然幢信息属性表"
    If CURZDID <> "" Then
        ZDGUID = SSProcess.GetObjectAttr(CURZDID,"[ZDGUID]")
        If ZDGUID = "" Or ZDGUID = "*" Or ZDGUID = "{00000000-0000-0000-0000-000000000000}" Then
            ZDGUID = SSProcess.GetObjectAttr(CURZDID,"[FeatureGUID]")
            SSProcess.SetObjectAttr CURZDID,"[ZDGUID]",ZDGUID
        End If
        ids = SSProcess.SearchInnerObjIDs (CURZDID, 2, ZRZCode, 1)
        Dim strs(500),count
        SSFunc.ScanString ids,",",strs,count
        For i = 0 To count - 1
            ZRZZDGUID = SSProcess.GetObjectAttr(strs(i),"[ZDGUID]")
            If ZRZZDGUID <> ZDGUID Then SSProcess.SetObjectAttr strs(i),"[ZDGUID]",ZDGUID
            ZRZGUID = SSProcess.GetObjectAttr(strs(i),"[ZRZGUID]")
            If ZRZGUID = "" Or ZRZGUID = "*" Or ZRZGUID = "{00000000-0000-0000-0000-000000000000}" Then
                ZRZGUID = SSProcess.GetObjectAttr(strs(i),"[FeatureGUID]")
                SSProcess.SetObjectAttr strs(i),"[ZRZGUID]",ZRZGUID
            End If
            LJZHLB = SSProcess.GetObjectAttr(strs(i),"[LJZHLB]")
            If LJZHLB = ""  Or LJZHLB = "*"  Then
                SSProcess.SetObjectAttr strs(i),"[LJZHLB]","1"
            End If
        Next
    End If
    SSProcess.MapMethod "clearattrbuffer", "FC_自然幢信息属性表"
End Function

Function SaveCurIDInfos(ByVal zrzsql,ByVal ljzsql)
    SSProcess.MapMethod "clearattrbuffer", "FC_自然幢信息属性表"
    SSProcess.MapMethod "clearattrbuffer", "FC_逻辑幢信息表"
    SSProcess.OpenAccessMdb mdbName
    If zrzsql <> "" Then SSProcess.ExecuteAccessSql mdbName, zrzsql
    If ljzsql <> "" Then SSProcess.ExecuteAccessSql mdbName, ljzsql
    SSProcess.CloseAccessMdb mdbName
    SSProcess.MapMethod "clearattrbuffer", "FC_自然幢信息属性表"
    SSProcess.MapMethod "clearattrbuffer", "FC_逻辑幢信息表"
End Function

Sub OnExitFreeScript()
    '添加代码
End Sub

Sub OnInitFreeScript(scriptHandle)
    'SSProcess.MapCallBackFunction1 31,"SCRIPT:.\\系统消息\\自然幢逻辑幢号列表赋值.vbs",0
    Dim Newlcxx(10)
    dlgHandle = SSProcess.CreateFreeScriptDlg(scriptHandle, 1)
    Rstate = GetZDID()  '获取要填属性的自然幢ID,
    If Rstate = 0 Then
        MsgBox "图上无宗地面，无需使用该功能X"
        SSProcess.CloseFreeScriptDlg dlgHandle
        Exit Sub
    End If
    
    If Rstate = 2 Then
        SSProcess.CloseFreeScriptDlg dlgHandle
        Exit Sub
    End If
    GetAllZRZAndLJZInfo()
    If ZRZCOUNT = 0 Then
        MsgBox "图上宗地面下无自然幢，无需使用该功能X"
        SSProcess.CloseFreeScriptDlg dlgHandle
        Exit Sub
    End If
    ZRZHS = ""
    LJZHS = ""
    ZRZH = ""
    LJZH = ""
    For i = 0 To ZRZCOUNT - 1
        If i = 0 Then
            ZRZHX = ZRZInfos(i,1)
            ZRZH = ZRZInfos(i,2) & "(" & ZRZInfos(i,1) & ")" & "(ID:" & ZRZInfos(i,0) & ")"
            ZRZHS = ZRZH
            CURZRZID = ZRZInfos(i,0)
        Else
            ZRZHS = ZRZHS & "," & ZRZInfos(i,2) & "(" & ZRZInfos(i,1) & ")" & "(ID:" & ZRZInfos(i,0) & ")"
            ZRZHX = ZRZHX
        End If
    Next
    XMMC = SSProcess.GetObjectAttr (CURZDID,"[XiangMMC]")
    ZL = SSProcess.GetObjectAttr (CURZDID,"[ZL]")
    mode = 1'0 有模 1 无模
    title = "自然幢与逻辑信息编辑"
    dlgTemplateName = "FC_自然幢与逻辑信息编辑-1"  ' "单位面属性表"
    dlgWidth = 820
    dlgHeight = 500
    colCount = 0
    titleWidth = 0
    valueWidth = 0
    IniDlgParameter()  '获取对话框Ini文本设置的参数
    Dim strs(100)
    For i = 0 To ConCount - 1
        If CURZRZID <> "" Then
            SSFunc.ScanString GFields(i),",",strs,scount
            Valuestr = ""
            For j = 0 To scount - 1
                If j = 0 Then Valuestr = SSProcess.GetObjectAttr (CURZRZID,"[" & strs(j) & "]")
                If j > 0 And Valuestr <> "" Then Valuestr = Valuestr & "、" & SSProcess.GetObjectAttr (CURZRZID,"[" & strs(j) & "]")
            Next
            
            'If ConIDS(i) = "[ZRZH]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[ZRZH]",AltValues(i)
            'If ConIDS(i) = "[CHZT]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[CHZT]",AltValues(i)
            'If ConIDS(i) = "[FWJG]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[FWJG]",AltValues(i)
            'If ConIDS(i) = "[ChanB]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[ChanB]",AltValues(i)
            'If ConIDS(i) = "[FWXZ]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[FWXZ]",AltValues(i)
            'If ConIDS(i) = "[SFZYYT]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[SFZYYT]",AltValues(i)
            'If ConIDS(i) = "[GHYT]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[GHYT]",AltValues(i)
            'If ConIDS(i) = "[FWYT]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[FWYT]",AltValues(i)
            'If ConIDS(i) = "[QD]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[QD]",AltValues(i)
            'If ConIDS(i) = "[QN]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[QN]",AltValues(i)
            'If ConIDS(i) = "[QX]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[QX]",AltValues(i)
            'If ConIDS(i) = "[QB]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[QB]",AltValues(i)
            GetAllGuiHXKZBH allghxkz
            allghxkz = "," & allghxkz
            If Valuestr = "" Then
                If ConIDS(i) = "[ghxkzbh]"  Then
                    SSProcess.AddInputParameter_ex dlgHandle,ConIDS(i),"",0,allghxkz,""
                ElseIf  ConIDS(i) = "[XMMC]"  Then
                    SSProcess.AddInputParameter_ex dlgHandle,ConIDS(i),XMMC,0,AltValues(i),""
                ElseIf  ConIDS(i) = "[ZL]"  Then
                    SSProcess.AddInputParameter_ex dlgHandle,ConIDS(i),ZL,0,AltValues(i),""
                Else
                    SSProcess.AddInputParameter_ex dlgHandle,ConIDS(i),DefValues(i),0,AltValues(i),""
                End If
            Else
                If ConIDS(i) = "[ghxkzbh]"  Then
                    GHXKZBH = Valuestr
                    SSProcess.AddInputParameter_ex dlgHandle,ConIDS(i),Valuestr,0,allghxkz,""   '获取当前已有值填充
                ElseIf  ConIDS(i) = "[GHJZWMC]"  Then
                    GetAllDH GHXKZBH,AllDH
                    SSProcess.AddInputParameter_ex dlgHandle,ConIDS(i),Valuestr,0,"," & AllDH,""   '获取当前已有值填充  
                Else
                    SSProcess.AddInputParameter_ex dlgHandle,ConIDS(i),Valuestr,0,AltValues(i),""   '获取当前已有值填充
                End If
            End If
        Else
            SSProcess.AddInputParameter_ex dlgHandle,ConIDS(i),DefValues(i),0,AltValues(i),""
        End If
    Next
    
    SSProcess.ShowFreeScriptDlg dlgHandle, title, dlgTemplateName, dlgWidth, dlgHeight, colCount, titleWidth, valueWidth, dockMode
    'SSProcess.ShowScriptUserDefDlgEx  mode, title, dlgTemplateName, dlgWidth, dlgHeight, colCount, titleWidth, valueWidth
    If ZRZH <> "" Then
        SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[自然幢号列表]",ZRZHS
        SSProcess.SetScriptDlgCellValue_ex dlgHandle, "[自然幢号列表]",ZRZH
    End If
End Sub

'下拉框选择项改变
Function OnComboBoxSelChange( tableName, objectID, fieldName, fieldValue )
    Dim AllDH,ghjzwbs
    If fieldName = "[ghxkzbh]" Then
        
        GHJZWMC = SSProcess.GetScriptDlgCellValue_ex (dlgHandle,"[GHJZWMC]")
        GHXKZBH = SSProcess.GetScriptDlgCellValue_ex (dlgHandle,"[ghxkzbh]")
        GetAllDH GHXKZBH,AllDH
        If GHXKZBH = ""  Then
            SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[GHJZWMC]",""
            SSProcess.SetScriptDlgCellValue_ex dlgHandle, "[JZWMCGUID]",""
        End If
        If GHJZWMC <> "" Then
            If  Replace("," & AllDH & ",","," & GHJZWMC & ",","") <> "," & AllDH & ","  Then
                Getjzwmcguid GHXKZBH,GHJZWMC,jzwmcguid
                SSProcess.SetScriptDlgCellValue_ex dlgHandle, "[JZWMCGUID]",jzwmcguid
            End If
        End If
        
        SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[GHJZWMC]","," & AllDH
    ElseIf fieldName = "[GHJZWMC]" Then
        GHJZWMC = SSProcess.GetScriptDlgCellValue_ex (dlgHandle,"[GHJZWMC]")
        GHXKZBH = SSProcess.GetScriptDlgCellValue_ex (dlgHandle,"[ghxkzbh]")
        Getjzwmcguid GHXKZBH,GHJZWMC,jzwmcguid
        SSProcess.SetScriptDlgCellValue_ex dlgHandle, "[JZWMCGUID]",jzwmcguid
    End If
End Function



Function GetAllGuiHXKZBH(AllGuiHXKZBH)
    AllGuiHXKZBH = ""
    projectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb projectName
    sql = "SELECT JG_建设工程规划许可证信息属性表.GuiHXKZBH FROM ((JG_建设工程规划许可证信息属性表 Inner Join ZD_宗地基本信息属性表 on ZD_宗地基本信息属性表.YDHXGUID = JG_建设工程规划许可证信息属性表.YDHXGUID) Inner Join GeoAreaTB on GeoAreaTB.ID = ZD_宗地基本信息属性表.ID) where ( (GeoAreaTB.mark mod 2) <> 0 ) ORDER BY JG_建设工程规划许可证信息属性表.GuiHXKZBH;"
    SSProcess.OpenAccessRecordset projectName, sql
    rscount = SSProcess.GetAccessRecordCount (projectName, sql )
    If rscount > 0 Then
        SSProcess.AccessMoveFirst projectName, sql
        While (SSProcess.AccessIsEOF (projectName, sql ) = False)
            SSProcess.GetAccessRecord projectName, sql, fields, values
            If values <> "" And  values <> "*" And values <> "NULL" Then
                If AllGuiHXKZBH = "" Then
                    AllGuiHXKZBH = values
                Else
                    AllGuiHXKZBH = AllGuiHXKZBH & "," & values
                End If
            End If
            SSProcess.AccessMoveNext projectName, sql
        WEnd
    End If
    SSProcess.CloseAccessRecordset projectName, sql
    SSProcess.CloseAccessMdb projectName
End Function

'根据工规证获取证下所有栋号
Function GetAllDH(GHXKZBH,AllDH)
    AllDH = ""
    projectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb projectName
    sql = "SELECT JG_建设工程建筑单体信息属性表.JianZWMC FROM ((JG_建设工程建筑单体信息属性表 inner join JG_建设工程规划许可证信息属性表 on JG_建设工程建筑单体信息属性表.JSGHXKZGUID = JG_建设工程规划许可证信息属性表.JSGHXKZGUID)  inner join ZD_宗地基本信息属性表 on ZD_宗地基本信息属性表.YDHXGUID = JG_建设工程规划许可证信息属性表.YDHXGUID) inner join GeoAreaTB on GeoAreaTB.ID = ZD_宗地基本信息属性表.ID  WHERE (([JG_建设工程建筑单体信息属性表].[GuiHXKZBH] = '" & GHXKZBH & "') And ((GeoAreaTB.mark mod 2) <> 0))  ORDER BY JG_建设工程建筑单体信息属性表.JianZWMC;"
    SSProcess.OpenAccessRecordset projectName, sql
    rscount = SSProcess.GetAccessRecordCount (projectName, sql )
    
    If rscount > 0 Then
        SSProcess.AccessMoveFirst projectName, sql
        While (SSProcess.AccessIsEOF (projectName, sql ) = False)
            SSProcess.GetAccessRecord projectName, sql, fields, values
            If values <> "" And  values <> "*" And values <> "NULL" Then
                If AllDH = "" Then
                    AllDH = values
                Else
                    AllDH = AllDH & "," & values
                End If
            End If
            SSProcess.AccessMoveNext projectName, sql
        WEnd
    End If
    SSProcess.CloseAccessRecordset projectName, sql
    SSProcess.CloseAccessMdb projectName
End Function

'根据工规证获取证下所有栋号
Function Getjzwmcguid(GHXKZBH,jzwmc,jzwmcguid)
    AllDH = ""
    projectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb projectName
    sql = "SELECT JG_建设工程建筑单体信息属性表.JZWMCGUID FROM ((JG_建设工程建筑单体信息属性表 inner join JG_建设工程规划许可证信息属性表 on JG_建设工程建筑单体信息属性表.JSGHXKZGUID = JG_建设工程规划许可证信息属性表.JSGHXKZGUID)  inner join ZD_宗地基本信息属性表 on ZD_宗地基本信息属性表.YDHXGUID = JG_建设工程规划许可证信息属性表.YDHXGUID) inner join GeoAreaTB on GeoAreaTB.ID = ZD_宗地基本信息属性表.ID  WHERE (([JG_建设工程建筑单体信息属性表].[GuiHXKZBH] = '" & GHXKZBH & "' AND [JG_建设工程建筑单体信息属性表].[JianZWMC] = '" & jzwmc & "') And ((GeoAreaTB.mark mod 2) <> 0))  ORDER BY JG_建设工程建筑单体信息属性表.JianZWMC;"
    SSProcess.OpenAccessRecordset projectName, sql
    rscount = SSProcess.GetAccessRecordCount (projectName, sql )
    
    If rscount > 0 Then
        SSProcess.AccessMoveFirst projectName, sql
        While (SSProcess.AccessIsEOF (projectName, sql ) = False)
            SSProcess.GetAccessRecord projectName, sql, fields, values
            If values <> "" And  values <> "*" And values <> "NULL" Then
                If jzwmcguid = "" Then
                    jzwmcguid = values
                Else
                    jzwmcguid = jzwmcguid & "," & values
                End If
            End If
            SSProcess.AccessMoveNext projectName, sql
        WEnd
    End If
    SSProcess.CloseAccessRecordset projectName, sql
    SSProcess.CloseAccessMdb projectName
End Function


'获取要编辑属性的宗地ID，优先取选择集的宗地，若无选择，全图找宗地，找到唯一个宗地，提取ZRZID，找到多个宗地，则提醒人工选择
Function GetZDID()
    '判断选择集是否有宗地
    GetZDID = 1
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.UpdateSysSelection 0
    geoCount = SSProcess.GetSelGeoCount
    LSID = ""
    ZRZCount = 0
    For i = 0 To geoCount - 1
        geoCode = SSProcess.GetSelGeoValue (i, "SSObj_Code")
        If  InStr("," & ZDCode & ",","," & geoCode & ",") > 0 Then
            LSID = SSProcess.GetSelGeoValue (i, "SSObj_ID")
            ZDCount = ZDCount + 1
        End If
    Next
    If ZDCount = 1 Then
        CURZDID = LSID
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
            CURZDID = SSProcess.GetSelGeoValue (0, "SSObj_ID")
        Else
            GetZDID = 2
            MsgBox "图上有多个宗地，且当前选择集内不包含宗地，请选择宗地，再执行该功能X"
        End If
    End If
End Function

'编辑项失去焦点
Function OnEditKillFocus( tableName, objectID, fieldName, fieldValue )
    If fieldName = "[ZRZH]" Then
        ZRZH = SSProcess.GetScriptDlgCellValue_EX(dlgHandle,"[ZRZH]")
        ZL = SSProcess.GetScriptDlgCellValue_EX(dlgHandle,"[ZL]")
        If InStr(ZL,ZRZH) = 0  Then
            SSProcess.SetScriptDlgCellValue_ex dlgHandle, "[ZL]",ZL & ZRZH
        End If
    End If
End Function

'选择集发生变化时，修改ZDID的值和界面的显示
Function OnSelectionChange( )
    '判断选择集是否有自然幢
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.UpdateSysSelection 0
    geoCount = SSProcess.GetSelGeoCount
    LSID = ""
    ZRZCount = 0
    For i = 0 To geoCount - 1
        geoCode = SSProcess.GetSelGeoValue (i, "SSObj_Code")
        If  InStr("," & ZDCode & ",","," & geoCode & ",") > 0 Then
            LSID = SSProcess.GetSelGeoValue (i, "SSObj_ID")
            ZDCount = ZDCount + 1
        End If
    Next
    Dim strs(10)
    If ZDCount = 1 Then
        ids = SSProcess.SearchInnerObjIDs (LSID, 2, ZRZCode, 1)
        ' If ids = "" Then Msgbox "图上宗地面下无自然幢，无需使用该功能X" : Exit Function
        
        CURZDID = LSID
        GetAllZRZAndLJZInfo()
        XMMC = SSProcess.GetObjectAttr (CURZDID,"[XiangMMC]")
        ZL = SSProcess.GetObjectAttr (CURZDID,"[ZL]")
        
        ZRZHS = ""
        LJZHS = ""
        ZRZH = ""
        LJZH = ""
        For i = 0 To ZRZCOUNT - 1
            If i = 0 Then
                CURZRZID = ZRZInfos(i,0)
                ZRZH = ZRZInfos(i,2) & "(" & ZRZInfos(i,1) & ")" & "(ID:" & ZRZInfos(i,0) & ")"
                ZRZHS = ZRZH
            Else
                ZRZHS = ZRZHS & "," & ZRZInfos(i,2) & "(" & ZRZInfos(i,1) & ")" & "(ID:" & ZRZInfos(i,0) & ")"
            End If
        Next
        
        For i = 0 To ConCount - 1
            If CURZRZID <> "" Then
                SSFunc.ScanString GFields(i),",",strs,scount
                Valuestr = ""
                For j = 0 To scount - 1
                    If j = 0 Then Valuestr = SSProcess.GetObjectAttr (CURZRZID,"[" & strs(j) & "]")
                    If j > 0 And Valuestr <> "" Then Valuestr = Valuestr & "、" & SSProcess.GetObjectAttr (CURZRZID,"[" & strs(j) & "]")
                Next
                If ConIDS(i) = "[ZRZH]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[ZRZH]",AltValues(i)
                If ConIDS(i) = "[CHZT]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[CHZT]",AltValues(i)
                If ConIDS(i) = "[FWJG]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[FWJG]",AltValues(i)
                If ConIDS(i) = "[ChanB]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[ChanB]",AltValues(i)
                If ConIDS(i) = "[FWXZ]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[FWXZ]",AltValues(i)
                If ConIDS(i) = "[SFZYYT]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[SFZYYT]",AltValues(i)
                If ConIDS(i) = "[GHYT]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[GHYT]",AltValues(i)
                If ConIDS(i) = "[FWYT]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[FWYT]",AltValues(i)
                If ConIDS(i) = "[QD]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[QD]",AltValues(i)
                If ConIDS(i) = "[QN]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[QN]",AltValues(i)
                If ConIDS(i) = "[QX]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[QX]",AltValues(i)
                If ConIDS(i) = "[QB]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[QB]",AltValues(i)
                '数据字典
                GetAllGuiHXKZBH allghxkz
                allghxkz = "," & allghxkz
                If Valuestr = "" Then
                    If ConIDS(i) = "[ghxkzbh]"  Then
                        SSProcess.SetScriptDlgCellOptions_ex dlgHandle,ConIDS(i),allghxkz
                    ElseIf ConIDS(i) = "[XMMC]"  Then
                        SSProcess.SetScriptDlgCellValue_ex dlgHandle,ConIDS(i),XMMC
                    ElseIf ConIDS(i) = "[ZL]"  Then
                        SSProcess.SetScriptDlgCellValue_ex dlgHandle,ConIDS(i),ZL
                    Else
                        SSProcess.SetScriptDlgCellValue_ex dlgHandle,ConIDS(i),DefValues(i)
                    End If
                Else
                    
                    SSProcess.SetScriptDlgCellValue_ex dlgHandle,ConIDS(i),Valuestr      '获取当前已有值填充
                End If
            Else
                SSProcess.SetScriptDlgCellValue_ex dlgHandle,ConIDS(i),DefValues(i)
            End If
        Next
        
        SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[自然幢号列表]",ZRZHS
        SSProcess.SetScriptDlgCellValue_ex dlgHandle, "[自然幢号列表]",ZRZH
        
    ElseIf ZDCount = 0 Then
        ' Msgbox "选择集内没有宗地，请重新选择X"
    Else
        'Msgbox "选择集内有多个宗地，请重新选择X"
    End If
End Function

Sub OnExitScript()
    '添加代码
End Sub

Sub OnOK()
    '添加代码
End Sub

Function GetAllljz(ByRef ljzguid)
    ljzguid = ""
    projectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb projectName
    sql = "Select LJZGUID from FC_逻辑幢信息表 where ID = " & CURLJZID & ""
    SSProcess.OpenAccessRecordset projectName, sql
    rscount = SSProcess.GetAccessRecordCount (projectName, sql )
    If rscount > 0 Then
        SSProcess.AccessMoveFirst projectName, sql
        While (SSProcess.AccessIsEOF (projectName, sql ) = False)
            SSProcess.GetAccessRecord projectName, sql, fields, values
            If values <> "" And  values <> "*" And values <> "NULL" Then
                If ljzguid = "" Then
                    ljzguid = values
                Else
                    ljzguid = ljzguid & "," & values
                End If
            End If
            SSProcess.AccessMoveNext projectName, sql
        WEnd
    End If
    SSProcess.CloseAccessRecordset projectName, sql
    SSProcess.CloseAccessMdb projectName
End Function

'所有按钮的实现在这里
Function OnButtonClick( tableName, objectID, fieldName, fieldValue )
    Dim strs(100),strs1(100),LZJarRecordList(),LJZInfo(2)
    mark = 1
    If fieldName = "[UPDATEXX]" Then            '自然幢逻辑幢更新
        '宗地信息
        zddm = SSProcess.GetObjectAttr(CURZDID,"[ZDDM]")
        zdguid = SSProcess.GetObjectAttr(CURZDID,"[ZDGUID]")
        '自然幢信息
        zrzh = SSProcess.GetObjectAttr(CURZRZID,"[ZRZH]")
        zrzguid = SSProcess.GetObjectAttr(CURZRZID,"[ZRZGUID]")
        
        '批量更新
        '户、面积块、各种线
        SSProcess.PushUndoMark
        SSProcess.ClearSelection
        SSProcess.ClearSelectCondition
        SSProcess.SetSelectCondition "[ZRZGUID]", "==", zrzguid
        SSProcess.SetSelectCondition "SSObj_LayerName", "<>", "自然幢,使用权宗地"
        SSProcess.SelectFilter
        SSProcess.ChangeSelectionObjAttr "[ZDDM]", zddm
        SSProcess.ChangeSelectionObjAttr "[zdguid]", zdguid
        SSProcess.ChangeSelectionObjAttr "[zrzh]", zrzh
        SSProcess.ChangeSelectionObjAttr "[zrzguid]", zrzguid
        
        '逻辑幢
        condition = "ljzguid='" & ljzguid & "'"
        setFieldValues = "zrzh=" & "'" & zrzh & "',zrzguid='" & zrzguid & "'"
        SSProcess.OpenAccessMdb mdbName
        ModifyRecord  "FC_逻辑幢信息表", condition, setFieldValues
        SSProcess.CloseAccessMdb mdbName
        '自然幢
        SSProcess.PushUndoMark
        SSProcess.ClearSelection
        SSProcess.ClearSelectCondition
        SSProcess.SetSelectCondition "[zrzguid]", "==", zrzguid
        SSProcess.SetSelectCondition "SSObj_LayerName", "=", "自然幢"
        SSProcess.SelectFilter
        SSProcess.ChangeSelectionObjAttr "[ZDDM]", zddm
        SSProcess.ChangeSelectionObjAttr "[zdguid]", zdguid
    End If
    
    If fieldName = "[SAVE]" Then                    '点击保存按钮
        'SSProcess.SetObjectAttr CURZRZID,"[LJZHLB]","1"
        LCFZXX = SSProcess.GetScriptDlgCellValue_EX(dlgHandle,"[LCFZXX]")
        ZRZH = SSProcess.GetScriptDlgCellValue_EX(dlgHandle,"[ZRZH]")
        ZRZSXH = SSProcess.GetScriptDlgCellValue_EX(dlgHandle,"[ZRZSXH]")
        If LCFZXX = "" Or LCFZXX = "*"  Then
            MsgBox "请填写楼层分组信息！"
            mark = 0
        Else
            LCFZXX = Replace(LCFZXX,",","|")
            artemp = Split(LCFZXX,";")
            For jj = 0 To UBound(artemp)
                artemp1 = Split(artemp(jj),"+")
                For jjj = 0 To UBound(artemp1)
                    If IsNumeric(artemp1(jjj)) = False  Then MsgBox "楼层分组信息填写有误！"
                    mark = 0
                Next
            Next
        End If
        If ZRZH = "" Or ZRZH = "*"  Then
            MsgBox "请填写自然幢号！"
            mark = 0
        End If
        If ZRZSXH = "" Or ZRZSXH = "*"  Then
            MsgBox "请填写自然幢顺序号！"
            mark = 0
        End If
        
        If mark = 0 Then
            MsgBox "保存失败！"
        Else
            SaveButtonProess()
            
            SSProcess.MapMethod "clearattrbuffer", "FC_自然幢信息属性表"
            SSProcess.MapMethod "clearattrbuffer", "FC_逻辑幢信息表"
            SSProcess.ExecuteSDLFunction "$SDL.SSProject.Display.RedrawExtend", 0
            MsgBox "保存成功！"
        End If
    End If
    
    If fieldName = "[Location]" Then'点击定位按钮
        SSProcess.GetObjectFocusPoint CURZRZID, x, y
        SSProcess.MapMethod  "addtrinkleobject",CURZRZID
        SSProcess.MoveScreen x ,y,1
        'CURZRZID = ""
    End If
    
    If fieldName = "[DeleteZRZ]" Then'点击删除自然幢
        ret = MsgBox ( "将删除自然幢及相关图形,继续请按确定键,按取消键放弃。", 1)
        If ret = 2 Then Exit Function
        zrzguid = SSProcess.GetObjectAttr(CURZRZID,"[ZRZGUID]")
        If zrzguid <> "" Then
            '删除逻辑幢
            condition = "ZRZGUID ='" & zrzguid & "'"
            DeleteRecord1   "FC_逻辑幢信息表", condition
            
            SSProcess.PushUndoMark
            SSProcess.ClearSelection
            SSProcess.ClearSelectCondition
            SSProcess.SetSelectCondition "[ZRZGUID]", "==", zrzguid
            SSProcess.SelectFilter
            SSProcess.DeleteSelectionObj
        End If
        SSProcess.RefreshView()
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
            arSQLRecord(iRecordCount) = values
            iRecordCount = iRecordCount + 1
            '移动记录游标
            SSProcess.AccessMoveNext mdbName, sql
        WEnd
    End If
    '关闭记录集
    SSProcess.CloseAccessRecordset mdbName, sql
    SSProcess.CloseAccessMDB mdbName
End Function


Function SaveButtonProess()
    Dim strs(10),strs1(10)
    zrzsql = ""
    If CURZRZID <> "" Then
        
        For i = 0 To ConCount - 1
            value_ = SSProcess.GetScriptDlgCellValue_ex (dlgHandle,ConIDS(i))
            SSFunc.ScanString GFields(i),",",strs,scount                'GFields(i)  对应地物字段
            SSFunc.ScanString value_,"、",strs1,scount1
            If scount = 1 Then
                If scount1 > 1 Then
                    For j = 0 To scount - 1
                        SSProcess.SetObjectAttr CURZRZID,"[" & strs(j) & "]",strs1(j)
                    Next
                End If
            End If
        Next
        
        GetAllZRZAndLJZInfo()
        
        ZRZHS = ""
        LJZHS = ""
        ZRZH = ""
        LJZH = ""
        For i = 0 To ZRZCOUNT - 1
            If i = 0 Then
                DQZRZID = ZRZInfos(i,0)
                ZRZH = ZRZInfos(i,2) & "(" & ZRZInfos(i,1) & ")" & "(ID:" & ZRZInfos(i,0) & ")"
                ZRZHS = ZRZH
            Else
                ZRZHS = ZRZHS & "," & ZRZInfos(i,2) & "(" & ZRZInfos(i,1) & ")" & "(ID:" & ZRZInfos(i,0) & ")"
            End If
        Next
        SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[自然幢号列表]",ZRZHS
        SSProcess.SetScriptDlgCellValue_ex dlgHandle, "[自然幢号列表]",ZRZH
        
        XMMC = SSProcess.GetObjectAttr (CURZDID,"[XiangMMC]")
        ZL = SSProcess.GetObjectAttr (CURZDID,"[ZL]")
        
        For i = 0 To ConCount - 1
            If DQZRZID <> "" Then
                SSFunc.ScanString GFields(i),",",strs,scount
                Valuestr = ""
                For j = 0 To scount - 1
                    If j = 0 Then Valuestr = SSProcess.GetObjectAttr (DQZRZID,"[" & strs(j) & "]")
                    If j > 0 And Valuestr <> "" Then Valuestr = Valuestr & "、" & SSProcess.GetObjectAttr (DQZRZID,"[" & strs(j) & "]")
                Next
                If ConIDS(i) = "[ZRZH]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[ZRZH]",AltValues(i)
                If ConIDS(i) = "[CHZT]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[CHZT]",AltValues(i)
                If ConIDS(i) = "[FWJG]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[FWJG]",AltValues(i)
                If ConIDS(i) = "[ChanB]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[ChanB]",AltValues(i)
                If ConIDS(i) = "[FWXZ]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[FWXZ]",AltValues(i)
                If ConIDS(i) = "[SFZYYT]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[SFZYYT]",AltValues(i)
                If ConIDS(i) = "[GHYT]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[GHYT]",AltValues(i)
                If ConIDS(i) = "[FWYT]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[FWYT]",AltValues(i)
                If ConIDS(i) = "[QD]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[QD]",AltValues(i)
                If ConIDS(i) = "[QN]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[QN]",AltValues(i)
                If ConIDS(i) = "[QX]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[QX]",AltValues(i)
                If ConIDS(i) = "[QB]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[QB]",AltValues(i)
                
                If Valuestr = "" Then
                    If ConIDS(i) = "[ghxkzbh]"  Then
                        SSProcess.SetScriptDlgCellOptions_ex dlgHandle,ConIDS(i),allghxkz
                    ElseIf ConIDS(i) = "[XMMC]"  Then
                        SSProcess.SetScriptDlgCellValue_ex dlgHandle,ConIDS(i),XMMC
                    ElseIf ConIDS(i) = "[ZL]"  Then
                        SSProcess.SetScriptDlgCellValue_ex dlgHandle,ConIDS(i),ZL
                    Else
                        SSProcess.SetScriptDlgCellValue_ex dlgHandle,ConIDS(i),DefValues(i)
                    End If
                Else
                    SSProcess.SetScriptDlgCellValue_ex dlgHandle,ConIDS(i),Valuestr      '获取当前已有值填充
                End If
            Else
                SSProcess.SetScriptDlgCellValue_ex dlgHandle,ConIDS(i),DefValues(i)
            End If
        Next
        
        
        ZDGUID = SSProcess.GetObjectAttr(CURZRZID,"[ZDGUID]")
        ZRZGUID = SSProcess.GetObjectAttr(CURZRZID,"[ZRZGUID]")
        CHZT = SSProcess.GetObjectAttr(CURZRZID,"[CHZT]")
        ZRZH = SSProcess.GetObjectAttr(CURZRZID,"[ZRZH]")
        LCFZXX = SSProcess.GetObjectAttr(CURZRZID,"[LCFZXX]")
        
        SSProcess.OpenAccessMdb mdbName
        Dim LCarRecordList(),LCRecordListCount
        sql = "Select FC_楼层信息属性表.ID,CH From FC_楼层信息属性表 INNER JOIN GeoAreaTB ON FC_楼层信息属性表.ID=GeoAreaTB.ID WHERE (GeoAreaTB.Mark mod 2)<>0 And FC_楼层信息属性表.ZRZGUID = " & ZRZGUID & " Order By CH;"
        GetSQLRecordAll mdbName,sql,LCarRecordList,LCRecordListCount
        
        
        
        For i = 0 To LCRecordListCount - 1
            artempch = Split(LCarRecordList(i),",")
            artemp = Split(LCFZXX,";")
            bzcbs = 0
            For jj = 0 To UBound(artemp)
                artemp1 = Split(artemp(jj),"+")
                For jjj = 0 To UBound(artemp1)
                    If  CDbl(artempch(1)) = CDbl(artemp1(0))  Then
                        bzcbs = 1
                        jjjj = jj
                    End If
                Next
            Next
            If bzcbs = 1  Then
                SSProcess.SetObjectAttr artempch(0), "[LCXX]", artemp(jjjj)
            Else
                SSProcess.SetObjectAttr artempch(0), "[LCXX]", ""
            End If
        Next
        
        Dim LZJarRecordList(),LZJRecordListCount
        
        sql = "SELECT FC_逻辑幢信息表.ID FROM FC_逻辑幢信息表 WHERE FC_逻辑幢信息表.ZRZGUID =" & zrzguid
        GetSQLRecordAll mdbName,sql,LZJarRecordList,LZJRecordListCount
        
        If LZJRecordListCount < 1  Then
            
            newGUID = GetNewGUID()
            sql = "Insert Into FC_逻辑幢信息表(FeatureGUID,ZDGUID,ZRZGUID,LJZGUID,ZRZH,LJZH,CHZT,LCFZXX) values('" & newGUID & "','" & ZDGUID & "','" & ZRZGUID & "','" & newGUID & "','" & ZRZH & "','1'," & CHZT & ",'" & LCFZXX & "')"
            SSProcess.ExecuteAccessSql mdbName, sql
            
        ElseIf LZJRecordListCount = 1  Then
            setsql = "CHZT" & "=" & CHZT
            sql = "update FC_逻辑幢信息表 set " & setsql & " where ID=" & LZJarRecordList(0)
            SSProcess.ExecuteAccessSql mdbName, sql
            
            setsql = "ZRZH" & "='" & ZRZH & "'"
            sql = "update FC_逻辑幢信息表 set " & setsql & " where ID=" & LZJarRecordList(0)
            SSProcess.ExecuteAccessSql mdbName, sql
            
            setsql = "LCFZXX" & "='" & LCFZXX & "'"
            sql = "update FC_逻辑幢信息表 set " & setsql & " where ID=" & LZJarRecordList(0)
            SSProcess.ExecuteAccessSql mdbName, sql
        Else
            MsgBox "逻辑幢存在多条记录，请检查！"
        End If
        SSProcess.CloseAccessMdb mdbName
        SSProcess.MapMethod "clearattrbuffer", "FC_逻辑幢信息表"
        
    End If
    
End Function

'列表框选择项改变
Function OnListBoxSelChange( tableName, objectID, fieldName, fieldValue )
    Dim strs(100)
    If fieldName = "[自然幢号列表]" Then
        index = InStr(fieldValue,":")
        STR = Right(fieldValue,Len(fieldValue) - index)
        CURZRZID = Left(STR,Len(STR) - 1)
        For i = 0 To ConCount - 1
            If CURZDID <> "" Then
                SSFunc.ScanString GFields(i),",",strs,scount
                Valuestr = ""
                For j = 0 To scount - 1
                    If j = 0 Then Valuestr = SSProcess.GetObjectAttr (CURZRZID,"[" & strs(j) & "]")
                    If j > 0 And Valuestr <> "" Then Valuestr = Valuestr & "、" & SSProcess.GetObjectAttr (CURZRZID,"[" & strs(j) & "]")
                Next
                If ConIDS(i) = "[ZRZH]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[ZRZH]",AltValues(i)
                If ConIDS(i) = "[CHZT]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[CHZT]",AltValues(i)
                If ConIDS(i) = "[FWJG]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[FWJG]",AltValues(i)
                If ConIDS(i) = "[ChanB]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[ChanB]",AltValues(i)
                If ConIDS(i) = "[FWXZ]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[FWXZ]",AltValues(i)
                If ConIDS(i) = "[SFZYYT]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[SFZYYT]",AltValues(i)
                If ConIDS(i) = "[GHYT]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[GHYT]",AltValues(i)
                If ConIDS(i) = "[FWYT]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[FWYT]",AltValues(i)
                If ConIDS(i) = "[QD]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[QD]",AltValues(i)
                If ConIDS(i) = "[QN]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[QN]",AltValues(i)
                If ConIDS(i) = "[QX]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[QX]",AltValues(i)
                If ConIDS(i) = "[QB]" Then SSProcess.SetScriptDlgCellOptions_ex dlgHandle, "[QB]",AltValues(i)
                '数据字典
                If Valuestr = ""  Then
                    '图幅号
                    SSProcess.SetScriptDlgCellValue_ex dlgHandle,ConIDS(i),DefValues(i)
                Else
                    SSProcess.SetScriptDlgCellValue_ex dlgHandle,ConIDS(i),Valuestr      '获取当前已有值填充
                End If
            Else
                SSProcess.SetScriptDlgCellValue_ex dlgHandle,ConIDS(i),DefValues(i)
            End If
        Next
        
    End If
End Function

Sub OnCancel()
    '添加代码
End Sub

Function GetNewGUID()
    Set TypeLib = CreateObject("Scriptlet.TypeLib")
    GetNewGUID = Left(TypeLib.Guid,38)
    Set TypeLib = Nothing
End Function

'修改记录
Function ModifyRecord(tableName, condition, setfieldValues)
    sqlString = "Update " & tableName & " Set " & setfieldValues & " where " & condition
    ModifyRecord = SSProcess.ExecuteSql (sqlString)
End Function


'解析DLG对应的ini设置记录
'入参，Ininame：ini文件的名称
'入参返回：ControlIDS：控件ID数组形式，GeoFields：对应地物的字段名数组形式，DefaultValues：默认值数组形式，AlternativeValues：备选项内容数组形式，MemoryValues：是否记忆上次传输的值数组形式，ControlCount ：控件个数
Function ReadIniInfo(ByVal Ininame,ByRef  ControlIDS(),ByRef GeoFields(),ByRef DefaultValues(),ByRef AlternativeValues(),ByRef MemoryValues(),ByRef ControlCount)
    ControlCount = 0
    TemplateFileName = SSProcess.GetTemplateFileName
    Ininame_ = Left(TemplateFileName,Len(TemplateFileName) - 4) & "\" & Ininame
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

'删除记录
Function DeleteRecord1( tableName, condition)
    sqlString = "Delete from " & tableName & " where " & condition
    DeleteRecord1 = SSProcess.ExecuteSql (sqlString)
End Function


'房屋结构
Function SetFWJG(ByVal fwjg, ByVal JGNAEM)
    If fwjg = 1 Then SSProcess.SetObjectAttr CURZRZID,JGNAEM,"钢结构"
    If fwjg = 2 Then SSProcess.SetObjectAttr CURZRZID,JGNAEM,"钢和钢筋混凝土结构"
    If fwjg = 3 Then SSProcess.SetObjectAttr CURZRZID,JGNAEM,"钢筋混凝土结构"
    If fwjg = 4 Then SSProcess.SetObjectAttr CURZRZID,JGNAEM,"混合结构"
    If fwjg = 5 Then SSProcess.SetObjectAttr CURZRZID,JGNAEM,"砖木结构"
    If fwjg = 6 Then SSProcess.SetObjectAttr CURZRZID,JGNAEM,"其它结构"
    If fwjg = "" Then SSProcess.SetObjectAttr CURZRZID,JGNAEM,""
End Function


'开库
Function inidatabase(ByRef adoConnection)
    Set adoConnection = CreateObject("ADODB.Connection")
    adoConnection.connectionstring = "Provider=ORAOLEDB.ORACLE.1;Data Source=192.168.2.15/orcl;user id=zhuser;password=zhuser"
    adoConnection.Open
End Function

'关库
Function releasedatabase(ByVal adoConnection)
    adoConnection.Close
    Set adoConnection = Nothing
End Function

'获取oracle系统时间
Function getSysDate(ByVal objconn,ByRef sysYear,ByRef sysDate,ByRef sysMinute)
    seqsql = "select to_char(sysdate,'yyyy'), to_char(sysdate,'yyyymm'), to_char(sysdate,'yyyy-mm-dd hh24:mi:ss') from dual"
    Set adoRs = CreateObject("ADODB.RECORDSET")
    adoRs.open seqsql, objconn, 3, 1
    sysYear = adoRs(0)
    sysDate = adoRs(1)
    sysMinute = adoRs(2)
    adoRs.close
    Set adoRs = Nothing
End Function

'插入记录
Function InsertRecord(ByVal tablename,ByVal FieldValues)
    Fieldstr = "YEWID,ZDGUID,ZDDM,ZRZGUID,ZRZH,SQR,SQSJ"
    sql = "insert into " & tablename & "(" & fieldstr & ") values(" & FieldValues & ")"
    adoConnection.Execute(sql)
    adoConnection.Execute("commit")
End Function


'删除记录
Function DeleteRecord(ByVal tablename,ByVal condition)
    delsql = "delete from " & tablename & " where " & condition
    adoConnection.Execute(delsql)
    adoConnection.Execute("commit")
End Function


'获取Oracle库属性
Function GetOracleValue(ByVal adoConnection, ByVal sql,ByRef rs(),ByRef rscount,ByVal fieldcount)
    rscount = 0
    Set adoRs = CreateObject("ADODB.recordset")
    adoRs.Open sql,adoConnection,3,3
    While adoRs.Eof = False
        For i = 0 To fieldcount - 1
            rs(rscount,i) = adoRs(i) & ""
        Next
        rscount = rscount + 1
        adoRs.MoveNext
    WEnd
    adoRs.Close
    Set adoRs = Nothing
End Function

Function GetZdInfo(ByRef ywid,ByRef zdguid,ByRef zddm)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
    SSProcess.SetSelectCondition "SSObj_Code", "=", "6803163,6803153"
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount
    If geoCount <> 1 Then Exit Function
    ywid = SSProcess.GetSelGeoValue (0, "[YeWID]")
    zdguid = SSProcess.GetSelGeoValue (0, "[ZDGUID]")
    zddm = SSProcess.GetSelGeoValue (0, "[ZDDM]")
End Function

Function GetZrzInfo(ByVal zrzid,ByRef zrzguid,ByRef zrzh)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_ID", "=", zrzid
    SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
    SSProcess.SetSelectCondition "SSObj_Code", "=", "3120033"
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount
    If geoCount <> 1 Then Exit Function
    zrzguid = SSProcess.GetSelGeoValue (0, "[ZRZGUID]")
    zrzh = SSProcess.GetSelGeoValue (0, "[ZRZH]")
End Function

'重新组织自然幢号
Function changezrzh(ByVal input,ByRef output)
    b = Int(Right(input,4) + 1)
    If Len(b) = 1 Then  output = "F" & String(3,"0") & b
    If Len(b) = 2 Then output = "F" & String(2,"0") & b
    If Len(b) = 3 Then output = "F" & String(1,"0") & b
    If Len(b) = 4 Then output = "F" & b
End Function


'数据字典
Function GetFieldDIC(fieldName)
    strResult = ""
    '获取属性字典的文件位置
    strFilename = SSProcess.GetTemplateFileName
    strFilename = Left(strFilename,Len(strFilename) - 4)
    If Right(fieldName,4) = "NAME"  Then
        strFilename = strFilename & "\" & Left(fieldName,Len(fieldName) - 4) & ".DIC"
    Else
        strFilename = strFilename & "\" & fieldName & ".DIC"
    End If
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    'msgbox   "strFilename=" & strFilename  & ",exist=" & fso.fileExists(strFilename)  
    If fso.fileExists(strFilename) = True Then
        'msgbox   "fieldName=" & fieldName & ",strFilename=" & strFilename        
        Set ts = fso.OpenTextFile(strFilename, 1)
        strDicFields = strDicFields & "," & fieldName '记录有下拉列表的字段
        Do While Not ts.AtEndOfStream
            chLine = ts.ReadLine '& NewLine             
            If  chLine <> "" Then
                arList = Split(chLine,",")
                nListCount = UBound(arList)
                If nListCount >= 1 Then
                    strResult = strResult & "," & arlist(0) & "(" & arlist(1) & ")"
                    i = i + 1
                End If
            End If
        Loop
        ts.Close
    End If
    Set fso = Nothing
    'msgbox  fieldName & "-strResult=" & strResult   
    If  strResult <> "" Then
        GetFieldDIC = Mid(strResult,2)
    Else
        GetFieldDIC = strResult
    End If
End Function

Function CalHuValue(HuName,fieldoptions)
    hxlist = Split(fieldoptions,",")
    hxresult = ""
    For i = 0 To UBound(hxlist)
        hxlistvalue = hxlist(i)
        If InStr(hxlistvalue,HuName) > 0 Then
            hxresult = Mid(hxlistvalue,1,InStr(hxlistvalue,"(") - 1)
            Exit For
        End If
    Next
    CalHuValue = hxresult
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



