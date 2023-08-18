
'========================================================Doc操作对象和文件路径操作对象================================================================

'Doc全局对象
Dim Global_Word
Set Global_Word = CreateObject ("asposewordscom.asposewordshelper")

'路径操作对象
Dim FileSysObj
Set FileSysObj = CreateObject("Scripting.FileSystemObject")

'============================================================字段&替换字段配置====================================================

'KeyStr = "编号,项目名称,项目地址,设计单位,建设单位,委托单位,测绘单位,外业时间,点最大较差值,高程最大较差值,深度最大较差值"

'TemplateVal = "BH,XMMC,XMDZ,SJDW,JSDW,WTDW,CHDW,WYSJ,MaxPoi,MaxHei,MaxDeep"

XMZD = "BH,XMMC,XMDZ,SJDW,JSDW,WTDW,CHDW,WYSJ,CHSJ"

KeyStr = "编号,项目名称,项目地址,设计单位,建设单位,委托单位,测绘单位,外业时间,测绘时间"

ReplaceVal = "CHSJ,CGTMC"

'===========================================管线信息=======================================================

'管线项目信息数组
Dim GXProjectInfo()

'===========================================功能入口========================================================

'总入口
Sub OnClick()
    
    If  TypeName (Global_Word) = "AsposeWordsHelper" Then
        Global_Word.CreateDocumentByTemplate  SSProcess.GetSysPathName (7) & "输出模板\" & "管线输出模板.doc"
    Else
        MsgBox "请先注册Aspose.Word插件"
        Exit Sub
    End If
    
    AllVisible
    
    InputInfo ExportFormat,BoolStr,GXProjectInfo,XMMC
    
    If BoolStr = 0 Then
        MsgBox"取消输出，已退出"
        Exit Sub
    End If
    
    ReplaceValue KeyStr,XMZD,DelCount,DelNodeRow
    
    DelNodeParagraph 0,DelCount,DelNodeRow
    
    InnerGXTable 2,1
    
    InnerGZTable 3,1,HjRow,ExportFormat
    
    InnerHj 3,1,HjRow
    
    Global_Word.SaveEx  SSProcess.GetSysPathName(5) & XMMC & "管线报告.doc"
    
    Ending
    
End Sub' OnClick

'===========================================信息录入======================================================

'窗口信息录入函数
Function InputInfo(ByRef ExportFormat,ByRef BoolStr,ByRef GXProjectInfo(),ByRef XMMC)
    
    ReDim GXProjectInfo(8)
    
    ProJectName = SSProcess.GetProjectFileName
    KeyArr = Split(KeyStr,",", - 1,1)
    XMZDArr = Split(XMZD,",", - 1,1)
    
    SqlStr = "Select 管线项目信息表." & XMZD & " From 管线项目信息表  WHERE 管线项目信息表.ID=1"
    
    GetSQLRecordAll SqlStr,ProJectInfoArr,ResultCount
    
    InfoArr = Split(ProJectInfoArr(0), ",", - 1,1)
    XMMC = InfoArr(1)
    SSProcess.ClearInputParameter
    SSProcess.AddInputParameter "管线长度统计方式","二维长度",0,"二维长度,三维长度","管线长度统计方式"
    For i = 0 To  8
        SSProcess.AddInputParameter KeyArr(i), InfoArr(i), 0, "", ""
    Next 'i
    
    BoolStr = SSProcess.ShowInputParameterDlg ("管线报表输出方式")
    ExportFormat = SSProcess.GetInputParameter("管线长度统计方式")
    
    If BoolStr = 1 Then
        SSProcess.OpenAccessMdb ProJectName
        For i = 0  To 8
            GXProjectInfo(i) = SSProcess.GetInputParameter(KeyArr(i))
            SqlStr = "Update  管线项目信息表 Set " & XMZDArr(i) & " = '" & GXProjectInfo(i) & "'Where 管线项目信息表.ID= 1"
            SSProcess.ExecuteAccessSql  ProJectName,SqlStr
        Next
        SSProcess.CloseAccessMdb ProJectName
    End If
    
End Function' InputInfo

'==========================================================获取小类名称&填写表格=======================================================

EngStr = "CD,CT,CY,CQ,CS,QT,BM,FQ,DL,GD,LD,DC,XH,TX,DX,YD,LT,JX,JK,EX,DS,BZ,JS,XF,PS,YS,WS,FS,RQ,MQ,TR,YH,RL,RS,ZQ,SY,GS"

CheStr = "长输输电,长输通信,油主管道,天然气主管道,水主管道,其他主管道,不明,废弃,电力,供电,路灯,电车,交通信号,综 合,电信,移动,联通,军用,监控,电力通讯,广播电视,保密专用,生活工业用水,消防水,排水,雨水,污水,生活废水,燃气,煤气,天然气,液化气,热力,热水,蒸汽,石油,工业废水"

'填写管线测绘取舍标准表
Function InnerGXTable(ByVal TableIndex,ByVal StartRow) 'TableIndex 表格索引,StartRow 起始行数
    StrString = "Select DISTINCT GXLX From 地下管线点属性表 inner join GeoPointTB on 地下管线点属性表.ID = GeoPointTB.ID WHERE (GeoPointTB.Mark Mod 2)<>0 And 地下管线点属性表.GXLX <>'*' And 地下管线点属性表.GXLX <>''"
    GetSQLRecordAll StrString,LxArr,LxCount
    If LxCount > 1 Then
        Global_Word.CloneTableRow TableIndex,StartRow,1,LxCount - 1,False
        For i = 0 To LxCount - 1
            Global_Word.SetCellText TableIndex,i + StartRow,0,ToChinese(LxArr(i)),True,False
            If ToChinese(LxArr(i)) = "生活工业用水" Then
                Global_Word.SetCellText TableIndex,i + StartRow,1,"管径≥50mm",True,False
            ElseIf ToChinese(LxArr(i)) = "排水" Then
                Global_Word.SetCellText TableIndex,i + StartRow,1,"管径≥200mm或方沟≥400mm×400mm",True,False
            ElseIf  ToChinese(LxArr(i)) <> "" Then
                Global_Word.SetCellText TableIndex,i + StartRow,1,"全测",True,False
            End If
        Next 'i
    Else
        For i = 0 To LxCount - 1
            Global_Word.SetCellText TableIndex,i + StartRow,0,ToChinese(LxArr(i)),True,False
            If ToChinese(LxArr(i)) = "生活工业用水" Then
                Global_Word.SetCellText TableIndex,i + StartRow,1,"管径≥50mm",True,False
            ElseIf ToChinese(LxArr(i)) = "排水" Then
                Global_Word.SetCellText TableIndex,i + StartRow,1,"管径≥200mm或方沟≥400mm×400mm",True,False
            Else
                Global_Word.SetCellText TableIndex,i + StartRow,1,"全测",True,False
            End If
        Next 'i
    End If
End Function' InnerGXTable

'填写各专业管线工作量统计表
Function InnerGZTable(ByVal TableIndex,ByVal StartRow,ByRef HjRow,ByVal LenTypes) 'TableIndex 表格索引,StartRow 起始行数,HjRow 合计行数值(返回值),LenTypes 长度类型
    StrString = "Select DISTINCT GXLX From 地下管线线属性表 inner join GeoLineTB on 地下管线线属性表.ID = GeoLineTB.ID WHERE (GeoLineTB.Mark Mod 2)<>0 And 地下管线线属性表.GXLX <>'*' And 地下管线线属性表.GXLX <>''"
    GetSQLRecordAll StrString,LxArr,LxCount
    HjRow = StartRow + LxCount
    If LxCount > 1 Then
        Global_Word.CloneTableRow TableIndex,StartRow,1,LxCount - 1,False
        For i = 0 To LxCount - 1
            Global_Word.SetCellText TableIndex,i + StartRow,0,ToChinese(LxArr(i)),True,False
            InnerPoiCount LxArr(i),TableIndex,i + StartRow
            InnerLineLen LxArr(i),TableIndex,i + StartRow,LenTypes
        Next 'i
    Else
        For i = 0 To LxCount - 1
            Global_Word.SetCellText TableIndex,i + StartRow,0,ToChinese(LxArr(i)),True,False
            InnerPoiCount LxArr(i),TableIndex,i + StartRow
            InnerLineLen LxArr(i),TableIndex,i + StartRow,LenTypes
        Next 'i
    End If
End Function' InnerGZTable

'填写明显点和隐蔽点个数
Function InnerPoiCount(ByVal GxName,ByVal TableIndex,ByVal InsertRow) 'GxName 管线类型名称,TableIndex 表索引,InsertRow 指定插入行
    StrString = "Select FSW From 地下管线点属性表 inner join GeoPointTB on 地下管线点属性表.ID = GeoPointTB.ID WHERE (GeoPointTB.Mark Mod 2)<>0 And 地下管线点属性表.GXLX =" & "'" & GxName & "'"
    GetSQLRecordAll StrString,FswArr,PoiCount
    OuterPoiCount = 0
    InnerPoiCount = 0
    For i = 0 To PoiCount - 1
        If FswArr(i) = "" Then
            InnerPoiCount = InnerPoiCount + 1
        ElseIf FswArr(i) = "*"  Then
            InnerPoiCount = InnerPoiCount + 1
        ElseIf FswArr(i) = Null Then
            InnerPoiCount = InnerPoiCount + 1
        Else
            OuterPoiCount = OuterPoiCount + 1
        End If
    Next 'i
    Global_Word.SetCellText TableIndex,InsertRow,1,OuterPoiCount,True,False
    Global_Word.SetCellText TableIndex,InsertRow,2,InnerPoiCount,True,False
    Global_Word.SetCellText TableIndex,InsertRow,3,PoiCount,True,False
End Function' InnerPoiCount

'填写管线长度
Function InnerLineLen(ByVal GxName,ByVal TableIndex,ByVal InsertRow,ByVal LenTypes) 'GxName 管线类型名称,TableIndex 表索引,InsertRow 指定插入行,LenTypes 长度类型
    SelFeatures GxName,LineCount,LineArr
    If LenTypes = "二维长度" Then
        For i = 0 To LineCount - 1
            TotalLength = TotalLength + Round(Transform(SSProcess.GetObjectAttr(LineArr(i),"SSObj_Length")),0)
            'msgbox   TotalLength
        Next 'i
    ElseIf LenTypes = "三维长度" Then
        For i = 0 To LineCount - 1
            TotalLength = TotalLength + Round(Transform(SSProcess.GetObjectAttr(LineArr(i),"SSObj_3DLength")),0)
        Next 'i
    End If
    Global_Word.SetCellText TableIndex,InsertRow,4,TotalLength,True,False
End Function' InnerLineLen

'填写合计数
Function InnerHj(ByVal TableIndex,ByVal StartRow,ByVal HjRow) 'TableIndex 表索引,StartRow 起始行数,HjRow 合计行
    MxCount = 0
    YbCount = 0
    ZCount = 0
    LineLen = 0
    For i = StartRow To HjRow - 1
        MxCount = MxCount + Transform(GetSelCellVal(Global_Word.GetCellText(TableIndex,i,1,False)))
        YbCount = YbCount + Transform(GetSelCellVal(Global_Word.GetCellText(TableIndex,i,2,False)))
        ZCount = ZCount + Transform(GetSelCellVal(Global_Word.GetCellText(TableIndex,i,3,False)))
        LineLen = LineLen + Transform(GetSelCellVal(Global_Word.GetCellText(TableIndex,i,4,False)))
    Next 'i
    Global_Word.SetCellText TableIndex,HjRow,1,MxCount,True,False
    Global_Word.SetCellText TableIndex,HjRow,2,YbCount,True,False
    Global_Word.SetCellText TableIndex,HjRow,3,ZCount,True,False
    Global_Word.SetCellText TableIndex,HjRow,4,LineLen,True,False
End Function' InnerHj

'删除指定行
Function DelNodeParagraph(ByVal PageIndex,ByVal DelCount,ByVal DelNodeRow)
    If DelCount > 1 Then
        NodePosArr = Split(DelNodeRow,",", - 1,1)
        Count = UBound(NodePosArr)
        For i = 0 To Count
            Global_Word.MoveToSectionParagraph PageIndex,Transform(NodePosArr(i))
            Global_Word.DeleteCurrentParagraph
            For j = i + 1 To Count
                NodePosArr(j) = Transform(NodePosArr(j)) - 1
            Next 'j
        Next 'i
        Global_Word.MoveToSectionParagraph PageIndex,16 - DelCount
        For i = 1 To DelCount
            Global_Word.Writeln ""
        Next 'i
    ElseIf DelCount = 1 Then
        Global_Word.MoveToSectionParagraph PageIndex,Transform(DelNodeRow)
        Global_Word.DeleteCurrentParagraph
        Global_Word.MoveToSectionParagraph PageIndex,16 - DelCount
        For i = 1 To DelCount
            Global_Word.Writeln ""
        Next 'i
    End If
End Function' DelNodeParagraph

'层名转化为中文
Function ToChinese(ByVal EngLayerName) 'EngLayerName 图层名称(英文)
    EngArr = Split(EngStr,",", - 1,1)
    CheArr = Split(CheStr,",", - 1,1)
    ToChinese = ""
    For i = 0 To UBound(EngArr)
        If EngArr(i) = EngLayerName Then
            ToChinese = CheArr(i)
        End If
    Next 'i
End Function' ToChinese

'=========================================================字符串替换=======================================================

' [管线报告信息]
' 编号 = ""
' 项目名称 = ""
' 项目地址 = ""
' 设计单位 = ""
' 建设单位 = ""
' 委托单位 = ""
' 外业时间 = ""
' 测绘时间 = ""
' 点最大较差值 = ""
' 高程最大较差值 = ""
' 深度最大较差值 = ""

'字符替换函数


'字符串替换函数
Function ReplaceValue(ByVal ReplaceStr,ByVal OriginVal,ByRef DelCount,ByRef DelNodeRow)
    
    DelCount = 0
    DelNodeRow = ""
    
    ReplaceArr = Split(ReplaceStr,",", - 1,1)
    OriginArr = Split(OriginVal,",", - 1,1)
    
    For i = 0 To UBound(ReplaceArr)
        Global_Word.Replace "{" & OriginArr(i) & "}",GXProjectInfo(i),0
    Next 'i
    
    For i = 3 To 6
        Val = GXProjectInfo(i)
        If Val = "" Then
            DelCount = DelCount + 1
            If DelNodeRow = "" Then
                DelNodeRow = CStr(i + 8)
            Else
                DelNodeRow = DelNodeRow & "," & CStr(i + 8)
            End If
        End If
    Next 'i
    
    StrString = "Select DISTINCT GXLX From 地下管线线属性表 inner join GeoLineTB on 地下管线线属性表.ID = GeoLineTB.ID WHERE (GeoLineTB.Mark Mod 2)<>0"
    GetSQLRecordAll StrString,LxArr,LxCount
    For i = 0 To LxCount - 1
        If LxArr(i) <> "*" Then
            GXTStr = GXTStr & GXProjectInfo(1) & ToChinese(LxArr(i)) & "管线图" & Chr(13)
        End If
    Next 'i
    
    ExtraVal = ToBigDate(GetNowTime) & "," & GXTStr
    ExtraArr = Split(ExtraVal,",", - 1,1)
    
    NameArr = Split(ReplaceVal,",", - 1,1)
    For i = 0 To UBound(ExtraArr)
        Global_Word.Replace "{" & NameArr(i) & "}",ExtraArr(i),0
    Next 'i
    
    SqlStr = "Select 管线项目信息表." & "DWZDJC,GCZDJC,MSZDJC" & " From 管线项目信息表  WHERE 管线项目信息表.id=1"
    GetSQLRecordAll SqlStr,MaxNumArr,ResultCount
    
    ValArr = Split(MaxNumArr(0), ",", - 1,1)
    
    Global_Word.Replace "{" & "DWZDJC" & "}",ValArr(0),0
    Global_Word.Replace "{" & "GCZDJC" & "}",ValArr(1),0
    Global_Word.Replace "{" & "MSZDJC" & "}",ValArr(2),0
    
End Function' ReplaceValue

'==========================================================工具类函数=======================================================

'打开所有图层
Function AllVisible()
    count = SSProcess.GetLayerCount
    For i = 0 To count - 1
        layername = SSProcess.GetLayerName (i)
        SSProcess.SetLayerStatus layername, 1, 1
    Next
    SSProcess.RefreshView
End Function

'年分转大写
Function YearChange(ByVal YearName)
    Number = "1,2,3,4,5,6,7,8,9,0"
    BigNumber = "一,二,三,四,五,六,七,八,九,〇"
    NumberArr = Split(Number,",", - 1,1)
    BigNumberArr = Split(BigNumber,",", - 1,1)
    For i = 1 To 4
        For j = 0 To UBound(NumberArr)
            If Mid(YearName,i,1) = NumberArr(j) Then
                YearChange = YearChange & BigNumberArr(j)
            End If
        Next 'j
    Next 'i
    YearChange = YearChange & "年"
End Function' YearChange

'月份转大写
Function MonthChange(ByVal MonthName)
    Number = "1,2,3,4,5,6,7,8,9,10,11,12"
    BigNumber = "一,二,三,四,五,六,七,八,九,十,十一,十二"
    NumberArr = Split(Number,",", - 1,1)
    BigNumberArr = Split(BigNumber,",", - 1,1)
    For i = 0 To UBound(NumberArr)
        If MonthName = NumberArr(i) Then
            MonthChange = BigNumberArr(i) & "月"
        End If
    Next 'i
End Function' MonthChange

'日期转大写
Function ToBigDate(ByVal DateStr)
    YearMonStr = Split(DateStr,"月", - 1,1)
    YeraName = Left(YearMonStr(0),4)
    MonName = Mid(YearMonStr(0),6)
    ToBigDate = YearChange(YeraName) & MonthChange(MonName)
End Function

'获取当前系统时间
Function GetNowTime()
    GetNowTime = FormatDateTime(Now(),1)
End Function' GetNowTime

'选择指定地物并返回个数
Function SelFeatures(ByVal EngLayerName,ByRef Count,ByRef IdArr()) 'EngLayerName 图层名称(英文),Count 个数(返回值),IdArr() Id数组(返回值)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", EngLayerName
    SSProcess.SetSelectCondition "SSObj_Type", "==", "LINE"
    SSProcess.SelectFilter
    Count = SSProcess.GetSelGeoCount
    ReDim IdArr(Count)
    For i = 0 To Count - 1
        IdArr(i) = SSProcess.GetSelGeoValue(i,"SSObj_ID")
    Next 'i
End Function' SelFeatures

'获取所有记录
Function GetSQLRecordAll(ByVal StrSqlStatement, ByRef SQLRecord(), ByRef iRecordCount)
    ProJectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb ProJectName
    If StrSqlStatement = "" Then
        MsgBox "查询语句为空，操作停止！",48
    End If
    iRecordCount =  - 1
    SSProcess.OpenAccessRecordset ProJectName, StrSqlStatement
    RecordCount = SSProcess.GetAccessRecordCount (ProJectName, StrSqlStatement)
    If RecordCount > 0 Then
        iRecordCount = 0
        ReDim SQLRecord(RecordCount)
        SSProcess.AccessMoveFirst ProJectName, StrSqlStatement
        iRecordCount = 0
        While SSProcess.AccessIsEOF (ProJectName, StrSqlStatement) = 0
            fields = ""
            values = ""
            SSProcess.GetAccessRecord ProJectName, StrSqlStatement, fields, values
            SQLRecord(iRecordCount) = values
            iRecordCount = iRecordCount + 1
            SSProcess.AccessMoveNext ProJectName, StrSqlStatement
        WEnd
    End If
    SSProcess.CloseAccessRecordset ProJectName, StrSqlStatement
    SSProcess.CloseAccessMdb ProJectName
End Function

'获取单元格值
Function GetSelCellVal(ByVal CellContent)
    GetSelCellVal = Left(CellContent,Len(CellContent) - 1)
End Function' GetSelCellVal

'数据类型转换
Function Transform(ByVal Values)
    If Values <> "" Then
        Values = CDbl(Values)
    Else
        Values = 0
        Exit Function
    End If
    Transform = Values
End Function'Transform

'完成提示
Function Ending()
    MsgBox "输出完成"
End Function' Ending