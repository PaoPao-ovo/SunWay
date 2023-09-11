Dim XiangMMC
Dim tihuanzdz
Dim f_docName
Dim outputMode
fwhTableName = "FC_LPB_户信息表"
zrzTableName = "ZRZ_LP_信息表"
'ado 全局变量
Dim adoConnection
'Aspose.Word 全局变量
Dim g_docObj
Dim ProjectName
Dim rptPathName
jmzbArray = Array("架空","架空层")
Sub OnClick()
    ' ProjectName = SSProcess.GetProjectFileName()
    ' rptPathName = Left(ProjectName,InStrRev(ProjectName,"\")) & "成果文件\房产测量报告\"
    ' If IsfolderExists(rptPathName) = False  Then CreateFolders rptPathName
    ' If rptPathName = "" Then Exit Sub
    ' ExtractYTarea
    ' '=============获取宗地属性
    ' 'filename=SSProcess.GetSysPathName(5)
    ' SQL = "SELECT YeWLX,XiangMMC,HeTBH,JianSDW,CeLDW,CeLRQ FROM ZD_XM信息属性表 WHERE ID =1"
    ' GetSQLRecordAll ProjectName,SQL,ZDRecord,RecordZDCount
    ' If RecordZDCount = 1 Then
    '     '================理论只有一宗地
    '     zdsx = Split(ZDRecord(0),",")
    '     YeWLX = zdsx(0)
    '     XiangMMC = zdsx(1)
    '     HeTBH = zdsx(2)
    '     JianSDW = zdsx(3)
    '     CeLDW = zdsx(4)
    '     CeLRQ = zdsx(5)
    '     tihuanzdz = XiangMMC & "$" & HeTBH & "$" & JianSDW & "$" & CeLDW & "$" & CeLRQ
    '     'msgbox  tihuanzdz
    '     '==========================根据不同的业务类型，选择不同的输出模板输出报告
    '     If YeWLX = "房开项目"  Then
    '         SSProcess.ClearInputParameter
    '         SSProcess.AddInputParameter "输出模式", "单体", 0, "单体,汇总", "【单体】：每个自然幢输出一份房产测绘成果" & vbCrLf & "【汇总】：输出所有自然幢汇总房产测绘成果"
    '         result = SSProcess.ShowInputParameterDlg ("输出模式")
    '         If result = 1 Then
    '             SSProcess.UpdateScriptDlgParameter 1
    '             outputMode = SSProcess.GetInputParameter ("输出模式" )
    '             '==========================输出单体房产报告（房开）
    '             If  outputMode = "单体"    Then
    '                 f_docName = "单体房产测量报告_房开.docx"
    '             Else
    '                 f_docName = "单体房产测量报告_房开.docx"
    '             End If
    '             '============================获取自然幢的个数，循环输出
    '             getzrz
    '         Else
    '             '未选择输出模式，推出
    '             MsgBox "未选择输出模式，已退出"
    '             Exit Sub
    '         End If
    '     End If
    
    ' Else
    '     '================多宗地进行添加
    
    ' End If
    ' FCFK_FXArea
    FCFK_TS
End Sub

'房产测量报告_房开（分项面积统计表）
Function FCFK_FXArea()
    
    Set g_docObj = CreateObject("asposewordscom.asposewordshelper")
    
    '模板路径
    Dim TamplateFilePath
    
    '输出路径
    Dim FilePath
    
    '表索引(分项面积统计表)
    Dim TableIndex
    
    '地上起始行
    Dim DS_StartRow
    
    '地上结束行
    Dim DS_EndRow
    
    '最大个数（超过需要加列）
    Dim MaxColumn
    
    '参数初始化
    TamplateFilePath = SSProcess.GetSysPathName(7) & "输出模板\房产测量报告_房开.docx"
    FilePath = SSProcess.GetSysPathName(5) & "成果文件\房产测量报告\单体房产测量报告_房开.docx"
    TableIndex = 3
    DS_StartRow = 1
    DS_EndRow = 1
    MaxColumn = 14
    
    '根据模板创建Word文档
    g_docObj.CreateDocumentByTemplate TamplateFilePath
    
    '获取所有的幢【ZRZH】并填值
    SqlStr = "Select DISTINCT ZRZ_LP_信息表.ZRZH,ZRZGUID From ZRZ_LP_信息表 Where ZRZ_LP_信息表.ID > 0 "
    MdbName = SSProcess.GetProjectFileName
    GetSQLRecordAll MdbName,SqlStr,InfoArr,ZrZhCount
    If ZrZhCount <= MaxColumn Then
        For ColumnIndex = 1 To ZrZhCount
            ZRZHArr = Split((InfoArr(ColumnIndex - 1)),",", - 1,1)
            g_docObj.SetCellText TableIndex,0,ColumnIndex,ZRZHArr(0),True,False
        Next 'i
    Else
        g_docObj.InsertTableColumn TableIndex,ZrZhCount - MaxColumn,False
        MaxColumn = ZrZhCount
        For ColumnIndex = 1 To ZrZhCount
            g_docObj.SetCellText TableIndex,0,ColumnIndex,ZRZHArr(0),True,False
        Next 'i
    End If
    
    '获取所有的使用功能
    SqlStr = "Select DISTINCT FC_LPB_户信息表.SYGN From FC_LPB_户信息表 Where FC_LPB_户信息表.ID > 0 And FC_LPB_户信息表.CH > 0 And FC_LPB_户信息表.SHBW Not Like '*计入地下*' "
    GetSQLRecordAll MdbName,SqlStr,SYGNArr,SYGNCount
    
    DS_EndRow = SYGNCount - 1 + DS_StartRow
    
    g_docObj.CloneTableRow TableIndex,DS_StartRow,1,SYGNCount - 1,False
    
    '填写使用功能
    For i = 0 To SYGNCount - 1
        g_docObj.SetCellText TableIndex,DS_StartRow,1,SYGNArr(i),True,False
        DS_StartRow = DS_StartRow + 1
    Next 'i
    
    For i = 0 To ZrZhCount - 1
        '获取每一幢每一个列别的面积和并填值
        GUIDArr = Split((InfoArr(i)),",", - 1,1)
        DS_StartRow = 1
        For j = 0 To SYGNCount - 1
            SqlStr = "Select Sum(FC_LPB_户信息表.JZMJ) From FC_LPB_户信息表 Where FC_LPB_户信息表.ZRZGUID = " & GUIDArr(1) & " And FC_LPB_户信息表.SYGN = " & "'" & SYGNArr(j) & "'" & " And FC_LPB_户信息表.CH > 0 And FC_LPB_户信息表.SHBW Not Like '*计入地下*' "
            GetSQLRecordAll MdbName,SqlStr,AreaArr,SumCount
            SqlStr = ""
            If AreaArr(0) <> "" Then
                g_docObj.SetCellText TableIndex,DS_StartRow,i + 2,GetFormatNumber(AreaArr(0),2),True,False
            End If
            DS_StartRow = DS_StartRow + 1
        Next 'j
    Next 'i
    
    '地下起始行
    Dim DX_StartRow
    
    '地下结束行
    Dim DX_EndRow
    
    '参数初始化
    DX_StartRow = DS_EndRow + 3
    DX_EndRow = DX_StartRow
    
    '获取所有的使用功能
    SqlStr = "Select DISTINCT FC_LPB_户信息表.SYGN From FC_LPB_户信息表 Where FC_LPB_户信息表.ID > 0 And FC_LPB_户信息表.CH < 0 "
    GetSQLRecordAll MdbName,SqlStr,SYGNArr1,SYGNCount1
    
    SqlStr = "Select DISTINCT FC_LPB_户信息表.SYGN From FC_LPB_户信息表 Where FC_LPB_户信息表.ID > 0 And FC_LPB_户信息表.CH > 0 And FC_LPB_户信息表.SHBW  Like '*计入地下*' "
    GetSQLRecordAll MdbName,SqlStr,SYGNArr2,SYGNCount2
    
    SYGNCount = SYGNCount1 + SYGNCount2 - 1
    
    ReDim SYGNArr(SYGNCount1 + SYGNCount2 - 2)
    
    For i = 0 To SYGNCount1 + SYGNCount2 - 2
        If i <= SYGNCount1 - 1 Then
            SYGNArr(i) = SYGNArr1(i)
        Else
            SYGNArr(i) = SYGNArr2(i - SYGNCount1)
        End If
    Next 'i
    
    DX_EndRow = SYGNCount - 1 + DX_EndRow
    
    g_docObj.CloneTableRow TableIndex,DX_StartRow,1,SYGNCount - 1,False
    
    '填写使用功能
    For i = 0 To SYGNCount - 1
        g_docObj.SetCellText TableIndex,DX_StartRow,1,SYGNArr(i),True,False
        DX_StartRow = DX_StartRow + 1
    Next 'i
    
    For i = 0 To ZrZhCount - 1
        '获取每一幢每一个列别的面积和并填值
        GUIDArr = Split((InfoArr(i)),",", - 1,1)
        DX_StartRow = DS_EndRow + 3
        For j = 0 To SYGNCount - 1
            SqlStr = "Select Sum(FC_LPB_户信息表.JZMJ) From FC_LPB_户信息表 Where FC_LPB_户信息表.ZRZGUID = " & GUIDArr(1) & " And FC_LPB_户信息表.SYGN = " & "'" & SYGNArr(j) & "'" & " And FC_LPB_户信息表.CH < 0 "
            GetSQLRecordAll MdbName,SqlStr,AreaArr1,SumCount1
            SqlStr = "Select Sum(FC_LPB_户信息表.JZMJ) From FC_LPB_户信息表 Where FC_LPB_户信息表.ZRZGUID = " & GUIDArr(1) & " And FC_LPB_户信息表.SYGN = " & "'" & SYGNArr(j) & "'" & " And FC_LPB_户信息表.CH > 0 And FC_LPB_户信息表.SHBW Like '*计入地下*'"
            GetSQLRecordAll MdbName,SqlStr,AreaArr2,SumCount2
            Area = Transform(AreaArr1(0)) + Transform(AreaArr2(0))
            If Area - 0 <> 0 Then
                g_docObj.SetCellText TableIndex,DX_StartRow,i + 2,GetFormatNumber(Area,2),True,False
            End If
            DX_StartRow = DX_StartRow + 1
        Next 'j
    Next 'i
    
    '重新初始化
    DS_StartRow = 1
    DX_StartRow = DS_EndRow + 3
    
    '填写地上小计
    For i = 0 To ZrZhCount - 1
        DS_Sum = 0
        For j = DS_StartRow To DS_EndRow
            SingleArea = Transform(GetSelCellVal(g_docObj.GetCellText(TableIndex,j,i + 2,False)))
            DS_Sum = DS_Sum + SingleArea
        Next 'j
        If DS_Sum - 0 <> 0 Then
            g_docObj.SetCellText TableIndex,DS_EndRow + 1,i + 2,DS_Sum,True,False
        End If
    Next 'i
    
    '填写地下小计
    For i = 0 To ZrZhCount - 1
        DX_Sum = 0
        For j = DX_StartRow To DX_EndRow
            SingleArea = Transform(GetSelCellVal(g_docObj.GetCellText(TableIndex,j,i + 2,False)))
            DX_Sum = DX_Sum + SingleArea
        Next 'j
        If DX_Sum - 0 <> 0 Then
            g_docObj.SetCellText TableIndex,DX_EndRow + 1,i + 2,DX_Sum,True,False
        End If
    Next 'i
    
    '填写不架空总建筑面积合计
    For i = 0 To ZrZhCount - 1
        GUIDArr = Split((InfoArr(i)),",", - 1,1)
        SqlStr = "Select Sum(FC_LPB_户信息表.JZMJ) From FC_LPB_户信息表 Where FC_LPB_户信息表.ZRZGUID = " & GUIDArr(1) & " And FC_LPB_户信息表.SYGN Not Like '*架空*'"
        GetSQLRecordAll MdbName,SqlStr,FJKArr,FJKCount
        If FJKArr(0) <> "" Then
            g_docObj.SetCellText TableIndex,DX_EndRow + 2,i + 1,GetFormatNumber(FJKArr(0),2),True,False
        End If
    Next 'i
    
    '填写架空总建筑面积合计
    For i = 0 To ZrZhCount - 1
        GUIDArr = Split((InfoArr(i)),",", - 1,1)
        SqlStr = "Select Sum(FC_LPB_户信息表.JZMJ) From FC_LPB_户信息表 Where FC_LPB_户信息表.ZRZGUID = " & GUIDArr(1) & " And FC_LPB_户信息表.SYGN Like '*架空*'"
        GetSQLRecordAll MdbName,SqlStr,JKArr,FJKCount
        If JKArr(0) <> "" Then
            g_docObj.SetCellText TableIndex,DX_EndRow + 3,i + 1,GetFormatNumber(JKArr(0),2),True,False
        End If
    Next 'i
    
    '填写合计
    
    '合计列
    Dim HjColmun
    HjColmun = MaxColumn + 2
    
    For i = DS_StartRow To DS_EndRow + 1
        DS_HJ = 0
        For j = 0 To ZrZhCount - 1
            DS_HJ = DS_HJ + Transform(GetSelCellVal(g_docObj.GetCellText(TableIndex,i,j + 2,False)))
        Next 'j
        If DS_HJ - 0 <> 0 Then
            g_docObj.SetCellText TableIndex,i,HjColmun,GetFormatNumber(DS_HJ,2),True,False
        End If
    Next 'i
    
    For i = DX_StartRow To DX_EndRow + 1
        DX_HJ = 0
        For j = 0 To ZrZhCount - 1
            DX_HJ = DX_HJ + Transform(GetSelCellVal(g_docObj.GetCellText(TableIndex,i,j + 2,False)))
        Next 'j
        If DX_HJ - 0 <> 0 Then
            g_docObj.SetCellText TableIndex,i,HjColmun,GetFormatNumber(DX_HJ,2),True,False
        End If
    Next 'i
    
    For i = 0 To ZrZhCount - 1
        JK_HJ = JK_HJ + Transform(GetSelCellVal(g_docObj.GetCellText(TableIndex,DX_EndRow + 2,i + 1,False)))
        FJK_HJ = FJK_HJ + Transform(GetSelCellVal(g_docObj.GetCellText(TableIndex,DX_EndRow + 3,i + 1,False)))
    Next 'i
    
    If JK_HJ - 0 <> 0 Then
        g_docObj.SetCellText TableIndex,DX_EndRow + 2,HjColmun - 1,GetFormatNumber(JK_HJ,2),True,False
    End If
    
    If FJK_HJ - 0 <> 0 Then
        g_docObj.SetCellText TableIndex,DX_EndRow + 3,HjColmun - 1,GetFormatNumber(FJK_HJ,2),True,False
    End If
    
    '保存文档
    g_docObj.SaveEx FilePath
    
End Function' FCFK_FXArea

'房产测量报告_房开（套数统计表）
Function FCFK_TS()
    
    Set g_docObj = CreateObject("asposewordscom.asposewordshelper")
    
    '模板路径
    Dim TamplateFilePath
    
    '输出路径
    Dim FilePath
    
    '表索引(套数统计表)
    Dim TableIndex
    
    '地上起始行
    Dim DS_StartRow
    
    '地上结束行
    Dim DS_EndRow
    
    '最大个数（超过需要加列）
    Dim MaxColumn
    
    '参数初始化
    TamplateFilePath = SSProcess.GetSysPathName(7) & "输出模板\房产测量报告_房开.docx"
    FilePath = SSProcess.GetSysPathName(5) & "成果文件\房产测量报告\单体房产测量报告_房开.docx"
    TableIndex = 4
    DS_StartRow = 1
    DS_EndRow = 1
    MaxColumn = 14
    
    
    '根据模板创建Word文档
    g_docObj.CreateDocumentByTemplate TamplateFilePath
    
    '获取所有的幢【ZRZH】并填值
    SqlStr = "Select DISTINCT ZRZ_LP_信息表.ZRZH,ZRZGUID From ZRZ_LP_信息表 Where ZRZ_LP_信息表.ID > 0 "
    MdbName = SSProcess.GetProjectFileName
    GetSQLRecordAll MdbName,SqlStr,InfoArr,ZrZhCount
    If ZrZhCount <= MaxColumn Then
        For ColumnIndex = 1 To ZrZhCount
            ZRZHArr = Split((InfoArr(ColumnIndex - 1)),",", - 1,1)
            g_docObj.SetCellText TableIndex,0,ColumnIndex,ZRZHArr(0),True,False
        Next 'i
    Else
        g_docObj.InsertTableColumn TableIndex,ZrZhCount - MaxColumn,False
        MaxColumn = ZrZhCount
        For ColumnIndex = 1 To ZrZhCount
            g_docObj.SetCellText TableIndex,0,ColumnIndex,ZRZHArr(0),True,False
        Next 'i
    End If
    
    '获取所有的使用功能
    SqlStr = "Select DISTINCT FC_LPB_户信息表.SYGN From FC_LPB_户信息表 Where FC_LPB_户信息表.ID > 0 And FC_LPB_户信息表.CH > 0 And FC_LPB_户信息表.SHBW Not Like '*计入地下*' "
    GetSQLRecordAll MdbName,SqlStr,SYGNArr,SYGNCount
    
    DS_EndRow = SYGNCount - 1 + DS_StartRow
    
    g_docObj.CloneTableRow TableIndex,DS_StartRow,1,SYGNCount - 1,False
    
    '填写使用功能
    For i = 0 To SYGNCount - 1
        g_docObj.SetCellText TableIndex,DS_StartRow,1,SYGNArr(i),True,False
        DS_StartRow = DS_StartRow + 1
    Next 'i
    
    For i = 0 To ZrZhCount - 1
        '获取每一幢每一个列别的个数并填值
        GUIDArr = Split((InfoArr(i)),",", - 1,1)
        DS_StartRow = 1
        For j = 0 To SYGNCount - 1
            SqlStr = "Select FC_LPB_户信息表.ID From FC_LPB_户信息表 Where FC_LPB_户信息表.ZRZGUID = " & GUIDArr(1) & " And FC_LPB_户信息表.SYGN = " & "'" & SYGNArr(j) & "'" & " And FC_LPB_户信息表.CH > 0  And FC_LPB_户信息表.SHBW Not Like '*计入地下*' "
            GetSQLRecordAll MdbName,SqlStr,AreaArr,SumCount
            If SumCount > 0 Then
                g_docObj.SetCellText TableIndex,DS_StartRow,i + 2,SumCount,True,False
            End If
            DS_StartRow = DS_StartRow + 1
        Next 'j
    Next 'i
    
    '地下起始行
    Dim DX_StartRow
    
    '地下结束行
    Dim DX_EndRow
    
    '参数初始化
    DX_StartRow = DS_EndRow + 3
    DX_EndRow = DX_StartRow
    
    '获取所有的使用功能
    SqlStr = "Select DISTINCT FC_LPB_户信息表.SYGN From FC_LPB_户信息表 Where FC_LPB_户信息表.ID > 0 And FC_LPB_户信息表.CH < 0 "
    GetSQLRecordAll MdbName,SqlStr,SYGNArr1,SYGNCount1
    
    SqlStr = "Select DISTINCT FC_LPB_户信息表.SYGN From FC_LPB_户信息表 Where FC_LPB_户信息表.ID > 0 And FC_LPB_户信息表.CH > 0 And FC_LPB_户信息表.SHBW  Like '*计入地下*' "
    GetSQLRecordAll MdbName,SqlStr,SYGNArr2,SYGNCount2
    
    SYGNCount = SYGNCount1 + SYGNCount2 - 1
    
    ReDim SYGNArr(SYGNCount1 + SYGNCount2 - 2)
    
    For i = 0 To SYGNCount1 + SYGNCount2 - 2
        If i <= SYGNCount1 - 1 Then
            SYGNArr(i) = SYGNArr1(i)
        Else
            SYGNArr(i) = SYGNArr2(i - SYGNCount1)
        End If
    Next 'i
    
    DX_EndRow = SYGNCount - 1 + DX_EndRow
    
    g_docObj.CloneTableRow TableIndex,DX_StartRow,1,SYGNCount - 1,False
    
    '填写使用功能
    For i = 0 To SYGNCount - 1
        g_docObj.SetCellText TableIndex,DX_StartRow,1,SYGNArr(i),True,False
        DX_StartRow = DX_StartRow + 1
    Next 'i
    
    For i = 0 To ZrZhCount - 1
        '获取每一幢每一个列别的个数和并填值
        GUIDArr = Split((InfoArr(i)),",", - 1,1)
        DX_StartRow = DS_EndRow + 3
        For j = 0 To SYGNCount - 1
            SqlStr = "Select FC_LPB_户信息表.ID From FC_LPB_户信息表 Where FC_LPB_户信息表.ZRZGUID = " & GUIDArr(1) & " And FC_LPB_户信息表.SYGN = " & "'" & SYGNArr(j) & "'" & " And FC_LPB_户信息表.CH < 0 "
            GetSQLRecordAll MdbName,SqlStr,AreaArr1,SumCount1
            SqlStr = "Select Sum(FC_LPB_户信息表.JZMJ) From FC_LPB_户信息表 Where FC_LPB_户信息表.ZRZGUID = " & GUIDArr(1) & " And FC_LPB_户信息表.SYGN = " & "'" & SYGNArr(j) & "'" & " And FC_LPB_户信息表.CH > 0 And FC_LPB_户信息表.SHBW Like '*计入地下*'"
            GetSQLRecordAll MdbName,SqlStr,AreaArr2,SumCount2
            SumCount = SumCount1 + SumCount2
            If SumCount > 0 Then
                g_docObj.SetCellText TableIndex,DX_StartRow,i + 2,SumCount,True,False
            End If
            DX_StartRow = DX_StartRow + 1
        Next 'j
    Next 'i
    
    '重新初始化
    DS_StartRow = 1
    DX_StartRow = DS_EndRow + 3
    
    '填写地上小计
    For i = 0 To ZrZhCount - 1
        DS_Sum = 0
        For j = DS_StartRow To DS_EndRow
            SingleCount = Transform(GetSelCellVal(g_docObj.GetCellText(TableIndex,j,i + 2,False)))
            DS_Sum = DS_Sum + SingleCount
        Next 'j
        If DS_Sum - 0 <> 0 Then
            g_docObj.SetCellText TableIndex,DS_EndRow + 1,i + 2,DS_Sum,True,False
        End If
    Next 'i
    
    '填写地下小计
    For i = 0 To ZrZhCount - 1
        DX_Sum = 0
        For j = DX_StartRow To DX_EndRow
            SingleCount = Transform(GetSelCellVal(g_docObj.GetCellText(TableIndex,j,i + 2,False)))
            DX_Sum = DX_Sum + SingleCount
        Next 'j
        If DX_Sum - 0 <> 0 Then
            g_docObj.SetCellText TableIndex,DX_EndRow + 1,i + 2,DX_Sum,True,False
        End If
    Next 'i
    
    '填写合计
    
    '合计列
    Dim HjColmun
    HjColmun = MaxColumn + 2
    
    For i = DS_StartRow To DS_EndRow + 1
        DS_HJ = 0
        For j = 0 To ZrZhCount - 1
            DS_HJ = DS_HJ + Transform(GetSelCellVal(g_docObj.GetCellText(TableIndex,i,j + 2,False)))
        Next 'j
        If DS_HJ - 0 <> 0 Then
            g_docObj.SetCellText TableIndex,i,HjColmun,DS_HJ,True,False
        End If
    Next 'i
    
    For i = DX_StartRow To DX_EndRow + 1
        DX_HJ = 0
        For j = 0 To ZrZhCount - 1
            DX_HJ = DX_HJ + Transform(GetSelCellVal(g_docObj.GetCellText(TableIndex,i,j + 2,False)))
        Next 'j
        If DX_HJ - 0 <> 0 Then
            g_docObj.SetCellText TableIndex,i,HjColmun,DX_HJ,True,False
        End If
    Next 'i
    
    For i = 0 To ZrZhCount - 1
        If Transform(GetSelCellVal(g_docObj.GetCellText(TableIndex,DS_EndRow + 1,i + 2,False))) + Transform(GetSelCellVal(g_docObj.GetCellText(TableIndex,DX_EndRow + 1,i + 2,False))) - 0 <> 0 Then
            g_docObj.SetCellText TableIndex,DX_EndRow + 2,i + 1,Transform(GetSelCellVal(g_docObj.GetCellText(TableIndex,DS_EndRow + 1,i + 2,False))) + Transform(GetSelCellVal(g_docObj.GetCellText(TableIndex,DX_EndRow + 1,i + 2,False))),True,False
        End If
    Next 'i
    
    For i = 0 To ZrZhCount - 1
        HJ_Num = HJ_Num + Transform(GetSelCellVal(g_docObj.GetCellText(TableIndex,DX_EndRow + 2,i + 1,False)))
    Next 'i
    
    If HJ_Num - 0 <> 0 Then
        g_docObj.SetCellText TableIndex,DX_EndRow + 2,HjColmun - 1,HJ_Num,True,False
    End If
    
    g_docObj.SaveEx FilePath
    
End Function' FCFK_FXArea

'获取单元格值
Function GetSelCellVal(ByVal CellContent)
    GetSelCellVal = Left(CellContent,Len(CellContent) - 1)
End Function' GetSelCellVal

'数值转换
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

'============================获取自然幢的个数，循环输出
Function getzrz
    SQL = "SELECT QiuH,ZRZGUID,ZRZH,JGRQ,JZMJ,CHLB,QiuH,FWJGNAME,ZCS,DSCS,DXCS FROM ZRZ_LP_信息表 WHERE ID <>0"
    GetSQLRecordAll ProjectName,SQL,ZrzRecord,RecordZrzCount
    If RecordZrzCount > 0 Then
        zrzcount = RecordZrzCount
        For j = 0 To  zrzcount - 1
            '========================处理替换属性值
            ZRZSXZ = Split(ZrzRecord(J),",")
            ZRZH = ZRZSXZ(2)
            ZRZGUID = ZRZSXZ(1)
            QiuH = ZRZSXZ(0)
            For K = 2 To UBound(ZRZSXZ)
                tihuanzdz = tihuanzdz & "$" & ZRZSXZ(k)
            Next
            OutputFCBG rptPathName,outputMode,ZRZH,ZRZGUID,QiuH
        Next
    Else
        MsgBox "不存在自然幢信息，已退出"
        Exit Function
    End If
End Function

'================================替换函数
Function ReplaceTableField()
    ZDFIELD = "{ZD_XM信息属性表.XiangMMC},{ZD_XM信息属性表.HeTBH},{ZD_XM信息属性表.JianSDW},{ZD_XM信息属性表.CeLDW},{ZD_XM信息属性表.CeLRQ},{ZRZ_LP_信息表.ZRZH},{ZRZ_LP_信息表.JGRQ},{ZRZ_LP_信息表.JZMJ},{ZRZ_LP_信息表.CHLB},{ZRZ_LP_信息表.QiuH},{ZRZ_LP_信息表.FWJGNAME},{ZRZ_LP_信息表.ZCS},{ZRZ_LP_信息表.DSCS},{ZRZ_LP_信息表.DXCS}"
    oldZDFIELD = Split(ZDFIELD,",")
    replaceZDFIELD = Split(tihuanzdz,"$")
    For i = 0 To UBound(oldZDFIELD)
        g_docObj.Replace oldZDFIELD(i),replaceZDFIELD(i),0
    Next
End Function

'================================输出报告函数
Function OutputFCBG(ByVal rptPathName,ByVal outputMode,ZRZH,ZRZGUID,QiuH)
    If  outputMode = "汇总"  Then   f_docName = "房产测量报告_房开.docx" 'else f_docName="单体房产测量报告_民房不动产.docx"
    If  outputMode = "单体"  Then   f_docName = "单体房产测量报告_房开.docx"
    If  outputMode = "房产报告"  Then   f_docName = "单体房产测量报告_民房竣工.docx"
    If  outputMode = "不动产房产报告"  Then   f_docName = "单体房产测量报告_民房不动产.docx"
    If  outputMode = "单一产权"  Then   f_docName = "房产测量报告_厂房.docx"
    strDocFileName = SSProcess.GetSysPathName (7) & "输出模板\" & f_docName
    Set g_docObj = CreateObject ("asposewordscom.asposewordshelper")
    If  TypeName (g_docObj) = "AsposeWordsHelper" Then
        g_docObj.CreateDocumentByTemplate strDocFileName
    Else
        MsgBox "请先注册Aspose.Word插件"
        Exit Function
    End If
    
    If outputMode = "单体" Then
        'ReplaceTableField  zdTableName,"", "SpatialData", "2"
        OutputTable1 1,ZRZGUID,outputMode
        OutputFWSYQ  3,ZRZGUID,outputMode,QiuH
        ReplaceTableField
        g_docObj.SaveEx rptPathName & XiangMMC & zrzh & "房产测量报告.docx"
    End If
End Function
'============================属性处理
'//提取面积块（阳台）的面积到户上
Function ExtractYTarea()
    SQL = "select  BZGUID from FC_LPB_户信息表  WHERE  FC_LPB_户信息表.id >1 and HMJKLY LIKE '*阳*'"
    GetSQLRecordAll ProjectName,SQL,HRecord,RecordCount
    If RecordCount > 0 Then
        For i = 0 To RecordCount - 1
            BZGUID = HRecord(I)'户的编组ID
            sql1 = "Select SUM (FC_面积块信息属性表.JZMJ) From FC_面积块信息属性表 INNER JOIN GeoAreaTB ON FC_面积块信息属性表.ID = GeoAreaTB.ID WHERE FC_面积块信息属性表.BZGUID= '" & BZGUID & "' and ([GeoAreaTB].[Mark] Mod 2)<>0 AND FC_面积块信息属性表.MJKMC like '*阳*'"
            GetSQLRecordAll ProjectName,SQL1,mjkRecord,mjkRecordCount
            
            If  mjkRecordCount > 0 Then
                If mjkRecord(0) <> "" Then
                    sql3 = "update FC_LPB_户信息表 set ytjzmj=" & mjkRecord(0) & " where FC_LPB_户信息表.BZGUID='" & BZGUID & "'"
                    mdbName = SSProcess.GetProjectFileName
                    SSProcess.OpenAccessMdb mdbName
                    SSProcess.ExecuteAccessSql mdbName,sql3
                    SSProcess.CloseAccessMdb mdbName
                End If
                
            End If
        Next
    End If
End Function

'=========================================分摊表
Function OutPutFTB(ByVal section,ByVal zrzguid)
    
    
End Function

'//输出 公用配套建筑分摊系数计算表
' Function OutputTable7(ByVal section,ByVal zrzguid)
'     g_docObj.MoveToSection section
'     tableIndex = 0
'     iniRow = 1
'     iniCol = 1
'     startRow = iniRow
'     startCol = iniCol
'     allCols = g_docObj.GetTableColCount(tableIndex,0)

'     '************************************************获取公用部分面积
'     ftxmcCount = 0
'     ReDim ftxmcList(ftxmcCount)
'     publicCount = 0
'     ReDim publicList(publicCount)
'     strKeyName = "GetNewApportionInfo"
'     SSProcess.MapCallBackFunction strKeyName, ZRZGUID, flags
'     defaultValut = ""
'     SSParameter.GetParameterSTR "MapCallBackFunction", strKeyName, defaultValut, ftmsxx
'     ftmsxxList = Split(ftmsxx,"$")
'     For i = 0 To UBound(ftmsxxList)
'         ftms = ftmsxxList(i)
'         ftmsList = Split(ftms,"^")
'         If  UBound(ftmsList) = 9 Then
'             ftxmc = ftmsList(3)
'             ctfw = ftmsList(4)
'             ftlx = ftmsList(5)
'             gjmj = ftmsList(6)
'             sjftmj = ftmsList(7)
'             tnmj = ftmsList(8)
'             gjftxs = ftmsList(9)
'             gjftxs = GetFormatNumber(gjftxs,6)
'             gjmj = GetFormatNumber(gjmj,4)
'             sjftmj = GetFormatNumber(sjftmj,4)
'             '公用面积=公建面积+上级分摊面积
'             gymj = CDbl(gjmj) + CDbl(sjftmj)
'             gymj = GetFormatNumber(gymj,4)
'             If ftlx <> "不分摊" Then
'                 values = ftlx & "||" & ftxmc & "||" & gymj & "||" & gjftxs & "||" & gjmj & "||" & sjftmj
'                 ReDim Preserve ftxmcList(ftxmcCount)
'                 ftxmcList(ftxmcCount) = values
'                 ftxmcCount = ftxmcCount + 1
'                 If ftlx <> "大楼分摊" And ftlx <> "一级分摊" Then
'                     ReDim Preserve publicList(publicCount)
'                     publicList(publicCount) = values
'                     publicCount = publicCount + 1
'                 End If
'             End If
'         End If
'     Next

'     '计算删除表格多余列
'     deleteCount = (allCols - 4) - ftxmcCount
'     If deleteCount > 13 Then deleteCount = 13
'     endCol = allCols - deleteCount - 1
'     For i = 0 To deleteCount
'         g_docObj.DeleteCol tableIndex, 2,True
'     Next

'     g_docObj.SetCellText tableIndex,5,endCol,"分摊受益后建筑面积",True
'     g_docObj.SetCellText tableIndex,5,endCol - 1,"公用分摊系数",True
'     g_docObj.MergeCell tableIndex, 2,  endCol - 2,  4, endCol - 1 '合并 合计
'     g_docObj.MergeCell tableIndex, 0,  0,  0, endCol - 1 '合并 标题
'     g_docObj.SetCellText tableIndex,2,endCol - 2,"合计",True


'     '************************************************按分摊级别填充公用部分面积
'     For i = 0 To ftxmcCount - 1
'         values = ftxmcList(i)
'         valuesList = Split(values,"||")
'         ftlx = valuesList(0)
'         ftxmc = valuesList(1)
'         gymj = valuesList(2)
'         gjftxs = valuesList(3)
'         g_docObj.SetCellText tableIndex,startRow,startCol,ftlx,True
'         g_docObj.SetCellText tableIndex,startRow + 1,startCol,ftxmc,True
'         g_docObj.SetCellText tableIndex,startRow + 2,startCol,gymj,True
'         g_docObj.SetCellText tableIndex,startRow + 3,startCol,gjftxs,True
'         g_docObj.SetCellText tableIndex,startRow + 4,startCol + 1,"分摊受益",True
'         startCol = startCol + 1
'     Next
'     '合并分摊类型一致项
'     startCol = iniCol
'     allFtlx = ""
'     mergeIndex = 0
'     For i = 0 To ftxmcCount - 1
'         values = ftxmcList(i)
'         valuesList = Split(values,"||")
'         ftlx = valuesList(0)
'         If InStr(allFtlx,"'" & ftlx & "'") = 0 Then
'             mergeCount = 0
'             For j = 0 To ftxmcCount - 1
'                 values1 = ftxmcList(j)
'                 valuesList1 = Split(values1,"||")
'                 ftlx1 = valuesList1(0)
'                 If  ftlx1 = ftlx Then  mergeCount = mergeCount + 1
'             Next
'             If mergeCount > 1 Then
'                 g_docObj.MergeCell tableIndex, startRow,  startCol,  startRow, startCol + mergeCount - 1 '合并
'             End If
'             allFtlx = allFtlx & "," & "'" & ftlx & "'"
'             mergeIndex = mergeIndex + 1
'         End If
'         startCol = startCol + 1
'     Next
'     If mergeIndex > 0 Then g_docObj.MergeCell tableIndex, 1,  mergeIndex,  1, deleteCount + 2 '合并


'     '************************************************按分摊级别填充私有分组部分面积
'     SplitDomainEx zrzguid,domainValueList,domainCount,isLftjcqList
'     startRow = iniRow + 5
'     startCol = iniCol - 1
'     syCol = iniCol + 1
'     If domainCount > 1 Then  g_docObj.CloneTableRow tableIndex,  startRow, 1,domainCount - 1, True
'     sumCount = 0
'     ReDim  sumAreaList(sumCount)
'     sumTnArea = 0
'     For i = 0 To domainCount - 1
'         fwhids = domainValueList(i)
'         fwhidsList = Split(fwhids,",")
'         sumArea = 0
'         If fwhids <> "" Then
'             valueCount = GetProjectTableList (fwhTableName,"sum(round(sctnjzmj,4))",fwhTableName & ".ID IN (" & fwhids & ") ","SpatialData","2",valueList,fieldCount)
'             If valueCount = 1 Then sumArea = valueList(0,0)
'         End If
'         If IsNumeric(sumArea) = False Then sumArea = 0
'         If UBound(fwhidsList) >= 0 Then
'             sygn = SSProcess.GetObjectAttr (fwhidsList(0),"[sygn]")
'             If UBound(fwhidsList) = 0 Then str = ""  Else    str = "及其他" & sygn & ""
'             shbw = SSProcess.GetObjectAttr (fwhidsList(0),"[SHBW]") & str
'             ftxs = SSProcess.GetObjectAttr (fwhidsList(0),"[ftxs]")
'         Else
'             shbw = ""
'             ftxs = 0
'         End If
'         ftxs = GetFormatNumber(ftxs,6)
'         sumArea = GetFormatNumber(sumArea,4)
'         sumSyfthjzmj = 0'受益分摊后建筑面积
'         syCol = iniCol + 1
'         For j = 0 To ftxmcCount - 1
'             values = ftxmcList(j)
'             valuesList = Split(values,"||")
'             ftlx = valuesList(0)
'             ftxmc = valuesList(1)
'             gymj = valuesList(2)
'             gjftxs = valuesList(3)
'             ftsyArea = CDbl(sumArea) * CDbl(gjftxs)
'             ftsyArea = GetFormatNumber(ftsyArea,4)
'             If CDbl(ftsyArea) > CDbl(gymj) Then ftsyArea = gymj
'             ftsyArea = GetFormatNumber(ftsyArea,4)

'             isftsy = False
'             '判断当前私有分组是否分摊受益
'             For jj = 0 To UBound(fwhidsList)
'                 CTXMCLB = SSProcess.GetObjectAttr (fwhidsList(jj),"[CTXMCLB]")
'                 CTXMCLBList = Split(CTXMCLB,":")
'                 For jjj = 0 To UBound(CTXMCLBList)
'                     allPublic = CTXMCLBList(jjj)
'                     allPublicList = Split(allPublic,"->")
'                     If UBound(allPublicList) = 4 Then
'                         publicName = allPublicList(3)
'                         If publicName = ftxmc Then
'                             isftsy = True
'                             Exit For
'                         End If
'                     Next
'                     If isftsy = True Then Exit For
'                 Next
'                 If isftsy = True Then
'                     g_docObj.SetCellText tableIndex,startRow,syCol,ftsyArea,True
'                     sumSyfthjzmj = sumSyfthjzmj + CDbl(ftsyArea)
'                     If ftlx <> "一级分摊" Then
'                         publicValues = ftxmc & "||" & syCol & "||" & sumArea
'                         ReDim Preserve sumAreaList(sumCount)
'                         sumAreaList(sumCount) = publicValues
'                         sumCount = sumCount + 1
'                     End If
'                 End If
'                 syCol = syCol + 1
'             Next
'             syfthjzmj = CDbl(sumSyfthjzmj) + CDbl(sumArea)
'             syfthjzmj = GetFormatNumber(syfthjzmj,2)

'             g_docObj.SetCellText tableIndex,startRow,startCol,shbw,True
'             g_docObj.SetCellText tableIndex,startRow,startCol + 1,sumArea,True
'             g_docObj.SetCellText tableIndex,startRow,endCol - 1,ftxs,True
'             g_docObj.SetCellText tableIndex,startRow,endCol,syfthjzmj,True
'             startRow = startRow + 1
'             sumTnArea = sumTnArea + CDbl(sumArea)
'         Next

'         '************************************************填充公用部位面积
'         sumGjmj = 0
'         If domainCount = 0 Then     startRow = startRow + 2  Else  startRow = startRow + 1
'         If publicCount > 1 Then  g_docObj.CloneTableRow tableIndex,  startRow, 1,publicCount - 1, True
'         For i = 0 To publicCount - 1
'             values = publicList(i)
'             valuesList = Split(values,"||")
'             ftlx = valuesList(0)
'             ftxmc = valuesList(1)
'             gjmj = valuesList(4)
'             sjftmj = valuesList(5)
'             g_docObj.SetCellText tableIndex,startRow,startCol,ftxmc,True
'             g_docObj.SetCellText tableIndex,startRow,startCol + 1,gjmj,True
'             g_docObj.SetCellText tableIndex,startRow,startCol + 2,sjftmj,True
'             startRow = startRow + 1
'             sumGjmj = sumGjmj + CDbl(gjmj)
'         Next

'         '************************************************汇总分摊受益套内面积
'         If publicCount = 0 Then startRow = startRow + 1
'         startCol = iniCol + 2
'         sumSYTNArea = 0
'         allCol = ""
'         For  i = 0 To sumCount - 1
'             publicValues = sumAreaList(i)
'             publicValuesList = Split(publicValues,"||")
'             ftxmc = publicValuesList(0)
'             col = publicValuesList(1)
'             area = publicValuesList(2)
'             sumArea = 0
'             If InStr(allCol,"'" & col & "'") = 0 Then
'                 allCol = allCol & "," & "'" & col & "'"
'                 For j = 0 To sumCount - 1
'                     publicValues1 = sumAreaList(j)
'                     publicValuesList1 = Split(publicValues1,"||")
'                     ftxmc1 = publicValuesList1(0)
'                     col1 = publicValuesList1(1)
'                     area1 = publicValuesList1(2)
'                     If col = col1 Then
'                         sumArea = sumArea + CDbl(area1)
'                     End If
'                 Next
'                 g_docObj.SetCellText tableIndex,startRow,col - 1,sumArea,True
'                 sumSYTNArea = sumSYTNArea + CDbl(sumArea)
'             End If
'         Next
'         sumAllArea = CDbl(sumGjmj) + CDbl(sumTnArea)
'         sumAllArea = GetFormatNumber(sumAllArea,2)
'         g_docObj.SetCellText tableIndex,startRow,1,sumAllArea,True
'     End Function

'     '//输出 房屋所有权面积测算汇总表
'     Function OutputFWSYQ(ByVal section,ByVal zrzguid,ByVal outputMode,ByVal qiuh)
'         Dim MyArray() '首先定义一个一维动态数组
'         sumFwhArea = 0
'         dhIndex = 1
'         g_docObj.MoveToSection section
'         tableIndex = 0
'         iniRow = 5
'         iniCol = 0
'         startRow = iniRow
'         startCol = iniCol
'         SQL = "select SHBW,JZMJ,TNJZMJ,YTJZMJ,FTXS,FTJZMJ,CH,SYGN from FC_LPB_户信息表  WHERE  FC_LPB_户信息表.id >1"
'         GetSQLRecordAll ProjectName,SQL,HRecord,RecordCount
'         If RecordCount > 0 Then
'             fwhCount = RecordCount
'             copyRows = 28
'             For i = 0 To fwhCount - 1
'                 If i > 0 And i Mod copyRows = 0 Then
'                     g_docObj.CloneTable  tableIndex,  1
'                 End If
'             Next
'             'redim fwhList(fwhCount,8)
'             bgs = fwhCount \ copyRows
'             bgs = bgs + 1
'             szdx = bgs * 28
'             ReDim MyArray(szdx,8)
'             For I = 0 To RecordCount - 1
'                 HUSXZ = Split(HRecord(i),",")
'                 shbw = HUSXZ(0)
'                 scjzmj = GetFormatNumber(HUSXZ(1),2)
'                 scytjzmj = GetFormatNumber(HUSXZ(3),2)
'                 scftxs = GetFormatNumber(HUSXZ(4),6)
'                 scftjzmj = GetFormatNumber(HUSXZ(5),2)
'                 ch = HUSXZ(6)
'                 sygn = HUSXZ(7)
'                 sctnjzmj = GetFormatNumber(HUSXZ(2),2)
'                 fjytArea = sctnjzmj
'                 fangjArea = CDbl(sctnjzmj) - CDbl(scytjzmj)
'                 If  qiuh <> "" And  qiuh <> "*"  Then   If CDbl(scftxs) > 0 Then   dih = qiuh & "-" & dhIndex
'                 dhIndex = dhIndex + 1  Else dih = ""

'                 If i > 0 And i Mod copyRows = 0 Then     tableIndex = tableIndex + 1
'                 startRow = iniRow

'                 '每条属性保存在二维数组方便每页求和
'                 For j = 0 To UBound(HUSXZ)
'                     MyArray(i,j) = HUSXZ(j)
'                 Next
'                 '输出每页合计面积 
'                 If i Mod copyRows = 0 Then '满一页
'                     sum_scjzmj = 0
'                     sum_fangjArea = 0
'                     sum_scytjzmj = 0
'                     sum_fjytArea = 0
'                     sum_scftjzmj = 0
'                     If k < fwhCount  Then
'                         For k = copyRows * tableIndex To copyRows * (tableIndex + 1) - 1

'                             sum_scjzmj = sum_scjzmj + CDbl(GetFormatNumber(MyArray(k,1),2))
'                             sum_fangjArea = sum_fangjArea + (CDbl(GetFormatNumber(MyArray(k,2),2)) - CDbl(GetFormatNumber(MyArray(k,3),2)))
'                             sum_scytjzmj = sum_scytjzmj + CDbl(GetFormatNumber(MyArray(k,3),2))
'                             sum_fjytArea = sum_fjytArea + CDbl(GetFormatNumber(MyArray(k,2),2))
'                             sum_scftjzmj = sum_scftjzmj + CDbl(GetFormatNumber(MyArray(k,5),2))
'                         Next
'                     End If
'                     'msgbox  sum_scjzmj
'                     g_docObj.SetCellText tableIndex,iniRow + copyRows,1,GetFormatNumber(sum_scjzmj,2),True
'                     g_docObj.SetCellText tableIndex,iniRow + copyRows,2,GetFormatNumber(sum_fangjArea,2),True
'                     g_docObj.SetCellText tableIndex,iniRow + copyRows,3,GetFormatNumber(sum_scytjzmj,2),True
'                     g_docObj.SetCellText tableIndex,iniRow + copyRows,4,GetFormatNumber(sum_fjytArea,2),True
'                     g_docObj.SetCellText tableIndex,iniRow + copyRows,6,GetFormatNumber(sum_scftjzmj,2),True
'                 End If
'                 startCol = iniCol
'                 '子幢号    室号     地号    总建筑面积 房间面积    阳台面积     房间+阳台    公用分摊系数    公用分摊面积 所在层次    规划用途
'                 values = zizh & "||" & shbw & "||" & dih & "||" & scjzmj & "||" & fangjArea & "||" & scytjzmj & "||" & fjytArea & "||" & scftxs & "||" & scftjzmj & "||" & ch & "||" & sygn
'                 valuesList = Split(values,"||")
'                 For j = 0 To UBound(valuesList)
'                     If j = 3 Or j = 4  Or j = 5  Or j = 6  Or j = 8   Then valuesList(j) = GetFormatNumber(valuesList(j),2)
'                     g_docObj.SetCellText tableIndex,startRow,startCol,valuesList(j)
'                     startCol = startCol + 1
'                 Next
'                 startRow = startRow + 1
'                 'sumFwhArea=sumFwhArea+cdbl(scjzmj)

'             Next
'             '每页合计
'             'g_docObj.Replace "{幢总面积}",GetFormatNumber(sumFwhArea,2),0
'         End If
'     End Function


'=============房屋基本信息表
Function OutputTable1(ByVal section,ByVal zrzguid,ByVal outputMode)
    g_docObj.MoveToSection section
    tableIndex = 0
    iniRow = 3
    iniCol = 1
    startRow = iniRow
    startCol = iniCol
    SQL = "select  round(sum(jzmj),2),sygn from FC_LPB_户信息表  WHERE  FC_LPB_户信息表.id >1 and ch>0  and zrzguid='" & zrzguid & "'  group by sygn"
    GetSQLRecordAll ProjectName,SQL,HRecord,RecordHCount
    If  RecordHCount > 0 Then
        fwhCount = RecordHCount
        ReDim dsList(fwhCount)
        'redim hList(2,fwhCount)
        If fwhCount > 10 Then  g_docObj.CloneTableRow tableIndex, 4, 1, fwhCount - 10
        DSfwhCount = fwhCount
        For I = 0 To RecordHCount - 1
            hsxz = Split(HRecord(I),",")
            '=======================获取使用功能
            sygn = hsxz(1)
            sumSygnArea = hsxz(0)
            sumArea = sumArea + CDbl(sumSygnArea)'地上总面积
            sql2 = "select  jzmj from FC_LPB_户信息表  WHERE  FC_LPB_户信息表.id >1 and ch>0  and zrzguid='" & zrzguid & "' and sygn ='" & sygn & "'"
            GetSQLRecordAll ProjectName,SQL2,HRecord2,RecordHCount2
            If IsNumeric(RecordHCount2) = False Then RecordHCount2 = 0
            dsValues = RecordHCount2 & "||" & sygn & "||" & sumSygnArea
            dsList(i) = dsValues
        Next
        '冒泡排序，套数从大到小排
        For i = 0 To fwhCount - 1
            For j = 0 To fwhCount - 1 - i - 1
                dsValues = dsList(j)
                dsValues1 = dsList(j + 1)
                dsValuesList = Split(dsValues,"||")
                dsValuesList1 = Split(dsValues1,"||")
                tsIndex = dsValuesList(0)
                tsIndex1 = dsValuesList1(0)
                If Int(tsIndex) < Int(tsIndex1) Then
                    temp = dsList(j)
                    dsList(j) = dsList(j + 1)
                    dsList(j + 1) = temp
                End If
            Next
        Next
        '填充
        For i = 0 To  fwhCount - 1
            'msgbox  dsValues
            dsValues = dsList(i)
            dsValuesList = Split(dsValues,"||")
            fwhCount1 = dsValuesList(0)
            sygn = dsValuesList(1)
            sumSygnArea = dsValuesList(2)
            g_docObj.SetCellText tableIndex,startRow,startCol,sygn & "(" & fwhCount1 & ")套"
            If sumSygnArea > 0  Then
                g_docObj.SetCellText tableIndex,startRow,startCol + 1,FormatNumber(sumSygnArea,2, - 1, - 1,0)
            End If
            startRow = startRow + 1
        Next
        If fwhCount > 10 Then
            If sumArea > 0 Then  g_docObj.SetCellText tableIndex,startRow,iniCol + 1,sumArea
            iniRow = startRow + 1
            iniCol = 1
        Else
            If sumArea > 0 Then  g_docObj.SetCellText tableIndex,13,iniCol + 1,sumArea
            iniRow = 14
            iniCol = 1
        End If
        sumDsArea = sumArea
    End If
    '=================================地下部分
    startRow = iniRow
    startCol = iniCol
    sumArea = 0
    
    SQL = "select  round(sum(jzmj),2),sygn from FC_LPB_户信息表  WHERE  FC_LPB_户信息表.id >1 and ch<0  and zrzguid='" & zrzguid & "'  group by sygn"
    GetSQLRecordAll ProjectName,SQL,HRecord,RecordHCount
    
    If RecordHCount > 0 Then
        fwhCount = RecordHCount
        ReDim dxList(fwhCount)
        
        For i = 0 To fwhCount - 1
            If i > 0 And i Mod 10 = 0 Then startRow = iniRow
            startCol = startCol + 2
            hsxzdx = Split(HRecord(i),",")
            sygn = hsxzdx(1)
            sumSygnArea = hsxzdx(0)
            sumArea = sumArea + CDbl(sumSygnArea)'地下总面积
            sql2 = "select  jzmj from FC_LPB_户信息表  WHERE  FC_LPB_户信息表.id >1 and ch<0  and zrzguid='" & zrzguid & "' and sygn ='" & sygn & "'"
            GetSQLRecordAll ProjectName,SQL2,HRecord2,RecordHCount2
            If IsNumeric(RecordHCount2) = False Then RecordHCount2 = 0
            dxValues = RecordHCount2 & "||" & sygn & "||" & sumSygnArea
            'if i=3 then msgbox  dxValues
            dxList(i) = dxValues
        Next
        
        'msgbox dsList(3)
        '冒泡排序，套数从大到小排
        
        For i = 0 To fwhCount - 1
            For j = 0 To fwhCount - 1 - i - 1
                
                dxValues = dxList(j)
                dxValues1 = dxList(j + 1)
                dxValuesList = Split(dxValues,"||")
                dxValuesList1 = Split(dxValues1,"||")
                
                tsIndex = dxValuesList(0)
                tsIndex1 = dxValuesList1(0)
                If Int(tsIndex) < Int(tsIndex1) Then
                    temp = dxList(j)
                    dxList(j) = dxList(j + 1)
                    dxList(j + 1) = temp
                End If
            Next
        Next
        '填充
        If DSfwhCount > 10  Then
            If fwhCount > 10  Then
                g_docObj.CloneTableRow tableIndex, startRow + 1, 1, fwhCount - 10
            End If
        Else
            If fwhCount > 10  Then
                g_docObj.CloneTableRow tableIndex, 15, 1, fwhCount - 10
            End If
        End If
        For i = 0 To  fwhCount - 1
            dxValues = dxList(i)
            dxValuesList = Split(dxValues,"||")
            fwhCount1 = dxValuesList(0)
            sygn = dxValuesList(1)
            sumSygnArea = dxValuesList(2)
            g_docObj.SetCellText tableIndex,startRow,startCol,sygn & "(" & fwhCount1 & ")套"
            If sumSygnArea > 0  Then
                g_docObj.SetCellText tableIndex,startRow,startCol + 1,FormatNumber(sumSygnArea,2, - 1, - 1,0)
            End If
            startRow = startRow + 1
        Next
        nTableRowCount = g_docObj.GetTableRowCount(tableIndex)
        If sumArea > 0 Then g_docObj.SetCellText tableIndex,nTableRowCount - 1,iniCol + 1,sumArea
    End If
    
    
End Function

Function FuncStart(ByVal rptPathName,ByVal outputMode)
    If  outputMode = "汇总"  Then   f_docName = "房产测量报告_房开.docx" 'else f_docName="单体房产测量报告_民房不动产.docx"
    If  outputMode = "单体"  Then   f_docName = "单体房产测量报告_房开.docx"
    If  outputMode = "房产报告"  Then   f_docName = "单体房产测量报告_民房竣工.docx"
    If  outputMode = "不动产房产报告"  Then   f_docName = "单体房产测量报告_民房不动产.docx"
    If  outputMode = "单一产权"  Then   f_docName = "房产测量报告_厂房.docx"
    strDocFileName = SSProcess.GetSysPathName (7) & "输出模板\" & f_docName
    
    Set g_docObj = CreateObject ("asposewordscom.asposewordshelper")
    If  TypeName (g_docObj) = "AsposeWordsHelper" Then
        g_docObj.CreateDocumentByTemplate strDocFileName
    Else
        MsgBox "请先注册Aspose.Word插件"
        Exit Function
    End If
    Dim arSeletionRecord(), nSeletionCount
    If outputMode = "汇总" Then
        InitDB()
        UpdateFwhArea True
        '插入wmf
        'InsertAllImage Arypictures
        '输出 套内阳台、飘窗、设备平台指标统计表  
        OutputTable5 10
        
        '输出 套数统计表 
        OutputTable4 9
        
        '输出 分项面积统计表 
        OutputTable3 8
        
        '输出 房屋建筑面积实测报告单 
        OutputTable2 7
        
        '输出 房屋基本信息 
        OutputTable1 5,zrzguid,outputMode
        'word模板字段统一替换
        ReplaceTableField  zdTableName,"", "SpatialData", "2"
        UpdateFwhArea False
        ReleaseDB()
        '''f_docName=replace(f_docName,".doc",".pdf")
        g_docObj.SaveEx rptPathName & XiangMMC & "房产测量报告_汇总.docx"
    End If
    If outputMode = "单体" Then
        InitDB()
        UpdateFwhArea True
        '模板位置
        strTableName = "FC_自然幢信息属性表"
        strCondition = "FC_自然幢信息属性表.ID > 0 Order By FC_自然幢信息属性表.ZDDM,FC_自然幢信息属性表.ZRZH"
        nSeletionCount = SSRETools.SearchGeoTableAttr(strTableName, strCondition, "GeoAreaTB", "ZDDM & '->幢号：' & ZRZH & ' ID.' & " & strTableName & ".ID", arSeletionRecord)
        ' msgbox nSeletionCount
        ' 如果有多块宗地，将弹出对话框用户选择
        If nSeletionCount > 1 Then
            Result_Dlg = SSFunc.SelectListAttr("选择列表", "待选数据列表", "选中数据列表", arSeletionRecord, nSeletionCount)
            If Result_Dlg = 0 Then Exit Function
            If Result_Dlg = 1 And nSeletionCount = 0 Then MsgBox "未选，或未将选中内容加到 “选中数据列表” ，退出输出。"
            Exit Function
        End If
        For i = 0 To nSeletionCount - 1
            
            arTemp = Split(arSeletionRecord(i), " ")
            zrzid = Replace(arTemp(1), "ID.", "")
            arTemp2 = Split(arTemp(0), "->幢号：")
            zrzh = arTemp2(1)
            zrzguid = SSProcess.GetObjectAttr (zrzid, "[ZRZGUID]" )
            qiuh = SSProcess.GetObjectAttr (zrzid, "[QiuH]" )
            
            SetFwhGnmc zrzguid
            '插入成果图-插入前需进行打印处理
            'InsertImage 9,zrzh,"分层分户成果总图"
            '输出 公用配套建筑分摊系数计算表
            OutputTable7 2,zrzguid
            '房屋所有权面积测算汇总表
            OutputTable8 3,zrzguid,outputMode,qiuh
            '输出 房屋基本信息 
            OutputTable1 1,zrzguid,outputMode
            'word模板字段统一替换
            ReplaceTableField  zdTableName,"", "SpatialData", "2"
            ReplaceTableField  zrzTableName,zrzTableName & ".id=" & zrzid & "", "SpatialData", "2"
            g_docObj.SaveEx rptPathName & XiangMMC & zrzh & "房产测量报告.docx"
        Next
        UpdateFwhArea False
        ReleaseDB()
    End If
    If outputMode = "不动产房产报告" Then
        InitDB()
        UpdateFwhArea True
        '模板位置
        strTableName = "FC_自然幢信息属性表"
        strCondition = "FC_自然幢信息属性表.ID > 0 Order By FC_自然幢信息属性表.ZDDM,FC_自然幢信息属性表.ZRZH"
        nSeletionCount = SSRETools.SearchGeoTableAttr(strTableName, strCondition, "GeoAreaTB", "ZDDM & '->幢号：' & ZRZH & ' ID.' & " & strTableName & ".ID", arSeletionRecord)
        ' msgbox nSeletionCount
        ' 如果有多块宗地，将弹出对话框用户选择
        If nSeletionCount > 1 Then
            Result_Dlg = SSFunc.SelectListAttr("选择列表", "待选数据列表", "选中数据列表", arSeletionRecord, nSeletionCount)
            If Result_Dlg = 0 Then Exit Function
            If Result_Dlg = 1 And nSeletionCount = 0 Then MsgBox "未选，或未将选中内容加到 “选中数据列表” ，退出输出。"
            Exit Function
        End If
        For i = 0 To nSeletionCount - 1
            
            arTemp = Split(arSeletionRecord(i), " ")
            zrzid = Replace(arTemp(1), "ID.", "")
            arTemp2 = Split(arTemp(0), "->幢号：")
            zrzh = arTemp2(1)
            zrzguid = SSProcess.GetObjectAttr (zrzid, "[ZRZGUID]" )
            qiuh = SSProcess.GetObjectAttr (zrzid, "[QiuH]" )
            
            SetFwhGnmc zrzguid
            '插入成果图-插入前需进行打印处理
            'InsertImage 9,zrzh,"分层分户成果总图"
            '输出 公用配套建筑分摊系数计算表
            'OutputTable7 7,zrzguid
            '房屋所有权面积测算汇总表
            OutputTable8 2,zrzguid,outputMode,qiuh
            '输出 房屋基本信息 
            OutputTable1 1,zrzguid,outputMode
            'word模板字段统一替换
            ReplaceTableField  zdTableName,"", "SpatialData", "2"
            ReplaceTableField  zrzTableName,zrzTableName & ".id=" & zrzid & "", "SpatialData", "2"
            g_docObj.SaveEx rptPathName & XiangMMC & zrzh & "房产测量报告_不动产.docx"
        Next
        UpdateFwhArea False
        ReleaseDB()
    End If
    If outputMode = "房产报告" Then
        InitDB()
        UpdateFwhArea True
        '模板位置
        strTableName = "FC_自然幢信息属性表"
        strCondition = "FC_自然幢信息属性表.ID > 0 Order By FC_自然幢信息属性表.ZDDM,FC_自然幢信息属性表.ZRZH"
        nSeletionCount = SSRETools.SearchGeoTableAttr(strTableName, strCondition, "GeoAreaTB", "ZDDM & '->幢号：' & ZRZH & ' ID.' & " & strTableName & ".ID", arSeletionRecord)
        ' msgbox nSeletionCount
        ' 如果有多块宗地，将弹出对话框用户选择
        If nSeletionCount > 1 Then
            Result_Dlg = SSFunc.SelectListAttr("选择列表", "待选数据列表", "选中数据列表", arSeletionRecord, nSeletionCount)
            If Result_Dlg = 0 Then Exit Function
            If Result_Dlg = 1 And nSeletionCount = 0 Then MsgBox "未选，或未将选中内容加到 “选中数据列表” ，退出输出。"
            Exit Function
        End If
        For i = 0 To nSeletionCount - 1
            
            arTemp = Split(arSeletionRecord(i), " ")
            zrzid = Replace(arTemp(1), "ID.", "")
            arTemp2 = Split(arTemp(0), "->幢号：")
            zrzh = arTemp2(1)
            zrzguid = SSProcess.GetObjectAttr (zrzid, "[ZRZGUID]" )
            qiuh = SSProcess.GetObjectAttr (zrzid, "[QiuH]" )
            
            SetFwhGnmc zrzguid
            '插入成果图-插入前需进行打印处理
            'InsertImage 9,zrzh,"分层分户成果总图"
            '输出 公用配套建筑分摊系数计算表
            OutputTable7 5,zrzguid
            '房屋所有权面积测算汇总表
            OutputTable8 6,zrzguid,outputMode,qiuh
            '输出 房屋基本信息 
            'OutputTable1 1,zrzguid,outputMode
            'word模板字段统一替换
            ReplaceTableField  zdTableName,"", "SpatialData", "2"
            ReplaceTableField  zrzTableName,zrzTableName & ".id=" & zrzid & "", "SpatialData", "2"
            g_docObj.SaveEx rptPathName & XiangMMC & zrzh & "房产测量报告_房产.docx"
        Next
        UpdateFwhArea False
        ReleaseDB()
    End If
    If outputMode = "单一产权" Then
        
        InitDB()
        UpdateFwhArea True
        '模板位置
        strTableName = "FC_自然幢信息属性表"
        strCondition = "FC_自然幢信息属性表.ID > 0 Order By FC_自然幢信息属性表.ZDDM,FC_自然幢信息属性表.ZRZH"
        nSeletionCount = SSRETools.SearchGeoTableAttr(strTableName, strCondition, "GeoAreaTB", "ZDDM & '->幢号：' & ZRZH & ' ID.' & " & strTableName & ".ID", arSeletionRecord)
        ' msgbox nSeletionCount
        ' 如果有多块宗地，将弹出对话框用户选择
        If nSeletionCount > 1 Then
            Result_Dlg = SSFunc.SelectListAttr("选择列表", "待选数据列表", "选中数据列表", arSeletionRecord, nSeletionCount)
            If Result_Dlg = 0 Then Exit Function
            If Result_Dlg = 1 And nSeletionCount = 0 Then MsgBox "未选，或未将选中内容加到 “选中数据列表” ，退出输出。"
            Exit Function
        End If
        For i = 0 To nSeletionCount - 1
            
            arTemp = Split(arSeletionRecord(i), " ")
            zrzid = Replace(arTemp(1), "ID.", "")
            arTemp2 = Split(arTemp(0), "->幢号：")
            zrzh = arTemp2(1)
            zrzguid = SSProcess.GetObjectAttr (zrzid, "[ZRZGUID]" )
            qiuh = SSProcess.GetObjectAttr (zrzid, "[QiuH]" )
            
            SetFwhGnmc zrzguid
            '插入成果图-插入前需进行打印处理
            'InsertImage 9,zrzh,"分层分户成果总图"
            '输出 公用配套建筑分摊系数计算表
            'OutputTable7 5,zrzguid
            '房屋所有权面积测算汇总表
            OutputTable8 4,zrzguid,outputMode,qiuh
            '输出 房屋基本信息 
            OutputTable9 3,zrzguid,outputMode
            'word模板字段统一替换
            ReplaceTableField  zdTableName,"", "SpatialData", "2"
            ReplaceTableField  zrzTableName,zrzTableName & ".id=" & zrzid & "", "SpatialData", "2"
            g_docObj.SaveEx rptPathName & XiangMMC & zrzh & "房产测量报告.docx"
        Next
        UpdateFwhArea False
        ReleaseDB()
    End If
    MsgBox "输出完成，请在成果文件夹下查阅！"
End Function

Function GetSQLRecordAll(ByRef mdbName, ByRef StrSqlStatement, ByRef arSQLRecord(), ByRef iRecordCount)
    
    If StrSqlStatement = "" Then
        MsgBox "查询语句为空，操作停止！",48
    End If
    iRecordCount =  - 1
    sql = StrSqlStatement
    SSProcess.OpenAccessMdb mdbName
    SSProcess.OpenAccessRecordset mdbName, sql
    RecordCount = SSProcess.GetAccessRecordCount (mdbName, sql)
    If RecordCount > 0 Then
        iRecordCount = 0
        ReDim arSQLRecord(RecordCount)
        SSProcess.AccessMoveFirst mdbName, sql
        iRecordCount = 0
        While SSProcess.AccessIsEOF (mdbName, sql) = 0
            fields = ""
            values = ""
            SSProcess.GetAccessRecord mdbName, sql, fields, values
            arSQLRecord(iRecordCount) = values
            iRecordCount = iRecordCount + 1
            SSProcess.AccessMoveNext mdbName, sql
        WEnd
    End If
    SSProcess.CloseAccessRecordset mdbName, sql
    SSProcess.CloseAccessMdb mdbName
End Function

'判断文件夹是否存在
Function IsfolderExists(folder)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.folderExists(folder) Then
        IsfolderExists = True
    Else
        IsfolderExists = False
    End If
End Function


'创建文件夹

Function CreateFolders(path)
    Set fso = CreateObject("scripting.filesystemobject")
    CreateFolderEx fso,path
    Set fso = Nothing
End Function
Function CreateFolderEx(fso,path)
    If fso.FolderExists(path) Then
        Exit Function
    End If
    If Not fso.FolderExists(fso.GetParentFolderName(path)) Then
        CreateFolderEx fso,fso.GetParentFolderName(path)
    End If
    fso.CreateFolder(path)
End Function
'====================辅助函数
'//数字进位
Function GetFormatNumber(ByVal number,ByVal numberDigit)
    If IsNumeric(numberDigit) = False Then numberDigit = 2
    If IsNumeric(number) = False Then number = 0
    number = FormatNumber(Round(number + 0.00000001,numberDigit),numberDigit, - 1,0,0)
    GetFormatNumber = (number)
End Function