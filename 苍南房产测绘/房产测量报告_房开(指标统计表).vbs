
'入口
Sub Onclick()
    
    FCFX_ZBTJ
    
End Sub' Onclick

'=======================================业务函数==============================================

'套内阳台、飘窗、设备平台指标统计表
Function FCFX_ZBTJ()
    
    'Doc对象
    Dim Global_Word
    
    '模板路径
    Dim TamplateFilePath
    
    '输出路径
    Dim FilePath
    
    '表索引
    Dim TableIndex
    
    '复制数目
    Dim CloneCount
    
    '套内阳台、飘窗、设备平台指标统计表起始行
    Dim ZB_StartRow
    
    '套内阳台、飘窗、设备平台指标统计表结束行
    Dim ZB_EndRow
    
    '数据库
    Dim MdbName
    
    '参数初始化
    Set Global_Word = CreateObject ("asposewordscom.asposewordshelper")
    TamplateFilePath = SSProcess.GetSysPathName(7) & "输出模板\房产测量报告_房开.docx"
    FilePath = SSProcess.GetSysPathName(5) & "成果文件\房产测量报告\单体房产测量报告_房开.docx"
    TableIndex = 5
    CloneCount = 0
    ZB_StartRow = 2
    ZB_EndRow = 2
    MdbName = SSProcess.GetProjectFileName
    
    '根据模板创建Word文档
    Global_Word.CreateDocumentByTemplate TamplateFilePath
    
    '获取所有的幢
    SqlStr = "Select DISTINCT ZRZ_LP_信息表.ZRZH,ZRZGUID From ZRZ_LP_信息表 Where ZRZ_LP_信息表.ID > 0 "
    GetSQLRecordAll MdbName,SqlStr,InfoArr,ZrZhCount
    
    CloneCount = ZrZhCount - 1
    
    '复制表格
    For i = 1 To CloneCount
        Global_Word.CloneTable TableIndex,0,0,False
    Next 'i
    
    '自然幢GUID
    Dim GUID
    
    '编组GUID
    Dim BZGUID

    '自然幢号
    Dim ZRZH
    
    '层号
    Dim CH
    
    '室号
    Dim SHBW
    
    '套内面积
    Dim TNMJ
    
    '套内阳台面积
    Dim TNYTMJ
    
    '飘窗面积
    Dim PCMJ
    
    '不封闭阳台面积
    Dim BFBYTMJ
    
    '设备平台面积
    Dim SBPTMJ
    
    '备注
    Dim BZ
    
    '超出面积
    Dim OverArea
    
    '复制表格，所以所有的表格所有自增（初始为5）
    For i = 0 To ZrZhCount - 1
        ZRZInfoArr = Split(InfoArr(i),",", - 1,1)
        GUID = ZRZInfoArr(1)
        ZRZH = ZRZInfoArr(0)
        Global_Word.SetCellText TableIndex,ZB_StartRow,0,ZRZH,True,False
        
        SqlStr = "Select DISTINCT FC_LPB_户信息表.CH,SHBW From FC_LPB_户信息表 Where FC_LPB_户信息表.ID > 0 And FC_LPB_户信息表.ZRZGUID = " & GUID & " And FC_LPB_户信息表.SYGN Like '*住宅*' Order By FC_LPB_户信息表.CH"
        GetSQLRecordAll MdbName,SqlStr,CHSHArr,SHCount
        
        
        Global_Word.CloneTableRow TableIndex,ZB_StartRow,1,SHCount - 1,False
        
        ZB_EndRow = ZB_StartRow + SHCount - 1
        
        For j = ZB_StartRow To ZB_EndRow
            CS_Arr = Split(CHSHArr(j - ZB_StartRow),",", - 1,1)

            CH = CS_Arr(0)
            SHBW = CS_Arr(1)

            SqlStr = "Select FC_LPB_户信息表.BZGUID,YTJZMJ,PCTXMJ,BFMYTTXMJ,SBPTTXMJ From FC_LPB_户信息表 Where FC_LPB_户信息表.ZRZGUID = " & GUID & " And FC_LPB_户信息表.CH = " & CH & " And FC_LPB_户信息表.SHBW = '" & SHBW & "'"
            GetSQLRecordAll MdbName,SqlStr,T_InfoArr,SearchCount
            
            S_InfoArr = Split(T_InfoArr(0),",", - 1,1)
            
            BZGUID = S_InfoArr(0)
            TNYTMJ = S_InfoArr(1)
            PCMJ = S_InfoArr(2)
            BFBYTMJ = S_InfoArr(3)
            SBPTMJ = S_InfoArr(4)
            
            SqlStr = "Select FC_面积块信息属性表.JZMJ From FC_面积块信息属性表 Where FC_面积块信息属性表.BZGUID = " & BZGUID
            GetSQLRecordAll MdbName,SqlStr,TNMJ_Arr,TNMJ_Count
            
            TNMJ = TNMJ_Arr(0)
            
            SqlStr = "Select FC_LPB_户信息表.BZ From FC_LPB_户信息表 Where FC_LPB_户信息表.ZRZGUID = " & GUID & " And FC_LPB_户信息表.CH = " & CH & " And FC_LPB_户信息表.SHBW = '" & SHBW & "'"
            GetSQLRecordAll MdbName,SqlStr,BZ_Arr,BZCount

            BZ = BZ_Arr(0)
            
            GetOverArea BZ,TNYTMJ,OverArea
            
            Global_Word.SetCellText TableIndex,j,1,CH,True,False
            Global_Word.SetCellText TableIndex,j,2,SHBW,True,False
            
            If TNMJ <> "" Then
                If TNMJ <> 0 Then
                    Global_Word.SetCellText TableIndex,j,3,GetFormatNumber(TNMJ,2),True,False
                End If
            End If
            
            If TNYTMJ <> "" Then
                If TNYTMJ <> 0 Then
                    Global_Word.SetCellText TableIndex,j,4,GetFormatNumber(TNYTMJ,2),True,False
                End If
            End If
            
            If PCMJ <> "" Then
                If PCMJ <> 0 Then
                    Global_Word.SetCellText TableIndex,j,6,GetFormatNumber(PCMJ,2),True,False
                End If
            End If
            
            If BFBYTMJ <> "" Then
                If BFBYTMJ <> 0 Then
                    Global_Word.SetCellText TableIndex,j,7,GetFormatNumber(BFBYTMJ,2),True,False
                End If
            End If
            
            If SBPTMJ <> ""  Then
                If SBPTMJ <> 0 Then
                    Global_Word.SetCellText TableIndex,j,8,GetFormatNumber(SBPTMJ,2),True,False
                End If   
            End If
            
            If OverArea <> "" Then
                If OverArea <> 0 Then
                    Global_Word.SetCellText TableIndex,j,9,GetFormatNumber(OverArea,2),True,False
                End If
            End If
            
            If TNYTMJ <> "" And  TNMJ <> "" Then
                If TNYTMJ <> 0 And  TNMJ <> 0 Then
                    Global_Word.SetCellText TableIndex,j,5,GetFormatNumber(TNYTMJ / TNMJ,4),True,False
                End If
            End If
        Next 'j
        
        Global_Word.MergeCell TableIndex,ZB_StartRow,0,ZB_EndRow,0,False
        
        TableIndex = TableIndex + 1
        
    Next 'i
    
    '保存文档
    Global_Word.SaveEx FilePath
    
End Function' FCFX_ZBTJ

Function GetOverArea(ByVal BZ,ByVal TNYTMJ,ByRef OverArea)
    
    OverArea = ""
    If InStrRev(BZ,"半算为：", - 1,1) <> 0 Then
        
        HalfArea_StartPos = InStrRev(BZ,"半算为：", - 1,1) + 4
        HalfArea_EndPos = Len(BZ) - 1
        
        TotalArea_StartPos = InStrRev(BZ,"超出部分全算为：", - 1,1) + 8
        
        TotalArea_EndPos = HalfArea_StartPos - 5
        
        HalfArea_NumLen = HalfArea_EndPos - HalfArea_StartPos + 1
        TotalArea_NumLen = TotalArea_EndPos - TotalArea_StartPos
        
        HalfArea = Transform((Mid(BZ,HalfArea_StartPos,HalfArea_NumLen)))
        
        TotalArea = Transform((Mid(BZ,TotalArea_StartPos,TotalArea_NumLen)))
        
        OverArea = HalfArea + TotalArea - Transform(TNYTMJ)
        
    End If
End Function' GetOverArea

'=======================================工具类函数============================================

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

'数字进位
Function GetFormatNumber(ByVal number,ByVal numberDigit)
    If IsNumeric(numberDigit) = False Then numberDigit = 2
    If IsNumeric(number) = False Then number = 0
    number = FormatNumber(Round(number + 0.00000001,numberDigit),numberDigit, - 1,0,0)
    GetFormatNumber = (number)
End Function
