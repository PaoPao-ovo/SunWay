' 一、获取报告模板路径并复制到输出路径
' 1、获取当前的模板路径并创建Word
' 2、设置Word保存路径和文件名

' 二、根据项目红线属性值进行字符串替换
' 1、获取项目红线的属性（获取项目信息时需要去除括号）

'========================================================Doc操作对象和文件路径操作对象================================================================

'Doc全局对象
Dim Global_Word
Set Global_Word = CreateObject ("asposewordscom.asposewordshelper")

'路径操作对象
Dim FileSysObj
Set FileSysObj = CreateObject("Scripting.FileSystemObject")

'===================================================================功能入口=======================================================================

Sub OnClick()
    If  TypeName (Global_Word) = "AsposeWordsHelper" Then
        Global_Word.CreateDocumentByTemplate  SSProcess.GetSysPathName (7) & "输出模板\" & "人防测量报告模板.doc"
    Else
        MsgBox "请先注册Aspose.Word插件"
        Exit Sub
    End If
    
    AllVisible
    ReplaceValue Yjmj,SPYJmj,ZhuZhaiXs,FZhuZhaiXs,HxId
    If ToDecimal(ZhuZhaiXs) = 0 Or ToDecimal(FZhuZhaiXs) = 0 Then
        MsgBox "住宅系数或非住宅系数为零"
        Exit Sub
    End If
    RfResultTableInner RfCount,RfValArr,Yjmj,SPYJmj
    RfYJTableInner Yjmj,ZhuZhaiXs,FZhuZhaiXs,HxId
    
    Global_Word.SaveEx  SSProcess.GetSysPathName(5) & "成果文件" & "\人防报告.doc"
    
    MsgBox "输出完成"
End Sub' OnClick

'======================================================字符串替换==============================================================================

'字符串替换
Function ReplaceValue(ByRef Yjmj,ByRef SPYJmj,ByRef ZhuZhaiXs,ByRef FZhuZhaiXs,ByRef HxId)
    
    RePlaceStr = "XiangMMC,XiangMDZ,SheJDW,JianSDW,WeiTDW,CeLDW,CeLRQ,JianZJG,JunGCLDSJZMJ,DSCS,JunGCLZTS,JunGCLDXJZMJ,ZZRFYJMJ,QTRFYJMJ,HLHTMJ,FKYBMJ,ZBYTHD,BPGC,SPRFYJMJ"
    
    SelFeatures "9130223",HxCount,HxId
    GetInnerFeatures "9130223","9210123",ZrzCount,ZrzArr
    
    ReDim DateArr(ZrzCount)
    For i = 0 To ZrzCount - 1
        DateArr(i) = Transform(Replace(Replace(Replace(SSProcess.GetObjectAttr(ZrzArr(i),"[JGRQ]"),"年",""),"月",""),"日",""))
    Next 'i
    UpDateTime = DateArr(0)
    pos = 0
    For i = 1 To ZrzCount - 1
        If UpDateTime < DateArr(i) Then
            UpDateTime = DateArr(i)
            pos = i
        End If
    Next 'i
    UpDateTime = SSProcess.GetObjectAttr(ZrzArr(pos),"[JGRQ]")
    
    If HxCount = 1 Then
        RePlaceArr = Split(RePlaceStr,",", - 1,1)
        For i = 0 To UBound(RePlaceArr)
            szaa = SSProcess.GetSelGeoValue(0, "[" & RePlaceArr(i) & "]")
            If szaa = "0" Or szaa = "0.0"  Then
                Global_Word.Replace "{" & RePlaceArr(i) & "}","",0
            Else
                Global_Word.Replace "{" & RePlaceArr(i) & "}",SSProcess.GetSelGeoValue(0, "[" & RePlaceArr(i) & "]"),0
            End If
        Next 'i
    End If
    
    SqlStr = "Select DISTINCT PSGN From 人防防护单元属性表 inner join GeoAreaTB on 人防防护单元属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0"
    GetSQLRecordAll SqlStr,DISTINCTArr,LxCount
    Global_Word.Replace "{PSGN}",Replace(DISTINCTArr(0),",","、"),0
    
    Yjmj = Transform(SSProcess.GetSelGeoValue(0,"[ZZRFYJMJ]")) + Transform(SSProcess.GetSelGeoValue(0,"[QTRFYJMJ]"))
    SPYJmj = Transform(SSProcess.GetSelGeoValue(0,"[SPRFYJMJ]"))
    ZhuZhaiXs = SSProcess.GetSelGeoValue(0,"[ZhuZXS]")
    FZhuZhaiXs = SSProcess.GetSelGeoValue(0,"[FZhuZXS]")
    RFBH = Replace(SSProcess.GetSelGeoValue(0,"[HeTBH]"),"FCCL","")
    StartPos = InStr(1,RFBH,"(",1)
    EndPos = InStr(1,RFBH,")",1)
    If StartPos > 0 Then
        Leftstr = Left(RFBH,StartPos - 1)
    Else
        Leftstr = ""
    End If
    If EndPos > 0 Then
        Rightstr = Right(RFBH,Len(RFBH) - EndPos)
    Else
        Rightstr = ""
    End If
    If Leftstr & Rightstr = "" Then
        Leftstr = RFBH
        Rightstr = ""
    End If
    Global_Word.Replace "{HeTBH}",Leftstr & Rightstr,0
    Global_Word.Replace "｛YJMJ｝",Yjmj,0
    Global_Word.Replace "{TTT}",UpDateTime,0
End Function' ReplaceValue

'==========================================================人防测量成果表====================================================================

'人防测量成果表入口函数
Function RfResultTableInner(ByRef RfCount,ByRef RfValArr(),ByVal Yjmj,ByVal SPYJmj)
    RfClMjInert RfCount,RfValArr,Yjmj,SPYJmj
End Function' RfResultTableInner

'人防测量面积表格填值
Function RfClMjInert(ByRef RfCount,ByRef RfValArr(),ByVal Yjmj,ByVal SPYJmj)
    
    SqlStr = "Select 人防防护单元属性表.ID From 人防防护单元属性表 Inner Join GeoAreaTB on 人防防护单元属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0"
    GetSQLRecordAll SqlStr,RfArr,RfCount
    FildArr = Split("FHDYBH,JZMJ,YBMJ,GYMJ,ZSGN,FHUDJ,FHUADJ,KBDYS,KBSL,SZCS,PSGN,TCWS,FJDCS,BZ",",", - 1,1)
    ReDim RfValArr(RfCount)
    Start = 0
    
    For i = 0 To RfCount - 1
        For j = 0 To UBound(FildArr)
            If RfValArr(Start) = "" Then
                RfValArr(Start) = SSProcess.GetObjectAttr(RfArr(i),"[" & FildArr(j) & "]")
            Else
                RfValArr(Start) = RfValArr(Start) & "," & SSProcess.GetObjectAttr(RfArr(i),"[" & FildArr(j) & "]")
            End If
        Next 'j
        Start = Start + 1
    Next 'i
    
    InsertTable RfValArr,RfCount,3,5,1,8,14,HjRow
    
    GetMaxUnderFloor 3,5,HjRow - 1,10,MaxUnderFloor
    
    Global_Word.SetCellText 3,1,8,MaxUnderFloor,True,False

    Global_Word.Replace "{" & "DXCS" & "}",MaxUnderFloor,0

    InsertSum 3,5,HjRow - 1,HjRow,2
    InsertSum 3,5,HjRow - 1,HjRow,3
    InsertSum 3,5,HjRow - 1,HjRow,8
    InsertSum 3,5,HjRow - 1,HjRow,9
    InsertSum 3,5,HjRow - 1,HjRow,12
    InsertSum 3,5,HjRow - 1,HjRow,13
    InsertSm 5,HjRow - 1,RfCount,Yjmj,SPYJmj
    
End Function' RfClMjInert

'获取最大的地下层数
Function GetMaxUnderFloor(ByVal TableIndex,ByVal StartRow,ByVal EndRow,ByVal CalCol,ByRef MaxUnderFloor)
    MaxUnderFloor = 0
    For i = StartRow To EndRow
        CurrentFloor = GetCurrentFloor(GetSelCellVal(Global_Word.GetCellText(TableIndex,i,CalCol,False)))
        If MaxUnderFloor = 0 Then
            MaxUnderFloor = CurrentFloor
        Else
            If MaxUnderFloor < CurrentFloor Then
                MaxUnderFloor = CurrentFloor
            Else
                MaxUnderFloor = MaxUnderFloor
            End If
        End If
    Next 'i
End Function' GetMaxUnderFloor

'获取当前地下层数
Function GetCurrentFloor(ByVal Content)
    If Content <> "" Then
        BigNum = Mid(Content,3,Len(Content) - 3)
        '最多地下十层
        Number = "1,2,3,4,5,6,7,8,9,10"
        BigNumber = "一,二,三,四,五,六,七,八,九,十"
        NumberArr = Split(Number,",", - 1,1)
        BigNumberArr = Split(BigNumber,",", - 1,1)
        If Len(BigNum) = 1 Then
            For i = 0 To UBound(BigNumberArr)
                If BigNumberArr(i) = BigNum Then
                    SmallNum = NumberArr(i)
                End If
            Next 'i
        ElseIf Len(BigNum) = 2 Then
            For i = 0 To UBound(BigNumberArr)
                If BigNumberArr(i) = Mid(BigNum,2,1) Then
                    SmallNum = NumberArr(i)
                    SmallNum = "1" & SmallNum
                End If
            Next 'i
        ElseIf Len(BigNum) = 3 Then
            For i = 0 To UBound(BigNumberArr)
                If BigNumberArr(i) = Mid(BigNum,1,1) Then
                    SmallNum1 = NumberArr(i)
                End If
                If BigNumberArr(i) = Mid(BigNum,3,1) Then
                    SmallNum2 = NumberArr(i)
                End If
                SmallNum = SmallNum1 & SmallNum2
            Next 'i
        End If
        GetCurrentFloor = CInt(SmallNum)
    Else
        GetCurrentFloor = 0
    End If
    
End Function' GetCurrentFloor

'填写合计值
Function InsertSum(ByVal TableIndex,ByVal StartRow,ByVal EndRow,ByVal HjRow,ByVal CalCol)
    TotalArea = 0
    For i = StartRow To EndRow
        If Global_Word.GetCellText(TableIndex,StartRow,CalCol,False) <> "" Then
            TotalArea = TotalArea + Transform(GetSelCellVal(Global_Word.GetCellText(TableIndex,i,CalCol,False)))
        End If
    Next 'i
    Global_Word.SetCellText Tableindex,HjRow,CalCol,TotalArea,True,False
End Function' InsertSum

'填写人防区域说明区域
Function InsertSm(ByVal StartRow,ByVal EndRow,ByVal RfCount,ByVal Yjmj,ByVal SPYJmj)
    RePlaceArr = Split("STR1,STR2,STR3",",", - 1,1)
    GetYbArea StartRow,EndRow,AreaArr,YBString
    ValArr = Split(RfCount & ";" & Transform(GetSelCellVal(Global_Word.GetCellText(3,EndRow + 1,2,False))) & ";" & YBString,";", - 1,1)
    For i = 0 To UBound(RePlaceArr)
        Global_Word.Replace "｛" & RePlaceArr(i) & "｝",ValArr(i),0
    Next 'i
    If (Transform(GetSelCellVal(Global_Word.GetCellText(3,EndRow + 1,2,False))) - SPYJmj) >= 0 Then
        Replace4 = "实际超面积" & Round(Transform(GetSelCellVal(Global_Word.GetCellText(3,EndRow + 1,2,False))) - SPYJmj,2)
    Else
        Replace4 = "实际小于面积" & Abs(Round(Transform(GetSelCellVal(Global_Word.GetCellText(3,EndRow + 1,2,False))) - SPYJmj,2))
    End If
    Global_Word.Replace "｛" & "STR4" & "｝",Replace4,0
    Global_Word.Replace "｛" & "STR5" & "｝","实测应建人防面积" & Yjmj,0
    If (Transform(GetSelCellVal(Global_Word.GetCellText(3,EndRow + 1,2,False))) - Yjmj) >= 0 Then
        Replace6 = "实际超面积" & Transform(GetSelCellVal(Global_Word.GetCellText(3,EndRow + 1,2,False))) - Yjmj
    Else
        Replace6 = "实际小于面积" & Abs(Transform(GetSelCellVal(Global_Word.GetCellText(3,EndRow + 1,2,False))) - Yjmj)
    End If
    Global_Word.Replace "｛" & "STR6" & "｝",Replace6,0
    '4:Transform(GetSelCellVal(Global_Word.GetCellText(2,EndRow + 1,2,False))) - SPYJmj
    '5:Yjmj
    '6:Transform(GetSelCellVal(Global_Word.GetCellText(2,EndRow + 1,2,False))) - Yjmj
End Function' InsertSm

'获取当前掩蔽单元的类型
Function GetYbLx(ByVal StartRow,ByVal EndRow,ByRef Lxstring,ByRef lxCount)
    GnStr = ""
    For i = StartRow To EndRow
        If GnStr = "" And GetSelCellVal(Global_Word.GetCellText(3,i,5,False)) <> "" Then
            GnStr = GetSelCellVal(Global_Word.GetCellText(3,i,5,False))
        ElseIf GetSelCellVal(Global_Word.GetCellText(3,i,5,False)) <> "" Then
            GnStr = GnStr & "," & GetSelCellVal(Global_Word.GetCellText(3,i,5,False))
        End If
    Next 'i
    DelRepeat Split(GnStr,",", - 1,1),Lxstring,lxCount
End Function' GetYbLx

'获取对应的掩蔽区面积字符串
Function GetYbArea(ByVal StartRow,ByVal EndRow,ByRef AreaArr(),ByRef TotalString)
    TotalString = ""
    GetYbLx StartRow,EndRow,Lxstring,lxcount
    LxArr = Split(Lxstring,",", - 1,1)
    ReDim AreaArr(lxcount,2)
    For i = 0 To UBound(LxArr)
        SingleLxCount = 0
        For j = StartRow To EndRow
            If LxArr(i) = GetSelCellVal(Global_Word.GetCellText(3,j,5,False)) Then
                SingleLxCount = SingleLxCount + 1
                AreaArr(i,1) = AreaArr(i,1) + Transform(GetSelCellVal(Global_Word.GetCellText(3,j,3,False)))
                AreaArr(i,0) = LxArr(i)
            End If
        Next 'j
        If TotalString = "" Then
            TotalString = "其中" & SingleLxCount & "个" & AreaArr(i,0) & "单元计" & AreaArr(i,1) & "平方米"
        Else
            TotalString = TotalString & "," & SingleLxCount & "个" & AreaArr(i,0) & "单元计" & AreaArr(i,1) & "平方米"
        End If
    Next 'i
End Function' GetYbArea

'========================================================人防应建面积计算表============================================================================

'人防应建面积计算表入口函数
Function RfYJTableInner(ByVal Yjmj,ByVal ZhuZhaiXs,ByVal FZhuZhaiXs,ByVal HxId)
    InsertZhuZ 4,HxId,2,EndRow,HjRow,ZhuZhaiXs
    InsertFZhuZ HxId,2,EndRow,4,HjRow,FZhuZhaiXs
End Function' RfResultTableInner

'填写住宅建筑部分
Function InsertZhuZ(ByVal TableIndex,ByVal HxId,ByVal StartRow,ByRef EndRow,ByRef HjRow,ByVal ZhuZhaiXs)
    SearchZrz HxId,ZrzArr,ZrzCount
    EndRow = StartRow + ZrzCount - 1
    If EndRow <= 13 Then
        HjRow = 14
        For i = 0 To ZrzCount - 1
            SingleZrzValArr = Split(ZrzArr(i),",", - 1,1)
            For j = 0 To 3
                Select Case j
                    Case 0
                    Global_Word.SetCellText TableIndex,i + StartRow,j,SingleZrzValArr(0),True,False
                    Case 1
                    Global_Word.SetCellText TableIndex,i + StartRow,j,Round((SingleZrzValArr(1) / ToDecimal(ZhuZhaiXs)),2),True,False
                    Case 2
                    Global_Word.SetCellText TableIndex,i + StartRow,j,ZhuZhaiXs,True,False
                    Case 3
                    Global_Word.SetCellText TableIndex,i + StartRow,j,Round(SingleZrzValArr(1),2),True,False
                End Select
            Next 'j
        Next 'i
    Else
        Global_Word.CloneTableRow TableIndex,StartRow,1,ZrzCount - 12, False
        HjRow = 2 + ZrzCount
        For i = 0 To ZrzCount - 1
            SingleZrzValArr = Split(ZrzArr(i),",", - 1,1)
            For j = StartCol To EndCol
                Select Case j
                    Case 0
                    Global_Word.SetCellText TableIndex,i + StartRow,j,SingleZrzValArr(0),True,False
                    Case 1
                    Global_Word.SetCellText TableIndex,i + StartRow,j,Round((SingleZrzValArr(1) / ToDecimal(ZhuZhaiXs)),2),True,False
                    Case 2
                    Global_Word.SetCellText TableIndex,i + StartRow,j,ZhuZhaiXs,True,False
                    Case 3
                    Global_Word.SetCellText TableIndex,i + StartRow,j,Round(SingleZrzValArr(1),2),,True,False
                End Select
            Next 'j
        Next 'i
    End If
    For i = StartRow To EndRow
        TotalArea = TotalArea + Transform(GetSelCellVal(Global_Word.GetCellText(TableIndex,i,3,False)))
    Next 'i
    Global_Word.SetCellText TableIndex,HjRow,1,TotalArea,True,False
End Function' InsertZhuZ

'填写非住宅建筑部分
Function InsertFZhuZ(ByVal HxId,ByVal StartRow,ByRef EndRow,ByVal TableIndex,ByRef HjRow,ByVal FZhuZhaiXs)
    SearchMjk ZrzCount,HxId,2,EndRow,TableIndex,FZhuZhaiXs,HjRow
    SearchH ZrzCount,HxId,EndRow,iEndRow,TableIndex,FZhuZhaiXs,HjRow
    InsertFZhuZSum TableIndex,2,iEndRow,HjRow
End Function' InsertFZhuZ

'搜索ZDGUID相同的自然幢数组
Function SearchZrz(ByVal HxId,ByRef ZrzArr,ByRef ZrzCount)
    SqlString = "Select ZRZH,ZhuZMJ,ZRZGUID From FC_自然幢信息属性表 inner join GeoAreaTB on FC_自然幢信息属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0 AND FC_自然幢信息属性表.ZDGUID = " & SSProcess.GetObjectAttr(HxId,"[ZDGUID]")
    GetSQLRecordAll SqlString,ZrzArr,ZrzCount
End Function' SearchZrz

'搜索ZRZGUID相同的面积块并填值
Function SearchMjk(ByRef ZrzCount,ByVal HxId,ByRef StartRow,ByRef EndRow,ByVal TableIndex,ByVal FZhuZhaiXs,ByRef HjRow)
    SearchZrz HxId,ZrzArr,ZrzCount
    StartRow = StartRow
    EndRow = StartRow
    If ZrzCount > 0 Then
        For i = 0 To ZrzCount - 1
            SingleArr = Split(ZrzArr(i),",", - 1,1)
            SqlString = "Select DISTINCT MJKMC From FC_面积块信息属性表 inner join GeoAreaTB on FC_面积块信息属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0 AND FC_面积块信息属性表.ZRZGUID = " & SingleArr(2) & " AND FC_面积块信息属性表.FTLX = '不分摊'"
            GetSQLRecordAll SqlString,MjkMcArr,MjkLxCount
            If MjkLxCount > 0 Then
                For j = 0 To MjkLxCount - 1
                    SqlString = "Select Sum(FC_面积块信息属性表.KZMJ) From FC_面积块信息属性表 inner join GeoAreaTB on FC_面积块信息属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0 AND FC_面积块信息属性表.MJKMC = " & "'" & MjkMcArr(j) & "'" & " AND FC_面积块信息属性表.FTLX = '不分摊'"
                    GetSQLRecordAll SqlString,FZhuZMjArr,FZhuZCount
                    If StartRow + MjkLxCount - 1 <= HjRow - 1 Then
                        Global_Word.SetCellText TableIndex,StartRow + j ,4,SingleArr(0) & MjkMcArr(j),True,False
                        Global_Word.SetCellText TableIndex,StartRow + j ,5,Round(FZhuZMjArr(0),2),True,False
                        Global_Word.SetCellText TableIndex,StartRow + j ,6,FZhuZhaiXs,True,False
                        Global_Word.SetCellText TableIndex,StartRow + j ,7,Round(ToDecimal(FZhuZhaiXs) * FZhuZMjArr(0),2),True,False
                    ElseIf StartRow + MjkLxCount - 1 > HjRow - 1 Then
                        Global_Word.CloneTableRow TableIndex,StartRow,1,HjRow - StartRow - MjkLxCount  , False
                        HjRow = HjRow + HjRow - StartRow - MjkLxCount
                        Global_Word.SetCellText TableIndex,StartRow + j ,4,SingleArr(0) & MjkMcArr(j),True,False
                        Global_Word.SetCellText TableIndex,StartRow + j ,5,Round(FZhuZMjArr(0),2),True,False
                        Global_Word.SetCellText TableIndex,StartRow + j ,6,FZhuZhaiXs,True,False
                        Global_Word.SetCellText TableIndex,StartRow + j ,7,Round(ToDecimal(FZhuZhaiXs) * FZhuZMjArr(0),2),True,False
                    End If
                Next 'j
                StartRow = MjkLxCount + StartRow
                EndRow = StartRow - 1
            Else
                StartRow = StartRow
                EndRow = StartRow
            End If
        Next 'i
    End If
End Function' SearchMjk

'搜索ZRZGUID相同的户并填值
Function SearchH(ByRef ZrzCount,ByVal HxId,ByRef StartRow,ByRef EndRow,ByVal TableIndex,ByVal FZhuZhaiXs,ByRef HjRow)
    SearchZrz HxId,ZrzArr,ZrzCount
    For i = 0 To ZrzCount - 1
        SingleArr = Split(ZrzArr(i),",", - 1,1)
        SqlString = "Select DISTINCT SYGN From FC_户信息属性表 inner join GeoAreaTB on FC_户信息属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0 AND NOT INSTR(FC_户信息属性表.SYGN,'住宅') AND FC_户信息属性表.ZRZGUID = " & "'" & SingleArr(2) & "'"
        GetSQLRecordAll SqlString,SYGNArr,SYGNCount
        For j = 0 To SYGNCount - 1
            SqlString = "Select Sum(FC_户信息属性表.JZMJ) From FC_户信息属性表 inner join GeoAreaTB on FC_户信息属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0 AND FC_户信息属性表.SYGN = " & "'" & SYGNArr(j) & "'" & " AND FC_户信息属性表.ZRZGUID = " & "'" & SingleArr(2) & "'"
            GetSQLRecordAll SqlString,MJZArr,MJZCount
            If StartRow + SYGNCount - 1 <= HjRow - 1 Then
                Global_Word.SetCellText TableIndex,StartRow + j ,4,SingleArr(0) & SYGNArr(j),True,False
                Global_Word.SetCellText TableIndex,StartRow + j ,5,Round(MJZArr(0),2),True,False
                Global_Word.SetCellText TableIndex,StartRow + j ,6,FZhuZhaiXs,True,False
                Global_Word.SetCellText TableIndex,StartRow + j ,7,Round(ToDecimal(FZhuZhaiXs) * MJZArr(0),2),True,False
            ElseIf StartRow + SYGNCount - 1 > HjRow - 1 Then
                Global_Word.CloneTableRow TableIndex,StartRow,1,HjRow - StartRow - SYGNCount  , False
                HjRow = HjRow + HjRow - StartRow - SYGNCount
                Global_Word.SetCellText TableIndex,StartRow + j ,4,SingleArr(0) & SYGNArr(j),True,False
                Global_Word.SetCellText TableIndex,StartRow + j ,5,Round(MJZArr(0),2),True,False
                Global_Word.SetCellText TableIndex,StartRow + j ,6,FZhuZhaiXs,True,False
                Global_Word.SetCellText TableIndex,StartRow + j ,7,Round(ToDecimal(FZhuZhaiXs) * MJZArr(0),2),True,False
            End If
        Next 'j
        StartRow = SYGNCount + StartRow
        EndRow = StartRow - 1
    Next 'i
End Function' SearchH

'计算非住宅合计值
Function InsertFZhuZSum(ByVal TableIndex,ByVal StartRow,ByVal EndRow,ByVal HjRow)
    For i = StartRow To EndRow
        TotalArea = TotalArea + Transform(GetSelCellVal(Global_Word.GetCellText(TableIndex,i,7,False)))
    Next 'i
    Global_Word.SetCellText TableIndex,HjRow,3,TotalArea,True,False
End Function' InsertFZhuZSum
'==============================================================工具类函数===============================================================

'打开所有图层
Function AllVisible()
    count = SSProcess.GetLayerCount
    For i = 0 To count - 1
        layername = SSProcess.GetLayerName (i)
        SSProcess.SetLayerStatus layername, 1, 1
    Next
    SSProcess.RefreshView
End Function

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

'去除字符串中重复值
Function DelRepeat(ByVal StrArr(),ByRef ToTalVal,ByRef LxCount)
    ToTalVal = ""
    For i = 0 To UBound(StrArr)
        If ToTalVal = "" Then
            ToTalVal = "'" & StrArr(i) & "'"
        ElseIf Replace(ToTalVal,StrArr(i),"") = ToTalVal Then
            ToTalVal = ToTalVal & "," & "'" & StrArr(i) & "'"
        End If
    Next 'i
    ToTalVal = Replace(ToTalVal,"'","")
    LxCount = UBound(Split(ToTalVal,",", - 1,1)) + 1
End Function' DelRepeat

'选择指定地物并返回个数
Function SelFeatures(ByVal Code,ByRef Count,ByRef ID)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_code", "==", Code
    SSProcess.SelectFilter
    Count = SSProcess.GetSelGeoCount
    ID = SSProcess.GetSelGeoValue(0,"SSObj_ID")
End Function' SelFeatures

'选择指定地物内部的指定地物并返回地物id数组
Function GetInnerFeatures(ByVal OuterCode,ByVal InnerCode ,ByRef InnerCount,ByRef InnerArr())
    SelFeatures OuterCode,OuterCount,HxID
    InnerArr = Split(SSProcess.SearchInnerObjIDs(HxID,2,InnerCode,0),",", - 1,1)
    InnerCount = UBound(InnerArr) + 1
End Function' GetInnerFeatures

'获取单元格值
Function GetSelCellVal(ByVal CellContent)
    GetSelCellVal = Left(CellContent,Len(CellContent) - 1)
End Function' GetSelCellVal

'百分号转小数
Function ToDecimal(ByVal Percentage)
    ToDecimal = Transform(Left(Percentage,Len(Percentage) - 1)) * 0.01
End Function' ToDecimal

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

'填写表格
Function InsertTable(ByVal ValueArr(),ByVal ArrSize,ByVal TableIndex,ByVal StartRow,ByVal StartCol,ByVal MaxRow,ByVal EndCol,ByRef HjRow)
    'MsgBox ArrSize
    If ArrSize <= MaxRow - StartRow + 1 Then
        HjRow = 9
        For i = 0 To ArrSize - 1
            ValArr = Split(ValueArr(i),",", - 1,1)
            For j = StartCol To EndCol
                If j = 3 And Transform(ValArr(3 - StartCol)) = 0 Then
                    Global_Word.SetCellText Tableindex,i + StartRow,j,"",True,False
                ElseIf j = 4 And Transform(ValArr(4 - StartCol)) = 0 Then
                    Global_Word.SetCellText Tableindex,i + StartRow,j,"",True,False
                Else
                    Global_Word.SetCellText Tableindex,i + StartRow,j,ValArr(j - StartCol),True,False
                End If
            Next 'j
        Next 'i
    Else
        Global_Word.CloneTableRow Tableindex,StartRow,1,ArrSize - MaxRow + StartRow - 1, False
        HjRow = ArrSize + StartRow
        For i = 0 To ArrSize - 1
            ValArr = Split(ValueArr(i),",", - 1,1)
            For j = StartCol To EndCol
                If j = 3 And Transform(ValArr(3 - StartCol)) = 0 Then
                    Global_Word.SetCellText Tableindex,i + StartRow,j,"",True,False
                ElseIf j = 4 And Transform(ValArr(4 - StartCol)) = 0 Then
                    Global_Word.SetCellText Tableindex,i + StartRow,j,"",True,False
                Else
                    Global_Word.SetCellText Tableindex,i + StartRow,j,ValArr(j - StartCol),True,False
                End If
            Next 'j
        Next 'i
    End If
End Function' InsertTable

'判断是否包含指定字符串
Function IsConTain(ByVal TempStr,ByVal ReplaceValue)
    If Replace(TempStr,ReplaceValue,"") = TempStr Then
        IsConTain = 0
    Else
        IsConTain = 1
    End If
End Function' IsConTain