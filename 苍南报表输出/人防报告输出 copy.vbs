' һ����ȡ����ģ��·�������Ƶ����·��
' 1����ȡ��ǰ��ģ��·��������Word
' 2������Word����·�����ļ���

' ����������Ŀ��������ֵ�����ַ����滻
' 1����ȡ��Ŀ���ߵ����ԣ���ȡ��Ŀ��Ϣʱ��Ҫȥ�����ţ�

'========================================================Doc����������ļ�·����������================================================================

'Docȫ�ֶ���
Dim Global_Word
Set Global_Word = CreateObject ("asposewordscom.asposewordshelper")

'·����������
Dim FileSysObj
Set FileSysObj = CreateObject("Scripting.FileSystemObject")

'===================================================================�������=======================================================================

Sub OnClick()
    If  TypeName (Global_Word) = "AsposeWordsHelper" Then
        Global_Word.CreateDocumentByTemplate  SSProcess.GetSysPathName (7) & "���ģ��\" & "�˷���������ģ��.doc"
    Else
        MsgBox "����ע��Aspose.Word���"
        Exit Sub
    End If
    
    AllVisible
    ReplaceValue Yjmj,SPYJmj,ZhuZhaiXs,FZhuZhaiXs,HxId
    If ToDecimal(ZhuZhaiXs) = 0 Or ToDecimal(FZhuZhaiXs) = 0 Then
        MsgBox "סլϵ�����סլϵ��Ϊ��"
        Exit Sub
    End If
    RfResultTableInner RfCount,RfValArr,Yjmj,SPYJmj
    RfYJTableInner Yjmj,ZhuZhaiXs,FZhuZhaiXs,HxId
    
    Global_Word.SaveEx  SSProcess.GetSysPathName(5) & "�ɹ��ļ�" & "\�˷�����.doc"
    MsgBox "������"
End Sub' OnClick

'======================================================�ַ����滻==============================================================================

'�ַ����滻
Function ReplaceValue(ByRef Yjmj,ByRef SPYJmj,ByRef ZhuZhaiXs,ByRef FZhuZhaiXs,ByRef HxId)
    
    RePlaceStr = "XiangMMC,XiangMDZ,SheJDW,JianSDW,WeiTDW,CeLDW,CeLRQ,DXCS,JianZJG,JunGCLDSJZMJ,DSCS,JunGCLZTS,JunGCLDXJZMJ,ZZRFYJMJ,QTRFYJMJ,HLHTMJ,FKYBMJ,ZBYTHD,BPGC,SPRFYJMJ"
    
    SelFeatures "9130223",HxCount,HxId
    GetInnerFeatures "9130223","9210123",ZrzCount,ZrzArr
    
    ReDim DateArr(ZrzCount)
    For i = 0 To ZrzCount - 1
        DateArr(i) = Transform(Replace(Replace(Replace(SSProcess.GetObjectAttr(ZrzArr(i),"[JGRQ]"),"��",""),"��",""),"��",""))
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
            If RePlaceArr(i) = "DXCS" Then
                DXCS = SSProcess.GetSelGeoValue(0, "[" & RePlaceArr(i) & "]")
                DXCS = Replace(DXCS,"-","")
                Global_Word.Replace "{" & RePlaceArr(i) & "}",dxcs,0
            Else
                szaa = SSProcess.GetSelGeoValue(0, "[" & RePlaceArr(i) & "]")
                If szaa = "0" Or szaa = "0.0"  Then
                    Global_Word.Replace "{" & RePlaceArr(i) & "}","",0
                Else
                    Global_Word.Replace "{" & RePlaceArr(i) & "}",SSProcess.GetSelGeoValue(0, "[" & RePlaceArr(i) & "]"),0
                End If
            End If
        Next 'i
    End If
    
    SqlStr = "Select DISTINCT PSGN From �˷�������Ԫ���Ա� inner join GeoAreaTB on �˷�������Ԫ���Ա�.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0"
    GetSQLRecordAll SqlStr,DISTINCTArr,LxCount
    Global_Word.Replace "{PSGN}",DISTINCTArr(0),0
    
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
    Global_Word.Replace "��YJMJ��",Yjmj,0
    Global_Word.Replace "{TTT}",UpDateTime,0
End Function' ReplaceValue

'==========================================================�˷������ɹ���====================================================================

'�˷������ɹ�����ں���
Function RfResultTableInner(ByRef RfCount,ByRef RfValArr(),ByVal Yjmj,ByVal SPYJmj)
    RfClMjInert RfCount,RfValArr,Yjmj,SPYJmj
End Function' RfResultTableInner

'�˷�������������ֵ
Function RfClMjInert(ByRef RfCount,ByRef RfValArr(),ByVal Yjmj,ByVal SPYJmj)
    'GetInnerFeatures "9130223","9530226",RfCount,RfArr
    SqlStr = "Select �˷�������Ԫ���Ա�.ID From �˷�������Ԫ���Ա� Inner Join GeoAreaTB on �˷�������Ԫ���Ա�.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0"
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
    
    InsertTable RfValArr,RfCount,2,5,1,8,14,HjRow
    InsertSum 2,5,HjRow - 1,HjRow,2
    InsertSum 2,5,HjRow - 1,HjRow,3
    InsertSum 2,5,HjRow - 1,HjRow,8
    InsertSum 2,5,HjRow - 1,HjRow,9
    InsertSum 2,5,HjRow - 1,HjRow,12
    InsertSum 2,5,HjRow - 1,HjRow,13
    InsertSm 5,HjRow - 1,RfCount,Yjmj,SPYJmj
End Function' RfClMjInert

'��д�ϼ�ֵ
Function InsertSum(ByVal TableIndex,ByVal StartRow,ByVal EndRow,ByVal HjRow,ByVal CalCol)
    TotalArea = 0
    For i = StartRow To EndRow
        If Global_Word.GetCellText(TableIndex,StartRow,CalCol,False) <> "" Then
            TotalArea = TotalArea + Transform(GetSelCellVal(Global_Word.GetCellText(TableIndex,i,CalCol,False)))
        End If
    Next 'i
    Global_Word.SetCellText Tableindex,HjRow,CalCol,TotalArea,True,False
End Function' InsertSum

'��д�˷�����˵������
Function InsertSm(ByVal StartRow,ByVal EndRow,ByVal RfCount,ByVal Yjmj,ByVal SPYJmj)
    RePlaceArr = Split("STR1,STR2,STR3",",", - 1,1)
    GetYbArea StartRow,EndRow,AreaArr,YBString
    ValArr = Split(RfCount & ";" & Transform(GetSelCellVal(Global_Word.GetCellText(2,EndRow + 1,2,False))) & ";" & YBString,";", - 1,1)
    For i = 0 To UBound(RePlaceArr)
        Global_Word.Replace "��" & RePlaceArr(i) & "��",ValArr(i),0
    Next 'i
    If (Transform(GetSelCellVal(Global_Word.GetCellText(2,EndRow + 1,2,False))) - SPYJmj) >= 0 Then
        Replace4 = "ʵ�ʳ����" & Round(Transform(GetSelCellVal(Global_Word.GetCellText(2,EndRow + 1,2,False))) - SPYJmj,2)
    Else
        Replace4 = "ʵ��С�����" & Abs(Round(Transform(GetSelCellVal(Global_Word.GetCellText(2,EndRow + 1,2,False))) - SPYJmj,2))
    End If
    Global_Word.Replace "��" & "STR4" & "��",Replace4,0
    Global_Word.Replace "��" & "STR5" & "��","ʵ��Ӧ���˷����" & Yjmj,0
    If (Transform(GetSelCellVal(Global_Word.GetCellText(2,EndRow + 1,2,False))) - Yjmj) >= 0 Then
        Replace6 = "ʵ�ʳ����" & Transform(GetSelCellVal(Global_Word.GetCellText(2,EndRow + 1,2,False))) - Yjmj
    Else
        Replace6 = "ʵ��С�����" & Abs(Transform(GetSelCellVal(Global_Word.GetCellText(2,EndRow + 1,2,False))) - Yjmj)
    End If
    Global_Word.Replace "��" & "STR6" & "��",Replace6,0
    '4:Transform(GetSelCellVal(Global_Word.GetCellText(2,EndRow + 1,2,False))) - SPYJmj
    '5:Yjmj
    '6:Transform(GetSelCellVal(Global_Word.GetCellText(2,EndRow + 1,2,False))) - Yjmj
End Function' InsertSm

'��ȡ��ǰ�ڱε�Ԫ������
Function GetYbLx(ByVal StartRow,ByVal EndRow,ByRef Lxstring,ByRef lxCount)
    GnStr = ""
    For i = StartRow To EndRow
        If GnStr = "" And GetSelCellVal(Global_Word.GetCellText(2,i,5,False)) <> "" Then
            GnStr = GetSelCellVal(Global_Word.GetCellText(2,i,5,False))
        ElseIf GetSelCellVal(Global_Word.GetCellText(2,i,5,False)) <> "" Then
            GnStr = GnStr & "," & GetSelCellVal(Global_Word.GetCellText(2,i,5,False))
        End If
    Next 'i
    DelRepeat Split(GnStr,",", - 1,1),Lxstring,lxCount
End Function' GetYbLx

'��ȡ��Ӧ���ڱ�������ַ���
Function GetYbArea(ByVal StartRow,ByVal EndRow,ByRef AreaArr(),ByRef TotalString)
    TotalString = ""
    GetYbLx StartRow,EndRow,Lxstring,lxcount
    LxArr = Split(Lxstring,",", - 1,1)
    ReDim AreaArr(lxcount,2)
    For i = 0 To UBound(LxArr)
        SingleLxCount = 0
        For j = StartRow To EndRow
            If LxArr(i) = GetSelCellVal(Global_Word.GetCellText(2,j,5,False)) Then
                SingleLxCount = SingleLxCount + 1
                AreaArr(i,1) = AreaArr(i,1) + Transform(GetSelCellVal(Global_Word.GetCellText(2,j,3,False)))
                AreaArr(i,0) = LxArr(i)
            End If
        Next 'j
        If TotalString = "" Then
            TotalString = "����" & SingleLxCount & "��" & AreaArr(i,0) & "��Ԫ��" & AreaArr(i,1) & "ƽ����"
        Else
            TotalString = TotalString & "," & SingleLxCount & "��" & AreaArr(i,0) & "��Ԫ��" & AreaArr(i,1) & "ƽ����"
        End If
    Next 'i
End Function' GetYbArea

'========================================================�˷�Ӧ����������============================================================================

'�˷�Ӧ������������ں���
Function RfYJTableInner(ByVal Yjmj,ByVal ZhuZhaiXs,ByVal FZhuZhaiXs,ByVal HxId)
    InsertZhuZ 3,HxId,2,EndRow,HjRow,ZhuZhaiXs
    InsertFZhuZ HxId,2,EndRow,3,HjRow,FZhuZhaiXs
End Function' RfResultTableInner

'��дסլ��������
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

'��д��סլ��������
Function InsertFZhuZ(ByVal HxId,ByVal StartRow,ByRef EndRow,ByVal TableIndex,ByRef HjRow,ByVal FZhuZhaiXs)
    SearchMjk ZrzCount,HxId,2,EndRow,TableIndex,FZhuZhaiXs,HjRow
    SearchH ZrzCount,HxId,EndRow,iEndRow,TableIndex,FZhuZhaiXs,HjRow
    InsertFZhuZSum TableIndex,2,iEndRow,HjRow
End Function' InsertFZhuZ

'����ZDGUID��ͬ����Ȼ������
Function SearchZrz(ByVal HxId,ByRef ZrzArr,ByRef ZrzCount)
    SqlString = "Select ZRZH,ZhuZMJ,ZRZGUID From FC_��Ȼ����Ϣ���Ա� inner join GeoAreaTB on FC_��Ȼ����Ϣ���Ա�.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0 AND FC_��Ȼ����Ϣ���Ա�.ZDGUID = " & SSProcess.GetObjectAttr(HxId,"[ZDGUID]")
    GetSQLRecordAll SqlString,ZrzArr,ZrzCount
End Function' SearchZrz

'����ZRZGUID��ͬ������鲢��ֵ
Function SearchMjk(ByRef ZrzCount,ByVal HxId,ByRef StartRow,ByRef EndRow,ByVal TableIndex,ByVal FZhuZhaiXs,ByRef HjRow)
    SearchZrz HxId,ZrzArr,ZrzCount
    StartRow = StartRow
    EndRow = StartRow
    If ZrzCount > 0 Then
        For i = 0 To ZrzCount - 1
            SingleArr = Split(ZrzArr(i),",", - 1,1)
            SqlString = "Select DISTINCT MJKMC From FC_�������Ϣ���Ա� inner join GeoAreaTB on FC_�������Ϣ���Ա�.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0 AND FC_�������Ϣ���Ա�.ZRZGUID = " & SingleArr(2) & " AND FC_�������Ϣ���Ա�.FTLX = '����̯'"
            GetSQLRecordAll SqlString,MjkMcArr,MjkLxCount
            If MjkLxCount > 0 Then
                For j = 0 To MjkLxCount - 1
                    SqlString = "Select Sum(FC_�������Ϣ���Ա�.KZMJ) From FC_�������Ϣ���Ա� inner join GeoAreaTB on FC_�������Ϣ���Ա�.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0 AND FC_�������Ϣ���Ա�.MJKMC = " & "'" & MjkMcArr(j) & "'" & " AND FC_�������Ϣ���Ա�.FTLX = '����̯'"
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

'����ZRZGUID��ͬ�Ļ�����ֵ
Function SearchH(ByRef ZrzCount,ByVal HxId,ByRef StartRow,ByRef EndRow,ByVal TableIndex,ByVal FZhuZhaiXs,ByRef HjRow)
    SearchZrz HxId,ZrzArr,ZrzCount
    For i = 0 To ZrzCount - 1
        SingleArr = Split(ZrzArr(i),",", - 1,1)
        SqlString = "Select DISTINCT SYGN From FC_����Ϣ���Ա� inner join GeoAreaTB on FC_����Ϣ���Ա�.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0 AND NOT INSTR(FC_����Ϣ���Ա�.SYGN,'סլ') AND FC_����Ϣ���Ա�.ZRZGUID = " & "'" & SingleArr(2) & "'"
        GetSQLRecordAll SqlString,SYGNArr,SYGNCount
        For j = 0 To SYGNCount - 1
            SqlString = "Select Sum(FC_����Ϣ���Ա�.JZMJ) From FC_����Ϣ���Ա� inner join GeoAreaTB on FC_����Ϣ���Ա�.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0 AND FC_����Ϣ���Ա�.SYGN = " & "'" & SYGNArr(j) & "'" & " AND FC_����Ϣ���Ա�.ZRZGUID = " & "'" & SingleArr(2) & "'"
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

'�����סլ�ϼ�ֵ
Function InsertFZhuZSum(ByVal TableIndex,ByVal StartRow,ByVal EndRow,ByVal HjRow)
    For i = StartRow To EndRow
        TotalArea = TotalArea + Transform(GetSelCellVal(Global_Word.GetCellText(TableIndex,i,7,False)))
    Next 'i
    Global_Word.SetCellText TableIndex,HjRow,3,TotalArea,True,False
End Function' InsertFZhuZSum
'==============================================================�����ຯ��===============================================================

'������ͼ��
Function AllVisible()
    count = SSProcess.GetLayerCount
    For i = 0 To count - 1
        layername = SSProcess.GetLayerName (i)
        SSProcess.SetLayerStatus layername, 1, 1
    Next
    SSProcess.RefreshView
End Function

'��������ת��
Function Transform(ByVal Values)
    If Values <> "" Then
        Values = CDbl(Values)
    Else
        Values = 0
        Exit Function
    End If
    Transform = Values
End Function'Transform

'ȥ���ַ������ظ�ֵ
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

'ѡ��ָ�����ﲢ���ظ���
Function SelFeatures(ByVal Code,ByRef Count,ByRef ID)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_code", "==", Code
    SSProcess.SelectFilter
    Count = SSProcess.GetSelGeoCount
    ID = SSProcess.GetSelGeoValue(0,"SSObj_ID")
End Function' SelFeatures

'ѡ��ָ�������ڲ���ָ�����ﲢ���ص���id����
Function GetInnerFeatures(ByVal OuterCode,ByVal InnerCode ,ByRef InnerCount,ByRef InnerArr())
    SelFeatures OuterCode,OuterCount,HxID
    InnerArr = Split(SSProcess.SearchInnerObjIDs(HxID,2,InnerCode,0),",", - 1,1)
    InnerCount = UBound(InnerArr) + 1
End Function' GetInnerFeatures

'��ȡ��Ԫ��ֵ
Function GetSelCellVal(ByVal CellContent)
    GetSelCellVal = Left(CellContent,Len(CellContent) - 1)
End Function' GetSelCellVal

'�ٷֺ�תС��
Function ToDecimal(ByVal Percentage)
    ToDecimal = Transform(Left(Percentage,Len(Percentage) - 1)) * 0.01
End Function' ToDecimal

'��ȡ���м�¼
Function GetSQLRecordAll(ByVal StrSqlStatement, ByRef SQLRecord(), ByRef iRecordCount)
    SSProcess.OpenAccessMdb SSProcess.GetProjectFileName
    If StrSqlStatement = "" Then
        MsgBox "��ѯ���Ϊ�գ�����ֹͣ��",48
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

'��д���
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

'�ж��Ƿ����ָ���ַ���
Function IsConTain(ByVal TempStr,ByVal ReplaceValue)
    If Replace(TempStr,ReplaceValue,"") = TempStr Then
        IsConTain = 0
    Else
        IsConTain = 1
    End If
End Function' IsConTain