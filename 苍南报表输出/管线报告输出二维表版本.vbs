
'========================================================Doc����������ļ�·����������================================================================

'Docȫ�ֶ���
Dim Global_Word
Set Global_Word = CreateObject ("asposewordscom.asposewordshelper")

'·����������
Dim FileSysObj
Set FileSysObj = CreateObject("Scripting.FileSystemObject")

'============================================================�ֶ�&�滻�ֶ�����====================================================

'KeyStr = "���,��Ŀ����,��Ŀ��ַ,��Ƶ�λ,���赥λ,ί�е�λ,��浥λ,��ҵʱ��,�����ϲ�ֵ,�߳����ϲ�ֵ,������ϲ�ֵ"

'TemplateVal = "BH,XMMC,XMDZ,SJDW,JSDW,WTDW,CHDW,WYSJ,MaxPoi,MaxHei,MaxDeep"

XMZD = "BH,XMMC,XMDZ,SJDW,JSDW,WTDW,CHDW,WYSJ,CHSJ"

KeyStr = "���,��Ŀ����,��Ŀ��ַ,��Ƶ�λ,���赥λ,ί�е�λ,��浥λ,��ҵʱ��,���ʱ��"

ReplaceVal = "CHSJ,CGTMC"

'===========================================������Ϣ=======================================================

'������Ŀ��Ϣ����
Dim GXProjectInfo()

'===========================================�������========================================================

'�����
Sub OnClick()
    
    If  TypeName (Global_Word) = "AsposeWordsHelper" Then
        Global_Word.CreateDocumentByTemplate  SSProcess.GetSysPathName (7) & "���ģ��\" & "�������ģ��.doc"
    Else
        MsgBox "����ע��Aspose.Word���"
        Exit Sub
    End If
    
    AllVisible
    
    InputInfo ExportFormat,BoolStr,GXProjectInfo,XMMC
    
    If BoolStr = 0 Then
        MsgBox"ȡ����������˳�"
        Exit Sub
    End If
    
    ReplaceValue KeyStr,XMZD,DelCount,DelNodeRow
    
    DelNodeParagraph 0,DelCount,DelNodeRow
    
    InnerGXTable 2,1
    
    InnerGZTable 3,1,HjRow,ExportFormat
    
    InnerHj 3,1,HjRow
    
    Global_Word.SaveEx  SSProcess.GetSysPathName(5) & XMMC & "���߱���.doc"
    
    Ending
    
End Sub' OnClick

'===========================================��Ϣ¼��======================================================

'������Ϣ¼�뺯��
Function InputInfo(ByRef ExportFormat,ByRef BoolStr,ByRef GXProjectInfo(),ByRef XMMC)
    
    ReDim GXProjectInfo(8)
    
    ProJectName = SSProcess.GetProjectFileName
    KeyArr = Split(KeyStr,",", - 1,1)
    XMZDArr = Split(XMZD,",", - 1,1)
    
    SqlStr = "Select ������Ŀ��Ϣ��." & XMZD & " From ������Ŀ��Ϣ��  WHERE ������Ŀ��Ϣ��.ID=1"
    
    GetSQLRecordAll SqlStr,ProJectInfoArr,ResultCount
    
    InfoArr = Split(ProJectInfoArr(0), ",", - 1,1)
    XMMC = InfoArr(1)
    SSProcess.ClearInputParameter
    SSProcess.AddInputParameter "���߳���ͳ�Ʒ�ʽ","��ά����",0,"��ά����,��ά����","���߳���ͳ�Ʒ�ʽ"
    For i = 0 To  8
        SSProcess.AddInputParameter KeyArr(i), InfoArr(i), 0, "", ""
    Next 'i
    
    BoolStr = SSProcess.ShowInputParameterDlg ("���߱��������ʽ")
    ExportFormat = SSProcess.GetInputParameter("���߳���ͳ�Ʒ�ʽ")
    
    If BoolStr = 1 Then
        SSProcess.OpenAccessMdb ProJectName
        For i = 0  To 8
            GXProjectInfo(i) = SSProcess.GetInputParameter(KeyArr(i))
            SqlStr = "Update  ������Ŀ��Ϣ�� Set " & XMZDArr(i) & " = '" & GXProjectInfo(i) & "'Where ������Ŀ��Ϣ��.ID= 1"
            SSProcess.ExecuteAccessSql  ProJectName,SqlStr
        Next
        SSProcess.CloseAccessMdb ProJectName
    End If
    
End Function' InputInfo

'==========================================================��ȡС������&��д���=======================================================

EngStr = "CD,CT,CY,CQ,CS,QT,BM,FQ,DL,GD,LD,DC,XH,TX,DX,YD,LT,JX,JK,EX,DS,BZ,JS,XF,PS,YS,WS,FS,RQ,MQ,TR,YH,RL,RS,ZQ,SY,GS"

CheStr = "�������,����ͨ��,�����ܵ�,��Ȼ�����ܵ�,ˮ���ܵ�,�������ܵ�,����,����,����,����,·��,�糵,��ͨ�ź�,�� ��,����,�ƶ�,��ͨ,����,���,����ͨѶ,�㲥����,����ר��,���ҵ��ˮ,����ˮ,��ˮ,��ˮ,��ˮ,�����ˮ,ȼ��,ú��,��Ȼ��,Һ����,����,��ˮ,����,ʯ��,��ҵ��ˮ"

'��д���߲��ȡ���׼��
Function InnerGXTable(ByVal TableIndex,ByVal StartRow) 'TableIndex �������,StartRow ��ʼ����
    StrString = "Select DISTINCT GXLX From ���¹��ߵ����Ա� inner join GeoPointTB on ���¹��ߵ����Ա�.ID = GeoPointTB.ID WHERE (GeoPointTB.Mark Mod 2)<>0 And ���¹��ߵ����Ա�.GXLX <>'*' And ���¹��ߵ����Ա�.GXLX <>''"
    GetSQLRecordAll StrString,LxArr,LxCount
    If LxCount > 1 Then
        Global_Word.CloneTableRow TableIndex,StartRow,1,LxCount - 1,False
        For i = 0 To LxCount - 1
            Global_Word.SetCellText TableIndex,i + StartRow,0,ToChinese(LxArr(i)),True,False
            If ToChinese(LxArr(i)) = "���ҵ��ˮ" Then
                Global_Word.SetCellText TableIndex,i + StartRow,1,"�ܾ���50mm",True,False
            ElseIf ToChinese(LxArr(i)) = "��ˮ" Then
                Global_Word.SetCellText TableIndex,i + StartRow,1,"�ܾ���200mm�򷽹���400mm��400mm",True,False
            ElseIf  ToChinese(LxArr(i)) <> "" Then
                Global_Word.SetCellText TableIndex,i + StartRow,1,"ȫ��",True,False
            End If
        Next 'i
    Else
        For i = 0 To LxCount - 1
            Global_Word.SetCellText TableIndex,i + StartRow,0,ToChinese(LxArr(i)),True,False
            If ToChinese(LxArr(i)) = "���ҵ��ˮ" Then
                Global_Word.SetCellText TableIndex,i + StartRow,1,"�ܾ���50mm",True,False
            ElseIf ToChinese(LxArr(i)) = "��ˮ" Then
                Global_Word.SetCellText TableIndex,i + StartRow,1,"�ܾ���200mm�򷽹���400mm��400mm",True,False
            Else
                Global_Word.SetCellText TableIndex,i + StartRow,1,"ȫ��",True,False
            End If
        Next 'i
    End If
End Function' InnerGXTable

'��д��רҵ���߹�����ͳ�Ʊ�
Function InnerGZTable(ByVal TableIndex,ByVal StartRow,ByRef HjRow,ByVal LenTypes) 'TableIndex �������,StartRow ��ʼ����,HjRow �ϼ�����ֵ(����ֵ),LenTypes ��������
    StrString = "Select DISTINCT GXLX From ���¹��������Ա� inner join GeoLineTB on ���¹��������Ա�.ID = GeoLineTB.ID WHERE (GeoLineTB.Mark Mod 2)<>0 And ���¹��������Ա�.GXLX <>'*' And ���¹��������Ա�.GXLX <>''"
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

'��д���Ե�����ε����
Function InnerPoiCount(ByVal GxName,ByVal TableIndex,ByVal InsertRow) 'GxName ������������,TableIndex ������,InsertRow ָ��������
    StrString = "Select FSW From ���¹��ߵ����Ա� inner join GeoPointTB on ���¹��ߵ����Ա�.ID = GeoPointTB.ID WHERE (GeoPointTB.Mark Mod 2)<>0 And ���¹��ߵ����Ա�.GXLX =" & "'" & GxName & "'"
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

'��д���߳���
Function InnerLineLen(ByVal GxName,ByVal TableIndex,ByVal InsertRow,ByVal LenTypes) 'GxName ������������,TableIndex ������,InsertRow ָ��������,LenTypes ��������
    SelFeatures GxName,LineCount,LineArr
    If LenTypes = "��ά����" Then
        For i = 0 To LineCount - 1
            TotalLength = TotalLength + Round(Transform(SSProcess.GetObjectAttr(LineArr(i),"SSObj_Length")),0)
            'msgbox   TotalLength
        Next 'i
    ElseIf LenTypes = "��ά����" Then
        For i = 0 To LineCount - 1
            TotalLength = TotalLength + Round(Transform(SSProcess.GetObjectAttr(LineArr(i),"SSObj_3DLength")),0)
        Next 'i
    End If
    Global_Word.SetCellText TableIndex,InsertRow,4,TotalLength,True,False
End Function' InnerLineLen

'��д�ϼ���
Function InnerHj(ByVal TableIndex,ByVal StartRow,ByVal HjRow) 'TableIndex ������,StartRow ��ʼ����,HjRow �ϼ���
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

'ɾ��ָ����
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

'����ת��Ϊ����
Function ToChinese(ByVal EngLayerName) 'EngLayerName ͼ������(Ӣ��)
    EngArr = Split(EngStr,",", - 1,1)
    CheArr = Split(CheStr,",", - 1,1)
    ToChinese = ""
    For i = 0 To UBound(EngArr)
        If EngArr(i) = EngLayerName Then
            ToChinese = CheArr(i)
        End If
    Next 'i
End Function' ToChinese

'=========================================================�ַ����滻=======================================================

' [���߱�����Ϣ]
' ��� = ""
' ��Ŀ���� = ""
' ��Ŀ��ַ = ""
' ��Ƶ�λ = ""
' ���赥λ = ""
' ί�е�λ = ""
' ��ҵʱ�� = ""
' ���ʱ�� = ""
' �����ϲ�ֵ = ""
' �߳����ϲ�ֵ = ""
' ������ϲ�ֵ = ""

'�ַ��滻����


'�ַ����滻����
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
    
    StrString = "Select DISTINCT GXLX From ���¹��������Ա� inner join GeoLineTB on ���¹��������Ա�.ID = GeoLineTB.ID WHERE (GeoLineTB.Mark Mod 2)<>0"
    GetSQLRecordAll StrString,LxArr,LxCount
    For i = 0 To LxCount - 1
        If LxArr(i) <> "*" Then
            GXTStr = GXTStr & GXProjectInfo(1) & ToChinese(LxArr(i)) & "����ͼ" & Chr(13)
        End If
    Next 'i
    
    ExtraVal = ToBigDate(GetNowTime) & "," & GXTStr
    ExtraArr = Split(ExtraVal,",", - 1,1)
    
    NameArr = Split(ReplaceVal,",", - 1,1)
    For i = 0 To UBound(ExtraArr)
        Global_Word.Replace "{" & NameArr(i) & "}",ExtraArr(i),0
    Next 'i
    
    SqlStr = "Select ������Ŀ��Ϣ��." & "DWZDJC,GCZDJC,MSZDJC" & " From ������Ŀ��Ϣ��  WHERE ������Ŀ��Ϣ��.id=1"
    GetSQLRecordAll SqlStr,MaxNumArr,ResultCount
    
    ValArr = Split(MaxNumArr(0), ",", - 1,1)
    
    Global_Word.Replace "{" & "DWZDJC" & "}",ValArr(0),0
    Global_Word.Replace "{" & "GCZDJC" & "}",ValArr(1),0
    Global_Word.Replace "{" & "MSZDJC" & "}",ValArr(2),0
    
End Function' ReplaceValue

'==========================================================�����ຯ��=======================================================

'������ͼ��
Function AllVisible()
    count = SSProcess.GetLayerCount
    For i = 0 To count - 1
        layername = SSProcess.GetLayerName (i)
        SSProcess.SetLayerStatus layername, 1, 1
    Next
    SSProcess.RefreshView
End Function

'���ת��д
Function YearChange(ByVal YearName)
    Number = "1,2,3,4,5,6,7,8,9,0"
    BigNumber = "һ,��,��,��,��,��,��,��,��,��"
    NumberArr = Split(Number,",", - 1,1)
    BigNumberArr = Split(BigNumber,",", - 1,1)
    For i = 1 To 4
        For j = 0 To UBound(NumberArr)
            If Mid(YearName,i,1) = NumberArr(j) Then
                YearChange = YearChange & BigNumberArr(j)
            End If
        Next 'j
    Next 'i
    YearChange = YearChange & "��"
End Function' YearChange

'�·�ת��д
Function MonthChange(ByVal MonthName)
    Number = "1,2,3,4,5,6,7,8,9,10,11,12"
    BigNumber = "һ,��,��,��,��,��,��,��,��,ʮ,ʮһ,ʮ��"
    NumberArr = Split(Number,",", - 1,1)
    BigNumberArr = Split(BigNumber,",", - 1,1)
    For i = 0 To UBound(NumberArr)
        If MonthName = NumberArr(i) Then
            MonthChange = BigNumberArr(i) & "��"
        End If
    Next 'i
End Function' MonthChange

'����ת��д
Function ToBigDate(ByVal DateStr)
    YearMonStr = Split(DateStr,"��", - 1,1)
    YeraName = Left(YearMonStr(0),4)
    MonName = Mid(YearMonStr(0),6)
    ToBigDate = YearChange(YeraName) & MonthChange(MonName)
End Function

'��ȡ��ǰϵͳʱ��
Function GetNowTime()
    GetNowTime = FormatDateTime(Now(),1)
End Function' GetNowTime

'ѡ��ָ�����ﲢ���ظ���
Function SelFeatures(ByVal EngLayerName,ByRef Count,ByRef IdArr()) 'EngLayerName ͼ������(Ӣ��),Count ����(����ֵ),IdArr() Id����(����ֵ)
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

'��ȡ���м�¼
Function GetSQLRecordAll(ByVal StrSqlStatement, ByRef SQLRecord(), ByRef iRecordCount)
    ProJectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb ProJectName
    If StrSqlStatement = "" Then
        MsgBox "��ѯ���Ϊ�գ�����ֹͣ��",48
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

'��ȡ��Ԫ��ֵ
Function GetSelCellVal(ByVal CellContent)
    GetSelCellVal = Left(CellContent,Len(CellContent) - 1)
End Function' GetSelCellVal

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

'�����ʾ
Function Ending()
    MsgBox "������"
End Function' Ending