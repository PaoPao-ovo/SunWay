
'==================================================���߱���������߱���========================================================

MainCode = "54311203,54324004,54323004,54412004,54423004,54452004,54511004,54512114,54534114,54523114,54611114,54612004,54623004,54111003,54112003,54123003,54145003,54134003,54211003,54212003,54223003,54234003,54245003,54256003,54267003,54278003,54289003,54720114,54730114,54030003,54040003,51011203,52011203,53011204,53022204,53033204,53044204"

HiddenCode = "54100004,54200304,54245304,54256304,54267304,54412005,54423005,54452005,54111004,54211304,54400005,54411005,54212304,54223304,54120004,54130004,54140004,54150004,54234304,54278304,54289304"

'=======================================================�������=========================================================

'�����
Sub OnClick()
    ConFirmWay Way,res,GroupStr
    'Way = "�ۺϹ���ͼ"
    If res = 1 Then
        If Way = "�ۺϹ���ͼ" Then
            AllVisible
            DelTk
            GxVisible "CD,CT,CY,CQ,CS,QT,BM,FQ,DL,GD,LD,DC,XH,TX,DX,YD,LT,JX,JK,EX,DS,BZ,JS,XF,PS,YS,WS,FS,RQ,MQ,TR,YH,RL,RS,ZQ,SY,GS"
            CreatMap Way,ContinueIf
            If ContinueIf = 0 Then
                MsgBox "�����ڹ�������"
                Exit Sub
            End If
            AllVisible
            DelTk
            FYNOTE GroupStr
            Ending
        ElseIf Way = "�ֲ����" Then
            AllVisible
            DelTk
            FCExport Way
            AllVisible
            DelTk
            FYNOTE GroupStr
            Ending
        Else
            MsgBox "δѡ�������ʽ"
            Exit Sub
        End If
    End If
End Sub' OnClick
Function FYNOTE(STR)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", "CD��ע��,CT��ע��,CY��ע��,CQ��ע��,CS��ע��,QT��ע��,BM��ע��,FQ��ע��,DL��ע��,GD��ע��,LD��ע��,DC��ע��,XH��ע��,TX��ע��,DX��ע��,YD��ע��,LT��ע��,JX��ע��,JK��ע��,EX��ע��,DS��ע��,BZ��ע��,JS��ע��,XF��ע��,PS��ע��,YS��ע��,WS��ע��,FS��ע��,RQ��ע��,MQ��ע��,TR��ע��,YH��ע��,RL��ע��,RS��ע��,ZQ��ע��,SY��ע��,GS��ע��"
    SSProcess.SetSelectCondition "SSObj_Type", "==", "NOTE"
    SSProcess.SelectFilter
    Count = SSProcess.GetSelNoteCount
    
    For j = 0 To Count - 1
        FormerVal = SSProcess.GetSelNoteValue(j,"SSObj_FontString")
        IDStr = SSProcess.GetSelNoteValue(j,"SSObj_ID")
        ws = Len(str) + 2
        qbwz = Len(FormerVal)
        rwz = qbwz - ws
        hmzte = Right(FormerVal,rwz)
        q2 = Left(FormerVal,2)
        fystr = q2 & hmzte
        'if j=0 then msgbox  fystr
        SSProcess.SetObjectAttr IDStr, "SSObj_FontString", fystr
    Next
End Function
'===================================================��չ�����޸�========================================================

' [����CAD���]
' ��ע=
' ͼ������=
' ��ҵ��λ=�����ز��Ժ
' ί�е�λ=
' ��������=2023��7�¼������ͼ
' ƽ��������ϵ=���ϳ�������ϵ
' �߳���ϵ=1985���Ҹ̻߳�׼���ȸ߾�0.5�ס�
' ͼʽ=2017���ͼʽ
' ̽��Ա=����
' ����Ա=����
' ��ͼԱ=����
' ���Ա=����

AttrStr = "����Ȩ��λ,ί�е�λ,��������,ƽ��������ϵ,�߳���ϵ,ͼʽ,̽��Ա,����Ա,��ͼԱ,���Ա"
KeyStr = "��ҵ��λ,ί�е�λ,��������,ƽ��������ϵ,�߳���ϵ,ͼʽ,̽��Ա,����Ա,��ͼԱ,���Ա"

Function ModifyAttr(ByVal Code,ByVal Way,ByVal TkId,ByRef XmMc,ByRef Count)
    SelFeature Code,TkId,Count
    TkArr = Split(TkId,",", - 1,1)
    If Count = 0 Then Exit Function
    AttrArr = Split(AttrStr,",", - 1,1)
    KeyArr = Split(KeyStr,",", - 1,1)
    For i = 0 To UBound(AttrArr)
        For j = 0 To UBound(TkArr)
            SSProcess.SetObjectAttr TkArr(j),"[" & AttrArr(i) & "]",SSProcess.ReadEpsIni("����CAD���", KeyArr(i) ,"")
        Next 'j
    Next 'i
    SqlStr = "Select XMMC From ������Ŀ��Ϣ�� Where ������Ŀ��Ϣ��.ID = 1"
    GetSQLRecordAll SqlStr,XmmcArr,Count
    If Count > 0 Then
        XmMc = XmmcArr(0)
    End If
    For i = 0 To UBound(TkArr)
        SSProcess.SetObjectAttr TkArr(i),"[ͼ������]",XmMc
        SSProcess.SetObjectAttr TkArr(i),"[��ע]",Way
        SSProcess.ObjectDeal TkArr(i), "FreeDisplayList", Parameters, Result
    Next 'i
    SSProcess.RefreshView
End Function' ModifyAttr

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

Function SetFcAttr(ByVal Code,ByRef TkId,ByRef XmMc,ByVal BigName)
    AttrArr = Split(AttrStr,",", - 1,1)
    KeyArr = Split(KeyStr,",", - 1,1)
    For i = 0 To UBound(AttrArr)
        SSProcess.SetObjectAttr TkId,"[" & AttrArr(i) & "]",SSProcess.ReadEpsIni("����CAD���", KeyArr(i) ,"")
    Next 'i
    SqlStr = "Select XMMC From ������Ŀ��Ϣ�� Where ������Ŀ��Ϣ��.ID = 1"
    GetSQLRecordAll SqlStr,XmmcArr,Count
    If Count > 0 Then
        XmMc = XmmcArr(0)
    End If
    SSProcess.SetObjectAttr TkId,"[ͼ������]",XmMc
    SSProcess.SetObjectAttr TkId,"[��ע]",BigName & "���¹���ͼ"
    SSProcess.ObjectDeal TkId, "FreeDisplayList", Parameters, Result
    SSProcess.RefreshView
End Function' SetFcAttr

'ѡ��ǰͼ��������ͼ��ID
Function SelFeature(ByVal Code,ByVal ID,ByRef Count)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", Code
    SSProcess.SetSelectCondition "SSObj_ID", "==", ID
    SSProcess.SelectFilter
    Count = SSProcess.GetSelGeoCount
End Function' SelFeature

'ѡ��ǰͼ�����ݲ����ظ���
Function SelData(ByVal LayerName)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", LayerName
    SSProcess.SelectFilter
    SelData = SSProcess.GetSelGeoCount
End Function' SelData

'��ȡ��ǰͼ�������еĹ���ͼ������(����)
Function GetAllLayerName(ByVal OuterId,ByRef SmallArr(),ByRef LayArr())
    ' LayArr = Split("CD,CT,CY,CQ,CS,QT,BM,FQ,DL,GD,LD,DC,XH,TX,DX,YD,LT,JX,JK,EX,DS,BZ,JS,XF,PS,YS,WS,FS,RQ,MQ,TR,YH,RL,RS,ZQ,SY,GS",",", - 1,1)
    ' For j = 0 To UBound(LayArr)
    '     If SelData(LayArr(j)) > 0 Then
    '         If LayerStr = "" Then
    '             LayerStr = LayArr(j)
    '         Else
    '             LayerStr = LayerStr & "," & LayArr(j)
    '         End If
    '     End If
    ' Next 'j
    ' SmallArr = Split(LayerStr,",", - 1,1)
    AllVisible
    AllIdStr = SSProcess.SearchInPolyObjIDs(OuterId,10,"",0,1,1)
    AllArr = Split(AllIdStr,",", - 1,1)
    ReDim SmallArr(UBound(AllArr))
    For i = 0 To UBound(AllArr)
        SmallArr(i) = SSProcess.GetObjectAttr(AllArr(i),"SSObj_LayerName")
    Next 'i
    DelRepeat SmallArr,SmallLayStr,LayerCount
    SmallArr = Split(SmallLayStr,",", - 1,1)
    Count = 0
    ReDim BigArr(Count)
    For i = 0 To UBound(SmallArr)
        If SmallArr(i) = "CD" Then
            BigArr(Count) = "CDD"
            Count = Count + 1
            ReDim Preserve BigArr(Count)
        ElseIf SmallArr(i) = "CT" Then
            BigArr(Count) = "CXD"
            Count = Count + 1
            ReDim  Preserve BigArr(Count)
        ElseIf SmallArr(i) = "CY" Or SmallArr(i) = "CQ" Or SmallArr(i) = "CS" Or SmallArr(i) = "QT" Then
            BigArr(Count) = "CYD"
            Count = Count + 1
            ReDim Preserve BigArr(Count)
        ElseIf SmallArr(i) = "BM" Or SmallArr(i) = "FQ" Then
            BigArr(Count) = "CSD"
            Count = Count + 1
            ReDim Preserve BigArr(Count)
        ElseIf SmallArr(i) = "DL" Or SmallArr(i) = "GD" Or SmallArr(i) = "LD" Or SmallArr(i) = "DC" Or SmallArr(i) = "XH" Then
            BigArr(Count) = "DLD"
            Count = Count + 1
            ReDim Preserve BigArr(Count)
        ElseIf SmallArr(i) = "TX" Or SmallArr(i) = "DX" Or SmallArr(i) = "YD" Or SmallArr(i) = "LT" Or SmallArr(i) = "JX" Or SmallArr(i) = "EX" Or SmallArr(i) = "DS" Or SmallArr(i) = "BZ" Then
            BigArr(Count) = "TXD"
            Count = Count + 1
            ReDim Preserve BigArr(Count)
        ElseIf SmallArr(i) = "JS" Or SmallArr(i) = "XF" Then
            BigArr(Count) = "JSD"
            Count = Count + 1
            ReDim Preserve BigArr(Count)
        ElseIf SmallArr(i) = "PS" Or SmallArr(i) = "YS" Or SmallArr(i) = "WS" Or SmallArr(i) = "FS" Then
            BigArr(Count) = "PSD"
            Count = Count + 1
            ReDim Preserve BigArr(Count)
        ElseIf SmallArr(i) = "RQ" Or SmallArr(i) = "MQ" Or SmallArr(i) = "TR" Or SmallArr(i) = "YH" Then
            BigArr(Count) = "RQD"
            Count = Count + 1
            ReDim Preserve BigArr(Count)
        ElseIf SmallArr(i) = "RL" Or SmallArr(i) = "RS" Or SmallArr(i) = "ZQ" Then
            BigArr(Count) = "RLD"
            Count = Count + 1
            ReDim Preserve BigArr(Count)
        ElseIf SmallArr(i) = "SY" Or SmallArr(i) = "GS" Then
            BigArr(Count) = "GYD"
            Count = Count + 1
            ReDim Preserve BigArr(Count)
        End If
    Next 'i
    DelCF BigArr,LayerStr,LayerCount
    LayArr = Split(LayerStr,",", - 1,1)
    For i = 0 To UBound(LayArr)
        LayArr(i) = ToChinese(LayArr(i))
    Next 'i
End Function' GetAllLayerName

'ȥ���ַ������ظ�ֵ
Function DelCF(ByVal StrArr(),ByRef ToTalVal,ByRef LxCount)
    ToTalVal = ""
    For i = 0 To UBound(StrArr) - 1
        If ToTalVal = "" Then
            ToTalVal = "'" & StrArr(i) & "'"
        ElseIf Replace(ToTalVal,StrArr(i),"") = ToTalVal Then
            ToTalVal = ToTalVal & "," & "'" & StrArr(i) & "'"
        End If
    Next 'i
    ToTalVal = Replace(ToTalVal,"'","")
    LxCount = UBound(Split(ToTalVal,",", - 1,1)) + 1
End Function' DelCF

'==================================================CAD���======================================================================

Function SZDWT(ByVal TkId,ByVal FilePath)
    SSProcess.SetFeatureCodeTB "FeatureCodeTB_500", "SymbolScriptTB_500"
    SSProcess.SetNotetemplateTB "NoteTemplateTB_500"
    
    SSProcess.ClearDataXParameter
    SSProcess.SetDataXParameter "DataType", "1"      '���ݸ�ʽ��ʽ��0(ArcGIS SDE)�� 1(DWG)��2(DXF)�� 3(E00)�� 4(Coverage)�� 5(Shp)
    SSProcess.SetDataXParameter "Version", "2008"    'AutoCad���ݰ汾�š�2000,2004,2006
    SSProcess.SetDataXParameter "FeatureCodeTBName", "FeatureCodeTB_500"
    SSProcess.SetDataXParameter "SymbolScriptTBName", "SymbolScriptTB_500"
    SSProcess.SetDataXParameter "NoteTemplateTBName", "NoteTemplateTB_500"
    SSProcess.SetDataXParameter "ExportPathName", FilePath                    '����ļ���(����·����),���Ϊ��ʱ,���Զ������Ի���ѡ��
    SSProcess.SetDataXParameter "DataBoundMode", "2"                    '���������Χ��ʽ�� 0(��������)�� 1(ѡ������)�� 2(��ǰͼ��)��
    SSProcess.SetDataXParameter "ZeroLineWidth", "10"
    SSProcess.SetDataXParameter "AcadColorMethod", "0"
    SSProcess.SetDataXParameter "ExportLayerCount", "0"
    SSProcess.SetDataXParameter "ColorUseStatus", "1"       '��ɫʹ��״̬��0����������趨��ɫ�������1���������趨��ɫ�����
    SSProcess.SetDataXParameter "ExplodeObjColorStatus", "1"
    SSProcess.SetDataXParameter "FontWidthScale", "0.7"            '���ע���ֿ����ű�
    SSProcess.SetDataXParameter "FontHeightScale", "0.7"        '���ע���ָ����ű�  
    SSProcess.SetDataXParameter "FontSizeUseStatus","1"               '�����Сʹ��״̬ 0 ����ע�Ƿ���������ָ߿�������� 1 ����ע�������ָ߿������
    SSProcess.SetDataXParameter "OthersExportMode", "3"'���AutoCAD����ʱ����������ʽ�� 0��������룩�� 1��������еĺ�ȣ��� 2��������еı�������3���ó�0��
    SSProcess.SetDataXParameter "OthersExportToZFactor", "1"       '���AutoCAD����ʱ������������Z������ʽ�� 0����������� 1�������
    SSProcess.SetDataXParameter "ExplodeNoteStatus","0"
    SSProcess.SetDataXParameter "SymbolExplodeMode", "1"   '���Ŵ�ɢ��ʽ�� 0���Զ���ɢ���� 1�����ݱ�����趨��ɢ���� 2��ȫ������ɢ��
    SSProcess.SetDataXParameter "LayerUseStatus", "1"     '�����������ʹ��״̬��0����������趨�����������1���������趨���������
    SSProcess.SetDataXParameter "ExplodeObjLayerStatus", "0"  '��Ƕ����ͼ�������ʽ��0�������������趨������� 1����������ͬ�������
    SSProcess.SetDataXParameter "LineExportMode", "1" '���AutoCAD����ʱ�������������ʽ�� 0 ��ȱʡ��ʽ������ͬ�߳�ʱ��3DPolyline��������ఴ2DPolyline������� 1��ǿ�ư�2DPolyline������� 2�� ǿ�ư�3DPolyline����� 3�� ǿ�ư�Polyline�����
    SSProcess.SetDataXParameter "LineWidthUseStatus", "0"
    SSProcess.SetDataXParameter "GotoPointsMode", "1"                     '���ͼ�����߻���ʽ�� 0 �������߻����� 1 ��ֻ���߻����ߣ��� 2 ������ͼ�����߻���
    SSProcess.SetDataXParameter "AcadLineWidthMode", "3"
    SSProcess.SetDataXParameter "AcadLineScaleMode", "0"                'Acad���ͱ��������ʽ��0 ������߳�������� 1 ���ǰ�1���
    SSProcess.SetDataXParameter "AcadLineWeightMode","0"               'Acad���������ʽ��0 �����߿� 1 ��� 2 ��� 3 ���߶���
    SSProcess.SetDataXParameter "AcadBlockUseColorMode", "1"        'Acadͼ�������ɫʹ�÷�ʽ��0 ��� 1 ��� 2 �����ʵ��
    SSProcess.SetDataXParameter "AcadLinetypeGenerateMode", "1"
    SSProcess.SetDataXParameter "ExplodeObjMakeGroup ", "0"       'AutoCAD����ʱ����ɢ������������ʽ�� 0�������飩�� 1�� ���飬ͬʱҪ��FeatureCodeTB���е�ExtraInfo=1 
    SSProcess.SetDataXParameter "AcadUsePersonalBlockScaleCodes ", "1=7601023"       'AcadUsePersonalBlockScaleCodes ָ��ʹ�����������ı��롣��ʽ1�� ����1=����1,����2;����2=����1,����2��ʽ2�� ���� (�÷�ʽָ�����б����ʹ��ָ���Ŀ����) 
    SSProcess.SetDataXParameter "AcadDwtFileName", SSProcess.GetSysPathName (0) & "\Acadlin\acad.dwt"
    
    startIndex = 0
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"DEFAULT"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���ӷ���������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ע��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ά��ͼ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ͼ����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�������Ե�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�����Ե�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���ؽ�ַ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ͼ����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�������Ƶ�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������Ƶ�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���Ƶ�����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ѧ����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ͼ����Χ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ˮϵ����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ˮϵ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ˮϵ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ˮϵ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ˮϵ������ʩ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ˮϵ������ʩ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ˮϵ������ʩ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ַ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"¥ַ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"POI"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ص�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ظ�����ʩ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ظ�����ʩ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ظ�����ʩ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��·��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ͨ������ʩ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ͨ������ʩ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ͨ������ʩ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��·��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��·��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��·����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ͨ������ʩ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��·�������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���ߵ�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ȸ���"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�̵߳�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ò��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ò��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ò��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ֲ����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ֲ�������ʵ�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ֲ����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�����ע��Ϣ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������ͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ͼ��Χ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ע��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ͼ��ͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���ش�ע��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���ؽ�ַ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ؼ���"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ؼ�����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���ۿ��Ƶ�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����Ȩ�ڵ�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"GPS����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���۲�վ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ʵ���վ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"֧����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ʹ��Ȩ�ڵ�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�������õط�Χ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ڵؽ�ַ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ڵؽ�ַ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ʵ����Ƶ�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ں�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ں���ַ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ں���ַ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���۷�����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ʵ�������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ڵ�ͼ����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���ע��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��Ȼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��״������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��״������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��״������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"¥��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ǽ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ǽ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ֻ���ǽ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ּ���ǽ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ͼ����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������Ϣ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�滮��������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ƫ���"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�滮������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�滮�����ﷶΧ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�����ﷶΧ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��������׷�Χ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������Ĥ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�滮Χǽ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���������Ϣ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������ע��Ϣ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ͼ��Χ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�滮�����ɹ�ͼͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"λ��ͼͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ռ�����������ͼͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�����߶ȼ���߲�����ͼͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�����Ǹ���"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������ͣ��λ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ǻ�����ͣ��λ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ͣ��λ�ֲ�ͼͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�̵ط�Χ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�̵ؿ���ͼͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������Ϣ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"Ժ�������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"TERP"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"GTFA"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"GTFL"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ط�����ԭ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ͼ�߸�����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���Ҳ�һ��ͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��һ��ͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ط���һ��ͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"Ȩ������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������Ŀ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ڵ�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ͼ������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ƽ��ͼͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�����Ա�ͼͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���غ���ͼͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ͣ��λ��Χ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ͼ����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"KZ_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"KZ_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"SX_G"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"SX_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"SX_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"SX_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"JMD_G"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"JMD_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"JMD_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"JMD_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"JT_G"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"JT_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"JT_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"JT_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"GX_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"GX_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"GX_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"JJ_G"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"JJ_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"JJ_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"JJ_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"DM_G"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"DM_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"DM_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"DM_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ZB_G"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ZB_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ZB_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ZB_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"QT"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"TK"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"PSANNEXE"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"FSANNEXE"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"YSANNEXE"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"WSANNEXE"
    
    
    
    LayStr = "CD,CT,CY,CQ,CS,QT,BM,FQ,DL,GD,LD,DC,XH,TX,DX,YD,LT,JX,JK,EX,DS,BZ,JS,XF,PS,YS,WS,FS,RQ,MQ,TR,YH,RL,RS,ZQ,SY,GS"
    LayArr = Split(LayStr,",", - 1,1)
    For i = 0 To UBound(LayArr)
        SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),LayArr(i)
    Next 'i
    
    For i = 0 To UBound(LayArr)
        SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),LayArr(i) & "��ע��"
        SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),LayArr(i) & "ע��"
    Next 'i
    
    startIndex = 0
    SSProcess.SetDataXParameter "LayerRelationCount", "100"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CD:CDPOINT:CDLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CT:CTPOINT:CTLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CY:CYPOINT:CYLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CQ:CQPOINT:CQLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CS:CSPOINT:CSLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "QT:QTPOINT:QTLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "BM:BMPOINT:BMLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "FQ:FQPOINT:FQLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DL:DLPOINT:DLLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "GD:GDPOINT:GDLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "LD:LDPOINT:LDLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DC:DCPOINT:DCLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "XH:XHPOINT:XHLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "TX:TXPOINT:TXLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DX:DXPOINT:DXLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YD:DYPOINT:YDLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "LT:LTPOINT:LTLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JX:JXPOINT:JXLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JK:JKPOINT:JKLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "EX:EXPOINT:EXLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DS:DSPOINT:DSLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "BZ:BZPOINT:BZLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JS:JSPOINT:JSLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "XF:XFPOINT:XFLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "PS:PSPOINT:PSLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YS:YSPOINT:YSLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "WS:WSPOINT:WSLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "FS:FSPOINT:FSLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RQ:RQPOINT:RQLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "MQ:MQPOINT:MQLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "TR:TRPOINT:TRLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YH:YHPOINT:YHLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RL:RLPOINT:RLLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RS:RSPOINT:RSLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "ZQ:ZQPOINT:ZQLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "SY:SYPOINT:SYLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "GS:GSPOINT:GSLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "����ͼ����:TK:TK:TK:TK:TK"
    
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CD��ע��::::CDTEXT:CDTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CT��ע��::::CTTEXT:CTTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CY��ע��::::CYTEXT:CYTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CQ��ע��::::CQTEXT:CQTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CS��ע��::::CSTEXT:CSTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "QT��ע��::::QTTEXT:QTTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "BM��ע��::::BMTEXT:BMTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "FQ��ע��::::FQTEXT:FQTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DL��ע��::::DLTEXT:DLTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "GD��ע��::::GDTEXT:GDTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "LD��ע��::::LDTEXT:LDTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DC��ע��::::DCTEXT:DCTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "XH��ע��::::XHTEXT:XHTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "TX��ע��::::TXTEXT:TXTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DX��ע��::::DXTEXT:DXTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YD��ע��::::YDTEXT:YDTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "LT��ע��::::LTTEXT:LTTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JX��ע��::::JXTEXT:JXTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JK��ע��::::JKTEXT:JKTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "EX��ע��::::EXTEXT:EXTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DS��ע��::::DSTEXT:DSTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "BZ��ע��::::BZTEXT:BZTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JS��ע��::::JSTEXT:JSTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "XF��ע��::::XFTEXT:XFTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "PS��ע��::::PSTEXT:PSTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YS��ע��::::YSTEXT:YSTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "WS��ע��::::WSTEXT:WSTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "FS��ע��::::FSTEXT:FSTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RQ��ע��::::RQTEXT:RQTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "MQ��ע��::::MQTEXT:MQTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "TR��ע��::::TRTEXT:TRTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YH��ע��::::YHTEXT:YHTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RL��ע��::::RLTEXT:RLTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RS��ע��::::RSTEXT:RSTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "ZQ��ע��::::ZQTEXT:ZQTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "SY��ע��::::SYTEXT:SYTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "GS��ע��::::GSTEXT:GSTEXT"
    
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CDע��::::CDMARK:CDMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CTע��::::CTMARK:CTMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CYע��::::CYMARK:CYMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CQע��::::CQMARK:CQMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CSע��::::CSMARK:CSMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "QTע��::::QTMARK:QTMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "BMע��::::BMMARK:BMMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "FQע��::::FQMARK:FQMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DLע��::::DLMARK:DLMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "GDע��::::GDMARK:GDMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "LDע��::::LDMARK:LDMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DCע��::::DCMARK:DCMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "XHע��::::XHMARK:XHMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "TXע��::::TXMARK:TXMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DXע��::::DXMARK:DXMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YDע��::::YDMARK:YDMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "LTע��::::LTMARK:LTMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JXע��::::JXMARK:JXMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JKע��::::JKMARK:JKMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "EXע��::::EXMARK:EXMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DSע��::::DSMARK:DSMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "BZע��::::BZMARK:BZMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JSע��::::JSMARK:JSMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "XFע��::::XFMARK:XFMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "PSע��::::PSMARK:PSMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YSע��::::YSMARK:YSMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "WSע��::::WSMARK:WSMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "FSע��::::FSMARK:FSMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RQע��::::RQMARK:RQMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "MQע��::::MQMARK:MQMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "TRע��::::TRMARK:TRMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YHע��::::YHMARK:YHMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RLע��::::RLMARK:RLMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RSע��::::RSMARK:RSMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "ZQע��::::ZQMARK:ZQMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "SYע��::::SYMARK:SYMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "GSע��::::GSMARK:GSMARK"
    startIndex = 0
    SSProcess.SetDataXParameter "TableFieldDefCount","3000"
    'QT���Ա��㣩
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QT,0,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QT,0,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    'QT���Ա��ߣ�
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QT,1,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QT,1,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '���Ƶ㣨�㣩
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"KZD,0,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"KZD,0,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '���Ƶ㣨�ߣ�
    'SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"SCP_LN,1,Z,Z,Z,Z,,dbDouble,16,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"KZD,1,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"KZD,1,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '���Ƶ�ע��
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"KZD,3,FontClass,FontClass,south:1000,Byname,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"KZD,3,Memo,Memo,NAME:1000,,,dbText,50,0"
    '�̵߳㣨�㣩
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"GCD,0,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"GCD,0,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '�̵߳㣨�ߣ�
    'SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"SCP_LN,1,Z,Z,Z,Z,,dbDouble,16,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"GCD,1,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"GCD,1,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '�̵߳�ע��
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"GCD,3,FontClass,FontClass,south:1000,Byname,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"GCD,3,Memo,Memo,NAME:1000,,,dbText,50,0"
    '�ȸ��ߣ��㣩
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DGX,0,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DGX,0,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '�ȸ��ߣ��ߣ�
    'SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"SCP_LN,1,Z,Z,Z,Z,,dbDouble,16,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DGX,1,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DGX,1,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DSX,1,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DSX,1,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '�ȸ��ߣ��棩
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DGX,2,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DGX,2,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '�ȸ���ע��
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DGX,3,FontClass,FontClass,south:1000,Byname,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DGX,3,Memo,Memo,NAME:1000,,,dbText,50,0"
    'ˮϵ��
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"SXSS,0,Code,Code,south:1000,Others,,dbText,20,0"
    'SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"SXSS,0,Code,Code,YSDM:1000,code,,dbText,100,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"SXSS,0,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    'ˮϵ��
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"SXSS,1,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"SXSS,1,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    'ˮϵ��
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"SXSS,2,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"SXSS,2,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    'ˮϵע��
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"SXSS,3,FontClass,FontClass,south:1000,Byname,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"SXSS,3,Memo,Memo,NAME:1000,,,dbText,50,0"
    '�������
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"JMD,2,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"JMD,2,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '����ص�
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"JMD,0,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"JMD,0,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '�������
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"JMD,1,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"JMD,1,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '�����ע��
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"JMD,3,FontClass,FontClass,south:1000,Byname,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"JMD,3,Memo,Memo,NAME:1000,,,dbText,50,0"
    '��ַ����
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"JZD,2,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"JZD,2,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '��ַ���
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"JZD,0,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"JZD,0,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '��ַ����
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"JZD,1,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"JZD,1,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '��ַ��ע��
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"JZD,3,FontClass,FontClass,south:1000,Byname,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"JZD,3,Memo,Memo,NAME:1000,,,dbText,50,0"
    '�����
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DLDW,0,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DLDW,0,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '�����
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DLDW,1,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DLDW,1,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '�����
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DLDW,2,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DLDW,2,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '�����
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DLDW,3,FontClass,FontClass,south:1000,Byname,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DLDW,3,Memo,Memo,NAME:1000,,,dbText,50,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DLSS,0,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DLSS,0,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '��ͨ��
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DLSS,1,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DLSS,1,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '��ͨ��
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DLSS,2,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DLSS,2,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '��ͨע��
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DLSS,3,FontClass,FontClass,south:1000,Byname,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DLSS,3,Memo,Memo,NAME:1000,,,dbText,50,0"
    '���ߵ�
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"GXYZ,0,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"GXYZ,0,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '������
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"GXYZ,1,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"GXYZ,1,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '������
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"GXYZ,2,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"GXYZ,2,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '����ע��
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"GXYZ,3,FontClass,FontClass,south:1000,Byname,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"GXYZ,3,Memo,Memo,NAME:1000,,,dbText,50,0"
    '������������
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"JJ,0,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"JJ,0,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '������������
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"JJ,1,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"JJ,1,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '������������
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"JJ,2,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"JJ,2,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '����ע��
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"JJ,3,FontClass,FontClass,south:1000,Byname,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"JJ,3,Memo,Memo,NAME:1000,,,dbText,50,0"
    '��ò��
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DMTZ,0,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DMTZ,0,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '��ò��
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DMTZ,1,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DMTZ,1,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '��ò��
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DMTZ,2,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DMTZ,2,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    '��òע��
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DMTZ,3,FontClass,FontClass,south:1000,Byname,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"DMTZ,3,Memo,Memo,NAME:1000,,,dbText,50,0"
    'ֲ�������ʵ�
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"ZBTZ,0,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"ZBTZ,0,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    'ֲ����������
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"ZBTZ,1,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"ZBTZ,1,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    'ֲ����������
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"ZBTZ,2,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"ZBTZ,2,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    'ֲ��������ע��
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"ZBTZ,3,FontClass,FontClass,south:1000,Byname,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"ZBTZ,3,Memo,Memo,NAME:1000,,,dbText,50,0"
    'ע��
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"TK,3,FontClass,FontClass,south:1000,Byname,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"TK,3,Memo,Memo,NAME:1000,,,dbText,50,0"
    '�Ǽ���
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"ASSIST,1,Code,Code,south:1000,Others,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"ASSIST,1,ObjectName,ObjectName,NAME:1000,,,dbText,50,0"
    SSProcess.ExportData
    SSProcess.SetFeatureCodeTB "FeatureCodeTB_500", "SymbolScriptTB_500"
    SSProcess.SetNotetemplateTB "NoteTemplateTB_500"
End Function

'�����Զ�����
Function AddOne(ByRef StartIndex)
    StartIndex = StartIndex + 1
    AddOne = StartIndex
End Function

'����ת��Ϊ����
Function ToChinese(ByVal EngLayerName) 'EngLayerName ͼ������(Ӣ��)
    EngStr = "CDD,CXD,CYD,CSD,DLD,TXD,JSD,PSD,RQD,RLD,GYD"
    CheStr = "�������,����ͨ��,��������ˮ,���й���,����,ͨ��,��ˮ,��ˮ,ȼ��,����,��ҵ"
    EngArr = Split(EngStr,",", - 1,1)
    CheArr = Split(CheStr,",", - 1,1)
    ToChinese = ""
    For j = 0 To UBound(EngArr)
        If EngArr(j) = EngLayerName Then
            ToChinese = CheArr(j)
        End If
    Next 'j
End Function' ToChinese

'�ر�����ͼ��
Function AllDisVisible()
    count = SSProcess.GetLayerCount
    For i = 0 To count - 1
        layername = SSProcess.GetLayerName (i)
        SSProcess.SetLayerStatus layername, 0, 1
    Next
    SSProcess.RefreshView
End Function

'����ͼ����ͼ��
Function CreatMap(ByVal Way,ByRef ContinueIf)
    SSProcess.CreateMapFrame
    SSProcess.MapMethod "LoadData","ͼ����"
    FrameCount = SSProcess.GetMapFrameCount()
    For i = 0 To FrameCount - 1
        SSProcess.GetMapFrameCenterPoint i, CenterX, CenterY
        SSProcess.SetFrameCode("59999999")
        SSProcess.SetCurMapFrame CenterX, CenterY, 0, ""
        'CreateNote SSProcess.GetCurMapFrame()
        'GetAllLayerName SSProcess.GetCurMapFrame(),SmallArr,LayArr
        ModifyAttr "59999999",Way,SSProcess.GetCurMapFrame(),XmMc,ContinueIf
        FilePath = SSProcess.GetSysPathName(5) & "רҵ����ͼ\" & SSProcess.GetObjectAttr(SSProcess.GetCurMapFrame(),"[MapNumber]") & "�ۺϹ���ͼ.dwg"
        SZDWT SSProcess.GetCurMapFrame(),FilePath
        DelTk
    Next
    SSProcess.FreeMapFrame
End Function

'�ֲ����
Function FCExport(ByVal Way)
    SSProcess.CreateMapFrame
    SSProcess.MapMethod "LoadData","ͼ����"
    FrameCount = SSProcess.GetMapFrameCount()
    For i = 0 To FrameCount - 1
        SSProcess.GetMapFrameCenterPoint i, CenterX, CenterY
        SSProcess.SetFrameCode("59999999")
        SSProcess.SetCurMapFrame CenterX, CenterY, 0, ""
        'CreateNote SSProcess.GetCurMapFrame()
        GetAllLayerName SSProcess.GetCurMapFrame(),SmallArr,BigArr
        For k = 0 To UBound(BigArr)
            Select Case BigArr(k)
                Case "�������"
                AllDisVisible
                
                SSProcess.SetLayerStatus "CD", 1, 1
                SSProcess.SetLayerStatus "CDע��", 1, 1
                SSProcess.SetLayerStatus "CD��ע��", 1, 1
                SSProcess.SetLayerStatus "TK", 1, 1
                SSProcess.SetLayerStatus "����ͼ����", 1, 1
                SSProcess.SetLayerStatus "ͼ����", 1, 1
                FilePath = SSProcess.GetSysPathName(5) & "רҵ����ͼ\" & XmMc & BigArr(k) & SSProcess.GetObjectAttr(SSProcess.GetCurMapFrame(),"[MapNumber]") & "���¹���ͼ.dwg"
                SetFcAttr "59999999",SSProcess.GetCurMapFrame(),XmMc,BigArr(k)
                SZDWT TkId,FilePath
                Case "����ͨ��"
                AllDisVisible
                SetFcAttr "59999999",SSProcess.GetCurMapFrame(),XmMc,BigArr(k)
                SSProcess.SetLayerStatus "CT", 1, 1
                SSProcess.SetLayerStatus "TK", 1, 1
                SSProcess.SetLayerStatus "CTע��", 1, 1
                SSProcess.SetLayerStatus "CT��ע��", 1, 1
                SSProcess.SetLayerStatus "����ͼ����", 1, 1
                SSProcess.SetLayerStatus "ͼ����", 1, 1
                
                FilePath = SSProcess.GetSysPathName(5) & "רҵ����ͼ\" & XmMc & BigArr(k) & SSProcess.GetObjectAttr(SSProcess.GetCurMapFrame(),"[MapNumber]") & "���¹���ͼ.dwg"
                
                SZDWT TkId,FilePath
                Case "��������ˮ"
                AllDisVisible
                SetFcAttr "59999999",SSProcess.GetCurMapFrame(),XmMc,BigArr(k)
                SSProcess.SetLayerStatus "CY", 1, 1
                SSProcess.SetLayerStatus "CQ", 1, 1
                SSProcess.SetLayerStatus "CS", 1, 1
                SSProcess.SetLayerStatus "QT", 1, 1
                SSProcess.SetLayerStatus "TK", 1, 1
                SSProcess.SetLayerStatus "CYע��", 1, 1
                SSProcess.SetLayerStatus "CQע��", 1, 1
                SSProcess.SetLayerStatus "CSע��", 1, 1
                SSProcess.SetLayerStatus "QTע��", 1, 1
                SSProcess.SetLayerStatus "CY��ע��", 1, 1
                SSProcess.SetLayerStatus "CQ��ע��", 1, 1
                SSProcess.SetLayerStatus "CS��ע��", 1, 1
                SSProcess.SetLayerStatus "QT��ע��", 1, 1
                SSProcess.SetLayerStatus "����ͼ����", 1, 1
                SSProcess.SetLayerStatus "ͼ����", 1, 1
                
                FilePath = SSProcess.GetSysPathName(5) & "רҵ����ͼ\" & XmMc & BigArr(k) & SSProcess.GetObjectAttr(SSProcess.GetCurMapFrame(),"[MapNumber]") & "���¹���ͼ.dwg"
                
                SZDWT TkId,FilePath
                Case "���й���"
                AllDisVisible
                SetFcAttr "59999999",SSProcess.GetCurMapFrame(),XmMc,BigArr(k)
                SSProcess.SetLayerStatus "BM", 1, 1
                SSProcess.SetLayerStatus "FQ", 1, 1
                SSProcess.SetLayerStatus "BMע��", 1, 1
                SSProcess.SetLayerStatus "FQע��", 1, 1
                SSProcess.SetLayerStatus "BM��ע��", 1, 1
                SSProcess.SetLayerStatus "FQ��ע��", 1, 1
                SSProcess.SetLayerStatus "TK", 1, 1
                SSProcess.SetLayerStatus "����ͼ����", 1, 1
                SSProcess.SetLayerStatus "ͼ����", 1, 1
                
                FilePath = SSProcess.GetSysPathName(5) & "רҵ����ͼ\" & XmMc & BigArr(k) & SSProcess.GetObjectAttr(SSProcess.GetCurMapFrame(),"[MapNumber]") & "���¹���ͼ.dwg"
                
                SZDWT TkId,FilePath
                Case "����"
                AllDisVisible
                SSProcess.SetLayerStatus "XH", 1, 1
                SSProcess.SetLayerStatus "DC", 1, 1
                SSProcess.SetLayerStatus "LD", 1, 1
                SSProcess.SetLayerStatus "GD", 1, 1
                SSProcess.SetLayerStatus "DL", 1, 1
                SSProcess.SetLayerStatus "XHע��", 1, 1
                SSProcess.SetLayerStatus "DCע��", 1, 1
                SSProcess.SetLayerStatus "LDע��", 1, 1
                SSProcess.SetLayerStatus "GDע��", 1, 1
                SSProcess.SetLayerStatus "DLע��", 1, 1
                SSProcess.SetLayerStatus "XH��ע��", 1, 1
                SSProcess.SetLayerStatus "DC��ע��", 1, 1
                SSProcess.SetLayerStatus "LD��ע��", 1, 1
                SSProcess.SetLayerStatus "GD��ע��", 1, 1
                SSProcess.SetLayerStatus "DL��ע��", 1, 1
                SSProcess.SetLayerStatus "TK", 1, 1
                SSProcess.SetLayerStatus "����ͼ����", 1, 1
                SSProcess.SetLayerStatus "ͼ����", 1, 1
                
                SetFcAttr "59999999",SSProcess.GetCurMapFrame(),XmMc,BigArr(k)
                FilePath = SSProcess.GetSysPathName(5) & "רҵ����ͼ\" & XmMc & BigArr(k) & SSProcess.GetObjectAttr(SSProcess.GetCurMapFrame(),"[MapNumber]") & "���¹���ͼ.dwg"
                
                SZDWT TkId,FilePath
                Case "ͨ��"
                AllDisVisible
                SSProcess.SetLayerStatus "BZ", 1, 1
                SSProcess.SetLayerStatus "DX", 1, 1
                SSProcess.SetLayerStatus "YD", 1, 1
                SSProcess.SetLayerStatus "LT", 1, 1
                SSProcess.SetLayerStatus "JX", 1, 1
                SSProcess.SetLayerStatus "JK", 1, 1
                SSProcess.SetLayerStatus "EX", 1, 1
                SSProcess.SetLayerStatus "DS", 1, 1
                SSProcess.SetLayerStatus "TX", 1, 1
                SSProcess.SetLayerStatus "BZע��", 1, 1
                SSProcess.SetLayerStatus "DXע��", 1, 1
                SSProcess.SetLayerStatus "YDע��", 1, 1
                SSProcess.SetLayerStatus "LTע��", 1, 1
                SSProcess.SetLayerStatus "JXע��", 1, 1
                SSProcess.SetLayerStatus "JKע��", 1, 1
                SSProcess.SetLayerStatus "EXע��", 1, 1
                SSProcess.SetLayerStatus "DSע��", 1, 1
                SSProcess.SetLayerStatus "TXע��", 1, 1
                SSProcess.SetLayerStatus "BZ��ע��", 1, 1
                SSProcess.SetLayerStatus "DX��ע��", 1, 1
                SSProcess.SetLayerStatus "YD��ע��", 1, 1
                SSProcess.SetLayerStatus "LT��ע��", 1, 1
                SSProcess.SetLayerStatus "JX��ע��", 1, 1
                SSProcess.SetLayerStatus "JK��ע��", 1, 1
                SSProcess.SetLayerStatus "EX��ע��", 1, 1
                SSProcess.SetLayerStatus "DS��ע��", 1, 1
                SSProcess.SetLayerStatus "TX��ע��", 1, 1
                SSProcess.SetLayerStatus "TK", 1, 1
                SSProcess.SetLayerStatus "����ͼ����", 1, 1
                SSProcess.SetLayerStatus "ͼ����", 1, 1
                SetFcAttr "59999999",SSProcess.GetCurMapFrame(),XmMc,BigArr(k)
                FilePath = SSProcess.GetSysPathName(5) & "רҵ����ͼ\" & XmMc & BigArr(k) & SSProcess.GetObjectAttr(SSProcess.GetCurMapFrame(),"[MapNumber]") & "���¹���ͼ.dwg"
                SZDWT TkId,FilePath
                Case "��ˮ"
                AllDisVisible
                SSProcess.SetLayerStatus "JS", 1, 1
                SSProcess.SetLayerStatus "XF", 1, 1
                SSProcess.SetLayerStatus "JSע��", 1, 1
                SSProcess.SetLayerStatus "XFע��", 1, 1
                SSProcess.SetLayerStatus "JS��ע��", 1, 1
                SSProcess.SetLayerStatus "XF��ע��", 1, 1
                SSProcess.SetLayerStatus "TK", 1, 1
                SSProcess.SetLayerStatus "����ͼ����", 1, 1
                SSProcess.SetLayerStatus "ͼ����", 1, 1
                SetFcAttr "59999999",SSProcess.GetCurMapFrame(),XmMc,BigArr(k)
                FilePath = SSProcess.GetSysPathName(5) & "רҵ����ͼ\" & XmMc & BigArr(k) & SSProcess.GetObjectAttr(SSProcess.GetCurMapFrame(),"[MapNumber]") & "���¹���ͼ.dwg"
                SZDWT TkId,FilePath
                Case "��ˮ"
                AllDisVisible
                SSProcess.SetLayerStatus "FS", 1, 1
                SSProcess.SetLayerStatus "WS", 1, 1
                SSProcess.SetLayerStatus "YS", 1, 1
                SSProcess.SetLayerStatus "PS", 1, 1
                SSProcess.SetLayerStatus "FSע��", 1, 1
                SSProcess.SetLayerStatus "WSע��", 1, 1
                SSProcess.SetLayerStatus "YSע��", 1, 1
                SSProcess.SetLayerStatus "PSע��", 1, 1
                SSProcess.SetLayerStatus "FS��ע��", 1, 1
                SSProcess.SetLayerStatus "WS��ע��", 1, 1
                SSProcess.SetLayerStatus "YS��ע��", 1, 1
                SSProcess.SetLayerStatus "PS��ע��", 1, 1
                SSProcess.SetLayerStatus "TK", 1, 1
                SSProcess.SetLayerStatus "����ͼ����", 1, 1
                SSProcess.SetLayerStatus "ͼ����", 1, 1
                SetFcAttr "59999999",SSProcess.GetCurMapFrame(),XmMc,BigArr(k)
                FilePath = SSProcess.GetSysPathName(5) & "רҵ����ͼ\" & XmMc & BigArr(k) & SSProcess.GetObjectAttr(SSProcess.GetCurMapFrame(),"[MapNumber]") & "���¹���ͼ.dwg"
                SZDWT TkId,FilePath
                Case "ȼ��"
                AllDisVisible
                SSProcess.SetLayerStatus "YH", 1, 1
                SSProcess.SetLayerStatus "MQ", 1, 1
                SSProcess.SetLayerStatus "TR", 1, 1
                SSProcess.SetLayerStatus "RQ", 1, 1
                SSProcess.SetLayerStatus "YHע��", 1, 1
                SSProcess.SetLayerStatus "MQע��", 1, 1
                SSProcess.SetLayerStatus "TRע��", 1, 1
                SSProcess.SetLayerStatus "RQע��", 1, 1
                SSProcess.SetLayerStatus "YH��ע��", 1, 1
                SSProcess.SetLayerStatus "MQ��ע��", 1, 1
                SSProcess.SetLayerStatus "TR��ע��", 1, 1
                SSProcess.SetLayerStatus "RQ��ע��", 1, 1
                SSProcess.SetLayerStatus "TK", 1, 1
                SSProcess.SetLayerStatus "����ͼ����", 1, 1
                SSProcess.SetLayerStatus "ͼ����", 1, 1
                SetFcAttr "59999999",SSProcess.GetCurMapFrame(),XmMc,BigArr(k)
                FilePath = SSProcess.GetSysPathName(5) & "רҵ����ͼ\" & XmMc & BigArr(k) & SSProcess.GetObjectAttr(SSProcess.GetCurMapFrame(),"[MapNumber]") & "���¹���ͼ.dwg"
                SZDWT TkId,FilePath
                Case "����"
                AllDisVisible
                SSProcess.SetLayerStatus "ZQ", 1, 1
                SSProcess.SetLayerStatus "RL", 1, 1
                SSProcess.SetLayerStatus "RS", 1, 1
                SSProcess.SetLayerStatus "ZQע��", 1, 1
                SSProcess.SetLayerStatus "RLע��", 1, 1
                SSProcess.SetLayerStatus "RSע��", 1, 1
                SSProcess.SetLayerStatus "ZQ��ע��", 1, 1
                SSProcess.SetLayerStatus "RL��ע��", 1, 1
                SSProcess.SetLayerStatus "RS��ע��", 1, 1
                SSProcess.SetLayerStatus "TK", 1, 1
                SSProcess.SetLayerStatus "����ͼ����", 1, 1
                SSProcess.SetLayerStatus "ͼ����", 1, 1
                SetFcAttr "59999999",SSProcess.GetCurMapFrame(),XmMc,BigArr(k)
                FilePath = SSProcess.GetSysPathName(5) & "רҵ����ͼ\" & XmMc & BigArr(k) & SSProcess.GetObjectAttr(SSProcess.GetCurMapFrame(),"[MapNumber]") & "���¹���ͼ.dwg"
                SZDWT TkId,FilePath
                Case "��ҵ"
                AllDisVisible
                SSProcess.SetLayerStatus "GS", 1, 1
                SSProcess.SetLayerStatus "SY", 1, 1
                SSProcess.SetLayerStatus "GSע��", 1, 1
                SSProcess.SetLayerStatus "SYע��", 1, 1
                SSProcess.SetLayerStatus "GS��ע��", 1, 1
                SSProcess.SetLayerStatus "SY��ע��", 1, 1
                SSProcess.SetLayerStatus "TK", 1, 1
                SSProcess.SetLayerStatus "����ͼ����", 1, 1
                SSProcess.SetLayerStatus "ͼ����", 1, 1
                SetFcAttr "59999999",SSProcess.GetCurMapFrame(),XmMc,BigArr(k)
                FilePath = SSProcess.GetSysPathName(5) & "רҵ����ͼ\" & XmMc & BigArr(k) & SSProcess.GetObjectAttr(SSProcess.GetCurMapFrame(),"[MapNumber]") & "���¹���ͼ.dwg"
                SZDWT TkId,FilePath
            End Select
        Next 'z
        DelTk
    Next
    SSProcess.FreeMapFrame
End Function' FCExport

Function CreateNote(ByVal MapId)
    
    SSProcess.GetObjectPoint MapId, 2, StandX, StandY, StandZ, PointType, Name '���Ͻǵ�����ֵ
    
    BorderStartX = StandX - 10 - 20
    BorderStartY = StandY - 10
    BorderEndX = StandX - 14
    FeatureY = BorderStartY - 2 - 2
    
    SelAll MapId,CodeVal,CodeCount
    
    If CodeCount > 0 Then
        CodeArr = Split(CodeVal,",", - 1,1)
        For j = 0 To CodeCount - 1
            If SSProcess.GetFeatureCodeInfo(CodeArr(j),"Type") = 0 Then
                DrawPoint BorderStartX + 3.5,FeatureY,CodeArr(j)
                FeatureY = FeatureY - 2.25
            Else
                DrawLine BorderStartX + 2,BorderStartX + 5,FeatureY,CodeArr(j)
                FeatureY = FeatureY - 2.25
            End If
        Next 'j
    End If
    
    DrawBorder BorderStartX,BorderEndX,BorderStartY,FeatureY
    
End Function' CreateNote

'��ȡ���еĵ����Ҫ������
Function SelAll(ByVal OuterId,ByRef DisplayCode,ByRef CodeCount)
    PoiIds = SSProcess.SearchInPolyObjIDs(OuterId,0,"",0,1,1)
    LinIds = SSProcess.SearchInPolyObjIDs(OuterId,1,"",0,1,1)
    PoiArr = Split(PoiIds,",", - 1,1)
    LinArr = Split(LinIds,",", - 1,1)
    For i = 0 To UBound(PoiArr)
        Select Case SSProcess.GetObjectAttr(PoiArr(i),"SSObj_LayerName")
            Case "CD"
            If CDCodeStr = "" Then
                CDCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                CDCodeStr = CDCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            
            Case "CT"
            If CTCodeStr = "" Then
                CTCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                CTCodeStr = CTCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "CY"
            If CYCodeStr = "" Then
                CYCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                CYCodeStr = CYCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "CQ"
            If CQCodeStr = "" Then
                CQCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                CQCodeStr = CQCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "CS"
            If CSCodeStr = "" Then
                CSCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                CSCodeStr = CSCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "QT"
            If QTCodeStr = "" Then
                QTCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                QTCodeStr = QTCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "BM"
            If BMCodeStr = "" Then
                BMCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                BMCodeStr = BMCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "FQ"
            If FQCodeStr = "" Then
                FQCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                FQCodeStr = FQCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "DL"
            If DLCodeStr = "" Then
                DLCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                DLCodeStr = DLCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "GD"
            If GDCodeStr = "" Then
                GDCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                GDCodeStr = GDCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "LD"
            If LDCodeStr = "" Then
                LDCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                LDCodeStr = LDCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "DC"
            If DCCodeStr = "" Then
                DCCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                DCCodeStr = DCCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "XH"
            If XHCodeStr = "" Then
                XHCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                XHCodeStr = XHCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "TX"
            If TXCodeStr = "" Then
                TXCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                TXCodeStr = TXCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "DX"
            If DXCodeStr = "" Then
                DXCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                DXCodeStr = DXCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "YD"
            If YDCodeStr = "" Then
                YDCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                YDCodeStr = YDCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "LT"
            If LTCodeStr = "" Then
                LTCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                LTCodeStr = LTCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "JX"
            If JXCodeStr = "" Then
                JXCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                JXCodeStr = JXCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "JK"
            If JKCodeStr = "" Then
                JKCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                JKCodeStr = JKCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "DS"
            If DSCodeStr = "" Then
                DSCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                DSCodeStr = DSCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "BZ"
            If BZCodeStr = "" Then
                BZCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                BZCodeStr = BZCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "JS"
            If JSCodeStr = "" Then
                JSCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                JSCodeStr = JSCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "XF"
            If XFCodeStr = "" Then
                XFCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                XFCodeStr = XFCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "PS"
            If PSCodeStr = "" Then
                PSCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                PSCodeStr = PSCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "YS"
            If YSCodeStr = "" Then
                YSCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                YSCodeStr = YSCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "WS"
            If WSCodeStr = "" Then
                WSCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                WSCodeStr = WSCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "FS"
            If FSCodeStr = "" Then
                FSCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                FSCodeStr = FSCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "RQ"
            If RQCodeStr = "" Then
                RQCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                RQCodeStr = RQCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "MQ"
            If MQCodeStr = "" Then
                MQCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                MQCodeStr = MQCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "YH"
            If YHCodeStr = "" Then
                YHCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                YHCodeStr = YHCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "RL"
            If RLCodeStr = "" Then
                RLCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                RLCodeStr = RLCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "RS"
            If RSCodeStr = "" Then
                RSCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                RSCodeStr = RSCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "ZQ"
            If ZQCodeStr = "" Then
                ZQCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                ZQCodeStr = ZQCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "SY"
            If SYCodeStr = "" Then
                SYCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                SYCodeStr = SYCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "GS"
            If GSCodeStr = "" Then
                GSCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                GSCodeStr = GSCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "EX"
            If EXCodeStr = "" Then
                EXCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                EXCodeStr = EXCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "TR"
            If TRCodeStr = "" Then
                TRCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                TRCodeStr = TRCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
        End Select
    Next 'i
    
    For i = 0 To UBound(LinArr)
        Select Case SSProcess.GetObjectAttr(LinArr(i),"SSObj_LayerName")
            
            Case "CD"
            If CDCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    CDCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    CDCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    CDCodeStr = CDCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    CDCodeStr = CDCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "CT"
            If CTCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    CTCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    CTCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    CTCodeStr = CTCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    CTCodeStr = CTCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "CQ"
            If CQCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    CQCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    CQCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    CQCodeStr = CQCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    CQCodeStr = CQCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "CS"
            If CSCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    CSCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    CSCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    CSCodeStr = CSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    CSCodeStr = CSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "QT"
            If QTCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    QTCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    QTCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    QTCodeStr = QTCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    QTCodeStr = QTCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "BM"
            If BMCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    BMCodeStr = BMCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    BMCodeStr = BMCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    BMCodeStr = BMCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    BMCodeStr = BMCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "FQ"
            If FQCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    FQCodeStr = FQCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    FQCodeStr = FQCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    FQCodeStr = FQCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    FQCodeStr = FQCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "DL"
            If DLCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    DLCodeStr = DLCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    DLCodeStr = DLCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    DLCodeStr = DLCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    DLCodeStr = DLCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "GD"
            If GDCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    GDCodeStr = GDCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    GDCodeStr = GDCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    GDCodeStr = GDCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    GDCodeStr = GDCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "LD"
            If LDCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    LDCodeStr = LDCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    LDCodeStr = LDCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    LDCodeStr = LDCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    LDCodeStr = LDCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "DC"
            If DCCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    DCCodeStr = DCCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    DCCodeStr = DCCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    DCCodeStr = DCCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    DCCodeStr = DCCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "XH"
            If XHCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    XHCodeStr = XHCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    XHCodeStr = XHCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    XHCodeStr = XHCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    XHCodeStr = XHCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "TX"
            If TXCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    TXCodeStr = TXCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    TXCodeStr = TXCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    TXCodeStr = TXCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    TXCodeStr = TXCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "DX"
            If DXCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    DXCodeStr = DXCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    DXCodeStr = DXCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    DXCodeStr = DXCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    DXCodeStr = DXCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "YD"
            If YDCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    YDCodeStr = YDCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    YDCodeStr = YDCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    YDCodeStr = YDCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    YDCodeStr = YDCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "LT"
            If LTCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    LTCodeStr = LTCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    LTCodeStr = LTCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    LTCodeStr = LTCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    LTCodeStr = LTCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "JX"
            If JXCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    JXCodeStr = JXCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    JXCodeStr = JXCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    JXCodeStr = JXCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    JXCodeStr = JXCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "JK"
            If JKCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    JKCodeStr = JKCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    JKCodeStr = JKCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    JKCodeStr = JKCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    JKCodeStr = JKCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "DS"
            If DSCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    DSCodeStr = DSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    DSCodeStr = DSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    DSCodeStr = DSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    DSCodeStr = DSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "BZ"
            If BZCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    BZCodeStr = BZCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    BZCodeStr = BZCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    BZCodeStr = BZCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    BZCodeStr = BZCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "JS"
            If JSCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    JSCodeStr = JSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    JSCodeStr = JSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    JSCodeStr = JSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    JSCodeStr = JSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "XF"
            If XFCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    XFCodeStr = XFCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    XFCodeStr = XFCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    XFCodeStr = XFCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    XFCodeStr = XFCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "PS"
            If PSCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    PSCodeStr = PSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    PSCodeStr = PSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    PSCodeStr = PSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    PSCodeStr = PSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "YS"
            If YSCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    YSCodeStr = YSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    YSCodeStr = YSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    YSCodeStr = YSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    YSCodeStr = YSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "WS"
            If WSCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    WSCodeStr = WSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    WSCodeStr = WSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    WSCodeStr = WSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    WSCodeStr = WSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "FS"
            If FSCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    FSCodeStr = FSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    FSCodeStr = FSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    FSCodeStr = FSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    FSCodeStr = FSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "RQ"
            If RQCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    RQCodeStr = RQCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    RQCodeStr = RQCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    RQCodeStr = RQCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    RQCodeStr = RQCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "MQ"
            If MQCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    MQCodeStr = MQCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    MQCodeStr = MQCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    MQCodeStr = MQCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    MQCodeStr = MQCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "YH"
            If YHCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    YHCodeStr = YHCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    YHCodeStr = YHCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    YHCodeStr = YHCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    YHCodeStr = YHCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "RL"
            If RLCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    RLCodeStr = RLCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    RLCodeStr = RLCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    RLCodeStr = RLCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    RLCodeStr = RLCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "RS"
            If RSCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    RSCodeStr = RSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    RSCodeStr = RSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    RSCodeStr = RSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    RSCodeStr = RSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "ZQ"
            If ZQCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    ZQCodeStr = ZQCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    ZQCodeStr = ZQCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    ZQCodeStr = ZQCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    ZQCodeStr = ZQCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "SY"
            If SYCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    SYCodeStr = SYCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    SYCodeStr = SYCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    SYCodeStr = SYCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    SYCodeStr = SYCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "GS"
            If GSCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    GSCodeStr = GSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    GSCodeStr = GSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    GSCodeStr = GSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    GSCodeStr = GSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "EX"
            If EXCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    EXCodeStr = EXCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    EXCodeStr = EXCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    EXCodeStr = EXCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    EXCodeStr = EXCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
            Case "TR"
            If TRCodeStr = "" Then
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    TRCodeStr = TRCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    TRCodeStr = TRCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            Else
                If SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") = "�ǿ���" Then
                    TRCodeStr = TRCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                ElseIf SSProcess.GetSelGeoValue(LinArr(i),"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(LinArr(i),"[YYKS]") <> "0" Then
                    TRCodeStr = TRCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
                End If
            End If
            
        End Select
    Next 'i
    ' ReDim CodeStr(UBound(PoiArr) + UBound(LinArr))
    ' For i = 0 To UBound(PoiArr) + UBound(LinArr)
    '     If i <= UBound(PoiArr) Then
    '         CodeStr(i) = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")      
    '     Else
    '         CodeStr(i) = SSProcess.GetObjectAttr(LinArr(i - UBound(PoiArr) ),"SSObj_Code")
    '     End If
    ' Next 'i
    CodeNameVal = CDCodeStr & ";" & CTCodeStr & ";" & CYCodeStr & ";" & CQCodeStr & ";" & CSCodeStr & ";" & QTCodeStr & ";" & BMCodeStr & ";" & FQCodeStr & ";" & DLCodeStr & ";" & GDCodeStr & ";" & LDCodeStr & ";" & DCCodeStr & ";" & XHCodeStr & ";" & TXCodeStr & ";" & DXCodeStr & ";" & YDCodeStr & ";" & LTCodeStr & ";" & JXCodeStr & ";" & JKCodeStr & ";" & DSCodeStr & ";" & BZCodeStr & ";" & JSCodeStr & ";" & XFCodeStr & ";" & PSCodeStr & ";" & YSCodeStr & ";" & WSCodeStr & ";" & FSCodeStr & ";" & RQCodeStr & ";" & MQCodeStr & ";" & YHCodeStr & ";" & RLCodeStr & ";" & RSCodeStr & ";" & ZQCodeStr & ";" & SYCodeStr & ";" & GSCodeStr & ";" & EXCodeStr & ";" & TRCodeStr
    CodeNameArr = Split(CodeNameVal,";", - 1,1)
    For i = 0 To UBound(CodeNameArr)
        If CodeNameArr(i) <> "" Then
            If TempCodeStr = "" Then
                TempCodeStr = CodeNameArr(i)
            Else
                TempCodeStr = TempCodeStr & "," & CodeNameArr(i)
            End If
        End If
    Next 'i
    CodeStr = Split(TempCodeStr,",", - 1,1)
    DelRepeat CodeStr,CodeVal,Count
    DelHiddenLine CodeVal,DisplayCode,CodeCount
End Function' SelAllPoi


'���Ƶ�ע��
Function DrawPoint(ByVal X,ByVal Y,ByVal Code)
    SSProcess.CreateNewObj 0
    SSProcess.SetNewObjValue "SSObj_Code", Code
    SSProcess.SetNewObjValue "SSObj_LayerName", "����ͼ����"
    'SSProcess.SetNewObjValue "SSObj_GroupID", GropuId
    SSProcess.SetNewObjValue "SSObj_Color", SSProcess.GetFeatureCodeInfo(Code,"LineColor")
    SSProcess.AddNewObjPoint X,Y,0,0,""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
    DrawPointNote X + 2.5,Y,Code,150,150
End Function

'���Ƶ�ע����
Function DrawPointNote(ByVal X,ByVal Y,ByVal Code,ByVal Width,ByVal Height)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontString", SSProcess.GetFeatureCodeInfo(Code,"ObjectName")
    SSProcess.SetNewObjValue "SSObj_LayerName", "����ͼ����"
    'SSProcess.SetNewObjValue "SSObj_GroupID", GropuId
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "3"
    SSProcess.SetNewObjValue "SSObj_FontWidth", Width
    SSProcess.SetNewObjValue "SSObj_FontHeight", Height
    SSProcess.AddNewObjPoint X, Y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

'������ע��
Function DrawLine(ByVal X1,ByVal X2,ByVal Y,ByVal Code)
    SSProcess.CreateNewObj 1
    SSProcess.SetNewObjValue "SSObj_Code", Code
    SSProcess.SetNewObjValue "SSObj_LayerName", "����ͼ����"
    'SSProcess.SetNewObjValue "SSObj_GroupID", GropuId
    SSProcess.SetNewObjValue "SSObj_Color", SSProcess.GetFeatureCodeInfo(Code,"LineColor")
    SSProcess.AddNewObjPoint X1, Y, 0, 0, ""
    SSProcess.AddNewObjPoint X2, Y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
    DrawLineNote X2 + 1,Y,Code,150,150
End Function

'������ע����
Function DrawLineNote(ByVal X,ByVal Y,ByVal Code,ByVal Width,ByVal Height)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontString", SSProcess.GetFeatureCodeInfo(Code,"ObjectName")
    SSProcess.SetNewObjValue "SSObj_LayerName", "����ͼ����"
    'SSProcess.SetNewObjValue "SSObj_GroupID", GropuId
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "3"
    SSProcess.SetNewObjValue "SSObj_FontWidth", Width
    SSProcess.SetNewObjValue "SSObj_FontHeight", Height
    SSProcess.AddNewObjPoint X, Y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

'������߻���
Function DrawBorder(ByVal StartX,ByVal EndX,ByVal StartY,ByVal EndY)
    
    SSProcess.CreateNewObj 2
    SSProcess.SetNewObjValue "SSObj_Code", "51111111"
    'SSProcess.SetNewObjValue "SSObj_GroupID", GroupId
    SSProcess.SetNewObjValue "SSObj_LayerName", "����ͼ����"
    SSProcess.SetNewObjValue "SSObj_Color", "RGB(255,255,255)"
    SSProcess.AddNewObjPoint StartX,StartY,0,0,""
    SSProcess.AddNewObjPoint EndX, StartY,0,0,""
    SSProcess.AddNewObjPoint EndX,EndY,0, 0,""
    SSProcess.AddNewObjPoint StartX,EndY,0,0,""
    SSProcess.AddNewObjPoint StartX,StartY,0,0,""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
    
    '���Ʊ���
    DrawTitle (StartX + EndX) / 2,StartY - 1,200,200
    
End Function

'���Ʊ���
Function DrawTitle(ByVal X,ByVal Y,ByVal Width, ByVal Height)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontString", "ͼ ��"
    SSProcess.SetNewObjValue "SSObj_Color", "RGB(255,255,255)"
    SSProcess.SetNewObjValue "SSObj_LayerName", "����ͼ����"
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "0"
    SSProcess.SetNewObjValue "SSObj_FontWidth", Width
    SSProcess.SetNewObjValue "SSObj_FontHeight", Height
    'SSProcess.SetNewObjValue "SSObj_GroupID", GroupId
    SSProcess.AddNewObjPoint X, Y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

'ȷ�������ʽ
Function ConFirmWay(ByRef Way,ByRef res,ByRef GroupStr)
    SSProcess.ClearInputParameter
    SSProcess.AddInputParameter "ѡ�������ʽ","�ۺϹ���ͼ",0,"�ۺϹ���ͼ,�ֲ����",""
    'SSProcess.AddInputParameter "ѡ�����","",0,"",""
    res = SSProcess.ShowInputParameterDlg ("����ͼ�����ʽ")
    SSProcess.RefreshView
    If res = 1  Then
        Way = SSProcess.GetInputParameter("ѡ�������ʽ")
    End If
    GroupStr = ""
    ' GroupStr = SSProcess.GetInputParameter("ѡ�����")
    If GroupStr <> "" Then
        SetPoiNote GroupStr
    End If
End Function' ConFirmWay

'����ע����
Function SetPoiNote(ByVal GroupStr)
    LayArr = Split("CD,CT,CY,CQ,CS,QT,BM,FQ,DL,GD,LD,DC,XH,TX,DX,YD,LT,JX,JK,EX,DS,BZ,JS,XF,PS,YS,WS,FS,RQ,MQ,TR,YH,RL,RS,ZQ,SY,GS",",", - 1,1)
    For i = 0 To UBound(LayArr)
        SelNote LayArr(i) & "��ע��",GroupStr
    Next 'i
End Function' SetPoiNote

'�������е�ע��
Function SelNote(ByVal LayerName,ByVal GroupStr)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", LayerName
    SSProcess.SetSelectCondition "SSObj_Type", "==", "NOTE"
    SSProcess.SelectFilter
    Count = SSProcess.GetSelNoteCount
    For j = 0 To Count - 1
        FormerVal = SSProcess.GetSelNoteValue(j,"SSObj_FontString")
        Prefix = Left(FormerVal,2)
        Suffix = Right(FormerVal,Len(FormerVal) - 2)
        CurrentVal = Prefix & GroupStr & Suffix
        SSProcess.SetSelNoteValue j,"SSObj_FontString",CurrentVal
    Next 'i
End Function' SelNote

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

'ȥ��������Code
Function DelHiddenLine(ByVal CodeStr,ByRef DisplayCode,ByRef DisPlayCount)
    HiddenArr = Split(HiddenCode,",", - 1,1)
    CodeArr = Split(CodeStr,",", - 1,1)
    For i = 0 To UBound(CodeArr)
        For j = 0 To UBound(HiddenArr)
            If CodeArr(i) = HiddenArr(j) Then
                CodeArr(i) = ""
            End If
        Next 'i
    Next 'i
    
    DisplayCode = ""
    
    For i = 0 To UBound(CodeArr)
        If CodeArr(i) <> "" Then
            If DisplayCode = "" Then
                DisplayCode = CodeArr(i)
            Else
                DisplayCode = DisplayCode & "," & CodeArr(i)
            End If
        End If
    Next 'i
    
    DisPlayArr = Split(DisplayCode,",", - 1,1)
    DisPlayCount = UBound(DisPlayArr) + 1
End Function' DelHiddenLine

Function GxVisible(ByVal LayString)
    count = SSProcess.GetLayerCount
    For i = 0 To count - 1
        layername = SSProcess.GetLayerName (i)
        SSProcess.SetLayerStatus layername, 1, 0
    Next
    LayArr = Split(LayString,",", - 1,1)
    For i = 0 To UBound(LayArr)
        SSProcess.SetLayerStatus LayArr(i), 1, 1
    Next 'i
    SSProcess.SetLayerStatus "ͼ����", 1, 1
    SSProcess.SetLayerStatus "TK", 1, 1
    SSProcess.SetLayerStatus "��ѧ����", 1, 1
    SSProcess.SetLayerStatus "����ͼ����", 1, 1
    SSProcess.RefreshView
End Function

Function AllVisible()
    count = SSProcess.GetLayerCount
    For i = 0 To count - 1
        layername = SSProcess.GetLayerName (i)
        SSProcess.SetLayerStatus layername, 1, 1
    Next
    SSProcess.RefreshView
End Function

Function DelTk()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", "����ͼ����"
    SSProcess.SelectFilter
    SSProcess.DeleteSelectionObj()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", "ͼ����"
    SSProcess.SelectFilter
    SSProcess.DeleteSelectionObj()
End Function' DelTk

'�����ʾ
Function Ending()
    MsgBox "������"
End Function' Ending