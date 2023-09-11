
Dim g_docObj
XMZD = "BH,XMMC,XMDZ,SJDW,JSDW,WTDW,CHDW,WYSJ,CHSJ"
GXLayerName = "CD,CT,CY,CQ,CS,QT,BM,FQ,DL,GD,LD,DC,XH,TX,DX,YD,LT,JX,JK,EX,DS,BZ,JS,XF,PS,YS,WS,FS,RQ,MQ,TR,YH,RL,RS,ZQ,SY,GS"

GXPointField0 = "TSDH,WTDH"
GXLineField0 = "GXQDDH^GXZDDH"
GXPointField1 = "TZ,FSW,Round(QDX,3),Round(QDY,3),DMGC"
GXLineField1 = "GC,GJ^DMCC,GXQDMS,GXZDMS,GXQDGDGC,GXZDGDGC,DYZ,SJYL,ZKS/YYKS,XLTS,LX,QSDW,FSFS,JCNY,BZ"
Sub OnClick()
    Set g_docObj = CreateObject ("AsposeCellsCom.AsposeCellsHelper")
    If  TypeName (g_docObj) <> "AsposeCellsHelper" Then
        MsgBox "����ע��Aspose.Excel���"
        Exit Sub
    End If
    InitDB()
    str = GXAddInputParameter(filename,ExportMark,frameCount)
    If str = True Then
        pathName = SSProcess.GetSysPathName(5) & "�ɹ���\"
        If pathName <> "" Then
            If ExportMark = "" Then
                '��Ŀ���
                ExportMap pathName, filename
            Else
                'ͼ�����
                ExportFrame pathName, filename, frameCount
            End If
        Else
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    ReleaseDB()
    MsgBox "������"
End Sub


Function GXAddInputParameter(ByRef filename,ByRef ExportMark,ByRef frameCount)
    str = True
    SSProcess.ClearInputParameter
    SSProcess.AddInputParameter "���߱��������ʽ", "��ͼ�����", 0, "��ͼ�����,����Ŀ���", ""
    result = SSProcess.ShowInputParameterDlg ("���߱��������ʽ")
    If result = 1 Then
        res = SSProcess.GetInputParameter ("���߱��������ʽ")
        If res = "��ͼ�����" Then
            filename = SSProcess.GetSysPathName (7) & "\" & "���߱���ģ�壨ͼ����.xlsx"
            SSProcess.ClearInputParameter
            SSProcess.AddInputParameter "ͼ����������ʽ", "����ͼ�����", 0, "����ͼ�����,����Χ�߷ַ����,ȫͼ�ַ����", ""
            result1 = SSProcess.ShowInputParameterDlg ("ͼ����������ʽ")
            If result1 = 1 Then
                res1 = SSProcess.GetInputParameter ("ͼ����������ʽ")
                SSProcess.ClearSelection
                SSProcess.ClearSelectCondition
                If res1 = "����ͼ�����" Then
                    frameID = SSProcess.GetCurMapFrame()
                    SSProcess.CreateMapFrameByRegionID  frameID
                    frameCount = SSProcess.GetMapFrameCount()
                ElseIf res1 = "����Χ�߷ַ����" Then
                    SSProcess.CreateMapFrameByRegion 1
                    frameCount = SSProcess.GetMapFrameCount()
                ElseIf res1 = "ȫͼ�ַ����" Then
                    SSProcess.CreateMapFrame
                    frameCount = SSProcess.GetMapFrameCount()
                End If
                ExportMark = 1
            Else
                Exit Function
                str = False
            End If
        Else
            filename = SSProcess.GetSysPathName (7) & "\" & "���߱���ģ�壨��Ŀ��.xlsx"
            ExportMark = ""
        End If
    Else
        Exit Function
        str = False
    End If
    GXAddInputParameter = str
End Function

'��Ŀ���
Function ExportMap(ByVal pathName,ByVal filename)
    g_docObj.CreateDocumentByTemplate filename
    '��ͷ
    'XMMC=SSProcess.ReadEpsIni("���߱�����Ϣ", "��Ŀ����" ,"")
    'CHDW=SSProcess.ReadEpsIni("���߱�����Ϣ", "��浥λ" ,""):RQDate=year(Date())&"��"&month(Date())&"��"
    GetXMXX GXXMXX
    XMMC = GXXMXX(1)
    CHDW = GXXMXX(6)
    RQDate = Year(Date()) & "��" & Month(Date()) & "��"
    '���Excel·��
    exlfilePathName = pathName & XMMC & ".xlsx"
    '��ȡȫͼ�ڵĹ���ͼ��
    strLayer = GetMapGXLayer(strLayer)
    strLayerList = Split(strLayer,",")
    '����excelsheet
    For i1 = 0 To UBound(strLayerList)
        g_docObj.CopySheet "Sheet1",strLayerList(i1)
    Next
    
    '����ͼ��˳����sheetֵ
    For i1 = 0 To UBound(strLayerList)
        g_docObj.SetActiveSheet strLayerList(i1)
        
        g_docObj.SetCellValueEx 0,0,"�������ƣ�" & XMMC
        
        ids = GetGXLayerID( strLayerList(i1))
        idsList = Split(ids,",")
        ReDim CellList(Count)
        Count = 0
        For i2 = 0 To UBound(idsList)
            value = ""
            CellValue = ""
            layer = SSProcess.GetObjectAttr( idsList(i2), "SSObj_LayerName")
            If layer = strLayerList(i1) Then
                value = SSProcess.GetObjectAttr( idsList(i2), "[WTDH]")
                WTDH = value
                CellValue = WTDH
                
                '���ߵ���ֵ�жϱ�ʶ
                rMarkCount = 0
                '���ӵ��
                ZDDHCount = GetProjectTableList( "���¹��������Ա�", "GXZDDH,���¹��������Ա�.ID", "GXQDDH='" & WTDH & "'", "SpatialData", "1", ZDDHList, fieldCount)
                If ZDDHCount > 0 Then
                    For i3 = 0 To ZDDHCount - 1
                        If rMarkCount > 0 Then CellValue = ""
                        '��ȡ�ܵ�Ϊ����������������
                        GetCellValueList i3,ZDDHList,"GXQDDH",WTDH, CellValue, CellList, Count,rMarkCount
                    Next
                End If
                QDDHCount = GetProjectTableList( "���¹��������Ա�", "GXQDDH,���¹��������Ա�.ID", "GXZDDH='" & WTDH & "'", "SpatialData", "1", QDDHList, fieldCount)
                If QDDHCount > 0 Then
                    For i3 = 0 To QDDHCount - 1
                        If rMarkCount > 0 Then CellValue = ""
                        '��ȡ�ܵ�Ϊ�����յ����������
                        GetCellValueList i3,QDDHList,"GXZDDH",WTDH, CellValue, CellList, Count,rMarkCount
                    Next
                End If
            End If
        Next
        
        '������
        g_docObj.CopySheetRows 3,1,Count - 1
        
        '��ֵ
        startRow = 3
        For i2 = 0 To Count - 1
            CellValueList = Split(CellList(i2),",")
            For i3 = 0 To UBound(CellValueList)
                CellValue = CellValueList(i3)
                g_docObj.SetCellValueEx startRow,i3,CellValue
            Next
            startRow = startRow + 1
        Next
        
        '��д��β
        ' g_docObj.SetCellValueEx Count + 3,0,"��ҵ��λ��" & CHDW
        ' g_docObj.SetCellValueEx Count + 3,10,"���ڣ�" & RQDate
        
        'ɾ����һ��
        g_docObj.DeleteSheetRows 0,1
        
        HeaderStr = "�������ƣ�" & XMMC
        g_docObj.PageSetup2 0,HeaderStr,"��ҵ��λ��" & CHDW & Space(30) & " �Ʊ��ߣ�������"
        g_docObj.PageSetup2 1,"","У���ߣ������" & Space(30) & "���ڣ�" & RQDate
    Next
    
    'ɾ��sheet1
    g_docObj.DeleteSheet "Sheet1"
    g_docObj.SaveEx exlfilePathName, 0
    
End Function


'ͼ�����
Function ExportFrame(ByVal pathName,ByVal filename,ByVal frameCount)
    For i = 0 To frameCount - 1
        g_docObj.CreateDocumentByTemplate filename
        
        SSProcess.GetMapFrameCenterPoint i, x, y
        SSProcess.SetCurMapFrame x, y, 0, ""
        frameID = SSProcess.GetCurMapFrame()
        mapNumber = SSProcess.GetCurMapFrameNumber()
        '���Excel·��
        exlfilePathName = pathName & mapNumber & ".xlsx"
        
        ids = SSProcess.SearchInPolyObjIDs(frameID, 0, "", 1, 1, 1)
        idsList = Split(ids,",")
        '��ȡͼ���ڵĹ���ͼ��
        strLayer = GetFrameGXLayer( ids, strLayer)
        strLayerList = Split(strLayer,",")
        '����excelsheet
        For i1 = 0 To UBound(strLayerList)
            g_docObj.CopySheet "Sheet1",strLayerList(i1)
        Next
        '��ȡ��Ŀ��Ϣ
        GetXMXX GXXMXX
        XMMC = GXXMXX(1)
        CHDW = GXXMXX(6)
        '����ͼ��˳����sheetֵ
        For i1 = 0 To UBound(strLayerList)
            g_docObj.SetActiveSheet strLayerList(i1)
            '��ͷ
            ''δ��д
            'XMMC=SSProcess.ReadEpsIni("���߱�����Ϣ", "��Ŀ����" ,""):XMBH=SSProcess.ReadEpsIni("���߱�����Ϣ", "���" ,"")
            'CHDW=SSProcess.ReadEpsIni("���߱�����Ϣ", "��浥λ" ,""):RQDate=year(Date())&"��"&month(Date())&"��"
            
            g_docObj.SetCellValueEx 0,0,"�������ƣ�" & XMMC
            g_docObj.SetCellValueEx 0,8,"���̱�ţ�" & XMBH
            g_docObj.SetCellValueEx 0,17,"ͼ���ţ�" & mapNumber
            
            ReDim CellList(Count)
            Count = 0
            For i2 = 0 To UBound(idsList)
                value = ""
                CellValue = ""
                layer = SSProcess.GetObjectAttr( idsList(i2), "SSObj_LayerName")
                If layer = strLayerList(i1) Then
                    Point0List = Split(GXPointField0,",")
                    str = ""
                    For i3 = 0 To UBound(Point0List)
                        value = SSProcess.GetObjectAttr( idsList(i2), "[" & Point0List(i3) & "]")
                        If i3 = 1 Then WTDH = value
                        CellValue = GetValueString( i3, value,str)
                    Next
                    '���ߵ���ֵ�жϱ�ʶ
                    rMarkCount = 0
                    '���ӵ��
                    ZDDHCount = GetProjectTableList( "���¹��������Ա�", "GXZDDH,���¹��������Ա�.ID", "GXQDDH='" & WTDH & "'", "SpatialData", "1", ZDDHList, fieldCount)
                    If ZDDHCount > 0 Then
                        For i3 = 0 To ZDDHCount - 1
                            If rMarkCount > 0 Then CellValue = ","
                            '��ȡ�ܵ�Ϊ����������������
                            GetCellValueList i3,ZDDHList,"GXQDDH",WTDH, CellValue, CellList, Count,rMarkCount
                        Next
                    End If
                    QDDHCount = GetProjectTableList( "���¹��������Ա�", "GXQDDH,���¹��������Ա�.ID", "GXZDDH='" & WTDH & "'", "SpatialData", "1", QDDHList, fieldCount)
                    If QDDHCount > 0 Then
                        For i3 = 0 To QDDHCount - 1
                            If rMarkCount > 0 Then CellValue = ","
                            '��ȡ�ܵ�Ϊ�����յ����������
                            GetCellValueList i3,QDDHList,"GXZDDH",WTDH, CellValue, CellList, Count,rMarkCount
                        Next
                    End If
                End If
            Next
            
            '������
            g_docObj.CopySheetRows 3,1,Count - 1
            
            '��ֵ
            startRow = 3
            For i2 = 0 To Count - 1
                CellValueList = Split(CellList(i2),",")
                For i3 = 0 To UBound(CellValueList)
                    CellValue = CellValueList(i3)
                    g_docObj.SetCellValueEx startRow,i3,CellValue
                Next
                startRow = startRow + 1
            Next
            
            '��д��β
            ' g_docObj.SetCellValueEx Count + 3,0,"��ҵ��λ��" & CHDW
            ' g_docObj.SetCellValueEx Count + 3,10,"���ڣ�" & RQDate
            
            'ɾ����һ��
            g_docObj.DeleteSheetRows 0,1
            
            HeaderStr = "�������ƣ�" & XMMC
            g_docObj.PageSetup2 0,HeaderStr,"��ҵ��λ��" & CHDW & Space(30) & " �Ʊ��ߣ�������"
            g_docObj.PageSetup2 1,"","У���ߣ������" & Space(30) & "���ڣ�" & RQDate
            
        Next
        
        
        'ɾ��sheet1
        g_docObj.DeleteSheet "Sheet1"
        g_docObj.SaveEx exlfilePathName, 0
    Next
    SSProcess.FreeMapFrame()
End Function

'��ȡ������������
Function GetCellValueList(ByVal i,ByVal DHList,ByVal GXLineDH,ByVal WTDH,ByVal CellValue,ByRef CellList(),ByRef Count,ByRef rMarkCount)
    GXCellDH = DHList(i,0)
    GXID = DHList(i,1)
    CellValue = CellValue & "," & GXCellDH
    PonitValueCount = GetProjectTableList( "���¹��ߵ����Ա�", GXPointField1, "WTDH='" & WTDH & "'", "SpatialData", "0", GXDList, fieldCount)
    If PonitValueCount = 1 Then
        If rMarkCount = 0 Then
            For i4 = 0 To fieldCount - 1
                value1 = GXDList(0,i4)
                CellValue = GetValueString( i3, value1,CellValue)
            Next
        Else
            CellValue = CellValue & ",,,,,"
        End If
    End If
    
    '�������ԣ�^Ϊ���ߣ�/Ϊ����
    GXLineField1List = Split(GXLineField1,",")
    For i3 = 0 To UBound(GXLineField1List)
        If InStr(GXLineField1List(i3),"^") > 0 Then
            GJList = Split(GXLineField1List(i3),"^")
            GJvalue = SSProcess.GetObjectAttr( GXID, "[GJ]")
            DMCCValue = SSProcess.GetObjectAttr( GXID, "[DMCC]")
            '�ܾ��Ͷ���ߴ��ֵ�ж����ĸ�ֵ
            If GJvalue <> "" And DMCCValue = "" Then
                GXLineFields = GJList(0)
            ElseIf  GJvalue = "" And DMCCValue <> "" Then
                GXLineFields = GJList(1)
            End If
        ElseIf InStr(GXLineField1List(i3),"/") > 0 Then
            KSList = Split(GXLineField1List(i3),"/")
            GXLineFields = KSList(0) & "," & KSList(1)
        Else
            GXLineFields = GXLineField1List(i3)
        End If
        If InStr(GXLineFields,",") > 0 Then
            GXLineCount = GetProjectTableList( "���¹��������Ա�", GXLineFields, "" & GXLineDH & "='" & WTDH & "' and ���¹��������Ա�.ID=" & GXID, "SpatialData", "1", ValueList, fieldCount)
            ZKS = ValueList(0,0)
            YYKS = ValueList(0,1)
            '����ܿ��������ÿ���
            If ZKS <> "" And YYKS = "" Then
                value1 = ZKS
            ElseIf ZKS <> "" And YYKS <> "" Then
                value1 = ZKS & "/" & YYKS
            ElseIf ZKS = "" And YYKS = "" Then
                value1 = ""
            End If
            CellValue = GetValueString( i3, value1,CellValue)
        Else
            GXLineCount = GetProjectTableList( "���¹��������Ա�", GXLineFields, "" & GXLineDH & "='" & WTDH & "' and ���¹��������Ա�.ID=" & GXID, "SpatialData", "1", ValueList, fieldCount)
            value1 = ValueList(0,0)
            CellValue = GetValueString( i3, value1,CellValue)
        End If
    Next
    CellValue = Replace(CellValue,"*","")
    ReDim Preserve CellList(Count)
    CellList(Count) = CellValue
    Count = Count + 1
    rMarkCount = rMarkCount + 1
End Function

'��ȡȫͼ�Ĺ��ߵ�ͼ��
Function GetMapGXLayer(ByRef strLayer)
    strLayer = ""
    GXLXCount = GetProjectTableList( "���¹��������Ա�", "distinct(GXLX)", "", "SpatialData", "1", ValueList, fieldCount)
    For i = 0 To GXLXCount - 1
        GXLX = ValueList(i,0)
        strLayer = GetString( GXLX, "," , strLayer)
    Next
    GetMapGXLayer = strLayer
End Function

'ѡ�񼯻�ȡָ��ͼ��Ĺܵ�id
Function GetGXLayerID(ByVal strLayer)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName","==",strLayer
    SSProcess.SetSelectCondition "SSObj_Type","==","POINT"
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    str = ""
    For i = 0 To geocount - 1
        id = SSProcess.GetSelGeoValue( i, "SSObj_ID")
        ids = GetString(id, "," , str)
    Next
    GetGXLayerID = ids
End Function


'��ȡͼ���ڵĹ���ͼ��
Function GetFrameGXLayer(ByVal ids,ByRef strLayer)
    idlist = Split(ids,",")
    strLayer = ""
    For i = 0 To UBound(idlist)
        layer = SSProcess.GetObjectAttr( idlist(i), "SSObj_LayerName")
        If InStr(GXLayerName,layer) > 0  Then
            If InStr(strLayer,layer) = 0 Then
                strLayer = GetString( layer, "," , strLayer)
            End If
        End If
    Next
    GetFrameGXLayer = strLayer
End Function

'������ַ���
Function GetString(ByVal value,ByVal splitMark , str)
    If str = "" Then
        str = value
    Else
        str = str & splitMark & value
    End If
    GetString = str
End Function

'�����ֶ�ֵ�ַ���
Function GetValueString(ByVal index,ByVal value,ByRef str)
    If str = "" Then
        If index = 0 Then
            str = value
        Else
            str = str & "," & value
        End If
    Else
        str = str & "," & value
    End If
    GetValueString = str
End Function

'//����
Dim  adoConnection
Function InitDB()
    accessName = SSProcess.GetProjectFileName
    Set adoConnection = CreateObject("adodb.connection")
    strcon = "DBQ=" & accessName & ";DRIVER={Microsoft Access Driver (*.mdb)};"
    adoConnection.Open strcon
End Function

'//�ؿ�
Function ReleaseDB()
    adoConnection.Close
    Set adoConnection = Nothing
End Function


'//strTableName ��
'//strFields �ֶ�
'//strAddCondition ���� 
'//strTableType "AttributeData�������Ա� ,SpatialData���������Ա�" 
'//strGeoType �������� �㡢�ߡ��桢ע��(0�㣬1�ߣ�2�棬3ע��)
'//rs ���¼��ά����rs(��,��)
'//fieldCount �ֶθ���
'//����ֵ ��sql��ѯ���¼����
Function GetProjectTableList(ByVal strTableName,ByVal strFields,ByVal strAddCondition,ByVal strTableType,ByVal strGeoType,ByRef rs(),ByRef fieldCount)
    GetProjectTableList = 0
    values = ""
    rsCount = 0
    fieldCount = 0
    If strTableName = "" Or strFields = "" Then Exit Function
    '���õ�������
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
    '��ȡ��ǰ����edb���¼
    AccessName = SSProcess.GetProjectFileName
    '�жϱ��Ƿ����
    'if  IsTableExits(AccessName,strTableName)=false then exit function 
    'set adoConnection=createobject("adodb.connection")
    'strcon="DBQ="& AccessName &";DRIVER={Microsoft Access Driver (*.mdb)};"  
    'adoConnection.Open strcon
    Set adoRs = CreateObject("ADODB.recordset")
    count = 0
    adoRs.cursorLocation = 3
    adoRs.cursorType = 3
    'msgbox sql
    adoRs.open sql,adoConnection,3,3
    rcdCount = adoRs.RecordCount
    fieldCount = adoRs.Fields.Count
    ReDim rs(rcdCount,fieldCount)
    'erase rs
    While adoRs.Eof = False
        nowValues = ""
        For i = 0 To fieldCount - 1
            value = adoRs(i)
            If IsNull(value) Then value = ""
            value = Replace(value,",","��")
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


Function GetXMXX(ByRef XMXXSZ())
    
    mdbName = SSProcess.GetProjectFileName
    sql = "Select ������Ŀ��Ϣ��." & XMZD & " From ������Ŀ��Ϣ��  WHERE ������Ŀ��Ϣ��.id=1"
    GetSQLRecordAll mdbName,sql,arSQLRecord,iRecordCount
    For i = 0 To iRecordCount - 1
        XMXXSZ = Split(arSQLRecord(i), ",")
    Next
End Function


Function GetSQLRecordAll(ByRef mdbName, ByRef StrSqlStatement, ByRef arSQLRecord(), ByRef iRecordCount)
    SSProcess.OpenAccessMdb mdbName
    iRecordCount =  - 1
    sql = StrSqlStatement
    '�򿪼�¼��
    SSProcess.OpenAccessRecordset mdbName, sql
    '��ȡ��¼����
    RecordCount = SSProcess.GetAccessRecordCount (mdbName, sql)
    If RecordCount > 0 Then
        iRecordCount = 0
        ReDim arSQLRecord(RecordCount)
        '����¼�α��Ƶ���һ��
        SSProcess.AccessMoveFirst mdbName, sql
        '�����¼
        While SSProcess.AccessIsEOF (mdbName, sql) = 0
            fields = ""
            values = ""
            '��ȡ��ǰ��¼����
            SSProcess.GetAccessRecord mdbName, sql, fields, values
            arSQLRecord(iRecordCount) = values
            iRecordCount = iRecordCount + 1
            '�ƶ���¼�α�
            SSProcess.AccessMoveNext mdbName, sql
        WEnd
    End If
    '�رռ�¼��
    SSProcess.CloseAccessRecordset mdbName, sql
    SSProcess.CloseAccessMdb mdbName
End Function