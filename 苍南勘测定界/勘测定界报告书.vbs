
'========================================================�ļ�·����������================================================================

'·����������
Dim FileSysObj
Set FileSysObj = CreateObject("Scripting.FileSystemObject")

'���ƺ�·���ַ���
Dim CopyPathStr
CopyPathStr = ""

'Word��������
Dim g_docObj

'==============================================================�������==================================================================

Sub OnClick()
    
    strTempFileName = "�������ڵ�λ��ͼģ��.docx"
    strTempFilePath = SSProcess.GetSysPathName (7) & "\����ģ��\" & strTempFileName
    
    Set g_docObj = CreateObject ("asposewordscom.asposewordshelper")
    
    If TypeName (g_docObj) = "AsposeWordsHelper" Then
        g_docObj.CreateDocumentByTemplate strTempFilePath
    Else
        MsgBox "����ע��Aspose.Word���"
        Exit Sub
    End If
    
    pathName = GetFilePath
    
    InitDB()
    
    ReplaceCell ZhuJi '�滻��Ԫ������
    
    ReleaseDB()
    
    strWordFileName = Replace(strTempFileName, "ģ��", "")
    
    strFileSavePath = pathName & "����ɹ�\" & strWordFileName
    
    g_docObj.SaveEx strFileSavePath
    
    OpenProject strFileSavePath,pathName,strWordFileName,ZhuJi
    
    Set TempFile = FileSysObj.GetFile(strFileSavePath)
    
    TempFile.Delete
    
    MsgBox "������"
    
End Sub

Function OpenProject(ByVal strFileSavePath,ByVal pathName,ByVal strWordFileName,ByVal ZhuJi)
    
    '��������
    EdbNameStr = ""
    
    '�������ط�Χ�棨504����ճ����
    CloneArea
    
    'ѡ���ļ���·��(��ѡ֮����","���зָ�)
    FilePathStr = SSProcess.SelectFileName(1,"ѡ���ļ�",1,"EDB Files (*.edb)|*.edb|All Files (*.*)|*.*||")
    
    FilePathArr = Split(FilePathStr,",", - 1,1)
    
    '���ƹ���
    For i = 0 To UBound(FilePathArr)
        Set EdbFile = FileSysObj.GetFile(FilePathArr(i))
        EdbFile.Copy SSProcess.GetSysPathName(5) & "����ɹ�\" & EdbFile.Name
        If CopyPathStr = "" Then
            CopyPathStr = SSProcess.GetSysPathName(5) & "����ɹ�\" & EdbFile.Name
        Else
            CopyPathStr = CopyPathStr & "," & SSProcess.GetSysPathName(5) & "����ɹ�\" & EdbFile.Name
        End If
        If EdbNameStr = "" Then
            EdbNameStr = EdbFile.Name
        Else
            EdbNameStr = EdbNameStr & "," & EdbFile.Name
        End If
    Next 'i
    
    EdbNameArr = Split(EdbNameStr,",", - 1,1)
    
    For i = 0 To UBound(EdbNameArr)
        CreatPath = pathName & "����ɹ�\" & Replace(EdbNameArr(i),".edb","") & strWordFileName
        g_docObj.CreateDocumentByTemplate strFileSavePath
        ReadDT Replace(EdbNameArr(i),".edb","")
        g_docObj.SaveEx CreatPath
    Next 'i 
    
    EdbScale = SSProcess.GetMapScale()
    
    '��ѡ��Ĺ��̲�ճ��
    CopyPathArr = Split(CopyPathStr,",", - 1,1)
    
    For i = 0 To UBound(CopyPathArr)
        
        SSProcess.OpenDatabase CopyPathArr(i)
        SSProcess.SetMapScale(EdbScale)
        SSProcess.AddClipBoardObjToMap 0,0
        CopyArr = Split(CopyPathArr(i),"\", - 1,1)
        EdbName = Replace(CopyArr(UBound(CopyArr)),".edb","")
        g_docObj.OpenDocument pathName & "����ɹ�\" & EdbName & strWordFileName
        
        NotePosition X,Y
        
        CJCTFWX MinX,MinY,MaxX,MaxY
        
        ZhuJiArr = Split(ZhuJi,";", - 1,1)
        
        For j = 0 To UBound(ZhuJiArr)
            DrawNote ZhuJiArr(j),X,Y + j * 6,500 / EdbScale
        Next 'j
        
        DrawCompass EdbScale,MinX,MinY,MaxX,MaxY
        
        InsterPicture MinX,MinY,MaxX,MaxY
        
        g_docObj.SaveEx pathName & "����ɹ�\" & EdbName & strWordFileName
        'SSProcess.CloseDatabase
    Next 'i
    
End Function' OpenProject()

'����ͼ����ָ����
Function DrawCompass(ByVal EdbScale,ByRef MinX,ByRef MinY,ByRef MaxX,ByRef MaxY)
    If EdbScale = 500 Then
        makePoint MaxX - 10,MaxY - 10,"9120066",RGB(255,255,255),polygonID  '����ָ����
        makePoint MinX + 10,MinY + 10,"9120046",RGB(255,255,255),polygonID  '����ͼ��
    ElseIf EdbScale = 5000 Then
        makePoint MaxX,MaxY,"9120066",RGB(255,255,255),polygonID  '����ָ����
        makePoint MinX,MinY,"9120047",RGB(255,255,255),polygonID  '����ͼ��
    End If
    'makeArea MinX,MinY,MaxX,MinY,MaxX,MaxY,MinX,MaxY,2,RGB(255,255,255)
End Function' DrawNote

'����ע��λ��
Function NotePosition(ByRef X,ByRef Y)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 504
    SSProcess.SelectFilter
    SelCount = SSProcess.GetSelGeoCount
    ID = SSProcess.GetSelGeoValue(0,"SSObj_ID")
    SSProcess.GetObjectFocusPoint ID,X,Y
End Function' NotePosition

'ѡ��Ҫ�ظ��Ƶ�ճ����
Function CloneArea()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 504
    SSProcess.SelectFilter
    SSProcess.SelectionObjToClipBoard
End Function' CloneArea

'����ͼƬ
Function InsterPicture(ByVal MinX,ByVal MinY,ByVal MaxX,ByVal MaxY)
    
    Path = SSProcess.GetSysPathName(7) & "Pictures\"
    StrBmpFile = Path & "RFT" & i & ".wmf"
    Dpi = 300

    SSFunc.DrawToImage MinX,MinY,MaxX,MaxY,"100" & "X" & "100",Dpi,StrBmpFile
    Rotation = 0
    
    Width = 100 * 4.28
    Height = 100 * 4.28
    
    g_docObj.MoveToCell TableIndex,5,0,0
    
    g_docObj.InsertImage StrBmpFile,Width,Height,Rotation
    
End Function' InsterPicture

'��ȡ���ط�Χ����    
Function CJCTFWX(ByRef MinX,ByRef MinY,ByRef MaxX,ByRef MaxY)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_code", "==", "504"
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount
    dh = 0
    For i = 0 To geocount - 1
        pointcount = SSProcess.GetSelGeoPointCount(i)
        objID = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        Dim x, y, z, pointtype, name
        For j = 0 To pointcount - 1
            dh = dh + 1
            SSProcess.GetObjectPoint objID, j, x, y, z, pointtype, name
            If dh <> 1 Then
                If  x > maxx Then  maxx = x
                If  x < minx Then  minx = x
                If  y > maxy Then  maxy = y
                If  y < miny Then  miny = y
            Else
                maxx = x
                minx = x
                maxy = y
                miny = y
            End If
        Next
    Next
    
    '��С������
    MinX = MinX - 10
    MaxX = MaxX + 10
    MinY = MinY - 10
    MaxY = MaxY + 10
    
End Function

Function makeArea(x1,y1,x2,y2,x3,y3,x4,y4,code,color)
    
    SSProcess.CreateNewObj 2
    SSProcess.SetNewObjValue "SSObj_Code", code
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjPoint x2, y2, 0, 0, ""
    SSProcess.AddNewObjPoint x3, y3, 0, 0, ""
    SSProcess.AddNewObjPoint x4,y4, 0, 0, ""
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
    
End Function

Function makePoint(x,y,code,color,polygonID)
    SSProcess.CreateNewObj 0
    SSProcess.SetNewObjValue "SSObj_Code", code
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.AddNewObjPoint x, y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

Function ReadDT(ByVal YearTime)
    Tablename = "����������������״ͼ��" & YearTime & "���ڵ���λͼ"
    g_docObj.Replace "{Tablename}",Tablename,0
End Function

Function ReplaceCell(ByRef Hr)
    SuoZXZ = WriteFormHX  '�Ӻ��߶�ȡ������дģ�壬��������������
    Res_value = WriteFormDLTB '��ͼ�߶�ȡ������дģ�壬����������Ȩ���͵�ͼҪʹ�õ�ע��
    
    ZhuJi = Split(Res_value, "||")(0) '��ͼ�ϵ�ע��
    ZhuJi = Replace(ZhuJi, " ", "��")
    ZhuJi = Right(ZhuJi, Len(ZhuJi) - 1)
    ZjArr = Split(ZhuJi,"��", - 1,1)
    Hr = ""
    For i = 0 To UBound(ZjArr)
        If Hr = "" Then
            Hr = ZjArr(i)
        Else
            If i > 0 And i Mod 3 = 0 Then
                Hr = Hr & ";" & ZjArr(i)
            Else
                Hr = Hr & "��" & ZjArr(i)
            End If
        End If
    Next 'i
    TuDQS = SuoZXZ & Split(Res_value, "||")(1) '����Ȩ��
    
    g_docObj.Replace "{TuDQS}",TuDQS,0
    
    MapScale = SSProcess.GetMapScale
    MapScale = "1:" & MapScale
    g_docObj.Replace "{MapScale}",MapScale,0
    
End Function

Function DrawNote(ByVal BZStr,ByVal X,ByVal Y,ByVal NoteScale)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontString", BZStr
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "5"
    SSProcess.SetNewObjValue "SSObj_FontWidth", 1000 * NoteScale
    SSProcess.SetNewObjValue "SSObj_Color", "RGB(0,255,191)"
    SSProcess.SetNewObjValue "SSObj_FontHeight", 1000 * NoteScale
    SSProcess.SetNewObjValue "SSObj_FontDirection", 0
    SSProcess.AddNewObjPoint X,Y,0,0,""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
    SSProcess.RefreshView
End Function' DrawNote

Function WriteFormDLTB
    '���ݵ���ͼ���������Ȩ�����������͡�{TuDQS}��{TuDLX} 
    TuDLX = ""
    ZhuJI = ""
    TuDQS = ""
    mdbName = SSProcess.GetProjectFileName
    sql = "Select DISTINCT ����ͼ�����Ա�.dlmc From ����ͼ�����Ա� INNER JOIN GeoAreaTB ON ����ͼ�����Ա�.ID = GeoAreaTB.ID WHERE([GeoAreaTB].[Mark] Mod 2)<>0"
    GetSQLRecordAll mdbName,sql,arSQLRecord,iRecordCount
    For i = 0 To iRecordCount - 1
        dlmc = arSQLRecord (i)
        TuDLX = TuDLX & " " & dlmc
        
        sql1 = "Select sum (����ͼ�����Ա�.tbmj) From ����ͼ�����Ա� INNER JOIN GeoAreaTB ON ����ͼ�����Ա�.ID = GeoAreaTB.ID WHERE  ([GeoAreaTB].[Mark] Mod 2)<>0 and ����ͼ�����Ա�.dlmc = '" & dlmc & "'"
        GetSQLRecordAll mdbName,sql1,arSQLRecord1,iRecordCount1
        For j = 0 To iRecordCount1 - 1
            message = dlmc & arSQLRecord1(j) & "����"
            ZhuJI = ZhuJI & " " & message
        Next
    Next
    
    sql2 = "Select DISTINCT ����ͼ�����Ա�.qsdw From ����ͼ�����Ա� INNER JOIN GeoAreaTB ON ����ͼ�����Ա�.ID = GeoAreaTB.ID WHERE([GeoAreaTB].[Mark] Mod 2)<>0"
    GetSQLRecordAll mdbName,sql2,arSQLRecord2,iRecordCount2
    For k = 0 To iRecordCount2 - 1
        qsdw = arSQLRecord2 (k)
        TuDQS = TuDQS & "��" & qsdw
        
        TuDQS = Right(TuDQS, Len(TuDQS) - 1)
        
    Next
    TuDLX = Replace(TuDLX, " ", "��")
    
    TuDLX = Right(TuDLX, Len(TuDLX) - 1)
    
    g_docObj.Replace "{TuDLX}",TuDLX,0
    WriteFormDLTB = ZhuJI & "||" & TuDQS
    
End Function


Function WriteFormHX
    '���ݺ��������е��õص�λ,��Ŀ����,�ؿ���� {XMMC},{YDDW},{DKMJ}��ģ���еĶ�Ӧ�ֶ�
    ' ����������,{SuoZXZ}����
    values = "XMMC,YDDW,DKMJ"
    valuesList = Split(values,",")
    SqlStr = "Select �������Ա�.DKMJ,YDDW,XMMC From �������Ա� Inner Join GeoAreaTB on �������Ա�.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0"
    ProJectName = SSProcess.GetProjectFileName
    GetSQLRecordAll ProJectName,SqlStr,InfoArr,DKCount
    Mj = 0
    For i = 0 To DKCount - 1
        TotalArr = Split(InfoArr(i),",", - 1,1)
        If Mj = 0 Then
            Mj = Transform(TotalArr(0))
        Else
            Mj = Mj + Transform(TotalArr(0))
        End If
    Next 'i
    First = Split(InfoArr(0),",", - 1,1)
    ValStr = First(1) & "," & First(2) & "," & Mj
    strFieldValue = Split(ValStr,",", - 1,1)
    For i = 0 To UBound(valuesList)
        'strFieldValue = ""
        strField = valuesList(i)
        'listCount = GetProjectTableList ("�������Ա�",strField," �������Ա�.ID>0 ","SpatialData","2",list,fieldCount)
        'If listCount = 1 Then strFieldValue = list(0,0)
        g_docObj.Replace "{" & strField & "}",strFieldValue(i),0
    Next
    
    listCount = GetProjectTableList ("�������Ա�","SuoZXZ"," �������Ա�.ID>0 ","SpatialData","2",list,fieldCount)
    If listCount = 1 Then
        SuoZXZ = list(0,0)
    End If
    WriteFormHX = SuoZXZ
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
            arSQLRecord(iRecordCount) = values    '��ѯ��¼
            iRecordCount = iRecordCount + 1        '��ѯ��¼��
            '�ƶ���¼�α�
            SSProcess.AccessMoveNext mdbName, sql
        WEnd
    End If
    '�رռ�¼��
    SSProcess.CloseAccessRecordset mdbName, sql
    SSProcess.CloseAccessMdb mdbName
End Function


'��ȡ�ɹ�Ŀ¼·��
Function  GetFilePath
    projectFileName = SSProcess.GetSysPathName (5)
    GetFilePath = projectFileName
End Function

'***********************���ݿ��������*********************
'//����
Dim  adoConnection
Function InitDB()
    accessName = SSProcess.GetProjectFileName
    Set adoConnection = CreateObject("adodb.connection")
    strcon = "DBQ=" & accessName & ";DRIVER={Microsoft Access Driver (*.mdb)};"
    adoConnection.Open strcon
End Function


'�ؿ�
Function ReleaseDB()
    adoConnection.Close
    Set adoConnection = Nothing
End Function


'�ݹ鴴���༶Ŀ¼
Function CreateFolder(path)
    Set fso = CreateObject("scripting.filesystemobject")
    If fso.FolderExists(path) Then
        Exit Function
    End If
    If Not fso.FolderExists(fso.GetParentFolderName(path)) Then
        CreateFolder fso.GetParentFolderName(path)
    End If
    fso.CreateFolder(path)
    Set fso = Nothing
End Function


'SQL��ѯ�ֶ�
Function GetProjectTableList(ByVal strTableName,ByVal strFields,ByVal strAddCondition,ByVal strTableType,ByVal strGeoType,ByRef rs(),ByRef fieldCount)
    'strTableName ��
    'strFields �ֶ�
    'strAddCondition ���� 
    'strTableType AttributeData(�����Ա�) ,SpatialData(�������Ա�)
    'strGeoType �������� �㡢�ߡ��桢ע��(0��,1��,2��,3ע��)
    'rs ���¼��ά����rs(��,��)
    'fieldCount �ֶθ���
    '����ֵ :sql��ѯ���¼����
    
    
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
    
    '��ȡ��ǰ����edb���¼
    AccessName = SSProcess.GetProjectFileName
    '�жϱ��Ƿ����
    Set adoRs = CreateObject("ADODB.recordset")
    count = 0
    adoRs.cursorLocation = 3
    adoRs.cursorType = 3
    
    adoRs.open sql,adoConnection,3,3
    rcdCount = adoRs.RecordCount
    fieldCount = adoRs.Fields.Count
    ReDim rs(rcdCount,fieldCount)
    
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
    
    GetProjectTableList = rsCount
End Function

'��������ת��
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