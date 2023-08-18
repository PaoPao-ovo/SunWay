
Set FileSystemObject = CreateObject("Scripting.FileSystemObject")

Sub OnClick()
    
    CurrentProPath = Mid(SSProcess.GetProjectFileName,1,Len(SSProcess.GetProjectFileName) - 4) & "����" & ".edb"
    Set FormerFileObj = FileSystemObject.GetFile(SSProcess.GetProjectFileName)
    FormerFileObj.Copy CurrentProPath
    SSProcess.OpenDatabase   CurrentProPath
    ToChinese
    changeLayerCode
    Dim filenames(1000),filecount,strs(1000),count
    PrjPathName = SSProcess.GetSysPathName (2)
    propaths = ""
    GetAllFiles PrjPathName, "prj", filecount, filenames
    If filecount > 0 Then
        For i = 0 To filecount - 1
            If propaths = "" Then
                propaths = Replace(filenames(i),PrjPathName,"")
            Else
                propaths = propaths & "," & Replace(filenames(i),PrjPathName,"")
            End If
        Next
    End If
    SSProcess.ClearInputParameter
    SSProcess.AddInputParameter "ѡ��ͶӰ�ļ�","" ,0, propaths, "ѡ��ͶӰ�ļ�,û�в�ѡ����Ҫ��ͶӰ�ļ��ŵ�eps�����µ�commĿ¼�¡�"
    'SSProcess.AddInputParameter "�ղ���������" , "",0,"", ""
    ret = SSProcess.ShowInputParameterDlg ("��������")
    
    If ret = 0 Then
        Exit Sub
    End If
    If ret = 1 Then
        SSProcess.UpdateScriptDlgParameter 1
        'PCQYMC  = SSProcess.GetInputParameter ("�ղ���������")
    End If
    
    SXZQMC = SSProcess.ReadEpsIni("��������������", "LastAttr" ,"")
    XXZQMC = SSProcess.ReadEpsIni("��������������", "LastAttr" ,"")
    PCQYMC = SXZQMC & XXZQMC
    
    SSProcess.UpdateScriptDlgParameter 1
    SystemFileName = SSProcess.GetInputParameter("ѡ��ͶӰ�ļ�")
    
    'pathName = SSProcess.SelectPathName( )
    SSProcess.ExecuteSDLFunction "$SDL.SSProject.Display.RedrawExtend", 0
    
    mdbName = SSProcess.GetProjectFileName()
    
    strOutputPath = Replace(mdbName, ".edb", "\")
    
    fileName = "�±�" & ".mdb"
    
    If FolderExist(strOutputPath) = False Then CreateFolder strOutputPath
    
    fileName = strOutputPath & "\" & fileName
    
    '���GDB
    ExportGDB fileName,SystemFileName,PrjPathName
    
    MsgBox "OK"
    
End Sub

SSProcess.AccessIsEOF mdbName, sql
Function AddOne( ByRef startIndex )
    startIndex = startIndex + 1
    AddOne = startIndex
End Function

Function ToChinese()
    SqlStr = "Select ���¹��������Ա�.ID,���¹��������Ա�.FSFS From ���¹��������Ա� Inner Join GeoLineTB on ���¹��������Ա�.ID = GeoLineTB.ID WHERE (GeoLineTB.Mark Mod 2)<>0"
    GetSQLRecordAll SqlStr,LineArr,LineCount
    For i = 0 To LineCount - 1
        SingleLineArr = Split(LineArr(i),",", - 1,1)
        If IsNumeric(SingleLineArr(1)) Then
            Select Case SingleLineArr(1)
                '���Զ��գ�0,1,2,3,4,5,6,7,8,9,10,11,12
                'ֱ��,����,�ܿ�,�ܹ�,�ܿ�,����,�ϼ�,Сͨ��,�ۺϹ��ȣ�����,�˷�,��������,����,ˮ��
                Case "0"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","ֱ��"
                Case "1"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","����"
                Case "2"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","�ܿ�"
                Case "3"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","�ܹ�"
                Case "4"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","�ܿ�"
                Case "5"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","����"
                Case "6"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","�ϼ�"
                Case "7"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","Сͨ��"
                Case "8"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","�ۺϹ��ȣ�����"
                Case "9"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","�˷�"
                Case "10"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","��������"
                Case "11"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","����"
                Case "12"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","ˮ��"
            End Select
        End If
    Next 'i
    
    SqlStr = "Select ���¹��������Ա�.ID,���¹��������Ա�.SJYL From ���¹��������Ա� Inner Join GeoLineTB on ���¹��������Ա�.ID = GeoLineTB.ID WHERE (GeoLineTB.Mark Mod 2)<>0"
    GetSQLRecordAll SqlStr,LineArr,LineCount
    For i = 0 To LineCount - 1
        SingleLineArr = Split(LineArr(i),",", - 1,1)
        If IsNumeric(SingleLineArr(1)) Then
            '���Զ��գ�0,1,2,3,4,5,6,7,8
            '��ѹ,��ѹA��,��ѹB��,�θ�ѹA��,�θ�ѹB��,��ѹ,��ѹA��,��ѹB��,��ѹ
            Select Case SingleLineArr(1)
                Case "0"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","��ѹ"
                Case "1"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","��ѹA��"
                Case "2"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","��ѹB��"
                Case "3"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","�θ�ѹA��"
                Case "4"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","�θ�ѹB��"
                Case "5"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","��ѹ"
                Case "6"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","��ѹA��"
                Case "7"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","��ѹB��"
                Case "8"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","��ѹ"
            End Select
        End If
    Next 'i
    
    SqlStr = "Select ���¹��������Ա�.ID,���¹��������Ա�.GC From ���¹��������Ա� Inner Join GeoLineTB on ���¹��������Ա�.ID = GeoLineTB.ID WHERE (GeoLineTB.Mark Mod 2)<>0"
    GetSQLRecordAll SqlStr,LineArr,LineCount
    For i = 0 To LineCount - 1
        SingleLineArr = Split(LineArr(i),",", - 1,1)
        If IsNumeric(SingleLineArr(1)) Then
            Select Case SingleLineArr(1)
                '���Զ��գ�0,1,2,3,4,5,6,7,8,9,10,11,12
                '���Ӹֹ�,�޷�ֹ�,�ҿ�������,��ī������,��������,�����ֹ�,PVC,PE��,ͭ,����,�ֽ�������,שʯ,����
                Case "0"
                SSProcess.SetObjectAttr SingleLineArr(0),"[GC]","���Ӹֹ�"
                Case "1"
                SSProcess.SetObjectAttr SingleLineArr(0),"[GC]","�޷�ֹ�"
                Case "2"
                SSProcess.SetObjectAttr SingleLineArr(0),"[GC]","�ҿ�������"
                Case "3"
                SSProcess.SetObjectAttr SingleLineArr(0),"[GC]","��ī������"
                Case "4"
                SSProcess.SetObjectAttr SingleLineArr(0),"[GC]","��������"
                Case "5"
                SSProcess.SetObjectAttr SingleLineArr(0),"[GC]","�����ֹ�"
                Case "6"
                SSProcess.SetObjectAttr SingleLineArr(0),"[GC]","PVC"
                Case "7"
                SSProcess.SetObjectAttr SingleLineArr(0),"[GC]","PE��"
                Case "8"
                SSProcess.SetObjectAttr SingleLineArr(0),"[GC]","ͭ"
                Case "9"
                SSProcess.SetObjectAttr SingleLineArr(0),"[GC]","����"
                Case "10"
                SSProcess.SetObjectAttr SingleLineArr(0),"[GC]","�ֽ�������"
                Case "11"
                SSProcess.SetObjectAttr SingleLineArr(0),"[GC]","שʯ"
                Case "12"
                SSProcess.SetObjectAttr SingleLineArr(0),"[GC]","����"
            End Select
        End If
    Next 'i
End Function' ToChinese

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


Function break
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Type", "==", "POINT,LINE,AREA"
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    For i = 0 To geocount - 1
        'д������Ϣ
        sid = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        '���Ŵ�ɢ��ʽ�� 0���Զ���ɢ���� 1�����ݱ�����趨��ɢ��
        SSProcess.ExplodeObj sid, 1, 1, "AfterExplodeObj"
    Next
    
End Function

Function AfterExplodeObj()
    'ȡ��ɢ����ID
    SSParameter.GetParameterINT "AfterExplodeObj", "ExplodeObjID", "0", objID
    objType = SSProcess.GetObjectAttr (objID, "SSObj_Type" )
    geoCode = SSProcess.GetObjectAttr (objID, "SSObj_Code" )
    layername = SSProcess.GetObjectAttr (objID, "SSObj_LayerName" )
    If layername = "GXFSLN" Then
        DATASOURCE = SSProcess.GetObjectAttr (objID, "[DATASOURCE]" )
        GXRQ = SSProcess.GetObjectAttr (objID, "[GXRQ]" )
        CHDW = SSProcess.GetObjectAttr (objID, "[CHDW]" )
        DMGC = SSProcess.GetObjectAttr (objID, "[DMGC]" )
    End If
    geoCount = SSProcess.GetSelGeoCount
    For i = 0 To geoCount - 1
        geoID = SSProcess.GetSelGeoValue (i, "SSObj_ID" )
        SSProcess.SetObjectAttr geoID, "[DATASOURCE]", DATASOURCE
        SSProcess.SetObjectAttr geoID, "[GXRQ]", GXRQ
        SSProcess.SetObjectAttr geoID, "[CHDW]", CHDW
        SSProcess.SetObjectAttr geoID, "[DMGC]", DMGC
    Next
End Function

Function CreateFolder(ByVal strFolderPath)
    Dim FSO,OutProject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    strDir = Left(strFolderPath,3)
    strFolderPath = Right( strFolderPath, Len(strFolderPath) - 3 )
    mulFolderPath = Split(strFolderPath,"\")
    nCount = UBound(mulFolderPath)
    OutFolderPath = strDir
    For i = 0 To nCount - 1
        OutFolderPath = OutFolderPath & "\" & mulFolderPath(i)
        If FSO.FolderExists(OutFolderPath) = False Then
            Set OutProject = FSO.CreateFolder(OutFolderPath)
            Set OutProject = Nothing
        End If
    Next
    Set FSO = Nothing
End Function




Function ExportGDB(fileName,SystemFileName,PrjPathName)
    
    '-----------���ת������--------------
    SSProcess.ClearDataXParameter
    '-----------���û���ת������------------------
    SSProcess.SetDataXParameter "DataType", "22"
    SSProcess.SetDataXParameter "FeatureCodeTBName", "FeatureCodeTB_500"
    SSProcess.SetDataXParameter "SymbolScriptTBName", "SymbolScriptTB_500"
    SSProcess.SetDataXParameter "NoteTemplateTBName", "NoteTemplateTB_500"
    SSProcess.SetDataXParameter "ExportPathName", fileName
    SSProcess.SetDataXParameter "DataBoundMode", "0"
    SSProcess.SetDataXParameter "SymbolExplodeMode", "1"
    SSProcess.SetDataXParameter "AddSystemFieldMode", "0"
    SSProcess.SetDataXParameter "LayerUseStatus", "0"
    SSProcess.SetDataXParameter "EXCHANGE_PDB_ExportNoteMode", "0"
    SSProcess.SetDataXParameter "EXCHANGE_PDB_ExportEmptyLayer", "1"
    'SSProcess.SetDataXParameter "EXCHANGE_PDB_ExportShapeMode", "1"
    SSProcess.SetDataXParameter "EXCHANGE_PDB_ExportShapeMode", "1"
    SSProcess.SetDataXParameter "FormatAttrValueStatus","1"
    'SSProcess.SetDataXParameter"EXCHANGE_PDB_ExportShortDate","1"
    SSProcess.SetDataXParameter "EXCHANGE_PDB_PrjFile", PrjPathName & SystemFileName
    
    SSProcess.SetDataXParameter"EXCHANGE_PDB_SpatialRF_MinX","190672.27475"
    SSProcess.SetDataXParameter"EXCHANGE_PDB_SpatialRF_MinY","1001661.35705"
    SSProcess.SetDataXParameter"EXCHANGE_PDB_SpatialRF_MaxX","808944.90075"
    SSProcess.SetDataXParameter"EXCHANGE_PDB_SpatialRF_MaxY","6219933.98305"
    
    startIndex = 0
    SSProcess.SetDataXParameter "ExportLayerCount", "100"
    ' SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "GXSSPT"
    ' SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "GXSSLN"
    ' SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "GXSSPY"
    ' SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "GXYJPT"
    ' SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "GXFSLN"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "JS"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "ZS"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "PS"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "YS"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "WS"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "DL"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "GD"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "LD"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "DC"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "XH"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "TX"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "DX"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "YD"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "LT"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "JX"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "JK"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "EX"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "DS"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "BZ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "RQ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "MQ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "TR"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "YH"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "RL"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "RS"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "ZQ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "BM"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "CD"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "CP"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "CT"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "CH"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "CY"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "CQ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "CS"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "QT"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "FQ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "XF"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "FS"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "SY"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)), "GS"
    
    
    
    startIndex = 0
    SSProcess.SetDataXParameter "LayerRelationCount", "100"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JS:JSPOINT,��ˮ���ߣ��㣩:JSLINE,��ˮ����:,::"
    ' SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JY:POINT,���㣩:LINE,���ߣ�:,::"
    ' SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "ZS:POINT,���㣩:LINE,���ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "PS:PSPOINT,��ˮ���ߣ��㣩:PSLINE,��ˮ���ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YS:YSPOINT,��ˮ���ߣ��㣩:YSLINE,��ˮ���ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "WS:WSPOINT,��ˮ���ߣ��㣩:WSLINE,��ˮ���ߣ��ߣ�:,::"
    ' SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "HS:HSPOINT,���㣩:HSLINE,���ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DL:DLPOINT,�������ߣ��㣩:DLLINE,�������ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "GD:GDPOINT,������ߣ��㣩:GDLINE,������ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "LD:LDPOINT,·�ƹ��ߣ��㣩:LDLINE,·�ƹ��ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DC:DCPOINT,�糵���ߣ��㣩:DCLINE,�糵���ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "XH:XHPOINT,��ͨ�źŹ��ߣ��㣩:XHLINE,��ͨ�źŹ��ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "TX:TXPOINT,�ۺϹ��ߣ��㣩:TXLINE,�ۺϹ��ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DX:DXPOINT,���Ź��ߣ��㣩:DXLINE,���Ź��ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YD:YDPOINT,�ƶ����ߣ��㣩:YDLINE,�ƶ����ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "LT:LTPOINT,��ͨ���ߣ��㣩:LTLINE,��ͨ���ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JX:JXPOINT,���ù��ߣ��㣩:JXLINE,���ù��ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JK:JKPOINT,��ع��ߣ��㣩:JKLINE,��ع��ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "EX:EXPOINT,����ͨѶ���ߣ��㣩:EXLINE,����ͨѶ���ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DS:DSPOINT,�㲥���ӹ��ߣ��㣩:DSLINE,�㲥���ӹ��ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "BZ:BZPOINT,����ר�ù��ߣ��㣩:BZLINE,�籣��ר�ù��ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RQ:RQPOINT,ȼ�����ߣ��㣩:RQLINE,ȼ�����ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "MQ:MQPOINT,ú�����ߣ��㣩:MQLINE,ú�����ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "TR:TRPOINT,��Ȼ�����ߣ��㣩:TRLINE,��Ȼ�����ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YH:YHPOINT,Һ�������ߣ��㣩:YHLINE,Һ�������ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RL:RLPOINT,�������ߣ��㣩:RLLINE,�������ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RS:RSPOINT,��ˮ���ߣ��㣩:RSLINE,��ˮ���ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "ZQ:ZQPOINT,�������ߣ��㣩:ZQLINE,�������ߣ��ߣ�:,::"
    ' SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "GY:POINT,���㣩:LINE,���ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "BM:BMPOINT,�������ߣ��㣩:BMLINE,�������ߣ��ߣ�:,::"
    ' SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "ZH:POINT,���㣩:LINE,���ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CD:CDPOINT,���������ߣ��㣩:CDLINE,���������ߣ��ߣ�:,::"
    ' SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CP:POINT,���㣩:LINE,���ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CT:CTPOINT,����ͨ�ţ��㣩:CTLINE,����ͨ�ţ��ߣ�:,::"
    ' SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CH:POINT,���㣩:LINE,���ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CY:CYPOINT,�����ܵ����ߣ��㣩:CYLINE,�����ܵ����ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CQ:CQPOINT,��Ȼ�����ܵ����ߣ��㣩:CQLINE,��Ȼ�����ܵ����ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CS:CSPOINT,ˮ���ܵ����ߣ��㣩:CSLINE,ˮ���ܵ����ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "QT:QTPOINT,�������ܵ����ߣ��㣩:QTLINE,�������ܵ����ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "FQ:FQPOINT,�������ߣ��㣩:FQLINE,�������ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "XF:XFPOINT,����ˮ���ߣ��㣩:XFLINE,����ˮ���ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "FS:FSPOINT,�����ˮ���ߣ��㣩:FSLINE,�����ˮ���ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "SY:SYPOINT,ʯ�͹��ߣ��㣩:SYLINE,ʯ�͹��ߣ��ߣ�:,::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "GS:GSPOINT,��ҵ��ˮ���ߣ��㣩:GSLINE,��ҵ��ˮ���ߣ��ߣ�:,::"
    
    startIndex = 0
    SSProcess.SetDataXParameter "TableFieldDefCount", "100000"
    
    
    
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,FCODE,FCODE,Ҫ�ط������,byname,,dbText:1,10,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,PCDYBH,PCDYBH,�ղ鵥Ԫ���,,,dbText:1,11,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,FEATUREID,FEATUREID,��ʩ����,,,dbText,14,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,SSWZ,SSWZ,��ʩλ��,,,dbText,128,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,GXDDH,GXDDH,���ߵ���,,,dbText:1,17,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,TZ,TZ,����,,,dbText,20,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,FSW,FSW,������,,,dbText,16,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,DMGC,DMGC,����߳�,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,ORIENTATION,ORIENTATION,����,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,PXJW,PXJW,ƫ�ľ�λ,,,dbText,20,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,DATASOURCE,DATASOURCE,����Դ,,,dbText:1,30,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,CHDW,CHDW,��浥λ,,,dbText:1,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,GXRQ,GXRQ,��������,,,dbText:1,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPT,0,BZ,BZ,��ע,,,dbText,255,0"
    
    
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,FCODE,FCODE,Ҫ�ط������,byname,,dbText:1,10,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,PCDYBH,PCDYBH,�ղ鵥Ԫ���,,,dbText:1,11,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,FEATUREID,FEATUREID,��ʩ����,,,dbText:1,14,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SSMC,SSMC,��ʩ����,,,dbText,64,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SSWZ,SSWZ,��ʩλ��,,,dbText,128,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,ZFZGBM,ZFZGBM,�������ܲ���,,,dbText:1,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,YGDW,YGDW,�˹ܵ�λ,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,QSDW,QSDW,Ȩ����λ,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,TXJYDW,TXJYDW,����Ӫ��λ,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,JSDW,JSDW,���赥λ,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SJDW,SJDW,��Ƶ�λ,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,KCDW,KCDW,���쵥λ,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SGDW,SGDW,ʩ����λ,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,JCNY,JCNY,��������,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,KSSYNY,KSSYNY,��ʼʹ������,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SJBCSJ,SJBCSJ,��Ʊ���ʱ��,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GXDBH,GXDBH,���߶α��,,,dbText:1,35,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GXQDDH,GXQDDH,���������,,,dbText:1,17,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GXZDDH,GXZDDH,�����յ���,,,dbText:1,17,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GXQDMS,GXQDMS,�����������,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GXZDMS,GXZDMS,�����յ�����,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GXQDDMGC,GXQDDMGC,����������߳�,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GXZDDMGC,GXZDDMGC,�����յ����߳�,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GXQDGDGC,GXQDGDGC,�������ܵ��߳�,,,dbDouble:1,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GXZDGDGC,GXZDGDGC,�����յ�ܵ��߳�,,,dbDouble:1,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GC,GC,�ܲ�,,,dbText:1,16,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SSJZ,SSJZ,���ͽ���,,,dbText,16,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SJWD,SJWD,����¶�,,,dbText,16,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GDBWCL,GDBWCL,�ܵ����²���,,,dbText,32,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GBHD,GBHD,�ܱں��,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GWYL,GWYL,����ѹ��,,,dbinteger,1,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SJYL,SJYL,���ѹ��,,,dbText,16,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,DYZ,DYZ,��ѹֵ,,,dbText,16,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,LL,LL,����,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,LX,LX,����,,,dbText,16,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,FSFS,FSFS,���跽ʽ,,,dbText:1,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GJ,GJ,�ܾ���DN��,,,dbText,16,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,DMCC,DMCC,����ߴ�,,,dbText,16,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GDGCWZ,GDGCWZ,�ܵ��߳�λ��,,,dbinteger:1,1,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,XLTS,XLTS,��������,,,dbLong,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,ZKS,ZKS,�ܿ���,,,dbLong,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,YYKS,YYKS,���ÿ���,,,dbLong,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,KJ,KJ,�׾�,,,dbText,16,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GXZT,GXZT,����״̬,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GDJKXS,GDJKXS,�ܵ��ӿ���ʽ,,,dbText,16,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GXAFXS,GXAFXS,���߰�����ʽ,,,dbText,16,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SFMYSZX,SFMYSZX,�Ƿ�����ʾ����,,,dbInteger,1,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,DJQK,DJQK,�ػ����,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,JCXS,JCXS,������ʽ,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SJDXSW,SJDXSW,��Ƶ���ˮλ,,,dbText,20,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,XKDXSW,XKDXSW,�ֿ�����ˮλ,,,dbText,20,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,DXSSFYFSX,DXSSFYFSX,����ˮ�Ƿ��и�ʴ��,,,dbInteger,1,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SGFS,SGFS,ʩ����ʽ,,,dbText,32,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SJSYNX,SJSYNX,���ʹ������,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,JGSJAQDJ,JGSJAQDJ,�ṹ��ư�ȫ�ȼ�,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,KZSFLD,KZSFLD,��������Ҷ�,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,KZSFLB,KZSFLB,����������,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,DMHZSJBZ,DMHZSJBZ,���������Ʊ�׼,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SFCYDZDLD,SFCYDZDLD,�Ƿ��ڵ�����Ѵ�,,,dbInteger,1,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SFCZBLDZ,SFCZBLDZ,�Ƿ���ڲ�������,,,dbInteger,1,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,SFCYQBSCZ,SFCYQBSCZ,�Ƿ���ǳ��ɰ����,,,dbInteger,1,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,MZGDWGJC,MZGDWGJC,��װ�ܵ���ۼ��,,,dbText,32,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,DATASOURCE,DATASOURCE,����Դ,,,dbText:1,30,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,GXRQ,GXRQ,��������,,,dbText:1,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,CHDW,CHDW,��浥λ,,,dbText:1,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,BZ,BZ,��ע,,,dbText,255,0"
    ' 'SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSLN,1,FJ,FJ,����,,,dbText,255,0"
    
    
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,FCODE,FCODE,Ҫ�ط������,byname,,dbText:1,10,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,PCDYBH,PCDYBH,�ղ鵥Ԫ���,,,dbText:1,11,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,FEATUREID,FEATUREID,��ʩ����,,,dbText:1,14,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,SSWZ,SSWZ,��ʩλ��,,,dbText,128,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,SSMC,SSMC,��ʩ����,,,dbText,64,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,ZFZGBM,ZFZGBM,�������ܲ���,,,dbText:1,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,YGDW,YGDW,�˹ܵ�λ,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,QSDW,QSDW,Ȩ����λ,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,TXJYDW,TXJYDW,����Ӫ��λ,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,JSDW,JSDW,���赥λ,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,SJDW,SJDW,��Ƶ�λ,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,KCDW,KCDW,���쵥λ,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,SGDW,SGDW,ʩ����λ,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,JCNY,JCNY,��������,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,KSSYNY,KSSYNY,��ʼʹ������,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,SJBCSJ,SJBCSJ,��Ʊ���ʱ��,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,SJSYNX,SJSYNX,���ʹ������,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,QDDMGC,QDDMGC,������߳�,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,ZDDMGC,ZDDMGC,�յ����߳�,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,QDJGDBDMGC,QDJGDBDMGC,���ṹ���嶥��߳�,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,ZDJGDBDMGC,ZDJGDBDMGC,�յ�ṹ���嶥��߳�,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,QDDBFTHD,QDDBFTHD,��㶥�帲�����,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,ZDDBFTHD,ZDDBFTHD,�յ㶥�帲�����,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,GLCSSL,GLCSSL,���Ȳ�������,,,dbText:1,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,GLNYXGXZL,GLNYXGXZL,���������й�������,,,dbText,32,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,DJQK,DJQK,�ػ����,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,JGXS,JGXS,�ṹ��ʽ,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,ZJFW,ZJFW,ע����Χ,,,dbText,32,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,SJDXSW,SJDXSW,��Ƶ���ˮλ,,,dbText,20,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,XKDXSW,XKDXSW,�ֿ�����ˮλ,,,dbText,20,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,DXSSFYFSX,DXSSFYFSX,����ˮ�Ƿ��и�ʴ��,,,dbInteger,1,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,SGFS,SGFS,ʩ����ʽ,,,dbText,32,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,JGSJAQDJ,JGSJAQDJ,�ṹ��ư�ȫ�ȼ�,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,KZSFLD,KZSFLD,��������Ҷ�,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,KZSFLB,KZSFLB,����������,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,DMHZSJBZ,DMHZSJBZ,���������Ʊ�׼,,,dbText,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,SFCYDZDLD,SFCYDZDLD,�Ƿ��ڵ�����Ѵ�,,,dbInteger,1,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,SFCZBLDZ,SFCZBLDZ,�Ƿ���ڲ�������,,,dbInteger,1,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,SFCYQBSCZ,SFCYQBSCZ,�Ƿ���ǳ��ɰ����,,,dbInteger,1,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,DATASOURCE,DATASOURCE,����Դ,,,dbText:1,30,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,GXRQ,GXRQ,��������,,,dbText:1,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,CHDW,CHDW,��浥λ,,,dbText:1,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,BZ,BZ,��ע,,,dbText,255,0"
    ' 'SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXSSPY,2,FJ,FJ,����,,,dbText,255,0"
    
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,FCODE,FCODE,Ҫ�ط������,byname,,dbText:1,10,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,PCDYBH,PCDYBH,�ղ鵥Ԫ���,,,dbText:1,11,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,FEATUREID,FEATUREID,��ʩ����,,,dbText:1,14,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,SSWZ,SSWZ,��ʩλ��,,,dbText,128,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,ZFZGBM,ZFZGBM,�������ܲ���,,,dbText:1,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,YGDW,YGDW,�˹ܵ�λ,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,QSDW,QSDW,Ȩ����λ,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,TXJYDW,TXJYDW,����Ӫ��λ,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,JSDW,JSDW,���赥λ,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,SJDW,SJDW,��Ƶ�λ,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,KCDW,KCDW,���쵥λ,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,SGDW,SGDW,ʩ����λ,,,dbText,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,GXDDH,GXDDH,���ߵ���,,,dbText:1,17,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,TZ,TZ,����,,,dbText,20,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,FSW,FSW,������,,,dbText,16,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,DMGC,DMGC,����߳�,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,ORIENTATION,ORIENTATION,����,,,dbDouble,12,3"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,JGXZ,JGXZ,������״,,,dbText:1,20,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,JGZJHDMCC,JGZJHDMCC,����ֱ�������ߴ�,,,dbText,20,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,JGCZ,JGCZ,���ǲ���,,,dbText:1,20,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,JGXS,JGXS,�ṹ��ʽ,,,dbText,20,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,JBS,JBS,������,,,dbDouble,12,2"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,JS,JS,����,,,dbDouble,12,2"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,JBCC,JBCC,����ֱ�������ߴ�,,,dbText,20,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,JCC,JCC,��ֱ�������ߴ�,,,dbText,20,20"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,SFYAQW,SFYAQW,�Ƿ��а�ȫ��,,,dbInteger,1,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,DATASOURCE,DATASOURCE,����Դ,,,dbText:1,30,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,GXRQ,GXRQ,��������,,,dbText:1,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,CHDW,CHDW,��浥λ,,,dbText:1,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXYJPT,0,BZ,BZ,��ע,,,dbText,255,0"
    
    
    
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXFSLN,1,FCODE,FCODE,Ҫ�ط������,byname,,dbText:1,10,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXFSLN,1,PCDYBH,PCDYBH,�ղ鵥Ԫ���,,,dbText:1,11,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXFSLN,1,DATASOURCE,DATASOURCE,����Դ,,,dbText:1,30,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXFSLN,1,GXRQ,GXRQ,��������,,,dbText:1,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXFSLN,1,CHDW,CHDW,��浥λ,,,dbText:1,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXFSLN,1,BZ,BZ,��ע,,,dbText,255,0"
    
    
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXFSLN,2,FCODE,FCODE,Ҫ�ط������,byname,,dbText:1,10,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXFSLN,2,PCDYBH,PCDYBH,�ղ鵥Ԫ���,,,dbText:1,11,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXFSLN,2,DATASOURCE,DATASOURCE,����Դ,,,dbText:1,30,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXFSLN,2,GXRQ,GXRQ,��������,,,dbText:1,8,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXFSLN,2,CHDW,CHDW,��浥λ,,,dbText:1,60,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GXFSLN,2,BZ,BZ,��ע,,,dbText,255,0"
    
    
    
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,WTDH,Exp_No,��̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,WTTSDH,Map_NO,ͼ�ϵ��,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,DMGC,Surf_H,����߳�,,,dbDouble,8,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,TZ,Feature,����,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,FSW,Subsid,������,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,PXJW,Offset,ƫ�ľ�λ,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,ORIENTATION,Angle,��X�᷽��Ϊ0��,ȡ��ʱ�뷽��,,,dbDouble,5,1"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,MapNumber,MapNum,ͼ����,MapNumber,,dbText,15,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,X,X,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,Y,Y,������,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,MapNo_X,MapNo_X,λ�ƺ��ͼ�ϵ��λ��X����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,MapNo_Y,MapNo_Y,λ�ƺ��ͼ�ϵ��λ��y����,,,dbDouble,12,3"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,JCNY,Mdata,�������,,,dbDate,,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,QSDW,B_Code,Ȩ����λ���������ֶ�,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,SSWZ,Road,·�����������ֶ�,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,0,XMMC,PID,���̺�,,,dbText,20,0"
    
    '����˳�򣺲���,����,EPS�ֶ���,�ͻ��ֶ���,�ͻ��ֶα���,ϵͳ�ֶ���,ȱʡֵ,�ֶ�����,�ֶγ���,С��λ
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CD,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CT,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CY,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,ZKS,Hole_Num,�ܿ���,,,dbLong,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CQ,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "CS,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "QT,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BM,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FQ,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DL,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GD,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LD,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DC,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XH,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TX,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DX,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YD,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "LT,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JX,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JK,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "EX,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "DS,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "BZ,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "JS,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "XF,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "PS,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YS,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "WS,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "FS,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RQ,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "MQ,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "TR,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "YH,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RL,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "RS,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "ZQ,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "SY,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,GXQDDH,S_point,�����̽���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,GXZDDH,E_point,���ӷ���,,,dbText,12,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,GXQDMS,S_Deep,�������,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,GXZDMS,E_Deep,�յ�����,,,dbDouble,4,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,GC,Material,����,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,DMCC,D_S,�ܾ���ܿ�,,,dbText,16,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,JCNY,Mdata,�������,,,dbDate,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,XLTS,Cab_Count,��������,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,DYZ,Voltage,��ѹֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,ZKS,Hole_Num,�ܿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,LX,FlowDirect,��ˮ����,,,dbInteger,1,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,GXLX,P_Type,��������,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,BZ,Memo,��ע,,,dbText,20,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,FSFS,D_Type,��������,,,dbText,8,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,QSDW,B_Code,Ȩ����λ����,,,dbText,60,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,SJYL,Pressure,ѹ��ֵ,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,WYKS,Hole_Used,δ�ÿ���,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,D_Dia,D_Dia,�׹ܳߴ�,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,SSWZ,Road,·������,,,dbText,128,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)), "GS,1,XMBH,PID,���̺�,,,dbText,20,0"
    
    SSProcess.ExportData
    
End Function

Function ScanString(ByVal str, ByVal sep, ByRef strs(), ByRef count)
    Dim sepidx1, sepidx2, strtemp
    count = 0
    sepidx1 = 1
    sepidx2 = InStr(sepidx1 , str, sep, 1)
    While (sepidx2 > 0)
        strs(count) = Mid( str, sepidx1, sepidx2 - sepidx1)
        sepidx1 = sepidx2 + 1
        sepidx2 = InStr(sepidx1, str, sep, 1)
        count = count + 1
    WEnd
    strs(count) = Mid( str, sepidx1, Len(str) + 1 - sepidx1)
    count = count + 1
End Function


Function AngleFz
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Type", "==", "POINT"
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    If geocount > 0 Then
        For i = 0 To geocount - 1
            id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
            code = SSProcess.GetSelGeoValue(i, "SSObj_Code")
            FX = SSProcess.GetSelGeoValue(i, "[ORIENTATION]")
            angle = SSProcess.GetObjectAttr(id, "SSObj_Angle")
            angle = SSProcess.AdjustAngle(angle)
            Pi = 3.1415926
            angle1 = SSProcess.RadianToDms(angle)
            If FX = "" Then
                If angle1 < 90 And angle1 > 0 Then
                    angle = SSProcess.RadianToDms  (Pi / 2 - angle)
                    angle = Left(angle,InStr(angle,".") + 3)
                    SSProcess.SetObjectAttr id, "[ORIENTATION]", angle
                End If
                If angle1 > 90 Or angle1 = 90 Then
                    angle = SSProcess.RadianToDms  ((5 * Pi / 2) - angle)
                    angle = Left(angle,InStr(angle,".") + 3)
                    SSProcess.SetObjectAttr id, "[ORIENTATION]", angle
                End If
            ElseIf FX <> "" And angle <> FX Then
                If angle1 < 90 And angle1 > 0 Then
                    angle = SSProcess.RadianToDms  (Pi / 2 - angle)
                    angle = Left(angle,InStr(angle,".") + 3)
                    SSProcess.SetObjectAttr id, "[ORIENTATION]", angle
                End If
                If angle1 > 90 Or angle1 = 90 Then
                    angle = SSProcess.RadianToDms  ((5 * Pi / 2) - angle)
                    angle = Left(angle,InStr(angle,".") + 3)
                    SSProcess.SetObjectAttr id, "[ORIENTATION]", angle
                End If
            End If
        Next
    End If
    
End Function

Function DateRq
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Type", "==", "Point,Line,AREA"
    SSProcess.SelectFilter
    count = SSProcess.GetSelGeoCount
    If count > 0 Then
        For i = 0 To count - 1
            id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
            GXSJ = SSProcess.GetObjectAttr(id,"[GXRQ]")
            JCNY = SSProcess.GetObjectAttr(id,"[JCNY]")
            KSSYNY = SSProcess.GetObjectAttr(id,"[KSSYNY]")
            SJBCSJ = SSProcess.GetObjectAttr(id,"[SJBCSJ]")
            TCRQ = SSProcess.GetObjectAttr(id,"[TCRQ]")
            ZJYCDZXHGKJSJ = SSProcess.GetObjectAttr(id,"[ZJYCDZXHGKJSJ]")
            ZSSJ = SSProcess.GetObjectAttr(id,"[ZSSJ]")
            PCRQ = SSProcess.GetObjectAttr(id,"[PCRQ]")
            If GXSJ <> "" Then
                b = InStr(GXSJ," ")
                If b = 0 Then  SSProcess.SetObjectAttr id, "[GXSJ]", CDate(GXSJ)
                If b <> 0 Then  SSProcess.SetObjectAttr id, "[GXSJ]", CDate(Left(GXSJ,b - 1))
            End If
            If JCNY <> "" Then
                b = InStr(JCNY," ")
                If b = 0 Then  SSProcess.SetObjectAttr id, "[JCNY]", CDate(JCNY)
                If b <> 0 Then  SSProcess.SetObjectAttr id, "[JCNY]", CDate(Left(JCNY,b - 1))
            End If
            If KSSYNY <> "" Then
                b = InStr(KSSYNY," ")
                If b = 0 Then  SSProcess.SetObjectAttr id, "[KSSYNY]", CDate(KSSYNY)
                If b <> 0 Then  SSProcess.SetObjectAttr id, "[KSSYNY]", CDate(Left(KSSYNY,b - 1))
            End If
            If SJBCSJ <> "" Then
                b = InStr(SJBCSJ," ")
                If b = 0 Then  SSProcess.SetObjectAttr id, "[SJBCSJ]", CDate(SJBCSJ)
                If b <> 0 Then  SSProcess.SetObjectAttr id, "[SJBCSJ]", CDate(Left(SJBCSJ,b - 1))
            End If
            If TCRQ <> "" Then
                b = InStr(TCRQ," ")
                If b = 0 Then  SSProcess.SetObjectAttr id, "[TCRQ]", CDate(TCRQ)
                If b <> 0 Then  SSProcess.SetObjectAttr id, "[TCRQ]", CDate(Left(TCRQ,b - 1))
            End If
            If ZJYCDZXHGKJSJ <> "" Then
                b = InStr(ZJYCDZXHGKJSJ," ")
                If b = 0 Then  SSProcess.SetObjectAttr id, "[ZJYCDZXHGKJSJ]", CDate(ZJYCDZXHGKJSJ)
                If b <> 0 Then  SSProcess.SetObjectAttr id, "[ZJYCDZXHGKJSJ]", CDate(Left(ZJYCDZXHGKJSJ,b - 1))
            End If
            If ZSSJ <> "" Then
                b = InStr(ZSSJ," ")
                If b = 0 Then  SSProcess.SetObjectAttr id, "[ZSSJ]", CDate(ZSSJ)
                If b <> 0 Then  SSProcess.SetObjectAttr id, "[ZSSJ]", CDate(Left(ZSSJ,b - 1))
            End If
            If PCRQ <> "" Then
                b = InStr(PCRQ," ")
                If b = 0 Then  SSProcess.SetObjectAttr id, "[PCRQ]", CDate(PCRQ)
                If b <> 0 Then  SSProcess.SetObjectAttr id, "[PCRQ]", CDate(Left(PCRQ,b - 1))
            End If
        Next
    End If
End Function

Function SmallNumber
    mdbName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb mdbName
    sql = "update  ���¹��ߵ����Ա� set  DMGC=Round(DMGC,3),ORIENTATION=Round(ORIENTATION,3),JBS=Round(JBS,2),ORIENTATION=Round(JS,2)"
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  ���¹��������Ա� set  GXZDDMGC = round(GXZDDMGC,3),GXQDMS=Round(GXQDMS,3) ,GXZDMS=Round(GXZDMS,3) ,GXQDDMGC=Round(GXQDDMGC ,3),GXQDGDGC=Round(GXQDGDGC ,3),GXZDGDGC=Round(GXZDGDGC ,3),LL=Round(LL ,3)"
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  ���¹��������Ա� set  QDDMGC=Round(QDDMGC,3),ZDDMGC=Round(ZDDMGC,3),QDJGDBDMGC=Round(QDJGDBDMGC,3),ZDJGDBDMGC=Round(ZDJGDBDMGC,3),QDDBFTHD=Round(QDDBFTHD ,3),ZDDBFTHD=Round(ZDDBFTHD ,3)"
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  SZDC_����ͨ�������Ա� set ORIENTATION=Round(ORIENTATION,3) "
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  SZDC_����ͨ�������Ա� set  QDDMGC=Round(QDDMGC,3),ZDDMGC=Round(ZDDMGC,3),QDJGDBDMGC=Round(QDJGDBDMGC,3),ZDJGDBDMGC=Round(ZDJGDBDMGC,3),QDDBFTHD=Round(QDDBFTHD ,3),ZDDBFTHD=Round(ZDDBFTHD ,3),CG=Round(CG,3),JZMJ=Round(JZMJ,3),SZCCG=Round(SZCCG,3),SZCJZMJ=Round(SZCJZMJ,3),LFKDA=Round(LFKDA,3),LFKDA1=Round(LFKDA1,3) ,LFKDA2=Round(LFKDA2,3) ,LF=Round(LF,3),PS=Round(PS,3),BJYCJ=Round(BJYCJ,3)"
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  SZDC_����ͣ���������Ա� set  CG=Round(CG,3),SZCCG=Round(SZCCG,3),SZCJZMJ=Round(SZCJZMJ,3),DMGC=Round(DMGC,3),SJJZMJ=Round(SJJZMJ, 3),JGDBDMGC=Round(JGDBDMGC ,3),DBFTHD=Round(DBFTHD,3) "
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  SZDC_�˷����̵����Ա� set ORIENTATION=Round(ORIENTATION,3) "
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  SZDC_�˷����������Ա� set  CG=Round(CG,3),SZCCG=Round(SZCCG,3),JZMJ=Round(JZMJ,3),YJMJ=Round(YJMJ,3),YBMJ=Round(YBMJ,3),SZCYBMJ=Round(SZCYBMJ, 3),JGDBDMGC=Round(JGDBDMGC ,3),DBFTHD=Round(DBFTHD,3) ,DMGC=Round(DMGC,3),SZCJZMJ=Round(SZCJZMJ,3)"
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  SZDC_���ºӵ������Ա� set  ORIENTATION=Round(ORIENTATION,3)"
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  SZDC_���ºӵ������Ա� set  QDDMGC=Round(QDDMGC,3),ZDDMGC=Round(ZDDMGC,3)"
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  SZDC_�������̵����Ա� set  DMGC=Round(DMGC,3),ORIENTATION=Round(ORIENTATION,3),JBS=Round(JBS,2),ORIENTATION=Round(JS,2)"
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  SZDC_�������������Ա� set  GXQDMS=Round(GXQDMS,3) ,GXZDMS=Round(GXZDMS,3) ,GXQDDMGC=Round(GXQDDMGC ,3),GXQDGDGC=Round(GXQDGDGC ,3),GXZDGDGC=Round(GXZDGDGC ,3)"
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  SZDC_�������������Ա� set  DMGC=Round(DMGC,3),JGDBDMGC=Round(JGDBDMGC ,3),DBFTHD=Round(DBFTHD,3) "
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  SZDC_�������¿ռ���ʩ�����Ա� set ORIENTATION=Round(ORIENTATION,3) "
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  SZDC_�������¿ռ���ʩ�����Ա� set  DMGC=Round(DMGC,3),JGDBDMGC=Round(JGDBDMGC ,3),DBFTHD=Round(DBFTHD,3),CG=Round(CG,3),SZCCG=Round(SZCCG,3),SZCJZMJ=Round(SZCJZMJ,3)"
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  SZDC_�ղ鵥Ԫ��Ϣ���Ա� set  XLF=Round(XLF,3),WLF=Round(WLF,3),TP=Round(TP,3),CZBL=Round(CZBL,3),MM=Round(MM,3),GL=Round(GL,3),SS=Round(SS,3),CX=Round(CX,3),KC=Round(KC,3),FJ=Round(FJ,3),DB=Round(DB,3),KD=Round(KD,3),BDTK=Round(BDTK,3),PSui=Round(PSui,3),CT=Round(CT,3),PSun=Round(PSun,3),GQRQ=Round(GQRQ,3),TFLSH=Round(TFLSH,3),JCJXC=Round(JCJXC,3),JBLMSH=Round(JBLMSH,3)"
    SSProcess.ExecuteAccessSql  mdbName,sql
    SSProcess.CloseAccessMdb mdbName
End Function


'///����ļ���������ָ���ļ�
Function GetAllFiles(ByRef pathname, ByRef fileExt, ByRef filecount, ByRef filenames())
    Dim fso, folder, file, files, subfolder,folder0, fcount
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(pathname)
    Set files = folder.Files
    '�����ļ�
    For Each file In files
        extname = fso.GetExtensionName(file.name)
        If UCase(extname) = UCase(fileExt) Then
            filenames(filecount) = pathname & file.name
            filecount = filecount + 1
        End If
    Next
End Function

#include "Function_beforeExportProcessFunc.vbs"

'********====<�ж��ļ����Ƿ����>==========&&&&&&&&&&
Function FolderExist(FolderName)
    FolderExist = True
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.folderExists(FolderName)  Then
        FolderExist = False
    End If
    Set fso = Nothing
End Function

Function CreateFolder(ByVal strFolderPath)
    Set Fso = CreateObject("Scripting.FileSystemObject")
    strDir = Left(strFolderPath,3)
    strFolderPath = Right( strFolderPath, Len(strFolderPath) - 3 )
    mulFolderPath = Split(strFolderPath,"\")
    nCount = UBound(mulFolderPath)
    strDirPath = strDir
    For i = 0 To nCount - 1
        strDirPath = strDirPath & "\" & mulFolderPath(i)
        If Fso.FolderExists(strDirPath) = False Then
            Fso.CreateFolder(strDirPath)
        End If
    Next
    Set Fso = Nothing
End Function