'ControlIDS���ؼ�ID������ʽ��GeoFields����Ӧ������ֶ���������ʽ��DefaultValues��Ĭ��ֵ������ʽ��AlternativeValues����ѡ������������ʽ��MemoryValues���Ƿ�����ϴδ����ֵ������ʽ��ControlCount ���ؼ�����
Dim ConIDS(300),GFields(300),DefValues(300),AltValues(300), MemValues(300),ConCount,ZDID
Dim dlgHandle, dlgHandle1, dlgHandle2, g_scriptHandle

mdbName = SSProcess.GetProjectFileName
'��¼����Ҫ��д���ڵ�ID
ZDID = ""
ZDCode = "9130223"
Sub OnInitFreeScript(scriptHandle)
    mapHandle = SSProject.GetActiveMap
    mapType = SSProject.GetMapInfo(mapHandle, "MapType")
    If mapType <> 2 Then
        MsgBox "������ֻ֧���ڵ���ͼ����ִ�У�"
        Exit Sub
    End If
    
    g_scriptHandle = scriptHandle
    dlgHandle = SSProcess.CreateFreeScriptDlg(scriptHandle, 1)
    Rstate = GetZDID()  '��ȡҪ�����Ե��ڵ�ID,
    If Rstate = 0 Then
        MsgBox "ͼ�����ڵ��棬����ʹ�øù���X"
        SSProcess.CloseScriptDlg
        Exit Sub
    End If
    Dim strs(100)
    IniDlgParameter()  '��ȡ�Ի���Ini�ı����õĲ���
    mode = 1'0 ��ģ 1 ��ģ
    title = "��Ŀ��Ϣ¼��"
    dlgTemplateName = "��Ŀ��Ϣ¼�����Ի���"
    dlgWidth = 750
    dlgHeight = 610
    colCount = 0
    titleWidth = 0
    valueWidth = 0
    For i = 0 To ConCount - 1
        If MemValues(i) = "1" Then
            'DefValues(i) = SSProcess.ReadEpsIni ("ZhuHaiInfoPut", "CLInfo_" & ConIDS(i) , "" )
        End If
        If DefValues(i) = "Date" Then DefValues(i) = Date
        If ZDID <> "" Then
            SSFunc.ScanString GFields(i),",",strs,scount
            Valuestr = ""
            For j = 0 To scount - 1
                If j = 0 Then Valuestr = SSProcess.GetObjectAttr (ZDID,"[" & strs(j) & "]")
                If j > 0 And Valuestr <> "" Then Valuestr = Valuestr & "��" & SSProcess.GetObjectAttr (ZDID,"[" & strs(j) & "]")
            Next
            If Valuestr = "" Then
                SSProcess.AddInputParameter_ex dlgHandle,ConIDS(i),DefValues(i),0,AltValues(i),""
            Else
                SSProcess.AddInputParameter_ex dlgHandle,ConIDS(i),Valuestr,0,AltValues(i),""   '��ȡ��ǰ����ֵ���
            End If
        Else
            SSProcess.AddInputParameter_ex dlgHandle,ConIDS(i),DefValues(i),0,AltValues(i),""
        End If
    Next
    SSProcess.AddInputParameter_ex dlgHandle, "{��ǩ}", "��Ŀ��Ϣ", 0, "��Ŀ��Ϣ,�����������Ϣ", ""
    SSProcess.ShowFreeScriptDlg dlgHandle, title, dlgTemplateName, dlgWidth, dlgHeight, colCount, titleWidth, valueWidth, dockMode
    OnTabCtrlSelChange "", 0, "{��ǩ}", "��Ŀ��Ϣ"
    OnTabCtrlSelChange "", 0, "{��ǩ}", "�����������Ϣ"
    OnTabCtrlSelChange "", 0, "{��ǩ}", "��Ŀ��Ϣ"
End Sub

'��ȡҪ�༭���Ե��ڵ�ID������ȡѡ�񼯵��ڵأ�����ѡ��ȫͼ���ڵأ��ҵ�Ψһ���ڵأ���ȡZDID���ҵ�����ڵأ��������˹�ѡ��
Function GetZDID()
    '�ж�ѡ���Ƿ����ڵ�
    GetZDID = 1
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.UpdateSysSelection 0
    geoCount = SSProcess.GetSelGeoCount
    LSID = ""
    ZDCount = 0
    For i = 0 To geoCount - 1
        geoCode = SSProcess.GetSelGeoValue (i, "SSObj_Code")
        If  InStr("," & ZDCode & ",","," & geoCode & ",") > 0 Then
            LSID = SSProcess.GetSelGeoValue (i, "SSObj_ID")
            ZDCount = ZDCount + 1
        End If
    Next
    If ZDCount = 1 Then
        ZDID = LSID
    Else
        SSProcess.ClearSelection
        SSProcess.ClearSelectCondition
        SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
        SSProcess.SetSelectCondition "SSObj_Code", "=", ZDCode
        SSProcess.SelectFilter
        geoCount = SSProcess.GetSelGeoCount
        If geoCount = 0 Then
            GetZDID = 0
        ElseIf geoCount = 1 Then
            ZDID = SSProcess.GetSelGeoValue (0, "SSObj_ID")
        Else
            MsgBox "ͼ���ж���ڵأ��ҵ�ǰѡ����ѡ���ڵأ����ڵ����Ի����ѡ���ڵ�����д����X"
        End If
    End If
End Function

Sub OnExitFreeScript()
    '��Ӵ���
End Sub

Sub OnOK()
    '��Ӵ���
End Sub

'�༭��ʧȥ����
'Function OnEditKillFocus( tableName, objectID, fieldName, fieldValue )
'��Ӵ���
'End Function

Sub OnCancel()
    If dlgHandle1 <> 0 Then SSProcess.CloseChildFreeScriptDlg dlgHandle1
    If dlgHandle2 <> 0 Then SSProcess.CloseChildFreeScriptDlg dlgHandle2
End Sub

'��ǩ�ؼ�ѡ����ı�
Function OnTabCtrlSelChange( tableName, objectID, fieldName, fieldValue )
    '��Ӵ���
    Dim strs(100), mapnumberInfo(100000)
    If fieldName = "{��ǩ}" Then
        '�����ضԻ���
        If dlgHandle1 <> 0 Then SSProcess.ShowScriptDlgWindow dlgHandle1, 0
        If dlgHandle2 <> 0 Then SSProcess.ShowScriptDlgWindow dlgHandle2, 0
        If fieldValue = "��Ŀ��Ϣ" Then
            If dlgHandle1 = 0 Then
                dlgHandle1 = SSProcess.CreateFreeScriptDlg(g_scriptHandle, 1)
                dlgTemplateName = "��Ŀ��Ϣ¼��"
                dockCtrlID = "{ͣ���ؼ�}"
                SSProcess.DockScriptDlg dlgHandle1, dlgTemplateName, dlgHandle, dockCtrlID
                For i = 0 To ConCount - 1
                    If MemValues(i) = "1" Then
                        'DefValues(i) = SSProcess.ReadEpsIni ("ZhuHaiInfoPut", "CLInfo_" & ConIDS(i) , "" )
                    End If
                    If DefValues(i) = "Date" Then DefValues(i) = Date
                    If ZDID <> "" Then
                        SSFunc.ScanString GFields(i),",",strs,scount
                        Valuestr = ""
                        For j = 0 To scount - 1
                            If j = 0 Then Valuestr = SSProcess.GetObjectAttr (ZDID,"[" & strs(j) & "]")
                            If j > 0 And Valuestr <> "" Then Valuestr = Valuestr & "��" & SSProcess.GetObjectAttr (ZDID,"[" & strs(j) & "]")
                        Next
                        
                        If ConIDS(i) = "[JianZXS]" Then  SSProcess.SetScriptDlgCellOptions_ex dlgHandle1, "[JianZXS]",AltValues(i)
                        If ConIDS(i) = "[JianZJG]" Then  SSProcess.SetScriptDlgCellOptions_ex dlgHandle1, "[JianZJG]",AltValues(i)
                        '�����ֵ�
                        If Valuestr = ""  Then
                            'ͼ����
                            SSProcess.SetScriptDlgCellValue_ex dlgHandle1,ConIDS(i),DefValues(i)
                        Else
                            SSProcess.SetScriptDlgCellValue_ex dlgHandle1,ConIDS(i),Valuestr      '��ȡ��ǰ����ֵ���
                        End If
                    Else
                        SSProcess.SetScriptDlgCellValue_ex dlgHandle1,ConIDS(i),DefValues(i)
                    End If
                Next
            Else
                SSProcess.ShowScriptDlgWindow dlgHandle1, 1
            End If
        End If
        If fieldValue = "�����������Ϣ" Then
            If dlgHandle2 = 0 Then
                dlgHandle2 = SSProcess.CreateFreeScriptDlg(g_scriptHandle, 1)
                dlgTemplateName = "��Ŀ��Ϣ¼��-������"
                dockCtrlID = "{ͣ���ؼ�}"
                SSProcess.DockScriptDlg dlgHandle2, dlgTemplateName, dlgHandle, dockCtrlID
                For i = 0 To ConCount - 1
                    If MemValues(i) = "1" Then
                        'DefValues(i) = SSProcess.ReadEpsIni ("ZhuHaiInfoPut", "CLInfo_" & ConIDS(i) , "" )
                    End If
                    If DefValues(i) = "Date" Then DefValues(i) = Date
                    If ZDID <> "" Then
                        SSFunc.ScanString GFields(i),",",strs,scount
                        Valuestr = ""
                        For j = 0 To scount - 1
                            If j = 0 Then Valuestr = SSProcess.GetObjectAttr (ZDID,"[" & strs(j) & "]")
                            If j > 0 And Valuestr <> "" Then Valuestr = Valuestr & "��" & SSProcess.GetObjectAttr (ZDID,"[" & strs(j) & "]")
                        Next
                        
                        If Valuestr = "" Then
                            SSProcess.SetScriptDlgCellValue_ex dlgHandle2,ConIDS(i),DefValues(i)
                        Else
                            SSProcess.SetScriptDlgCellValue_ex dlgHandle2,ConIDS(i),Valuestr      '��ȡ��ǰ����ֵ���
                        End If
                    Else
                        SSProcess.SetScriptDlgCellValue_ex dlgHandle2,ConIDS(i),DefValues(i)
                    End If
                Next
            Else
                SSProcess.ShowScriptDlgWindow dlgHandle2, 1
            End If
        End If
    End If
End Function


Function GetXMFZR(strs_,strs1_)
    Dim  strs(1000)
    returnStr = ""
    str = ""
    fileName = SSProcess.GetSysPathName(7) & "���������.txt"
    Dim fso, ts, chLine
    Set fso = CreateObject("Scripting.FileSystemObject")'��Ҫ���ڴ��������Ӧ�ö����ű������ʵ����
    Set ts = fso.OpenTextFile(fileName, 1)' 1����ֻ����ʽ���ļ�������д����ļ���2����д��ʽ���ļ���8�����ļ������ļ�ĩβ��ʼд��
    Do While Not ts.AtEndOfStream'���λ���ļ�ĩ���򷵻� True�����򷵻� False��
        chLine = ts.ReadLine'��ȡһ����
        chLine = Trim(chLine)
        If str = ""  Then str = chLine
    Else str = str & "," & chLine
    Loop
    SSProcess.SaveBufferObjToDatabase
    ts.Close
    
    'str = "������3300300039,���½���3310300635,�����ࡢ3300300032"
    SSFunc.ScanString str,",",strs,strs1count
    ResVal_Dlg = SSFunc.SelectListAttr("ѡ���б�", "��ѡ�����б�", "ѡ�������б�", strs, strs1count)
    strs_ = ""
    strs1_ = ""
    If strs1count = 0 Then
        Exit Function
    End If
    If ResVal_Dlg = 1 Then
        If strs1count = 1 Then
            arinfo = Split(strs(0), "��")
            strs_ = arinfo(0)
            strs1_ = arinfo(1)
        Else
            MsgBox "��ѡ��"
            Exit Function
        End If
    End If
End Function

'���а�ť��ʵ��������
Function OnButtonClick( tableName, objectID, fieldName, fieldValue )
    '��Ӵ���
    If fieldName = "[XMFZR]" Then
        GetXMFZR  strs_,strs1 _
        If strs_ <> ""  Then SSProcess.SetScriptDlgCellValue_ex dlgHandle2 ,"[XiangMFZR]", strs _
        If strs1_ <> "" Then SSProcess.SetScriptDlgCellValue_ex dlgHandle2 ,"[XiangMFZRZSH]",strs1 _
    End If
    
    If fieldName = "[CLY]" Then
        GetXMFZR  strs_,strs1 _
        If strs_ <> ""  Then SSProcess.SetScriptDlgCellValue_ex dlgHandle2 ,"[CeLY]", strs _
        If strs1_ <> "" Then SSProcess.SetScriptDlgCellValue_ex dlgHandle2 ,"[CeLYZSH]",strs1 _
    End If
    If fieldName = "[ZTY]" Then
        GetXMFZR  strs_,strs1 _
        If strs_ <> ""  Then SSProcess.SetScriptDlgCellValue_ex dlgHandle2 ,"[ZhiTY]", strs _
        If strs1_ <> "" Then SSProcess.SetScriptDlgCellValue_ex dlgHandle2 ,"[ZhiTYZSH]",strs1 _
    End If
    If fieldName = "[JCY]" Then
        GetXMFZR  strs_,strs1 _
        If strs_ <> ""  Then SSProcess.SetScriptDlgCellValue_ex dlgHandle2 ,"[JianCY]", strs _
        If strs1_ <> "" Then SSProcess.SetScriptDlgCellValue_ex dlgHandle2 ,"[JianCYZSH]",strs1 _
    End If
    If fieldName = "[SHY]" Then
        GetXMFZR  strs_,strs1 _
        If strs_ <> ""  Then SSProcess.SetScriptDlgCellValue_ex dlgHandle2 ,"[ShenHR]", strs _
        If strs1_ <> "" Then SSProcess.SetScriptDlgCellValue_ex dlgHandle2 ,"[ShenHRZSH]",strs1 _
    End If
    
    If fieldName = "[CHZRR]"  Then
        
        fileName = SSProcess.GetSysPathName(7) & "���������.txt"
        Set oShell = CreateObject ("Wscript.shell")
        oShell.run   fileName
        Set oshell = Nothing
    End If
    
    Dim strs(10),strs1(10)
    mark = 1
    If fieldName = "[SAVE]" Then'������水ť
        lhcybh = SSProcess.GetScriptDlgCellValue_ex (dlgHandle,"[HeTBH]")
        xmmc = SSProcess.GetScriptDlgCellValue_ex (dlgHandle,"[XiangMMC]")
        ywlx = SSProcess.GetScriptDlgCellValue_ex (dlgHandle,"[YeWLX]")
        If lhcybh = ""  Or  xmmc = ""  Or    ywlx = ""     Then mark = 0
        MsgBox "����д��Ŀ���ơ���ͬ�š�ҵ�����ͣ�"
        MsgBox "����ʧ�ܣ�"
        Exit Function
        For i = 0 To ConCount - 1
            lhcybh = SSProcess.GetScriptDlgCellValue_ex (dlgHandle,"[HeTBH]")
            xmmc = SSProcess.GetScriptDlgCellValue_ex (dlgHandle,"[XiangMMC]")
            If lhcybh = ""  And xmmc = ""  Then mark = 0
            MsgBox "����д��Ŀ���ƺͺ�ͬ�ţ�"
            Exit For
            If i < 5 Then value = SSProcess.GetScriptDlgCellValue_ex (dlgHandle,ConIDS(i))
            If i > 4 And i < 21 Then value = SSProcess.GetScriptDlgCellValue_ex (dlgHandle1,ConIDS(i))
            If i > 20 Then value = SSProcess.GetScriptDlgCellValue_ex (dlgHandle2,ConIDS(i))
            If ConIDS(i) = "[ZhiTRQ]"   Or  ConIDS(i) = "[CeLRQ]"  Or  ConIDS(i) = "[KSSCSJ]"  Or  ConIDS(i) = "[JSSCSJ]"  Or  ConIDS(i) = "[JianCRQ]" Or  ConIDS(i) = "[ShenHRQ]"  Then value = FormatDateTime(value,1)
            SSFunc.ScanString GFields(i),",",strs,scount                'GFields(i)  ��Ӧ�����ֶ�
            SSFunc.ScanString value,"��",strs1,scount1
            If scount = 1 Then
                SSProcess.SetObjectAttr ZDID,"[" & GFields(i) & "]",value
            Else
                If scount1 > 1 Then
                    For j = 0 To scount - 1
                        SSProcess.SetObjectAttr ZDID,"[" & strs(j) & "]",strs1(j)
                    Next
                End If
            End If
        Next
        Dim zdarRecordList(),zdRecordListCount,zrzarRecordList(),zrzRecordListCount
        SSProcess.OpenAccessMdb mdbName
        Dim LCarRecordList(),LCRecordListCount
        sql = "Select ZDGUID,CeLY,CeLRQ,ZhiTY,ZhiTRQ,JianCY,JianCRQ,ShenHR,ShenHRQ,DiaoCR,DiaoCRQ From ZD_�ڵػ�����Ϣ���Ա� INNER JOIN GeoAreaTB ON ZD_�ڵػ�����Ϣ���Ա�.ID=GeoAreaTB.ID WHERE (GeoAreaTB.Mark mod 2)<>0 And ZD_�ڵػ�����Ϣ���Ա�.ID = " & ZDID
        GetSQLRecordAll mdbName,sql,zdarRecordList,zdRecordListCount
        
        If zdRecordListCount = 1  Then
            artempzd = Split(zdarRecordList(0),",")
            sql = "select FC_��Ȼ����Ϣ���Ա�.ID from FC_��Ȼ����Ϣ���Ա� inner join GeoAreaTB ON FC_��Ȼ����Ϣ���Ա�.ID = GeoAreaTB.ID Where (GeoAreaTB.Mark mod 2)<>0 and FC_��Ȼ����Ϣ���Ա�.ZDGUID =" & artempzd(0)
            GetSQLRecordAll mdbName,sql,zrzarRecordList,zrzRecordListCount
            
            For j = 0 To zrzRecordListCount - 1
                
                SSProcess.SetObjectAttr zrzarRecordList(j),"[CeLY],[CeLRQ],[ZhiTY],[HuiTRQ],[JianCY],[JianCRQ],[ShenHR],[ShenHRQ],[DiaoCR],[DiaoCRQ]",artempzd(1) & "," & artempzd(2) & "," & artempzd(3) & "," & artempzd(4) & "," & artempzd(5) & "," & artempzd(6) & "," & artempzd(7) & "," & artempzd(8) & "," & artempzd(9) & "," & artempzd(10)
                
            Next
        End If
        SSProcess.CloseAccessMdb mdbName
        
        KSSCSJ = SSProcess.GetScriptDlgCellValue_ex (dlgHandle1,"[KSSCSJ]")
        scsj = Left(KSSCSJ,4)
        SSProcess.SetObjectAttr ZDID,"[ShiCNF]",ZDID
        
        MsgBox "����ɹ���"
    End If
    
    If fieldName = "[CLOSE]" Then'����رհ�ť
        SSProcess.CloseFreeScriptDlg dlgHandle
        Exit Function
    End If
End Function

'��ȡָ��sql����µ�  �������
Function GetSQLRecordAll(ByRef mdbName, ByRef StrSqlStatement, ByRef arSQLRecord(), ByRef iRecordCount)
    SSProcess.OpenAccessMDB mdbName
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
            arSQLRecord(iRecordCount) = values                                        '��ѯ��¼
            iRecordCount = iRecordCount + 1                                                    '��ѯ��¼��
            '�ƶ���¼�α�
            SSProcess.AccessMoveNext mdbName, sql
        WEnd
    End If
    '�رռ�¼��
    SSProcess.CloseAccessRecordset mdbName, sql
    SSProcess.CloseAccessMDB mdbName
End Function
'ѡ��������ı�
'Function OnComboBoxSelChange( tableName, objectID, fieldName, fieldValue )
'��Ӵ���
'End Function

'���ڴ�С�����ı�
Function OnSize( dHandle, cx, cy )
    If dHandle = dlgHandle Then
        ctrlSize = SSProcess.GetScriptDlgCellRect_ex (dlgHandle,"{�༭��}")
        Dim vArray(4), nCount
        SSFunc.ScanString ctrlSize, ",", vArray, nCount
        rectleft = CLng(vArray(0))
        recttop = CLng(vArray(1))
        rectright = CLng(vArray(2))
        rectbottom = CLng(vArray(3))
        rectright = cx - 10
        rectbottom = cy - 10
        SSProcess.SetScriptDlgCellRect_ex dlgHandle, "{�༭��}", rectLeft, rectTop, rectRight, rectBottom
    End If
End Function


Function IniDlgParameter()
    Ininame = "BDC_��Ŀ��Ϣ¼��-1.ini"
    ReadIniInfo Ininame, ConIDS,GFields,DefValues,AltValues, MemValues,ConCount
End Function


'����DLG��Ӧ��ini���ü�¼
'��Σ�Ininame��ini�ļ�������
'��η��أ�ControlIDS���ؼ�ID������ʽ��GeoFields����Ӧ������ֶ���������ʽ��DefaultValues��Ĭ��ֵ������ʽ��AlternativeValues����ѡ������������ʽ��MemoryValues���Ƿ�����ϴδ����ֵ������ʽ��ControlCount ���ؼ�����
Function ReadIniInfo(ByVal Ininame,ByRef  ControlIDS(),ByRef GeoFields(),ByRef DefaultValues(),ByRef AlternativeValues(),ByRef MemoryValues(),ByRef ControlCount)
    ControlCount = 0
    
    TemplateFileName = SSProcess.GetTemplateFileName
    Ininame_ = Left(TemplateFileName,Len(TemplateFileName) - 4) & "\" & Ininame
    MsgBox Ininame _
    Dim strs(50)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(Ininame_, 1, False)
    While (f.atEndOfStream = False)
        strLine = f.ReadLine()
        If strLine <> "" And Len(strLine) > 4 Then
            If Left(strLine,2) <> "//" Then
                SSFunc.ScanString strLine,"|",strs,scount
                If scount = 5 Then
                    ControlIDS(ControlCount) = strs(0)
                    GeoFields(ControlCount) = strs(1)
                    DefaultValues(ControlCount) = strs(2)
                    AlternativeValues(ControlCount) = strs(3)
                    MemoryValues(ControlCount) = strs(4)
                    ControlCount = ControlCount + 1
                End If
            Else
                ReadIniInfo = strLine
            End If
        End If
    WEnd
    f.Close()
    Set f = Nothing
    Set fso = Nothing
End Function



