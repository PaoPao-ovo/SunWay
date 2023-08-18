
'=========================================================ͼ����������=======================================================

LayStr = "CD,CT,CY,CQ,CS,QT,BM,FQ,DL,GD,LD,DC,XH,TX,DX,YD,LT,JX,JK,EX,DS,BZ,JS,XF,PS,YS,WS,FS,RQ,MQ,TR,YH,RL,RS,ZQ,SY,GS"

MainCode = "54311203,54324004,54323004,54412004,54423004,54452004,54511004,54512114,54534114,54523114,54611114,54612004,54623004,54111003,54112003,54123003,54145003,54134003,54211003,54212003,54223003,54234003,54245003,54256003,54267003,54278003,54289003,54720114,54730114,54030003,54040003,51011203,52011203,53011204,53022204,53033204,53044204"

Table_LineName = "���¹��������Ա�"

'===========================================�������========================================================

'�����
Sub OnClick()
    
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", "ͼ����"
    SSProcess.SetSelectCondition "SSObj_Code", "==", "59999999"
    SSProcess.SelectFilter
    SelCount = SSProcess.GetSelGeoCount
    If SelCount <= 0 Then
        MsgBox "������ͼ��"
        Exit Sub
    End If
    
    AllVisible
    
    DelFormerTl
    
    GetMapBorderPoision StandX,StandY
    
    LayArr = Split(LayStr,",", - 1,1)
    
    BorderStartX = StandX - 10 - 20 - 4
    BorderStartY = StandY - 10
    BorderEndX = StandX - 14 + 4
    FeatureY = BorderStartY - 2 - 2
    
    For i = 0 To UBound(LayArr)
        SelAllLine LayArr(i),CodeVal,CodeCount
        If CodeCount > 0 Then
            CodeArr = Split(CodeVal,",", - 1,1)
            For j = 0 To CodeCount - 1
                If j = 0 Then
                    DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),200,200
                    FeatureY = FeatureY - 2.25
                Else
                    DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),150,150
                    FeatureY = FeatureY - 2.25
                End If
            Next 'j
        End If
        SelAllPoi LayArr(i),CodeVal,CodeCount
        If CodeCount > 0 Then
            CodeArr = Split(CodeVal,",", - 1,1)
            For j = 0 To CodeCount - 1
                DrawPoint BorderStartX + 5,FeatureY,CodeArr(j),150,150
                FeatureY = FeatureY - 2.25
            Next 'j
        End If
    Next 'i
    
    DrawBorder BorderStartX,BorderEndX,BorderStartY,FeatureY
    
    Ending
    
End Sub' OnClick

'==============================================ͼ������==========================================================

'��ȡͼ�������Ͻ�����ֵ
Function GetMapBorderPoision(ByRef X,ByRef Y)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", "ͼ����"
    SSProcess.SetSelectCondition "SSObj_Code", "==", "59999999"
    SSProcess.SelectFilter
    ID = SSProcess.GetSelGeoValue(0,"SSObj_ID")
    SSProcess.GetObjectPoint ID, 2, X, Y, Z, PointType, Name '���Ͻǵ�����ֵ
End Function' GetMapBorderPoision

'��ȡ���е���Ҫ������
Function SelAllLine(ByVal LayerName,ByRef CodeVal,ByRef CodeCount)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", LayerName
    SSProcess.SetSelectCondition "SSObj_Type", "==", "LINE"
    SSProcess.SetSelectCondition "SSObj_Code", "<>", "54100004,54200304,54245304,54256304,54267304,54412005,54423005,54452005,54111004,54211304,54400005,54411005,54212304,54223304,54120004,54130004,54140004,54150004,54234304,54278304,54289304,0"
    SSProcess.SelectFilter
    LineCount = SSProcess.GetSelGeoCount
    Dim k
    k = 0
    ReDim Preserve CodeStr(k)
    For i = 0 To LineCount - 1
        If SSProcess.GetSelGeoValue(i,"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(i,"[FSFS]") = "�ǿ���" Then
            CodeStr(k) = SSProcess.GetSelGeoValue(i,"SSObj_Code")
            k = k + 1
            ReDim Preserve CodeStr(k)
        ElseIf SSProcess.GetSelGeoValue(i,"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(i,"[YYKS]") <> "0" Then
            CodeStr(k) = SSProcess.GetSelGeoValue(i,"SSObj_Code")
            k = k + 1
            ReDim Preserve CodeStr(k)
        End If
    Next 'i
    DelRepeatLine CodeStr,CodeVal,CodeCount
    MainCodeArr = Split(MainCode,",", - 1,1)
    
    CodeArr = Split(CodeVal,",", - 1,1)
    For i = 0 To CodeCount - 1
        For j = 0 To UBound(MainCodeArr)
            If CodeArr(i) = MainCodeArr(j) Then
                Temp = CodeArr(0)
                CodeArr(0) = CodeArr(i)
                CodeArr(i) = Temp
            End If
        Next 'j
    Next 'i
    
    CodeVal = ""
    For i = 0 To UBound(CodeArr)
        If CodeVal = "" Then
            CodeVal = CodeArr(i)
        Else
            CodeVal = CodeVal & "," & CodeArr(i)
        End If
    Next 'i
End Function' SelAllLine

'��ȡ���е���Ҫ������
Function SelAllPoi(ByVal LayerName,ByRef CodeVal,ByRef CodeCount)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", LayerName
    SSProcess.SetSelectCondition "SSObj_Type", "==", "Point"
    SSProcess.SetSelectCondition "SSObj_Code", "<>", "0"
    SSProcess.SelectFilter
    PoiCount = SSProcess.GetSelGeoCount
    ReDim CodeStr(PoiCount - 1)
    For i = 0 To PoiCount - 1
        CodeStr(i) = SSProcess.GetSelGeoValue(i,"SSObj_Code")
    Next 'i
    DelRepeat CodeStr,CodeVal,CodeCount
End Function' SelAllPoi

'���Ƶ�ע��
Function DrawPoint(ByVal X,ByVal Y,ByVal Code,ByVal Width,ByVal Height)
    SSProcess.CreateNewObj 0
    SSProcess.SetNewObjValue "SSObj_Code", Code
    SSProcess.SetNewObjValue "SSObj_LayerName", "����ͼ����"
    'SSProcess.SetNewObjValue "SSObj_GroupID", GropuId
    SSProcess.SetNewObjValue "SSObj_Color", SSProcess.GetFeatureCodeInfo(Code,"LineColor")
    SSProcess.AddNewObjPoint X,Y,0,0,""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
    DrawPointNote X + 7,Y,Code,Width,Height
End Function

'���Ƶ�ע����
Function DrawPointNote(ByVal X,ByVal Y,ByVal Code,ByVal Width,ByVal Height)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontString", SSProcess.GetFeatureCodeInfo(Code,"ObjectName")
    SSProcess.SetNewObjValue "SSObj_LayerName", "����ͼ����"
    'SSProcess.SetNewObjValue "SSObj_GroupID", GropuId
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "3"
    SSProcess.SetNewObjValue "SSObj_FontWidth", Width
    SSProcess.SetNewObjValue "SSObj_Color", SSProcess.GetFeatureCodeInfo(Code,"LineColor")
    SSProcess.SetNewObjValue "SSObj_FontHeight", Height
    SSProcess.AddNewObjPoint X + 2, Y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

'������ע��
Function DrawLine(ByVal X1,ByVal X2,ByVal Y,ByVal Code,ByVal Width,ByVal Height)
    SSProcess.CreateNewObj 1
    SSProcess.SetNewObjValue "SSObj_Code", Code
    SSProcess.SetNewObjValue "SSObj_LayerName", "����ͼ����"
    SSProcess.SetNewObjValue "SSObj_Color", SSProcess.GetFeatureCodeInfo(Code,"LineColor")
    SSProcess.AddNewObjPoint X1, Y, 0, 0, ""
    SSProcess.AddNewObjPoint X2, Y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
    DrawLineNote X2 + 1,Y,Code,Width,Height
End Function

'������ע����
Function DrawLineNote(ByVal X,ByVal Y,ByVal Code,ByVal Width,ByVal Height)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontString", SSProcess.GetFeatureCodeInfo(Code,"ObjectName")
    SSProcess.SetNewObjValue "SSObj_LayerName", "����ͼ����"
    SSProcess.SetNewObjValue "SSObj_Color", SSProcess.GetFeatureCodeInfo(Code,"LineColor")
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "3"
    SSProcess.SetNewObjValue "SSObj_FontWidth", Width
    SSProcess.SetNewObjValue "SSObj_FontHeight", Height
    SSProcess.AddNewObjPoint X + 4, Y, 0, 0, ""
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
    DrawTitle (StartX + EndX) / 2,StartY - 1,250,250
    
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

'ɾ��ͼ��
Function DelFormerTl()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", "����ͼ����"
    SSProcess.SelectFilter
    SSProcess.DeleteSelectionObj()
End Function' DelFormerTl

'=================================================================���ߺ���=======================================================

Function AllVisible()
    Count = SSProcess.GetLayerCount
    For i = 0 To Count - 1
        SSProcess.SetLayerStatus SSProcess.GetLayerName(i), 1, 1
    Next
    SSProcess.RefreshView
End Function

'ȥ���ַ������ظ�ֵ(����)
Function DelRepeatLine(ByVal StrArr(),ByRef ToTalVal,ByRef LxCount)
    ToTalVal = ""
    For i = 0 To UBound(StrArr) - 1
        If ToTalVal = "" Then
            ToTalVal = "'" & StrArr(i) & "'"
        ElseIf Replace(ToTalVal,StrArr(i),"") = ToTalVal Then
            ToTalVal = ToTalVal & "," & "'" & StrArr(i) & "'"
            'MsgBox i & "----" & StrArr(i)
        End If
    Next 'i
    ToTalVal = Replace(ToTalVal,"'","")
    LxCount = UBound(Split(ToTalVal,",", - 1,1)) + 1
End Function' DelRepeatLine

'ȥ���ַ������ظ�ֵ
Function DelRepeat(ByVal StrArr(),ByRef ToTalVal,ByRef LxCount)
    ToTalVal = ""
    For i = 0 To UBound(StrArr)
        If ToTalVal = "" Then
            ToTalVal = "'" & StrArr(i) & "'"
        ElseIf Replace(ToTalVal,StrArr(i),"") = ToTalVal Then
            ToTalVal = ToTalVal & "," & "'" & StrArr(i) & "'"
            'MsgBox i & "----" & StrArr(i)
        End If
    Next 'i
    ToTalVal = Replace(ToTalVal,"'","")
    LxCount = UBound(Split(ToTalVal,",", - 1,1)) + 1
End Function' DelRepeat

Function Ending()
    MsgBox "�������"
End Function' Ending

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