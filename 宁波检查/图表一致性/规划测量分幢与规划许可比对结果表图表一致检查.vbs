
'==========================================================�������============================================================

'��鼯��Ŀ����
Dim strGroupName
strGroupName = "�滮�����ִ���滮��ɱȶԽ����"

'��鼯������
Dim strCheckName
strCheckName = "ͼ��һ���Լ��"

'���ģ������
Dim CheckmodelName
CheckmodelName = "�Զ���ű������->ͼ��һ���Լ��"

'�������
Dim strDescription
strDescription = "�滮�����ִ���滮��ɱȶԽ����,��Ȼ���С�ID_ZRZ����ʵ��㡾ID_ZRZ�����Ҳ�����ͬ��ֵ"

'================================================================��������======================================================

'��Ȼ�����Ա�
Dim FxTable
FxTable = "FC_��Ȼ����Ϣ���Ա�"

'���������Ա�
Dim RealTable
RealTable = "JG_ʵ������Ա�"

'�滮���������Ա�
Dim GuiHuaTable
GuiHuaTable = "JG_�ܱ߹�ϵУ�˱�ע���Ա�"

'�ݶ��߶����Ա�
Dim WdHeighTable
WdHeighTable = "JZWDGDXX"

'����ͼ��ע���Ա�
Dim LmtBzTbale
LmtBzTbale = "JG_����ͼ��ע���Ա�"

'�����ע���Ա�
Dim LenBzTable
LenBzTable = "JG_�����ע���Ա�"

'=============================================================�������=======================================================================

Sub OnClick()
    
    AddRecordInner
    
    
End Sub' OnClick

'=============================================================����ֶ��жϲ���Ӽ���¼================================================

'��Ӽ���¼���
Function AddRecordInner()
    ClearCheckRecord
    FxPoiInfo FxDhArr,DhCount
    ConfirmScPoi FxDhArr,DhCount
    ConfirmGhArea FxDhArr,DhCount
    ConfirmHeigh  FxDhArr,DhCount
    ConfirmLmt FxDhArr,DhCount
    ConfirmLen FxDhArr,DhCount
    ShowCheckRecord
End Function' AddRecordInner

'��ȡ��Ȼ����ID��ID_ZRZ
Function FxPoiInfo(ByRef FxDhArr(),ByRef DhCount)
    SqlStr = "Select " & FxTable & ".ID_ZRZ," & FxTable & ".ID" & " From " & FxTable & " Inner Join GeoAreaTB on " & FxTable & ".ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0 And " & FxTable & ".ID_ZRZ <> " & "'" & "*" & "'"
    GetSQLRecordAll SqlStr,FxDhArr,DhCount
End Function' FxPoiInfo

'�ж�ʵ����Ƿ����
Function ConfirmScPoi(ByVal FxDhArr(),ByVal DhCount)
    For i = 0 To DhCount - 1
        FxArr = Split(FxDhArr(i),",", - 1,1)
        If FxArr(0) <> "" Then
            SqlStr = "Select " & RealTable & ".ID From " & RealTable & " Inner Join GeoPointTB on " & RealTable & ".ID = GeoPointTB.ID WHERE (GeoPointTB.Mark Mod 2)<>0 And " & RealTable & ".ID_ZRZ = " & "'" & FxArr(0) & "'"
            GetSQLRecordAll SqlStr,ScDhArr,ScCount
            If ScCount <= 0 Then
                SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(FxArr(1),"SSObj_X"),SSProcess.GetObjectAttr(FxArr(1),"SSObj_Y"),0,0,FxArr(1),""
            End If
        End If
    Next 'i
End Function' ConfirmScPoi

'�ж��ܱ߱�ע�Ƿ����
Function ConfirmGhArea(ByVal FxDhArr(),ByVal DhCount)
    For i = 0 To DhCount - 1
        FxArr = Split(FxDhArr(i),",", - 1,1)
        If FxArr(0) <> "" Then
            SqlStr = "Select " & GuiHuaTable & ".ID From " & GuiHuaTable & " Inner Join GeoLineTB on " & GuiHuaTable & ".ID = GeoLineTB.ID WHERE (GeoLineTB.Mark Mod 2)<>0" & " And " & GuiHuaTable & ".ID_ZRZ_QD= " & "'" & FxArr(0) & "'"
            GetSQLRecordAll SqlStr,ScDhArr,ScCount
            'MsgBox ScCount
            If ScCount <= 0 Then
                strDescription = "�����ｨ��������ܱ�,��Ȼ���С�ID_ZRZ�����ܱ߹�ϵУ�˱�ע��ID_ZRZ_QD�����Ҳ�����ͬ��ֵ"
                SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(FxArr(1),"SSObj_X"),SSProcess.GetObjectAttr(FxArr(1),"SSObj_Y"),0,0,FxArr(1),""
            End If
        End If
    Next 'i
End Function' ConfirmGhArea

'�жϸ߶��Ƿ����
Function ConfirmHeigh(ByVal FxDhArr(),ByVal DhCount)
    For i = 0 To DhCount - 1
        FxArr = Split(FxDhArr(i),",", - 1,1)
        If FxArr(0) <> "" Then
            SqlStr = "Select " & WdHeighTable & ".ID From " & WdHeighTable & " WHERE " & WdHeighTable & ".ID_ZRZ = " & "'" & FxArr(0) & "'"
            GetSQLRecordAll SqlStr,ScDhArr,ScCount
            If ScCount <= 0 Then
                strDescription = "�����ｨ��������ܱ�,��Ȼ���С�ID_ZRZ���ڽ������ݶ��߶���Ϣ��ID_ZRZ�����Ҳ�����ͬ��ֵ"
                SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(FxArr(1),"SSObj_X"),SSProcess.GetObjectAttr(FxArr(1),"SSObj_Y"),0,0,FxArr(1),""
            End If
        End If
    Next 'i
End Function' ConfirmHeigh

'�ж�����ͼע���Ƿ����
Function ConfirmLmt(ByVal FxDhArr(),ByVal DhCount)
    For i = 0 To DhCount - 1
        FxArr = Split(FxDhArr(i),",", - 1,1)
        If FxArr(0) <> "" Then
            SqlStr = "Select " & LmtBzTbale & ".ID From " & LmtBzTbale & " Inner Join GeoLineTB on " & LmtBzTbale & ".ID = GeoLineTB.ID WHERE (GeoLineTB.Mark Mod 2)<>0" & " And " & LmtBzTbale & ".ID_ZRZ= " & "'" & FxArr(0) & "'"
            GetSQLRecordAll SqlStr,ScDhArr,ScCount
            If ScCount <= 0 Then
                strDescription = "�����ｨ��������ܱ�,��Ȼ���С�ID_ZRZ��������ͼ��ע���Ա�ID_ZRZ�����Ҳ�����ͬ��ֵ"
                SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(FxArr(1),"SSObj_X"),SSProcess.GetObjectAttr(FxArr(1),"SSObj_Y"),0,0,FxArr(1),""
            End If
        End If
    Next 'i
End Function' ConfirmLmt

'�жϾ���ע���Ƿ����
Function ConfirmLen(ByVal FxDhArr(),ByVal DhCount)
    For i = 0 To DhCount - 1
        FxArr = Split(FxDhArr(i),",", - 1,1)
        If FxArr(0) <> "" Then
            SqlStr = "Select " & LenBzTable & ".ID From " & LenBzTable & " Inner Join GeoLineTB on " & LenBzTable & ".ID = GeoLineTB.ID WHERE (GeoLineTB.Mark Mod 2)<>0" & " And " & LenBzTable & ".ID_ZRZ= " & "'" & FxArr(0) & "'"
            GetSQLRecordAll SqlStr,ScDhArr,ScCount
            If ScCount <= 0 Then
                strDescription = "�����ｨ��������ܱ�,��Ȼ���С�ID_ZRZ���ھ����ע���Ա�ID_ZRZ�����Ҳ�����ͬ��ֵ"
                SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(FxArr(1),"SSObj_X"),SSProcess.GetObjectAttr(FxArr(1),"SSObj_Y"),0,0,FxArr(1),""
            End If
        End If
    Next 'i
End Function' ConfirmLen

'==============================================================���ߺ���==========================================================

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
End Function' GetSQLRecordAll

'��ռ�鼯
Function ClearCheckRecord()
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
End Function' ClearCheckRecord

'��ʾ����¼
Function ShowCheckRecord()
    SSProcess.ShowCheckOutput
    SSProcess.SaveCheckRecord
End Function' ShowCheckRecord
