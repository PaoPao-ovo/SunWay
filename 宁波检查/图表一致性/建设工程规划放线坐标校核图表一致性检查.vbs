
'==========================================================�������============================================================

'��鼯��Ŀ����
Dim strGroupName
strGroupName = "���蹤�̹滮��������У�˱�"

'��鼯������
Dim strCheckName
strCheckName = "ͼ��һ���Լ��"

'���ģ������
Dim CheckmodelName
CheckmodelName = "�Զ���ű������->ͼ��һ���Լ��"

'�������
Dim strDescription
strDescription = "���蹤�̹滮��������У�˱�,ʵ����С�DH���ڷ���׮�㡾DH�����Ҳ�����ͬ��ֵ"

'================================================================��������======================================================

'����׮�����Ա�
Dim FxTable
FxTable = "FX_����׮�����Ա�"

'����ʵ������Ա�
Dim RealTable
RealTable = "FX_����ʵ��У�˵����Ա�"

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
    ShowCheckRecord
End Function' AddRecordInner

'��ȡ����׮��ĵ��
Function FxPoiInfo(ByRef FxDhArr(),ByRef DhCount)
    SqlStr = "Select " & FxTable & ".DH," & FxTable & ".ID" & " From " & FxTable & " Inner Join GeoPointTB on " & FxTable & ".ID = GeoPointTB.ID WHERE (GeoPointTB.Mark Mod 2)<>0 And " & FxTable & ".DH <> " & "'" & "*" & "'"
    GetSQLRecordAll SqlStr,FxDhArr,DhCount
End Function' FxPoiInfo

'�ж�ʵ����Ƿ����
Function ConfirmScPoi(ByVal FxDhArr(),ByVal DhCount)
    For i = 0 To DhCount - 1
        FxArr = Split(FxDhArr(i),",", - 1,1)
        If FxArr(0) <> "" Then
            SqlStr = "Select " & RealTable & ".ID From " & RealTable & " Inner Join GeoPointTB on " & RealTable & ".ID = GeoPointTB.ID WHERE (GeoPointTB.Mark Mod 2)<>0 And " & RealTable & ".DH = " & "'" & FxArr(0) & "'"
            GetSQLRecordAll SqlStr,ScDhArr,ScCount
            If ScCount <= 0 Then
                SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(FxArr(1),"SSObj_X"),SSProcess.GetObjectAttr(FxArr(1),"SSObj_Y"),0,0,FxArr(1),""
            End If
        End If
    Next 'i
End Function' ConfirmScPoi

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
