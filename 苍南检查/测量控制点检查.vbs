
'=============================================��鼯����==============================================

'�������Ŀ����
Dim strGroupName
strGroupName = "�������Ƶ�����ֵ������"

'��鼯������
Dim strCheckName
strCheckName = "�������Ƶ�����ֵ������"

'���ģ������
Dim CheckmodelName
CheckmodelName = "�Զ���ű������->�������Ƶ�����ֵ������"

'�������
Dim strDescription
strDescription = "�������Ƶ�FCODE,NAME,GRADE,XYCOOR,X,Y,ELEVATION,ZCOOR����ֵ��Ϊ��"

'========================================����ֶ�����=================================================

Dim FildsName
FildsName = "FCODE,NAME,GRADE,XYCOOR,X,Y,ELEVATION,ZCOOR"

'==========================================��������===================================================

'�������
Sub OnClick()
    
    AllVisible

    ClearCheckRecord
    
    SelFeatures "�������Ƶ�",IdCount,IdArr
    RecordsInner IdCount,IdArr
    
    Ending

End Sub' OnClick

'=========================================��鼯���=============================================

'��鼯������
Function RecordsInner(ByVal IdCount,ByVal IdArr())
    ExportRecords IdCount,IdArr
End Function' RecordsInner

'�����鼯
Function ExportRecords(ByVal IdCount,ByVal IdArr())
    For i = 0 To IdCount - 1
        If IsEmpty(IdArr(i),FildsName) Then
            AddCheckRecord IdArr(i)
        End If
    Next 'i
    SSProcess.ShowCheckOutput
    SSProcess.SaveCheckRecord
End Function' ExportRecords

'=========================================�����ຯ��===========================================

'������ͼ��
Function AllVisible()
    count = SSProcess.GetLayerCount
    For i = 0 To count - 1
        layername = SSProcess.GetLayerName (i)
        SSProcess.SetLayerStatus layername, 1, 1
    Next
    SSProcess.RefreshView
End Function

'��ռ�鼯
Function ClearCheckRecord()
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
End Function' ClearCheckRecord

'ѡ��Ҫ��
Function SelFeatures(ByVal LayerName,ByRef TotalCount,ByRef AllIdArr())
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", LayerName
    SSProcess.SelectFilter
    TotalCount = SSProcess.GetSelGeoCount
    ReDim AllIdArr(TotalCount)
    For i = 0 To TotalCount - 1
        AllIdArr(i) = SSProcess.GetSelGeoValue(i,"SSObj_ID")
    Next 'i
End Function' SelAllFeatures

'�жϿ���ֶ��Ƿ�Ϊ��
Function IsEmpty(ByVal Id,ByVal KeyString)
    SplitKeyString KeyString,KeyArr,KeyCount
    IsEmpty = False
    For i = 0 To KeyCount - 1
        If SSProcess.GetObjectAttr(Id,KeyArr(i)) = "" Then IsEmpty = True
    Next 'i
End Function' IsEmpty

'�ֽ���ַ���
Function SplitKeyString(ByVal StringName,ByRef StrArr(),ByRef StrCount)
    StrArr = Split(StringName,",", - 1,1)
    StrCount = UBound(StrArr) + 1
    For i = 0 To StrCount - 1
        StrArr(i) = "[" & StrArr(i) & "]"
    Next 'i
End Function' SplitKeyString

'��ӵ�����¼
Function AddCheckRecord(ByVal Id)
    SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(Id,"SSObj_X"),SSProcess.GetObjectAttr(Id,"SSObj_Y"),0,1,Id,""
End Function' AddCheckRecord

'��������
Function Ending()
    MsgBox "������"
End Function' Ending