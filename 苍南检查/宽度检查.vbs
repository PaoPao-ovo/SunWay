
'=============================================��鼯����==============================================

'�������Ŀ����
Dim strGroupName:strGroupName = "����ֵ������"

'��鼯������
Dim strCheckName:strCheckName = "��ȱ�����"

'���ģ������
Dim CheckmodelName:CheckmodelName = "�Զ���ű������->��ȱ�����"

'�������
Dim strDescription:strDescription = "���Ϊ��"

'==========================================��������===================================================

SSObj_Codes = "2206005,2207005,2208005,2209005,2701013,2702035,2702045,2702065,2702075,2702085,2705003,2705005,2706005,2706015,4201015,4201025,4201018,4201028,4202015,4202025,4202018,4202028,4203015,4203025,4204005,4204007,4205005,4205007,4206005,4206007,4208005,4208007,4302005,4303005,4303015,4304005,4305015,4305025,4305035,4305045,4306005,4307005,4401005,4402005,4503015,4503065,4503025,4503035,4503045,4503055,4503075,4505015,4505025,4505075,4505045,4505055,4505056,4505065,4505085,4506015,4506025,4507004,4507007,4904003,4905074,4905072,7505022,7505023"

'==========================================��������===================================================

'�������
Sub OnClick()

    AllVisible

    ClearCheckRecord
    
    SelFeatures SSObj_Codes,IdCount,IdArr
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
        If IsWidthEmpty(IdArr(i)) Then
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
Function SelFeatures(ByVal Codes,ByRef TotalCount,ByRef AllIdArr())
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", Codes
    SSProcess.SelectFilter
    TotalCount = SSProcess.GetSelGeoCount
    ReDim AllIdArr(TotalCount)
    For i = 0 To TotalCount - 1
        AllIdArr(i) = SSProcess.GetSelGeoValue(i,"SSObj_ID")
    Next 'i
End Function' SelAllFeatures

'�жϿ���ֶ��Ƿ�Ϊ��
Function IsWidthEmpty(ByVal Id)
    If SSProcess.GetObjectAttr(Id,"[WIDTH]") = "" Then
        IsWidthEmpty = True
    Else
        IsWidthEmpty = False
    End If
End Function' IsWidthEmpty

'��ӵ�����¼
Function AddCheckRecord(ByVal Id)
    SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(Id,"SSObj_X"),SSProcess.GetObjectAttr(Id,"SSObj_Y"),0,1,Id,""
End Function' AddCheckRecord

Function Ending()
    MsgBox "������"
End Function' Ending