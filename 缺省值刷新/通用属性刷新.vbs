
' 1��ˢ��ͨ�������е�����Դ�͸���ʱ�䣬����ʱ��ѡ��ǰϵͳ�µ�ʱ��
' 2����ͨ��������ֵ��ԭֵ����û��ֵ����ȱʡֵˢ��

'=====================================================��ֵ������==============================================================

'ͨ�������ֶ�
Dim CommonString
CommonString = "UPDATEDATE,DATASOURCE,FEATURESTATUS"

'ͨ������ȱʡֵ
Dim CommonValString
CommonValString = GetNowTime & "," & "1:500��������ͼ,2"

'==========================================================�������==================================================================

'�������
Sub OnClick()
    
    AllVisible
    
    SelAllFeatures IdCount,IdArr
    CommonAttrInner IdCount,IdArr
    
    Ending
    
End Sub' OnClick

'================================================================ͨ������ˢ��=================================================================

'ͨ������ˢ�����
Function CommonAttrInner(ByVal IdCount,ByVal IdArr())
    UpDateAttribute IdCount,IdArr
End Function' CommonAttrInner

'ͨ������ˢ�º���
Function UpDateAttribute(ByVal IdCount,ByVal IdArr())
    SplitKeyString CommonString,CommonArr,CommonCount
    SplitString CommonValString,ValArr,ValCount
    For i = 0 To IdCount - 1
        For j = 0 To ValCount - 1
            If SSProcess.GetObjectAttr(IdArr(i),CommonArr(j)) = "" Then
                SSProcess.SetObjectAttr IdArr(i),CommonArr(j),ValArr(j)
            End If
        Next 'j
    Next 'i
End Function' UpDateAttribute

'=================================================================�����ຯ��==================================================================

'������ͼ��
Function AllVisible()
    count = SSProcess.GetLayerCount
    For i = 0 To count - 1
        layername = SSProcess.GetLayerName (i)
        SSProcess.SetLayerStatus layername, 1, 1
    Next
    SSProcess.RefreshView
End Function

'ѡ������Ҫ��
Function SelAllFeatures(ByRef TotalCount,ByRef AllIdArr())
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Type", "==", "POINT,LINE,AREA,NOTE"
    SSProcess.SelectFilter
    TotalCount = SSProcess.GetSelGeoCount
    ReDim AllIdArr(TotalCount)
    For i = 0 To TotalCount - 1
        AllIdArr(i) = SSProcess.GetSelGeoValue(i,"SSObj_ID")
    Next 'i
End Function' SelAllFeatures

'�ֽ��ַ���
Function SplitString(ByVal StringName,ByRef StrArr(),ByRef StrCount)
    StrArr = Split(StringName,",", - 1,1)
    StrCount = UBound(StrArr) + 1
    MsgBox StrCount
End Function' SplitString

'�ֽ���ַ���
Function SplitKeyString(ByVal StringName,ByRef StrArr(),ByRef StrCount)
    StrArr = Split(StringName,",", - 1,1)
    StrCount = UBound(StrArr) + 1
    For i = 0 To StrCount - 1
        StrArr(i) = "[" & StrArr(i) & "]"
    Next 'i
End Function' SplitKeyString

'��ȡ��ǰϵͳʱ��
Function GetNowTime()
    GetNowTime = FormatDateTime(Now(),2)
End Function' GetNowTime

Function Ending()
    MsgBox "ˢ�����"
End Function' Ending