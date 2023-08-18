'======================================================��鼯����=====================================================

'��鼯��Ŀ����
Dim strGroupName
strGroupName = "���߼��"

'��鼯������
Dim strCheckName
strCheckName = "���߸̼߳��"

'���ģ������
Dim CheckmodelName
CheckmodelName = "�Զ���ű������->���߸̼߳��"

'�������
Dim strDescription
strDescription = "���߸߳�Ϊ��"

'==================================================ʵ����������=========================================================

'ʵ���Ԥ���Ӧ��ϵ��
' ʵ�����            ����                ���۱���
' 9130512           GPS����            1103021
' 9130412           ˮ׼��                1102021
' 9130311           ���Ƶ㣨��ʯ��         9130211
' 9130312           ���Ƶ㣨����ʯ��         9130212
' 9130217           ��վ��                9130216
' 9130511           ������               9130411


ScdCodes = "9130224"

'===================================================��������==========================================================


'��ں���
Sub OnClick()
    ClearCheckRecord()
    ExportRecords ScdCodes
End Sub' OnClick

'��ռ�鼯
Function ClearCheckRecord()
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
End Function' ClearCheckRecord

'�����鼯
Function ExportRecords(code)
    SelRealPoi code
    SelCount = SSProcess.GetSelGeoCount()
    If SelCount > 0  Then
        polygonID = SSProcess.GetSelGeoValue(0,"SSObj_ID")
        idstr = SSProcess.SearchInnerObjIDs(polygonID, 0, "9130611",0)
        idarr = Split(idstr,",", - 1,1)
        For i = 0 To UBound(idarr)
            yxgc = SSProcess.GetObjectAttr(idarr(i),"[YanXGC]")
            x = SSProcess.GetObjectAttr(idarr(i),"SSObj_X")
            y = SSProcess.GetObjectAttr(idarr(i),"SSObj_Y")
            z = SSProcess.GetObjectAttr(idarr(i),"SSObj_Z")
            yxgc = transform(yxgc)
            'MsgBox yxgc
            If yxgc = 0 Then SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription, x, y, 0,0,idarr(i), ""
        Next 'i
    End If
    SSProcess.ShowCheckOutput
    SSProcess.SaveCheckRecord
End Function' ExportRecords

'ѡ��ʵ���
Function SelRealPoi(Code)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", Code
    SSProcess.SelectFilter
End Function' CheckRealPoi

Function SelLlPoi(Code,poiname)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", Code
    SSProcess.SetSelectCondition "SSObj_PointName", "==", poiname
    SSProcess.SelectFilter
End Function' SelLlPoi

'��������ת��
Function transform(content)
    If content <> "" Then
        content = CDbl(content)
    Else
        content = 0
    End If
    transform = content
End Function