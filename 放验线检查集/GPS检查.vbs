'======================================================��鼯����=====================================================

'��鼯��Ŀ����
Dim strGroupName
strGroupName = "���߼��"

'��鼯������
Dim strCheckName
strCheckName = "GPS������"

'���ģ������
Dim CheckmodelName
CheckmodelName = "�Զ���ű������->GPS������"

'�������
Dim strDescription1
strDescription1 = "GPS���㸽�����������ۿ��Ƶ�"
Dim strDescription2
strDescription2 = "GPS���㸽�����ڶ�����ۿ��Ƶ�"
'==================================================ʵ����������=========================================================

'ʵ���Ԥ���Ӧ��ϵ��
' ʵ�����            ����                ���۱���
' 9130512           GPS����            1103021
' 9130412           ˮ׼��                1102021
' 9130311           ���Ƶ㣨��ʯ��         9130211
' 9130312           ���Ƶ㣨����ʯ��         9130212
' 9130217           ��վ��                9130216
' 9130511           ������               9130411


ScdCodes = "9130215"

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
    ' Dim idarr(SelCount)
    If SelCount > 0 Then
        For i = 0 To SelCount - 1
            id = SSProcess.GetSelGeoValue(i,"SSObj_ID")
            x = SSProcess.GetSelGeoValue(i,"SSObj_X")
            y = SSProcess.GetSelGeoValue(i,"SSObj_Y")
            z = SSProcess.GetObjectAttr(i,"SSObj_Z")
            idstr = SSProcess.SearchNearObjIDs(x, y, 0.1, 0, "9130211,9130212", 0)
            idarr = Split(idstr,",",-1,1)
            nearcount = UBound(idarr) + 1
            If idstr = "" Then
                SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription1, x, y, 0, 0,id, ""
            End If
            If nearcount > 1 Then
                SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription2, x, y, 0, 0,id, ""
            End If
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